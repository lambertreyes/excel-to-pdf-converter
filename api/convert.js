// api/convert.js
import FormData from 'form-data';
import fetch from 'node-fetch';
import AdmZip from 'adm-zip';
import ExcelJS from 'exceljs';

export const config = {
  api: {
    bodyParser: {
      sizeLimit: '10mb',
    },
  },
  maxDuration: 30,
};

const GOTENBERG_URL = process.env.GOTENBERG_URL || 'https://your-gotenberg.railway.app';

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { file, filename } = req.body;
    
    if (!file) {
      return res.status(400).json({ error: 'No file provided' });
    }

    const buffer = Buffer.from(file, 'base64');
    
    console.log('Step 1: Loading workbook...');
    
    // Load the workbook with ExcelJS first
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    
    console.log('Step 2: Converting tables to ranges...');
    
    // Process each worksheet
    workbook.eachSheet((worksheet, sheetId) => {
      console.log(`Processing sheet: ${worksheet.name}`);
      
      // Get all tables in the worksheet
      const tables = worksheet.tables || [];
      
      if (tables.length > 0) {
        console.log(`Found ${tables.length} tables in ${worksheet.name}`);
        
        tables.forEach(table => {
          console.log(`Converting table: ${table.name || 'unnamed'}`);
          
          // Get table range
          const tableRef = table.ref || table.tableRef;
          if (!tableRef) return;
          
          // Parse the range (e.g., "A1:E10")
          const match = tableRef.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
          if (!match) return;
          
          const [, startCol, startRow, endCol, endRow] = match;
          const startRowNum = parseInt(startRow);
          const endRowNum = parseInt(endRow);
          
          // Convert column letters to numbers
          const colToNum = (col) => {
            let num = 0;
            for (let i = 0; i < col.length; i++) {
              num = num * 26 + (col.charCodeAt(i) - 64);
            }
            return num;
          };
          
          const startColNum = colToNum(startCol);
          const endColNum = colToNum(endCol);
          
          // Iterate through all cells in the table range
          for (let rowNum = startRowNum; rowNum <= endRowNum; rowNum++) {
            const row = worksheet.getRow(rowNum);
            
            for (let colNum = startColNum; colNum <= endColNum; colNum++) {
              const cell = row.getCell(colNum);
              
              // Convert formulas to values
              if (cell.formula || cell.type === ExcelJS.ValueType.Formula) {
                const value = cell.value;
                // If it's a formula result, get the calculated value
                if (value && typeof value === 'object' && 'result' in value) {
                  cell.value = value.result;
                } else if (cell.text) {
                  // Use displayed text as fallback
                  cell.value = cell.text;
                }
              }
              
              // Clear any table-specific styling that might cause issues
              // Keep basic formatting but remove table references
              if (cell.style) {
                // Remove any named styles that might reference tables
                delete cell.style.name;
              }
            }
          }
        });
        
        // Clear all table definitions from the worksheet
        worksheet.tables = [];
      }
      
      // Set optimal page setup for PDF conversion
      worksheet.pageSetup = {
        paperSize: 9, // A4
        orientation: 'portrait',
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0, // Allow multiple pages vertically if needed
        margins: {
          left: 0.25,
          right: 0.25,
          top: 0.25,
          bottom: 0.25,
          header: 0,
          footer: 0
        },
        horizontalCentered: true,
        verticalCentered: false,
        printArea: worksheet.pageSetup?.printArea
      };
      
      // Ensure all rows are properly committed
      worksheet.commit();
    });
    
    console.log('Step 3: Removing table definitions from XML structure...');
    
    // Write the modified workbook to buffer
    const modifiedBuffer = await workbook.xlsx.writeBuffer();
    
    // Additional cleanup: Remove table XML files from the ZIP
    const zip = new AdmZip(modifiedBuffer);
    const zipEntries = zip.getEntries();
    
    let tablesRemoved = 0;
    const entriesToRemove = [];
    
    // Find and mark table files for removal
    zipEntries.forEach(entry => {
      const name = entry.entryName;
      if (name.startsWith('xl/tables/') && name.endsWith('.xml')) {
        entriesToRemove.push(name);
      }
    });
    
    // Delete table files
    entriesToRemove.forEach(name => {
      zip.deleteFile(name);
      tablesRemoved++;
      console.log(`Removed: ${name}`);
    });
    
    // Clean up Content_Types.xml
    const contentTypesEntry = zip.getEntry('[Content_Types].xml');
    if (contentTypesEntry) {
      let content = contentTypesEntry.getData().toString('utf8');
      content = content.replace(/<Override[^>]*PartName="\/xl\/tables\/[^"]*"[^>]*\/>/g, '');
      zip.updateFile('[Content_Types].xml', Buffer.from(content, 'utf8'));
    }
    
    // Clean up worksheet relationships
    zipEntries.forEach(entry => {
      const name = entry.entryName;
      if (name.includes('/_rels/') && name.endsWith('.xml.rels')) {
        let content = entry.getData().toString('utf8');
        const originalLength = content.length;
        content = content.replace(/<Relationship[^>]*Target="[^"]*\/tables\/[^"]*"[^>]*\/>/g, '');
        if (content.length !== originalLength) {
          zip.updateFile(name, Buffer.from(content, 'utf8'));
        }
      }
    });
    
    console.log(`Tables removed from ZIP: ${tablesRemoved}`);
    
    // Get final cleaned buffer
    const finalBuffer = zip.toBuffer();
    
    console.log('Step 4: Converting to PDF with Gotenberg...');
    
    // Send to Gotenberg
    const form = new FormData();
    form.append('files', finalBuffer, {
      filename: filename || 'document.xlsx',
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    
    form.append('landscape', 'false');
    form.append('nativePageRanges', '1-50');
    form.append('exportFormFields', 'false');
    form.append('losslessImageCompression', 'true');
    form.append('quality', '100');
    form.append('pdfa', 'PDF/A-1b'); // Better compatibility
    
    const gotenbergResponse = await fetch(`${GOTENBERG_URL}/forms/libreoffice/convert`, {
      method: 'POST',
      body: form,
      headers: form.getHeaders(),
    });

    if (!gotenbergResponse.ok) {
      const errorText = await gotenbergResponse.text();
      throw new Error(`Gotenberg conversion failed: ${errorText}`);
    }

    const pdfArrayBuffer = await gotenbergResponse.arrayBuffer();
    const pdfBuffer = Buffer.from(pdfArrayBuffer);
    
    console.log('Conversion successful!');
    
    res.status(200).json({
      success: true,
      pdf: pdfBuffer.toString('base64'),
      filename: filename ? filename.replace(/\.xlsx?$/i, '.pdf') : 'converted.pdf'
    });
    
  } catch (error) {
    console.error('Conversion error:', error);
    res.status(500).json({ 
      error: 'Conversion failed', 
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
}
