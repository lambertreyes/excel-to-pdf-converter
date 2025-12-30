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
    
    console.log('Step 1: Aggressive table removal from ZIP structure...');
    
    // STEP 1: Aggressively remove ALL table-related XML
    const zip = new AdmZip(buffer);
    const zipEntries = zip.getEntries();
    
    // Remove table XML files
    const tablesToRemove = [];
    zipEntries.forEach(entry => {
      if (entry.entryName.startsWith('xl/tables/')) {
        tablesToRemove.push(entry.entryName);
      }
    });
    
    tablesToRemove.forEach(name => {
      zip.deleteFile(name);
      console.log(`Deleted: ${name}`);
    });
    
    // Clean Content_Types.xml - remove table definitions
    const contentTypesEntry = zip.getEntry('[Content_Types].xml');
    if (contentTypesEntry) {
      let content = contentTypesEntry.getData().toString('utf8');
      content = content.replace(/<Override[^>]*PartName="\/xl\/tables\/[^"]*"[^>]*\/>/g, '');
      zip.updateFile('[Content_Types].xml', Buffer.from(content, 'utf8'));
      console.log('Cleaned [Content_Types].xml');
    }
    
    // Clean ALL worksheet XML files - remove tableParts completely
    zipEntries.forEach(entry => {
      const name = entry.entryName;
      if (name.startsWith('xl/worksheets/') && name.endsWith('.xml')) {
        let content = entry.getData().toString('utf8');
        const original = content;
        
        // Remove tableParts section entirely
        content = content.replace(/<tableParts[^>]*>[\s\S]*?<\/tableParts>/g, '');
        content = content.replace(/<tableParts[^>]*\/>/g, '');
        content = content.replace(/<tablePart[^>]*\/>/g, '');
        
        if (content !== original) {
          zip.updateFile(name, Buffer.from(content, 'utf8'));
          console.log(`Cleaned worksheet: ${name}`);
        }
      }
    });
    
    // Clean ALL relationship files - remove table relationships
    zipEntries.forEach(entry => {
      const name = entry.entryName;
      if (name.endsWith('.xml.rels')) {
        let content = entry.getData().toString('utf8');
        const original = content;
        
        content = content.replace(/<Relationship[^>]*Type="[^"]*table"[^>]*\/>/gi, '');
        content = content.replace(/<Relationship[^>]*Target="[^"]*\/tables\/[^"]*"[^>]*\/>/g, '');
        
        if (content !== original) {
          zip.updateFile(name, Buffer.from(content, 'utf8'));
          console.log(`Cleaned relationships: ${name}`);
        }
      }
    });
    
    const cleanedBuffer = zip.toBuffer();
    
    console.log('Step 2: Loading with ExcelJS and converting to values...');
    
    // STEP 2: Load with ExcelJS and convert everything to values
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(cleanedBuffer);
    
    // Create a NEW workbook and copy data as pure values
    const newWorkbook = new ExcelJS.Workbook();
    
    workbook.eachSheet((sourceSheet, sheetId) => {
      console.log(`Processing sheet: ${sourceSheet.name}`);
      
      const newSheet = newWorkbook.addWorksheet(sourceSheet.name, {
        pageSetup: {
          paperSize: 9,
          orientation: 'portrait',
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
          margins: {
            left: 0.25,
            right: 0.25,
            top: 0.25,
            bottom: 0.25,
            header: 0,
            footer: 0
          }
        }
      });
      
      // Copy column widths
      sourceSheet.columns?.forEach((col, idx) => {
        if (col.width) {
          newSheet.getColumn(idx + 1).width = col.width;
        }
      });
      
      // Copy all rows with values only (no formulas)
      sourceSheet.eachRow((sourceRow, rowNumber) => {
        const newRow = newSheet.getRow(rowNumber);
        
        // Set row height
        if (sourceRow.height) {
          newRow.height = sourceRow.height;
        }
        
        sourceRow.eachCell({ includeEmpty: true }, (sourceCell, colNumber) => {
          const newCell = newRow.getCell(colNumber);
          
          // Get the ACTUAL VALUE (not formula)
          let cellValue = sourceCell.value;
          
          // If it's a formula, get the result
          if (sourceCell.type === ExcelJS.ValueType.Formula) {
            cellValue = sourceCell.result || sourceCell.text || '';
          } else if (cellValue && typeof cellValue === 'object') {
            // Handle rich text and other complex values
            if (cellValue.richText) {
              cellValue = cellValue.richText.map(t => t.text).join('');
            } else if (cellValue.result !== undefined) {
              cellValue = cellValue.result;
            } else if (cellValue.text) {
              cellValue = cellValue.text;
            }
          }
          
          // Set the value
          newCell.value = cellValue;
          
          // Copy styling (but not table-specific styles)
          if (sourceCell.style) {
            newCell.style = {
              font: sourceCell.font,
              alignment: sourceCell.alignment,
              border: sourceCell.border,
              fill: sourceCell.fill,
              numFmt: sourceCell.numFmt
            };
          }
        });
        
        newRow.commit();
      });
      
      // Copy merged cells
      if (sourceSheet.model?.merges) {
        sourceSheet.model.merges.forEach(merge => {
          newSheet.mergeCells(merge);
        });
      }
      
      // Set print area if exists
      if (sourceSheet.pageSetup?.printArea) {
        newSheet.pageSetup.printArea = sourceSheet.pageSetup.printArea;
      }
    });
    
    console.log('Step 3: Saving clean workbook...');
    
    // Save the NEW workbook (completely table-free)
    const finalBuffer = await newWorkbook.xlsx.writeBuffer();
    
    console.log('Step 4: Converting to PDF with Gotenberg...');
    
    // STEP 3: Send to Gotenberg
    const form = new FormData();
    form.append('files', finalBuffer, {
      filename: filename || 'document.xlsx',
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    
    // Gotenberg parameters for best quality
    form.append('landscape', 'false');
    form.append('nativePageRanges', '1-100');
    form.append('exportFormFields', 'false');
    form.append('losslessImageCompression', 'true');
    form.append('quality', '100');
    
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
    
    console.log('✅ Conversion successful!');
    
    res.status(200).json({
      success: true,
      pdf: pdfBuffer.toString('base64'),
      filename: filename ? filename.replace(/\.xlsx?$/i, '.pdf') : 'converted.pdf'
    });
    
  } catch (error) {
    console.error('❌ Conversion error:', error);
    res.status(500).json({ 
      error: 'Conversion failed', 
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
}
