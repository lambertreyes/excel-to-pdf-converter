// api/convert.js
import FormData from 'form-data';
import fetch from 'node-fetch';
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

    // Decode base64 Excel file
    const buffer = Buffer.from(file, 'base64');
    
    // Load workbook
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const worksheet = workbook.worksheets[0];
    
    console.log('Starting table removal process...');
    
    // COMPREHENSIVE TABLE REMOVAL
    
    // Step 1: Get table ranges before removal
    const tableRanges = [];
    if (worksheet.tables && worksheet.tables.length > 0) {
      console.log(`Found ${worksheet.tables.length} Excel tables`);
      worksheet.tables.forEach(table => {
        console.log(`Table: ${table.name}, Range: ${table.ref}`);
        tableRanges.push({
          name: table.name,
          ref: table.ref,
          displayName: table.displayName
        });
      });
    }
    
    // Step 2: Force remove tables by clearing the tables array
    // This is more aggressive than removeTable()
    if (worksheet.tables) {
      worksheet.tables.length = 0; // Clear all tables
      console.log('Cleared worksheet.tables array');
    }
    
    // Step 3: Clear internal table references
    if (worksheet._tables) {
      worksheet._tables = [];
      console.log('Cleared internal _tables');
    }
    
    // Step 4: Remove table parts from workbook
    if (workbook._tables) {
      workbook._tables = [];
    }
    
    // Step 5: Remove AutoFilter (commonly used with tables)
    if (worksheet.autoFilter) {
      worksheet.autoFilter = null;
      console.log('Removed autofilter');
    }
    
    // Step 6: Clear defined names that reference tables
    if (workbook.definedNames) {
      const namesToRemove = [];
      workbook.definedNames.forEach((name, index) => {
        if (name && (name.includes('Table') || name.includes('_FilterDatabase'))) {
          namesToRemove.push(index);
        }
      });
      // Remove in reverse order to maintain indices
      namesToRemove.reverse().forEach(index => {
        workbook.definedNames.splice(index, 1);
      });
      if (namesToRemove.length > 0) {
        console.log(`Removed ${namesToRemove.length} table-related defined names`);
      }
    }
    
    // Step 7: Ensure all cells have explicit styling (not table-inherited)
    // This forces cells to keep their visual appearance even after table removal
    tableRanges.forEach(tableInfo => {
      try {
        const range = worksheet.getCell(tableInfo.ref.split(':')[0]).address + ':' + 
                     worksheet.getCell(tableInfo.ref.split(':')[1]).address;
        
        // Parse range
        const match = tableInfo.ref.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        if (match) {
          const startCol = columnLetterToNumber(match[1]);
          const startRow = parseInt(match[2]);
          const endCol = columnLetterToNumber(match[3]);
          const endRow = parseInt(match[4]);
          
          // Iterate through all cells in table range
          for (let row = startRow; row <= endRow; row++) {
            for (let col = startCol; col <= endCol; col++) {
              const cell = worksheet.getCell(row, col);
              
              // Force explicit styling if cell has value
              if (cell.value !== null && cell.value !== undefined) {
                // Ensure borders are explicit
                if (!cell.border || Object.keys(cell.border).length === 0) {
                  cell.border = {
                    top: {style: 'thin', color: {argb: 'FF000000'}},
                    left: {style: 'thin', color: {argb: 'FF000000'}},
                    bottom: {style: 'thin', color: {argb: 'FF000000'}},
                    right: {style: 'thin', color: {argb: 'FF000000'}}
                  };
                }
              }
            }
          }
        }
      } catch (err) {
        console.error('Error processing table range:', err.message);
      }
    });
    
    console.log('Table removal complete');
    
    // Ensure print settings are optimal
    worksheet.pageSetup = {
      paperSize: 9, // A4
      orientation: 'portrait',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
      horizontalCentered: false,
      verticalCentered: false,
      margins: {
        left: 0.25,
        right: 0.25,
        top: 0.25,
        bottom: 0.25,
        header: 0,
        footer: 0
      }
    };
    
    // Set print area if not already set
    if (!worksheet.pageSetup.printArea) {
      worksheet.pageSetup.printArea = 'A1:O49';
    }
    
    // Re-save the modified Excel (this finalizes the table removal)
    console.log('Re-saving Excel without tables...');
    const modifiedBuffer = await workbook.xlsx.writeBuffer();
    console.log('Excel re-saved successfully');
    
    // Create form data for Gotenberg
    const form = new FormData();
    form.append('files', modifiedBuffer, {
      filename: filename || 'document.xlsx',
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    
    // LibreOffice conversion parameters
    form.append('landscape', 'false');
    form.append('nativePageRanges', '1-50');
    form.append('exportFormFields', 'false');
    form.append('losslessImageCompression', 'true');
    form.append('quality', '100');
    
    // Call Gotenberg API
    console.log('Sending to Gotenberg...');
    const gotenbergResponse = await fetch(`${GOTENBERG_URL}/forms/libreoffice/convert`, {
      method: 'POST',
      body: form,
      headers: form.getHeaders(),
    });

    if (!gotenbergResponse.ok) {
      const errorText = await gotenbergResponse.text();
      throw new Error(`Gotenberg conversion failed: ${errorText}`);
    }

    // Get PDF buffer
    const pdfBuffer = await gotenbergResponse.buffer();
    console.log('PDF conversion successful');
    
    // Return PDF as base64
    res.status(200).json({
      success: true,
      pdf: pdfBuffer.toString('base64'),
      filename: filename ? filename.replace(/\.xlsx?$/i, '.pdf') : 'converted.pdf'
    });
    
  } catch (error) {
    console.error('Conversion error:', error);
    res.status(500).json({ 
      error: 'Conversion failed', 
      details: error.message 
    });
  }
}

function columnLetterToNumber(letter) {
  let column = 0;
  for (let i = 0; i < letter.length; i++) {
    column = column * 26 + letter.charCodeAt(i) - 64;
  }
  return column;
}
