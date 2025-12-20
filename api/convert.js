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
    
    console.log('Loading original workbook...');
    
    // Load original workbook
    const originalWorkbook = new ExcelJS.Workbook();
    await originalWorkbook.xlsx.load(buffer);
    const originalSheet = originalWorkbook.worksheets[0];
    
    console.log('Creating new workbook without tables...');
    
    // Create a brand new workbook (no tables possible)
    const newWorkbook = new ExcelJS.Workbook();
    const newSheet = newWorkbook.addWorksheet(originalSheet.name || 'Sheet1');
    
    // Copy worksheet properties
    newSheet.properties = { ...originalSheet.properties };
    
    // Copy page setup
    if (originalSheet.pageSetup) {
      newSheet.pageSetup = {
        paperSize: 9, // A4
        orientation: 'portrait',
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 1,
        margins: {
          left: 0.25,
          right: 0.25,
          top: 0.25,
          bottom: 0.25,
          header: 0,
          footer: 0
        },
        printArea: originalSheet.pageSetup.printArea || 'A1:O49'
      };
    } else {
      newSheet.pageSetup = {
        paperSize: 9,
        orientation: 'portrait',
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 1,
        margins: {
          left: 0.25,
          right: 0.25,
          top: 0.25,
          bottom: 0.25,
          header: 0,
          footer: 0
        },
        printArea: 'A1:O49'
      };
    }
    
    // Copy column widths
    originalSheet.columns.forEach((col, index) => {
      if (col && col.width) {
        const newCol = newSheet.getColumn(index + 1);
        newCol.width = col.width;
      }
    });
    
    // Copy all rows with their properties
    console.log('Copying cells...');
    originalSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const newRow = newSheet.getRow(rowNumber);
      
      // Copy row height
      if (row.height) {
        newRow.height = row.height;
      }
      
      // Copy each cell
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const newCell = newRow.getCell(colNumber);
        
        // Copy value
        newCell.value = cell.value;
        
        // Copy all styling
        if (cell.style) {
          newCell.style = {
            numFmt: cell.style.numFmt,
            font: cell.style.font ? { ...cell.style.font } : undefined,
            alignment: cell.style.alignment ? { ...cell.style.alignment } : undefined,
            protection: cell.style.protection ? { ...cell.style.protection } : undefined,
            border: cell.style.border ? {
              top: cell.style.border.top ? { ...cell.style.border.top } : undefined,
              left: cell.style.border.left ? { ...cell.style.border.left } : undefined,
              bottom: cell.style.border.bottom ? { ...cell.style.border.bottom } : undefined,
              right: cell.style.border.right ? { ...cell.style.border.right } : undefined,
              diagonal: cell.style.border.diagonal ? { ...cell.style.border.diagonal } : undefined
            } : undefined,
            fill: cell.style.fill ? { ...cell.style.fill } : undefined
          };
        }
      });
      
      newRow.commit();
    });
    
    // Copy merged cells
    if (originalSheet._merges) {
      console.log('Copying merged cells...');
      Object.values(originalSheet._merges).forEach(merge => {
        try {
          newSheet.mergeCells(merge);
        } catch (err) {
          console.log('Could not merge:', merge);
        }
      });
    }
    
    // Copy images
    console.log('Copying images...');
    if (originalWorkbook.model && originalWorkbook.model.media) {
      originalWorkbook.model.media.forEach((media, index) => {
        try {
          const imageId = newWorkbook.addImage({
            buffer: media.buffer,
            extension: media.extension
          });
          
          // Find where this image is placed in original
          if (originalSheet.getImages) {
            const images = originalSheet.getImages();
            images.forEach(img => {
              if (img.imageId === index) {
                newSheet.addImage(imageId, {
                  tl: img.range.tl,
                  br: img.range.br,
                  editAs: img.range.editAs
                });
              }
            });
          }
        } catch (err) {
          console.log('Could not copy image:', err.message);
        }
      });
    }
    
    console.log('Creating clean Excel file...');
    
    // Save the new workbook (completely clean, no tables)
    const cleanBuffer = await newWorkbook.xlsx.writeBuffer();
    
    console.log('Clean Excel created successfully');
    
    // Create form data for Gotenberg
    const form = new FormData();
    form.append('files', cleanBuffer, {
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
