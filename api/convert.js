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
    
    // Load workbook to remove table formatting
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const worksheet = workbook.worksheets[0];
    
    // Remove all Excel tables (convert to regular ranges)
    if (worksheet.tables && worksheet.tables.length > 0) {
      // Get table info before removing
      const tablesToRemove = [];
      worksheet.tables.forEach(table => {
        tablesToRemove.push({
          name: table.name,
          ref: table.ref
        });
      });
      
      // Remove tables (this preserves formatting but removes table structure)
      tablesToRemove.forEach(tableInfo => {
        try {
          worksheet.removeTable(tableInfo.name);
        } catch (err) {
          console.log('Could not remove table:', tableInfo.name);
        }
      });
    }
    
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
    
    // Re-save the modified Excel
    const modifiedBuffer = await workbook.xlsx.writeBuffer();
    
    // Create form data for Gotenberg
    const form = new FormData();
    form.append('files', modifiedBuffer, {
      filename: filename || 'document.xlsx',
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    
    // LibreOffice conversion parameters
    form.append('pdfa', 'PDF/A-1a');
    form.append('pdfua', 'true');
    form.append('landscape', 'false');
    form.append('nativePdfFormat', 'true');
    form.append('exportFormFields', 'false');
    form.append('exportBookmarks', 'false');
    form.append('losslessImageCompression', 'true');
    form.append('reduceImageResolution', 'false');
    form.append('quality', '100');
    form.append('merge', 'false');
    
    // Call Gotenberg API
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
