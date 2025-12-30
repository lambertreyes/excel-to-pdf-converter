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
    
    console.log('Step 1: Removing table XML files...');
    
    // STEP 1: Remove table XML files directly from ZIP
    const zip = new AdmZip(buffer);
    const zipEntries = zip.getEntries();
    
    let tablesRemoved = 0;
    const entriesToRemove = [];
    
    // Find and mark table files for removal
    zipEntries.forEach(entry => {
      const name = entry.entryName;
      if (name.startsWith('xl/tables/') && name.endsWith('.xml')) {
        entriesToRemove.push(name);
        console.log(`Marking for removal: ${name}`);
      }
    });
    
    // Delete table files
    entriesToRemove.forEach(name => {
      zip.deleteFile(name);
      tablesRemoved++;
    });
    
    // Update [Content_Types].xml
    const contentTypesEntry = zip.getEntry('[Content_Types].xml');
    if (contentTypesEntry) {
      let content = contentTypesEntry.getData().toString('utf8');
      content = content.replace(/<Override[^>]*PartName="\/xl\/tables\/[^"]*"[^>]*\/>/g, '');
      zip.updateFile('[Content_Types].xml', Buffer.from(content, 'utf8'));
    }
    
    // Update worksheet XML to remove tableParts
    zipEntries.forEach(entry => {
      const name = entry.entryName;
      if (name.startsWith('xl/worksheets/sheet') && name.endsWith('.xml')) {
        let content = entry.getData().toString('utf8');
        content = content.replace(/<tableParts[^>]*>[\s\S]*?<\/tableParts>/g, '');
        content = content.replace(/<tablePart[^>]*\/>/g, '');
        zip.updateFile(name, Buffer.from(content, 'utf8'));
      }
    });
    
    // Update worksheet relationships
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
    
    console.log(`Tables removed: ${tablesRemoved}`);
    
    // Get the modified ZIP buffer
    const modifiedZipBuffer = zip.toBuffer();
    
    console.log('Step 2: Loading with ExcelJS to normalize...');
    
    // STEP 2: Load the modified file with ExcelJS and re-save
    // This normalizes the workbook structure (like Office Script does)
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(modifiedZipBuffer);
    
    // Force optimal page setup
    const worksheet = workbook.worksheets[0];
    worksheet.pageSetup = {
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
      printArea: worksheet.pageSetup?.printArea || 'A1:O49'
    };
    
    // Re-save to normalize the structure
    console.log('Step 3: Re-saving workbook to normalize structure...');
    const normalizedBuffer = await workbook.xlsx.writeBuffer();
    
    console.log('Step 4: Converting to PDF with Gotenberg...');
    
    // STEP 3: Send normalized file to Gotenberg
    const form = new FormData();
    form.append('files', normalizedBuffer, {
      filename: filename || 'document.xlsx',
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    
    form.append('landscape', 'false');
    form.append('nativePageRanges', '1-50');
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

    // ✅ FIXED: Use arrayBuffer() instead of buffer()
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
      details: error.message 
    });
  }
}
