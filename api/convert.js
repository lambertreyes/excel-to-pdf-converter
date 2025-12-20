// api/convert.js
import FormData from 'form-data';
import fetch from 'node-fetch';
import AdmZip from 'adm-zip';

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
    
    console.log('Removing Excel tables by manipulating XML...');
    
    // Excel files are ZIP archives - we'll extract, modify, and repack
    const zip = new AdmZip(buffer);
    const zipEntries = zip.getEntries();
    
    let tablesRemoved = 0;
    let tablesFound = 0;
    
    // Step 1: Find and remove table XML files
    const entriesToRemove = [];
    zipEntries.forEach(entry => {
      const entryName = entry.entryName;
      
      // Remove table definition files (xl/tables/tableX.xml)
      if (entryName.startsWith('xl/tables/table') && entryName.endsWith('.xml')) {
        entriesToRemove.push(entryName);
        tablesFound++;
        console.log(`Found table file: ${entryName}`);
      }
    });
    
    // Delete table files
    entriesToRemove.forEach(entryName => {
      zip.deleteFile(entryName);
      tablesRemoved++;
      console.log(`Removed: ${entryName}`);
    });
    
    // Step 2: Update worksheet XML to remove table references
    zipEntries.forEach(entry => {
      const entryName = entry.entryName;
      
      // Modify worksheet files (xl/worksheets/sheetX.xml)
      if (entryName.startsWith('xl/worksheets/sheet') && entryName.endsWith('.xml')) {
        let content = entry.getData().toString('utf8');
        
        // Remove <tableParts> section entirely
        content = content.replace(/<tableParts[^>]*>[\s\S]*?<\/tableParts>/g, '');
        
        // Remove <tablePart> references
        content = content.replace(/<tablePart[^>]*\/>/g, '');
        
        // Update the entry
        zip.updateFile(entryName, Buffer.from(content, 'utf8'));
        console.log(`Updated worksheet: ${entryName}`);
      }
    });
    
    // Step 3: Update [Content_Types].xml to remove table content type
    const contentTypesEntry = zip.getEntry('[Content_Types].xml');
    if (contentTypesEntry) {
      let contentTypes = contentTypesEntry.getData().toString('utf8');
      
      // Remove table-related content types
      contentTypes = contentTypes.replace(/<Override[^>]*PartName="\/xl\/tables\/[^"]*"[^>]*\/>/g, '');
      
      zip.updateFile('[Content_Types].xml', Buffer.from(contentTypes, 'utf8'));
      console.log('Updated [Content_Types].xml');
    }
    
    // Step 4: Update xl/_rels/workbook.xml.rels to remove table relationships
    const workbookRelsEntry = zip.getEntry('xl/_rels/workbook.xml.rels');
    if (workbookRelsEntry) {
      let workbookRels = workbookRelsEntry.getData().toString('utf8');
      
      // Remove table relationship references
      workbookRels = workbookRels.replace(/<Relationship[^>]*Type="[^"]*\/table"[^>]*\/>/g, '');
      
      zip.updateFile('xl/_rels/workbook.xml.rels', Buffer.from(workbookRels, 'utf8'));
      console.log('Updated workbook relationships');
    }
    
    // Step 5: Update worksheet relationships
    zipEntries.forEach(entry => {
      const entryName = entry.entryName;
      
      if (entryName.startsWith('xl/worksheets/_rels/sheet') && entryName.endsWith('.xml.rels')) {
        let content = entry.getData().toString('utf8');
        
        // Remove table relationship references
        const originalLength = content.length;
        content = content.replace(/<Relationship[^>]*Target="..\/tables\/table[^"]*"[^>]*\/>/g, '');
        
        if (content.length !== originalLength) {
          zip.updateFile(entryName, Buffer.from(content, 'utf8'));
          console.log(`Updated worksheet relationships: ${entryName}`);
        }
      }
    });
    
    console.log(`Tables removed: ${tablesRemoved} of ${tablesFound} found`);
    
    // Generate the modified Excel file
    const modifiedBuffer = zip.toBuffer();
    console.log('Created clean Excel file without tables');
    
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
