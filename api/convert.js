// api/convert.js
import FormData from 'form-data';
import fetch from 'node-fetch';

export const config = {
  api: {
    bodyParser: {
      sizeLimit: '10mb',
    },
  },
  maxDuration: 30, // Allow up to 30 seconds
};

// Set your Gotenberg URL (Railway deployment)
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
    
    // Create form data for Gotenberg
    const form = new FormData();
    form.append('files', buffer, {
      filename: filename || 'document.xlsx',
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    
    // Optional: Add conversion parameters
    // form.append('landscape', 'false'); // Portrait mode
    // form.append('nativePageRanges', '1-1'); // First page only
    
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
