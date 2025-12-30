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
  maxDuration: 60,
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
    
    console.log('=== Starting Excel to PDF Conversion ===');
    console.log(`File: ${filename || 'document.xlsx'}`);
    console.log(`Size: ${(buffer.length / 1024).toFixed(2)} KB`);
    
    // Load Excel file
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const worksheet = workbook.worksheets[0];
    
    console.log(`Worksheet: ${worksheet.name}`);
    console.log(`Dimensions: ${worksheet.rowCount} rows x ${worksheet.columnCount} columns`);
    
    // Extract images with positions
    const images = [];
    const imageMap = new Map();
    
    if (worksheet.getImages) {
      console.log('Extracting images...');
      worksheet.getImages().forEach((img, idx) => {
        const imageId = workbook.model.media[img.imageId];
        if (imageId) {
          const extension = imageId.extension || 'png';
          const base64 = imageId.buffer.toString('base64');
          const mimeType = extension === 'png' ? 'image/png' : 
                          extension === 'jpg' || extension === 'jpeg' ? 'image/jpeg' : 
                          'image/' + extension;
          
          // Get image position from range
          const range = img.range;
          
          images.push({
            id: idx,
            data: `data:${mimeType};base64,${base64}`,
            range: range,
            extension: extension
          });
          
          // Map to cell position
          if (range && range.tl) {
            const cellKey = `${range.tl.nativeRow + 1}-${range.tl.nativeCol + 1}`;
            imageMap.set(cellKey, {
              id: idx,
              data: `data:${mimeType};base64,${base64}`,
              rowSpan: (range.br.nativeRow - range.tl.nativeRow + 1) || 1,
              colSpan: (range.br.nativeCol - range.tl.nativeCol + 1) || 1
            });
          }
          
          console.log(`  Image ${idx + 1}: ${extension}, positioned at row ${range?.tl?.nativeRow + 1}, col ${range?.tl?.nativeCol + 1}`);
        }
      });
      console.log(`Total images extracted: ${images.length}`);
    }
    
    // Get column widths (for proportional HTML table)
    const columnWidths = [];
    for (let i = 1; i <= worksheet.columnCount; i++) {
      const col = worksheet.getColumn(i);
      columnWidths.push(col.width || 10); // Default 10 if not set
    }
    const totalWidth = columnWidths.reduce((a, b) => a + b, 0);
    
    // Track merged cells
    const mergedCells = new Map();
    const skippedCells = new Set();
    
    if (worksheet.model?.merges) {
      worksheet.model.merges.forEach(merge => {
        const match = merge.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        if (match) {
          const startCol = colLetterToNumber(match[1]);
          const startRow = parseInt(match[2]);
          const endCol = colLetterToNumber(match[3]);
          const endRow = parseInt(match[4]);
          
          const colspan = endCol - startCol + 1;
          const rowspan = endRow - startRow + 1;
          
          mergedCells.set(`${startRow}-${startCol}`, { colspan, rowspan });
          
          // Mark cells to skip
          for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
              if (r !== startRow || c !== startCol) {
                skippedCells.add(`${r}-${c}`);
              }
            }
          }
        }
      });
    }
    
    // Generate HTML
    let html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    @page {
      size: A4 portrait;
      margin: 0.25in;
    }
    * {
      box-sizing: border-box;
    }
    body {
      font-family: 'Calibri', 'Arial', sans-serif;
      font-size: 11pt;
      margin: 0;
      padding: 0;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }
    table {
      border-collapse: collapse;
      width: 100%;
      table-layout: fixed;
    }
    td {
      border: 1px solid #000;
      padding: 2px 4px;
      vertical-align: middle;
      overflow: hidden;
      word-wrap: break-word;
      position: relative;
    }
    .img-cell {
      padding: 0;
      text-align: center;
      vertical-align: middle;
    }
    .img-cell img {
      max-width: 100%;
      max-height: 100%;
      display: block;
      margin: 0 auto;
    }
    .nowrap { white-space: nowrap; }
  </style>
</head>
<body>
  <table>
`;
    
    // Generate colgroup for column widths
    html += '    <colgroup>\n';
    columnWidths.forEach(width => {
      const percent = ((width / totalWidth) * 100).toFixed(2);
      html += `      <col style="width: ${percent}%">\n`;
    });
    html += '    </colgroup>\n';
    
    // Process rows
    for (let rowNum = 1; rowNum <= worksheet.rowCount; rowNum++) {
      const row = worksheet.getRow(rowNum);
      const rowHeight = row.height || 15;
      
      html += `    <tr style="height: ${rowHeight}pt;">\n`;
      
      for (let colNum = 1; colNum <= worksheet.columnCount; colNum++) {
        const cellKey = `${rowNum}-${colNum}`;
        
        // Skip if part of merged cell (not top-left)
        if (skippedCells.has(cellKey)) {
          continue;
        }
        
        const cell = row.getCell(colNum);
        
        // Get merge info
        const mergeInfo = mergedCells.get(cellKey);
        const colspan = mergeInfo?.colspan || 1;
        const rowspan = mergeInfo?.rowspan || 1;
        
        // Check if this cell has an image
        const imageInfo = imageMap.get(cellKey);
        
        // Build cell attributes
        let attrs = '';
        if (colspan > 1) attrs += ` colspan="${colspan}"`;
        if (rowspan > 1) attrs += ` rowspan="${rowspan}"`;
        
        // Build inline styles
        let styles = [];
        let classes = [];
        
        // Alignment
        if (cell.alignment) {
          if (cell.alignment.horizontal) {
            styles.push(`text-align: ${cell.alignment.horizontal}`);
          }
          if (cell.alignment.vertical) {
            const vAlign = cell.alignment.vertical === 'middle' ? 'middle' : 
                          cell.alignment.vertical === 'top' ? 'top' : 'bottom';
            styles.push(`vertical-align: ${vAlign}`);
          }
          if (cell.alignment.wrapText) {
            styles.push('white-space: normal');
          } else {
            classes.push('nowrap');
          }
        }
        
        // Font styling
        if (cell.font) {
          if (cell.font.bold) styles.push('font-weight: bold');
          if (cell.font.italic) styles.push('font-style: italic');
          if (cell.font.size) styles.push(`font-size: ${cell.font.size}pt`);
          if (cell.font.color?.argb) {
            const color = '#' + cell.font.color.argb.slice(2);
            styles.push(`color: ${color}`);
          }
          if (cell.font.name) styles.push(`font-family: '${cell.font.name}'`);
        }
        
        // Background fill
        if (cell.fill && cell.fill.type === 'pattern' && cell.fill.fgColor?.argb) {
          const bgColor = '#' + cell.fill.fgColor.argb.slice(2);
          styles.push(`background-color: ${bgColor}`);
        }
        
        // Border styling
        if (cell.border) {
          ['top', 'bottom', 'left', 'right'].forEach(side => {
            const border = cell.border[side];
            if (border && border.style && border.style !== 'none') {
              const color = border.color?.argb ? '#' + border.color.argb.slice(2) : '#000';
              const weight = border.style === 'medium' || border.style === 'thick' ? '2px' : '1px';
              styles.push(`border-${side}: ${weight} solid ${color}`);
            }
          });
        }
        
        const styleAttr = styles.length > 0 ? ` style="${styles.join('; ')}"` : '';
        const classAttr = classes.length > 0 ? ` class="${classes.join(' ')}"` : '';
        
        // Cell content
        let content = '';
        
        if (imageInfo) {
          // This cell contains an image
          classes.push('img-cell');
          const imgHeight = rowHeight * rowspan * 0.75; // Convert pt to approximate px
          content = `<img src="${imageInfo.data}" style="max-height: ${imgHeight}px;" alt="Image">`;
        } else {
          // Regular cell with text
          let value = '';
          
          if (cell.value !== null && cell.value !== undefined) {
            if (cell.type === ExcelJS.ValueType.Formula) {
              value = cell.result !== undefined ? cell.result : cell.text || '';
            } else if (typeof cell.value === 'object') {
              if (cell.value.richText) {
                value = cell.value.richText.map(t => t.text).join('');
              } else if (cell.value.text) {
                value = cell.value.text;
              } else if (cell.value.result !== undefined) {
                value = cell.value.result;
              } else {
                value = cell.text || '';
              }
            } else {
              value = cell.value.toString();
            }
          }
          
          content = escapeHtml(value);
        }
        
        html += `      <td${attrs}${classAttr}${styleAttr}>${content}</td>\n`;
      }
      
      html += '    </tr>\n';
    }
    
    html += `  </table>
</body>
</html>`;
    
    console.log('HTML generated successfully');
    console.log(`HTML size: ${(html.length / 1024).toFixed(2)} KB`);
    
    // Convert HTML to PDF using Gotenberg Chromium
    console.log('Converting HTML to PDF with Chromium...');
    
    const form = new FormData();
    form.append('files', Buffer.from(html, 'utf8'), {
      filename: 'index.html',
      contentType: 'text/html; charset=utf-8'
    });
    
    // Chromium options
    form.append('marginTop', '0.25');
    form.append('marginBottom', '0.25');
    form.append('marginLeft', '0.25');
    form.append('marginRight', '0.25');
    form.append('preferCssPageSize', 'true');
    form.append('printBackground', 'true');
    form.append('scale', '1.0');
    form.append('paperWidth', '8.27'); // A4 width in inches
    form.append('paperHeight', '11.69'); // A4 height in inches
    
    const gotenbergResponse = await fetch(`${GOTENBERG_URL}/forms/chromium/convert/html`, {
      method: 'POST',
      body: form,
      headers: form.getHeaders(),
    });

    if (!gotenbergResponse.ok) {
      const errorText = await gotenbergResponse.text();
      console.error('Gotenberg error:', errorText);
      throw new Error(`Gotenberg conversion failed: ${errorText}`);
    }

    const pdfArrayBuffer = await gotenbergResponse.arrayBuffer();
    const pdfBuffer = Buffer.from(pdfArrayBuffer);
    
    console.log(`✅ PDF generated successfully! Size: ${(pdfBuffer.length / 1024).toFixed(2)} KB`);
    console.log('=== Conversion Complete ===');
    
    res.status(200).json({
      success: true,
      pdf: pdfBuffer.toString('base64'),
      filename: filename ? filename.replace(/\.xlsx?$/i, '.pdf') : 'converted.pdf',
      debug: {
        imagesExtracted: images.length,
        rowsProcessed: worksheet.rowCount,
        columnsProcessed: worksheet.columnCount
      }
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

// Helper functions
function colLetterToNumber(col) {
  let num = 0;
  for (let i = 0; i < col.length; i++) {
    num = num * 26 + (col.charCodeAt(i) - 64);
  }
  return num;
}

function escapeHtml(text) {
  if (text === null || text === undefined) return '';
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  return String(text).replace(/[&<>"']/g, m => map[m]);
}
