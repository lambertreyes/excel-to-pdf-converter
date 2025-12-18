// api/convert.js
import ExcelJS from 'exceljs';
import PDFDocument from 'pdfkit';
import { Readable } from 'stream';

export const config = {
  api: {
    bodyParser: {
      sizeLimit: '10mb',
    },
  },
};

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
    
    // Load Excel workbook
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    
    // Get first worksheet
    const worksheet = workbook.worksheets[0];
    
    // Create PDF
    const pdfBuffer = await createPDF(worksheet);
    
    // Return PDF as base64
    res.status(200).json({
      success: true,
      pdf: pdfBuffer.toString('base64'),
      filename: filename ? filename.replace('.xlsx', '.pdf') : 'converted.pdf'
    });
    
  } catch (error) {
    console.error('Conversion error:', error);
    res.status(500).json({ 
      error: 'Conversion failed', 
      details: error.message 
    });
  }
}

async function createPDF(worksheet) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ 
      size: 'A4',
      margin: 40,
      layout: 'portrait'
    });
    
    const chunks = [];
    doc.on('data', chunk => chunks.push(chunk));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    // Calculate column widths
    const colWidths = {};
    let maxCol = 0;
    
    worksheet.eachRow((row, rowNum) => {
      row.eachCell((cell, colNum) => {
        maxCol = Math.max(maxCol, colNum);
        const width = worksheet.getColumn(colNum).width || 10;
        colWidths[colNum] = Math.max(colWidths[colNum] || 0, width * 7);
      });
    });

    // Calculate available width and adjust columns proportionally
    const totalWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
    const totalColWidth = Object.values(colWidths).reduce((a, b) => a + b, 0);
    const scale = totalWidth / totalColWidth;
    
    Object.keys(colWidths).forEach(col => {
      colWidths[col] *= scale;
    });

    let startY = doc.y;
    const rowHeight = 20;

    // Render each row
    worksheet.eachRow((row, rowNum) => {
      let currentX = doc.page.margins.left;
      const currentY = startY + (rowNum - 1) * rowHeight;

      // Check if we need a new page
      if (currentY + rowHeight > doc.page.height - doc.page.margins.bottom) {
        doc.addPage();
        startY = doc.page.margins.top;
        currentX = doc.page.margins.left;
      }

      row.eachCell({ includeEmpty: true }, (cell, colNum) => {
        const cellWidth = colWidths[colNum] || 60;
        const cellX = currentX;
        const cellY = currentY;

        // Draw cell background
        if (cell.style && cell.style.fill && cell.style.fill.fgColor) {
          const color = cell.style.fill.fgColor.argb;
          if (color && color !== 'FFFFFFFF') {
            const rgb = hexToRgb(color);
            if (rgb) {
              doc.rect(cellX, cellY, cellWidth, rowHeight)
                 .fillAndStroke(rgb, '#000000');
            }
          }
        }

        // Draw cell border
        doc.rect(cellX, cellY, cellWidth, rowHeight).stroke();

        // Draw cell text
        let cellValue = '';
        if (cell.value !== null && cell.value !== undefined) {
          if (typeof cell.value === 'object' && cell.value.text) {
            cellValue = cell.value.text;
          } else {
            cellValue = String(cell.value);
          }
        }

        if (cellValue) {
          // Set font style
          let fontStyle = 'Helvetica';
          if (cell.font) {
            if (cell.font.bold && cell.font.italic) {
              fontStyle = 'Helvetica-BoldOblique';
            } else if (cell.font.bold) {
              fontStyle = 'Helvetica-Bold';
            } else if (cell.font.italic) {
              fontStyle = 'Helvetica-Oblique';
            }
          }
          doc.font(fontStyle);

          // Set font size
          const fontSize = (cell.font && cell.font.size) ? cell.font.size : 10;
          doc.fontSize(fontSize);

          // Set font color
          if (cell.font && cell.font.color && cell.font.color.argb) {
            const textColor = hexToRgb(cell.font.color.argb);
            if (textColor) {
              doc.fillColor(textColor);
            }
          } else {
            doc.fillColor('#000000');
          }

          // Handle text alignment
          const align = cell.alignment && cell.alignment.horizontal ? cell.alignment.horizontal : 'left';
          const verticalAlign = cell.alignment && cell.alignment.vertical ? cell.alignment.vertical : 'middle';
          
          let textX = cellX + 3;
          if (align === 'center') {
            textX = cellX + cellWidth / 2;
          } else if (align === 'right') {
            textX = cellX + cellWidth - 3;
          }

          let textY = cellY + (rowHeight - fontSize) / 2;
          if (verticalAlign === 'top') {
            textY = cellY + 3;
          } else if (verticalAlign === 'bottom') {
            textY = cellY + rowHeight - fontSize - 3;
          }

          doc.text(cellValue, textX, textY, {
            width: cellWidth - 6,
            align: align,
            lineBreak: false,
            ellipsis: true
          });
        }

        currentX += cellWidth;
      });
    });

    doc.end();
  });
}

function hexToRgb(hex) {
  // Remove alpha channel if present (ARGB format)
  if (hex.length === 8) {
    hex = hex.substring(2);
  }
  
  const result = /^([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result ? [
    parseInt(result[1], 16),
    parseInt(result[2], 16),
    parseInt(result[3], 16)
  ] : null;
}
