// api/convert.js
import ExcelJS from 'exceljs';
import PDFDocument from 'pdfkit';

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
    const pdfBuffer = await createPDF(worksheet, workbook);
    
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

async function createPDF(worksheet, workbook) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ 
      size: 'A4',
      margin: 25,
      layout: 'portrait'
    });
    
    const chunks = [];
    doc.on('data', chunk => chunks.push(chunk));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    // Get merged cells info
    const mergedRanges = [];
    for (const key in worksheet._merges) {
      const range = worksheet._merges[key];
      if (typeof range === 'string') {
        mergedRanges.push(range);
      } else if (range.model) {
        // ExcelJS structure: {model: 'A1:B2'}
        mergedRanges.push(range.model);
      }
    }

    // Helper to check if cell is part of merged range
    function getMergeInfo(rowNum, colNum) {
      for (const range of mergedRanges) {
        const rangeStr = typeof range === 'string' ? range : String(range);
        const match = rangeStr.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        if (match) {
          const startCol = columnLetterToNumber(match[1]);
          const startRow = parseInt(match[2]);
          const endCol = columnLetterToNumber(match[3]);
          const endRow = parseInt(match[4]);
          
          if (rowNum >= startRow && rowNum <= endRow && 
              colNum >= startCol && colNum <= endCol) {
            return {
              isMerged: true,
              isTopLeft: rowNum === startRow && colNum === startCol,
              colSpan: endCol - startCol + 1,
              rowSpan: endRow - startRow + 1
            };
          }
        }
      }
      return { isMerged: false };
    }

    // Define range A1:O49
    const minRow = 1;
    const maxRow = 49;
    const minCol = 1; // Column A
    const maxCol = 15; // Column O

    // Calculate column widths
    const colWidths = {};
    for (let colNum = minCol; colNum <= maxCol; colNum++) {
      const column = worksheet.getColumn(colNum);
      const width = column.width || 10;
      colWidths[colNum] = width * 5.5; // Adjusted for A4 portrait
    }

    // Calculate available width and scale if needed
    const availableWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
    const totalColWidth = Object.values(colWidths).reduce((a, b) => a + b, 0);
    const scale = Math.min(1, availableWidth / totalColWidth);
    
    Object.keys(colWidths).forEach(col => {
      colWidths[col] *= scale;
    });

    // Dynamic row heights based on content
    const rowHeights = {};
    for (let rowNum = minRow; rowNum <= maxRow; rowNum++) {
      const row = worksheet.getRow(rowNum);
      const height = row.height || 15;
      rowHeights[rowNum] = Math.max(height * 1.2, 15); // Minimum 15pt
    }

    let currentY = doc.page.margins.top;
    const pageHeight = doc.page.height - doc.page.margins.bottom;

    // Extract images
    const images = {};
    if (worksheet.getImages) {
      worksheet.getImages().forEach(img => {
        const imageId = img.imageId;
        const image = workbook.model.media.find(m => m.index === imageId);
        if (image) {
          images[img.range.tl.nativeRow] = {
            buffer: image.buffer,
            extension: image.extension,
            row: img.range.tl.nativeRow,
            col: img.range.tl.nativeCol
          };
        }
      });
    }

    // Render each row
    for (let rowNum = minRow; rowNum <= maxRow; rowNum++) {
      const row = worksheet.getRow(rowNum);
      const rowHeight = rowHeights[rowNum];

      // Check if we need a new page
      if (currentY + rowHeight > pageHeight) {
        doc.addPage();
        currentY = doc.page.margins.top;
      }

      let currentX = doc.page.margins.left;

      // Render each cell in the row
      for (let colNum = minCol; colNum <= maxCol; colNum++) {
        const cell = row.getCell(colNum);
        const cellWidth = colWidths[colNum];
        const mergeInfo = getMergeInfo(rowNum, colNum);

        // Skip if this is a merged cell but not the top-left
        if (mergeInfo.isMerged && !mergeInfo.isTopLeft) {
          currentX += cellWidth;
          continue;
        }

        const cellX = currentX;
        const cellY = currentY;
        const finalWidth = mergeInfo.isMerged ? 
          Array.from({length: mergeInfo.colSpan}, (_, i) => colWidths[colNum + i]).reduce((a,b) => a+b, 0) : 
          cellWidth;
        const finalHeight = mergeInfo.isMerged ? 
          Array.from({length: mergeInfo.rowSpan}, (_, i) => rowHeights[rowNum + i]).reduce((a,b) => a+b, 0) : 
          rowHeight;

        // Draw cell background
        if (cell.style && cell.style.fill && cell.style.fill.fgColor) {
          const color = cell.style.fill.fgColor.argb;
          if (color && color !== 'FFFFFFFF' && color !== '00000000') {
            const rgb = hexToRgb(color);
            if (rgb) {
              doc.save();
              doc.rect(cellX, cellY, finalWidth, finalHeight).fill(rgb);
              doc.restore();
            }
          }
        }

        // Draw cell border
        const borderColor = '#000000';
        doc.strokeColor(borderColor).lineWidth(0.5);
        doc.rect(cellX, cellY, finalWidth, finalHeight).stroke();

        // Check for image in this cell
        const imgData = images[rowNum - 1]; // 0-indexed
        if (imgData && imgData.col === colNum - 1) {
          try {
            const imgWidth = finalWidth - 4;
            const imgHeight = finalHeight - 4;
            doc.image(imgData.buffer, cellX + 2, cellY + 2, {
              fit: [imgWidth, imgHeight],
              align: 'center',
              valign: 'center'
            });
          } catch (err) {
            console.error('Image rendering error:', err);
          }
        }

        // Draw cell text
        let cellValue = '';
        if (cell.value !== null && cell.value !== undefined) {
          if (typeof cell.value === 'object' && cell.value.text) {
            cellValue = cell.value.text;
          } else if (typeof cell.value === 'object' && cell.value.result !== undefined) {
            cellValue = String(cell.value.result);
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
          const fontSize = cell.font && cell.font.size ? Math.min(cell.font.size, 11) : 9;
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
          
          const padding = 3;
          const textWidth = finalWidth - (padding * 2);
          let textY = cellY + (finalHeight - fontSize) / 2;
          
          if (verticalAlign === 'top') {
            textY = cellY + padding;
          } else if (verticalAlign === 'bottom') {
            textY = cellY + finalHeight - fontSize - padding;
          }

          // Handle checkmarks and special characters
          const displayValue = cellValue === 'true' ? '✓' : cellValue;

          doc.text(displayValue, cellX + padding, textY, {
            width: textWidth,
            align: align,
            lineBreak: false,
            ellipsis: true
          });
        }

        currentX += cellWidth;
      }

      currentY += rowHeight;
    }

    doc.end();
  });
}

function columnLetterToNumber(letter) {
  let column = 0;
  for (let i = 0; i < letter.length; i++) {
    column = column * 26 + letter.charCodeAt(i) - 64;
  }
  return column;
}

function hexToRgb(hex) {
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
