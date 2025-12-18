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
    // Check for print area
    let minRow = 1, maxRow = 49, minCol = 1, maxCol = 15;
    
    if (worksheet.pageSetup && worksheet.pageSetup.printArea) {
      const printArea = worksheet.pageSetup.printArea;
      const match = printArea.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
      if (match) {
        minCol = columnLetterToNumber(match[1]);
        minRow = parseInt(match[2]);
        maxCol = columnLetterToNumber(match[3]);
        maxRow = parseInt(match[4]);
      }
    }

    // Detect actual data boundaries if no print area
    if (!worksheet.pageSetup || !worksheet.pageSetup.printArea) {
      let hasData = false;
      worksheet.eachRow((row, rowNum) => {
        row.eachCell({ includeEmpty: false }, (cell, colNum) => {
          if (cell.value !== null && cell.value !== undefined && cell.value !== '') {
            hasData = true;
            maxRow = Math.max(maxRow, rowNum);
            maxCol = Math.max(maxCol, colNum);
          }
        });
      });
      
      if (hasData) {
        // Add small buffer
        maxRow = Math.min(maxRow + 1, 49);
        maxCol = Math.min(maxCol, 15);
      }
    }

    // Determine page orientation from Excel settings
    const orientation = worksheet.pageSetup && worksheet.pageSetup.orientation === 'landscape' ? 'landscape' : 'portrait';
    
    const doc = new PDFDocument({ 
      size: 'A4',
      margin: 20,
      layout: orientation,
      bufferPages: true
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
              startCol: startCol,
              startRow: startRow,
              colSpan: endCol - startCol + 1,
              rowSpan: endRow - startRow + 1
            };
          }
        }
      }
      return { isMerged: false };
    }

    // Calculate column widths
    const colWidths = {};
    for (let colNum = minCol; colNum <= maxCol; colNum++) {
      const column = worksheet.getColumn(colNum);
      const width = column.width || 10;
      colWidths[colNum] = width * 5.2;
    }

    // Scale to fit page width
    const availableWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
    const totalColWidth = Object.values(colWidths).reduce((a, b) => a + b, 0);
    const scale = availableWidth / totalColWidth;
    
    Object.keys(colWidths).forEach(col => {
      colWidths[col] *= scale;
    });

    // Calculate row heights
    const rowHeights = {};
    for (let rowNum = minRow; rowNum <= maxRow; rowNum++) {
      const row = worksheet.getRow(rowNum);
      const height = row.height || 15;
      rowHeights[rowNum] = height * 1.1;
    }

    // Calculate total height and scale if needed to fit one page
    const totalHeight = Object.values(rowHeights).reduce((a, b) => a + b, 0);
    const availableHeight = doc.page.height - doc.page.margins.top - doc.page.margins.bottom;
    
    let heightScale = 1;
    if (totalHeight > availableHeight) {
      heightScale = availableHeight / totalHeight;
      Object.keys(rowHeights).forEach(row => {
        rowHeights[row] *= heightScale;
      });
    }

    // Extract images
    const images = {};
    if (worksheet.getImages) {
      worksheet.getImages().forEach(img => {
        const imageId = img.imageId;
        const image = workbook.model.media.find(m => m.index === imageId);
        if (image) {
          const row = img.range.tl.nativeRow + 1; // Convert to 1-based
          const col = img.range.tl.nativeCol + 1;
          images[`${row}-${col}`] = {
            buffer: image.buffer,
            extension: image.extension
          };
        }
      });
    }

    let currentY = doc.page.margins.top;

    // Render each row
    for (let rowNum = minRow; rowNum <= maxRow; rowNum++) {
      const row = worksheet.getRow(rowNum);
      const rowHeight = rowHeights[rowNum];
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
        
        let finalWidth = cellWidth;
        let finalHeight = rowHeight;
        
        if (mergeInfo.isMerged) {
          finalWidth = 0;
          for (let i = 0; i < mergeInfo.colSpan; i++) {
            finalWidth += colWidths[mergeInfo.startCol + i] || 0;
          }
          finalHeight = 0;
          for (let i = 0; i < mergeInfo.rowSpan; i++) {
            finalHeight += rowHeights[mergeInfo.startRow + i] || 0;
          }
        }

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
        const imgKey = `${rowNum}-${colNum}`;
        if (images[imgKey]) {
          try {
            const imgWidth = finalWidth - 4;
            const imgHeight = finalHeight - 4;
            doc.image(images[imgKey].buffer, cellX + 2, cellY + 2, {
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
          } else if (cell.value === true) {
            cellValue = '✓';
          } else {
            cellValue = String(cell.value);
          }
        }

        if (cellValue && !images[imgKey]) {
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

          // Set font size (scaled)
          let fontSize = cell.font && cell.font.size ? cell.font.size : 9;
          fontSize = Math.max(fontSize * heightScale, 7); // Minimum 7pt
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
          
          const padding = 2;
          const textWidth = finalWidth - (padding * 2);
          let textY = cellY + (finalHeight - fontSize) / 2;
          
          if (verticalAlign === 'top') {
            textY = cellY + padding;
          } else if (verticalAlign === 'bottom') {
            textY = cellY + finalHeight - fontSize - padding;
          }

          doc.text(cellValue, cellX + padding, textY, {
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
