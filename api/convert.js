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
    // Force A4 Portrait to match your checklist
    const doc = new PDFDocument({ 
      size: 'A4',
      margin: 15,
      layout: 'portrait',
      bufferPages: true
    });
    
    const chunks = [];
    doc.on('data', chunk => chunks.push(chunk));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    // Define print range - force A1:O49 for your checklist
    const minRow = 1;
    const maxRow = 49;
    const minCol = 1;
    const maxCol = 15; // Column O

    // Build merged cells map
    const mergedCellsMap = new Map();
    const processedMerges = new Set();
    
    if (worksheet._merges) {
      Object.values(worksheet._merges).forEach(merge => {
        const rangeStr = typeof merge === 'string' ? merge : (merge.model || merge);
        const match = String(rangeStr).match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        if (match) {
          const startCol = columnLetterToNumber(match[1]);
          const startRow = parseInt(match[2]);
          const endCol = columnLetterToNumber(match[3]);
          const endRow = parseInt(match[4]);
          
          for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
              mergedCellsMap.set(`${r}-${c}`, {
                isTopLeft: r === startRow && c === startCol,
                startRow, startCol, endRow, endCol,
                rowSpan: endRow - startRow + 1,
                colSpan: endCol - startCol + 1
              });
            }
          }
        }
      });
    }

    // Calculate column widths
    const colWidths = {};
    for (let colNum = minCol; colNum <= maxCol; colNum++) {
      const column = worksheet.getColumn(colNum);
      colWidths[colNum] = (column.width || 10) * 4.8;
    }

    // Scale to fit page width
    const availableWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
    const totalColWidth = Object.values(colWidths).reduce((a, b) => a + b, 0);
    const widthScale = availableWidth / totalColWidth;
    
    Object.keys(colWidths).forEach(col => {
      colWidths[col] *= widthScale;
    });

    // Calculate row heights
    const rowHeights = {};
    for (let rowNum = minRow; rowNum <= maxRow; rowNum++) {
      const row = worksheet.getRow(rowNum);
      rowHeights[rowNum] = (row.height || 15) * 0.95;
    }

    // FORCE FIT TO ONE PAGE - Scale height aggressively
    const totalHeight = Object.values(rowHeights).reduce((a, b) => a + b, 0);
    const availableHeight = doc.page.height - doc.page.margins.top - doc.page.margins.bottom;
    const heightScale = availableHeight / totalHeight;
    
    Object.keys(rowHeights).forEach(row => {
      rowHeights[row] *= heightScale;
    });

    // Extract images
    const images = new Map();
    if (worksheet.getImages) {
      worksheet.getImages().forEach(img => {
        try {
          const imageId = img.imageId;
          const image = workbook.model.media.find(m => m.index === imageId);
          if (image && image.buffer) {
            const row = img.range.tl.nativeRow + 1;
            const col = img.range.tl.nativeCol + 1;
            const rowEnd = img.range.br ? img.range.br.nativeRow + 1 : row;
            const colEnd = img.range.br ? img.range.br.nativeCol + 1 : col;
            
            images.set(`${row}-${col}`, {
              buffer: image.buffer,
              extension: image.extension,
              rowSpan: rowEnd - row + 1,
              colSpan: colEnd - col + 1
            });
          }
        } catch (err) {
          console.error('Image extraction error:', err);
        }
      });
    }

    let currentY = doc.page.margins.top;

    // Render all rows
    for (let rowNum = minRow; rowNum <= maxRow; rowNum++) {
      const row = worksheet.getRow(rowNum);
      const rowHeight = rowHeights[rowNum];
      let currentX = doc.page.margins.left;

      for (let colNum = minCol; colNum <= maxCol; colNum++) {
        const cellKey = `${rowNum}-${colNum}`;
        const mergeInfo = mergedCellsMap.get(cellKey);
        const cell = row.getCell(colNum);
        const cellWidth = colWidths[colNum];

        // Skip if merged but not top-left
        if (mergeInfo && !mergeInfo.isTopLeft) {
          currentX += cellWidth;
          continue;
        }

        const cellX = currentX;
        const cellY = currentY;
        
        // Calculate final dimensions for merged cells
        let finalWidth = cellWidth;
        let finalHeight = rowHeight;
        
        if (mergeInfo && mergeInfo.isTopLeft) {
          finalWidth = 0;
          for (let c = mergeInfo.startCol; c <= mergeInfo.endCol; c++) {
            finalWidth += colWidths[c] || 0;
          }
          finalHeight = 0;
          for (let r = mergeInfo.startRow; r <= mergeInfo.endRow; r++) {
            finalHeight += rowHeights[r] || 0;
          }
        }

        // Draw background color
        if (cell.style?.fill?.fgColor?.argb) {
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

        // Draw border
        doc.strokeColor('#000000').lineWidth(0.3);
        doc.rect(cellX, cellY, finalWidth, finalHeight).stroke();

        // Draw image if exists
        const imgData = images.get(cellKey);
        if (imgData) {
          try {
            const imgWidth = finalWidth - 2;
            const imgHeight = finalHeight - 2;
            doc.image(imgData.buffer, cellX + 1, cellY + 1, {
              fit: [imgWidth, imgHeight],
              align: 'center',
              valign: 'center'
            });
          } catch (err) {
            console.error('Image render error:', err);
          }
        }

        // Draw text content
        if (!imgData) {
          let cellValue = '';
          
          if (cell.value !== null && cell.value !== undefined) {
            if (typeof cell.value === 'boolean') {
              cellValue = cell.value ? '✓' : '';
            } else if (typeof cell.value === 'object' && cell.value.text) {
              cellValue = cell.value.text;
            } else if (typeof cell.value === 'object' && cell.value.result !== undefined) {
              cellValue = String(cell.value.result);
            } else {
              cellValue = String(cell.value);
            }
          }

          if (cellValue) {
            // Handle checkmark symbol
            if (cellValue === 'TRUE' || cellValue === 'true' || cellValue === '1') {
              cellValue = '✓';
            }
            
            // Font styling
            let fontStyle = 'Helvetica';
            if (cell.font?.bold && cell.font?.italic) {
              fontStyle = 'Helvetica-BoldOblique';
            } else if (cell.font?.bold) {
              fontStyle = 'Helvetica-Bold';
            } else if (cell.font?.italic) {
              fontStyle = 'Helvetica-Oblique';
            }
            doc.font(fontStyle);

            // Font size - scaled down
            const baseFontSize = cell.font?.size || 9;
            const fontSize = Math.max(Math.min(baseFontSize * heightScale, 9), 6);
            doc.fontSize(fontSize);

            // Font color
            const textColor = cell.font?.color?.argb ? hexToRgb(cell.font.color.argb) : null;
            doc.fillColor(textColor || '#000000');

            // Text alignment
            const align = cell.alignment?.horizontal || 'left';
            const vAlign = cell.alignment?.vertical || 'middle';
            
            const padding = 1.5;
            const textWidth = finalWidth - (padding * 2);
            
            let textY = cellY + (finalHeight - fontSize) / 2;
            if (vAlign === 'top') {
              textY = cellY + padding;
            } else if (vAlign === 'bottom') {
              textY = cellY + finalHeight - fontSize - padding;
            }

            // Handle text wrapping for merged cells or long text
            const shouldWrap = mergeInfo || cellValue.length > 30;
            
            try {
              doc.text(cellValue, cellX + padding, textY, {
                width: textWidth,
                align: align,
                lineBreak: shouldWrap,
                ellipsis: !shouldWrap,
                height: finalHeight - (padding * 2)
              });
            } catch (err) {
              // Fallback for Arabic or special characters
              doc.text(cellValue.replace(/[^\x00-\x7F]/g, '?'), cellX + padding, textY, {
                width: textWidth,
                align: align,
                lineBreak: false,
                ellipsis: true
              });
            }
          }
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
