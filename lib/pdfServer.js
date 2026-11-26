import { PDFDocument, rgb, degrees } from 'pdf-lib';
import { LABEL_FORMATS, DEFAULT_CATEGORIES, resolveLabelFormat } from './catalog.js';

const MM_TO_PT = 72 / 25.4;

function mmToPt(mm) {
  return mm * MM_TO_PT;
}

function parseDataUrl(dataUrl) {
  if (typeof dataUrl !== 'string') {
    throw new Error('Imagem de etiqueta ausente.');
  }
  const match = dataUrl.match(
    /^data:(?:image\/(?:png|jpeg|jpg)|application\/octet-stream);base64,([A-Za-z0-9+/=\s]+)$/i
  );
  if (!match) {
    throw new Error('Formato de imagem inválido (somente PNG ou JPEG).');
  }
  const bytes = Buffer.from(match[1].replace(/\s+/g, ''), 'base64');
  if (bytes.length >= 4 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) {
    return { mime: 'image/png', bytes };
  }
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) {
    return { mime: 'image/jpeg', bytes };
  }
  throw new Error('Conteúdo de imagem inválido (assinatura PNG ou JPEG não reconhecida).');
}

function drawCutGuides(page, spec, pageH_pt) {
  const strokeColor = rgb(200 / 255, 200 / 255, 200 / 255);
  const lineW_pt = mmToPt(0.15);

  if (spec.gapX === 0 && spec.gapY === 0) {
    const xStart = spec.marginX;
    const xEnd = spec.marginX + spec.cols * spec.labelW;
    const yStart = spec.marginY;
    const yEnd = spec.marginY + spec.rows * spec.labelH;

    for (let c = 0; c <= spec.cols; c++) {
      const xMm = spec.marginX + c * spec.labelW;
      drawVLine(page, xMm, yStart, yEnd, pageH_pt, strokeColor, lineW_pt);
    }
    for (let r = 0; r <= spec.rows; r++) {
      const yMm = spec.marginY + r * spec.labelH;
      drawHLine(page, xStart, xEnd, yMm, pageH_pt, strokeColor, lineW_pt);
    }
    return;
  }

  for (let r = 0; r < spec.rows; r++) {
    for (let c = 0; c < spec.cols; c++) {
      const xMm = spec.marginX + c * (spec.labelW + spec.gapX);
      const yTopMm = spec.marginY + r * (spec.labelH + spec.gapY);
      page.drawRectangle({
        x: mmToPt(xMm),
        y: pageH_pt - mmToPt(yTopMm + spec.labelH),
        width: mmToPt(spec.labelW),
        height: mmToPt(spec.labelH),
        borderColor: strokeColor,
        borderWidth: lineW_pt
      });
    }
  }
}

function drawVLine(page, xMm, y1TopMm, y2TopMm, pageH_pt, color, thickness) {
  const x = mmToPt(xMm);
  page.drawLine({
    start: { x, y: pageH_pt - mmToPt(y1TopMm) },
    end: { x, y: pageH_pt - mmToPt(y2TopMm) },
    thickness,
    color
  });
}

function drawHLine(page, x1Mm, x2Mm, yTopMm, pageH_pt, color, thickness) {
  const y = pageH_pt - mmToPt(yTopMm);
  page.drawLine({
    start: { x: mmToPt(x1Mm), y },
    end: { x: mmToPt(x2Mm), y },
    thickness,
    color
  });
}

/**
 * Generates an A4 Landscape PDF (297 x 210 mm) matching the client checkout layout:
 * - 6x3 cm: 4 cols x 6 rows = 24 labels/sheet
 * - 10x4 cm (Sucos): 7 cols x 2 rows = 14 labels/sheet rotated 90° clockwise, edge-to-edge (gap 0)
 */
export async function buildOrderPdfBuffer(items, categories = DEFAULT_CATEGORIES) {
  if (!Array.isArray(items) || items.length === 0) {
    throw new Error('Nenhum item informado para gerar o PDF.');
  }

  const pdfDoc = await PDFDocument.create();
  const embeddedCache = new Map();

  for (const item of items) {
    const { label, sheets } = item;
    if (!label || !sheets || sheets <= 0) continue;

    const formatKey = item.format || resolveLabelFormat(label, categories);
    const spec = LABEL_FORMATS[formatKey] || LABEL_FORMATS['6x3'];
    const pageW_pt = mmToPt(spec.pageW);
    const pageH_pt = mmToPt(spec.pageH);

    let embeddedImg = embeddedCache.get(label.id || label.dataUrl);
    if (!embeddedImg) {
      const { mime, bytes } = parseDataUrl(label.dataUrl);
      embeddedImg = mime === 'image/png' ? await pdfDoc.embedPng(bytes) : await pdfDoc.embedJpg(bytes);
      embeddedCache.set(label.id || label.dataUrl, embeddedImg);
    }

    for (let s = 0; s < sheets; s++) {
      const page = pdfDoc.addPage([pageW_pt, pageH_pt]);

      for (let r = 0; r < spec.rows; r++) {
        for (let c = 0; c < spec.cols; c++) {
          const xMm = spec.marginX + c * (spec.labelW + spec.gapX);
          const yTopMm = spec.marginY + r * (spec.labelH + spec.gapY);

          if (spec.rotate90) {
            // Rotate 100x40 image -90 deg (clockwise) around top-left of the 40x100 cell
            page.drawImage(embeddedImg, {
              x: mmToPt(xMm),
              y: pageH_pt - mmToPt(yTopMm),
              width: mmToPt(spec.labelH),
              height: mmToPt(spec.labelW),
              rotate: degrees(-90)
            });
          } else {
            page.drawImage(embeddedImg, {
              x: mmToPt(xMm),
              y: pageH_pt - mmToPt(yTopMm + spec.labelH),
              width: mmToPt(spec.labelW),
              height: mmToPt(spec.labelH)
            });
          }
        }
      }

      // Draw cutting grid lines ON TOP of labels so borders remain visible across all cells
      drawCutGuides(page, spec, pageH_pt);
    }
  }

  if (pdfDoc.getPageCount() === 0) {
    throw new Error('Nenhuma página válida gerada para o pedido.');
  }

  const pdfBytes = await pdfDoc.save();
  return Buffer.from(pdfBytes);
}
