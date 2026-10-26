'use strict';
// Writes a 3-page text PDF with embedded Windows fonts (Arial, Times New Roman) and one rotated page, for tests.
// node tests/make-text-pdf.js <out.pdf>
const fs = require('fs');
const path = require('path');
const { PDFDocument, rgb, degrees } = require('pdf-lib');
const fontkit = require('@pdf-lib/fontkit');

(async () => {
  const doc = await PDFDocument.create();
  doc.registerFontkit(fontkit);
  const fonts = path.join(process.env.WINDIR || 'C:\\Windows', 'Fonts');
  const arial = await doc.embedFont(fs.readFileSync(path.join(fonts, 'arial.ttf')), { subset: true });
  const arialBold = await doc.embedFont(fs.readFileSync(path.join(fonts, 'arialbd.ttf')), { subset: true });
  const times = await doc.embedFont(fs.readFileSync(path.join(fonts, 'times.ttf')), { subset: true });
  for (let n = 1; n <= 3; n++) {
    const page = doc.addPage([612, 792]);
    page.drawText(`Quarterly Report - page ${n}`, { x: 72, y: 720, size: 20, font: arialBold, color: rgb(0.1, 0.25, 0.55) });
    page.drawText('Revenue grew strongly this quarter across all regions.', { x: 72, y: 680, size: 12, font: arial });
    page.drawText('The board approved the new budget on Monday.', { x: 72, y: 660, size: 12, font: times });
    for (let k = 0; k < 20; k++) page.drawText(`Line ${k + 1}: lorem ipsum dolor sit amet, consectetur adipiscing elit.`, { x: 72, y: 620 - k * 18, size: 11, font: arial });
    if (n === 3) page.setRotation(degrees(90));
  }
  doc.setTitle('Doc Studio PDF test');
  fs.writeFileSync(process.argv[2], await doc.save());
  console.log('text pdf written');
})();
