'use strict';
// Writes a small fillable PDF form for tests: node tests/make-form-pdf.js <out.pdf>
const fs = require('fs');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');

(async () => {
  const doc = await PDFDocument.create();
  const page = doc.addPage([612, 792]);
  const font = await doc.embedFont(StandardFonts.Helvetica);
  page.drawText('Membership form', { x: 72, y: 720, size: 22, font, color: rgb(0.1, 0.2, 0.5) });
  page.drawText('Name:', { x: 72, y: 660, size: 12, font });
  page.drawText('Plan:', { x: 72, y: 620, size: 12, font });
  page.drawText('Newsletter:', { x: 72, y: 580, size: 12, font });
  page.drawText('Contact by:', { x: 72, y: 540, size: 12, font });
  const form = doc.getForm();
  const name = form.createTextField('name');
  name.addToPage(page, { x: 160, y: 652, width: 250, height: 22 });
  const plan = form.createDropdown('plan');
  plan.addOptions(['Basic', 'Plus', 'Pro']);
  plan.addToPage(page, { x: 160, y: 612, width: 120, height: 22 });
  const news = form.createCheckBox('newsletter');
  news.addToPage(page, { x: 160, y: 576, width: 16, height: 16 });
  const contact = form.createRadioGroup('contact');
  contact.addOptionToPage('email', page, { x: 160, y: 536, width: 16, height: 16 });
  contact.addOptionToPage('phone', page, { x: 240, y: 536, width: 16, height: 16 });
  page.drawText('Email', { x: 182, y: 540, size: 11, font });
  page.drawText('Phone', { x: 262, y: 540, size: 11, font });
  fs.writeFileSync(process.argv[2], await doc.save());
  console.log('form written');
})();
