// Netlify Function: generate-pdf.js
// Place in: netlify/functions/generate-pdf.js
// npm install pdfkit (add to package.json)

const PDFDocument = require('pdfkit');

const DKGREEN = '#1A5C3A';
const SAGE    = '#78BEA0';
const PALE    = '#E8F5EF';
const WHITE   = '#FFFFFF';
const GREY    = '#6B6B6B';
const BLACK   = '#1A1A1A';
const MAROON  = '#3C141E';

function hexToRgb(hex) {
  const r = parseInt(hex.slice(1,3),16);
  const g = parseInt(hex.slice(3,5),16);
  const b = parseInt(hex.slice(5,7),16);
  return [r,g,b];
}

function buildPDF(sections, offer) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ 
      size: 'A4', 
      margins: { top: 72, bottom: 72, left: 56, right: 56 },
      info: { Title: 'Pickwick Estates Document', Author: 'Pickwick Estates' }
    });

    const chunks = [];
    doc.on('data', chunk => chunks.push(chunk));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    const W = doc.page.width;
    const LM = doc.page.margins.left;
    const RM = doc.page.width - doc.page.margins.right;
    const CW = RM - LM;

    function addHeader() {
      // Green header bar
      doc.save();
      doc.rect(0, 0, W, 60).fill(DKGREEN);
      doc.rect(0, 60, W, 4).fill(SAGE);
      
      // Logo text
      doc.fillColor(WHITE).fontSize(14).font('Helvetica-Bold');
      doc.text('PICKWICK ESTATES', LM, 18);
      doc.fillColor(SAGE).fontSize(8).font('Helvetica');
      doc.text('ALPS ESTATES LTD  |  LETTINGS & PROPERTY MANAGEMENT', LM, 36);
      
      // Right side contact
      doc.fillColor(WHITE).fontSize(8).font('Helvetica');
      doc.text('maintenance@pickwickestates.com', LM, 18, { align: 'right', width: CW });
      doc.text('pickwickestates.co.uk', LM, 30, { align: 'right', width: CW });
      doc.restore();
      
      doc.y = 80;
    }

    function addFooter() {
      const footerY = doc.page.height - 40;
      doc.save();
      doc.rect(0, footerY - 4, W, 1).fill(SAGE);
      doc.rect(0, footerY - 3, W, 37).fill(DKGREEN);
      doc.fillColor(WHITE).fontSize(7).font('Helvetica');
      doc.text(
        'Pickwick Estates  |  Alps Estates Ltd  |  Member of The Property Redress Scheme  |  Propertymark CMP',
        LM, footerY + 8, { align: 'center', width: CW }
      );
      doc.restore();
    }

    function checkPageBreak(neededHeight) {
      const remaining = doc.page.height - doc.page.margins.bottom - doc.y;
      if (remaining < neededHeight) {
        addFooter();
        doc.addPage();
        addHeader();
      }
    }

    // First page header
    addHeader();

    // Process sections
    sections.forEach(sec => {
      if (sec.h) {
        checkPageBreak(40);
        doc.moveDown(0.5);
        doc.save();
        const y = doc.y;
        doc.rect(LM, y, CW, 26).fill(DKGREEN);
        doc.fillColor(WHITE).fontSize(10).font('Helvetica-Bold');
        doc.text(sec.h.toUpperCase(), LM + 8, y + 8, { width: CW - 16 });
        doc.restore();
        doc.y = y + 30;
        doc.moveDown(0.3);

      } else if (sec.h2) {
        checkPageBreak(30);
        doc.moveDown(0.6);
        doc.fillColor(DKGREEN).fontSize(10).font('Helvetica-Bold');
        doc.text(sec.h2, LM, doc.y, { width: CW });
        doc.moveDown(0.1);
        doc.save();
        doc.rect(LM, doc.y, CW, 1.5).fill(SAGE);
        doc.restore();
        doc.moveDown(0.4);

      } else if (sec.sub) {
        checkPageBreak(18);
        doc.fillColor(GREY).fontSize(9).font('Helvetica');
        doc.text(sec.sub, LM, doc.y, { width: CW });
        doc.moveDown(0.2);

      } else if (sec.body) {
        const text = (sec.body || '').replace(/GBP /g, '£');
        const lines = doc.heightOfString(text, { width: CW, fontSize: 9 });
        checkPageBreak(lines + 10);
        doc.fillColor(BLACK).fontSize(9).font('Helvetica');
        doc.text(text, LM, doc.y, { width: CW, lineGap: 3 });
        doc.moveDown(0.5);

      } else if (sec.table) {
        checkPageBreak(sec.table.length * 22 + 10);
        const col1W = CW * 0.38;
        const col2W = CW - col1W;
        sec.table.forEach((row, i) => {
          const rowY = doc.y;
          // Col 1 - label
          doc.save();
          doc.rect(LM, rowY, col1W, 20).fill(i % 2 === 0 ? PALE : '#DCF0E6');
          doc.restore();
          doc.fillColor(DKGREEN).fontSize(9).font('Helvetica-Bold');
          doc.text(row[0] || '', LM + 4, rowY + 5, { width: col1W - 8 });
          // Col 2 - value
          doc.save();
          doc.rect(LM + col1W, rowY, col2W, 20).fill(WHITE);
          doc.restore();
          doc.fillColor(BLACK).fontSize(9).font('Helvetica');
          doc.text(row[1] || '', LM + col1W + 4, rowY + 5, { width: col2W - 8 });
          // Border
          doc.save();
          doc.rect(LM, rowY, CW, 20).stroke('#D4E8DC');
          doc.restore();
          doc.y = rowY + 20;
        });
        doc.moveDown(0.5);

      } else if (sec.sig) {
        checkPageBreak(80);
        doc.moveDown(0.5);
        doc.save();
        doc.rect(LM, doc.y, CW, 70).fill(PALE);
        doc.restore();
        const sigY = doc.y;
        doc.fillColor(DKGREEN).fontSize(10).font('Helvetica-Bold');
        doc.text(sec.sig || 'Signature', LM + 8, sigY + 8);
        doc.fillColor(GREY).fontSize(9).font('Helvetica');
        doc.text('Name: ' + (sec.name || ''), LM + 8, sigY + 22);
        doc.text('Date: ' + (sec.date || ''), LM + 8, sigY + 34);
        if (offer && offer.gdprTimestamp) {
          doc.fillColor(GREY).fontSize(8);
          doc.text('Electronically signed: ' + new Date(offer.gdprTimestamp).toLocaleString('en-GB'), LM + 8, sigY + 46);
        }
        doc.save();
        doc.rect(LM, sigY, CW, 70).stroke(SAGE);
        doc.restore();
        doc.y = sigY + 74;
        doc.moveDown(0.5);

      } else if (sec.footer) {
        checkPageBreak(30);
        doc.moveDown(1);
        doc.save();
        doc.rect(LM, doc.y, CW, 1).fill(SAGE);
        doc.restore();
        doc.moveDown(0.3);
        doc.fillColor(GREY).fontSize(8).font('Helvetica');
        doc.text(sec.footer || '', LM, doc.y, { align: 'center', width: CW });
      }
    });

    addFooter();
    doc.end();
  });
}

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method not allowed' };
  }

  try {
    const body = JSON.parse(event.body);
    const { sections, filename, offer } = body;

    if (!sections || !Array.isArray(sections)) {
      return { statusCode: 400, body: 'Invalid sections data' };
    }

    const pdfBuffer = await buildPDF(sections, offer || {});
    const base64 = pdfBuffer.toString('base64');

    return {
      statusCode: 200,
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${filename || 'document'}.pdf"`,
        'Access-Control-Allow-Origin': '*',
      },
      body: base64,
      isBase64Encoded: true,
    };

  } catch (err) {
    console.error('PDF generation error:', err);
    return { statusCode: 500, body: 'PDF generation failed: ' + err.message };
  }
};
