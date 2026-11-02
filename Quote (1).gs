/**
 * Quotation / Invoice generator (Google Doc + PDF saved in Drive).
 *
 * Optional Script Properties (Project Settings):
 *   COMPANY_NAME, COMPANY_PHONE, COMPANY_ADDRESS
 *   ABA_INFO   e.g. "ABA 000 123 456 - Your Name"
 */

const PDF_FOLDER = 'Customer Tracker PDFs';
const KH_FONT = 'Noto Sans Khmer';

function folder_() {
  const it = DriveApp.getFoldersByName(PDF_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PDF_FOLDER);
}

function docsSheet_() {
  let s = ss_().getSheetByName('Documents');
  if (!s) {
    s = ss_().insertSheet('Documents');
    s.appendRow(['Time', 'ID', 'Customer', 'Type', 'Number', 'PDF', 'Doc']);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, 7).setFontWeight('bold');
  }
  return s;
}

function cell_(row, i, text, bold, align) {
  const c = row.getCell(i);
  c.setText(text);
  c.setPaddingTop(5).setPaddingBottom(5);
  const t = c.editAsText();
  t.setBold(!!bold);
  if (align) c.getChild(0).asParagraph().setAlignment(align);
}

function makeDocument(id, type) {
  type = type === 'Invoice' ? 'Invoice' : 'Quotation';
  const c = getAll().find(x => x.ID === id);
  if (!c) throw new Error('Customer not found');

  const p = props_();
  const company = p.getProperty('COMPANY_NAME') || 'Your Company';
  const phone = p.getProperty('COMPANY_PHONE') || '';
  const address = p.getProperty('COMPANY_ADDRESS') || '';
  const aba = p.getProperty('ABA_INFO') || '';

  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const valid = Utilities.formatDate(new Date(Date.now() + 14 * 864e5), tz, 'yyyy-MM-dd');
  const prefix = type === 'Invoice' ? 'INV-' : 'QT-';
  const no = prefix + c.ID.replace(/^C/, '');
  const title = type === 'Invoice' ? 'វិក្កយបត្រ / INVOICE' : 'ប័ណ្ណសម្រង់តម្លៃ / QUOTATION';
  const amount = '$' + Number(c.Amount || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  const doc = DocumentApp.create(no + ' ' + c.Customer);
  const body = doc.getBody();
  body.setMarginTop(40).setMarginBottom(40).setMarginLeft(50).setMarginRight(50);

  body.appendParagraph(company).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  if (phone || address) body.appendParagraph([phone, address].filter(String).join(' · '));
  body.appendParagraph('');
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('');

  const info = body.appendTable([
    ['លេខ / No.', no, 'កាលបរិច្ឆេទ / Date', today],
    ['អតិថិជន / Customer', c.Customer, type === 'Invoice' ? 'ទូរស័ព្ទ / Phone' : 'មានសុពលភាព / Valid until',
      type === 'Invoice' ? (c.Phone || '-') : valid],
    ['ក្រុមហ៊ុន / Company', c.Company || '-', 'ទូរស័ព្ទ / Phone', c.Phone || '-']
  ]);
  info.setBorderWidth(0.5);

  body.appendParagraph('');
  const items = body.appendTable([
    ['#', 'ពិពណ៌នា / Description', 'ចំនួន / Qty', 'តម្លៃ / Unit price', 'សរុប / Total'],
    ['1', c.Project || '-', '1', amount, amount]
  ]);
  items.setBorderWidth(0.5);
  for (let i = 0; i < 5; i++) items.getRow(0).getCell(i).setBackgroundColor('#e5e7eb').editAsText().setBold(true);

  body.appendParagraph('');
  const tot = body.appendParagraph('សរុប / TOTAL: ' + amount);
  tot.setAlignment(DocumentApp.HorizontalAlignment.RIGHT).editAsText().setBold(true).setFontSize(13);

  if (c.Notes) {
    body.appendParagraph('');
    body.appendParagraph('កំណត់ចំណាំ / Notes: ' + c.Notes);
  }

  body.appendParagraph('');
  if (type === 'Invoice') {
    const pay = c['Payment Method'] === 'ABA'
      ? 'ទូទាត់តាម ABA / Pay via ABA' + (aba ? ': ' + aba : '')
      : c['Payment Method'] === 'Deliver'
        ? 'ទូទាត់ពេលដឹកជញ្ជូន / Pay on delivery'
        : 'វិធីទូទាត់ / Payment: ABA ឬ Deliver';
    body.appendParagraph(pay);
  } else {
    body.appendParagraph('តម្លៃនេះមានសុពលភាព ១៤ ថ្ងៃ / This quotation is valid for 14 days.');
  }

  body.appendParagraph('');
  body.appendParagraph('');
  const sig = body.appendTable([['អតិថិជន / Customer', 'អ្នកលក់ / Seller'], ['\n\n\n', '\n\n\n']]);
  sig.setBorderWidth(0);

  body.editAsText().setFontFamily(KH_FONT);
  doc.saveAndClose();

  const docFile = DriveApp.getFileById(doc.getId());
  const pdfFile = folder_().createFile(docFile.getAs('application/pdf').setName(no + '.pdf'));
  docFile.moveTo(folder_());

  docsSheet_().appendRow([now_(), c.ID, c.Customer, type, no, pdfFile.getUrl(), docFile.getUrl()]);
  return { no: no, pdf: pdfFile.getUrl(), doc: docFile.getUrl() };
}
