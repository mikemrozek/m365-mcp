import { describe, it, expect } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import {
  extractText,
  UnsupportedFormatError,
  DEFAULT_MAX_CHARS,
  MAX_CHARS_LIMIT,
} from '../src/lib/extract-text.js';

/** Minimal but structurally real DOCX: the parts our extractor reads. */
function makeDocx(paragraphs: string[]): Buffer {
  const body = paragraphs
    .map((p) => `<w:p><w:r><w:t>${p}</w:t></w:r></w:p>`)
    .join('');
  const xml = `<?xml version="1.0"?><w:document xmlns:w="x"><w:body>${body}</w:body></w:document>`;
  return Buffer.from(
    zipSync({
      '[Content_Types].xml': strToU8('<Types/>'),
      'word/document.xml': strToU8(xml),
    })
  );
}

/** XLSX using the shared-string table, as Excel actually writes it. */
function makeXlsx(): Buffer {
  const shared =
    `<?xml version="1.0"?><sst xmlns="x" count="3" uniqueCount="3">` +
    `<si><t>Region</t></si><si><t>Revenue</t></si><si><t>Northeast</t></si></sst>`;
  const sheet =
    `<?xml version="1.0"?><worksheet xmlns="x"><sheetData>` +
    `<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>` +
    `<row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>41250</v></c></row>` +
    `</sheetData></worksheet>`;
  const workbook =
    `<?xml version="1.0"?><workbook xmlns="x"><sheets>` +
    `<sheet name="Q3 Summary" sheetId="1" r:id="rId1"/></sheets></workbook>`;
  return Buffer.from(
    zipSync({
      'xl/workbook.xml': strToU8(workbook),
      'xl/sharedStrings.xml': strToU8(shared),
      'xl/worksheets/sheet1.xml': strToU8(sheet),
    })
  );
}

/**
 * Builds a real, structurally valid PDF (xref table and all) with the given
 * lines as an actual text layer. Using a genuine PDF rather than a stub means
 * this test exercises the real pdf pipeline, which is the whole risk in this
 * module — pass `[]` for a page with no text layer, i.e. the scan case.
 */
function makePdf(lines: string[]): Buffer {
  const content = lines.length
    ? `BT /F1 24 Tf 72 700 Td ${lines.map((l, i) => (i ? `0 -30 Td (${l}) Tj ` : `(${l}) Tj `)).join('')}ET`
    : `0.5 0.5 0.5 rg 100 100 200 200 re f`; // a grey box, no text at all
  const objs = [
    `<</Type/Catalog/Pages 2 0 R>>`,
    `<</Type/Pages/Kids[3 0 R]/Count 1>>`,
    `<</Type/Page/Parent 2 0 R/MediaBox[0 0 612 792]/Contents 4 0 R/Resources<</Font<</F1 5 0 R>>>>>>`,
    `<</Length ${content.length}>>\nstream\n${content}\nendstream`,
    `<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>`,
  ];
  let pdf = '%PDF-1.4\n';
  const offsets: number[] = [];
  objs.forEach((body, i) => {
    offsets.push(pdf.length);
    pdf += `${i + 1} 0 obj\n${body}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n`;
  offsets.forEach((o) => {
    pdf += String(o).padStart(10, '0') + ' 00000 n \n';
  });
  pdf += `trailer\n<</Size ${objs.length + 1}/Root 1 0 R>>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(pdf, 'latin1');
}

describe('PDF', () => {
  it('extracts the text layer', async () => {
    const r = await extractText(makePdf(['Invoice Total 41250', 'Vendor: Acme Supply']), 'inv.pdf');
    expect(r.format).toBe('pdf');
    expect(r.text).toContain('Invoice Total 41250');
    expect(r.text).toContain('Vendor: Acme Supply');
  });

  it('reports the page count', async () => {
    const r = await extractText(makePdf(['One line']), 'inv.pdf');
    expect(r.detail?.pages).toBe(1);
  });

  it('detects a scan and says so instead of returning empty text', async () => {
    // A page with graphics but no text layer is what a scanned document is.
    await expect(extractText(makePdf([]), 'scanned.pdf')).rejects.toThrow(
      /no extractable text layer/i
    );
    await expect(extractText(makePdf([]), 'scanned.pdf')).rejects.toThrow(
      UnsupportedFormatError
    );
  });

  it('recognises a PDF from content type alone', async () => {
    const r = await extractText(makePdf(['Body text']), 'attachment', 'application/pdf');
    expect(r.format).toBe('pdf');
    expect(r.text).toContain('Body text');
  });
});

describe('plain-text formats', () => {
  it('reads a text file', async () => {
    const r = await extractText(Buffer.from('hello world'), 'note.txt');
    expect(r.text).toBe('hello world');
    expect(r.format).toBe('txt');
    expect(r.truncated).toBe(false);
  });

  it('reads CSV as-is, preserving structure', async () => {
    const csv = 'name,amount\nAcme,100\nBeta,200';
    const r = await extractText(Buffer.from(csv), 'data.csv');
    expect(r.text).toBe(csv);
  });

  it('strips tags from HTML but keeps the words', async () => {
    const html = '<html><body><p>First para</p><p>Second para</p></body></html>';
    const r = await extractText(Buffer.from(html), 'page.html');
    expect(r.text).toContain('First para');
    expect(r.text).toContain('Second para');
    expect(r.text).not.toContain('<p>');
  });

  it('falls back to content type when the name has no extension', async () => {
    const r = await extractText(Buffer.from('{"a":1}'), 'attachment', 'application/json');
    expect(r.format).toBe('json');
  });
});

describe('DOCX', () => {
  it('extracts paragraphs and keeps them separated', async () => {
    const r = await extractText(makeDocx(['First paragraph', 'Second paragraph']), 'doc.docx');
    expect(r.format).toBe('docx');
    expect(r.text).toContain('First paragraph');
    expect(r.text).toContain('Second paragraph');
    // Paragraphs must not run together — the bug this replaced.
    expect(r.text).not.toContain('First paragraphSecond paragraph');
  });

  it('decodes XML entities', async () => {
    const r = await extractText(makeDocx(['Smith &amp; Sons &lt;test&gt;']), 'doc.docx');
    expect(r.text).toContain('Smith & Sons <test>');
  });

  it('reports a docx with no document part rather than returning empty', async () => {
    const bogus = Buffer.from(zipSync({ 'other/thing.xml': strToU8('<x/>') }));
    await expect(extractText(bogus, 'doc.docx')).rejects.toThrow(UnsupportedFormatError);
  });
});

describe('XLSX', () => {
  it('resolves shared strings and inline numbers into tab-separated rows', async () => {
    const r = await extractText(makeXlsx(), 'book.xlsx');
    expect(r.format).toBe('xlsx');
    expect(r.text).toContain('Region\tRevenue');
    expect(r.text).toContain('Northeast\t41250');
  });

  it('labels output with the real sheet name', async () => {
    const r = await extractText(makeXlsx(), 'book.xlsx');
    expect(r.text).toContain('--- Q3 Summary ---');
    expect(r.detail?.sheets).toEqual(['Q3 Summary']);
  });
});

describe('truncation', () => {
  it('cuts at the cap and reports both the flag and the true length', async () => {
    const long = 'x'.repeat(1000);
    const r = await extractText(Buffer.from(long), 'big.txt', '', 100);
    expect(r.text).toHaveLength(100);
    expect(r.truncated).toBe(true);
    expect(r.totalChars).toBe(1000);
  });

  it('does not flag truncation when the text fits', async () => {
    const r = await extractText(Buffer.from('short'), 'a.txt', '', 100);
    expect(r.truncated).toBe(false);
    expect(r.totalChars).toBe(5);
  });

  it('refuses to exceed the hard ceiling even if asked', async () => {
    const long = 'y'.repeat(MAX_CHARS_LIMIT + 5_000);
    const r = await extractText(Buffer.from(long), 'huge.txt', '', MAX_CHARS_LIMIT * 10);
    expect(r.text.length).toBe(MAX_CHARS_LIMIT);
    expect(r.truncated).toBe(true);
  });

  it('defaults to the documented cap', async () => {
    const long = 'z'.repeat(DEFAULT_MAX_CHARS + 100);
    const r = await extractText(Buffer.from(long), 'd.txt');
    expect(r.text.length).toBe(DEFAULT_MAX_CHARS);
  });
});

describe('unsupported formats give actionable reasons', () => {
  it.each([
    ['scan.png', /OCR/i],
    ['old.doc', /Legacy/i],
    ['bundle.zip', /Archives/i],
    ['thing.bin', /No text extractor/i],
  ])('%s explains why and points elsewhere', async (name, pattern) => {
    await expect(extractText(Buffer.from('data'), name)).rejects.toThrow(pattern);
    // Every message should route the caller to the download path.
    await expect(extractText(Buffer.from('data'), name)).rejects.toThrow(
      /download-mail-attachment|\.docx|OCR/
    );
  });

  it('carries the detected format on the error for logging', async () => {
    await expect(extractText(Buffer.from('x'), 'scan.tiff')).rejects.toMatchObject({
      name: 'UnsupportedFormatError',
      format: 'tiff',
    });
  });
});
