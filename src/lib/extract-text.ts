import { unzipSync, strFromU8 } from 'fflate';

/**
 * Server-side text extraction from document bytes.
 *
 * Why this exists: the connector can already hand Claude a pre-authed download
 * URL for a file, but Claude's sandbox has no egress to our SharePoint host, so
 * it cannot fetch that URL to read the file (verified 2026-08-10:
 * `403 host_not_allowed`). Extracting the text here and returning that instead
 * sidesteps the network boundary entirely — and it costs far less context than
 * the base64 fallback, because prose is smaller than an encoded binary.
 *
 * Scope: text, not fidelity. This is for "what does this document say", not for
 * reproducing layout. Scanned images contain no extractable text and are
 * reported as such rather than returned empty — those still need the download
 * URL and a human.
 */

/** Hard ceiling regardless of what the caller asks for, to protect context. */
export const MAX_CHARS_LIMIT = 200_000;
/** Default cap — roughly 12k tokens, enough for most documents. */
export const DEFAULT_MAX_CHARS = 50_000;

export interface ExtractResult {
  text: string;
  /** What we decided the file was, after sniffing name and content type. */
  format: string;
  /** True when `text` was cut at the cap; the tail is missing, not empty. */
  truncated: boolean;
  /** Characters extracted before truncation, so callers can gauge the whole. */
  totalChars: number;
  /** Per-format extras: PDF page count, spreadsheet sheet names. */
  detail?: Record<string, unknown>;
}

export class UnsupportedFormatError extends Error {
  constructor(
    public readonly format: string,
    message: string
  ) {
    super(message);
    this.name = 'UnsupportedFormatError';
  }
}

const OOXML_MAIN: Record<string, RegExp> = {
  docx: /^word\/document\.xml$/,
  pptx: /^ppt\/slides\/slide\d+\.xml$/,
};

/**
 * Strips XML tags and decodes entities, inserting whitespace where tags implied
 * a break. Without the break handling, `<w:p>One</w:p><w:p>Two</w:p>` collapses
 * into "OneTwo".
 */
function xmlToText(xml: string): string {
  return xml
    // Paragraph, row, line-break and slide boundaries become real breaks.
    .replace(/<\/(w:p|a:p|w:tr|tr)>/g, '\n')
    .replace(/<(w:br|w:cr|a:br)\b[^>]*\/?>/g, '\n')
    .replace(/<\/(w:tc|tc)>/g, '\t')
    .replace(/<[^>]+>/g, '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&amp;/g, '&')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n');
}

function guessFormat(filename: string, contentType: string): string {
  const ext = (filename.match(/\.([a-z0-9]+)$/i)?.[1] || '').toLowerCase();
  if (ext) return ext;
  const ct = contentType.toLowerCase();
  if (ct.includes('pdf')) return 'pdf';
  if (ct.includes('wordprocessingml')) return 'docx';
  if (ct.includes('spreadsheetml')) return 'xlsx';
  if (ct.includes('presentationml')) return 'pptx';
  if (ct.includes('json')) return 'json';
  if (ct.includes('html')) return 'html';
  if (ct.includes('csv')) return 'csv';
  if (ct.startsWith('text/')) return 'txt';
  return 'unknown';
}

/** Column letters + row number → nothing; we only need the row for grouping. */
function rowOf(ref: string): number {
  return Number(ref.replace(/[^0-9]/g, '')) || 0;
}

/**
 * XLSX text extraction. Cells live in sheet XML as either inline values or
 * indexes into a shared string table, so the table has to be read first.
 * Output is tab-separated rows per sheet — compact and readable as data.
 */
function extractXlsx(files: Record<string, Uint8Array>): { text: string; sheets: string[] } {
  const shared: string[] = [];
  const sharedXml = files['xl/sharedStrings.xml'];
  if (sharedXml) {
    const xml = strFromU8(sharedXml);
    for (const si of xml.match(/<si>[\s\S]*?<\/si>/g) ?? []) {
      shared.push(xmlToText(si).replace(/\n/g, ' ').trim());
    }
  }

  // Sheet display names, in workbook order, so output is labelled usefully.
  const names: string[] = [];
  const wb = files['xl/workbook.xml'];
  if (wb) {
    for (const m of strFromU8(wb).matchAll(/<sheet[^>]*name="([^"]*)"/g)) names.push(m[1]);
  }

  const sheetPaths = Object.keys(files)
    .filter((p) => /^xl\/worksheets\/sheet\d+\.xml$/.test(p))
    .sort((a, b) => Number(a.match(/\d+/)![0]) - Number(b.match(/\d+/)![0]));

  const out: string[] = [];
  sheetPaths.forEach((path, i) => {
    const xml = strFromU8(files[path]);
    const label = names[i] ?? `Sheet${i + 1}`;
    const rows = new Map<number, string[]>();

    for (const cellMatch of xml.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
      const attrs = cellMatch[1];
      const body = cellMatch[2];
      const ref = attrs.match(/r="([A-Z]+\d+)"/)?.[1] ?? '';
      const type = attrs.match(/t="([^"]+)"/)?.[1] ?? 'n';
      const raw = body.match(/<v>([\s\S]*?)<\/v>/)?.[1];

      let value: string;
      if (type === 's') {
        value = raw !== undefined ? (shared[Number(raw)] ?? '') : '';
      } else if (type === 'inlineStr') {
        value = xmlToText(body).replace(/\n/g, ' ').trim();
      } else {
        value = raw !== undefined ? xmlToText(raw).trim() : '';
      }
      if (value === '') continue;

      const r = rowOf(ref);
      if (!rows.has(r)) rows.set(r, []);
      rows.get(r)!.push(value);
    }

    if (rows.size === 0) return;
    const ordered = [...rows.keys()].sort((a, b) => a - b);
    out.push(`--- ${label} ---`);
    for (const r of ordered) out.push(rows.get(r)!.join('\t'));
    out.push('');
  });

  return { text: out.join('\n').trim(), sheets: names.length ? names : sheetPaths };
}

function extractOoxml(buffer: Buffer, format: 'docx' | 'pptx'): string {
  const files = unzipSync(new Uint8Array(buffer));
  const pattern = OOXML_MAIN[format];
  const parts = Object.keys(files)
    .filter((p) => pattern.test(p))
    // slide2 must not sort before slide10 lexicographically.
    .sort((a, b) => (Number(a.match(/\d+/)?.[0] ?? 0) - Number(b.match(/\d+/)?.[0] ?? 0)) || a.localeCompare(b));
  if (parts.length === 0) {
    throw new UnsupportedFormatError(format, `No ${format} content found inside the archive.`);
  }
  return parts.map((p) => xmlToText(strFromU8(files[p]))).join('\n\n').trim();
}

async function extractPdf(buffer: Buffer): Promise<{ text: string; pages: number }> {
  // Imported lazily so the pdf machinery is only loaded when a PDF shows up —
  // it is the heaviest thing in the dependency tree.
  const { extractText: pdfExtract, getDocumentProxy } = await import('unpdf');
  const pdf = await getDocumentProxy(new Uint8Array(buffer));
  const { text, totalPages } = await pdfExtract(pdf, { mergePages: true });
  return { text: (Array.isArray(text) ? text.join('\n\n') : text) ?? '', pages: totalPages };
}

/**
 * Extracts readable text from a document.
 *
 * Throws UnsupportedFormatError for formats with no text layer (images,
 * archives, binaries) so callers can give a precise reason rather than an empty
 * string that looks like an empty document.
 */
export async function extractText(
  buffer: Buffer,
  filename: string,
  contentType = '',
  maxChars: number = DEFAULT_MAX_CHARS
): Promise<ExtractResult> {
  const format = guessFormat(filename, contentType);
  const cap = Math.max(1, Math.min(maxChars, MAX_CHARS_LIMIT));

  let text: string;
  let detail: Record<string, unknown> | undefined;

  switch (format) {
    case 'pdf': {
      const { text: t, pages } = await extractPdf(buffer);
      text = t;
      detail = { pages };
      if (text.trim() === '') {
        throw new UnsupportedFormatError(
          'pdf',
          `This PDF has ${pages} page(s) but no extractable text layer — it is almost certainly a ` +
            `scan or image-only export. Text extraction cannot help; use download-mail-attachment ` +
            `to get the file itself.`
        );
      }
      break;
    }
    case 'docx':
    case 'pptx':
      text = extractOoxml(buffer, format);
      break;
    case 'xlsx':
    case 'xlsm': {
      const { text: t, sheets } = extractXlsx(unzipSync(new Uint8Array(buffer)));
      text = t;
      detail = { sheets };
      break;
    }
    case 'html':
    case 'htm':
      text = xmlToText(buffer.toString('utf8'));
      break;
    case 'txt':
    case 'csv':
    case 'tsv':
    case 'json':
    case 'xml':
    case 'md':
    case 'log':
    case 'yml':
    case 'yaml':
      text = buffer.toString('utf8');
      break;
    // Legacy Office formats are OLE compound files, not ZIP — a different
    // parser entirely, and rare enough not to justify one.
    case 'doc':
    case 'xls':
    case 'ppt':
      throw new UnsupportedFormatError(
        format,
        `Legacy .${format} files are not supported (they use an older binary container). ` +
          `Ask the sender for a .${format}x version, or use download-mail-attachment.`
      );
    case 'png':
    case 'jpg':
    case 'jpeg':
    case 'gif':
    case 'bmp':
    case 'tif':
    case 'tiff':
    case 'heic':
      throw new UnsupportedFormatError(
        format,
        `Images contain no extractable text (this would need OCR, which is not available). ` +
          `Use download-mail-attachment to get the file.`
      );
    case 'zip':
    case '7z':
    case 'rar':
    case 'gz':
    case 'tar':
      throw new UnsupportedFormatError(
        format,
        `Archives are not unpacked by this tool. Use download-mail-attachment to get the file.`
      );
    default:
      throw new UnsupportedFormatError(
        format,
        `No text extractor for '${format}' (content type '${contentType || 'unknown'}'). ` +
          `Supported: pdf, docx, xlsx, pptx, and plain-text formats. ` +
          `Use download-mail-attachment for anything else.`
      );
  }

  const normalized = text.replace(/\r\n/g, '\n').trim();
  const truncated = normalized.length > cap;
  return {
    text: truncated ? normalized.slice(0, cap) : normalized,
    format,
    truncated,
    totalChars: normalized.length,
    detail,
  };
}
