import mammoth from 'mammoth';
import sanitizeHtml from 'sanitize-html';
import type { SourceFile, ProcessedArticle } from '../types';

export type UploadImageHandler = (name: string, bytes: ArrayBuffer, contentType: string) => Promise<string>;

export const SANITIZE_OPTS: sanitizeHtml.IOptions = {
  allowedTags: [
    'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'br', 'hr',
    'ul', 'ol', 'li', 'blockquote', 'pre', 'code',
    'strong', 'em', 'b', 'i', 'u', 's', 'sub', 'sup',
    'a', 'img', 'table', 'thead', 'tbody', 'tr', 'th', 'td',
    'div', 'span'
  ],
  allowedAttributes: {
    a: ['href', 'title', 'target', 'rel'],
    img: ['src', 'alt', 'title', 'width', 'height'],
    '*': ['class']
  },
  allowedSchemes: ['http', 'https', 'mailto', 'data'],
  transformTags: {
    a: (_tag, attribs) => ({
      tagName: 'a',
      attribs: { ...attribs, target: '_blank', rel: 'noopener noreferrer' }
    })
  }
};

export function sanitizeArticleHtml(html: string): string {
  return sanitizeHtml(html, SANITIZE_OPTS).trim();
}

function deriveTitle(html: string, fallback: string): string {
  const h1 = /<h1[^>]*>([\s\S]*?)<\/h1>/i.exec(html);
  if (h1) return stripTags(h1[1]).trim().slice(0, 200) || fallback;
  const title = /<title[^>]*>([\s\S]*?)<\/title>/i.exec(html);
  if (title) return stripTags(title[1]).trim().slice(0, 200) || fallback;
  return fallback;
}

function stripTags(s: string): string {
  return s.replace(/<[^>]+>/g, '');
}

function extractBody(html: string): string {
  const m = /<body[^>]*>([\s\S]*?)<\/body>/i.exec(html);
  return m ? m[1] : html;
}

async function pdfToHtml(buf: ArrayBuffer): Promise<{ html: string; warnings: string[] }> {
  const warnings: string[] = [];
  const pdfjs = await import('pdfjs-dist');
  // @ts-ignore — pdfjs-dist exposes a mutable GlobalWorkerOptions
  pdfjs.GlobalWorkerOptions.workerSrc = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${pdfjs.version}/build/pdf.worker.min.mjs`;
  const doc = await pdfjs.getDocument({ data: buf }).promise;
  const paragraphs: string[] = [];
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const text = await page.getTextContent();
    const lines: string[] = [];
    let current = '';
    let lastY: number | undefined;
    for (const item of text.items as Array<{ str?: string; hasEOL?: boolean; transform?: number[] }>) {
      const y = item.transform?.[5] ?? 0;
      if (lastY !== undefined && Math.abs(lastY - y) > 2) {
        if (current.trim()) lines.push(current.trim());
        current = '';
      }
      current += (item.str ?? '') + (item.hasEOL ? ' ' : '');
      lastY = y;
    }
    if (current.trim()) lines.push(current.trim());
    if (lines.length === 0) {
      warnings.push(`Page ${p} had no extractable text (image-only?)`);
      continue;
    }
    paragraphs.push(...lines.map(line => `<p>${escape(line)}</p>`));
    if (p < doc.numPages) paragraphs.push('<hr />');
  }
  return { html: paragraphs.join('\n'), warnings };
}

function escape(s: string): string {
  return s.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]!));
}

export async function processFile(file: SourceFile, content: ArrayBuffer, uploadImage?: UploadImageHandler): Promise<ProcessedArticle> {
  const warnings: string[] = [];
  let rawHtml = '';

  if (file.kind === 'docx') {
    let imageCounter = 0;
    const result = await mammoth.convertToHtml(
      { arrayBuffer: content },
      {
        styleMap: ["p[style-name='Title'] => h1:fresh", "p[style-name='Heading 1'] => h2:fresh"],
        convertImage: mammoth.images.imgElement(async image => {
          const contentType = image.contentType || 'image/png';
          const bytes = await image.readAsArrayBuffer();
          const baseName = file.name.replace(/\.[^.]+$/, '');
          const ext = extensionForContentType(contentType);
          const imageName = `${baseName}-image-${++imageCounter}.${ext}`;
          const src = uploadImage
            ? await uploadImage(imageName, bytes, contentType)
            : `data:${contentType};base64,${await image.readAsBase64String()}`;
          return { src };
        })
      }
    );
    rawHtml = result.value;
    for (const message of result.messages) warnings.push(`${message.type}: ${message.message}`);
  } else if (file.kind === 'html') {
    rawHtml = new TextDecoder('utf-8').decode(content);
  } else if (file.kind === 'pdf') {
    const out = await pdfToHtml(content);
    rawHtml = out.html;
    warnings.push(...out.warnings);
  } else if (file.kind === 'md') {
    rawHtml = mdToHtml(new TextDecoder('utf-8').decode(content));
  } else {
    warnings.push(`Unsupported file type: ${file.name}`);
    rawHtml = '';
  }

  const bodyHtml = extractBody(rawHtml);
  const sanitized = sanitizeArticleHtml(bodyHtml);
  const title = deriveTitle(rawHtml, file.name.replace(/\.[^.]+$/, ''));

  return {
    id: file.id,
    source: file,
    title,
    html: sanitized,
    rawHtml,
    warnings,
    findings: [],
    selected: true,
    loadStatus: 'pending'
  };
}

function extensionForContentType(contentType: string): string {
  switch (contentType.toLowerCase()) {
    case 'image/jpeg':
      return 'jpg';
    case 'image/gif':
      return 'gif';
    case 'image/svg+xml':
      return 'svg';
    case 'image/webp':
      return 'webp';
    default:
      return 'png';
  }
}

export function classify(name: string): SourceFile['kind'] {
  const lower = name.toLowerCase();
  if (lower.endsWith('.docx')) return 'docx';
  if (lower.endsWith('.html') || lower.endsWith('.htm')) return 'html';
  if (lower.endsWith('.pdf')) return 'pdf';
  if (lower.endsWith('.md') || lower.endsWith('.markdown')) return 'md';
  return 'unsupported';
}

/** Tiny Markdown → HTML converter (headings, lists, code fences, paragraphs, **bold** / *italic*). */
function mdToHtml(src: string): string {
  const lines = src.replace(/\r\n/g, '\n').split('\n');
  const out: string[] = [];
  let inList = false;
  let inCode = false;
  let para: string[] = [];
  const flushPara = () => {
    if (para.length) {
      out.push('<p>' + inline(para.join(' ')) + '</p>');
      para = [];
    }
  };
  for (const raw of lines) {
    const line = raw.trimEnd();
    if (line.startsWith('```')) {
      flushPara();
      if (inCode) {
        out.push('</code></pre>');
        inCode = false;
      } else {
        out.push('<pre><code>');
        inCode = true;
      }
      continue;
    }
    if (inCode) {
      out.push(escape(raw) + '\n');
      continue;
    }
    if (/^#{1,6}\s/.test(line)) {
      flushPara();
      if (inList) {
        out.push('</ul>');
        inList = false;
      }
      const n = line.match(/^#+/)![0].length;
      const text = line.replace(/^#+\s*/, '');
      out.push(`<h${n}>${inline(text)}</h${n}>`);
      continue;
    }
    if (/^[-*+]\s/.test(line)) {
      flushPara();
      if (!inList) {
        out.push('<ul>');
        inList = true;
      }
      out.push('<li>' + inline(line.replace(/^[-*+]\s+/, '')) + '</li>');
      continue;
    }
    if (line.trim() === '') {
      flushPara();
      if (inList) {
        out.push('</ul>');
        inList = false;
      }
      continue;
    }
    para.push(line);
  }
  flushPara();
  if (inList) out.push('</ul>');
  if (inCode) out.push('</code></pre>');
  return out.join('\n');
}

function inline(text: string): string {
  return escape(text)
    .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
    .replace(/(^|[^*])\*([^*]+)\*/g, '$1<em>$2</em>')
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>');
}
