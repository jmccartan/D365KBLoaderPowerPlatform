import mammoth from 'mammoth';
import sanitizeHtml from 'sanitize-html';
import type { SourceFile, ProcessedArticle } from '../types';

const SANITIZE_OPTS: sanitizeHtml.IOptions = {
  allowedTags: [
    'h1','h2','h3','h4','h5','h6','p','br','hr',
    'ul','ol','li','blockquote','pre','code',
    'strong','em','b','i','u','s','sub','sup',
    'a','img','table','thead','tbody','tr','th','td',
    'div','span'
  ],
  allowedAttributes: {
    a: ['href', 'title', 'target', 'rel'],
    img: ['src', 'alt', 'title', 'width', 'height'],
    '*': ['class']
  },
  allowedSchemes: ['http', 'https', 'mailto', 'data'],
  transformTags: {
    a: (tag, attribs) => ({
      tagName: 'a',
      attribs: { ...attribs, target: '_blank', rel: 'noopener noreferrer' }
    })
  }
};

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

export async function processFile(file: SourceFile, content: ArrayBuffer): Promise<ProcessedArticle> {
  const warnings: string[] = [];
  let rawHtml = '';

  if (file.kind === 'docx') {
    const result = await mammoth.convertToHtml(
      { arrayBuffer: content },
      { styleMap: ["p[style-name='Title'] => h1:fresh", "p[style-name='Heading 1'] => h2:fresh"] }
    );
    rawHtml = result.value;
    for (const m of result.messages) warnings.push(`${m.type}: ${m.message}`);
  } else if (file.kind === 'html') {
    rawHtml = new TextDecoder('utf-8').decode(content);
  } else {
    warnings.push(`Unsupported file type: ${file.name}`);
    rawHtml = '';
  }

  const bodyHtml = extractBody(rawHtml);
  const sanitized = sanitizeHtml(bodyHtml, SANITIZE_OPTS).trim();
  const title = deriveTitle(rawHtml, file.name.replace(/\.[^.]+$/, ''));

  return {
    id: file.id,
    source: file,
    title,
    html: sanitized,
    rawHtml,
    warnings,
    selected: true,
    loadStatus: 'pending'
  };
}

export function classify(name: string): SourceFile['kind'] {
  const lower = name.toLowerCase();
  if (lower.endsWith('.docx')) return 'docx';
  if (lower.endsWith('.html') || lower.endsWith('.htm')) return 'html';
  return 'unsupported';
}
