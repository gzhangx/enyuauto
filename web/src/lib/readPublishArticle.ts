import * as mammoth from 'mammoth';
import { getDocument, GlobalWorkerOptions } from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import {
    downloadDriveItemContent,
    listFolderFiles,
    listMatchingSubfolderFiles,
    type SiblingFileInfo,
} from './msGraphConfig';

GlobalWorkerOptions.workerSrc = pdfWorker;

type ArticleMeta = {
    title?: string;
    author?: string;
    category?: string;
};

type Log = { doLog: (msg: string, isError?: boolean) => void };

export type PublishArticleText = {
    source: string;
    text: string;
};

export type PublishArticleSearch = PublishArticleText & {
    charCount: number;
    terms: string[];
};

const META_LINE = /^(作者|日期|类别|Author|Date|Category|称呼)[：:\s]/i;
const STOP_START = /^(我们|他们|她们|你们|这个|那个|一个|可以|因为|所以|如果|但是|然而|不是|自己|这样|那样|什么|怎么|没有|已经|还是|或者|以及|并且|因此|于是|就是|也是|还要)/;

function pickSignedPdf(files: SiblingFileInfo[]): SiblingFileInfo | undefined {
    return files.find((file) => /_signed\.pdf$/i.test(file.name));
}

function pickDocx(files: SiblingFileInfo[]): SiblingFileInfo | undefined {
    return files.find((file) => /\.docx$/i.test(file.name) && !file.name.startsWith('~$'));
}

async function extractDocxText(buffer: ArrayBuffer): Promise<string> {
    const result = await mammoth.extractRawText({ arrayBuffer: buffer });
    return (result.value || '')
        .replace(/[ \t]+/g, ' ')
        .replace(/\n{3,}/g, '\n\n')
        .trim();
}

async function listPublishFiles(msToken: string, mainFolder: string): Promise<SiblingFileInfo[]> {
    try {
        const files = await listFolderFiles(msToken, mainFolder, '4 Publish');
        if (files.length) return files;
    } catch {
        // Folder name may differ slightly (spacing / punctuation).
    }
    return listMatchingSubfolderFiles(msToken, mainFolder, /4\s*publish/i);
}

async function extractPdfText(buffer: ArrayBuffer): Promise<string> {
    const task = getDocument({ data: new Uint8Array(buffer) });
    const pdf = await task.promise;
    const pages: string[] = [];
    const maxPages = Math.min(pdf.numPages, 8);
    for (let i = 1; i <= maxPages; i++) {
        const page = await pdf.getPage(i);
        const content = await page.getTextContent();
        const parts: string[] = [];
        let lastY: number | null = null;
        for (const item of content.items) {
            if (!('str' in item)) continue;
            const y = Array.isArray(item.transform) ? item.transform[5] : undefined;
            if (lastY != null && typeof y === 'number' && Math.abs(y - lastY) > 8) {
                parts.push('\n');
            }
            parts.push(item.str);
            if ('hasEOL' in item && item.hasEOL) parts.push('\n');
            if (typeof y === 'number') lastY = y;
        }
        const text = parts.join('').replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim();
        if (text) pages.push(text);
    }
    return pages.join('\n\n').trim();
}

function categoryFromArticle(text: string): string {
    const match = text.match(/(?:类别|Category)\s*[：:]\s*([^\n\r]{2,20})/i);
    return (match?.[1] || '').replace(/\s+/g, ' ').trim();
}

export function extractArticleSearchTerms(text: string, meta: ArticleMeta = {}): string[] {
    const terms: string[] = [];
    const seen = new Set<string>();
    const add = (raw: string) => {
        const value = raw.replace(/\s+/g, ' ').trim();
        const compact = value.replace(/\s+/g, '');
        if (compact.length < 2 || seen.has(compact)) return;
        seen.add(compact);
        terms.push(value);
    };

    add(categoryFromArticle(text));
    for (const match of text.matchAll(/《([^》]{2,20})》/g)) {
        add(match[1]);
    }

    const title = (meta.title || '').replace(/\s+/g, '').trim();
    const author = (meta.author || '').replace(/\s+/g, '').trim();
    const body = text
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter((line) => line && !META_LINE.test(line) && line !== title && line !== author)
        .join('\n');
    const sentences = body
        .split(/[。！？；\n]+/)
        .map((sentence) => sentence.replace(/\s+/g, '').trim())
        .filter((sentence) => {
            if (sentence.length < 8) return false;
            if (title && (sentence === title || sentence.includes(title))) return false;
            if (author && sentence.includes(author)) return false;
            return true;
        });

    for (const sentence of sentences.slice(0, 8)) {
        let chunk = sentence.slice(0, 6);
        if (STOP_START.test(chunk) && sentence.length > 8) {
            chunk = sentence.slice(2, 8);
        }
        if (chunk.length >= 4) add(chunk);
        if (terms.length >= 8) break;
    }

    return terms.slice(0, 8);
}

function splitSentences(text: string): string[] {
    return (text.match(/[^。！？]+[。！？]?/g) || [])
        .map((sentence) => sentence.replace(/\s+/g, '').trim())
        .filter((sentence) => sentence.length >= 8);
}

function isMetaOrHeading(line: string, meta: ArticleMeta): boolean {
    const compact = line.replace(/\s+/g, '').trim();
    if (!compact) return true;
    if (META_LINE.test(compact)) return true;
    if (meta.title && (compact === meta.title.replace(/\s+/g, '') || compact.includes(meta.title.replace(/\s+/g, '')))) return true;
    if (meta.author && compact === meta.author.replace(/\s+/g, '')) return true;
    if (/^(前言|后记|序|引言|结语|摘要|概要)$/.test(compact)) return true;
    if (compact.length <= 8 && !/[。！？]/.test(compact)) return true;
    return false;
}

function isProseSentence(sentence: string, meta: ArticleMeta): boolean {
    const compact = sentence.replace(/\s+/g, '').trim();
    if (compact.length < 15) return false;
    return !isMetaOrHeading(compact, meta);
}

function joinSentences(sentences: string[]): string {
    return sentences
        .map((sentence) => {
            const text = sentence.trim();
            return /[。！？]$/.test(text) ? text : `${text}。`;
        })
        .join('');
}

function bodyParagraphs(text: string, meta: ArticleMeta): string[] {
    return text
        .split(/\n+/)
        .map((line) => line.replace(/[ \t]+/g, ' ').trim())
        .filter((line) => !isMetaOrHeading(line, meta));
}

const OPENING_STORY = /^(今天|今日|那天|有一天|有一次|记得|早上|早晨|晚上|前几天|那天晚上)/;
const HIGHLIGHT_HINT = /神|主|圣经|信心|祷告|恩典|十字架|耶稣|圣灵|《/;

function allProseSentences(text: string, meta: ArticleMeta): string[] {
    return bodyParagraphs(text, meta).flatMap((paragraph) =>
        splitSentences(paragraph).filter((sentence) => isProseSentence(sentence, meta)));
}

function highlightScore(sentence: string): number {
    const compact = sentence.replace(/\s+/g, '');
    if (OPENING_STORY.test(compact)) return -10;
    if (compact.length < 18 || compact.length > 90) return -5;
    let score = 0;
    if (HIGHLIGHT_HINT.test(compact)) score += 4;
    if (/《[^》]{2,20}》/.test(sentence)) score += 3;
    if (compact.length >= 24 && compact.length <= 70) score += 2;
    return score;
}

function canUseInHighlight(sentence: string): boolean {
    return !OPENING_STORY.test(sentence.replace(/\s+/g, ''));
}

/** 3–5 consecutive sentences around a highlight, not the whole paragraph. */
function pickHighlight(text: string, meta: ArticleMeta): string | null {
    const sentences = allProseSentences(text, meta);
    if (sentences.length === 0) return null;

    const ranked = sentences
        .map((sentence, index) => ({ sentence, index, score: highlightScore(sentence) }))
        .filter((item) => item.score > 0)
        .sort((a, b) => b.score - a.score || a.index - b.index);

    const center = ranked[0]?.index ?? sentences.findIndex((sentence) => canUseInHighlight(sentence));
    if (center < 0) return null;

    const picked: string[] = [];
    for (let i = center; i < sentences.length && picked.length < 5; i++) {
        if (!canUseInHighlight(sentences[i])) break;
        picked.push(sentences[i]);
    }
    for (let i = center - 1; i >= 0 && picked.length < 3; i--) {
        if (!canUseInHighlight(sentences[i])) break;
        picked.unshift(sentences[i]);
    }
    if (picked.length < 3) return null;
    return joinSentences(picked.slice(0, 5));
}

function writeOwnAbstract(text: string, meta: ArticleMeta): string {
    const title = (meta.title || '').trim();
    const author = (meta.author || '').trim();
    const category = (meta.category || '').trim();
    const books = [...new Set([...text.matchAll(/《([^》]{2,20})》/g)].map((match) => match[1]))].slice(0, 2);
    const written: string[] = [];

    if (title && category) {
        written.push(`这是一篇「${category}」文章，题目是「${title}」`);
    } else if (title) {
        written.push(`本文题目是「${title}」`);
    } else if (category) {
        written.push(`这是一篇关于「${category}」的文章`);
    } else {
        written.push('这是一篇恩语网刊文章');
    }

    written.push(author
        ? `作者${author}在文中分享了自己的经历与思考`
        : '作者在文中分享了相关的经历与思考');

    if (books.length >= 2) {
        written.push(`文章提到《${books[0]}》与《${books[1]}》，并把经文里的看见写进日常处境`);
    } else if (books.length === 1) {
        written.push(`文章围绕《${books[0]}》展开，把经文与生活连在一起谈`);
    } else {
        written.push('文章把信仰里的看见，写进日常的处境与选择里');
    }

    written.push('读者可以从中思想神的带领，以及如何在生活中回应');
    return joinSentences(written.slice(0, 5));
}

/** A short highlight from the article, or a brief written 摘要. */
export function extractArticleAbstract(text: string, meta: ArticleMeta = {}): { abstract: string; foundParagraph: boolean } {
    const excerpt = pickHighlight(text, meta);
    if (excerpt) {
        return { abstract: excerpt, foundParagraph: true };
    }
    return { abstract: writeOwnAbstract(text, meta), foundParagraph: false };
}

function escapeHtml(text: string): string {
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

const BREAK_BOLD = '</b></strong></h1></h2></h3></h4></h5>';
const INDENT = '&nbsp;&nbsp;&nbsp;&nbsp;';

export function formatArticleAbstractHtml(abstract: string): string {
    const text = abstract.trim() ? escapeHtml(abstract) : '（未能生成文章摘要）';
    return `${BREAK_BOLD}<br>${INDENT}${text}`;
}

export async function readPublishArticleText(
    msToken: string,
    mainFolder: string,
    log?: Log,
): Promise<PublishArticleText> {
    const files = await listPublishFiles(msToken, mainFolder);
    const signedPdf = pickSignedPdf(files);
    if (!signedPdf) {
        throw new Error('No *_signed.pdf in 4 Publish');
    }

    const docx = pickDocx(files);
    if (docx) {
        log?.doLog(`processOperation: reading publish article text from ${docx.name} (${signedPdf.name} is present)`);
        const text = await extractDocxText(await downloadDriveItemContent(msToken, docx));
        if (text.length >= 20) {
            return { source: docx.name, text };
        }
        log?.doLog(`processOperation: ${docx.name} had little text, falling back to ${signedPdf.name}`);
    }

    log?.doLog(`processOperation: reading publish article ${signedPdf.name}`);
    const text = await extractPdfText(await downloadDriveItemContent(msToken, signedPdf));
    if (text.length < 20) {
        throw new Error(`Could not extract article text from ${signedPdf.name}`);
    }
    return { source: signedPdf.name, text };
}

export async function readPublishArticleForSearch(
    msToken: string,
    mainFolder: string,
    meta: ArticleMeta,
    log?: Log,
): Promise<PublishArticleSearch> {
    const article = await readPublishArticleText(msToken, mainFolder, log);
    return {
        ...article,
        charCount: article.text.length,
        terms: extractArticleSearchTerms(article.text, meta),
    };
}
