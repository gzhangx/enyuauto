import { CATEGORY_ALIASES } from './newProjectForm';

const ENYU_SITE = 'https://enyu.acccn.org';
const LAMBDA_URL = 'https://fk9u03bqm3.execute-api.us-east-1.amazonaws.com/default/enyu-auto-v2-EnyuAutoFunction-r0jJLctkXxHF';

export type RelatedArticle = {
    title: string;
    link: string;
};

export type ArticleLanguage = 'zh' | 'en';

const HAS_CHINESE = /[\u4e00-\u9fff]/;

export function detectArticleLanguage(...texts: string[]): ArticleLanguage {
    return texts.some((text) => HAS_CHINESE.test(text || '')) ? 'zh' : 'en';
}

function titleMatchesLanguage(title: string, language: ArticleLanguage): boolean {
    const chinese = HAS_CHINESE.test(title);
    return language === 'zh' ? chinese : !chinese;
}

const LISTING_TITLES = new Set([
    ...CATEGORY_ALIASES.flatMap((alias) => [alias.chinese, alias.english]),
    '近期文章',
    '推荐文章',
    'Recent Articles',
    'Editor Picks',
    'Editor’s Picks',
    "Editor's Picks",
]);

const LISTING_PATH = /(recent-articles|editor-picks|search|about|contact)(-ch|-en)?\/?$/i;

/** Real EnYu posts look like /life/2026-05-z09/ — not category or listing pages. */
function isRealArticle(title: string, link: string): boolean {
    if (LISTING_TITLES.has(title.trim())) return false;
    try {
        const path = new URL(link).pathname.replace(/\/+$/, '');
        const last = path.split('/').filter(Boolean).pop() || '';
        if (LISTING_PATH.test(path) || /-(ch|en)$/i.test(last)) return false;
        return /^\d{4}-\d{2}-[a-z]\d+/i.test(last);
    } catch {
        return /\/\d{4}-\d{2}-[a-z]\d+/i.test(link);
    }
}

function decodeTitle(html: string): string {
    return html
        .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1')
        .replace(/<[^>]+>/g, '')
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)))
        .replace(/&#x([0-9a-fA-F]+);/g, (_, n) => String.fromCharCode(parseInt(n, 16)))
        .replace(/&nbsp;/g, ' ')
        .trim();
}

function isLocalDev(): boolean {
    return typeof location !== 'undefined' && /^(localhost|127\.0\.0\.1)$/.test(location.hostname);
}

function rssPath(query: string): string {
    return `/search/${encodeURIComponent(query)}/feed/rss2/`;
}

async function fetchRssXml(query: string): Promise<string> {
    const path = rssPath(query);
    const urls: string[] = [];
    if (isLocalDev()) {
        urls.push(`/enyu-search${path}`);
    }
    urls.push(`${ENYU_SITE}${path}`);

    for (const url of urls) {
        try {
            const res = await fetch(url);
            if (res.ok) return await res.text();
        } catch {
            // Browser CORS blocks the live site; local Vite proxy is the fallback.
        }
    }

    const res = await fetch(LAMBDA_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'searchEnyuSite', query }),
    });
    if (!res.ok) {
        throw new Error(`Site search failed (${res.status} ${res.statusText})`.trim());
    }
    const data = await res.json() as { xml?: string; error?: string };
    if (!data.xml) {
        throw new Error(data.error || 'Site search returned no results feed');
    }
    return data.xml;
}

function parseRssItems(xml: string): Array<RelatedArticle & { draft: boolean }> {
    const items: Array<RelatedArticle & { draft: boolean }> = [];
    const itemRe = /<item>([\s\S]*?)<\/item>/gi;
    let match: RegExpExecArray | null;
    while ((match = itemRe.exec(xml))) {
        const block = match[1];
        const title = decodeTitle(block.match(/<title>([\s\S]*?)<\/title>/i)?.[1] || '');
        const link = decodeTitle(block.match(/<link>([\s\S]*?)<\/link>/i)?.[1] || '')
            .replace(/#.*$/, '')
            .trim();
        const draft = /\/editing\//i.test(link);
        if (title && link) {
            items.push({ title, link, draft });
        }
    }
    return items;
}

function addUnique(
    results: RelatedArticle[],
    seen: Set<string>,
    items: Array<RelatedArticle & { draft: boolean }>,
    excludeSlug: string,
    limit: number,
    language?: ArticleLanguage,
): void {
    for (const item of items) {
        if (results.length >= limit) return;
        if (item.draft) continue;
        if (!isRealArticle(item.title, item.link)) continue;
        if (language && !titleMatchesLanguage(item.title, language)) continue;
        const link = item.link.toLowerCase();
        if (excludeSlug && (link.includes(excludeSlug) || link.endsWith(`/${excludeSlug}/`))) continue;
        if (seen.has(link)) continue;
        seen.add(link);
        results.push({ title: item.title, link: item.link });
    }
}

export async function findRelatedEnyuArticles(
    query: string | string[],
    excludeSlug?: string,
    limit = 5,
    language?: ArticleLanguage,
): Promise<RelatedArticle[]> {
    const terms = [...new Set(
        (Array.isArray(query) ? query : query.split(/\s+/))
            .map((part) => part.replace(/\s+/g, ' ').trim())
            .filter(Boolean)
            .map((term) => {
                if (!language) return term;
                const alias = CATEGORY_ALIASES.find((item) =>
                    item.chinese === term || item.english.toLowerCase() === term.toLowerCase());
                return alias ? (language === 'zh' ? alias.chinese : alias.english) : term;
            }),
    )];
    if (terms.length === 0) return [];

    const exclude = (excludeSlug || '').trim().toLowerCase();
    const seen = new Set<string>();
    const results: RelatedArticle[] = [];

    // Search each term separately so a long "title + author + category" query
    // does not miss articles that only match one of those.
    for (const term of terms) {
        if (results.length >= limit) break;
        const xml = await fetchRssXml(term);
        addUnique(results, seen, parseRssItems(xml), exclude, limit, language);
    }

    return results.slice(0, limit);
}

const BREAK_BOLD = '</b></strong></h1></h2></h3></h4></h5>';
const INDENT = '&nbsp;&nbsp;&nbsp;&nbsp;';

export function formatRelatedArticlesHtml(articles: RelatedArticle[]): string {
    if (articles.length === 0) {
        return `${BREAK_BOLD}<br>${INDENT}（未找到相关文章）`;
    }
    const items = articles
        .map((article, index) => `${INDENT}${index + 1}. ${article.title}<br>${INDENT}${article.link}`)
        .join('<br>');
    return `${BREAK_BOLD}<br>${items}`;
}
