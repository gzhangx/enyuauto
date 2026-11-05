export type CategoryAlias = {
    chinese: string;
    english: string;
    letter: string;
};

/** English names are translated to these Chinese categories before the slug letter is chosen. */
export const CATEGORY_ALIASES: CategoryAlias[] = [
    { chinese: '主内生活', english: 'A life in Christ', letter: 'z' },
    { chinese: '灵修园地', english: 'Garden of Devotions', letter: 'l' },
    { chinese: '感恩见证', english: 'Thanksgiving Testimonies', letter: 'g' },
    { chinese: '牧者的话', english: 'Pastrol & Elder Messages', letter: 'm' },
    { chinese: '恩语艺廊', english: '', letter: 'e' },
    { chinese: '宣教无国界', english: 'Missions Beyond Borders', letter: 'x' },
];

const EXTRA_ENGLISH_ALIASES: { english: string; chinese: string }[] = [
    { english: 'Pastoral & Elder Messages', chinese: '牧者的话' },
    { english: '佳作分享', chinese: '恩语艺廊' },
    { english: '佳作欣赏', chinese: '恩语艺廊' },
    { english: 'Shared Book Passages', chinese: '恩语艺廊' },
];

export type ParsedFileInput = {
    category: string;
    /** YYYY-MM-DD */
    date: string;
    author: string;
    title: string;
};

export type SlugSourceRow = {
    文件?: string;
    文章类别?: string;
    slug?: string;
    作者?: string;
    作者电邮?: string;
    称呼?: string;
};

function normalizeAuthorName(name: string): string {
    return name.trim().replace(/\s+/g, ' ').toLowerCase();
}

/** Latest non-empty 作者电邮 for an author who already has a row. */
export function findAuthorEmail(author: string, rows: SlugSourceRow[]): string {
    const key = normalizeAuthorName(author);
    if (!key) {
        return '';
    }
    let email = '';
    for (const row of rows) {
        if (normalizeAuthorName(row.作者 || '') !== key) {
            continue;
        }
        const found = (row.作者电邮 || '').trim();
        if (found) {
            email = found;
        }
    }
    return email;
}

function canonicalChinese(category: string): string {
    const trimmed = category.trim();
    const lower = trimmed.toLowerCase();
    for (const alias of CATEGORY_ALIASES) {
        if (alias.chinese === trimmed || (alias.english && alias.english.toLowerCase() === lower)) {
            return alias.chinese;
        }
    }
    for (const extra of EXTRA_ENGLISH_ALIASES) {
        if (extra.english.toLowerCase() === lower) {
            return extra.chinese;
        }
    }
    return '';
}

/** Old 佳作分享 / 佳作欣赏 rows had no 文章名. 恩语艺廊 is a normal category. */
export function categoryHasNoArticleTitle(category: string): boolean {
    const raw = category.trim();
    return raw === '佳作分享' || raw === '佳作欣赏' || raw.toLowerCase() === 'shared book passages';
}

export function categoryLetter(category: string): string {
    const chinese = canonicalChinese(category);
    const found = CATEGORY_ALIASES.find(alias => alias.chinese === chinese);
    return found?.letter ?? '';
}

function isCjkToken(token: string): boolean {
    return /[\u4e00-\u9fff]/.test(token);
}

function isLatinNameToken(token: string): boolean {
    return /^[A-Za-z][A-Za-z.'’-]*$/.test(token);
}

function knownCategoryNames(): string[] {
    return CATEGORY_ALIASES
        .flatMap(alias => [
            alias.chinese,
            alias.english,
            ...EXTRA_ENGLISH_ALIASES.filter(extra => extra.chinese === alias.chinese).map(extra => extra.english),
        ])
        .filter(Boolean)
        .sort((a, b) => b.length - a.length);
}

/** `YYYY-MM-DD` from a 文件 name such as `2026-10-16 恩语艺廊 罗丽丹 留一个空间`. */
export function dateFromText(text: string): string {
    const match = text.trim().match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (!match) {
        return '';
    }
    return `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
}

function splitCategoryFromRest(rest: string): { category: string; remainder: string } | null {
    const trimmed = rest.trim();
    if (!trimmed) {
        return null;
    }
    for (const name of knownCategoryNames()) {
        if (trimmed.toLowerCase().startsWith(name.toLowerCase())) {
            return { category: name, remainder: trimmed.slice(name.length).trim() };
        }
    }
    const tokens = trimmed.split(/\s+/).filter(Boolean);
    return { category: tokens[0], remainder: tokens.slice(1).join(' ') };
}

function parseAuthorAndTitle(category: string, date: string, rest: string): ParsedFileInput | null {
    const tokens = rest.trim().split(/\s+/).filter(Boolean);
    const noTitle = categoryHasNoArticleTitle(category);
    if (!category || tokens.length < 1 || (!noTitle && tokens.length < 2)) {
        return null;
    }
    let authorCount = 1;
    if (!isCjkToken(tokens[0]) && isLatinNameToken(tokens[0])) {
        authorCount = 0;
        while (authorCount < tokens.length && isLatinNameToken(tokens[authorCount])) {
            authorCount += 1;
        }
        if (authorCount === 0 || (!noTitle && authorCount >= tokens.length)) {
            return null;
        }
    }
    const author = tokens.slice(0, authorCount).join(' ');
    const title = tokens.slice(authorCount).join(' ');
    if (!author || (!noTitle && !title)) {
        return null;
    }
    return { category, date, author, title };
}

/**
 * Accepts add-form `{文章类别} {date} {作者} {文章名}`
 * or sheet `{date} {文章类别} {作者} {文章名}`.
 * A Chinese author is one word. An English author keeps following name words, such as `Sherra Lee`, until the title starts.
 */
export function parseFileInput(input: string): ParsedFileInput | null {
    const text = input.trim();
    const match = text.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (!match || match.index === undefined) {
        return null;
    }
    const date = `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
    const before = text.slice(0, match.index).trim();
    const after = text.slice(match.index + match[0].length).trim();
    if (before) {
        return parseAuthorAndTitle(before, date, after);
    }
    const split = splitCategoryFromRest(after);
    if (!split) {
        return null;
    }
    return parseAuthorAndTitle(split.category, date, split.remainder);
}

/** Sheet form: `{y}-{m}-{d} {文章类别} {作者} {文章名}`, with no leading zeros on month or day. */
export function toSheetFileName(date: string, category: string, author: string, title: string): string {
    const match = date.trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    const categoryText = category.trim();
    const authorText = author.trim();
    const titleText = title.trim();
    if (!match || !categoryText || !authorText) {
        return '';
    }
    if (!titleText && !categoryHasNoArticleTitle(categoryText)) {
        return '';
    }
    const prefix = `${match[1]}-${Number(match[2])}-${Number(match[3])} ${categoryText} ${authorText}`;
    return titleText ? `${prefix} ${titleText}` : prefix;
}

function yearMonthInText(text: string): { year: string; month: string } | null {
    const match = text.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (!match) {
        return null;
    }
    return { year: match[1], month: String(Number(match[2])).padStart(2, '0') };
}

function categoryInFileName(fileName: string): string {
    const match = fileName.match(/\d{4}-\d{1,2}-\d{1,2}\s*(.*)$/);
    const rest = match?.[1]?.trim() ?? '';
    for (const name of knownCategoryNames()) {
        if (rest.toLowerCase().startsWith(name.toLowerCase())) {
            return canonicalChinese(name);
        }
    }
    return '';
}

/**
 * `YYYY-MM-{letter}{nn}`. `nn` is one past the larger of the highest existing slug
 * and the number of articles of the same category already in that month.
 */
export function nextSlug(date: string, category: string, rows: SlugSourceRow[]): string {
    const letter = categoryLetter(category);
    const when = yearMonthInText(date);
    if (!letter || !when) {
        return '';
    }
    const chinese = canonicalChinese(category);
    const prefix = `${when.year}-${when.month}-${letter}`;
    let maxNumber = 0;
    let count = 0;
    for (const row of rows) {
        const slug = (row.slug || '').trim();
        const slugMatch = slug.match(new RegExp(`^${prefix}(\\d+)$`, 'i'));
        if (slugMatch) {
            maxNumber = Math.max(maxNumber, Number(slugMatch[1]));
        }
        const rowChinese = canonicalChinese(row.文章类别 || '') || categoryInFileName(row.文件 || '');
        const rowWhen = yearMonthInText(row.文件 || '');
        if (rowChinese && rowChinese === chinese && rowWhen && rowWhen.year === when.year && rowWhen.month === when.month) {
            count += 1;
        }
    }
    const next = Math.max(maxNumber, count) + 1;
    return `${prefix}${String(next).padStart(2, '0')}`;
}
