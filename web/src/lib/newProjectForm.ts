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
    { chinese: '宣教无国界', english: 'Missions Beyond Borders', letter: 'x' },
    { chinese: '佳作分享', english: 'Shared Book Passages', letter: 'j' },
];

const EXTRA_ENGLISH_ALIASES: { english: string; chinese: string }[] = [
    { english: 'Pastoral & Elder Messages', chinese: '牧者的话' },
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
        if (alias.chinese === trimmed || alias.english.toLowerCase() === lower) {
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

/** 佳作分享 is a book passage, so it has no 文章名. */
export function categoryHasNoArticleTitle(category: string): boolean {
    return canonicalChinese(category) === '佳作分享';
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

/**
 * Input form: `{文章类别} {date} {作者} {文章名}`.
 * A Chinese author is one word. An English author keeps following name words, such as `Sherra Lee`, until the title starts.
 */
export function parseFileInput(input: string): ParsedFileInput | null {
    const text = input.trim();
    const match = text.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (!match || match.index === undefined) {
        return null;
    }
    const category = text.slice(0, match.index).trim();
    const tokens = text.slice(match.index + match[0].length).trim().split(/\s+/).filter(Boolean);
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
    const date = `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
    return { category, date, author, title };
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
    const names = CATEGORY_ALIASES
        .flatMap(alias => [alias.chinese, alias.english, ...EXTRA_ENGLISH_ALIASES.filter(extra => extra.chinese === alias.chinese).map(extra => extra.english)])
        .sort((a, b) => b.length - a.length);
    for (const name of names) {
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
