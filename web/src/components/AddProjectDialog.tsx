import { useMemo, useState } from 'react';
import {
    categoryLetter,
    findAuthorEmail,
    nextSlug,
    parseFileInput,
    toSheetFileName,
    type SlugSourceRow,
} from '../lib/newProjectForm';

export type NewProjectValues = {
    文件: string;
    mainFolder: string;
    作者: string;
    称呼: string;
    作者电邮: string;
    文章类别: string;
    文章名: string;
    文章链接: string;
    slug: string;
};

type AddProjectDialogProps = {
    categories: string[];
    titles: string[];
    projects: SlugSourceRow[];
    saving: boolean;
    onClose: () => void;
    onSave: (values: NewProjectValues) => void;
};

const emptyForm = {
    fileInput: '',
    mainFolder: '',
    author: '',
    titleOfAddress: '',
    email: '',
    category: '',
    title: '',
    link: '',
    slug: '',
    date: '',
};

export function AddProjectDialog({ categories, titles, projects, saving, onClose, onSave }: AddProjectDialogProps) {
    const [form, setForm] = useState(emptyForm);
    const [parseNote, setParseNote] = useState('');
    const [emailNote, setEmailNote] = useState('');

    const options = useMemo(() => {
        if (form.category && !categories.some(category => category.toLowerCase() === form.category.toLowerCase())) {
            return [...categories, form.category];
        }
        return categories;
    }, [categories, form.category]);

    const addressOptions = titles;

    const sheetFileName = toSheetFileName(form.date, form.category, form.author, form.title);

    const onFileInput = (fileInput: string) => {
        const parsed = parseFileInput(fileInput);
        if (!parsed) {
            setForm(prev => ({ ...prev, fileInput, date: '' }));
            setParseNote(fileInput.trim() ? 'Could not read 文件. Put the category, date, author, and article name in that order, separated by spaces.' : '');
            return;
        }
        setParseNote('');
        const category = categories.find(item => item.toLowerCase() === parsed.category.toLowerCase()) ?? parsed.category;
        setForm(prev => ({
            ...prev,
            fileInput,
            category,
            author: parsed.author,
            title: parsed.title,
            date: parsed.date,
            slug: nextSlug(parsed.date, category, projects),
        }));
    };

    const searchEmail = () => {
        const author = form.author.trim();
        if (!author) {
            setEmailNote('Fill in 作者 first.');
            return;
        }
        const email = findAuthorEmail(author, projects);
        if (!email) {
            setEmailNote(`No earlier email for ${author}.`);
            return;
        }
        setForm(prev => ({ ...prev, email }));
        setEmailNote('');
    };

    const onCategory = (category: string) => {
        setForm(prev => ({
            ...prev,
            category,
            slug: nextSlug(prev.date, category, projects),
        }));
    };

    const canSave = Boolean(sheetFileName && form.slug.trim() && !saving);

    const save = () => {
        if (!canSave) {
            return;
        }
        onSave({
            文件: sheetFileName,
            mainFolder: form.mainFolder.trim(),
            作者: form.author.trim(),
            称呼: form.titleOfAddress.trim(),
            作者电邮: form.email.trim(),
            文章类别: form.category.trim(),
            文章名: form.title.trim(),
            文章链接: form.link.trim(),
            slug: form.slug.trim(),
        });
    };

    return (
        <div className="add-project-backdrop" onClick={onClose}>
            <div className="add-project-dialog" onClick={(event) => event.stopPropagation()}>
                <div className="add-project-header">
                    <div>
                        <h2>Add New Project</h2>
                        <p>Leave 文件 to fill the article, author, and slug.</p>
                    </div>
                    <button type="button" className="add-project-close" onClick={onClose} disabled={saving} aria-label="Close">×</button>
                </div>

                <div className="add-project-body">
                <section>
                    <h3>稿件</h3>
                    <label className="add-project-field add-project-span">
                        文件
                        <input
                            value={form.fileInput}
                            autoFocus
                            onChange={(event) => {
                                const value = event.target.value;
                                setForm(prev => ({ ...prev, fileInput: value }));
                                if (parseNote) setParseNote('');
                            }}
                            onBlur={(event) => onFileInput(event.target.value)}
                        />
                    </label>
                    {parseNote && <div className="add-project-note add-project-span">{parseNote}</div>}
                    <div className="add-project-preview add-project-span">
                        <span>Saved 文件</span>
                        <strong>{sheetFileName || '—'}</strong>
                    </div>
                    <label className="add-project-field add-project-span">
                        mainFolder
                        <input value={form.mainFolder} onChange={(event) => setForm(prev => ({ ...prev, mainFolder: event.target.value }))} />
                    </label>
                </section>

                <section>
                    <h3>文章</h3>
                    <label className="add-project-field">
                        文章类别
                        <select value={form.category} onChange={(event) => onCategory(event.target.value)}>
                            <option value="">Select a category</option>
                            {options.map(option => (
                                <option key={option} value={option}>{option}</option>
                            ))}
                        </select>
                    </label>
                    <label className="add-project-field">
                        slug
                        <input value={form.slug} onChange={(event) => setForm(prev => ({ ...prev, slug: event.target.value }))} />
                    </label>
                    {!form.category || categoryLetter(form.category) ? null : (
                        <div className="add-project-note add-project-span">This category has no slug letter. Type the slug yourself.</div>
                    )}
                    <label className="add-project-field add-project-span">
                        文章名
                        <input value={form.title} onChange={(event) => setForm(prev => ({ ...prev, title: event.target.value }))} />
                    </label>
                    <label className="add-project-field add-project-span">
                        文章链接
                        <input value={form.link} onChange={(event) => setForm(prev => ({ ...prev, link: event.target.value }))} />
                    </label>
                </section>

                <section>
                    <h3>作者</h3>
                    <label className="add-project-field">
                        作者
                        <input value={form.author} onChange={(event) => setForm(prev => ({ ...prev, author: event.target.value }))} />
                    </label>
                    <label className="add-project-field">
                        称呼
                        <input list="author-title-options" value={form.titleOfAddress} onChange={(event) => setForm(prev => ({ ...prev, titleOfAddress: event.target.value }))} />
                        <datalist id="author-title-options">
                            {addressOptions.map(option => <option key={option} value={option} />)}
                        </datalist>
                    </label>
                    <div className="add-project-field add-project-span">
                        <span>作者电邮</span>
                        <div className="add-project-email">
                            <input value={form.email} onChange={(event) => { setForm(prev => ({ ...prev, email: event.target.value })); setEmailNote(''); }} />
                            <button type="button" onClick={searchEmail}>Search</button>
                        </div>
                        {emailNote && <div className="add-project-note">{emailNote}</div>}
                    </div>
                </section>
                </div>

                <div className="add-project-actions">
                    <button type="button" className="add-project-cancel" onClick={onClose} disabled={saving}>Cancel</button>
                    <button type="button" className="btn btn-create" onClick={save} disabled={!canSave}>{saving ? 'Saving…' : 'Save'}</button>
                </div>
            </div>
            <style>{`
                .add-project-backdrop {
                    position: fixed;
                    inset: 0;
                    background: rgba(33, 37, 41, 0.45);
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    z-index: 2100;
                    padding: 1.25rem;
                }
                .add-project-dialog {
                    display: flex;
                    flex-direction: column;
                    width: min(760px, 100%);
                    max-height: 90vh;
                    overflow: hidden;
                    background: #fff;
                    color: #212529;
                    border-radius: 12px;
                    box-shadow: 0 12px 40px rgba(15, 40, 70, 0.35);
                    text-align: left;
                }
                .add-project-header {
                    position: relative;
                    background: linear-gradient(135deg, #1f4b7a, #2f6aa8);
                    color: #fff;
                    padding: 1.1rem 3rem;
                    text-align: center;
                }
                .add-project-header h2 {
                    margin: 0;
                    font-size: 1.35rem;
                    letter-spacing: 0.02em;
                    line-height: 1.3;
                }
                .add-project-header p {
                    margin: 6px 0 0;
                    font-size: 0.95rem;
                    opacity: 0.92;
                }
                .add-project-close {
                    position: absolute;
                    top: 0.7rem;
                    right: 0.9rem;
                    border: none;
                    background: transparent;
                    color: rgba(255, 255, 255, 0.85);
                    font-size: 1.6rem;
                    line-height: 1;
                    padding: 0 0.15rem;
                    cursor: pointer;
                }
                .add-project-close:hover {
                    color: #fff;
                }
                .add-project-body {
                    flex: 1;
                    overflow-y: auto;
                    background: #f4f7fb;
                    padding: 1rem 1.25rem;
                    display: flex;
                    flex-direction: column;
                    gap: 10px;
                }
                .add-project-dialog section {
                    display: grid;
                    grid-template-columns: 1fr 1fr;
                    gap: 0.75rem 1rem;
                    margin: 0;
                    padding: 0.85rem 1rem 1rem;
                    background: #fff;
                    border: 1px solid #dde5ee;
                    border-left: 4px solid #2f6aa8;
                    border-radius: 8px;
                    box-shadow: 0 2px 8px rgba(31, 75, 122, 0.08);
                }
                .add-project-dialog h3 {
                    grid-column: 1 / -1;
                    justify-self: start;
                    margin: 0;
                    padding: 2px 12px;
                    border-radius: 999px;
                    background: #1f4b7a;
                    color: #fff;
                    font-size: 0.85rem;
                    letter-spacing: 0.04em;
                }
                .add-project-field {
                    display: flex;
                    flex-direction: column;
                    gap: 0.3rem;
                    font-size: 0.86rem;
                    font-weight: 600;
                    color: #1f4b7a;
                }
                .add-project-span { grid-column: 1 / -1; }
                .add-project-dialog input,
                .add-project-dialog select {
                    width: 100%;
                    box-sizing: border-box;
                    padding: 0.55rem 0.65rem;
                    border: 1px solid #c5d0dc;
                    border-radius: 6px;
                    font-size: 0.95rem;
                    font-weight: 400;
                    background: #fff;
                    color: #212529;
                    color-scheme: light;
                }
                .add-project-dialog input:focus,
                .add-project-dialog select:focus {
                    outline: none;
                    border-color: #2f6aa8;
                    box-shadow: 0 0 0 2px rgba(47, 106, 168, 0.2);
                }
                .add-project-email {
                    display: flex;
                    gap: 0.5rem;
                }
                .add-project-email button {
                    padding: 0.55rem 0.9rem;
                    border-radius: 6px;
                    border: 1px solid #2f6aa8;
                    background: #e3edf8;
                    color: #1f4b7a;
                    font-weight: 600;
                    cursor: pointer;
                    white-space: nowrap;
                }
                .add-project-email button:hover {
                    background: #d3e3f5;
                }
                .add-project-preview {
                    background: #eef4fb;
                    border: 1px solid #dde5ee;
                    border-radius: 8px;
                    padding: 0.7rem 0.8rem;
                }
                .add-project-preview span {
                    display: block;
                    color: #6c757d;
                    font-size: 0.78rem;
                    margin-bottom: 0.15rem;
                }
                .add-project-preview strong {
                    font-weight: 600;
                    color: #1f4b7a;
                    line-height: 1.4;
                }
                .add-project-note {
                    color: #8a4b00;
                    background: #fff7e6;
                    border: 1px solid #f5c26b;
                    border-radius: 6px;
                    font-weight: 400;
                    padding: 0.4rem 0.6rem;
                    font-size: 0.82rem;
                }
                .add-project-actions {
                    display: flex;
                    justify-content: flex-end;
                    gap: 10px;
                    padding: 0.85rem 1.25rem;
                    border-top: 1px solid #dde5ee;
                    background: #fff;
                }
                .add-project-cancel {
                    padding: 8px 18px;
                    border-radius: 6px;
                    border: 1px solid #c5d0dc;
                    background: #fff;
                    color: #374151;
                    cursor: pointer;
                    font-size: 14px;
                }
                .add-project-cancel:hover {
                    background: #f1f4f8;
                }
                .add-project-actions .btn:disabled {
                    opacity: 0.55;
                    cursor: not-allowed;
                }
                @media (max-width: 680px) {
                    .add-project-dialog section { grid-template-columns: 1fr; }
                    .add-project-span { grid-column: auto; }
                }
            `}</style>
        </div>
    );
}
