import { useMemo, useState } from 'react';
import {
    categoryLetter,
    categoryOptions,
    findAuthorEmail,
    nextSlug,
    parseFileInput,
    titleOptions,
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

export function AddProjectDialog({ projects, saving, onClose, onSave }: AddProjectDialogProps) {
    const [form, setForm] = useState(emptyForm);
    const [parseNote, setParseNote] = useState('');
    const [emailNote, setEmailNote] = useState('');

    const options = useMemo(() => {
        const fromRows = projects.map(row => row.文章类别 || '');
        const base = categoryOptions(fromRows);
        if (form.category && !base.includes(form.category)) {
            return [...base, form.category];
        }
        return base;
    }, [projects, form.category]);

    const addressOptions = useMemo(
        () => titleOptions(projects.map(row => row.称呼 || '')),
        [projects],
    );

    const sheetFileName = toSheetFileName(form.date, form.category, form.author, form.title);

    const onFileInput = (fileInput: string) => {
        const parsed = parseFileInput(fileInput);
        if (!parsed) {
            setForm(prev => ({ ...prev, fileInput, date: '' }));
            setParseNote(fileInput.trim() ? 'Could not read 文件. Put the category, date, author, and article name in that order, separated by spaces.' : '');
            return;
        }
        setParseNote('');
        setForm(prev => ({
            ...prev,
            fileInput,
            category: parsed.category,
            author: parsed.author,
            title: parsed.title,
            date: parsed.date,
            slug: nextSlug(parsed.date, parsed.category, projects),
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

    const fieldStyle = { display: 'flex', flexDirection: 'column' as const, gap: '0.35rem', textAlign: 'left' as const };
    const inputStyle = {
        width: '100%',
        boxSizing: 'border-box' as const,
        padding: '0.5rem',
        border: '1px solid #ced4da',
        borderRadius: '4px',
        fontSize: '14px',
        backgroundColor: '#fff',
        color: '#212529',
        colorScheme: 'light' as const,
    };

    return (
        <div
            style={{
                position: 'fixed',
                top: 0,
                left: 0,
                right: 0,
                bottom: 0,
                backgroundColor: 'rgba(0, 0, 0, 0.5)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                zIndex: 2100,
                padding: '1rem',
            }}
            onClick={onClose}
        >
            <div
                style={{
                    backgroundColor: 'white',
                    padding: '1.5rem',
                    borderRadius: '8px',
                    width: 'min(640px, 100%)',
                    maxHeight: '90vh',
                    overflow: 'auto',
                    textAlign: 'left',
                    color: '#212529',
                }}
                onClick={(event) => event.stopPropagation()}
            >
                <h2 style={{ margin: '0 0 1rem 0' }}>Add a New project</h2>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '0.85rem' }}>
                    <label style={fieldStyle}>
                        文件
                        <input
                            style={inputStyle}
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
                    {parseNote && <div style={{ color: '#856404', fontSize: '13px' }}>{parseNote}</div>}
                    <label style={fieldStyle}>
                        mainFolder
                        <input style={inputStyle} value={form.mainFolder} onChange={(event) => setForm(prev => ({ ...prev, mainFolder: event.target.value }))} />
                    </label>
                    <label style={fieldStyle}>
                        作者
                        <input style={inputStyle} value={form.author} onChange={(event) => setForm(prev => ({ ...prev, author: event.target.value }))} />
                    </label>
                    <label style={fieldStyle}>
                        称呼
                        <input style={inputStyle} list="author-title-options" value={form.titleOfAddress} onChange={(event) => setForm(prev => ({ ...prev, titleOfAddress: event.target.value }))} />
                        <datalist id="author-title-options">
                            {addressOptions.map(option => <option key={option} value={option} />)}
                        </datalist>
                    </label>
                    <div style={fieldStyle}>
                        <span>作者电邮</span>
                        <div style={{ display: 'flex', gap: '0.5rem' }}>
                            <input style={{ ...inputStyle, flex: 1 }} value={form.email} onChange={(event) => { setForm(prev => ({ ...prev, email: event.target.value })); setEmailNote(''); }} />
                            <button type="button" onClick={searchEmail} style={{ padding: '0.5rem 0.9rem', borderRadius: '4px', border: '1px solid #ced4da', background: 'white', color: '#212529', cursor: 'pointer', whiteSpace: 'nowrap' }}>
                                Search
                            </button>
                        </div>
                        {emailNote && <div style={{ color: '#856404', fontSize: '13px' }}>{emailNote}</div>}
                    </div>
                    <label style={fieldStyle}>
                        文章类别
                        <select style={inputStyle} value={form.category} onChange={(event) => onCategory(event.target.value)}>
                            <option value="">Select a category</option>
                            {options.map(option => (
                                <option key={option} value={option}>{option}</option>
                            ))}
                        </select>
                    </label>
                    {!form.category || categoryLetter(form.category) ? null : (
                        <div style={{ color: '#856404', fontSize: '13px' }}>This category has no slug letter. Type the slug yourself.</div>
                    )}
                    <label style={fieldStyle}>
                        文章名
                        <input style={inputStyle} value={form.title} onChange={(event) => setForm(prev => ({ ...prev, title: event.target.value }))} />
                    </label>
                    <label style={fieldStyle}>
                        文章链接
                        <input style={inputStyle} value={form.link} onChange={(event) => setForm(prev => ({ ...prev, link: event.target.value }))} />
                    </label>
                    <label style={fieldStyle}>
                        slug
                        <input
                            style={inputStyle}
                            value={form.slug}
                            onChange={(event) => setForm(prev => ({ ...prev, slug: event.target.value }))}
                        />
                    </label>
                    <div style={{ background: '#f8f9fa', borderRadius: '4px', padding: '0.75rem', fontSize: '14px' }}>
                        <div style={{ color: '#6c757d', marginBottom: '0.25rem' }}>Saved 文件</div>
                        <div>{sheetFileName || '—'}</div>
                    </div>
                </div>
                <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '0.5rem', marginTop: '1.25rem' }}>
                    <button type="button" onClick={onClose} disabled={saving} style={{ padding: '0.5rem 1rem', borderRadius: '4px', border: '1px solid #ced4da', background: 'white', cursor: 'pointer' }}>
                        Cancel
                    </button>
                    <button type="button" className="btn btn-create" onClick={save} disabled={!canSave}>
                        {saving ? 'Saving…' : 'Save'}
                    </button>
                </div>
            </div>
        </div>
    );
}
