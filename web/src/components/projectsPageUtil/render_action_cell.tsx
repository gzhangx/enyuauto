
import type { FreedCampComment, FreedCampLoginParams, ProjectTaskParams } from '../../../shared/freedcampTypes';
import { anyKeyUpdater, formatLocalDateYyyyMmDd, type DebugLog, type FreeCampAndUpdateOperations, type ICombinedOpsAndFreeCampData, deleteItemActionTask } from '../../../shared/main_ops';
import { getTaskIdColumnName, type IOperationWithLineNumber, type IOperationWithLineNumberAndParentTaskId, type IOpsConfig, type ISheetDataOps } from '../../../shared/opsTypes';
import type { ActionType } from '../../lib/api';
import { wordpressApi } from '../../lib/api';
import React, { useEffect, useRef, useState, type JSX } from 'react';
import { fetchOneDriveChildren } from '../MicrosoftOneDrivePage';
import { convertDocxToHtml, buildWpOutputPage } from '../../lib/docToWp';

/** Minimal DriveItem shape needed locally */
type MinDriveItem = { id: string; name: string; folder?: object; parentReference?: { driveId?: string } };
const SHOW_WP_Button = false;
async function fetchDriveItemChildren(token: string, driveId: string, itemId: string): Promise<MinDriveItem[]> {
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/children`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error((body as any)?.error?.message || res.statusText);
  }
  return ((await res.json()).value ?? []) as MinDriveItem[];
}

const PublishButton: React.FC<{
  p: IOperationWithLineNumberAndParentTaskId;
  msToken: string;
  wpToken?: string;
}> = ({ p, msToken, wpToken }) => {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [dialog, setDialog] = useState<{ wpBlocks: string; previewHtml: string; images: import('../../lib/docToWp').DocxImage[]; blobUrl: string } | null>(null);
  const [tab, setTab] = useState<'preview' | 'code'>('preview');
  const [copied, setCopied] = useState(false);
  const [uploadStatus, setUploadStatus] = useState<{ name: string; url?: string; error?: string }[]>([]);
  const [uploading, setUploading] = useState(false);

  const handlePublish = async () => {
    if (!p.mainFolder) { setError('No mainFolder for this project'); return; }
    setLoading(true);
    setError('');
    try {
      const articleChildren = await fetchOneDriveChildren(msToken, p.mainFolder) as MinDriveItem[];
      const publishFolder = articleChildren.find(item => item.folder && /4\s*publish/i.test(item.name));
      if (!publishFolder) throw new Error(`"4 Publish" folder not found in "${p['文件']}"`);
      const driveId = publishFolder.parentReference?.driveId;
      if (!driveId) throw new Error('Could not determine driveId for publish folder');
      const publishChildren = await fetchDriveItemChildren(msToken, driveId, publishFolder.id);
      const docxItem = publishChildren.find(item => /\.docx?$/i.test(item.name));
      if (!docxItem) throw new Error('No .docx file found in "4 Publish" folder');

      const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${docxItem.id}/content`;
      const res = await fetch(url, { headers: { Authorization: `Bearer ${msToken}` } });
      if (!res.ok) {
        const body = await res.json().catch(() => ({}));
        throw new Error((body as any)?.error?.message || res.statusText);
      }
      const arrayBuffer = await res.arrayBuffer();
      const converted = await convertDocxToHtml(arrayBuffer);
      const page = buildWpOutputPage(converted, wpToken);
      const blob = new Blob([page], { type: 'text/html;charset=utf-8' });
      setDialog({ ...converted, blobUrl: URL.createObjectURL(blob) });
      setTab('preview');
      setUploadStatus([]);
    } catch (err: any) {
      setError(err.message || String(err));
    } finally {
      setLoading(false);
    }
  };

  const handleCopy = async () => {
    if (!dialog) return;
    await navigator.clipboard.writeText(dialog.wpBlocks);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleUpload = async () => {
    if (!dialog || !wpToken) return;
    setUploading(true);
    const statuses: { name: string; url?: string; error?: string }[] = [];
    for (const img of dialog.images) {
      try {
        const data = await wordpressApi({
          subAction: 'uploadMedia',
          wpToken: 'Basic ' + btoa(wpToken),
          filename: img.name,
          mimeType: img.mimeType,
          b64: img.b64,
        });
        if (data.source_url) {
          statuses.push({ name: img.name, url: data.source_url });
        } else {
          statuses.push({ name: img.name, error: data.message || JSON.stringify(data) });
        }
      } catch (e: any) {
        statuses.push({ name: img.name, error: String(e) });
      }
      setUploadStatus([...statuses]);
    }
    setUploading(false);
  };

  const overlayStyle: React.CSSProperties = {
    position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 2000,
    display: 'flex', alignItems: 'flex-start', justifyContent: 'center', padding: '2rem 1rem', overflowY: 'auto',
  };
  const dlgStyle: React.CSSProperties = {
    background: '#fff', borderRadius: '8px', width: '100%', maxWidth: '860px',
    boxShadow: '0 8px 32px rgba(0,0,0,0.3)', display: 'flex', flexDirection: 'column', maxHeight: '90vh',
  };
  const tbStyle: React.CSSProperties = {
    display: 'flex', gap: '0.4rem', padding: '0.6rem 1rem', borderBottom: '1px solid #ddd',
    flexWrap: 'wrap', alignItems: 'center', background: '#f8f9fa', borderRadius: '8px 8px 0 0',
  };
  const btnBase = (bg: string, disabled = false): React.CSSProperties => ({
    padding: '0.35rem 0.9rem', border: 'none', borderRadius: '4px', cursor: disabled ? 'not-allowed' : 'pointer',
    fontSize: '0.85rem', color: '#fff', background: disabled ? '#aaa' : bg,
  });

  return (
    <span>
		  {SHOW_WP_Button &&<button
			  className="btn btn-create"
			  onClick={handlePublish}
			  disabled={loading}
			  style={{ background: '#e91e63', color: 'white', marginLeft: '4px' }}
		  >
			  {loading ? '…' : '发布 WP'}
		  </button>
		  }
      {error && <span style={{ color: 'red', fontSize: '11px', display: 'block' }}>{error}</span>}

      {dialog && (
        <div style={overlayStyle} onClick={e => { if (e.target === e.currentTarget) setDialog(null); }}>
          <div style={dlgStyle}>
            {/* Toolbar */}
            <div style={tbStyle}>
              <button style={btnBase(tab === 'preview' ? '#0078d4' : '#6c757d')} onClick={() => setTab('preview')}>Preview</button>
              <button style={btnBase(tab === 'code' ? '#495057' : '#6c757d')} onClick={() => setTab('code')}>WP Code</button>
              <button style={btnBase('#6f42c1')} onClick={handleCopy}>{copied ? 'Copied ✓' : 'Copy All'}</button>
              {dialog.images.length > 0 && (
                <button style={btnBase('#e91e63', uploading || !wpToken)} disabled={uploading || !wpToken} onClick={handleUpload} title={!wpToken ? 'No WP token configured' : undefined}>
                  {uploading ? 'Uploading…' : `Upload Images (${dialog.images.length})`}
                </button>
              )}
              <button style={btnBase('#17a2b8')} onClick={() => window.open(dialog.blobUrl, '_blank')}>Open in New Window</button>
              <button style={{ ...btnBase('#6c757d'), marginLeft: 'auto' }} onClick={() => setDialog(null)}>✕ Close</button>
            </div>

            {/* Upload status */}
            {uploadStatus.length > 0 && (
              <div style={{ padding: '0.4rem 1rem', fontSize: '0.8rem', borderBottom: '1px solid #eee' }}>
                {uploadStatus.map((s, i) => (
                  <div key={i} style={{ color: s.error ? 'red' : 'green' }}>
                    {s.error ? `✗ ${s.name}: ${s.error}` : <span>✓ <a href={s.url} target="_blank" rel="noreferrer">{s.name}</a></span>}
                  </div>
                ))}
              </div>
            )}

            {/* Panes */}
            <div style={{ overflowY: 'auto', flex: 1, padding: '1rem' }}>
              {tab === 'preview' && (
                <div style={{ maxWidth: '780px', margin: '0 auto', lineHeight: 1.7, fontSize: '1rem' }}
                  dangerouslySetInnerHTML={{ __html: dialog.previewHtml }} />
              )}
              {tab === 'code' && (
                <pre style={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word', fontSize: '0.82rem', margin: 0 }}>
                  {dialog.wpBlocks}
                </pre>
              )}
            </div>
          </div>
        </div>
      )}
    </span>
  );
};

function formatTaskWhen(ts?: number): string {
	if (!ts) return '—';
	return new Date(ts * 1000).toLocaleString();
}

function isPastDueOrSoon(dueMillis: number | undefined, days: number): boolean {
	if (typeof dueMillis !== 'number' || Number.isNaN(dueMillis)) return false;
	return dueMillis <= Date.now() + Math.max(0, days) * 24 * 60 * 60 * 1000;
}

function commentBody(comment: FreedCampComment): string {
	return (comment.description || comment.text || comment.comment || '').trim();
}

function commentAuthor(comment: FreedCampComment): string {
	return comment.created_by_fullname || comment.user_full_name || 'Unknown';
}

function asCommentList(value: unknown): FreedCampComment[] {
	if (Array.isArray(value)) return value as FreedCampComment[];
	if (value && typeof value === 'object') return Object.values(value as Record<string, FreedCampComment>);
	return [];
}

const TaskDetailButton: React.FC<{
	item: ProjectTaskParams;
	action: ActionType;
	daysShowAlertAfterComplete: number;
	freeCampOpsWithCache: FreeCampAndUpdateOperations;
	freedCampLoginParams: FreedCampLoginParams;
	children: JSX.Element;
}> = ({ item, action, daysShowAlertAfterComplete, freeCampOpsWithCache, freedCampLoginParams, children }) => {
	const [open, setOpen] = useState(false);
	const [detail, setDetail] = useState<ProjectTaskParams>(item);
	const [commentsStatus, setCommentsStatus] = useState<'idle' | 'loading' | 'ready' | 'error'>('idle');
	const [commentsError, setCommentsError] = useState('');
	const itemRef = useRef(item);
	const opsRef = useRef(freeCampOpsWithCache);
	const loginRef = useRef(freedCampLoginParams);
	itemRef.current = item;
	opsRef.current = freeCampOpsWithCache;
	loginRef.current = freedCampLoginParams;
	useEffect(() => {
		if (!open) return;
		const current = itemRef.current;
		setDetail(current);
		if (!current.id) {
			setCommentsStatus('ready');
			return;
		}
		let cancelled = false;
		setCommentsStatus('loading');
		setCommentsError('');
		const timeout = window.setTimeout(() => {
			if (cancelled) return;
			setCommentsStatus('error');
			setCommentsError('Timed out loading comments.');
		}, 8000);
		(async () => {
			try {
				const token = await opsRef.current.getFreedCampToken(loginRef.current, false);
				const processor = opsRef.current.getFreedCampProcessor(token);
				const full = await processor.getTask(current.id!);
				if (cancelled) return;
				window.clearTimeout(timeout);
				const comments = asCommentList(full.comments);
				setDetail({ ...current, ...full, comments });
				setCommentsStatus('ready');
			} catch (error: unknown) {
				if (cancelled) return;
				window.clearTimeout(timeout);
				setCommentsStatus('error');
				setCommentsError((error as Error)?.message || String(error));
			}
		})();
		return () => {
			cancelled = true;
			window.clearTimeout(timeout);
		};
	}, [open, item.id]);
	const assigned = detail.assigned_to_id && detail.assigned_to_id !== '0'
		? (detail.assigned_to_fullname || detail.assigned_to_id)
		: 'Unassigned';
	const dueSoon = isPastDueOrSoon(detail.due_ts ? detail.due_ts * 1000 : undefined, daysShowAlertAfterComplete);
	const dueText = detail.due_ts ? formatTaskWhen(detail.due_ts) : (detail.due_date || '—');
	const comments = detail.comments || item.comments || [];
	return (
		<>
			<button
				type="button"
				className="fc-task-link"
				title="Show task details"
				onClick={(event) => {
					event.stopPropagation();
					setOpen(true);
				}}
			>
				{children}
			</button>
			{open && (
				<div
					className="fc-task-overlay"
					onClick={(event) => {
						if (event.target === event.currentTarget) setOpen(false);
					}}
				>
					<div className="fc-task-dialog" onClick={(event) => event.stopPropagation()}>
						<div className="fc-task-dialog-head">
							<h2>{action} task</h2>
							<button type="button" className="fc-task-close" onClick={() => setOpen(false)}>Close</button>
						</div>
						<dl className="fc-task-meta">
							<div><dt>Title</dt><dd>{detail.title || '—'}</dd></div>
							<div><dt>Task ID</dt><dd>{detail.id || '—'}</dd></div>
							<div><dt>Status</dt><dd>{detail.status_title || '—'}</dd></div>
							<div><dt>Assigned</dt><dd>{assigned}</dd></div>
							<div><dt>Created</dt><dd>{formatTaskWhen(detail.created_ts)}</dd></div>
							<div>
								<dt>Due</dt>
								<dd style={{ color: dueSoon ? 'red' : undefined, fontWeight: dueSoon ? 'bold' : undefined }}>{dueText}</dd>
							</div>
							<div><dt>Completed</dt><dd>{formatTaskWhen(detail.completed_ts)}</dd></div>
							<div><dt>Priority</dt><dd>{detail.priority_title || '—'}</dd></div>
						</dl>
						<div className="fc-task-desc-label">Description</div>
						<div
							className="fc-task-desc"
							dangerouslySetInnerHTML={{ __html: (detail.description_processed || detail.description || '').trim() || 'No description on this task.' }}
						/>
						<div className="fc-task-desc-label">Comments</div>
						{commentsStatus === 'loading' && <div className="fc-task-comments-status">Loading comments…</div>}
						{commentsStatus === 'error' && <div className="fc-task-comments-status">{commentsError || 'Could not load comments.'}</div>}
						{commentsStatus === 'ready' && comments.length === 0 && (
							<div className="fc-task-comments-status">No comments on this task.</div>
						)}
						{comments.length > 0 && (
							<div className="fc-task-comments">
								{comments.map((comment, index) => (
									<div key={comment.id || index} className="fc-task-comment">
										<div className="fc-task-comment-meta">
											<strong>{commentAuthor(comment)}</strong>
											<span>{formatTaskWhen(comment.created_ts)}</span>
										</div>
										<div
											className="fc-task-comment-body"
											dangerouslySetInnerHTML={{ __html: commentBody(comment) || '—' }}
										/>
									</div>
								))}
							</div>
						)}
					</div>
				</div>
			)}
		</>
	);
};

const HoverTooltip: React.FC<{ content: JSX.Element | string; children: JSX.Element | string }> = ({ content, children }) => {
	const [show, setShow] = useState(false);
	return (
		<span style={{ position: 'relative', display: 'inline-block' }} onMouseEnter={() => setShow(true)} onMouseLeave={() => setShow(false)}>
			{children}
			{show && (
				<div style={{ position: 'absolute', zIndex: 2200, bottom: '100%', left: 0, marginBottom: 8 }}>
					<div style={{ background: '#fff', color: '#111', padding: '6px 8px', borderRadius: 6, whiteSpace: 'pre-wrap', fontSize: 12, maxWidth: 320, boxShadow: '0 4px 12px rgba(0,0,0,0.12)', border: '1px solid #e6e6e6' }}>
						{typeof content === 'string' ? <span>{content}</span> : content}
					</div>
				</div>
			)}
		</span>
	);
};


// Render a cell for syncing sheet with FreedCamp if data exists in FreedCamp but not in p
export function renderSyncActionCell(
	p: IOperationWithLineNumberAndParentTaskId,
	freedCampLoginParams: FreedCampLoginParams,
    deps: RenderActionCellDeps,
    desc: string,
) {
	const {
		sheetOpsRef,
		combined,
		//freeCampOpsWithCache,
		//completeDateUpdater,
		doLog,
		fetchData,
		setErrorDialog,		
	} = deps;

    if (!combined) {
        return <div>Waiting for Data</div>;
    }

    const ready = sheetOpsRef.current && combined;
    let tooltipText = 'Exists in FreedCamp but not in sheet.';
    if (!ready) {
        tooltipText += ' Sheet operations not ready.!!!!!';
        return <div title={tooltipText}>Waiting for Data</div>;
	}
	const syncData = p.syncFreeCampToSheetData;
	if (!syncData) {
		tooltipText += 'No data to sync.';
		return <div title={tooltipText}></div>;
	}
	const { parts, updates } = syncData;

	// Build tooltip text with FreedCamp info
	
	tooltipText += '\n' + parts.join('\n');

	// Button to sync sheet with FreedCamp
	return (
		<button
			className="btn btn-create"
			style={{ marginLeft: '5px', background: updates.length === 0 ? '#9e9e9e' : '#ff9800', color: 'white' }}
			title={tooltipText}
			disabled={updates.length === 0}
			onClick={async () => {
				if (sheetOpsRef.current && combined) {
					try {
						// Here you would update the sheet to match FreedCamp item
						// For example, set the cell to the FreedCamp ID or other info
						// This is a placeholder for the actual update logic
						// You may want to call a function similar to completeDateUpdater
						// For now, just log and refresh
						//doLog(`Syncing sheet for action ${action} with FreedCamp item for "${p.文件}"`);
						// Example: update the sheet with the FreedCamp ID
						// await updateSheetWithFreedCampId(sheetOpsRef.current, opsConfig, action, p, freedCampItem.id, logParam);						
						// For now, just refresh
						for (const update of updates) {
							p[update.sheetCol as ActionType] = update.value;
							await anyKeyUpdater(sheetOpsRef.current, combined.opsConfig, update.sheetCol as ActionType, p, { doLog });
						}
						await fetchData(freedCampLoginParams);
					} catch (error: any) {
						console.error('Error syncing sheet:', error);
						setErrorDialog({ show: true, message: `Failed to sync sheet:\n${error.message || String(error)}` });
					}
				}
			}}
		>
			{desc} { updates.length}
		</button>
	);
}


export type RenderActionCellDeps = {
	sheetOpsRef: React.RefObject<any>;
	//opsConfig: IOpsConfig | null;
	combined: ICombinedOpsAndFreeCampData | null;
	freeCampOpsWithCache: FreeCampAndUpdateOperations;
	freedCampLoginParams: FreedCampLoginParams;
	deleteItemActionTask: typeof deleteItemActionTask;
	completeDateUpdater: (ops: ISheetDataOps, opsConfig: IOpsConfig, key: ActionType, item: IOperationWithLineNumber, val: string, log: DebugLog) => Promise<void>;
	doLog: (msg: string) => void;
	fetchData: (freedCampCredentials: FreedCampLoginParams) => Promise<void>;
	setErrorDialog: (v: { show: boolean; message: string }) => void;
	//logParam: { doLog: (msg: string) => void };
	daysShowAlertAfterComplete: number;
	msToken?: string | null;
	wpToken?: string;
};

export function renderActionCell(
	p: IOperationWithLineNumberAndParentTaskId,
	action: ActionType,
	deps: RenderActionCellDeps,
	canCreate = false,
) {
	const {
		sheetOpsRef,
		combined,
		freeCampOpsWithCache,
		deleteItemActionTask,
		freedCampLoginParams,
		//completeDateUpdater,
		doLog,
		//fetchData,
		//setErrorDialog,
		daysShowAlertAfterComplete,
		msToken,
		wpToken,
	} = deps;

	const taskId = p[getTaskIdColumnName(action)];
	const freedCampItem = p[`${action} FreeCamp Item`];

	// Build tooltip text
	//if (p.文件.includes('宣教无国界')) {
		//console.log('fond 2026-1-25 宣教无国界 陈默 三福培训分享- ', p.文件)
	//}
	let tooltipText = '';
	if (freedCampItem) {
		const parts: string[] = [];
		if (freedCampItem.completed_ts && freedCampItem.status_title === 'Completed') {
			const completedDate = new Date(freedCampItem.completed_ts * 1000);
			parts.push(`Completed: ${completedDate.toLocaleString()}`);
		}
		if (freedCampItem.status_title) {
			parts.push(`Status: ${freedCampItem.status_id}:${freedCampItem.status_title}`);
		}
		tooltipText = parts.join('\n');
	} else {
		tooltipText = 'No corresponding FreedCamp item found.';
	}

	
	if (taskId === 'TaskIsEnglishAndIsDisabledForEnglish') {
		return <span style={{ color: 'orange', fontWeight: 'bold' }} title={tooltipText}>No English</span>;
	}
	if (taskId === 'done') {
		return <span style={{ color: 'green', fontWeight: 'bold' }} title={tooltipText}>Done</span>;
	}

	const hasCompletedDate = freedCampItem?.status_title === 'Completed' ? freedCampItem?.completed_ts : null;
	//if (hasCompletedDate) {
	//	console.log(freedCampItem, 'debugremove freedCampItem is in render_action_cell.tsx')
	//}

	const hasAssignedTo = freedCampItem?.assigned_to_id && freedCampItem?.assigned_to_id !== '0';

	const freedCampTaskId = freedCampItem?.id;
	const openTask = (node: JSX.Element) => (
		freedCampItem
			? (
				<TaskDetailButton
					item={freedCampItem}
					action={action}
					daysShowAlertAfterComplete={daysShowAlertAfterComplete}
					freeCampOpsWithCache={freeCampOpsWithCache}
					freedCampLoginParams={freedCampLoginParams}
				>
					{node}
				</TaskDetailButton>
			)
			: node
	);
	if (!taskId && (!hasCompletedDate && !hasAssignedTo)) {
		if (canCreate) return null;
		return <span style={{ color: 'orange' }} title={tooltipText}>NA { freedCampTaskId}</span>;
	}
	let dateEditorDsp: JSX.Element | null = null;
	if (hasCompletedDate) {
		const isRecent = daysShowAlertAfterComplete > 0 &&
			hasCompletedDate * 1000 >= Date.now() - daysShowAlertAfterComplete * 24 * 60 * 60 * 1000;
		const color = isRecent ? 'red' : 'green';
		dateEditorDsp = openTask(
			<span
				style={{ color, fontWeight: 'bold' }}
				title={freedCampItem?.assigned_to_fullname || undefined}
			>
				{formatLocalDateYyyyMmDd(hasCompletedDate)}
			</span>
		);
	}
	let assignedEditorDsp: JSX.Element | null = null;
	if (hasAssignedTo && !hasCompletedDate) {
		const rawDueDateValue = freedCampItem?.due_ts ? freedCampItem.due_ts * 1000 : undefined;
		let sheetDueMillis: number | undefined = undefined;
		const sheetDueRaw = (p[`${action} Due Date` as keyof IOperationWithLineNumberAndParentTaskId] as unknown as string) || '';
		if (!rawDueDateValue && sheetDueRaw) {
			const parsed = Date.parse(sheetDueRaw);
			if (!isNaN(parsed)) sheetDueMillis = parsed;
		}
		const dueMillis = rawDueDateValue ?? sheetDueMillis;
		const isDueSoon = isPastDueOrSoon(dueMillis, daysShowAlertAfterComplete);
		const dueDisplay = rawDueDateValue ? formatLocalDateYyyyMmDd(rawDueDateValue / 1000) : (sheetDueRaw || '');
		assignedEditorDsp = (
			<HoverTooltip
				content={
					<>
						<div>Assigned to: {freedCampItem?.assigned_to_fullname}</div>
						<div>Due: <span style={{ color: isDueSoon ? 'red' : undefined, fontWeight: isDueSoon ? 'bold' : 'normal' }}>{dueDisplay || 'N/A'}</span></div>
					</>
				}
			>
				{openTask(
					<span
						style={{
							fontWeight: isDueSoon ? 'bold' : 'normal',
							color: isDueSoon ? 'red' : undefined,
						}}
					>
						{freedCampItem?.assigned_to_fullname}
					</span>
				)}
			</HoverTooltip>
		);
	}
	return (
		<>
			{ !hasAssignedTo && !hasCompletedDate && <button
				className="btn btn-delete"
				title={tooltipText}
				onClick={async () => {
					if (sheetOpsRef.current && combined) {
							await deleteItemActionTask(sheetOpsRef.current, freeCampOpsWithCache, freedCampLoginParams, combined.opsConfig, p, action, { doLog });
					}
				}}
			>
				Delete {p[getTaskIdColumnName(action)]}
			</button>
			}
			{
				assignedEditorDsp
			}
			{
				dateEditorDsp
			}
			{action === '发布' && msToken && (
				<PublishButton p={p} msToken={msToken} wpToken={wpToken} />
			)}
		</>
	);
}
