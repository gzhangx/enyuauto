
import type { FreedCampLoginParams } from '../../../shared/freedcampTypes';
import { anyKeyUpdater, formatLocalDateYyyyMmDd, type DebugLog, type ICombinedOpsAndFreeCampData, deleteItemActionTask } from '../../../shared/main_ops';
import { getTaskIdColumnName, type IOperationWithLineNumber, type IOperationWithLineNumberAndParentTaskId, type IOpsConfig, type ISheetDataOps } from '../../../shared/opsTypes';
import type { ActionType } from '../../lib/api';
import { wordpressApi } from '../../lib/api';
import React, { useState, type JSX } from 'react';
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

const HoverTooltip: React.FC<{ content: JSX.Element | string; children: JSX.Element | string }> = ({ content, children }) => {
	const [show, setShow] = useState(false);
	return (
		<span style={{ position: 'relative', display: 'inline-block' }} onMouseEnter={() => setShow(true)} onMouseLeave={() => setShow(false)}>
			{children}
			{show && (
				<div style={{ position: 'absolute', zIndex: 2200, bottom: '100%', left: 0, marginBottom: 8 }}>
					<div style={{ background: '#222', color: '#fff', padding: '6px 8px', borderRadius: 6, whiteSpace: 'pre-wrap', fontSize: 12, maxWidth: 320 }}>
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
	freeCampOpsWithCache: any;
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
	deps: RenderActionCellDeps
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
	if (!taskId && (!hasCompletedDate && !hasAssignedTo)) {
		return <span style={{ color: 'orange' }} title={tooltipText}>NA { freedCampTaskId}</span>;
	}
	let dateEditorDsp: JSX.Element | null = null;
	if (hasCompletedDate) {
		const isRecent = daysShowAlertAfterComplete > 0 &&
			hasCompletedDate * 1000 >= Date.now() - daysShowAlertAfterComplete * 24 * 60 * 60 * 1000;
		const color = isRecent ? 'red' : 'green';
		dateEditorDsp = (
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
		const now = Date.now();
		const threshold = daysShowAlertAfterComplete * 24 * 60 * 60 * 1000;
		const isDueSoon = typeof dueMillis === 'number' && dueMillis <= (now + threshold);
		const dueDisplay = rawDueDateValue ? formatLocalDateYyyyMmDd(rawDueDateValue / 1000) : (sheetDueRaw || 'N/A');
		assignedEditorDsp = (
			<HoverTooltip
				content={
					<>
						<div>Assigned to: {freedCampItem?.assigned_to_fullname}</div>
						<div>Due: <span style={{ color: isDueSoon ? 'red' : undefined, fontWeight: isDueSoon ? 'bold' : 'normal' }}>{dueDisplay}</span></div>
					</>
				}
			>
				<span style={{ fontWeight: 'normal' }}>
					<span>{freedCampItem?.assigned_to_fullname}</span>
				</span>
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
