
import type { FreedCampLoginParams } from '../../../shared/freedcampTypes';
import { anyKeyUpdater, formatLocalDateYyyyMmDd, type DebugLog, type ICombinedOpsAndFreeCampData, deleteItemActionTask } from '../../../shared/main_ops';
import { getTaskIdColumnName, type IOperationWithLineNumber, type IOperationWithLineNumberAndParentTaskId, type IOpsConfig, type ISheetDataOps } from '../../../shared/opsTypes';
import type { ActionType } from '../../lib/api';
import React, { useState, type JSX } from 'react';
import { fetchOneDriveChildren } from '../MicrosoftOneDrivePage';
import { convertDocxToHtml, buildWpOutputPage } from '../../lib/docToWp';

/** Minimal DriveItem shape needed locally */
type MinDriveItem = { id: string; name: string; folder?: object; parentReference?: { driveId?: string } };

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

  const handlePublish = async () => {
    if (!p.mainFolder) { setError('No mainFolder for this project'); return; }
    setLoading(true);
    setError('');
    try {
      // 1. List mainFolder → find the article folder by name
      const articleChildren = await fetchOneDriveChildren(msToken, p.mainFolder) as MinDriveItem[];
      

      const publishFolder = articleChildren.find(item => item.folder && /4\s*publish/i.test(item.name));
		if (!publishFolder) throw new Error(`"4 Publish" folder not found in "${p['文件']}"`);
		
	  const driveId = publishFolder.parentReference?.driveId;
      if (!driveId) throw new Error('Could not determine driveId for publish folder');

      // 3. List "4 Publish" → find first .docx
      const publishChildren = await fetchDriveItemChildren(msToken, driveId, publishFolder.id);
      const docxItem = publishChildren.find(item => /\.docx?$/i.test(item.name));
      if (!docxItem) throw new Error('No .docx file found in "4 Publish" folder');

      // 4. Download
      const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${docxItem.id}/content`;
      const res = await fetch(url, { headers: { Authorization: `Bearer ${msToken}` } });
      if (!res.ok) {
        const body = await res.json().catch(() => ({}));
        throw new Error((body as any)?.error?.message || res.statusText);
      }
      const arrayBuffer = await res.arrayBuffer();

      // 5. Convert and open
      const converted = await convertDocxToHtml(arrayBuffer);
      const page = buildWpOutputPage(converted, wpToken);
      const blob = new Blob([page], { type: 'text/html;charset=utf-8' });
      window.open(URL.createObjectURL(blob), '_blank');
    } catch (err: any) {
      setError(err.message || String(err));
    } finally {
      setLoading(false);
    }
  };

  return (
    <span>
      <button
        className="btn btn-create"
        onClick={handlePublish}
        disabled={loading}
        style={{ background: '#e91e63', color: 'white', marginLeft: '4px' }}
      >
        {loading ? '…' : '发布 WP'}
      </button>
      {error && <span style={{ color: 'red', fontSize: '11px', display: 'block' }}>{error}</span>}
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
	if (p.文件.includes('宣教无国界')) {
		console.log('fond 2026-1-25 宣教无国界 陈默 三福培训分享- ', p.文件)
	}
	let tooltipText = '';
	if (freedCampItem) {
		const parts: string[] = [];
		if (freedCampItem.completed_ts) {
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

	const hasCompletedDate = freedCampItem?.completed_ts;
	if (hasCompletedDate) {
		console.log(freedCampItem, 'debugremove freedCampItem is in render_action_cell.tsx')
	}

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
				title={freedCampItem.assigned_to_fullname || undefined}
			>
				{formatLocalDateYyyyMmDd(hasCompletedDate)}
			</span>
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
				hasAssignedTo && !hasCompletedDate && (freedCampItem.assigned_to_fullname)
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
