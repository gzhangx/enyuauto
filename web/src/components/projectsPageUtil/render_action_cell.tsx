
import { anyKeyUpdater, type DebugLog, type ICombinedOpsAndFreeCampData } from '../../../shared/main_ops';
import { getTaskIdColumnName, type IOperationWithLineNumber, type IOperationWithLineNumberAndParentTaskId, type IOpsConfig } from '../../../shared/opsTypes';
import type { ActionType } from '../../lib/api';
import React from 'react';
import * as gs from '@gzhangx/googleapi';

function formatLocalDateYyyyMmDd(unixSeconds: number): string {
	const date = new Date(unixSeconds * 1000);
	return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

// Render a cell for syncing sheet with FreedCamp if data exists in FreedCamp but not in p
export function renderSyncActionCell(
	p: IOperationWithLineNumberAndParentTaskId,
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
	const parts: string[] = [];	
	const updateActions: (() => Promise<void>)[] = [];
    for (const action of combined.opsConfig.groupAndMainProjectMapping.actions) {
        const freedCampItem = p[`${action} FreeCamp Item`];
		if (!freedCampItem) {
            parts.push(`Action ${action}: No FreedCamp item, skipped`);
            continue;
        }
		// Map sheet columns to FreedCamp item keys
		const columnMapping: { [sheetCol: string]: keyof typeof freedCampItem } = {
			[action]: 'assigned_to_id',
			[`${action} Due Date`]: 'due_ts',
			[`${action} Complete Date`]: 'completed_ts',
        };
        
		// For each mapped column, if FreedCamp has data but sheet does not (or is different),
		// collect the update info: sheet column index, line number, FreedCamp value
		Object.entries(columnMapping).forEach(([sheetCol, freedCampKey]) => {
			// Sheet value (sheetCol is a key in p)
			const sheetVal = p[sheetCol as keyof typeof p];
			// FreedCamp value
			let fcVal = freedCampItem[freedCampKey] as string;
			let compareValues: string[] = [];			
			if (freedCampKey === 'assigned_to_id') {
				const userInfo = combined.userIdToInfoMap[fcVal];
				console.log(`Mapping user ID ${fcVal} to name:`, userInfo, freedCampItem, freedCampKey);
				if (userInfo) {
					fcVal = userInfo.full_name; // Replace ID with name for comparison and display
				}
				if (!fcVal || fcVal === '0') {
					return;
				}
			} else {
				//it is date
				if (fcVal) {
					const ts = Number(fcVal);
					if (!Number.isNaN(ts)) {
						const date = new Date(ts * 1000);
						const yyyyMmDd = formatLocalDateYyyyMmDd(ts);
						const mDyYyyy = `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
						compareValues = [yyyyMmDd, mDyYyyy];
						fcVal = yyyyMmDd;
					}
				}
			}
			// Find the column index in the sheet
			const colIdx = combined.opsConfig.headers.indexOf(sheetCol);
			if (colIdx === -1) {
				parts.push(`Column ${sheetCol} not found in sheet headers`);
				return;
			}
			
			const sheetValStr = String(sheetVal);
			
			
			// Only consider if FreedCamp has a value and it's different from sheet
			const matchesSheet = compareValues.length > 0
				? compareValues.includes(sheetValStr)
				: String(fcVal) === sheetValStr;
			if (fcVal !== undefined && fcVal !== null && !matchesSheet) {
				const displayVal = compareValues.length > 0 ? `${compareValues[0]} / ${compareValues[1]}` : String(fcVal);
				parts.push(
					`Action ${action}: SheetCol "${sheetCol}" (col ${colIdx + 1}, line ${p.itemPositionOnSheet}) will update to "${displayVal}" (was "${sheetVal}")`
				);
				updateActions.push(async () => { 
					p[sheetCol as ActionType] = fcVal; // Update local data immediately for better UX	
					await anyKeyUpdater(sheetOpsRef.current, combined.opsConfig, sheetCol as ActionType, p, { doLog });
				});
			}
		});
    }    

	// Build tooltip text with FreedCamp info
	
	tooltipText += '\n' + parts.join('\n');

	// Button to sync sheet with FreedCamp
	return (
		<button
			className="btn btn-create"
			style={{ marginLeft: '5px', background: updateActions.length === 0 ? '#9e9e9e' : '#ff9800', color: 'white' }}
			title={tooltipText}
			disabled={updateActions.length === 0}
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
						for (const updateAction of updateActions) {
							await updateAction();
						}
						await fetchData();
					} catch (error: any) {
						console.error('Error syncing sheet:', error);
						setErrorDialog({ show: true, message: `Failed to sync sheet:\n${error.message || String(error)}` });
					}
				}
			}}
		>
			{desc} { updateActions.length}
		</button>
	);
}


export type RenderActionCellDeps = {
	sheetOpsRef: React.RefObject<any>;
	//opsConfig: IOpsConfig | null;
	combined: ICombinedOpsAndFreeCampData | null;
	freeCampOpsWithCache: any;
	deleteItemActionTask: Function;
	completeDateUpdater: (ops: gs.gsAccount.IGetSheetOpsReturn, opsConfig: IOpsConfig, key: ActionType, item: IOperationWithLineNumber, val: string, log: DebugLog) => Promise<void>;
	doLog: (msg: string) => void;
	fetchData: () => Promise<void>;
	setErrorDialog: (v: { show: boolean; message: string }) => void;
	//logParam: { doLog: (msg: string) => void };
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
		//completeDateUpdater,
		doLog,
		//fetchData,
		//setErrorDialog,
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

	if (!taskId && (!hasCompletedDate && !hasAssignedTo)) {
		return <span style={{ color: 'orange' }} title={tooltipText}>NA</span>;
	}
	return (
		<>
			{ !hasAssignedTo && !hasCompletedDate && <button
				className="btn btn-delete"
				title={tooltipText}
				onClick={async () => {
					if (sheetOpsRef.current && combined) {
						await deleteItemActionTask(sheetOpsRef.current, freeCampOpsWithCache, combined.opsConfig, p, action, { getOperations: () => [], doLog });
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
				hasCompletedDate && (
					<span style={{ color: 'green', fontWeight: 'bold' }}>{formatLocalDateYyyyMmDd(hasCompletedDate)}</span>
				)
			// 	hasCompletedDate && (
			// 	<button
			// 		className="btn btn-create"
			// 		style={{ marginLeft: '5px' }}
			// 		title={`Mark as done with completion date: ${new Date(hasCompletedDate * 1000).toLocaleString()}`}
			// 		onClick={async () => {
			// 			if (sheetOpsRef.current && combined) {
			// 				try {
			// 					const completedDate = new Date(hasCompletedDate * 1000);
			// 					const formattedDate = completedDate.toLocaleDateString();
			// 					await completeDateUpdater(
			// 						sheetOpsRef.current,
			// 						combined.opsConfig,
			// 						action,
			// 						p,
			// 						formattedDate,
			// 						{ doLog }
			// 					);
			// 					doLog(`Marked ${action} as done for "${p.文件}" (completed: ${formattedDate})`);
			// 					await fetchData();
			// 				} catch (error: any) {
			// 					console.error('Error updating sheet:', error);
			// 					setErrorDialog({ show: true, message: `Failed to update sheet:\n${error.message || String(error)}` });
			// 				}
			// 			}
			// 		}}
			// 	>
			// 		Mark Done
			// 	</button>
			// )
				}
		</>
	);
}
