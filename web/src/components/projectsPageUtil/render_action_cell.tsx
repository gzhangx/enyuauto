
import type { DebugLog } from '../../../shared/main_ops';
import { getTaskIdColumnName, type IOperationWithLineNumber, type IOperationWithLineNumberAndParentTaskId, type IOpsConfig } from '../../../shared/opsTypes';
import type { ActionType } from '../../lib/api';
import React from 'react';
import * as gs from '@gzhangx/googleapi';

// Render a cell for syncing sheet with FreedCamp if data exists in FreedCamp but not in p
export function renderSyncActionCell(
	p: IOperationWithLineNumberAndParentTaskId,
    deps: RenderActionCellDeps,
    desc: string,
) {
	const {
		sheetOpsRef,
		opsConfig,
		//freeCampOpsWithCache,
		//completeDateUpdater,
		//doLog,
		fetchData,
		setErrorDialog,		
	} = deps;

    if (!opsConfig) {
        return <div>Waiting for Data</div>;
    }

    const ready = sheetOpsRef.current && opsConfig;
    let tooltipText = 'Exists in FreedCamp but not in sheet.';
    if (!ready) {
        tooltipText += ' Sheet operations not ready.!!!!!';
        return <div title={tooltipText}>Waiting for Data</div>;
    }
	const parts: string[] = [];	
    for (const action of opsConfig.groupAndMainProjectMapping.actions) {
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
			if (freedCampKey === 'assigned_to_id') {
				opsConfig.editorInfoMap
			}
			// Find the column index in the sheet
			const colIdx = opsConfig.headers.indexOf(sheetCol);
			if (colIdx === -1) {
				parts.push(`Column ${sheetCol} not found in sheet headers`);
				return;
			}
			// Sheet value (sheetCol is a key in p)
			const sheetVal = p[sheetCol as keyof typeof p];
			// FreedCamp value
			const fcVal = freedCampItem[freedCampKey];
			// Only consider if FreedCamp has a value and it's different from sheet
			if (fcVal !== undefined && fcVal !== null && String(fcVal) !== String(sheetVal)) {
				parts.push(
					`Action ${action}: SheetCol "${sheetCol}" (col ${colIdx + 1}, line ${p.itemPositionOnSheet}) will update to "${fcVal}" (was "${sheetVal}")`
				);
			}
		});
    }    

	// Build tooltip text with FreedCamp info
	
	tooltipText += '\n' + parts.join('\n');

	// Button to sync sheet with FreedCamp
	return (
		<button
			className="btn btn-create"
			style={{ marginLeft: '5px', background: '#ff9800', color: 'white' }}
			title={tooltipText}
			onClick={async () => {
				if (sheetOpsRef.current && opsConfig) {
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
						await fetchData();
					} catch (error: any) {
						console.error('Error syncing sheet:', error);
						setErrorDialog({ show: true, message: `Failed to sync sheet:\n${error.message || String(error)}` });
					}
				}
			}}
		>
			{desc}
		</button>
	);
}


type RenderActionCellDeps = {
	sheetOpsRef: React.RefObject<any>;
	opsConfig: IOpsConfig | null;
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
		opsConfig,
		freeCampOpsWithCache,
		deleteItemActionTask,
		completeDateUpdater,
		doLog,
		fetchData,
		setErrorDialog,
	} = deps;

	const taskId = p[getTaskIdColumnName(action)];
	const freedCampItem = p[`${action} FreeCamp Item`];

	// Build tooltip text
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
	}

	if (!taskId) {
		return <span style={{ color: 'green', fontWeight: 'bold' }} title={tooltipText}>NA</span>;
	}
	if (taskId === 'TaskIsEnglishAndIsDisabledForEnglish') {
		return <span style={{ color: 'orange', fontWeight: 'bold' }} title={tooltipText}>No English</span>;
	}
	if (taskId === 'done') {
		return <span style={{ color: 'green', fontWeight: 'bold' }} title={tooltipText}>Done</span>;
	}

	const hasCompletedDate = freedCampItem?.completed_ts;

	return (
		<>
			<button
				className="btn btn-delete"
				title={tooltipText}
				onClick={async () => {
					if (sheetOpsRef.current && opsConfig) {
						await deleteItemActionTask(sheetOpsRef.current, freeCampOpsWithCache, opsConfig, p, action, { getOperations: () => [], doLog });
					}
				}}
			>
				Delete {p[getTaskIdColumnName(action)]}
			</button>
			{hasCompletedDate && (
				<button
					className="btn btn-create"
					style={{ marginLeft: '5px' }}
					title={`Mark as done with completion date: ${new Date(hasCompletedDate * 1000).toLocaleString()}`}
					onClick={async () => {
						if (sheetOpsRef.current && opsConfig) {
							try {
								const completedDate = new Date(hasCompletedDate * 1000);
								const formattedDate = completedDate.toLocaleDateString();
								await completeDateUpdater(
									sheetOpsRef.current,
									opsConfig,
									action,
									p,
									formattedDate,
									{ doLog }
								);
								doLog(`Marked ${action} as done for "${p.文件}" (completed: ${formattedDate})`);
								await fetchData();
							} catch (error: any) {
								console.error('Error updating sheet:', error);
								setErrorDialog({ show: true, message: `Failed to update sheet:\n${error.message || String(error)}` });
							}
						}
					}}
				>
					Mark Done
				</button>
			)}
		</>
	);
}
