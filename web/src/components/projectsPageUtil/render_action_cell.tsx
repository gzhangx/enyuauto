import { getTaskIdColumnName, type IOperationWithLineNumberAndParentTaskId, type IOpsConfig } from '../../../shared/opsTypes';
import type { ActionType } from '../../lib/api';
import React from 'react';

type RenderActionCellDeps = {
	sheetOpsRef: React.RefObject<any>;
	opsConfig: IOpsConfig | null;
	freeCampOpsWithCache: any;
	deleteItemActionTask: Function;
	completeDateUpdater: Function;
	doLog: (msg: string) => void;
	fetchData: () => Promise<void>;
	setErrorDialog: (v: { show: boolean; message: string }) => void;
	logParam: any;
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
		logParam,
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
									logParam
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
