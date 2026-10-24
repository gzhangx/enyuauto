import { useEffect, useState, useRef, useCallback } from 'react'
import '../App.css'
import { useAuth } from '../contexts/AuthContext';
import {
  getFreeCampAndUpdateOperations, getOpsAndMainList, getSheetOps, processOperation, deleteItemActionTask,
  anyKeyUpdater,
  getActionsToPerform,
  type FreeCampAndUpdateOperations,
  type ActionsToPerformInfo,
  loadMainData,
  combineOpsConfigWithFreedCampData,
  type DebugLog,
  formatLocalDateYyyyMmDd,
  completeDateUpdater,  
  updateDoneParentIds,
  updateDoneColumn,
  appendMainSheetRow,
  updateMainSheetColumns,
  projectCommentHeader,
  projectCommentText,
  loadMainSheetHeaders,
  type IOneDriveDirInfo,
  createMsExcelDataOps,
  getEditorWorkload,
  getActionFilePlan,
  buildFilledTaskTemplate,
  findFreedCampUserForEditor,
} from '../../shared/main_ops';
import type { ActionType } from '../../shared/types';
import { ErrorDialog } from './ErrorDialog';
import { AddProjectDialog, type NewProjectValues } from './AddProjectDialog';
import { freedCampOps } from '../lib/util';
import type { FreedCampLoginParams, LoginResponse } from '../../shared/freedcampTypes';
import { getTaskIdColumnName, type IOperationWithLineNumberAndParentTaskId } from '../../shared/opsTypes';
import { renderActionCell, renderSyncActionCell, type RenderActionCellDeps } from './projectsPageUtil/render_action_cell';
import { fetchOneDriveChildren } from './MicrosoftOneDrivePage';
import { listFoldersFiles, listSiblingFiles, type SiblingFileInfo } from '../lib/msGraphConfig';
import { CATEGORY_ALIASES } from '../lib/newProjectForm';


type LogMessage = {
  id: number;
  text: string;
};

type UseLoggerResult = {
  logMessages: LogMessage[];
  criticalError: { show: boolean; message: string };
  closeCriticalError: () => void;
} & DebugLog;

const toDateOnly = (value: Date): Date => {
  const date = new Date(value);
  date.setHours(0, 0, 0, 0);
  return date;
};

const formatDateLabel = (value: Date): string => {
  const year = value.getFullYear();
  const month = String(value.getMonth() + 1).padStart(2, '0');
  const day = String(value.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};

/** Group cell `语音制作 Human` marks someone who records with a human voice for that action. */
const isHumanForAction = (group: string | undefined, action: string): boolean =>
  (group || '').split('/').some((part) => part.trim() === `${action} Human`);

function dateSearchTerms(fileName: string): string[] {
  const match = fileName.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!match) return [];
  const [, year, month, day] = match;
  const monthPad = month.padStart(2, '0');
  const dayPad = day.padStart(2, '0');
  return [
    `${year}-${month}-${day}`,
    `${year}-${monthPad}-${dayPad}`,
    `${year}/${month}/${day}`,
    `${year}/${monthPad}/${dayPad}`,
  ];
}

function categorySearchTerms(project: IOperationWithLineNumberAndParentTaskId): string[] {
  const blob = `${project.文章类别 || ''} ${project.文件 || ''}`;
  const terms: string[] = [];
  for (const alias of CATEGORY_ALIASES) {
    if (blob.includes(alias.chinese) || blob.toLowerCase().includes(alias.english.toLowerCase())) {
      terms.push(alias.chinese, alias.english);
    }
  }
  return terms;
}

function projectMatchesSearch(project: IOperationWithLineNumberAndParentTaskId, query: string): boolean {
  const tokens = query.trim().toLowerCase().split(/\s+/).filter(Boolean);
  if (tokens.length === 0) return true;
  const haystack = [
    project.文件,
    project.作者,
    project.文章名,
    project.文章类别,
    projectCommentText(project),
    ...categorySearchTerms(project),
    ...dateSearchTerms(project.文件 || ''),
  ].join('\n').toLowerCase();
  return tokens.every((token) => haystack.includes(token));
}

/** `1 Text Editor & First Proofread` → `folder 1`. */
const folderShortLabel = (folderName?: string): string => {
  if (!folderName) return 'the folder';
  const match = folderName.match(/^(\d+)/);
  return match ? `folder ${match[1]}` : folderName;
};

const emptySourceFileHint = (action: ActionType, plan: ReturnType<typeof getActionFilePlan>): string => {
  if (action === '二校') return 'Create 校对 first, or put the first-proof file in folder 1';
  if (action === '校对') return 'Put the original manuscript in folder 0';
  if (action === '美编') return 'Create 二校 first, or put the second-proof file in folder 2';
  if (plan?.fileExtensions?.length) {
    return `No ${plan.fileExtensions.join(', ')} file in ${plan.sourceFolders?.join(' or ') || 'the folder'}.`;
  }
  const folders = [...(plan?.sourceFolders || []), ...(plan?.fallbackFolders || [])];
  return `No files in ${folders.join(' or ') || 'the folder'}.`;
};

const findLeaveOverlap = ({
  editor,
  dueDays,
  editorInfoAry,
  onLeaveData,
}: {
  editor: string;
  dueDays: number;
  editorInfoAry: Array<{ fullNamePN?: string; shortNameOnFreedCamp?: string }>;
  onLeaveData: Array<{ name: string; startDate: Date; endDate: Date }>;
}) => {
  const editorInfo = editorInfoAry.find((info) => info.shortNameOnFreedCamp === editor);
  const editorNames = [editorInfo?.fullNamePN, editorInfo?.shortNameOnFreedCamp, editor]
    .filter(Boolean)
    .map((value) => value?.toLowerCase());

  const startDate = toDateOnly(new Date());
  const endDate = toDateOnly(new Date());
  endDate.setDate(startDate.getDate() + dueDays);

  return onLeaveData.find((item) => {
    const leaveName = item.name?.toLowerCase();
    if (!leaveName || !editorNames.includes(leaveName)) {
      return false;
    }
    const leaveStart = toDateOnly(item.startDate);
    const leaveEnd = toDateOnly(item.endDate);
    return leaveStart <= endDate && leaveEnd >= startDate;
  });
};

const getOnLeaveWarning = ({
  selectedEditor,
  dueDays,
  editorInfoAry,
  onLeaveData,
}: {
  selectedEditor?: string;
  dueDays?: number;
  editorInfoAry: Array<{ fullNamePN?: string; shortNameOnFreedCamp?: string }>;
  onLeaveData: Array<{ name: string; startDate: Date; endDate: Date }>;
}): string | null => {
  if (!selectedEditor || !Number.isFinite(dueDays) || (dueDays as number) <= 0) {
    return null;
  }

  const editorInfo = editorInfoAry.find((info) => info.shortNameOnFreedCamp === selectedEditor);
  const overlap = findLeaveOverlap({ editor: selectedEditor, dueDays: dueDays as number, editorInfoAry, onLeaveData });
  if (!overlap) {
    return null;
  }

  return `Warning: ${editorInfo?.fullNamePN || selectedEditor} is on leave from ${formatDateLabel(overlap.startDate)} to ${formatDateLabel(overlap.endDate)} and that period overlaps the due-date window.`;
};

const useLogger = (displayDuration = 60000): UseLoggerResult => {
  const [logMessages, setLogMessages] = useState<LogMessage[]>([]);
  const [criticalError, setCriticalError] = useState<{ show: boolean; message: string }>({ show: false, message: '' });
  const logIdCounterRef = useRef(0);
  
  const doLog = useCallback<DebugLog['doLog']>((msg, critical = false) => {
    if (critical) {
      console.error(msg);
      setCriticalError({ show: true, message: msg });
    } else {
      console.log('show msg in doLog',msg);
    }
    const newLog: LogMessage = {
      id: logIdCounterRef.current++,
      text: msg
    };
    setLogMessages(prev => {
      const next = [...prev, newLog];
      return next.length > 50 ? next.slice(next.length - 50) : next;
    });

    setTimeout(() => {
      setLogMessages(prev => prev.filter(log => log.id !== newLog.id));
    }, displayDuration);
  }, [displayDuration]);

  const closeCriticalError = useCallback(() => {
    setCriticalError({ show: false, message: '' });
  }, []);

  return { logMessages, doLog, criticalError, closeCriticalError };
};


export const ProjectsPage = ({ onNavigateToFreedCamp }: { onNavigateToFreedCamp?: () => void }) => {
  const { token, msToken, msAccount, msLoginRedirect, ensureMsToken, msLogout, sheetInfoCache, opsConfig, setOpsConfig, combinedOpsAndData, setCombinedOpsAndData, freedCampCredentials, useMsOps, authLoadingStatus } = useAuth();  
  const [fullProjectList, setFullProjectList] = useState<IOperationWithLineNumberAndParentTaskId[]>([]);
  const [projectList, setProjectList] = useState<IOperationWithLineNumberAndParentTaskId[]>([]);
  const [responseData, setResponseData] = useState<string>('');
  const [isLoading, setIsLoading] = useState({
    listLoading: '',
    freeCampLoading: '',
    projectButtonAction: '',
  });
  const [progressText, setProgressText] = useState('');
  const [errorDialog, setErrorDialog] = useState<{ show: boolean; message: string }>({ show: false, message: '' });
  const [successDialog, setSuccessDialog] = useState<{ show: boolean; message: string }>({ show: false, message: '' });
  type CreateTasksDialogState = {
    operation: IOperationWithLineNumberAndParentTaskId;
    oneDriveFolders: IOneDriveDirInfo[];
    actions: ActionsToPerformInfo[];
    selected: Partial<Record<ActionType, boolean>>;
    useEnglishTemplate: Partial<Record<ActionType, boolean>>;
    useAITemplate: Partial<Record<ActionType, boolean>>;
    selectedEditor?: Partial<Record<ActionType, string>>;
    dueDays?: Partial<Record<ActionType, number>>;
    sourceFiles?: Partial<Record<ActionType, SourceFilePick>>;
    focusedAction?: ActionType;
    previewHtml?: Partial<Record<ActionType, string>>;
    previewStatus?: Partial<Record<ActionType, 'idle' | 'loading' | 'ready' | 'error'>>;
    previewDirty?: Partial<Record<ActionType, boolean>>;
  };
  type SourceFilePick = {
    status: 'loading' | 'ready' | 'error';
    sourceColumn: string;
    destFolder?: string;
    files: SiblingFileInfo[];
    currentName?: string;
    selectedItemId?: string;
    error?: string;
  };
  const [createTasksDialog, setCreateTasksDialog] = useState<CreateTasksDialogState | null>(null);
  const sheetOpsRef = useRef<Awaited<ReturnType<typeof getSheetOps>> | null>(null);
  const previewRef = useRef<HTMLDivElement | null>(null);
  const lastPreviewKeyRef = useRef('');
  const { logMessages, doLog, criticalError, closeCriticalError } = useLogger(60000);
  const [showLogPanel, setShowLogPanel] = useState(false);
  const [cachedToken, setCachedToken] = useState<{ token: LoginResponse; timestamp: number } | null>(null);
  type OwnerInfoEntry = { article: string; name: string; assignedOn: string; status: string; taskId?: string; timestamp: number; };
  const [showAllProjects, setShowAllProjects] = useState(false);
  const [projectSearch, setProjectSearch] = useState('');
  const [daysShowAlertAfterComplete, setDaysShowAlertAfterComplete] = useState(3);
  const [doneColumnSavingLine, setDoneColumnSavingLine] = useState<number | null>(null);
  const [ownerInfoDialog, setOwnerInfoDialog] = useState<{ title: string; entries: OwnerInfoEntry[] } | null>(null);
  const [showAddProject, setShowAddProject] = useState(false);
  const [editingProject, setEditingProject] = useState<IOperationWithLineNumberAndParentTaskId | null>(null);
  const [addingProject, setAddingProject] = useState(false);
  const [createConfirm, setCreateConfirm] = useState<{ action: ActionType; taskTitle: string; assignee: string } | null>(null);

  const originalFreedCampOps = getFreeCampAndUpdateOperations(freedCampOps);

  const getFilePlan = (action: ActionType) => {
    if (!combinedOpsAndData || !createTasksDialog) return undefined;
    return getActionFilePlan(
      action,
      createTasksDialog.operation,
      combinedOpsAndData.opsConfig,
      !!createTasksDialog.selected['教审' as ActionType],
    );
  };

  useEffect(() => {
    if (!createTasksDialog || !msToken || !combinedOpsAndData) return;
    const creatingDoctrineReview = !!createTasksDialog.selected['教审' as ActionType];
    const toLoad = createTasksDialog.actions
      .map((row) => row.action)
      .filter((action) => {
        if (!createTasksDialog.selected[action]) return false;
        const plan = getActionFilePlan(action, createTasksDialog.operation, combinedOpsAndData.opsConfig, creatingDoctrineReview);
        if (!plan) return false;
        const existing = createTasksDialog.sourceFiles?.[action];
        const sourceKey = [...(plan.sourceFolders || []), ...(plan.fallbackFolders || [])].join('|') || plan.sourceColumn || '';
        return !existing || existing.sourceColumn !== sourceKey || (existing.destFolder || '') !== (plan.destFolder || '');
      });
    if (toLoad.length === 0) return;
    const operation = createTasksDialog.operation;
    const setPick = (action: ActionType, pick: SourceFilePick) => {
      setCreateTasksDialog((prev) => prev && prev.operation === operation
        ? { ...prev, sourceFiles: { ...(prev.sourceFiles || {}), [action]: pick } }
        : prev);
    };
    for (const action of toLoad) {
      const plan = getActionFilePlan(action, operation, combinedOpsAndData.opsConfig, creatingDoctrineReview)!;
      const sourceColumn = [...(plan.sourceFolders || []), ...(plan.fallbackFolders || [])].join('|') || plan.sourceColumn || '';
      const destFolder = plan.destFolder;
      if (plan.sourceFolders?.length) {
        if (!operation.mainFolder?.trim()) {
          setPick(action, { status: 'error', sourceColumn, destFolder, files: [], error: 'mainFolder is empty on the sheet.' });
          continue;
        }
        setPick(action, { status: 'loading', sourceColumn, destFolder, files: [] });
        const mainFolder = operation.mainFolder.trim();
        listFoldersFiles(msToken, mainFolder, plan.sourceFolders)
          .then(async (files) => {
            const raw = files.length > 0 || !plan.fallbackFolders?.length
              ? files
              : await listFoldersFiles(msToken, mainFolder, plan.fallbackFolders);
            const listed = plan.fileExtensions?.length
              ? raw.filter((file) => plan.fileExtensions!.some((ext) => file.name.toLowerCase().endsWith(ext.toLowerCase())))
              : raw;
            const destFile = destFolder ? listed.find((file) => file.folderName === destFolder) : undefined;
            setPick(action, listed.length > 0
              ? { status: 'ready', sourceColumn, destFolder, files: listed, selectedItemId: (destFile ?? listed[0]).itemId }
              : { status: 'error', sourceColumn, destFolder, files: [], error: emptySourceFileHint(action, plan) });
          })
          .catch((error: unknown) => {
            setPick(action, { status: 'error', sourceColumn, destFolder, files: [], error: `Could not list the files: ${(error as Error)?.message || String(error)}` });
          });
        continue;
      }
      const link = ((operation as any)[sourceColumn] as string | undefined)?.trim();
      if (!link) {
        setPick(action, { status: 'error', sourceColumn, destFolder, files: [], error: `${sourceColumn} is empty on the sheet.` });
        continue;
      }
      setPick(action, { status: 'loading', sourceColumn, destFolder, files: [] });
      listSiblingFiles(msToken, link)
        .then(({ files, currentName }) => {
          const current = files.find((file) => file.name === currentName);
          setPick(action, {
            status: 'ready',
            sourceColumn,
            destFolder,
            files,
            currentName,
            selectedItemId: (current ?? files[0])?.itemId,
          });
        })
        .catch((error: unknown) => {
          setPick(action, { status: 'error', sourceColumn, destFolder, files: [], error: `Could not list the files: ${(error as Error)?.message || String(error)}` });
        });
    }
  }, [createTasksDialog, msToken, combinedOpsAndData]);

  useEffect(() => {
    if (!createTasksDialog || !combinedOpsAndData) return;
    const action = createTasksDialog.focusedAction;
    if (!action || !createTasksDialog.selected[action]) return;
    if (createTasksDialog.previewDirty?.[action]) return;
    const pick = createTasksDialog.sourceFiles?.[action];
    if (pick?.status === 'loading') return;
    const editor = createTasksDialog.selectedEditor?.[action] || '';
    const useAI = createTasksDialog.useAITemplate[action] !== false;
    const useEnglish = createTasksDialog.useEnglishTemplate[action] === true;
    const fileId = pick?.selectedItemId || '';
    const previewKey = `${action}|${editor}|${fileId}|${useAI}|${useEnglish}`;
    if (lastPreviewKeyRef.current === previewKey) return;
    lastPreviewKeyRef.current = previewKey;
    let cancelled = false;
    const operation = {
      ...createTasksDialog.operation,
      ...(editor ? { [action]: editor } : {}),
    } as IOperationWithLineNumberAndParentTaskId;
    const pickedFile = pick?.files.find((file) => file.itemId === pick.selectedItemId);
    setCreateTasksDialog((prev) => prev && prev.focusedAction === action
      ? { ...prev, previewStatus: { ...(prev.previewStatus || {}), [action]: 'loading' } }
      : prev);
    const log: DebugLog = { doLog: (msg) => console.log(msg) };
    buildFilledTaskTemplate({
      action,
      operation,
      combined: combinedOpsAndData,
      oneDriveDirInfos: createTasksDialog.oneDriveFolders,
      useAI: createTasksDialog.actions.find((row) => row.action === action)?.hasAITemplate ? useAI : false,
      useEnglish,
      pickedSource: pickedFile,
      siteLogin: { name: msAccount?.name || '', email: msAccount?.username || '' },
      log,
      msToken: msToken || undefined,
    }).then((html) => {
      if (cancelled) return;
      setCreateTasksDialog((prev) => {
        if (!prev || prev.focusedAction !== action || prev.previewDirty?.[action]) return prev;
        return {
          ...prev,
          previewHtml: { ...(prev.previewHtml || {}), [action]: `<!--${previewKey}-->${html}` },
          previewStatus: { ...(prev.previewStatus || {}), [action]: 'ready' },
        };
      });
    }).catch((error: unknown) => {
      if (cancelled) return;
      setCreateTasksDialog((prev) => prev && prev.focusedAction === action
        ? {
            ...prev,
            previewStatus: { ...(prev.previewStatus || {}), [action]: 'error' },
            previewHtml: { ...(prev.previewHtml || {}), [action]: `Could not build preview: ${(error as Error)?.message || String(error)}` },
          }
        : prev);
    });
    return () => { cancelled = true; };
  }, [
    combinedOpsAndData,
    createTasksDialog?.focusedAction,
    createTasksDialog?.selectedEditor,
    createTasksDialog?.sourceFiles,
    createTasksDialog?.useAITemplate,
    createTasksDialog?.useEnglishTemplate,
    createTasksDialog?.previewDirty,
    msToken,
    msAccount?.name,
    msAccount?.username,
  ]);

  useEffect(() => {
    const action = createTasksDialog?.focusedAction;
    if (!action || !previewRef.current) return;
    if (createTasksDialog?.previewDirty?.[action]) return;
    const html = createTasksDialog?.previewHtml?.[action] || '';
    const visible = html.replace(/^<!--.*?-->/, '');
    if (previewRef.current.innerHTML !== visible) {
      previewRef.current.innerHTML = visible;
    }
  }, [createTasksDialog?.focusedAction, createTasksDialog?.previewHtml, createTasksDialog?.previewDirty]);

  const renderSourceFilePicker = (action: ActionType) => {
    const pick = createTasksDialog?.sourceFiles?.[action];
    if (!pick) return null;
    const plan = getFilePlan(action);
    const destFolder = pick.destFolder || plan?.destFolder;
    const selected = pick.files.find((file) => file.itemId === pick.selectedItemId);
    const sourceFolder = selected?.folderName || plan?.sourceFolders?.[0];
    const alreadyInDest = !!(destFolder && selected?.folderName === destFolder);
    const fileSelect = pick.status === 'ready' ? (
      <select
        className="ctd-input ctd-source-select"
        value={pick.selectedItemId ?? ''}
        onChange={(e) => {
          const itemId = e.target.value;
          lastPreviewKeyRef.current = '';
          setCreateTasksDialog((prev) => {
            if (!prev?.sourceFiles?.[action]) return prev;
            return {
              ...prev,
              sourceFiles: { ...prev.sourceFiles, [action]: { ...prev.sourceFiles[action]!, selectedItemId: itemId } },
              previewDirty: { ...(prev.previewDirty || {}), [action]: false },
            };
          });
        }}
      >
        {pick.files.map((file) => (
          <option key={file.itemId} value={file.itemId}>
            {file.folderName ? `/${file.folderName}/${file.name}` : file.name}
          </option>
        ))}
      </select>
    ) : null;
    const openLink = selected ? <a href={selected.webUrl} target="_blank" rel="noreferrer">Open File</a> : null;
    const destFolderName = destFolder ? (
      <strong className="ctd-dest-folder">/{destFolder}</strong>
    ) : null;

    if (pick.status === 'loading') {
      return (
        <div className="ctd-source">
          <span className="ctd-source-label">
            {destFolder ? <>Loading files to copy into folder <strong className="ctd-dest-folder">/{destFolder}</strong>…</> : 'Loading files…'}
          </span>
        </div>
      );
    }
    if (pick.status === 'error') {
      return (
        <div className="ctd-source">
          <span className="ctd-warning">
            {pick.error}
            {plan?.linkPlaceholder
              ? ` {${plan.linkPlaceholder}} will be left empty in the task description.`
              : ''}
          </span>
        </div>
      );
    }
    if (plan?.linkPlaceholder === 'voice_file') {
      return (
        <div className="ctd-source">
          <span className="ctd-source-label">Select 语音 file</span>
          {fileSelect}
          {openLink}
        </div>
      );
    }
    if (destFolder && selected && !alreadyInDest) {
      return (
        <div className="ctd-source">
          <span className="ctd-source-label">Will copy</span>
          {fileSelect}
          <span className="ctd-source-label">into folder {destFolderName}.</span>
          {openLink}
        </div>
      );
    }
    if (destFolder && selected && alreadyInDest) {
      return (
        <div className="ctd-source">
          <span className="ctd-source-label">Using</span>
          {fileSelect}
          <span className="ctd-source-label">already in folder {destFolderName}.</span>
          {openLink}
        </div>
      );
    }
    return (
      <div className="ctd-source">
        <span className="ctd-source-label">File from {folderShortLabel(sourceFolder)}</span>
        {fileSelect}
        {openLink}
      </div>
    );
  };

  const getEligibleEditors = (row: { action: ActionType; hasAITemplate?: boolean }) => {
    if (!combinedOpsAndData) return [];
    const aiOn = row.hasAITemplate && createTasksDialog?.useAITemplate[row.action] !== false;
    return combinedOpsAndData.opsConfig.editorInfoAry.filter((info) =>
      info.group?.indexOf(row.action) >= 0 && !(aiOn && isHumanForAction(info.group, row.action)));
  };

  const renderWorkloadTable = (row: { action: ActionType; hasAITemplate?: boolean }) => {
    if (!combinedOpsAndData) return null;
    const rows = getEligibleEditors(row).map((info) => ({
      info,
      load: getEditorWorkload(info.shortNameOnFreedCamp, row.action, combinedOpsAndData),
    }));
    if (rows.length === 0) return null;
    rows.sort((a, b) => b.load.lastAssignedTs - a.load.lastAssignedTs);
    const selectedName = createTasksDialog?.selectedEditor?.[row.action];
    const dueDays = createTasksDialog?.dueDays?.[row.action] ?? 7;
    const onLeaveData = combinedOpsAndData.opsConfig.onLeaveData ?? opsConfig?.onLeaveData ?? [];
    return (
      <table className="ctd-workload">
        <thead>
          <tr>
            <th>Editor</th>
            <th>Open now</th>
            <th>Last assigned</th>
            <th>Last completed</th>
            <th>Done in total</th>
          </tr>
        </thead>
        <tbody>
          {rows.map(({ info, load }) => {
            const leave = findLeaveOverlap({
              editor: info.shortNameOnFreedCamp,
              dueDays,
              editorInfoAry: combinedOpsAndData.opsConfig.editorInfoAry,
              onLeaveData,
            });
            return (
            <tr
              key={info.shortNameOnFreedCamp}
              title={leave ? `On leave ${formatDateLabel(leave.startDate)} – ${formatDateLabel(leave.endDate)}` : undefined}
              className={[
                'ctd-row-pickable',
                info.shortNameOnFreedCamp === selectedName ? 'ctd-row-selected' : '',
                leave ? 'ctd-row-leave' : '',
              ].join(' ').trim()}
              onClick={() => {
                setCreateTasksDialog((prev) => {
                  if (!prev) return prev;
                  const selectedEditor = { ...(prev.selectedEditor || {}) };
                  selectedEditor[row.action] = info.shortNameOnFreedCamp;
                  return { ...prev, selectedEditor };
                });
              }}
            >
              <td>
                {info.fullNamePN || info.shortNameOnFreedCamp}
                {leave && <span className="ctd-leave-tag">on leave until {formatDateLabel(leave.endDate)}</span>}
              </td>
              {load.matchedUser ? (
                <>
                  <td>
                    <span className={`ctd-open-badge${leave ? '' : load.openCount > 0 ? ' ctd-open-busy' : ' ctd-open-free'}`}>
                      {load.openCount}
                    </span>
                  </td>
                  <td>{load.lastAssignedTs ? formatLocalDateYyyyMmDd(load.lastAssignedTs) : <span className="ctd-never">never</span>}</td>
                  <td>{load.lastCompletedTs ? formatLocalDateYyyyMmDd(load.lastCompletedTs) : '—'}</td>
                  <td>{load.completedCount}</td>
                </>
              ) : (
                <td className="ctd-muted" colSpan={4}>not found in FreedCamp</td>
              )}
            </tr>
            );
          })}
        </tbody>
      </table>
    );
  };

  const closeOwnerInfoDialog = () => setOwnerInfoDialog(null);
  const getOwnerInfoEntriesForAction = (action: ActionType): OwnerInfoEntry[] => {
    const byName = new Map<string, OwnerInfoEntry>();
    for (const p of fullProjectList) {
      const freedCampItem = p[`${action} FreeCamp Item`];
      if (!freedCampItem || !freedCampItem.assigned_to_fullname) continue;
      const timestamp = freedCampItem.completed_ts ?? 0;
      const entry: OwnerInfoEntry = {
        article: p.文章名,
        name: freedCampItem.assigned_to_fullname,
        assignedOn: timestamp ? formatLocalDateYyyyMmDd(timestamp) : 'Unknown',
        status: freedCampItem.status_title || 'Unknown',
        taskId: freedCampItem.id,
        timestamp,
      };
      const existing = byName.get(entry.name);
      if (!existing || entry.timestamp > existing.timestamp) {
        byName.set(entry.name, entry);
      }
    }
    return Array.from(byName.values()).sort((a, b) => b.timestamp - a.timestamp || a.name.localeCompare(b.name));
  };
  const openOwnerInfoDialogForAction = (action: ActionType) => {
    setOwnerInfoDialog({
      title: `Owner assignment for ${action}`,
      entries: getOwnerInfoEntriesForAction(action),
    });
  };

  const openCreateDialog = async (p: IOperationWithLineNumberAndParentTaskId, focusAction?: ActionType) => {
    if (opsConfig === null) {
      setResponseData('Operation data not loaded yet.');
      return;
    }
    const ops = sheetOpsRef.current;
    if (!ops || !combinedOpsAndData) return;
    setIsLoading((prev) => ({ ...prev, projectButtonAction: 'Preparing…' }));
    try {
      let oneDriveFolders: IOneDriveDirInfo[] = [];
      if (p.mainFolder) {
        const graphToken = await ensureMsToken();
        if (!graphToken) {
          setErrorDialog({
            show: true,
            message: 'Your Microsoft sign-in expired. Sign in with Microsoft again, then click the cell.',
          });
          return;
        }
        oneDriveFolders = await fetchOneDriveChildren(graphToken, p.mainFolder);
      }
      const previewLog: DebugLog = { doLog: (msg) => console.log(msg) };
      const previewOp = { ...p } as IOperationWithLineNumberAndParentTaskId;
      const actions = getActionsToPerform(previewOp, combinedOpsAndData, p['文件'], previewLog);
      if (actions.length === 0) {
        setErrorDialog({
          show: true,
          message: 'No new FreedCamp tasks to create for this row (all actions already have tasks or are excluded).',
        });
        return;
      }
      if (focusAction && !actions.some((row) => row.action === focusAction)) {
        setErrorDialog({
          show: true,
          message: `${focusAction} already has a task or is not available for this row.`,
        });
        return;
      }
      const selected: Partial<Record<ActionType, boolean>> = {};
      const useEnglishTemplate: Partial<Record<ActionType, boolean>> = {};
      const useAITemplate: Partial<Record<ActionType, boolean>> = {};
      const selectedEditor: Partial<Record<ActionType, string>> = {};
      const dueDays: Partial<Record<ActionType, number>> = {};
      for (const a of actions) {
        if (a.hasAITemplate) useAITemplate[a.action] = true;
        const maybeEditor = (p as any)[a.action] as string | undefined;
        if (maybeEditor) {
          selectedEditor[a.action] = maybeEditor;
        } else if (a.editorName && combinedOpsAndData?.opsConfig?.editorInfoAry) {
          const entries = combinedOpsAndData.opsConfig.editorInfoAry;
          for (const info of entries) {
            const prettyName = info.fullNamePN || info.shortNameOnFreedCamp || '';
            const normalizedTitle = (info.title || '').toLowerCase();
            const isEng = normalizedTitle === 'brother' || normalizedTitle === 'sister';
            const display = isEng ? `${info.title} ${prettyName}` : `${prettyName}${info.title || ''}`;
            if (display === a.editorName) {
              selectedEditor[a.action] = info.shortNameOnFreedCamp;
              break;
            }
          }
        }
        if (useAITemplate[a.action] && selectedEditor[a.action]) {
          const chosen = combinedOpsAndData.opsConfig.editorInfoAry.find((info) => info.shortNameOnFreedCamp === selectedEditor[a.action]);
          if (isHumanForAction(chosen?.group, a.action)) {
            delete selectedEditor[a.action];
          }
        }
        dueDays[a.action] = 7;
      }
      if (focusAction) selected[focusAction] = true;
      lastPreviewKeyRef.current = '';
      setCreateTasksDialog({
        operation: p,
        oneDriveFolders,
        actions,
        selected,
        useEnglishTemplate,
        useAITemplate,
        selectedEditor,
        dueDays,
        focusedAction: focusAction,
      });
    } catch (error: unknown) {
      const err = error as { message?: string };
      console.error('Error preparing project creation:', error);
      const detail = err.message || String(error);
      const tokenExpired = /token lifetime|expired|unauthenticated|invalid authentication/i.test(detail);
      setErrorDialog({
        show: true,
        message: tokenExpired
          ? 'Your Microsoft sign-in expired. Sign in with Microsoft again (top right), then click the cell.'
          : `Could not open the article folder:\n${detail}\n\nFolder:\n${p.mainFolder || '(empty)'}`,
      });
    } finally {
      setIsLoading((prev) => ({ ...prev, projectButtonAction: '' }));
    }
  };

  const submitCreateTask = async () => {
    const dlg = createTasksDialog;
    if (!dlg || !combinedOpsAndData) return;
    const action = dlg.focusedAction;
    if (!action) return;
    const editor = dlg.selectedEditor?.[action];
    if (!editor) return;
    (dlg.operation as any)[action] = editor;
    const days = dlg.dueDays?.[action];
    const normalizedDays = Number.isFinite(days) && (days as number) > 0 ? days as number : 7;
    const dueDate = new Date();
    dueDate.setDate(dueDate.getDate() + normalizedDays);
    dlg.operation[`${action} Due Date`] = `${dueDate.getFullYear()}-${String(dueDate.getMonth() + 1).padStart(2, '0')}-${String(dueDate.getDate()).padStart(2, '0')}`;
    const editedHtml = (previewRef.current?.innerHTML || dlg.previewHtml?.[action] || '').replace(/^<!--.*?-->/, '');
    setCreateTasksDialog(null);
    const ops = sheetOpsRef.current;
    if (!ops) return;
    setIsLoading((prev) => ({ ...prev, projectButtonAction: 'Creating project...' }));
    const logs: DebugLog = {
      doLog: (msg, critical = false) => {
        console.log('doLog', msg);
        setProgressText(msg);
        if (critical) {
          setErrorDialog({ show: true, message: msg });
        }
      },
    };
    try {
      const useAITemplateMap = new Map<ActionType, boolean>();
      if (dlg.useAITemplate[action] !== false) useAITemplateMap.set(action, true);
      const sourceFileOverride = new Map<ActionType, SiblingFileInfo>();
      const pick = dlg.sourceFiles?.[action];
      const file = pick?.files.find((item) => item.itemId === pick.selectedItemId);
      if (pick?.status === 'ready' && file) sourceFileOverride.set(action, file);
      const useEnglishTemplateMap = new Map<ActionType, boolean>();
      if (dlg.useEnglishTemplate[action]) useEnglishTemplateMap.set(action, true);
      const descriptionOverrides = new Map<ActionType, string>();
      if (editedHtml.trim()) descriptionOverrides.set(action, editedHtml);
      await processOperation(
        ops,
        freeCampOpsWithCache,
        combinedOpsAndData,
        dlg.operation,
        dlg.oneDriveFolders,
        logs,
        '',
        new Set<ActionType>([action]),
        useAITemplateMap,
        useEnglishTemplateMap,
        msToken || undefined,
        sourceFileOverride,
        { name: msAccount?.name || '', email: msAccount?.username || '' },
        descriptionOverrides,
        true,
      );
      if (freedCampCredentials) {
        setIsLoading((prev) => ({ ...prev, projectButtonAction: 'Reloading...' }));
        await fetchData(freedCampCredentials, true);
      }
    } catch (error: unknown) {
      const err = error as { message?: string };
      console.error('Error creating project:', error);
      setErrorDialog({
        show: true,
        message: `Failed to create project:\n${err.message || String(error)}`,
      });
    } finally {
      setIsLoading((prev) => ({ ...prev, projectButtonAction: '' }));
      setProgressText('');
    }
  };


  const startupRunOnce = useRef(false);
  const freeCampOpsWithCache: FreeCampAndUpdateOperations = {
    ...originalFreedCampOps,
    getFreedCampToken: async (opsAndTemplates, forceReload) => {
      const now = Date.now();
      const oneHour = 60 * 60 * 1000; // 1 hour in milliseconds
      
      // Check if cached token exists and is still valid
      if (!forceReload && cachedToken && (now - cachedToken.timestamp) < oneHour) {
        doLog('Using cached FreedCamp token');
        return cachedToken.token;
      }
      if (forceReload) {
        doLog('Force reload requested for FreedCamp token');
      }
      
      // Get new token if cache is expired or doesn't exist
      doLog('Fetching new FreedCamp token');
      const newToken = await originalFreedCampOps.getFreedCampToken(opsAndTemplates, forceReload);
      setCachedToken({ token: newToken, timestamp: now });
      return newToken;
    },    
  };

  const logParam = { getOperations: () => [], doLog }
  async function fetchData(freedCampCredentials: FreedCampLoginParams, fullReload = false) {
    console.log('useEffect run');
    const activeToken = useMsOps ? msToken : token;
    if (activeToken) {
      const ops = useMsOps
        ? createMsExcelDataOps(msToken!)
        : await getSheetOps({ token: token! }, sheetInfoCache);
      if (opsConfig &&combinedOpsAndData && !fullReload) {
        setIsLoading(prev => ({ ...prev, listLoading: 'Loading projects only...' }));
        const res = await loadMainData(ops);
        setOpsConfig({ ...opsConfig, headers: res.headers, operationList: res.operationList });
        setCombinedOpsAndData({
          ...combinedOpsAndData,
          opsConfig: { ...combinedOpsAndData.opsConfig, headers: res.headers, operationList: res.operationList },
        });
        updateDoneParentIds(combinedOpsAndData, res.operationList, logParam);
        prepareData(res);
        setIsLoading(prev=>({ ...prev, listLoading: '' }));
      } else {
        setIsLoading(prev => ({ ...prev, listLoading: 'Loading projects and config...' }));
      
        
        await getOpsAndMainList(ops, logParam).then(async res => {
          setOpsConfig(res); // Save the complete response (including functions)
          //const { ops, operationList, groupAndMainProjectMapping, editorInfoMap, headers } = res;
          const combined = await combineOpsConfigWithFreedCampData(res, freeCampOpsWithCache, freedCampCredentials, logParam, fullReload);          

          //--------------------------- get freecamp data        
          const pr = freeCampOpsWithCache.getFreedCampProcessor(combined.loginToken);
          for (const action of combined.opsConfig.groupAndMainProjectMapping.actions) {
            const actionInfo = combined.opsConfig.groupAndMainProjectMapping.shortProjectNameToProjectId[action];
            if (actionInfo) {
              setIsLoading(prev => ({ ...prev, freeCampLoading: `Loading FreeCamp tasks for ${action}...` }));
              const prjs = await pr.getTasksForProjects(actionInfo.project_id, 1);
              combined.freedCampTasksByAction[action] = prjs;            }
          }
          setCombinedOpsAndData(combined);
          updateDoneParentIds(combined, res.operationList, logParam);
          setIsLoading(prev => ({ ...prev, freeCampLoading: '' }));      
          //---------------------------
          prepareData(res);
        }).catch(error => {
          console.error('Error loading projects:', error);
          setResponseData(`Error: ${error.message || String(error)}`);
          doLog(`Error: ${error.message || String(error)}`);
          setErrorDialog({ show: true, message: `Failed to load projects:\n${error.message || String(error)}` });
        }).finally(() => {
          setIsLoading(prev => ({ ...prev, listLoading: '' }));
        });
      }
      sheetOpsRef.current = ops;
    }

    function prepareData(dataRes: { operationList: IOperationWithLineNumberAndParentTaskId[];}) {
      const list = dataRes.operationList.map((item, index) => ({ ...item, line: index + 2 }));
      setFullProjectList(list);
    }
  }

  useEffect(() => {
    const query = projectSearch.trim();
    const filtered = fullProjectList.filter((item) => {
      if (!query && !showAllProjects && item.isFinished) return false;
      return projectMatchesSearch(item, query);
    });
    setProjectList(filtered);
  }, [fullProjectList, showAllProjects, projectSearch]);

  //useEffect(() => {
    //if (!freedCampCredentials && onNavigateToFreedCamp) {
      //onNavigateToFreedCamp();
    //}
  //}, [freedCampCredentials]);

  useEffect(() => {
    const ready = useMsOps ? !!msToken : !!token;
    if (ready && !startupRunOnce.current && freedCampCredentials) {
      startupRunOnce.current = true;
      fetchData(freedCampCredentials);
    }
  }, [token, msToken, freedCampCredentials, useMsOps]);
  

  const syncProjectsCount = fullProjectList.filter(p => (p.syncFreeCampToSheetData?.updates?.length ?? 0) > 0).length;
  const hasAnySyncable = syncProjectsCount > 0;

  const handleSyncAll = async () => {
    if (!sheetOpsRef.current || !combinedOpsAndData) {
      setErrorDialog({ show: true, message: 'Sheet operations or combined config not ready' });
      return;
    }
    const items = fullProjectList.filter(p => (p.syncFreeCampToSheetData?.updates?.length ?? 0) > 0);
    if (items.length === 0) return;
    try {
      let i = 0;
      setIsLoading(prev => ({ ...prev, projectButtonAction: `Syncing 0/${items.length}` }));
      for (const p of items) {
        const updates = p.syncFreeCampToSheetData?.updates || [];
        let j = 0;
        for (const update of updates) {
          (p as any)[update.sheetCol as ActionType] = update.value;
          setIsLoading(prev => ({ ...prev, projectButtonAction: `Syncing ${p.文章名} 0/${items.length} of ${j}/${updates.length}` }));
          j++;
          await anyKeyUpdater(sheetOpsRef.current, combinedOpsAndData.opsConfig, update.sheetCol as ActionType, p, { doLog });
        }
        i++;
        setIsLoading(prev => ({ ...prev, projectButtonAction: `Syncing ${i}/${items.length}` }));
      }
      await fetchData(freedCampCredentials!);
    } catch (e: any) {
      console.error('Error syncing all:', e);
      setErrorDialog({ show: true, message: `Failed to sync all:\n${e.message || String(e)}` });
    } finally {
      setIsLoading(prev => ({ ...prev, projectButtonAction: '' }));
    }
  };



  // ...function removed, now imported from render_action_cell.tsx

  // Calculate animation speed based on number of messages
  if (authLoadingStatus || isLoading.listLoading || isLoading.freeCampLoading || isLoading.projectButtonAction) {
    return (
      <div style={{
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        minHeight: '100vh',
        gap: '20px'
      }}>
        <div style={{
          width: '60px',
          height: '60px',
          border: '6px solid rgba(33, 150, 243, 0.2)',
          borderTop: '6px solid rgb(33, 150, 243)',
          borderRadius: '50%',
          animation: 'spin 1s linear infinite'
        }}></div>
        {authLoadingStatus && <h2 style={{ margin: 0, color: '#555' }}>{authLoadingStatus}</h2>}
        <h2 style={{ margin: 0, color: '#555' }}>{isLoading.listLoading}</h2>
        <h2 style={{ margin: 0, color: '#555' }}>{isLoading.freeCampLoading}</h2>
        <h2 style={{ margin: 0, color: '#555' }}>{isLoading.projectButtonAction}</h2>
        {progressText && (
          <div style={{ 
            maxWidth: '600px', 
            padding: '10px 20px',
            color: '#777',
            fontSize: '13px',
            textAlign: 'center',
            lineHeight: '1.4'
          }}>
            {progressText}
          </div>
        )}
        <style>{`
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
        `}</style>

      </div>
    );
  }

  const actionCellDeps: RenderActionCellDeps = {
    sheetOpsRef,
    combined: combinedOpsAndData,
    freeCampOpsWithCache,
    freedCampLoginParams: freedCampCredentials!,
    deleteItemActionTask,
    completeDateUpdater,
    doLog,
    fetchData,
    setErrorDialog,
    onTaskUpdated: () => setFullProjectList((prev) => prev.slice()),
    daysShowAlertAfterComplete,
    msToken,
    wpToken: freedCampCredentials?.enyu_wp_token,
  };

  const stickyHeaderCellStyle = {
    padding: '12px 14px',
    textAlign: 'left' as const,
    position: 'sticky' as const,
    top: 0,
    zIndex: 900,
    backgroundColor: '#1f4b7a',
    color: '#fff',
    fontWeight: 600,
    border: '1px solid #8eafd4',
    whiteSpace: 'nowrap' as const,
  };

  if (!freedCampCredentials) {
    return (
      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: '40vh', gap: '1rem' }}>
        <p style={{ color: '#555' }}>FreedCamp credentials not loaded. Please set them first.</p>
        {onNavigateToFreedCamp && (
          <button
            onClick={onNavigateToFreedCamp}
            style={{ padding: '0.5rem 1.25rem', backgroundColor: '#e65100', color: 'white', border: 'none', borderRadius: '4px', fontSize: '0.95rem', cursor: 'pointer' }}
          >
            Go to FreedCamp Login
          </button>
        )}
      </div>
    );
  }

  const sheetValuesFromForm = (values: NewProjectValues, headers: string[]): Record<string, string> => {
    const savedValues = { ...(values as unknown as Record<string, string>) };
    const commentHeader = projectCommentHeader(headers);
    if (commentHeader) {
      savedValues[commentHeader] = values.comment || '';
    }
    savedValues.comment = values.comment || '';
    return savedValues;
  };

  const requireCommentColumn = (values: NewProjectValues, headers: string[]) => {
    if ((values.comment || '').trim() && !projectCommentHeader(headers)) {
      throw new Error('Add a column named 备注 on the main Excel sheet header row, save the workbook, then try Save again.');
    }
  };

  const liveMainHeaders = async (): Promise<string[]> => {
    const ops = sheetOpsRef.current;
    const previous = opsConfig?.headers ?? [];
    if (!ops) return previous;
    try {
      const headers = await loadMainSheetHeaders(ops);
      const named = headers.filter(Boolean);
      const previousNamed = previous.filter(Boolean);
      const looksComplete = named.includes('文件') && named.length >= Math.min(previousNamed.length, 8);
      if (looksComplete) {
        setOpsConfig((prev) => prev ? { ...prev, headers } : prev);
        return headers;
      }
      doLog(`Kept previous main headers; refreshed list looked incomplete (${named.length} named columns).`);
    } catch (error: unknown) {
      doLog(`Could not refresh main headers: ${(error as Error)?.message || String(error)}`);
    }
    return previous;
  };

  const handleAddProject = async (values: NewProjectValues) => {
    const ops = sheetOpsRef.current;
    if (!ops || !opsConfig) {
      setErrorDialog({ show: true, message: 'Sheet is not loaded yet.' });
      return;
    }
    const duplicate = fullProjectList.some(row => (row.文件 || '').trim() === values.文件);
    if (duplicate) {
      setErrorDialog({ show: true, message: `A row already uses 文件:\n${values.文件}` });
      return;
    }
    setAddingProject(true);
    try {
      const headers = await liveMainHeaders();
      requireCommentColumn(values, headers);
      const savedValues = sheetValuesFromForm(values, headers);
      const lineNumber = await appendMainSheetRow(ops, headers, fullProjectList, savedValues, { doLog });
      const newRow = { isFinished: false, noNeedToCreate: false, publishCompleted: false, itemPositionOnSheet: lineNumber } as IOperationWithLineNumberAndParentTaskId;
      for (const header of headers) {
        (newRow as unknown as Record<string, string>)[header] = savedValues[header] ?? '';
      }
      newRow.comment = savedValues.comment || '';
      newRow.备注 = savedValues.comment || '';
      if (combinedOpsAndData) {
        updateDoneParentIds(combinedOpsAndData, [newRow], { doLog });
      }
      setFullProjectList(prev => {
        const rest = prev.filter(row => row.itemPositionOnSheet !== lineNumber);
        return [...rest, newRow].sort((a, b) => a.itemPositionOnSheet - b.itemPositionOnSheet);
      });
      setShowAddProject(false);
      setSuccessDialog({ show: true, message: `${values.文件} 文件 saved successfully` });
    } catch (error: unknown) {
      const err = error as { message?: string };
      setErrorDialog({ show: true, message: `Failed to add project:\n${err.message || String(error)}` });
    } finally {
      setAddingProject(false);
    }
  };

  const handleEditProject = async (values: NewProjectValues) => {
    const ops = sheetOpsRef.current;
    const row = editingProject;
    if (!ops || !opsConfig || !row) {
      setErrorDialog({ show: true, message: 'Sheet is not loaded yet.' });
      return;
    }
    const duplicate = fullProjectList.some(item =>
      item.itemPositionOnSheet !== row.itemPositionOnSheet && (item.文件 || '').trim() === values.文件
    );
    if (duplicate) {
      setErrorDialog({ show: true, message: `A row already uses 文件:\n${values.文件}` });
      return;
    }
    setAddingProject(true);
    try {
      const headers = await liveMainHeaders();
      requireCommentColumn(values, headers);
      const savedValues = sheetValuesFromForm(values, headers);
      const writable: Record<string, string> = {};
      for (const header of ['文件', 'mainFolder', '作者', '称呼', '作者电邮', '文章类别', '文章名', '文章链接', 'slug']) {
        if (headers.includes(header)) writable[header] = savedValues[header] ?? '';
      }
      const commentHeader = projectCommentHeader(headers);
      if (commentHeader) writable[commentHeader] = savedValues[commentHeader] ?? '';
      await updateMainSheetColumns(ops, headers, row, writable, { doLog });
      setFullProjectList(prev => prev.map(item => {
        if (item.itemPositionOnSheet !== row.itemPositionOnSheet) return item;
        return {
          ...item,
          ...writable,
          comment: savedValues.comment || '',
          备注: savedValues.comment || '',
        };
      }));
      setEditingProject(null);
      setSuccessDialog({ show: true, message: `${values.文件} 文件 saved successfully` });
    } catch (error: unknown) {
      const err = error as { message?: string };
      setErrorDialog({ show: true, message: `Failed to update project:\n${err.message || String(error)}` });
    } finally {
      setAddingProject(false);
    }
  };

  return (
    <>
      <ErrorDialog 
        show={errorDialog.show}
        message={errorDialog.message}
        onClose={() => setErrorDialog({ show: false, message: '' })}
      />
      <ErrorDialog
        show={successDialog.show}
        message={successDialog.message}
        variant="success"
        onClose={() => setSuccessDialog({ show: false, message: '' })}
      />

      <ErrorDialog 
        show={criticalError.show}
        message={criticalError.message}
        onClose={closeCriticalError}
      />      

      {createTasksDialog && combinedOpsAndData && (
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
          }}
          onClick={() => setCreateTasksDialog(null)}
        >
          <div className="ctd-dialog ctd-dialog-wide" onClick={(e) => e.stopPropagation()}>
            <div className="ctd-header">
              <h2>Create FreedCamp Task</h2>
              <div className="ctd-project">{createTasksDialog.operation['文件']}</div>
            </div>
            <div className="ctd-body">
              <div className="ctd-action-cells">
                {createTasksDialog.actions.map((row) => (
                  <button
                    key={row.action}
                    type="button"
                    className={`ctd-action-cell${createTasksDialog.focusedAction === row.action ? ' ctd-action-cell-on' : ''}`}
                    onClick={() => {
                      lastPreviewKeyRef.current = '';
                      setCreateTasksDialog((prev) => {
                        if (!prev) return prev;
                        const selected: Partial<Record<ActionType, boolean>> = {};
                        selected[row.action] = true;
                        return {
                          ...prev,
                          focusedAction: row.action,
                          selected,
                          previewDirty: { ...(prev.previewDirty || {}), [row.action]: false },
                        };
                      });
                    }}
                  >
                    <span className="ctd-action">{row.action}</span>
                  </button>
                ))}
              </div>
              {(() => {
                const row = createTasksDialog.actions.find((item) => item.action === createTasksDialog.focusedAction);
                if (!row) {
                  return <div className="ctd-muted ctd-pick-hint">Click an action to choose the editor and preview the FreedCamp task.</div>;
                }
                return (
                  <div className="ctd-sections">
                    <section className="ctd-section">
                      <h4>Editor</h4>
                      <div className="ctd-controls">
                        <select
                          className="ctd-input"
                          value={createTasksDialog.selectedEditor?.[row.action] ?? ''}
                          onChange={(e) => {
                            const val = e.target.value;
                            lastPreviewKeyRef.current = '';
                            setCreateTasksDialog((prev) => {
                              if (!prev) return prev;
                              const selectedEditor = { ...(prev.selectedEditor || {}) };
                              selectedEditor[row.action] = val;
                              return { ...prev, selectedEditor, previewDirty: { ...(prev.previewDirty || {}), [row.action]: false } };
                            });
                          }}
                        >
                          <option value="">(select editor)</option>
                          {getEligibleEditors(row).map((info) => {
                            const prettyName = info.fullNamePN || info.shortNameOnFreedCamp || '';
                            const normalizedTitle = (info.title || '').toLowerCase();
                            const isEng = normalizedTitle === 'brother' || normalizedTitle === 'sister';
                            const pretty = isEng ? `${info.title} ${prettyName}` : `${prettyName}${info.title || ''}`;
                            return <option key={info.shortNameOnFreedCamp} value={info.shortNameOnFreedCamp}>{info.fullNamePN || pretty}</option>;
                          })}
                        </select>
                        <label className="ctd-due">
                          Due in
                          <input
                            className="ctd-input"
                            type="number"
                            min="1"
                            step="1"
                            value={createTasksDialog.dueDays?.[row.action] ?? 7}
                            onChange={(e) => {
                              const rawValue = e.target.value;
                              const parsedValue = Number(rawValue);
                              const normalizedValue = rawValue === '' || !Number.isFinite(parsedValue) || parsedValue <= 0 ? 7 : parsedValue;
                              setCreateTasksDialog((prev) => {
                                if (!prev) return prev;
                                const dueDays = { ...(prev.dueDays || {}) };
                                dueDays[row.action] = normalizedValue;
                                return { ...prev, dueDays };
                              });
                            }}
                            style={{ width: '58px' }}
                          />
                          day(s)
                        </label>
                        {row.hasEnglishTemplate && (
                          <label className="ctd-option">
                            <input
                              type="checkbox"
                              checked={createTasksDialog.useEnglishTemplate[row.action] === true}
                              onChange={(e) => {
                                lastPreviewKeyRef.current = '';
                                setCreateTasksDialog((prev) => prev && {
                                  ...prev,
                                  useEnglishTemplate: { ...prev.useEnglishTemplate, [row.action]: e.target.checked },
                                  previewDirty: { ...(prev.previewDirty || {}), [row.action]: false },
                                });
                              }}
                            />
                            English template
                          </label>
                        )}
                        {row.hasAITemplate && (
                          <label className="ctd-option">
                            <input
                              type="checkbox"
                              checked={createTasksDialog.useAITemplate[row.action] !== false}
                              onChange={(e) => {
                                const checked = e.target.checked;
                                lastPreviewKeyRef.current = '';
                                setCreateTasksDialog((prev) => {
                                  if (!prev) return prev;
                                  const useAITemplate = { ...prev.useAITemplate, [row.action]: checked };
                                  const selectedEditor = { ...(prev.selectedEditor || {}) };
                                  const current = selectedEditor[row.action];
                                  const currentInfo = combinedOpsAndData.opsConfig.editorInfoAry.find((info) => info.shortNameOnFreedCamp === current);
                                  if (checked && current && isHumanForAction(currentInfo?.group, row.action)) {
                                    selectedEditor[row.action] = '';
                                  }
                                  return { ...prev, useAITemplate, selectedEditor, previewDirty: { ...(prev.previewDirty || {}), [row.action]: false } };
                                });
                              }}
                            />
                            AI template
                          </label>
                        )}
                      </div>
                      {(() => {
                        const warning = getOnLeaveWarning({
                          selectedEditor: createTasksDialog.selectedEditor?.[row.action],
                          dueDays: createTasksDialog.dueDays?.[row.action],
                          editorInfoAry: combinedOpsAndData?.opsConfig?.editorInfoAry ?? opsConfig?.editorInfoAry ?? [],
                          onLeaveData: combinedOpsAndData?.opsConfig?.onLeaveData ?? opsConfig?.onLeaveData ?? [],
                        });
                        return warning ? <div className="ctd-warning">{warning}</div> : null;
                      })()}
                      {renderWorkloadTable(row)}
                    </section>
                    {(createTasksDialog.sourceFiles?.[row.action] || getFilePlan(row.action)) && (
                      <section className="ctd-section">
                        <h4>File</h4>
                        {renderSourceFilePicker(row.action) || <div className="ctd-muted">Loading files…</div>}
                      </section>
                    )}
                    <section className="ctd-section">
                      <h4>FreedCamp preview</h4>
                      {createTasksDialog.previewStatus?.[row.action] === 'loading' && (
                        <div className="ctd-muted">Building preview…</div>
                      )}
                      <div
                        ref={previewRef}
                        className="ctd-preview"
                        contentEditable
                        suppressContentEditableWarning
                        onInput={() => {
                          const html = previewRef.current?.innerHTML || '';
                          setCreateTasksDialog((prev) => prev && {
                            ...prev,
                            previewHtml: { ...(prev.previewHtml || {}), [row.action]: html },
                            previewDirty: { ...(prev.previewDirty || {}), [row.action]: true },
                          });
                        }}
                      />
                    </section>
                  </div>
                );
              })()}
            </div>
            <div className="ctd-footer">
              <button
                type="button"
                className="ctd-cancel"
                onClick={() => { setCreateConfirm(null); setCreateTasksDialog(null); }}
              >
                Cancel
              </button>
              <button
                type="button"
                className="btn btn-create"
                onClick={() => {
                  const dlg = createTasksDialog;
                  if (!dlg || !combinedOpsAndData) return;
                  const action = dlg.focusedAction;
                  if (!action) {
                    setErrorDialog({
                      show: true,
                      message: 'Click an action first, then press Create Task.',
                    });
                    return;
                  }
                  const editor = dlg.selectedEditor?.[action];
                  if (!editor) {
                    setErrorDialog({
                      show: true,
                      message: 'Select an editor so the FreedCamp task can be assigned.',
                    });
                    return;
                  }
                  const editorInfo = combinedOpsAndData.opsConfig.editorInfoAry.find((info) => info.shortNameOnFreedCamp === editor);
                  const fcUser = findFreedCampUserForEditor(editor, combinedOpsAndData);
                  const assignee = editorInfo?.fullNamePN || fcUser?.full_name || editor;
                  const taskTitle = `${action} — ${dlg.operation['文件']}`;
                  setCreateConfirm({ action, taskTitle, assignee });
                }}
              >
                Create Task
              </button>
            </div>
          </div>
        </div>
      )}

      {createConfirm && (
        <div
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: 'rgba(0, 0, 0, 0.55)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 2300,
          }}
          onClick={() => setCreateConfirm(null)}
        >
          <div className="ctd-confirm" onClick={(e) => e.stopPropagation()}>
            <p>
              Are you sure to create this task <strong>{createConfirm.taskTitle}</strong>. assigned to <strong>{createConfirm.assignee}</strong>.
            </p>
            <div className="ctd-footer">
              <button type="button" className="ctd-cancel" onClick={() => setCreateConfirm(null)}>Cancel</button>
              <button
                type="button"
                className="btn btn-create"
                onClick={() => {
                  setCreateConfirm(null);
                  void submitCreateTask();
                }}
              >
                Create Task
              </button>
            </div>
          </div>
        </div>
      )}

      <button
        type="button"
        aria-label={showLogPanel ? 'Hide logs' : 'Show logs'}
        onClick={() => setShowLogPanel(prev => !prev)}
        style={{
          position: 'fixed',
          top: '12px',
          right: '12px',
          width: '36px',
          height: '36px',
          borderRadius: '18px',
          border: '1px solid #ddd',
          backgroundColor: '#ffffff',
          color: '#333',
          fontSize: '12px',
          fontWeight: 700,
          cursor: 'pointer',
          zIndex: 1200,
          boxShadow: '0 2px 6px rgba(0,0,0,0.15)'
        }}
      >
        {logMessages.length}
      </button>

      {showLogPanel && (
      <div style={{
        position: 'fixed',
        top: '56px',
        right: '12px',
        width: '340px',
        height: '70vh',
        backgroundColor: '#fff',
        border: '1px solid #ddd',
        borderRadius: '6px',
        overflowY: 'auto',
        zIndex: 1100,
        display: 'flex',
        flexDirection: 'column',
        gap: '6px',
        padding: '10px',
        boxShadow: '0 4px 12px rgba(0,0,0,0.2)'
      }}>
        {logMessages.length === 0 ? (
          <div style={{ color: '#777', fontSize: '12px' }}>No logs yet</div>
        ) : (
          [...logMessages].reverse().map((log) => (
            <div
              key={log.id}
              style={{
                backgroundColor: 'rgba(33, 150, 243, 0.9)',
                color: 'white',
                padding: '8px 12px',
                borderRadius: '4px',
                fontSize: '12px',
                wordWrap: 'break-word',
                boxShadow: '0 2px 4px rgba(0,0,0,0.2)'
              }}
            >
              {log.text}
            </div>
          ))
        )}
      </div>
      )}

      <style>{`
        @keyframes slideInFromBottom {
          from {
            transform: translateY(100%);
            opacity: 0;
          }
          to {
            transform: translateY(0);
            opacity: 1;
          }
        }
        .projects-heading {
          display: flex;
          flex-direction: column;
          align-items: center;
          gap: 14px;
          margin-bottom: 1.25rem;
        }
        .projects-heading h1 {
          margin: 0;
          font-size: 2.35rem;
          line-height: 1.2;
          color: #1f4b7a;
          letter-spacing: 0.01em;
        }
        .projects-toolbar {
          display: flex;
          align-items: center;
          justify-content: center;
          gap: 10px;
          flex-wrap: wrap;
        }
        .projects-toolbar label,
        .projects-toolbar input {
          color: #212529;
        }
        .projects-search-row {
          display: flex;
          align-items: center;
          justify-content: center;
          gap: 10px;
          width: min(720px, 100%);
        }
        .projects-search {
          flex: 1;
          min-width: 0;
          padding: 8px 12px;
          font-size: 14px;
          border: 1px solid #c5d4e6;
          border-radius: 8px;
          background: #fff;
          color: #212529;
        }
        .projects-search:focus {
          outline: none;
          border-color: #2f6aa8;
          box-shadow: 0 0 0 2px rgba(47, 106, 168, 0.2);
        }
        .projects-search-count {
          color: #5b6b7c;
          font-size: 13px;
          white-space: nowrap;
        }
        .projects-table {
          width: 100%;
          border-collapse: collapse;
          background: #fff;
          color: #212529;
          border-radius: 10px;
          box-shadow: 0 2px 10px rgba(31, 75, 122, 0.12);
        }
        .projects-table td {
          padding: 10px 14px;
          border: 1px solid #c5d0dc;
          text-align: left;
          vertical-align: middle;
        }
        .projects-table tbody tr:nth-child(odd) {
          background-color: #ffffff;
        }
        .projects-table tbody tr:nth-child(even) {
          background-color: #eef4fb;
        }
        .projects-table tbody tr:hover {
          background-color: #dce9f7;
        }
        .projects-table a {
          color: #1a56a0;
        }
        .projects-table .fc-task-link {
          color: inherit;
          background: none;
          border: none;
          padding: 0;
          font: inherit;
          text-decoration: none;
          cursor: pointer;
        }
        .projects-table .fc-task-link:hover {
          opacity: 0.85;
        }
        .projects-table .fc-unassigned {
          color: #e65100;
          margin-right: 6px;
        }
        .fc-task-overlay {
          position: fixed;
          inset: 0;
          background: rgba(0, 0, 0, 0.5);
          z-index: 2300;
          display: flex;
          align-items: center;
          justify-content: center;
          padding: 1.5rem;
        }
        .fc-task-dialog {
          background: #fff;
          color: #212529;
          border-radius: 12px;
          width: 92%;
          max-width: 720px;
          max-height: 86vh;
          overflow: hidden;
          display: flex;
          flex-direction: column;
          box-shadow: 0 12px 40px rgba(15, 40, 70, 0.35);
          text-align: left;
        }
        .fc-task-dialog-head {
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 12px;
          padding: 1rem 1.25rem 0.75rem;
          border-bottom: 1px solid #dde5ee;
        }
        .fc-task-dialog-head h2 {
          margin: 0;
          font-size: 1.15rem;
          color: #1f4b7a;
        }
        .fc-task-dialog-body {
          overflow: auto;
          flex: 1;
        }
        .fc-task-close {
          padding: 6px 12px;
          border-radius: 6px;
          border: 1px solid #c5d0dc;
          background: #fff;
          cursor: pointer;
        }
        .fc-task-close:disabled,
        .fc-task-save:disabled {
          opacity: 0.6;
          cursor: not-allowed;
        }
        .fc-task-meta {
          margin: 0;
          padding: 0.85rem 1.25rem;
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 10px 16px;
          font-size: 13px;
        }
        .fc-task-meta div {
          display: grid;
          gap: 2px;
        }
        .fc-task-meta .fc-task-wide {
          grid-column: 1 / -1;
        }
        .fc-task-meta dt {
          color: #6b7785;
          font-weight: 600;
        }
        .fc-task-meta dd {
          margin: 0;
          color: #212529;
        }
        .fc-task-input {
          width: 100%;
          box-sizing: border-box;
          padding: 5px 8px;
          border: 1px solid #c5d0dc;
          border-radius: 6px;
          font: inherit;
          color: inherit;
          background: #fff;
        }
        .fc-task-input:focus,
        .fc-task-desc:focus {
          outline: 2px solid #8eafd4;
          border-color: #2f6aa8;
        }
        .fc-task-desc-label {
          padding: 0 1.25rem;
          color: #6b7785;
          font-weight: 600;
          font-size: 13px;
        }
        .fc-task-desc {
          margin: 6px 1.25rem 1.15rem;
          padding: 10px 12px;
          border: 1px solid #dde5ee;
          border-radius: 8px;
          overflow: auto;
          min-height: 88px;
          max-height: 24vh;
          font-size: 13px;
          line-height: 1.5;
          text-align: left;
          cursor: text;
        }
        .fc-task-footer {
          display: flex;
          align-items: center;
          justify-content: flex-end;
          gap: 8px;
          padding: 0.75rem 1.25rem 1rem;
          border-top: 1px solid #dde5ee;
        }
        .fc-task-save {
          padding: 6px 14px;
          border-radius: 6px;
          border: none;
          background: #1f4b7a;
          color: #fff;
          cursor: pointer;
        }
        .fc-task-save-msg,
        .fc-task-save-err {
          margin-right: auto;
          font-size: 13px;
        }
        .fc-task-save-msg {
          color: #2e7d32;
        }
        .fc-task-save-err {
          color: #c62828;
        }
        .fc-task-desc * {
          text-align: left !important;
        }
        .fc-task-comments-status {
          padding: 0 1.25rem 1.15rem;
          color: #6b7785;
          font-size: 13px;
        }
        .fc-task-comments {
          margin: 6px 1.25rem 1.15rem;
          display: flex;
          flex-direction: column;
          gap: 10px;
          overflow: auto;
          max-height: 28vh;
        }
        .fc-task-comment {
          border: 1px solid #dde5ee;
          border-radius: 8px;
          padding: 8px 10px;
          background: #fafcfe;
        }
        .fc-task-comment-meta {
          display: flex;
          justify-content: space-between;
          gap: 8px;
          font-size: 12px;
          color: #6b7785;
          margin-bottom: 4px;
        }
        .fc-task-comment-body {
          font-size: 13px;
          line-height: 1.5;
          text-align: left;
        }
        .fc-task-comment-body * {
          text-align: left !important;
        }
        .projects-table .line-col {
          width: 4.5rem;
          white-space: nowrap;
          vertical-align: top;
        }
        .projects-table .file-col {
          min-width: 28rem;
          width: 28rem;
        }
        .projects-table .file-comment {
          margin-top: 4px;
          font-size: 0.82rem;
          color: #5b6b7c;
          white-space: pre-wrap;
          display: -webkit-box;
          -webkit-line-clamp: 2;
          -webkit-box-orient: vertical;
          overflow: hidden;
        }
        .projects-table .line-edit-btn {
          color: inherit;
          background: none;
          border: none;
          padding: 0;
          font: inherit;
          cursor: pointer;
          text-decoration: underline;
          text-underline-offset: 2px;
        }
        .projects-table .line-edit-btn:hover {
          color: #1a56a0;
        }
        .projects-table a.file-preview {
          display: block;
          margin-top: 4px;
          font-size: 0.85rem;
          color: #e67e00;
        }
        .projects-table a.file-preview:hover {
          color: #c65d00;
        }
        .projects-table a.file-live {
          display: block;
          margin-top: 4px;
          font-size: 0.85rem;
          color: #e67e00;
        }
        .projects-table a.file-live:hover {
          color: #c65d00;
        }
        .ctd-dialog {
          display: flex;
          flex-direction: column;
          background: #fff;
          color: #212529;
          border-radius: 12px;
          width: 92%;
          max-width: 760px;
          max-height: 86vh;
          overflow: hidden;
          box-shadow: 0 12px 40px rgba(15, 40, 70, 0.35);
        }
        .ctd-dialog-wide {
          max-width: 980px;
        }
        .ctd-action-cells {
          display: flex;
          flex-wrap: wrap;
          gap: 8px;
        }
        .ctd-action-cell {
          border: 1px solid #dde5ee;
          background: #fff;
          border-radius: 8px;
          padding: 8px 4px;
          cursor: pointer;
        }
        .ctd-action-cell-on {
          border-color: #2f6aa8;
          box-shadow: 0 2px 8px rgba(31, 75, 122, 0.12);
        }
        .ctd-action-cell-on .ctd-action {
          background: #1f4b7a;
          color: #fff;
        }
        .ctd-pick-hint {
          padding: 24px 8px;
          text-align: center;
        }
        .ctd-preview {
          min-height: 220px;
          max-height: 360px;
          overflow-y: auto;
          padding: 12px 14px;
          border: 1px solid #c5d0dc;
          border-radius: 8px;
          background: #fff;
          font-size: 14px;
          line-height: 1.5;
          color: #212529;
          text-align: left;
        }
        .ctd-preview, .ctd-preview * {
          text-align: left !important;
        }
        .ctd-preview:focus {
          outline: none;
          border-color: #2f6aa8;
          box-shadow: 0 0 0 2px rgba(47, 106, 168, 0.2);
        }
        .ctd-table-create-cell {
          cursor: pointer;
        }
        .ctd-table-create-cell:hover {
          background: #eef4fb;
        }
        .ctd-confirm {
          background: #fff;
          color: #212529;
          border-radius: 12px;
          width: 92%;
          max-width: 520px;
          padding: 1.25rem 1.4rem 0.75rem;
          box-shadow: 0 12px 40px rgba(15, 40, 70, 0.35);
        }
        .ctd-confirm p {
          margin: 0 0 8px;
          line-height: 1.55;
          font-size: 15px;
        }
        .ctd-header {
          background: linear-gradient(135deg, #1f4b7a, #2f6aa8);
          color: #fff;
          padding: 1.1rem 1.5rem;
          text-align: center;
        }
        .ctd-header h2 {
          margin: 0;
          font-size: 1.35rem;
          letter-spacing: 0.02em;
        }
        .ctd-project {
          margin-top: 6px;
          font-size: 0.95rem;
          opacity: 0.92;
        }
        .ctd-body {
          list-style: none;
          margin: 0;
          padding: 1rem 1.25rem;
          overflow-y: auto;
          flex: 1;
          background: #f4f7fb;
          display: flex;
          flex-direction: column;
          gap: 8px;
        }
        .ctd-card {
          background: #fff;
          border: 1px solid #dde5ee;
          border-radius: 8px;
          padding: 8px 12px;
          font-size: 14px;
          display: flex;
          flex-direction: column;
          gap: 8px;
          transition: border-color 0.15s, box-shadow 0.15s;
        }
        .ctd-card:hover {
          border-color: #9fbbe0;
        }
        .ctd-card-selected {
          border-color: #2f6aa8;
          border-left: 4px solid #2f6aa8;
          box-shadow: 0 2px 8px rgba(31, 75, 122, 0.12);
        }
        .ctd-card-head {
          display: flex;
          align-items: center;
          gap: 10px;
        }
        .ctd-card-head input[type="checkbox"] {
          width: 16px;
          height: 16px;
          accent-color: #2f6aa8;
          cursor: pointer;
        }
        .ctd-card-label {
          flex: 1;
          display: flex;
          align-items: center;
          gap: 10px;
          cursor: pointer;
        }
        .ctd-action {
          display: inline-block;
          min-width: 4.5em;
          padding: 2px 10px;
          border-radius: 999px;
          background: #e3edf8;
          color: #1f4b7a;
          font-weight: 700;
          text-align: center;
        }
        .ctd-card-selected .ctd-action {
          background: #1f4b7a;
          color: #fff;
        }
        .ctd-editor-name {
          color: #6b7785;
        }
        .ctd-sections {
          display: flex;
          flex-direction: column;
          gap: 8px;
          padding-left: 26px;
        }
        .ctd-section {
          background: #f7fafd;
          border: 1px solid #e3eaf3;
          border-radius: 8px;
          padding: 8px 10px;
        }
        .ctd-section h4 {
          margin: 0 0 8px;
          font-size: 12px;
          font-weight: 700;
          color: #1f4b7a;
          letter-spacing: 0.02em;
        }
        .ctd-controls {
          display: flex;
          align-items: center;
          gap: 12px;
          flex-wrap: wrap;
        }
        .ctd-input {
          padding: 5px 8px;
          font-size: 13px;
          border: 1px solid #c5d0dc;
          border-radius: 6px;
          background: #fff;
          color: #212529;
        }
        .ctd-input:focus {
          outline: none;
          border-color: #2f6aa8;
          box-shadow: 0 0 0 2px rgba(47, 106, 168, 0.2);
        }
        .ctd-due {
          display: flex;
          align-items: center;
          gap: 6px;
          font-size: 13px;
          color: #555;
        }
        .ctd-source {
          width: 100%;
          display: flex;
          flex-wrap: wrap;
          align-items: center;
          gap: 8px;
          font-size: 13px;
        }
        .ctd-source-label {
          color: #555;
        }
        .ctd-dest-folder {
          color: #1f4b7a;
          font-weight: 700;
        }
        .ctd-source-select {
          flex: 1;
          min-width: 0;
        }
        .ctd-source a {
          color: #e67e00;
          white-space: nowrap;
        }
        .ctd-warning {
          width: 100%;
          padding: 6px 10px;
          border-radius: 6px;
          background: #fff7e6;
          border: 1px solid #f5c26b;
          color: #8a4b00;
          font-size: 12px;
        }
        .ctd-option {
          display: flex;
          align-items: center;
          gap: 8px;
          font-size: 13px;
          color: #555;
        }
        .ctd-option input[type="checkbox"] {
          accent-color: #2f6aa8;
        }
        .ctd-option label {
          cursor: pointer;
        }
        .ctd-workload {
          width: 100%;
          border-collapse: separate;
          border-spacing: 0;
          font-size: 12px;
          color: #3a4552;
          border: 1px solid #dde5ee;
          border-radius: 6px;
          overflow: hidden;
        }
        .ctd-workload th {
          background: #eef4fb;
          color: #1f4b7a;
          font-weight: 600;
          text-align: left;
          padding: 5px 8px;
          border-bottom: 1px solid #dde5ee;
        }
        .ctd-workload td {
          padding: 4px 8px;
          border-bottom: 1px solid #eef1f5;
          text-align: left;
        }
        .ctd-workload tbody tr:last-child td {
          border-bottom: none;
        }
        .ctd-workload tbody tr:nth-child(even) {
          background: #fafcfe;
        }
        .ctd-workload tr.ctd-row-pickable {
          cursor: pointer;
        }
        .ctd-workload tbody tr.ctd-row-pickable:hover {
          background: #dce9f7;
        }
        .ctd-workload tr.ctd-row-selected,
        .ctd-workload tbody tr.ctd-row-selected:hover {
          background: #fff4e0;
          box-shadow: inset 3px 0 0 #e67e00;
          font-weight: 600;
        }
        .ctd-workload tr.ctd-row-leave {
          background: #fdecea;
          color: #c62828;
        }
        .ctd-leave-tag {
          display: inline-block;
          margin-left: 6px;
          padding: 0 6px;
          border-radius: 999px;
          background: #c62828;
          color: #fff;
          font-size: 11px;
          font-weight: 600;
        }
        .ctd-open-badge {
          display: inline-block;
          min-width: 1.6em;
          padding: 0 6px;
          border-radius: 999px;
          text-align: center;
          font-weight: 700;
        }
        .ctd-open-free {
          background: #e6f4ea;
          color: #2e7d32;
        }
        .ctd-open-busy {
          background: #fff1dc;
          color: #b45309;
        }
        .ctd-never {
          color: #2f6aa8;
          font-style: italic;
        }
        .ctd-muted {
          color: #999;
        }
        .ctd-footer {
          display: flex;
          justify-content: flex-end;
          gap: 10px;
          padding: 0.85rem 1.25rem;
          border-top: 1px solid #dde5ee;
          background: #fff;
        }
        .ctd-cancel {
          padding: 8px 18px;
          border-radius: 6px;
          border: 1px solid #c5d0dc;
          background: #fff;
          color: #374151;
          cursor: pointer;
          font-size: 14px;
        }
        .ctd-cancel:hover {
          background: #f1f4f8;
        }
      `}</style>

      {(showAddProject || editingProject) && (
        <AddProjectDialog
          key={editingProject ? `edit-${editingProject.itemPositionOnSheet}` : 'new'}
          heading={editingProject ? (editingProject.文件 || 'Edit Project') : 'Add New Project'}
          subheading={editingProject ? projectCommentText(editingProject) : undefined}
          initial={editingProject ? {
            文件: editingProject.文件 || '',
            mainFolder: editingProject.mainFolder || '',
            作者: editingProject.作者 || '',
            称呼: editingProject.称呼 || '',
            作者电邮: editingProject.作者电邮 || '',
            文章类别: editingProject.文章类别 || '',
            文章名: editingProject.文章名 || '',
            文章链接: editingProject.文章链接 || '',
            slug: editingProject.slug || '',
            comment: projectCommentText(editingProject),
          } : undefined}
          categories={opsConfig?.categories ?? combinedOpsAndData?.opsConfig.categories ?? []}
          titles={opsConfig?.titles ?? combinedOpsAndData?.opsConfig.titles ?? []}
          projects={fullProjectList}
          saving={addingProject}
          msToken={msToken || undefined}
          onClose={() => {
            if (addingProject) return;
            setShowAddProject(false);
            setEditingProject(null);
          }}
          onSave={editingProject ? handleEditProject : handleAddProject}
        />
      )}

      <div className="projects-heading">
        <h1>Enyu Project Management</h1>
        <div className="projects-search-row">
          <input
            className="projects-search"
            type="search"
            value={projectSearch}
            onChange={(e) => setProjectSearch(e.target.value)}
            placeholder="Search by author, article, date, category, or 备注"
            aria-label="Search by author, article, date, category, or 备注"
          />
          {projectSearch.trim() !== '' && (
            <span className="projects-search-count">{projectList.length} match{projectList.length === 1 ? '' : 'es'}</span>
          )}
        </div>
        <div className="projects-toolbar">
        <button className="btn btn-create" onClick={() => { setEditingProject(null); setShowAddProject(true); }} disabled={addingProject || !opsConfig}>Add New Project</button>
        <button className="btn btn-create" onClick={() => fetchData(freedCampCredentials,false)}>Reload</button>
        <button className="btn btn-create" onClick={() => fetchData(freedCampCredentials, true)}>Reload All</button>
        {hasAnySyncable && (
        
          <button
            className="btn btn-create"
            onClick={handleSyncAll}
            style={{ backgroundColor: '#0078d4', color: 'white' }}
          >
            Sync All ({syncProjectsCount})
          </button>        
      )}
        {msAccount && (
          <button className="btn btn-create" onClick={msLogout} style={{ backgroundColor: '#e3f2fd', color: '#1565c0', display:'none' }}>
            MS Logout ({msAccount.username})
          </button>
        )}
        {!msAccount && (
          <button className="btn btn-create" onClick={msLoginRedirect} style={{ backgroundColor: '#e3f2fd', color: '#1565c0' }}>
            MS Login
          </button>
        )}
        <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '14px' }}>
          <input
            type="checkbox"
            checked={showAllProjects}
            onChange={(e) => setShowAllProjects(e.target.checked)}
          />
          Show all
        </label>
        <input
          type="number"
          min={0}
          max={10}
          placeholder="days"
          value={daysShowAlertAfterComplete}
          onChange={(e) => {
            const v = parseInt(e.target.value, 10);
            if (!isNaN(v) && v >= 0 && v <= 10) setDaysShowAlertAfterComplete(v);
          }}
          style={{ width: '60px', fontSize: '14px', padding: '2px 6px', background: '#fff', border: '1px solid #c5d4e6', borderRadius: '4px' }}
        />
        </div>
      </div>
      <table className="projects-table">
        <thead>
          <tr>
            <th className="line-col" style={stickyHeaderCellStyle} title="Click the number to edit this project">Line</th>
            <th className="file-col" style={stickyHeaderCellStyle}>文件</th>
            {opsConfig?.groupAndMainProjectMapping.actions.map((action) => (
              <th key={action} style={stickyHeaderCellStyle}>
                <button
                  type="button"
                  onClick={() => openOwnerInfoDialogForAction(action)}
                  style={{
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: '0.35rem',
                    border: 'none',
                    background: 'none',
                    padding: 0,
                    margin: 0,
                    font: 'inherit',
                    color: 'inherit',
                    cursor: 'pointer',
                  }}
                  title={`Show all owners for ${action} sorted by date descending`}
                >
                  <span>{action}</span>
                  <span style={{ fontSize: '0.9em' }}>👤</span>
                </button>
              </th>
            ))}
            <th
              style={stickyHeaderCellStyle}
              title='Sets the sheet column "done" to Y so this row is hidden unless "Show all" is checked.'
            >
              Done
            </th>
          </tr>
        </thead>
        <tbody>
          {projectList.length === 0 && (
            <tr>
              <td colSpan={3 + (opsConfig?.groupAndMainProjectMapping.actions.length || 0)} style={{ textAlign: 'center', color: '#5b6b7c', padding: '1.25rem' }}>
                {projectSearch.trim()
                  ? `No projects match “${projectSearch.trim()}”.`
                  : 'No projects to show.'}
              </td>
            </tr>
          )}
          {
            projectList.map((p,idx) => {
              return <tr key={`${p.itemPositionOnSheet}-${idx}`}>
                <td className="line-col">
                  <button
                    type="button"
                    className="line-edit-btn"
                    title="Edit this project"
                    onClick={() => { setShowAddProject(false); setEditingProject(p); }}
                  >
                    {p.itemPositionOnSheet}
                  </button>
                  { renderSyncActionCell(p, freedCampCredentials, actionCellDeps) }
                </td>
                <td className="file-col">
                  <div>
                    {p.mainFolder
                      ? <a href={p.mainFolder} target="_blank" rel="noreferrer">{p.文件}</a>
                      : p.文件}
                  </div>
                  {projectCommentText(p) && (
                    <div className="file-comment" title={projectCommentText(p)}>{projectCommentText(p)}</div>
                  )}
                  {p.slug?.trim() && (p.done || '').toUpperCase() === 'Y' && (
                    <a
                      className="file-live"
                      href={`http://www.enyu.acccn.org/editing/${encodeURIComponent(p.slug.trim())}`}
                      target="_blank"
                      rel="noreferrer"
                    >
                      view on website
                    </a>
                  )}
                  {p.publishCompleted && p.slug?.trim() && (p.done || '').toUpperCase() !== 'Y' && (
                    <a
                      className="file-preview"
                      href={`http://www.enyu.acccn.org/editing/${encodeURIComponent(p.slug.trim())}`}
                      target="_blank"
                      rel="noreferrer"
                    >
                      preview
                    </a>
                  )}
                </td>
                {opsConfig?.groupAndMainProjectMapping.actions.map((action) => {
                  const taskId = String(p[getTaskIdColumnName(action)] || '').trim();
                  const canCreate = p.noNeedToCreate === false && !taskId;
                  return (
                    <td
                      key={action}
                      className={canCreate ? 'ctd-table-create-cell' : undefined}
                      title={canCreate ? `Click the Cell to Create ${action} Task` : undefined}
                      onClick={() => { if (canCreate) openCreateDialog(p, action); }}
                    >
                      {renderActionCell(p, action, actionCellDeps, canCreate)}
                    </td>
                  );
                })}
                <td style={{ textAlign: 'center' }}>
                  <input
                    type="checkbox"
                    title='Maps to sheet column "done": Y hides this row from the default list.'
                    checked={(p.done || '').toUpperCase() === 'Y'}
                    disabled={
                      !opsConfig ||
                      opsConfig.headers.indexOf('done') < 0 ||
                      doneColumnSavingLine === p.itemPositionOnSheet
                    }
                    onChange={async (e) => {
                      const checked = e.target.checked;
                      const ops = sheetOpsRef.current;
                      if (!ops || !opsConfig) return;
                      setDoneColumnSavingLine(p.itemPositionOnSheet);
                      try {
                        await updateDoneColumn(ops, opsConfig, p, checked, { doLog });
                        setFullProjectList((prev) =>
                          prev.map((row) =>
                            row.itemPositionOnSheet === p.itemPositionOnSheet
                              ? { ...row, done: checked ? 'Y' : '', isFinished: checked }
                              : row,
                          ),
                        );
                      } catch (error: unknown) {
                        const err = error as { message?: string };
                        setErrorDialog({
                          show: true,
                          message: err.message || String(error),
                        });
                      } finally {
                        setDoneColumnSavingLine(null);
                      }
                    }}
                  />
                </td>
              </tr>
            })
          }
        </tbody>
      </table>      
      {ownerInfoDialog && (
        <div
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            width: '100vw',
            height: '100vh',
            backgroundColor: 'rgba(0,0,0,0.35)',
            display: 'flex',
            justifyContent: 'center',
            alignItems: 'center',
            zIndex: 2000,
            padding: '1rem',
          }}
          onClick={closeOwnerInfoDialog}
        >
          <div
            style={{
              backgroundColor: 'white',
              borderRadius: '8px',
              boxShadow: '0 12px 30px rgba(0,0,0,0.25)',
              maxWidth: '600px',
              width: '100%',
              maxHeight: '90vh',
              overflow: 'auto',
              padding: '1rem 1.25rem',
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '1rem' }}>
              <h3 style={{ margin: 0, fontSize: '1.1rem' }}>{ownerInfoDialog.title}</h3>
              <button
                type="button"
                onClick={closeOwnerInfoDialog}
                style={{ border: 'none', background: '#6c757d', color: 'white', borderRadius: '4px', padding: '4px 10px', cursor: 'pointer' }}
              >
                Close
              </button>
            </div>
            <div style={{ marginTop: '1rem' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                <thead>
                  <tr>
                    <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid #ddd' }}>Article</th>
                    <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid #ddd' }}>Owner</th>
                    <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid #ddd' }}>Assigned</th>
                    <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid #ddd' }}>Status</th>
                    <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid #ddd' }}>Task ID</th>
                  </tr>
                </thead>
                <tbody>
                  {ownerInfoDialog.entries.length === 0 ? (
                    <tr>
                      <td colSpan={5} style={{ padding: '12px', borderBottom: '1px solid #eee', textAlign: 'center', color: '#555' }}>
                        No owner assignment details available.
                      </td>
                    </tr>
                  ) : (
                    ownerInfoDialog.entries.map((entry, index) => (
                      <tr key={`${entry.name}-${index}`}>
                        <td style={{ padding: '8px', borderBottom: '1px solid #eee' }}>{entry.article}</td>
                        <td style={{ padding: '8px', borderBottom: '1px solid #eee' }}>{entry.name}</td>
                        <td style={{ padding: '8px', borderBottom: '1px solid #eee' }}>{entry.assignedOn}</td>
                        <td style={{ padding: '8px', borderBottom: '1px solid #eee' }}>{entry.status}</td>
                        <td style={{ padding: '8px', borderBottom: '1px solid #eee' }}>{entry.taskId || '—'}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
      {responseData && (
        <div style={{ marginTop: '2rem', textAlign: 'left' }}>
          <h2>Response Data:</h2>
          <pre style={{ 
            backgroundColor: '#f5f5f5', 
            padding: '1rem', 
            borderRadius: '4px',
            overflow: 'auto',
            maxHeight: '400px'
          }}>
            {responseData}
          </pre>
        </div>
      )}
      <p className="read-the-docs">
        
      </p>
    </>
  )
}
