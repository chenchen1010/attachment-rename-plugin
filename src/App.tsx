import './App.css';
import {
  bitable,
  FieldType,
  ImageQuality,
  IOpenAttachment,
  IOpenCellValue,
  ITable,
} from '@lark-base-open/js-sdk';
import {
  Button,
  Input,
  InputNumber,
  Modal,
  Progress,
  Radio,
  Select,
  Space,
  Tabs,
  Toast,
  Typography,
} from '@douyinfe/semi-ui';
import { IconEyeClosed, IconEyeOpened } from '@douyinfe/semi-icons';
import { DragEvent, useCallback, useEffect, useMemo, useRef, useState } from 'react';

const PREVIEW_LIMIT = 50;
const BATCH_SIZE = 50;
const TOKEN_SEQ = '{{序号}}';

type Scope = 'selected' | 'view' | 'all';
type Mode = 'replace' | 'append';
type AppendPosition = 'prepend' | 'append' | 'insert';

type PreviewItem = {
  id: string;
  label: string;
  oldName: string;
  newName: string;
};

type UndoSnapshot = {
  fieldId: string;
  records: { recordId: string; value: IOpenAttachment[] }[];
};

type RenameConfig = {
  mode: Mode;
  replaceTemplate: string;
  appendPosition: AppendPosition;
  insertIndex: number;
  appendFront: string;
  appendBack: string;
  seqStart: number;
  seqPad: number;
};

function splitFileName(name: string): { base: string; ext: string } {
  const lastDot = name.lastIndexOf('.');
  if (lastDot <= 0) {
    return { base: name, ext: '' };
  }
  return { base: name.slice(0, lastDot), ext: name.slice(lastDot) };
}

function formatSequence(num: number, pad: number): string {
  const raw = String(num);
  if (pad <= 0) {
    return raw;
  }
  return raw.padStart(pad, '0');
}

function renderTemplate(
  input: string,
  seqText: string,
  fieldValues: Record<string, string>
): string {
  let result = input.split(TOKEN_SEQ).join(seqText);
  // 匹配 {{字段名}} 格式的变量，支持任意字段名
  result = result.replace(/\{\{([^{}]+)\}\}/g, (match, fieldName) => {
    const trimmedName = fieldName.trim();
    if (trimmedName === '序号') {
      return seqText;
    }
    return fieldValues[trimmedName] ?? '';
  });
  return result;
}

function ensureUniqueName(base: string, ext: string, used: Set<string>): string {
  let name = `${base}${ext}`;
  if (!used.has(name)) {
    used.add(name);
    return name;
  }
  let index = 1;
  while (used.has(`${base}_${index}${ext}`)) {
    index += 1;
  }
  name = `${base}_${index}${ext}`;
  used.add(name);
  return name;
}

function clampIndex(value: number, max: number): number {
  if (Number.isNaN(value) || !Number.isFinite(value)) {
    return 0;
  }
  if (value < 0) {
    return 0;
  }
  if (value > max) {
    return max;
  }
  return value;
}

function formatCellValue(value: unknown): string {
  if (value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'string') {
    return value;
  }
  if (typeof value === 'number') {
    return String(value);
  }
  if (typeof value === 'boolean') {
    return value ? '是' : '否';
  }
  if (Array.isArray(value)) {
    return value
      .map((item) => {
        if (item === null || item === undefined) {
          return '';
        }
        if (typeof item === 'string' || typeof item === 'number') {
          return String(item);
        }
        if (typeof item === 'object') {
          const obj = item as Record<string, unknown>;
          // 人员、群组、关联记录等
          if (typeof obj.name === 'string') {
            return obj.name;
          }
          // 富文本
          if (typeof obj.text === 'string') {
            return obj.text;
          }
          // 选项字段
          if (typeof obj.id === 'string' && typeof obj.text === 'string') {
            return obj.text;
          }
        }
        return '';
      })
      .filter(Boolean)
      .join(',');
  }
  if (typeof value === 'object') {
    const obj = value as Record<string, unknown>;
    // 进度字段: {status: "completed", value: 14}
    if ('value' in obj && (typeof obj.value === 'number' || typeof obj.value === 'string')) {
      return String(obj.value);
    }
    // 人员、关联记录等
    if (typeof obj.name === 'string') {
      return obj.name;
    }
    // 富文本
    if (typeof obj.text === 'string') {
      return obj.text;
    }
    // 标题字段
    if (typeof obj.title === 'string') {
      return obj.title;
    }
    // 链接字段: {link: "url", text: "显示文本"}
    if (typeof obj.link === 'string') {
      return typeof obj.text === 'string' && obj.text ? obj.text : obj.link;
    }
    // 地理位置: {location: "地址", name: "地点名"} 或 {address: "地址"}
    if (typeof obj.location === 'string') {
      return obj.location;
    }
    if (typeof obj.address === 'string') {
      return obj.address;
    }
  }
  // 其他未知类型，返回空字符串而不是 JSON
  return '';
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }
  try {
    return JSON.stringify(error);
  } catch {
    return String(error);
  }
}

function getDiffParts(oldName: string, newName: string): { prefix: string; diff: string; suffix: string } {
  let start = 0;
  const maxStart = Math.min(oldName.length, newName.length);
  while (start < maxStart && oldName[start] === newName[start]) {
    start += 1;
  }
  let endOld = oldName.length - 1;
  let endNew = newName.length - 1;
  while (endOld >= start && endNew >= start && oldName[endOld] === newName[endNew]) {
    endOld -= 1;
    endNew -= 1;
  }
  return {
    prefix: newName.slice(0, start),
    diff: newName.slice(start, endNew + 1),
    suffix: newName.slice(endNew + 1),
  };
}

function renameAttachments(
  attachments: IOpenAttachment[],
  config: RenameConfig,
  fieldValues: Record<string, string>
): { updated: IOpenAttachment[]; changed: boolean } {
  const used = new Set<string>();
  const updated = attachments.map((att, index) => {
    const { base, ext } = splitFileName(att.name);
    const seq = formatSequence(config.seqStart + index, config.seqPad);
    let newBase = base;

    if (config.mode === 'replace') {
      const rendered = renderTemplate(config.replaceTemplate, seq, fieldValues);
      newBase = rendered.trim() ? rendered : base;
    } else {
      const front = renderTemplate(config.appendFront, seq, fieldValues);
      const back = renderTemplate(config.appendBack, seq, fieldValues);
      const insertText = `${front}${seq}${back}`;
      if (config.appendPosition === 'prepend') {
        newBase = `${insertText}${base}`;
      } else if (config.appendPosition === 'append') {
        newBase = `${base}${insertText}`;
      } else {
        const safeIndex = clampIndex(config.insertIndex, base.length);
        newBase = `${base.slice(0, safeIndex)}${insertText}${base.slice(safeIndex)}`;
      }
    }

    const uniqueName = ensureUniqueName(newBase, ext, used);
    return { ...att, name: uniqueName };
  });

  const changed = updated.some((att, index) => att.name !== attachments[index]?.name);
  return { updated, changed };
}

function isImageAttachment(att: IOpenAttachment): boolean {
  const mime = (att as { type?: string }).type;
  if (mime && mime.startsWith('image/')) {
    return true;
  }
  const lower = att.name?.toLowerCase() || '';
  return ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.svg', '.heic', '.heif'].some((ext) => lower.endsWith(ext));
}

function isVideoAttachment(att: IOpenAttachment): boolean {
  const mime = (att as { type?: string }).type;
  if (mime && mime.startsWith('video/')) {
    return true;
  }
  const lower = att.name?.toLowerCase() || '';
  return ['.mp4', '.mov', '.avi', '.mkv', '.webm', '.m4v'].some((ext) => lower.endsWith(ext));
}

function getFileTag(name: string): string {
  const { ext } = splitFileName(name);
  if (!ext) {
    return 'FILE';
  }
  return ext.replace('.', '').toUpperCase();
}

function normalizeUrlList(list: string[], length: number): string[] {
  const result = Array.from({ length }, (_, index) => list[index] || '');
  return result;
}

function reorderList<T>(list: T[], from: number, to: number): T[] {
  const next = [...list];
  const [item] = next.splice(from, 1);
  next.splice(to, 0, item);
  return next;
}

export default function App() {
  const [table, setTable] = useState<ITable | null>(null);
  const [fieldOptions, setFieldOptions] = useState<{ label: string; value: string }[]>([]);
  const [variableOptions, setVariableOptions] = useState<{ label: string; value: string }[]>([]);
  const [fieldIdToName, setFieldIdToName] = useState<Record<string, string>>({});
  const [attachmentFieldId, setAttachmentFieldId] = useState('');

  const [scope, setScope] = useState<Scope>('view');
  const [manualSelectedRecordIds, setManualSelectedRecordIds] = useState<string[]>([]);
  const [mode, setMode] = useState<Mode>('replace');
  const [replaceTemplate, setReplaceTemplate] = useState('');
  const [appendPosition, setAppendPosition] = useState<AppendPosition>('append');
  const [insertIndex, setInsertIndex] = useState(0);
  const [appendFront, setAppendFront] = useState('');
  const [appendBack, setAppendBack] = useState('');
  const [seqStart, setSeqStart] = useState(1);
  const [seqPad, setSeqPad] = useState(0);
  const [activeInput, setActiveInput] = useState<'replace' | 'front' | 'back' | null>(null);
  const [cursorPosition, setCursorPosition] = useState<number | null>(null);

  const replaceInputRef = useRef<HTMLInputElement>(null);
  const frontInputRef = useRef<HTMLInputElement>(null);
  const backInputRef = useRef<HTMLInputElement>(null);

  const [estimatedCount, setEstimatedCount] = useState<number | null>(null);
  const [previewItems, setPreviewItems] = useState<PreviewItem[]>([]);
  const [previewLoading, setPreviewLoading] = useState(false);
  const [previewCollapsed, setPreviewCollapsed] = useState(true);

  const [selectedRecordId, setSelectedRecordId] = useState('');
  const [selectedRecordCount, setSelectedRecordCount] = useState(0);
  const [selectedAttachments, setSelectedAttachments] = useState<IOpenAttachment[]>([]);
  const [selectedAttachmentUrls, setSelectedAttachmentUrls] = useState<string[]>([]);
  const [selectedThumbnailUrls, setSelectedThumbnailUrls] = useState<string[]>([]);
  const [selectedPreviewIndex, setSelectedPreviewIndex] = useState(0);
  const [selectedLoading, setSelectedLoading] = useState(false);
  const [reorderSaving, setReorderSaving] = useState(false);
  const [dragIndex, setDragIndex] = useState<number | null>(null);
  const [dragOverIndex, setDragOverIndex] = useState<number | null>(null);

  const [processing, setProcessing] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0 });
  const [result, setResult] = useState<{ total: number; success: number; failed: number } | null>(null);
  const [undoStack, setUndoStack] = useState<UndoSnapshot[]>([]);

  const previewRequestRef = useRef(0);

  const pushUndoSnapshot = useCallback((snapshot: UndoSnapshot) => {
    setUndoStack((prev) => {
      const next = [snapshot, ...prev];
      return next.slice(0, 5);
    });
  }, []);

  const normalizedConfig = useMemo<RenameConfig>(() => {
    const safeSeqStart = Number.isFinite(seqStart) ? Math.max(0, Math.floor(seqStart)) : 1;
    const safeSeqPad = Number.isFinite(seqPad) ? Math.max(0, Math.floor(seqPad)) : 0;
    const safeInsert = Number.isFinite(insertIndex) ? Math.max(0, Math.floor(insertIndex)) : 0;
    return {
      mode,
      replaceTemplate,
      appendPosition,
      insertIndex: safeInsert,
      appendFront,
      appendBack,
      seqStart: safeSeqStart,
      seqPad: safeSeqPad,
    };
  }, [appendBack, appendFront, appendPosition, insertIndex, mode, replaceTemplate, seqPad, seqStart]);

  const refreshFields = useCallback(async () => {
    if (!table) {
      return;
    }
    try {
      const fields = await table.getFieldMetaList();
      const attachments = fields.filter((field) => field.type === FieldType.Attachment);
      const variables = fields.filter((field) => field.type !== FieldType.Attachment);

      setFieldOptions(attachments.map((field) => ({ label: field.name, value: field.id })));
      setVariableOptions(
        variables.map((field) => ({ label: field.name, value: field.id }))
      );

      // 建立字段 ID 到字段名的映射
      const idToName: Record<string, string> = {};
      for (const field of variables) {
        idToName[field.id] = field.name;
      }
      setFieldIdToName(idToName);

      const attachmentIds = new Set(attachments.map((field) => field.id));
      if (attachmentFieldId && !attachmentIds.has(attachmentFieldId)) {
        setAttachmentFieldId('');
      }
    } catch (error) {
      Toast.error(`字段加载失败：${getErrorMessage(error)}`);
    }
  }, [attachmentFieldId, table]);

  useEffect(() => {
    let mounted = true;
    const init = async () => {
      try {
        const activeTable = await bitable.base.getActiveTable();
        if (!mounted) {
          return;
        }
        setTable(activeTable);
      } catch (error) {
        Toast.error('无法获取表格，请刷新重试');
      }
    };
    init();
    return () => {
      mounted = false;
    };
  }, []);

  useEffect(() => {
    if (!table) {
      return;
    }
    refreshFields();
    const offAdd = table.onFieldAdd(refreshFields);
    const offModify = table.onFieldModify(refreshFields);
    const offDelete = table.onFieldDelete(refreshFields);
    return () => {
      offAdd?.();
      offModify?.();
      offDelete?.();
    };
  }, [refreshFields, table]);

  const getRecordIdsByScope = useCallback(async (targetScope: Scope): Promise<string[]> => {
    if (!table) {
      return [];
    }
    if (targetScope === 'all') {
      return await table.getRecordIdList();
    }
    const view = await table.getActiveView();
    if (targetScope === 'view') {
      const ids = await view.getVisibleRecordIdList();
      return ids.filter((id): id is string => Boolean(id));
    }
    // 使用 getSelection 获取选中的记录
    const selection = await bitable.base.getSelection();
    // 优先尝试获取多选记录
    const recordIds = (selection as { recordIds?: string[] })?.recordIds;
    if (Array.isArray(recordIds) && recordIds.length > 0) {
      return recordIds.filter((id): id is string => Boolean(id));
    }
    // 兜底：获取单条选中记录
    const recordId = (selection as { recordId?: string })?.recordId;
    return recordId ? [recordId] : [];
  }, [table]);

  const updateEstimatedCount = useCallback(async (targetScope: Scope) => {
    if (!table) {
      return;
    }
    try {
      const ids = await getRecordIdsByScope(targetScope);
      setEstimatedCount(ids.length);
    } catch (error) {
      setEstimatedCount(null);
      Toast.error('获取记录数量失败');
    }
  }, [getRecordIdsByScope, table]);

  useEffect(() => {
    if (!table) {
      return;
    }
    // 已选记录的数量由 handleScopeChange 设置，不需要这里更新
    if (scope !== 'selected') {
      updateEstimatedCount(scope);
    }
  }, [scope, table, updateEstimatedCount]);

  const buildPreview = useCallback(async () => {
    if (!table || !attachmentFieldId) {
      setPreviewItems([]);
      return;
    }
    const requestId = previewRequestRef.current + 1;
    previewRequestRef.current = requestId;
    setPreviewLoading(true);

    try {
      const recordIds = await getRecordIdsByScope(scope);
      const previewIds = recordIds.slice(0, PREVIEW_LIMIT);
      const items: PreviewItem[] = [];

      for (let i = 0; i < previewIds.length; i += 1) {
        const recordId = previewIds[i];
        const record = await table.getRecordById(recordId);
        const attachments = (record.fields?.[attachmentFieldId] as IOpenAttachment[] | null | undefined) || [];
        if (attachments.length === 0) {
          continue;
        }
        // 构建所有字段的值映射
        const fieldValues: Record<string, string> = {};
        for (const [fieldId, fieldName] of Object.entries(fieldIdToName)) {
          const rawValue = record.fields?.[fieldId];
          fieldValues[fieldName] = formatCellValue(rawValue);
        }
        const { updated } = renameAttachments(attachments, normalizedConfig, fieldValues);
        updated.forEach((att, index) => {
          items.push({
            id: `${recordId}-${index}`,
            label: `记录 ${i + 1} · 附件 ${index + 1}`,
            oldName: attachments[index]?.name || '',
            newName: att.name,
          });
        });
      }

      if (previewRequestRef.current === requestId) {
        setPreviewItems(items);
      }
    } catch (error) {
      if (previewRequestRef.current === requestId) {
        setPreviewItems([]);
      }
    } finally {
      if (previewRequestRef.current === requestId) {
        setPreviewLoading(false);
      }
    }
  }, [attachmentFieldId, fieldIdToName, getRecordIdsByScope, normalizedConfig, scope, table]);

  const loadSelectedPreview = useCallback(async () => {
    if (!table || !attachmentFieldId) {
      setSelectedRecordId('');
      setSelectedRecordCount(0);
      setSelectedAttachments([]);
      setSelectedAttachmentUrls([]);
      setSelectedThumbnailUrls([]);
      setSelectedPreviewIndex(0);
      return;
    }
    setSelectedLoading(true);
    try {
      const selectedIds = await getRecordIdsByScope('selected');
      setSelectedRecordCount(selectedIds.length);
      if (selectedIds.length !== 1) {
        setSelectedRecordId('');
        setSelectedAttachments([]);
        setSelectedAttachmentUrls([]);
        setSelectedThumbnailUrls([]);
        setSelectedPreviewIndex(0);
        return;
      }
      const recordId = selectedIds[0] || '';
      const record = await table.getRecordById(recordId);
      const attachments = (record.fields?.[attachmentFieldId] as IOpenAttachment[] | null | undefined) || [];
      setSelectedRecordId(recordId);
      setSelectedAttachments(attachments);
      const tokens = attachments.map((att) => att.token);
      if (tokens.length === 0) {
        setSelectedAttachmentUrls([]);
        setSelectedThumbnailUrls([]);
      } else {
        let thumbUrls: string[] = [];
        let fullUrls: string[] = [];
        try {
          thumbUrls = await table.getCellThumbnailUrls(tokens, attachmentFieldId, recordId, ImageQuality.Mid);
        } catch (error) {
          thumbUrls = [];
        }
        try {
          fullUrls = await table.getCellAttachmentUrls(tokens, attachmentFieldId, recordId);
        } catch (error) {
          fullUrls = [];
        }
        if (fullUrls.length === 0) {
          fullUrls = thumbUrls;
        }
        const normalizedFull = normalizeUrlList(fullUrls, tokens.length);
        const normalizedThumbs = normalizeUrlList(thumbUrls, tokens.length);
        const fallbackThumbs = normalizedThumbs.map((url, index) => {
          if (url) {
            return url;
          }
          if (isImageAttachment(attachments[index]) && normalizedFull[index]) {
            return normalizedFull[index];
          }
          return '';
        });
        setSelectedThumbnailUrls(fallbackThumbs);
        setSelectedAttachmentUrls(normalizedFull);
      }
      setSelectedPreviewIndex((prev) => {
        if (attachments.length === 0) {
          return 0;
        }
        return Math.min(prev, attachments.length - 1);
      });
    } catch (error) {
      setSelectedRecordId('');
      setSelectedAttachments([]);
      setSelectedAttachmentUrls([]);
      setSelectedThumbnailUrls([]);
      setSelectedPreviewIndex(0);
    } finally {
      setSelectedLoading(false);
    }
  }, [attachmentFieldId, getRecordIdsByScope, table]);

  useEffect(() => {
    if (!table) {
      return;
    }
    const offSelection = bitable.base.onSelectionChange(() => {
      loadSelectedPreview();
      if (scope === 'selected') {
        updateEstimatedCount(scope);
        buildPreview();
      }
    });
    const offRecordAdd = table.onRecordAdd(() => {
      if (scope === 'all') {
        updateEstimatedCount(scope);
        buildPreview();
      }
    });
    const offRecordDelete = table.onRecordDelete(() => {
      if (scope === 'all') {
        updateEstimatedCount(scope);
        buildPreview();
      }
    });
    return () => {
      offSelection?.();
      offRecordAdd?.();
      offRecordDelete?.();
    };
  }, [buildPreview, loadSelectedPreview, scope, table, updateEstimatedCount]);

  useEffect(() => {
    if (!table || !attachmentFieldId) {
      setPreviewItems([]);
      return;
    }
    const timer = window.setTimeout(() => {
      buildPreview();
    }, 300);
    return () => {
      window.clearTimeout(timer);
    };
  }, [attachmentFieldId, buildPreview, mode, table, scope, replaceTemplate, appendPosition, insertIndex, appendFront, appendBack, seqStart, seqPad]);

  useEffect(() => {
    loadSelectedPreview();
  }, [attachmentFieldId, loadSelectedPreview]);

  const handleInsertToken = useCallback((token: string) => {
    if (!activeInput) {
      Toast.info('先点击一个输入框，再插入变量');
      return;
    }

    // 获取当前输入框的引用和光标位置
    const getInputElement = () => {
      if (activeInput === 'replace') return replaceInputRef.current;
      if (activeInput === 'front') return frontInputRef.current;
      if (activeInput === 'back') return backInputRef.current;
      return null;
    };

    const inputEl = getInputElement();
    const pos = cursorPosition ?? (inputEl?.value?.length ?? 0);

    const insertAtPosition = (prev: string) => {
      const before = prev.slice(0, pos);
      const after = prev.slice(pos);
      return `${before}${token}${after}`;
    };

    if (activeInput === 'replace') {
      setReplaceTemplate(insertAtPosition);
      setCursorPosition(pos + token.length);
      return;
    }
    if (activeInput === 'front') {
      setAppendFront(insertAtPosition);
      setCursorPosition(pos + token.length);
      return;
    }
    if (activeInput === 'back') {
      setAppendBack(insertAtPosition);
      setCursorPosition(pos + token.length);
    }
  }, [activeInput, cursorPosition]);

  const handleDragStart = useCallback((index: number) => {
    setDragIndex(index);
  }, []);

  const handleDragOver = useCallback((event: DragEvent<HTMLDivElement>, index: number) => {
    event.preventDefault();
    setDragOverIndex(index);
  }, []);

  const handleDragEnd = useCallback(() => {
    setDragIndex(null);
    setDragOverIndex(null);
  }, []);

  const handleDrop = useCallback(async (index: number) => {
    if (dragIndex === null || dragIndex === index || reorderSaving) {
      setDragIndex(null);
      setDragOverIndex(null);
      return;
    }
    const previous = [...selectedAttachments];
    const prevUrls = [...selectedAttachmentUrls];
    const prevThumbs = [...selectedThumbnailUrls];
    const next = reorderList(previous, dragIndex, index);
    const nextUrls = reorderList(prevUrls, dragIndex, index);
    const nextThumbs = reorderList(prevThumbs, dragIndex, index);
    setSelectedAttachments(next);
    setSelectedAttachmentUrls(nextUrls);
    setSelectedThumbnailUrls(nextThumbs);
    setSelectedPreviewIndex((prev) => {
      if (prev === dragIndex) {
        return index;
      }
      if (dragIndex < prev && prev <= index) {
        return prev - 1;
      }
      if (dragIndex > prev && prev >= index) {
        return prev + 1;
      }
      return prev;
    });
    setDragIndex(null);
    setDragOverIndex(null);

    if (!table || !attachmentFieldId || !selectedRecordId) {
      return;
    }

    setReorderSaving(true);
    try {
      await table.setRecords([
        {
          recordId: selectedRecordId,
          fields: {
            [attachmentFieldId]: next,
          },
        },
      ]);
      pushUndoSnapshot({
        fieldId: attachmentFieldId,
        records: [{ recordId: selectedRecordId, value: previous }],
      });
      Toast.success('排序已更新');
    } catch (error) {
      setSelectedAttachments(previous);
      setSelectedAttachmentUrls(prevUrls);
      setSelectedThumbnailUrls(prevThumbs);
      Toast.error(`排序保存失败：${getErrorMessage(error)}`);
    } finally {
      setReorderSaving(false);
    }
  }, [attachmentFieldId, dragIndex, pushUndoSnapshot, reorderSaving, selectedAttachments, selectedAttachmentUrls, selectedRecordId, selectedThumbnailUrls, table]);

  const canStart = useMemo(() => {
    if (!attachmentFieldId || processing || reorderSaving) {
      return false;
    }
    if (mode === 'replace' && !replaceTemplate.trim()) {
      return false;
    }
    return true;
  }, [attachmentFieldId, mode, processing, reorderSaving, replaceTemplate]);

  const executeRename = useCallback(async (recordIds: string[]) => {
    if (!table) {
      Toast.error('无法获取表格');
      return;
    }
    if (!attachmentFieldId) {
      Toast.error('请先选择附件字段');
      return;
    }

    setProcessing(true);
    setProgress({ current: 0, total: recordIds.length });
    setResult(null);

    let success = 0;
    let failed = 0;
    const undoRecords: UndoSnapshot['records'] = [];

    try {
      for (let i = 0; i < recordIds.length; i += BATCH_SIZE) {
        const batchIds = recordIds.slice(i, i + BATCH_SIZE);
        const batchRecords = await Promise.all(
          batchIds.map(async (recordId) => {
            try {
              const record = await table.getRecordById(recordId);
              return { recordId, fields: record.fields as Record<string, IOpenCellValue> };
            } catch (error) {
              failed += 1;
              return null;
            }
          })
        );

        const validRecords = batchRecords.filter(Boolean) as {
          recordId: string;
          fields: Record<string, IOpenCellValue>;
        }[];

        const updates = validRecords
          .map((record) => {
            const attachments = (record.fields?.[attachmentFieldId] as IOpenAttachment[] | null | undefined) || [];
            if (attachments.length === 0) {
              return null;
            }
            // 构建所有字段的值映射
            const fieldValues: Record<string, string> = {};
            for (const [fieldId, fieldName] of Object.entries(fieldIdToName)) {
              const rawValue = record.fields?.[fieldId];
              fieldValues[fieldName] = formatCellValue(rawValue);
            }
            const { updated, changed } = renameAttachments(attachments, normalizedConfig, fieldValues);
            if (!changed) {
              return null;
            }
            undoRecords.push({ recordId: record.recordId, value: attachments });
            return {
              recordId: record.recordId,
              fields: {
                [attachmentFieldId]: updated,
              },
            };
          })
          .filter((item): item is { recordId: string; fields: Record<string, IOpenAttachment[]> } => Boolean(item));

        if (updates.length > 0) {
          try {
            await table.setRecords(updates);
            success += updates.length;
          } catch (error) {
            for (const update of updates) {
              try {
                await table.setRecords([update]);
                success += 1;
              } catch (innerError) {
                failed += 1;
              }
            }
          }
        }

        const processed = Math.min(i + batchIds.length, recordIds.length);
        setProgress({ current: processed, total: recordIds.length });
      }
    } finally {
      setProcessing(false);
      setResult({ total: recordIds.length, success, failed });
      if (undoRecords.length > 0) {
        pushUndoSnapshot({ fieldId: attachmentFieldId, records: undoRecords });
      }
      loadSelectedPreview();
      buildPreview();
    }
  }, [attachmentFieldId, buildPreview, fieldIdToName, loadSelectedPreview, normalizedConfig, pushUndoSnapshot, table]);

  const handleStart = useCallback(async () => {
    if (!table) {
      Toast.error('无法获取表格');
      return;
    }
    if (!attachmentFieldId) {
      Toast.error('请先选择附件字段');
      return;
    }
    if (mode === 'replace' && !replaceTemplate.trim()) {
      Toast.error('替换模式下，新名称不能为空');
      return;
    }

    let recordIds: string[] = [];

    if (scope === 'selected') {
      // 使用已选择的记录ID
      if (manualSelectedRecordIds.length === 0) {
        Toast.info('请先选择记录');
        return;
      }
      recordIds = manualSelectedRecordIds;
    } else {
      recordIds = await getRecordIdsByScope(scope);
      if (recordIds.length === 0) {
        Toast.info('当前范围没有记录');
        return;
      }
    }

    Modal.confirm({
      title: '确认重命名',
      content: `将处理 ${recordIds.length} 条记录的附件，是否继续？`,
      onOk: async () => {
        await executeRename(recordIds);
      },
    });
  }, [attachmentFieldId, executeRename, getRecordIdsByScope, mode, replaceTemplate, scope, table]);

  // 处理范围切换，选择"已选记录"时弹出选择对话框
  const handleScopeChange = useCallback(async (newScope: Scope) => {
    if (newScope === 'selected') {
      try {
        const selection = await bitable.base.getSelection();
        const tableId = selection?.tableId;
        const viewId = selection?.viewId;
        if (!tableId || !viewId) {
          Toast.error('无法获取当前表格信息');
          return;
        }
        const recordIds = await bitable.ui.selectRecordIdList(tableId, viewId);
        if (!recordIds || recordIds.length === 0) {
          Toast.info('未选择任何记录');
          return;
        }
        setManualSelectedRecordIds(recordIds);
        setScope('selected');
        setEstimatedCount(recordIds.length);
      } catch (error) {
        // 用户取消选择，不切换范围
      }
    } else {
      setManualSelectedRecordIds([]);
      setScope(newScope);
      updateEstimatedCount(newScope);
    }
  }, [updateEstimatedCount]);

  const handleUndo = useCallback(() => {
    if (!table || undoStack.length === 0) {
      return;
    }
    const snapshot = undoStack[0];
    Modal.confirm({
      title: '撤销上次操作',
      content: `将撤销 ${snapshot.records.length} 条记录的改动，是否继续？`,
      onOk: async () => {
        setProcessing(true);
        try {
          for (let i = 0; i < snapshot.records.length; i += BATCH_SIZE) {
            const batch = snapshot.records.slice(i, i + BATCH_SIZE).map((item) => ({
              recordId: item.recordId,
              fields: {
                [snapshot.fieldId]: item.value,
              },
            }));
            await table.setRecords(batch);
          }
          Toast.success('撤销完成');
          setUndoStack((prev) => prev.slice(1));
          loadSelectedPreview();
          buildPreview();
        } catch (error) {
          Toast.error(`撤销失败：${getErrorMessage(error)}`);
        } finally {
          setProcessing(false);
        }
      },
    });
  }, [buildPreview, loadSelectedPreview, table, undoStack]);

  const percent = progress.total > 0 ? Math.round((progress.current / progress.total) * 100) : 0;
  const { TabPane } = Tabs;
  const activeAttachment = selectedAttachments[selectedPreviewIndex];
  const activeAttachmentUrl = selectedAttachmentUrls[selectedPreviewIndex] || '';
  const activeThumbnailUrl = selectedThumbnailUrls[selectedPreviewIndex] || '';

  return (
    <main className="page">
      <header className="hero">
        <div className="hero-title">附件批量重命名</div>
        <div className="hero-subtitle">替换 / 追加 · 保留扩展名 · 自动序号</div>
        <div className="hero-tags">
          <span>作用范围可选</span>
          <span>同名自动加后缀</span>
          <span>最多撤销 5 次</span>
        </div>
      </header>

      <section className="card preview-top">
        <div className="card-header">
          <div className="section-title">选中单行即可进行附件预览</div>
          <button
            type="button"
            className="collapse-toggle"
            onClick={() => setPreviewCollapsed((prev) => !prev)}
            aria-label={previewCollapsed ? '展开预览' : '收起预览'}
          >
            {previewCollapsed ? <IconEyeClosed /> : <IconEyeOpened />}
            <span>{previewCollapsed ? '展开预览' : '收起预览'}</span>
          </button>
        </div>
        {previewCollapsed ? (
          <div className="collapse-hint">
            {selectedLoading && '加载中…'}
            {!selectedLoading && !attachmentFieldId && '请先选择附件字段'}
            {!selectedLoading && attachmentFieldId && selectedRecordCount === 0 && '请在表格中选中一行'}
            {!selectedLoading &&
              attachmentFieldId &&
              selectedRecordCount > 1 &&
              `已选 ${selectedRecordCount} 行，请只选一行进行预览`}
            {!selectedLoading &&
              attachmentFieldId &&
              selectedRecordCount === 1 &&
              '已选 1 行，点击眼睛展开预览'}
          </div>
        ) : (
          <div className="collapse-body">
            <div className="preview-top-body">
              <div className="preview-main">
                {selectedLoading && <div className="empty">加载中…</div>}
                {!selectedLoading && !attachmentFieldId && <div className="empty">请先选择附件字段</div>}
                {!selectedLoading && attachmentFieldId && selectedRecordCount === 0 && (
                  <div className="empty">请在表格中选中一行</div>
                )}
                {!selectedLoading && attachmentFieldId && selectedRecordCount > 1 && (
                  <div className="empty">已选 {selectedRecordCount} 行，请只选一行</div>
                )}
                {!selectedLoading &&
                  attachmentFieldId &&
                  selectedRecordCount === 1 &&
                  selectedAttachments.length === 0 && <div className="empty">该行没有附件</div>}
                {!selectedLoading &&
                  attachmentFieldId &&
                  selectedRecordCount === 1 &&
                  selectedAttachments.length > 0 &&
                  (() => {
                    if (!activeAttachment) {
                      return <div className="preview-placeholder">暂无预览</div>;
                    }
                    if (isVideoAttachment(activeAttachment) && activeAttachmentUrl) {
                      return (
                        <video className="preview-video" src={activeAttachmentUrl} controls>
                          当前浏览器不支持视频预览
                        </video>
                      );
                    }
                    if (isImageAttachment(activeAttachment) && (activeAttachmentUrl || activeThumbnailUrl)) {
                      const src = activeAttachmentUrl || activeThumbnailUrl;
                      return <img className="preview-image" src={src} alt={activeAttachment.name} />;
                    }
                    return <div className="preview-placeholder">{getFileTag(activeAttachment.name)}</div>;
                  })()}
              </div>
              <div className="preview-meta">
                <div className="meta-title">当前文件</div>
                <div className="file-name">{selectedRecordCount === 1 ? activeAttachment?.name || '—' : '—'}</div>
                <div className="meta-row">已选记录：{selectedRecordCount}</div>
                <div className="meta-row">附件数量：{selectedRecordCount === 1 ? selectedAttachments.length : 0}</div>
                <div className="meta-row">
                  {selectedRecordCount === 1 && selectedAttachments.length > 0
                    ? reorderSaving
                      ? '正在保存排序…'
                      : '拖拽缩略图可调整顺序'
                    : '请选择单行后查看预览'}
                </div>
              </div>
            </div>
            <div className="thumb-list">
              {selectedRecordCount === 1 &&
                selectedAttachments.map((att, index) => {
                  const thumbUrl = selectedThumbnailUrls[index] || '';
                  const isActive = index === selectedPreviewIndex;
                  const isDragging = index === dragIndex;
                  const isDragOver = index === dragOverIndex;
                  return (
                    <div
                      className={`thumb-item${isActive ? ' active' : ''}${isDragging ? ' dragging' : ''}${isDragOver ? ' drag-over' : ''}`}
                      key={`${att.token || att.name}-${index}`}
                      draggable={!reorderSaving}
                      onDragStart={() => handleDragStart(index)}
                      onDragOver={(event) => handleDragOver(event, index)}
                      onDrop={() => handleDrop(index)}
                      onDragEnd={handleDragEnd}
                      onClick={() => setSelectedPreviewIndex(index)}
                      title={att.name}
                    >
                      {thumbUrl ? (
                        <img className="thumb-image" src={thumbUrl} alt={att.name} />
                      ) : (
                        <div className="thumb-placeholder">{getFileTag(att.name)}</div>
                      )}
                      <div className="thumb-name">{att.name}</div>
                    </div>
                  );
                })}
              {!selectedLoading && selectedRecordCount !== 1 && <div className="empty">请先选中单行</div>}
              {!selectedLoading && selectedRecordCount === 1 && selectedAttachments.length === 0 && (
                <div className="empty">暂无附件缩略图</div>
              )}
            </div>
          </div>
        )}
      </section>

      <section className="card">
        <div className="section-title">1. 选择附件字段</div>
        <Select
          className="field-select"
          placeholder="请选择附件字段"
          value={attachmentFieldId || undefined}
          onChange={(value) => setAttachmentFieldId(value as string)}
          optionList={fieldOptions}
          filter
          disabled={processing}
        />
        <div className="hint">只会处理这个附件列</div>
      </section>

      <section className="card">
        <div className="section-title">2. 选择作用范围</div>
        <Radio.Group value={scope} onChange={(e) => handleScopeChange(e.target.value as Scope)} disabled={processing}>
          <Radio value="selected">已选记录</Radio>
          <Radio value="view">当前视图</Radio>
          <Radio value="all">全表</Radio>
        </Radio.Group>
        <div className="hint">
          预计处理记录数：{estimatedCount === null ? '—' : estimatedCount}
          {scope === 'selected' && manualSelectedRecordIds.length > 0 ? `（已选择 ${manualSelectedRecordIds.length} 条）` : ''}
        </div>
      </section>

      <section className="card">
        <div className="section-title">3. 变量设置</div>
        <div className="token-bar">
          <span className="token-label">可插入变量：</span>
          <Button size="small" onClick={() => handleInsertToken(TOKEN_SEQ)} disabled={processing}>
            序号
          </Button>
          {variableOptions.map((opt) => (
            <Button
              key={opt.value}
              size="small"
              onClick={() => handleInsertToken(`{{${opt.label}}}`)}
              disabled={processing}
            >
              {opt.label}
            </Button>
          ))}
          <span className="token-hint">先点击输入框，再点击按钮插入变量</span>
        </div>
        <div className="row">
          <span className="label">自动序号</span>
          <Space>
            <span className="inline-label">开始</span>
            <InputNumber
              min={0}
              value={seqStart}
              onChange={(value) => setSeqStart(value === null || value === undefined ? 1 : Number(value))}
              disabled={processing}
            />
            <span className="inline-label">位数</span>
            <InputNumber
              min={0}
              value={seqPad}
              onChange={(value) => setSeqPad(value === null || value === undefined ? 0 : Number(value))}
              disabled={processing}
            />
          </Space>
        </div>
        <div className="hint">每条记录内序号从起始值重新计数；变量格式为 {"{{"}字段名{"}}"}</div>
      </section>

      <section className="card">
        <div className="section-title">4. 命名规则</div>
        <Tabs activeKey={mode} onChange={(key) => setMode(key as Mode)} type="line">
          <TabPane tab="替换" itemKey="replace">
            <div className="row">
              <span className="label">新名称模板</span>
              <Input
                ref={replaceInputRef}
                value={replaceTemplate}
                onChange={(value) => setReplaceTemplate(value)}
                onFocus={() => { setActiveInput('replace'); setCursorPosition(replaceTemplate.length); }}
                onSelect={(e) => setCursorPosition((e.target as HTMLInputElement).selectionStart ?? replaceTemplate.length)}
                placeholder="可输入固定文本，并支持 {{序号}} {{字段名}}"
                disabled={processing}
              />
            </div>
            <div className="hint">仅替换文件名主体，扩展名保持不变</div>
          </TabPane>

          <TabPane tab="追加" itemKey="append">
            <div className="row">
              <span className="label">追加位置</span>
              <Radio.Group
                value={appendPosition}
                onChange={(e) => setAppendPosition(e.target.value as AppendPosition)}
                disabled={processing}
              >
                <Radio value="prepend">名称前</Radio>
                <Radio value="insert">指定位置</Radio>
                <Radio value="append">名称后</Radio>
              </Radio.Group>
            </div>
            {appendPosition === 'insert' && (
              <div className="row">
                <span className="label">插入索引</span>
                <InputNumber
                  min={0}
                  value={insertIndex}
                  onChange={(value) => setInsertIndex(value === null || value === undefined ? 0 : Number(value))}
                  disabled={processing}
                />
                <span className="hint-inline">索引从 0 开始，不包含扩展名</span>
              </div>
            )}

            <div className="row">
              <span className="label">前面字符</span>
              <Input
                ref={frontInputRef}
                value={appendFront}
                onChange={(value) => setAppendFront(value)}
                onFocus={() => { setActiveInput('front'); setCursorPosition(appendFront.length); }}
                onSelect={(e) => setCursorPosition((e.target as HTMLInputElement).selectionStart ?? appendFront.length)}
                placeholder="可插入 {{字段名}}"
                disabled={processing}
              />
            </div>
            <div className="row">
              <span className="label">后面字符</span>
              <Input
                ref={backInputRef}
                value={appendBack}
                onChange={(value) => setAppendBack(value)}
                onFocus={() => { setActiveInput('back'); setCursorPosition(appendBack.length); }}
                onSelect={(e) => setCursorPosition((e.target as HTMLInputElement).selectionStart ?? appendBack.length)}
                placeholder="可插入 {{字段名}}"
                disabled={processing}
              />
            </div>
            <div className="hint">追加规则：前面字符 + 序号 + 后面字符</div>
          </TabPane>
        </Tabs>
      </section>

      <section className="card">
        <div className="section-title">5. 预览（最多 {PREVIEW_LIMIT} 条）</div>
        {previewLoading && <div className="hint">预览生成中…</div>}
        {!previewLoading && previewItems.length === 0 && <div className="empty">暂无预览</div>}
        <div className="preview-list">
          {previewItems.map((item) => {
            const parts = getDiffParts(item.oldName, item.newName);
            return (
              <div className="preview-row" key={item.id}>
                <div className="preview-label">{item.label}</div>
                <div className="preview-names">
                  <span className="old-name">{item.oldName}</span>
                  <span className="arrow">→</span>
                  <span className="new-name">
                    {parts.prefix}
                    {parts.diff && <span className="diff">{parts.diff}</span>}
                    {parts.suffix}
                  </span>
                </div>
              </div>
            );
          })}
        </div>
      </section>

      <section className="footer">
        <Space>
          <Button theme="solid" type="primary" onClick={handleStart} disabled={!canStart}>
            开始重命名
          </Button>
          <Button
            theme="light"
            type="tertiary"
            onClick={handleUndo}
            disabled={undoStack.length === 0 || processing}
          >
            撤销上次（支持撤销 {undoStack.length} 次）
          </Button>
        </Space>
        <Typography.Text type="tertiary" className="footer-note">
          同名自动添加后缀（_1、_2），扩展名不变
        </Typography.Text>
      </section>

      <section className="card status">
        <div className="section-title">处理进度</div>
        <Progress percent={percent} showInfo={!processing} strokeWidth={6} />
        {processing && (
          <Typography.Text type="secondary">
            正在处理 {progress.current}/{progress.total}
          </Typography.Text>
        )}
        {result && (
          <div className="result">
            <span>总计：{result.total}</span>
            <span>成功：{result.success}</span>
            <span>失败：{result.failed}</span>
          </div>
        )}
      </section>
    </main>
  );
}
