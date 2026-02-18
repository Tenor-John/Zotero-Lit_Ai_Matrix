type MatrixAIKey =
  | "领域基础知识"
  | "研究背景"
  | "作者的问题意识"
  | "研究意义"
  | "研究结论"
  | "未来研究方向提及"
  | "未来研究方向思考";

type MatrixPayload = {
  schema: "lms.v1";
  sourceNoteID: number;
  sourceNoteTitle: string;
  sourceNoteDateModified: string;
  status: string;
  ai: Record<MatrixAIKey, string>;
};

const MATRIX_JSON_PREFIX = "LMS_MATRIX_JSON: ";
const MATRIX_INCLUDE_TAG = "lms-matrix";
const MATRIX_IGNORE_TAG = "lms-ignore";
const MATRIX_PRIMARY_TAG = "matrix-primary";
const MATRIX_TITLE_PREFIX = "[MATRIX]";
const MATRIX_EXCLUDE_PREFIXES = ["[ANNOT]", "[DRAFT]"];

const AI_FIELDS: MatrixAIKey[] = [
  "领域基础知识",
  "研究背景",
  "作者的问题意识",
  "研究意义",
  "研究结论",
  "未来研究方向提及",
  "未来研究方向思考",
];

const READ_STATUS_FIELDS = ["状态", "阅读状态"];
const READ_STATUS_MAP: Record<string, string> = {
  已读: "done",
  在读: "reading",
  未读: "unread",
  done: "done",
  reading: "reading",
  unread: "unread",
};

const COLUMN_DEFS: Array<{
  dataKey: string;
  label: string;
  iconPath?: string;
  provider: (item: Zotero.Item) => string;
}> = [
  {
    dataKey: "lms_title",
    label: "标题",
    provider: (item) => item.getDisplayTitle() || "",
  },
  {
    dataKey: "lms_journal",
    label: "期刊",
    provider: (item) => String(item.getField("publicationTitle") || ""),
  },
  {
    dataKey: "lms_year",
    label: "年份",
    provider: (item) => extractYear(String(item.getField("date") || "")),
  },
  {
    dataKey: "lms_status",
    label: "状态",
    provider: (item) => {
      const status = getMatrixPayload(item)?.status || "unread";
      return status === "done"
        ? "已读"
        : status === "reading"
          ? "在读"
          : "未读";
    },
  },
  {
    dataKey: "lms_tags",
    label: "标签",
    provider: (item) =>
      (item.getTags() || [])
        .map((t) => String(t.tag || "").trim())
        .filter(Boolean)
        .join(", "),
  },
  {
    dataKey: "lms_updated",
    label: "更新日期",
    provider: (item) => String(item.dateModified || "").slice(0, 10),
  },
  ...AI_FIELDS.map((field) => ({
    dataKey: `lms_ai_${toSafeKey(field)}`,
    label: field,
    provider: (item: Zotero.Item) => getMatrixPayload(item)?.ai[field] || "",
  })),
];

const registeredColumnKeys: string[] = [];
let notifierID = "";
let isRebuilding = false;
const CHUNK_SIZE = 200;
const MATRIX_UI_VERSION = "2026-02-18-2038";
const MATRIX_NAV_BUTTON_ID = "zotero-toolbarbutton-lms-matrix";
const MATRIX_PAGE_ROOT_ID = "lms-matrix-page-root";
const matrixTabIDs = new WeakMap<Window, string>();

// Auto-rebuild can be triggered by Zotero sync/bulk edits.
// Use a debounced, de-duplicated queue to avoid UI stutters.
const AUTO_REBUILD_DEBOUNCE_MS = 1500;
const AUTO_REBUILD_BATCH_SIZE = 25;
const AUTO_REBUILD_MAX_PER_FLUSH = 300;
const pendingAutoRebuildIDs = new Set<number>();
let autoRebuildTimer: Promise<void> | null = null;
let autoRebuildRunning = false;

export class MatrixFeature {
  static async startup() {
    await this.registerColumns();
    this.registerMenu();
    this.registerNotifier();
    Zotero.getMainWindows().forEach((win) => this.onMainWindowLoad(win));
  }

  static shutdown() {
    if (notifierID) {
      Zotero.Notifier.unregisterObserver(notifierID);
      notifierID = "";
    }
    pendingAutoRebuildIDs.clear();
    autoRebuildTimer = null;
    autoRebuildRunning = false;
    Zotero.getMainWindows().forEach((win) => this.onMainWindowUnload(win));
    for (const dataKey of registeredColumnKeys) {
      const unregisterColumns = (
        Zotero.ItemTreeManager as unknown as {
          unregisterColumns?: (pluginID: string, dataKey: string) => void;
        }
      ).unregisterColumns;
      if (unregisterColumns) {
        unregisterColumns(addon.data.config.addonID, dataKey);
      }
    }
    registeredColumnKeys.length = 0;
  }

  static onMainWindowLoad(win: _ZoteroTypes.MainWindow) {
    this.registerNavButton(win);
  }

  static onMainWindowUnload(win: Window) {
    const doc = win.document;
    doc.getElementById(MATRIX_NAV_BUTTON_ID)?.remove();
    const tabID = matrixTabIDs.get(win);
    if (tabID) {
      try {
        ztoolkit.getGlobal("Zotero_Tabs").close(tabID);
      } catch {
        // ignore
      }
      matrixTabIDs.delete(win);
    }
  }

  static async rebuildSelectedItems() {
    const selected = ztoolkit.getGlobal("ZoteroPane").getSelectedItems() || [];
    const regularItems = selected.filter((item: Zotero.Item) =>
      item.isRegularItem(),
    );
    if (!regularItems.length) {
      new ztoolkit.ProgressWindow(addon.data.config.addonName)
        .createLine({
          text: "未选中文献条目",
          type: "default",
          progress: 100,
        })
        .show();
      return;
    }
    await rebuildForRegularItems(regularItems);
    new ztoolkit.ProgressWindow(addon.data.config.addonName)
      .createLine({
        text: `矩阵缓存已更新 (${regularItems.length} 条)`,
        type: "success",
        progress: 100,
      })
      .show();
  }

  static async rebuildAllItems() {
    const popup = new ztoolkit.ProgressWindow(addon.data.config.addonName, {
      closeOnClick: true,
      closeTime: -1,
    })
      .createLine({
        text: "正在扫描全库文献...",
        type: "default",
        progress: 0,
      })
      .show();

    try {
      const allItems = await getAllRegularItems();
      if (!allItems.length) {
        popup.changeLine({
          text: "未找到可处理的文献条目",
          type: "default",
          progress: 100,
        });
        popup.startCloseTimer(2500);
        return;
      }

      let processed = 0;
      for (let i = 0; i < allItems.length; i += CHUNK_SIZE) {
        const chunk = allItems.slice(i, i + CHUNK_SIZE);
        await rebuildForRegularItems(chunk);
        processed += chunk.length;
        const progress = Math.min(
          99,
          Math.floor((processed / allItems.length) * 100),
        );
        popup.changeLine({
          text: `重建中... ${processed}/${allItems.length}`,
          type: "default",
          progress,
        });
      }
      popup.changeLine({
        text: `全库矩阵缓存重建完成 (${allItems.length} 条)`,
        type: "success",
        progress: 100,
      });
      popup.startCloseTimer(5000);
    } catch (e) {
      ztoolkit.log("rebuildAllItems error", e);
      popup.changeLine({
        text: "全库重建失败，请查看日志",
        type: "default",
        progress: 100,
      });
    }
  }

  static async setStatusForSelectedItems(
    status: "done" | "reading" | "unread",
  ) {
    const selected = ztoolkit.getGlobal("ZoteroPane").getSelectedItems() || [];
    const regularItems = selected.filter((item: Zotero.Item) =>
      item.isRegularItem(),
    );
    if (!regularItems.length) {
      new ztoolkit.ProgressWindow(addon.data.config.addonName)
        .createLine({
          text: "未选中文献条目",
          type: "default",
          progress: 100,
        })
        .show();
      return;
    }
    await updateStatusForItems(regularItems, status);
    const statusCN =
      status === "done" ? "已读" : status === "reading" ? "在读" : "未读";
    new ztoolkit.ProgressWindow(addon.data.config.addonName)
      .createLine({
        text: `已将 ${regularItems.length} 条文献标记为${statusCN}`,
        type: "success",
        progress: 100,
      })
      .show();
  }

  static openMatrixPage() {
    const win = Zotero.getMainWindow();
    if (!win) {
      return;
    }
    this.openMatrixInTab(win);
  }

  private static async registerColumns() {
    for (const column of COLUMN_DEFS) {
      await Zotero.ItemTreeManager.registerColumns({
        pluginID: addon.data.config.addonID,
        dataKey: column.dataKey,
        label: column.label,
        dataProvider: (item: Zotero.Item) => {
          if (!item.isRegularItem()) {
            return "";
          }
          return column.provider(item);
        },
        iconPath:
          column.iconPath || "chrome://zotero/skin/16/universal/book.svg",
      });
      registeredColumnKeys.push(column.dataKey);
    }
  }

  private static registerMenu() {
    ztoolkit.Menu.register("item", {
      tag: "menuitem",
      id: "zotero-itemmenu-lms-rebuild-matrix",
      label: "重建智能文献矩阵缓存",
      commandListener: () => {
        void MatrixFeature.rebuildSelectedItems();
      },
    });
    ztoolkit.Menu.register("item", {
      tag: "menuitem",
      id: "zotero-itemmenu-lms-mark-done",
      label: "矩阵状态: 标记已读",
      commandListener: () => {
        void MatrixFeature.setStatusForSelectedItems("done");
      },
    });
    ztoolkit.Menu.register("item", {
      tag: "menuitem",
      id: "zotero-itemmenu-lms-mark-reading",
      label: "矩阵状态: 标记在读",
      commandListener: () => {
        void MatrixFeature.setStatusForSelectedItems("reading");
      },
    });
    ztoolkit.Menu.register("item", {
      tag: "menuitem",
      id: "zotero-itemmenu-lms-mark-unread",
      label: "矩阵状态: 标记未读",
      commandListener: () => {
        void MatrixFeature.setStatusForSelectedItems("unread");
      },
    });
    ztoolkit.Menu.register("menuFile", {
      tag: "menuitem",
      id: "zotero-filemenu-lms-open-page",
      label: "打开智能文献矩阵页面",
      commandListener: () => {
        MatrixFeature.openMatrixPage();
      },
    });
    ztoolkit.Menu.register("menuFile", {
      tag: "menuitem",
      id: "zotero-filemenu-lms-rebuild-all",
      label: "重建全库智能文献矩阵缓存",
      commandListener: () => {
        void MatrixFeature.rebuildAllItems();
      },
    });
  }

  private static registerNavButton(win: _ZoteroTypes.MainWindow) {
    const doc = win.document;
    if (doc.getElementById(MATRIX_NAV_BUTTON_ID)) {
      return;
    }
    const btn = doc.createXULElement("toolbarbutton");
    btn.id = MATRIX_NAV_BUTTON_ID;
    btn.setAttribute("class", "zotero-tb-button");
    btn.setAttribute("tooltiptext", "打开智能文献矩阵");
    btn.setAttribute("type", "button");
    btn.setAttribute("style", "padding:0 2px;");
    const icon = doc.createXULElement("image");
    icon.setAttribute(
      "style",
      `list-style-image: url("chrome://${addon.data.config.addonRef}/content/ICON/Ai_matrix.png"); width: 18px; height: 18px;`,
    );
    btn.appendChild(icon);
    btn.addEventListener("command", () => {
      MatrixFeature.openMatrixPage();
    });

    const anchor = doc.querySelector("#zoterogpt") as Element | null;
    if (anchor?.parentElement) {
      anchor.parentElement.insertBefore(btn, anchor.nextSibling);
      return;
    }
    const toolbar = doc.querySelector(
      "#zotero-items-toolbar, #zotero-tb-actions, #zotero-toolbar, #zotero-title-bar, toolbar",
    ) as Element | null;
    toolbar?.appendChild(btn);
  }

  private static openMatrixInTab(win: _ZoteroTypes.MainWindow) {
    const tabs = ztoolkit.getGlobal("Zotero_Tabs");
    const existingTabID = matrixTabIDs.get(win);
    if (
      existingTabID &&
      tabs._tabs.some((t: _ZoteroTypes.TabInstance) => t.id === existingTabID)
    ) {
      tabs.select(existingTabID);
      const root = win.document.getElementById(
        `${MATRIX_PAGE_ROOT_ID}-${existingTabID}`,
      ) as HTMLDivElement | null;
      if (root) {
        void renderMatrixPage(win, root);
      }
      return;
    }

    const added = tabs.add({
      type: "library",
      title: "智能文献矩阵",
      data: { itemID: 0 },
      select: true,
      onClose: () => {
        matrixTabIDs.delete(win);
      },
    });
    const tabID = added.id;
    matrixTabIDs.set(win, tabID);
    const iconURL = `chrome://${addon.data.config.addonRef}/content/ICON/Ai_matrix.png`;
    const applyTabIcon = () => {
      try {
        const tabRoot = win.document.querySelector(
          `.tab[data-id="${tabID}"]`,
        ) as HTMLElement | null;
        const iconSpan = (tabRoot?.querySelector(".tab-icon") ||
          added.container.querySelector(".tab-icon")) as HTMLElement | null;
        if (!iconSpan) {
          return false;
        }
        iconSpan.style.backgroundImage = `url("${iconURL}")`;
        iconSpan.style.backgroundRepeat = "no-repeat";
        iconSpan.style.backgroundPosition = "center";
        iconSpan.style.backgroundSize = "16px 16px";
        iconSpan.style.width = "16px";
        iconSpan.style.height = "16px";
        iconSpan.style.display = "inline-block";
        iconSpan.style.minWidth = "16px";
        iconSpan.style.minHeight = "16px";
        iconSpan.classList.remove("icon-item-type");
        return true;
      } catch {
        return false;
      }
    };
    if (!applyTabIcon()) {
      win.setTimeout(() => {
        applyTabIcon();
      }, 120);
    }
    const root = win.document.createElementNS(
      "http://www.w3.org/1999/xhtml",
      "div",
    ) as HTMLDivElement;
    root.id = `${MATRIX_PAGE_ROOT_ID}-${tabID}`;
    root.style.height = "100%";
    root.style.overflow = "auto";
    root.style.background = "#fff";
    added.container.appendChild(root);
    void renderMatrixPage(win, root);
  }

  private static registerNotifier() {
    const callback = {
      notify: async (
        event: string,
        type: string,
        ids: number[] | string[],
        _extraData: { [key: string]: any },
      ) => {
        if (!addon?.data.alive || type !== "item" || isRebuilding) {
          return;
        }
        if (!["add", "modify"].includes(event)) {
          return;
        }
        const numberIDs = ids
          .map((id) => Number(id))
          .filter((id) => Number.isInteger(id) && id > 0);
        if (!numberIDs.length) {
          return;
        }
        const items = await Zotero.Items.getAsync(numberIDs);
        const targetRegularItems: Zotero.Item[] = [];
        for (const item of items) {
          if (!item) {
            continue;
          }
          if (item.isRegularItem()) {
            targetRegularItems.push(item);
            continue;
          }
          if (item.isNote() && item.parentID) {
            const parentItem = Zotero.Items.get(item.parentID);
            if (parentItem?.isRegularItem()) {
              targetRegularItems.push(parentItem);
            }
          }
        }
        if (!targetRegularItems.length) {
          return;
        }

        enqueueAutoRebuild(targetRegularItems);
      },
    };

    notifierID = Zotero.Notifier.registerObserver(callback, ["item"]);
  }
}

function enqueueAutoRebuild(items: Zotero.Item[]) {
  for (const item of items) {
    if (item?.id) {
      pendingAutoRebuildIDs.add(item.id);
    }
  }
  scheduleAutoRebuildFlush();
}

function scheduleAutoRebuildFlush() {
  if (autoRebuildTimer) {
    return;
  }
  autoRebuildTimer = Zotero.Promise.delay(AUTO_REBUILD_DEBOUNCE_MS).then(
    async () => {
      autoRebuildTimer = null;
      await flushAutoRebuildQueue();
    },
  );
}

async function flushAutoRebuildQueue() {
  if (autoRebuildRunning) {
    // Another flush is in progress; it will pick up new IDs.
    return;
  }
  if (!addon?.data.alive) {
    pendingAutoRebuildIDs.clear();
    return;
  }

  autoRebuildRunning = true;
  try {
    let processedThisFlush = 0;

    while (pendingAutoRebuildIDs.size > 0) {
      if (!addon?.data.alive) {
        pendingAutoRebuildIDs.clear();
        return;
      }
      if (processedThisFlush >= AUTO_REBUILD_MAX_PER_FLUSH) {
        // Avoid long blocking; continue later.
        scheduleAutoRebuildFlush();
        return;
      }
      if (isRebuilding) {
        // A manual rebuild is running; postpone.
        scheduleAutoRebuildFlush();
        return;
      }

      const batchIDs: number[] = [];
      for (const id of pendingAutoRebuildIDs) {
        batchIDs.push(id);
        if (batchIDs.length >= AUTO_REBUILD_BATCH_SIZE) {
          break;
        }
      }
      batchIDs.forEach((id) => pendingAutoRebuildIDs.delete(id));

      const batchItems = await Zotero.Items.getAsync(batchIDs);
      const regularItems = batchItems.filter((it): it is Zotero.Item =>
        Boolean(it && it.isRegularItem()),
      );

      if (regularItems.length) {
        isRebuilding = true;
        try {
          for (const item of regularItems) {
            await rebuildSingleItemMatrix(item);
          }
        } finally {
          isRebuilding = false;
        }
      }

      processedThisFlush += batchIDs.length;

      // Yield to keep UI responsive.
      await Zotero.Promise.delay(0);
    }
  } catch (e) {
    ztoolkit.log("flushAutoRebuildQueue error", e);
  } finally {
    autoRebuildRunning = false;
    // If new items arrived while we were running, schedule another flush.
    if (pendingAutoRebuildIDs.size > 0) {
      scheduleAutoRebuildFlush();
    }
  }
}

async function rebuildForRegularItems(items: Zotero.Item[]) {
  if (!items.length) {
    return;
  }
  isRebuilding = true;
  try {
    for (const item of items) {
      await rebuildSingleItemMatrix(item);
    }
  } finally {
    isRebuilding = false;
  }
}

async function getAllRegularItems(): Promise<Zotero.Item[]> {
  const search = new Zotero.Search();
  search.addCondition("itemType", "isNot", "attachment");
  search.addCondition("itemType", "isNot", "note");
  const ids = await search.search();
  const rawItems = await Zotero.Items.getAsync(ids);
  return rawItems.filter((item): item is Zotero.Item => {
    return Boolean(item && item.isRegularItem());
  });
}

async function updateStatusForItems(
  items: Zotero.Item[],
  status: "done" | "reading" | "unread",
) {
  isRebuilding = true;
  try {
    for (const item of items) {
      const payload =
        getMatrixPayload(item) ||
        ({
          schema: "lms.v1",
          sourceNoteID: 0,
          sourceNoteTitle: "",
          sourceNoteDateModified: "",
          status: "unread",
          ai: Object.fromEntries(AI_FIELDS.map((k) => [k, ""])) as Record<
            MatrixAIKey,
            string
          >,
        } as MatrixPayload);
      payload.status = status;
      await upsertMatrixPayload(item, payload);
    }
  } finally {
    isRebuilding = false;
  }
}

async function rebuildSingleItemMatrix(item: Zotero.Item) {
  const noteIDs = item.getNotes();
  if (!noteIDs.length) {
    await clearMatrixPayload(item);
    return;
  }
  const notes = await Zotero.Items.getAsync(noteIDs);
  const candidates: Array<{
    note: Zotero.Item;
    title: string;
    plainText: string;
    tags: string[];
    score: number;
    modifiedMs: number;
  }> = [];

  for (const note of notes) {
    if (!note?.isNote()) {
      continue;
    }
    const title = String(note.getField("title") || "").trim();
    const tags = (note.getTags() || []).map((t) => String(t.tag || "").trim());
    const plainText = normalizeText(note.getNote());
    const judge = judgeMatrixNote(title, tags, plainText);
    if (!judge.accepted) {
      continue;
    }
    candidates.push({
      note,
      title,
      plainText,
      tags,
      score: judge.score,
      modifiedMs: parseDateMs(note.dateModified),
    });
  }

  if (!candidates.length) {
    await clearMatrixPayload(item);
    return;
  }

  candidates.sort((a, b) => {
    if (b.score !== a.score) {
      return b.score - a.score;
    }
    return b.modifiedMs - a.modifiedMs;
  });

  const winner = candidates[0];
  const ai = extractAIMatrixFields(winner.plainText);
  const status = extractReadStatus(winner.plainText);
  const payload: MatrixPayload = {
    schema: "lms.v1",
    sourceNoteID: winner.note.id,
    sourceNoteTitle: winner.title,
    sourceNoteDateModified: winner.note.dateModified || "",
    status,
    ai,
  };
  await upsertMatrixPayload(item, payload);
}

function judgeMatrixNote(title: string, tags: string[], plainText: string) {
  const hasIgnoreTag = tags.includes(MATRIX_IGNORE_TAG);
  if (hasIgnoreTag) {
    return { accepted: false, fieldCount: 0, score: 0 };
  }
  if (MATRIX_EXCLUDE_PREFIXES.some((prefix) => title.startsWith(prefix))) {
    return { accepted: false, fieldCount: 0, score: 0 };
  }
  const includeByTag = tags.includes(MATRIX_INCLUDE_TAG);
  const includeByTitle = title.startsWith(MATRIX_TITLE_PREFIX);
  const fieldCount = AI_FIELDS.filter((k) =>
    plainText.includes(`${k}::`),
  ).length;
  const includeByFields = fieldCount >= 3;
  if (!includeByTag && !includeByTitle && !includeByFields) {
    return { accepted: false, fieldCount: 0, score: 0 };
  }
  if (fieldCount === 0) {
    return { accepted: false, fieldCount: 0, score: 0 };
  }
  let score = 0;
  if (tags.includes(MATRIX_PRIMARY_TAG)) {
    score += 100;
  }
  if (includeByTitle) {
    score += 40;
  }
  if (includeByTag) {
    score += 20;
  }
  score += fieldCount * 2;
  return { accepted: true, fieldCount, score };
}

function extractAIMatrixFields(plainText: string): Record<MatrixAIKey, string> {
  const result = {} as Record<MatrixAIKey, string>;
  for (const field of AI_FIELDS) {
    result[field] = extractDataviewField(plainText, field);
  }
  return result;
}

function extractReadStatus(plainText: string): string {
  for (const key of READ_STATUS_FIELDS) {
    const raw = extractDataviewField(plainText, key);
    if (!raw) {
      continue;
    }
    const normalized = READ_STATUS_MAP[raw.trim().toLowerCase()];
    if (normalized) {
      return normalized;
    }
    const cnNormalized = READ_STATUS_MAP[raw.trim()];
    if (cnNormalized) {
      return cnNormalized;
    }
  }
  return "unread";
}

function extractDataviewField(text: string, key: string): string {
  const marker = `${key}::`;
  const start = text.indexOf(marker);
  if (start < 0) {
    return "";
  }
  const valueStart = start + marker.length;
  let valueEnd = text.length;
  for (const candidate of [...AI_FIELDS, ...READ_STATUS_FIELDS]) {
    if (candidate === key) {
      continue;
    }
    const nextIndex = text.indexOf(`${candidate}::`, valueStart);
    if (nextIndex >= 0 && nextIndex < valueEnd) {
      valueEnd = nextIndex;
    }
  }
  return text.slice(valueStart, valueEnd).replace(/\s+/g, " ").trim();
}

function normalizeText(noteHTML: string): string {
  return String(noteHTML || "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/\r/g, "")
    .trim();
}

function parseDateMs(dateLike: string): number {
  const d = new Date(dateLike);
  if (Number.isNaN(d.getTime())) {
    return 0;
  }
  return d.getTime();
}

function extractYear(dateText: string): string {
  const m = dateText.match(/\b(19|20)\d{2}\b/);
  return m ? m[0] : "";
}

function toSafeKey(key: string): string {
  return key
    .replace(/[^\w]/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "")
    .toLowerCase();
}

function getMatrixPayload(item: Zotero.Item): MatrixPayload | null {
  const extra = String(item.getField("extra") || "");
  const line = extra
    .split(/\r?\n/)
    .find((l) => l.trim().startsWith(MATRIX_JSON_PREFIX));
  if (!line) {
    return null;
  }
  const rawJSON = line.trim().slice(MATRIX_JSON_PREFIX.length).trim();
  if (!rawJSON) {
    return null;
  }
  try {
    const parsed = JSON.parse(rawJSON) as MatrixPayload;
    if (parsed?.schema !== "lms.v1") {
      return null;
    }
    return parsed;
  } catch (e) {
    ztoolkit.log("Invalid matrix payload JSON", e);
    return null;
  }
}

async function upsertMatrixPayload(item: Zotero.Item, payload: MatrixPayload) {
  const extra = String(item.getField("extra") || "");
  const lines = extra.split(/\r?\n/).filter((line) => line.length > 0);
  const serialized = `${MATRIX_JSON_PREFIX}${JSON.stringify(payload)}`;
  const idx = lines.findIndex((line) =>
    line.trim().startsWith(MATRIX_JSON_PREFIX),
  );
  if (idx >= 0) {
    lines[idx] = serialized;
  } else {
    lines.push(serialized);
  }
  const nextExtra = lines.join("\n");
  if (nextExtra === extra) {
    return;
  }
  item.setField("extra", nextExtra);
  await item.saveTx();
}

async function clearMatrixPayload(item: Zotero.Item) {
  const extra = String(item.getField("extra") || "");
  if (!extra.includes(MATRIX_JSON_PREFIX)) {
    return;
  }
  const lines = extra
    .split(/\r?\n/)
    .filter((line) => !line.trim().startsWith(MATRIX_JSON_PREFIX));
  const nextExtra = lines.join("\n");
  if (nextExtra === extra) {
    return;
  }
  item.setField("extra", nextExtra);
  await item.saveTx();
}

type MatrixPageRow = {
  itemID: number;
  title: string;
  author: string;
  journal: string;
  year: string;
  status: string;
  tags: string;
  added: string;
  updated: string;
  activityDate: string;
  hasPDF: boolean;
  ai: Record<MatrixAIKey, string>;
};

function hasPdfAttachment(item: Zotero.Item): boolean {
  try {
    if (item?.isPDFAttachment?.()) {
      return true;
    }
    if (!item?.isRegularItem?.()) {
      return false;
    }
    const attachmentIDs = item.getAttachments?.() || [];
    for (const id of attachmentIDs) {
      const att = Zotero.Items.get(id);
      if (att?.isPDFAttachment?.()) {
        return true;
      }
    }
  } catch {
    // ignore
  }
  return false;
}

async function renderMatrixPage(win: Window, rootOverride?: HTMLDivElement) {
  const doc = win.document;
  const root =
    rootOverride ||
    (doc.getElementById(MATRIX_PAGE_ROOT_ID) as HTMLDivElement | null);
  if (!root) {
    return;
  }
  root.innerHTML = `<div style="padding:16px;font-size:14px;">正在加载矩阵数据...</div>`;
  const rows = await buildMatrixPageRows();

  const years = Array.from(
    new Set(rows.map((r) => String(r.year || "").trim()).filter(Boolean)),
  ).sort((a, b) => Number(b) - Number(a));
  const journals = Array.from(
    new Set(rows.map((r) => normalizeVisibleText(r.journal) || "")),
  ).sort((a, b) =>
    a.localeCompare(b, "zh-CN", { numeric: true, sensitivity: "base" }),
  );

  const state = ((win as any).__lmsMatrixPageState as
    | {
        q: string;
        status: string;
        year: string;
        journal: string;
        updatedRange: string;
        tags: string;
        sortKey: string;
        sortDir: "asc" | "desc";
      }
    | undefined) || {
    q: "",
    status: "all",
    year: "all",
    journal: "all",
    updatedRange: "all",
    tags: "",
    sortKey: "added",
    sortDir: "desc",
  };
  (win as any).__lmsMatrixPageState = state;
  root.innerHTML = `
    <div xmlns="http://www.w3.org/1999/xhtml" style="display:flex;flex-direction:column;height:100%;font-family:system-ui,-apple-system,Segoe UI,sans-serif;color:#222;">
      <div style="padding:12px 12px 8px 12px;border-bottom:1px solid #ddd;background:#f8fafc;">
        <div style="display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;">
          <div style="font-size:18px;font-weight:700;">智能文献矩阵</div>
          <div style="font-size:12px;color:#64748b;">总文献: ${rows.length} | UI:${MATRIX_UI_VERSION}</div>
        </div>
        <div style="display:flex;gap:8px;margin-top:8px;align-items:center;flex-wrap:wrap;">
          <input id="lms-matrix-q" type="text" placeholder="搜索标题/期刊/AI字段" value="${escapeHTML(state.q)}" style="min-width:320px;flex:1 1 320px;padding:6px 10px;border:1px solid #cbd5e1;border-radius:6px;"/>
          <button id="lms-matrix-export-csv" style="padding:6px 10px;border:1px solid #94a3b8;background:#fff;border-radius:6px;cursor:pointer;">导出CSV</button>
          <button id="lms-matrix-export-md" style="padding:6px 10px;border:1px solid #94a3b8;background:#fff;border-radius:6px;cursor:pointer;">导出Markdown</button>
          <button id="lms-matrix-refresh" style="padding:6px 10px;border:1px solid #94a3b8;background:#fff;border-radius:6px;cursor:pointer;">刷新</button>
        </div>
      </div>
      <div id="lms-matrix-filters-wrap" style="padding:10px 12px;background:#fff;border-bottom:1px solid #eee;">
        <div style="border:1px solid #e2e8f0;border-radius:10px;padding:12px;background:#f8fafc;">
          <div style="display:grid;grid-template-columns:repeat(4,minmax(180px,1fr));gap:12px;align-items:end;">
            <div>
              <div style="font-size:12px;color:#475569;margin-bottom:6px;">阅读状态</div>
              <select id="lms-matrix-status" style="width:100%;padding:8px 10px;border:1px solid #cbd5e1;border-radius:8px;background:#fff;">
                <option value="all">全部状态</option>
                <option value="done">已读</option>
                <option value="reading">在读</option>
                <option value="unread">未读</option>
              </select>
            </div>
            <div>
              <div style="font-size:12px;color:#475569;margin-bottom:6px;">年份</div>
              <select id="lms-matrix-year" style="width:100%;padding:8px 10px;border:1px solid #cbd5e1;border-radius:8px;background:#fff;">
                <option value="all">全部年份</option>
                ${years
                  .map(
                    (y) =>
                      `<option value="${escapeHTML(y)}">${escapeHTML(y)}</option>`,
                  )
                  .join("")}
              </select>
            </div>
            <div>
              <div style="font-size:12px;color:#475569;margin-bottom:6px;">分类</div>
              <select id="lms-matrix-journal" style="width:100%;padding:8px 10px;border:1px solid #cbd5e1;border-radius:8px;background:#fff;">
                <option value="all">全部分类</option>
                <option value="__empty__">未分类</option>
                ${journals
                  .filter((j) => j)
                  .map(
                    (j) =>
                      `<option value="${escapeHTML(j)}">${escapeHTML(j)}</option>`,
                  )
                  .join("")}
              </select>
            </div>
            <div>
              <div style="font-size:12px;color:#475569;margin-bottom:6px;">修改时间</div>
              <select id="lms-matrix-updated-range" style="width:100%;padding:8px 10px;border:1px solid #cbd5e1;border-radius:8px;background:#fff;">
                <option value="all">不限</option>
                <option value="7d">近7天</option>
                <option value="30d">近30天</option>
                <option value="365d">近365天</option>
              </select>
            </div>
          </div>
          <div style="display:grid;grid-template-columns:2fr 1fr;gap:12px;align-items:end;margin-top:12px;">
            <div>
              <div style="font-size:12px;color:#475569;margin-bottom:6px;">标签</div>
              <input id="lms-matrix-tags" type="text" placeholder="全部标签" value="${escapeHTML(state.tags)}" style="width:100%;padding:8px 10px;border:1px solid #cbd5e1;border-radius:8px;background:#fff;"/>
            </div>
            <div>
              <div style="font-size:12px;color:#475569;margin-bottom:6px;">排序方式</div>
              <select id="lms-matrix-sort" style="width:100%;padding:8px 10px;border:1px solid #cbd5e1;border-radius:8px;background:#fff;">
                <option value="added:desc">最新导入</option>
                <option value="updated:desc">最新修改</option>
                <option value="year:desc">年份（新→旧）</option>
                <option value="year:asc">年份（旧→新）</option>
                <option value="title:asc">标题（A→Z）</option>
                <option value="title:desc">标题（Z→A）</option>
              </select>
            </div>
          </div>
        </div>
      </div>
      <div id="lms-matrix-stats-wrap" style="padding:8px 12px;background:#fff;border-bottom:1px solid #eee;"></div>
      <div id="lms-matrix-table-wrap" style="flex:1 1 auto;overflow:auto;padding:8px 12px 12px 12px;background:#fff;"></div>
    </div>
  `;
  const statusSel = doc.getElementById(
    "lms-matrix-status",
  ) as HTMLSelectElement | null;
  if (statusSel) {
    statusSel.value = state.status;
  }
  const yearSel = doc.getElementById(
    "lms-matrix-year",
  ) as HTMLSelectElement | null;
  if (yearSel) {
    yearSel.value = state.year;
  }
  const journalSel = doc.getElementById(
    "lms-matrix-journal",
  ) as HTMLSelectElement | null;
  if (journalSel) {
    journalSel.value = state.journal;
  }
  const updatedSel = doc.getElementById(
    "lms-matrix-updated-range",
  ) as HTMLSelectElement | null;
  if (updatedSel) {
    updatedSel.value = state.updatedRange;
  }
  const sortSel = doc.getElementById(
    "lms-matrix-sort",
  ) as HTMLSelectElement | null;
  if (sortSel) {
    sortSel.value = `${state.sortKey}:${state.sortDir}`;
  }
  const renderTable = () => {
    state.q =
      (doc.getElementById("lms-matrix-q") as HTMLInputElement | null)?.value ||
      "";
    state.status =
      (doc.getElementById("lms-matrix-status") as HTMLSelectElement | null)
        ?.value || "all";
    state.year =
      (doc.getElementById("lms-matrix-year") as HTMLSelectElement | null)
        ?.value || "all";
    state.journal =
      (doc.getElementById("lms-matrix-journal") as HTMLSelectElement | null)
        ?.value || "all";
    state.updatedRange =
      (
        doc.getElementById(
          "lms-matrix-updated-range",
        ) as HTMLSelectElement | null
      )?.value || "all";
    state.tags =
      (doc.getElementById("lms-matrix-tags") as HTMLInputElement | null)
        ?.value || "";
    const sortValue =
      (doc.getElementById("lms-matrix-sort") as HTMLSelectElement | null)
        ?.value || `${state.sortKey}:${state.sortDir}`;
    const [nextSortKey, nextSortDir] = sortValue.split(":");
    if (nextSortKey && (nextSortDir === "asc" || nextSortDir === "desc")) {
      state.sortKey = nextSortKey;
      state.sortDir = nextSortDir;
    }
    const tableWrap = doc.getElementById(
      "lms-matrix-table-wrap",
    ) as HTMLDivElement | null;
    const statsWrap = doc.getElementById(
      "lms-matrix-stats-wrap",
    ) as HTMLDivElement | null;
    if (!tableWrap) {
      return;
    }
    const filtered = filterMatrixRows(
      rows,
      state.q,
      state.status,
      state.year,
      state.journal,
      state.updatedRange,
      state.tags,
    );
    const sorted = sortMatrixRows(filtered, state.sortKey, state.sortDir);
    if (statsWrap) {
      statsWrap.innerHTML = renderMatrixStatsHTML(sorted);
    }
    bindHeatmapTooltip(doc);
    tableWrap.innerHTML = renderMatrixTableHTML(
      sorted,
      state.sortKey,
      state.sortDir,
    );

    const sortBtns =
      tableWrap.querySelectorAll<HTMLButtonElement>("[data-sort-key]");
    sortBtns.forEach((btn: HTMLButtonElement) => {
      btn.onclick = () => {
        const key = String(btn.dataset.sortKey || "");
        if (!key) {
          return;
        }
        if (state.sortKey === key) {
          state.sortDir = state.sortDir === "asc" ? "desc" : "asc";
        } else {
          state.sortKey = key;
          state.sortDir = "asc";
        }
        renderTable();
      };
    });
    const openPdfTargets: NodeListOf<HTMLElement> =
      tableWrap.querySelectorAll<HTMLElement>("[data-open-pdf-item-id]");
    openPdfTargets.forEach((el: HTMLElement) => {
      el.onclick = async (ev: Event) => {
        ev.preventDefault?.();
        ev.stopPropagation?.();
        const id = Number(el.dataset.openPdfItemId);
        if (!id) {
          return;
        }
        await openPdfForItem(id);
      };
    });

    const jumpTargets: NodeListOf<HTMLElement> =
      tableWrap.querySelectorAll<HTMLElement>("[data-jump-item-id]");
    jumpTargets.forEach((el: HTMLElement) => {
      el.onclick = (ev: Event) => {
        ev.preventDefault?.();
        ev.stopPropagation?.();
        const id = Number(el.dataset.jumpItemId);
        if (!id) {
          return;
        }
        ztoolkit.getGlobal("Zotero_Tabs").select("zotero-pane");
        ztoolkit.getGlobal("ZoteroPane").selectItem(id);
      };
    });
    const exportCsvBtn = doc.getElementById(
      "lms-matrix-export-csv",
    ) as HTMLButtonElement | null;
    const exportMdBtn = doc.getElementById(
      "lms-matrix-export-md",
    ) as HTMLButtonElement | null;
    if (exportCsvBtn) {
      exportCsvBtn.onclick = () => exportMatrixCSV(filtered);
    }
    if (exportMdBtn) {
      exportMdBtn.onclick = () => exportMatrixMarkdown(filtered);
    }

    const sortSelect = doc.getElementById(
      "lms-matrix-sort",
    ) as HTMLSelectElement | null;
    if (sortSelect) {
      sortSelect.value = `${state.sortKey}:${state.sortDir}`;
    }
  };
  doc.getElementById("lms-matrix-q")?.addEventListener("input", renderTable);
  doc
    .getElementById("lms-matrix-status")
    ?.addEventListener("change", renderTable);
  doc
    .getElementById("lms-matrix-year")
    ?.addEventListener("change", renderTable);
  doc
    .getElementById("lms-matrix-journal")
    ?.addEventListener("change", renderTable);
  doc
    .getElementById("lms-matrix-updated-range")
    ?.addEventListener("change", renderTable);
  doc.getElementById("lms-matrix-tags")?.addEventListener("input", renderTable);
  doc
    .getElementById("lms-matrix-sort")
    ?.addEventListener("change", renderTable);
  doc.getElementById("lms-matrix-refresh")?.addEventListener("click", () => {
    void renderMatrixPage(win);
  });
  renderTable();
}

function parseISODateMs(isoDate: string): number {
  const iso = String(isoDate || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) {
    return 0;
  }
  const t = Date.parse(`${iso}T00:00:00Z`);
  return Number.isFinite(t) ? t : 0;
}

function sortMatrixRows(
  rows: MatrixPageRow[],
  sortKey: string,
  sortDir: "asc" | "desc",
): MatrixPageRow[] {
  const dir = sortDir === "asc" ? 1 : -1;
  const statusWeight = (s: string) =>
    s === "done" ? 3 : s === "reading" ? 2 : s === "unread" ? 1 : 0;
  const parseYear = (y: string) => {
    const n = Number(String(y || "").trim());
    return Number.isFinite(n) ? n : -1;
  };
  const parseISODate = (d: string) => parseISODateMs(d) || -1;
  const getAIField = (key: string) =>
    key.startsWith("ai__") ? fromSafeAIKey(key.slice(4)) : "";

  const comparator = (a: MatrixPageRow, b: MatrixPageRow) => {
    if (sortKey === "title") {
      return (
        dir *
        String(a.title || "").localeCompare(String(b.title || ""), "zh-CN", {
          numeric: true,
          sensitivity: "base",
        })
      );
    }
    if (sortKey === "author") {
      return (
        dir *
        String(a.author || "").localeCompare(String(b.author || ""), "zh-CN", {
          numeric: true,
          sensitivity: "base",
        })
      );
    }
    if (sortKey === "journal") {
      return (
        dir *
        String(a.journal || "").localeCompare(
          String(b.journal || ""),
          "zh-CN",
          {
            numeric: true,
            sensitivity: "base",
          },
        )
      );
    }
    if (sortKey === "year") {
      return dir * (parseYear(a.year) - parseYear(b.year));
    }
    if (sortKey === "status") {
      return dir * (statusWeight(a.status) - statusWeight(b.status));
    }
    if (sortKey === "tags") {
      return (
        dir *
        String(a.tags || "").localeCompare(String(b.tags || ""), "zh-CN", {
          numeric: true,
          sensitivity: "base",
        })
      );
    }
    if (sortKey === "updated") {
      return dir * (parseISODate(a.updated) - parseISODate(b.updated));
    }
    if (sortKey === "added") {
      return dir * (parseISODate(a.added) - parseISODate(b.added));
    }
    const aiField = getAIField(sortKey);
    if (aiField) {
      return (
        dir *
        String(a.ai[aiField] || "").localeCompare(
          String(b.ai[aiField] || ""),
          "zh-CN",
          {
            numeric: true,
            sensitivity: "base",
          },
        )
      );
    }
    return dir * (a.itemID - b.itemID);
  };

  return [...rows].sort((a, b) => {
    const v = comparator(a, b);
    return v !== 0 ? v : a.itemID - b.itemID;
  });
}

function fromSafeAIKey(safe: string): MatrixAIKey | "" {
  const hit = AI_FIELDS.find((f) => toSafeKey(f) === safe);
  return hit || "";
}

async function buildMatrixPageRows(): Promise<MatrixPageRow[]> {
  const items = await getAllRegularItems();
  const rows: MatrixPageRow[] = [];
  for (const item of items) {
    const payload = getMatrixPayload(item);
    const itemTitle = getPreferredItemTitle(item);
    rows.push({
      itemID: item.id,
      title: itemTitle,
      author: String(item.firstCreator || "").trim(),
      journal: String(item.getField("publicationTitle") || ""),
      year: extractYear(String(item.getField("date") || "")),
      status: payload?.status || "unread",
      tags: (item.getTags() || [])
        .map((t) => String(t.tag || "").trim())
        .filter(Boolean)
        .join(", "),
      added: String(
        (item as any).dateAdded || (item as any).dateAddedUTC || "",
      ).slice(0, 10),
      updated: String(item.dateModified || "").slice(0, 10),
      activityDate: String(
        payload?.sourceNoteDateModified || item.dateModified || "",
      ).slice(0, 10),
      hasPDF: hasPdfAttachment(item),
      ai:
        payload?.ai ||
        (Object.fromEntries(AI_FIELDS.map((k) => [k, ""])) as Record<
          MatrixAIKey,
          string
        >),
    });
  }
  return rows;
}

function filterMatrixRows(
  rows: MatrixPageRow[],
  q: string,
  status: string,
  year: string,
  journal: string,
  updatedRange: string,
  tags: string,
) {
  const keyword = q.trim().toLowerCase();
  const tagsQuery = tags.trim().toLowerCase();
  const now = Date.now();
  const days =
    updatedRange === "7d"
      ? 7
      : updatedRange === "30d"
        ? 30
        : updatedRange === "365d"
          ? 365
          : 0;
  const threshold = days ? now - days * 24 * 60 * 60 * 1000 : 0;

  return rows.filter((row) => {
    if (!(status === "all" || row.status === status)) {
      return false;
    }
    if (!(year === "all" || String(row.year || "").trim() === year)) {
      return false;
    }
    const rowJournal = normalizeVisibleText(row.journal) || "";
    if (journal !== "all") {
      if (journal === "__empty__") {
        if (rowJournal) {
          return false;
        }
      } else if (rowJournal !== journal) {
        return false;
      }
    }
    if (threshold) {
      const t = parseISODateMs(row.updated);
      if (!t || t < threshold) {
        return false;
      }
    }
    if (tagsQuery) {
      const hay = String(row.tags || "").toLowerCase();
      if (!hay.includes(tagsQuery)) {
        return false;
      }
    }
    if (!keyword) {
      return true;
    }
    const text = [
      row.title,
      row.author,
      row.journal,
      row.tags,
      ...AI_FIELDS.map((f) => row.ai[f] || ""),
    ]
      .join(" ")
      .toLowerCase();
    return text.includes(keyword);
  });
}

function renderMatrixTableHTML(
  rows: MatrixPageRow[],
  sortKey: string,
  sortDir: "asc" | "desc",
) {
  const headers: Array<{ key: string; label: string }> = [
    { key: "title", label: "标题" },
    { key: "author", label: "作者" },
    { key: "journal", label: "期刊" },
    { key: "year", label: "年份" },
    { key: "status", label: "状态" },
    { key: "tags", label: "标签" },
    { key: "updated", label: "更新日期" },
    ...AI_FIELDS.map((f) => ({ key: `ai__${toSafeKey(f)}`, label: f })),
  ];
  const statusCN = (s: string) =>
    s === "done" ? "已读" : s === "reading" ? "在读" : "未读";
  const body = rows
    .map((row) => {
      const displayTitle = getGuaranteedTitleByItemID(row.itemID);
      const titleHTML = row.hasPDF
        ? `<a href="#" data-open-pdf-item-id="${row.itemID}" style="color:#0f172a;cursor:pointer;text-align:left;font-size:12px;line-height:1.4;white-space:normal;word-break:break-word;overflow-wrap:anywhere;text-decoration:underline;">${escapeHTML(displayTitle)}</a>`
        : `<a href="#" data-jump-item-id="${row.itemID}" title="无PDF附件，点击定位到条目" style="color:#0f172a;cursor:pointer;text-align:left;font-size:12px;line-height:1.4;white-space:normal;word-break:break-word;overflow-wrap:anywhere;text-decoration:underline;">${escapeHTML(displayTitle)}</a>`;
      const aiCells = AI_FIELDS.map(
        (field) =>
          `<td style="border:1px solid #e5e7eb;padding:6px 8px;vertical-align:top;white-space:normal;word-break:break-word;overflow-wrap:anywhere;max-width:320px;">${renderExpandableText(
            row.ai[field] || "",
            120,
          )}</td>`,
      ).join("");
      return `
      <tr>
        <td style="border:1px solid #e5e7eb;padding:6px 8px;white-space:normal;word-break:break-word;overflow-wrap:anywhere;max-width:360px;">
          ${titleHTML}
          <button data-jump-item-id="${row.itemID}" style="margin-top:4px;padding:0;border:none;background:none;color:#0f766e;cursor:pointer;text-align:left;font-size:11px;">定位到条目</button>
          <div style="margin-top:2px;font-size:10px;color:#64748b;">ID:${row.itemID}</div>
        </td>
        <td style="border:1px solid #e5e7eb;padding:6px 8px;white-space:normal;word-break:break-word;overflow-wrap:anywhere;max-width:140px;">${escapeHTML(row.author)}</td>
        <td style="border:1px solid #e5e7eb;padding:6px 8px;white-space:normal;word-break:break-word;overflow-wrap:anywhere;max-width:180px;">${escapeHTML(row.journal)}</td>
        <td style="border:1px solid #e5e7eb;padding:6px 8px;width:72px;">${escapeHTML(row.year)}</td>
        <td style="border:1px solid #e5e7eb;padding:6px 8px;width:72px;">${statusCN(row.status)}</td>
        <td style="border:1px solid #e5e7eb;padding:6px 8px;white-space:normal;word-break:break-word;overflow-wrap:anywhere;max-width:220px;">${escapeHTML(row.tags)}</td>
        <td style="border:1px solid #e5e7eb;padding:6px 8px;width:96px;">${escapeHTML(row.updated)}</td>
        ${aiCells}
      </tr>
    `;
    })
    .join("");
  return `
    <div style="margin-bottom:8px;font-size:12px;color:#64748b;">筛选后文献: ${rows.length}</div>
    <table style="border-collapse:collapse;table-layout:fixed;width:100%;font-size:12px;">
      <thead>
        <tr>
          ${headers
            .map(({ key, label }) => {
              const active = key === sortKey;
              const arrow = active ? (sortDir === "asc" ? " ▲" : " ▼") : "";
              return `<th style="position:sticky;top:0;background:#0f172a;color:#fff;border:1px solid #1e293b;padding:6px 8px;text-align:left;z-index:2;white-space:normal;word-break:break-word;overflow-wrap:anywhere;">
                <button data-sort-key="${escapeHTML(key)}" style="all:unset;cursor:pointer;display:inline;">
                  ${escapeHTML(label)}${arrow}
                </button>
              </th>`;
            })
            .join("")}
        </tr>
      </thead>
      <tbody>${body}</tbody>
    </table>
  `;
}

function renderExpandableText(input: string, limit: number) {
  const content = String(input || "").trim();
  const safeContent = escapeHTML(content);
  if (!content) {
    return "";
  }
  if (content.length <= limit) {
    return safeContent;
  }
  const short = `${escapeHTML(content.slice(0, limit))}...`;
  return `
    <details style="cursor:pointer;">
      <summary style="list-style-position:inside;color:#0f766e;">${short}</summary>
      <div style="margin-top:4px;color:#334155;white-space:normal;word-break:break-word;overflow-wrap:anywhere;">${safeContent}</div>
    </details>
  `;
}

function renderMatrixStatsHTML(rows: MatrixPageRow[]) {
  const done = rows.filter((r) => r.status === "done").length;
  const reading = rows.filter((r) => r.status === "reading").length;
  const unread = rows.filter((r) => r.status === "unread").length;
  const yearMap = new Map<string, number>();
  const journalMap = new Map<string, number>();
  rows.forEach((row) => {
    const year = row.year || "未知";
    yearMap.set(year, (yearMap.get(year) || 0) + 1);
    const journal = row.journal || "未分类";
    journalMap.set(journal, (journalMap.get(journal) || 0) + 1);
  });
  const topYears = [...yearMap.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 6);
  const topJournals = [...journalMap.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 6);
  const listHTML = (pairs: Array<[string, number]>) =>
    pairs
      .map(
        ([name, count]) =>
          `<div style="display:flex;justify-content:space-between;gap:8px;font-size:12px;line-height:1.4;">
            <span style="color:#334155;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${escapeHTML(name)}</span>
            <span style="color:#0f172a;font-weight:600;">${count}</span>
          </div>`,
      )
      .join("");
  return `
    <div style="display:grid;grid-template-columns:repeat(6,minmax(120px,1fr));gap:8px;">
      <div style="border:1px solid #e2e8f0;border-radius:8px;padding:8px;background:#f8fafc;">
        <div style="font-size:12px;color:#64748b;">筛选后总数</div>
        <div style="font-size:20px;font-weight:700;color:#0f172a;">${rows.length}</div>
      </div>
      <div style="border:1px solid #e2e8f0;border-radius:8px;padding:8px;background:#ecfeff;">
        <div style="font-size:12px;color:#155e75;">已读</div>
        <div style="font-size:20px;font-weight:700;color:#0f766e;">${done}</div>
      </div>
      <div style="border:1px solid #e2e8f0;border-radius:8px;padding:8px;background:#eff6ff;">
        <div style="font-size:12px;color:#1e40af;">在读</div>
        <div style="font-size:20px;font-weight:700;color:#1d4ed8;">${reading}</div>
      </div>
      <div style="border:1px solid #e2e8f0;border-radius:8px;padding:8px;background:#f8fafc;">
        <div style="font-size:12px;color:#475569;">未读</div>
        <div style="font-size:20px;font-weight:700;color:#334155;">${unread}</div>
      </div>
      <div style="border:1px solid #e2e8f0;border-radius:8px;padding:8px;background:#fff;">
        <div style="font-size:12px;color:#64748b;margin-bottom:4px;">Top 年份</div>
        ${listHTML(topYears)}
      </div>
      <div style="border:1px solid #e2e8f0;border-radius:8px;padding:8px;background:#fff;">
        <div style="font-size:12px;color:#64748b;margin-bottom:4px;">Top 期刊</div>
        ${listHTML(topJournals)}
      </div>
    </div>
    <div style="margin-top:8px;border:1px solid #e2e8f0;border-radius:8px;padding:10px;background:#fff;">
      <div style="font-size:12px;color:#64748b;margin-bottom:6px;">每日使用频率（最近365天）</div>
      ${renderUsageHeatmapHTML(rows)}
    </div>
  `;
}

function renderUsageHeatmapHTML(rows: MatrixPageRow[]) {
  const counts = new Map<string, number>();
  for (const row of rows) {
    const d = row.activityDate;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) {
      continue;
    }
    counts.set(d, (counts.get(d) || 0) + 1);
  }

  const year = new Date().getFullYear();
  const baseStart = new Date(year, 0, 1);
  const baseEnd = new Date(year, 11, 31);
  const start = new Date(baseStart);
  const startOffset = (start.getDay() + 6) % 7; // Monday = 0
  start.setDate(start.getDate() - startOffset);
  const end = new Date(baseEnd);
  const endOffset = 6 - ((end.getDay() + 6) % 7);
  end.setDate(end.getDate() + endOffset);

  let maxCount = 0;
  for (const v of counts.values()) {
    if (v > maxCount) {
      maxCount = v;
    }
  }

  const colorFor = (n: number, inRange: boolean) => {
    if (!inRange || n <= 0) return "#161b22";
    if (maxCount <= 1) return "#26a641";
    const ratio = n / maxCount;
    if (ratio < 0.25) return "#0e4429";
    if (ratio < 0.5) return "#006d32";
    if (ratio < 0.75) return "#26a641";
    return "#39d353";
  };

  const dateLabel = (d: Date) => {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  };

  const monthNames = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];

  type WeekCell = {
    date: Date;
    inRange: boolean;
    iso: string;
    count: number;
  };
  const weeks: WeekCell[][] = [];
  const monthLabels: Array<{ index: number; name: string }> = [];

  const dayMs = 24 * 60 * 60 * 1000;
  const totalDays = Math.round((end.getTime() - start.getTime()) / dayMs) + 1;
  const weeksCount = Math.ceil(totalDays / 7);
  for (let col = 0; col < weeksCount; col++) {
    const week: WeekCell[] = [];
    for (let row = 0; row < 7; row++) {
      const d = new Date(start);
      d.setDate(start.getDate() + col * 7 + row);
      const iso = dateLabel(d);
      const inRange = d >= baseStart && d <= baseEnd;
      week.push({
        date: d,
        inRange,
        iso,
        count: inRange ? counts.get(iso) || 0 : 0,
      });
    }
    weeks.push(week);

    const monthEntry = week.find((c) => c.date.getDate() === 1);
    if (monthEntry) {
      monthLabels.push({
        index: col,
        name: monthNames[monthEntry.date.getMonth()],
      });
    }
  }

  const cellGap = 3;
  const cells: string[] = [];
  for (const week of weeks) {
    for (const cell of week) {
      const color = colorFor(cell.count, cell.inRange);
      const borderColor = cell.inRange ? "#30363d" : "#21262d";
      cells.push(
        `<div class="lms-heat-cell" data-tip="${cell.iso} · ${cell.count} 次" style="border-radius:3px;background:${color};border:1px solid ${borderColor};"></div>`,
      );
    }
  }

  const monthRow = `
    <div class="lms-heat-month-row" style="display:grid;grid-template-columns:30px repeat(${weeksCount}, var(--lms-cell-size));column-gap:var(--lms-cell-gap);align-items:end;margin-bottom:6px;">
      <div></div>
      ${Array.from({ length: weeksCount })
        .map((_, idx) => {
          const label = monthLabels.find((m) => m.index === idx)?.name || "";
          return `<div style="font-size:11px;color:#8b949e;line-height:1;text-align:left;white-space:nowrap;overflow:visible;">${label}</div>`;
        })
        .join("")}
    </div>
  `;

  const weekLabels = `
    <div class="lms-heat-week-labels" style="display:grid;grid-template-rows:repeat(7,var(--lms-cell-size));row-gap:var(--lms-cell-gap);width:30px;">
      <div></div>
      <div style="font-size:11px;color:#8b949e;line-height:var(--lms-cell-size);">Mon</div>
      <div></div>
      <div style="font-size:11px;color:#8b949e;line-height:var(--lms-cell-size);">Wed</div>
      <div></div>
      <div style="font-size:11px;color:#8b949e;line-height:var(--lms-cell-size);">Fri</div>
      <div></div>
    </div>
  `;

  return `
    <div id="lms-heatmap-panel" data-weeks="${weeksCount}" style="--lms-cell-size:13px;--lms-cell-gap:${cellGap}px;background:#0d1117;border:1px solid #30363d;border-radius:8px;padding:10px 12px;">
      ${monthRow}
      <div style="display:flex;align-items:flex-start;gap:8px;overflow-x:auto;padding-bottom:2px;">
        ${weekLabels}
        <div style="position:relative;">
          <div id="lms-heatmap-tooltip" style="display:none;position:absolute;z-index:20;pointer-events:none;background:#111827;color:#e5e7eb;border:1px solid #374151;border-radius:6px;padding:4px 7px;font-size:11px;line-height:1.2;white-space:nowrap;box-shadow:0 4px 10px rgba(0,0,0,0.35);"></div>
          <div class="lms-heat-grid" style="display:grid;grid-auto-flow:column;grid-template-rows:repeat(7,var(--lms-cell-size));grid-auto-columns:var(--lms-cell-size);gap:var(--lms-cell-gap);">
            ${cells.join("")}
          </div>
        </div>
      </div>
      <div style="display:flex;justify-content:flex-end;align-items:center;gap:6px;margin-top:8px;">
        <span style="font-size:11px;color:#8b949e;">Less</span>
        <span style="width:10px;height:10px;border-radius:2px;background:#161b22;border:1px solid #30363d;display:inline-block;"></span>
        <span style="width:10px;height:10px;border-radius:2px;background:#0e4429;border:1px solid #30363d;display:inline-block;"></span>
        <span style="width:10px;height:10px;border-radius:2px;background:#006d32;border:1px solid #30363d;display:inline-block;"></span>
        <span style="width:10px;height:10px;border-radius:2px;background:#26a641;border:1px solid #30363d;display:inline-block;"></span>
        <span style="width:10px;height:10px;border-radius:2px;background:#39d353;border:1px solid #30363d;display:inline-block;"></span>
        <span style="font-size:11px;color:#8b949e;">More</span>
      </div>
    </div>
  `;
}

function bindHeatmapTooltip(doc: Document) {
  const panel = doc.getElementById(
    "lms-heatmap-panel",
  ) as HTMLDivElement | null;
  const tooltip = doc.getElementById(
    "lms-heatmap-tooltip",
  ) as HTMLDivElement | null;
  const cells = doc.querySelectorAll<HTMLDivElement>(".lms-heat-cell");
  if (!panel || !tooltip || !cells.length) {
    return;
  }

  const applyResponsiveCellSize = () => {
    const weeks = Number(panel.dataset.weeks || "53");
    const gap = 3;
    const labelsWidth = 42;
    const horizontalPadding = 24;
    const usable = Math.max(
      400,
      panel.clientWidth - labelsWidth - horizontalPadding,
    );
    const size = Math.floor((usable - (weeks - 1) * gap) / weeks);
    const clamped = Math.max(12, Math.min(20, size));
    panel.style.setProperty("--lms-cell-size", `${clamped}px`);
    panel.style.setProperty("--lms-cell-gap", `${gap}px`);
  };
  applyResponsiveCellSize();
  const win = doc.defaultView as
    | (Window & { __lmsHeatmapResizeHandler?: () => void })
    | null;
  if (win) {
    if (win.__lmsHeatmapResizeHandler) {
      win.removeEventListener("resize", win.__lmsHeatmapResizeHandler);
    }
    win.__lmsHeatmapResizeHandler = applyResponsiveCellSize;
    win.addEventListener("resize", win.__lmsHeatmapResizeHandler);
  }

  const showTip = (ev: MouseEvent, text: string) => {
    tooltip.textContent = text;
    tooltip.style.display = "block";
    const containerRect = tooltip.parentElement?.getBoundingClientRect();
    if (!containerRect) {
      return;
    }
    const x = ev.clientX - containerRect.left + 12;
    const y = ev.clientY - containerRect.top + 12;
    tooltip.style.left = `${x}px`;
    tooltip.style.top = `${y}px`;
  };
  cells.forEach((cell: HTMLDivElement) => {
    const tip = cell.dataset.tip || "";
    cell.onmouseenter = (ev: Event) => showTip(ev as MouseEvent, tip);
    cell.onmousemove = (ev: Event) => showTip(ev as MouseEvent, tip);
    cell.onmouseleave = () => {
      tooltip.style.display = "none";
    };
  });
}

function exportMatrixCSV(rows: MatrixPageRow[]) {
  const headers = [
    "标题",
    "作者",
    "期刊",
    "年份",
    "状态",
    "标签",
    "更新日期",
    ...AI_FIELDS,
  ];
  const statusCN = (s: string) =>
    s === "done" ? "已读" : s === "reading" ? "在读" : "未读";
  const lines = [
    headers.join(","),
    ...rows.map((row) =>
      [
        normalizeVisibleText(row.title) || `[item:${row.itemID}]`,
        row.author,
        row.journal,
        row.year,
        statusCN(row.status),
        row.tags,
        row.updated,
        ...AI_FIELDS.map((field) => row.ai[field] || ""),
      ]
        .map((v) => `"${String(v).replace(/"/g, '""')}"`)
        .join(","),
    ),
  ].join("\n");
  downloadTextFile(
    lines,
    `matrix-${dateStamp()}.csv`,
    "text/csv;charset=utf-8;",
  );
}

function exportMatrixMarkdown(rows: MatrixPageRow[]) {
  const headers = [
    "标题",
    "作者",
    "期刊",
    "年份",
    "状态",
    "标签",
    "更新日期",
    ...AI_FIELDS,
  ];
  const statusCN = (s: string) =>
    s === "done" ? "已读" : s === "reading" ? "在读" : "未读";
  const md = [
    `# 智能文献矩阵导出 (${new Date().toLocaleString()})`,
    "",
    `| ${headers.join(" | ")} |`,
    `| ${headers.map(() => "---").join(" | ")} |`,
    ...rows
      .map((row) =>
        [
          normalizeVisibleText(row.title) || `[item:${row.itemID}]`,
          row.author,
          row.journal,
          row.year,
          statusCN(row.status),
          row.tags,
          row.updated,
          ...AI_FIELDS.map((field) => row.ai[field] || ""),
        ]
          .map((v) => String(v).replace(/\|/g, "\\|").replace(/\n/g, " "))
          .join(" | "),
      )
      .map((line) => `| ${line} |`),
  ].join("\n");
  downloadTextFile(
    md,
    `matrix-${dateStamp()}.md`,
    "text/markdown;charset=utf-8;",
  );
}

function downloadTextFile(content: string, filename: string, mimeType: string) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = Zotero.getMainWindow().document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function dateStamp() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  const hh = String(d.getHours()).padStart(2, "0");
  const mi = String(d.getMinutes()).padStart(2, "0");
  return `${yyyy}${mm}${dd}-${hh}${mi}`;
}

function getPreferredItemTitle(item: Zotero.Item): string {
  const primary = normalizeVisibleText(item.getField("title", false, true));
  if (primary) {
    return primary;
  }
  const display = normalizeVisibleText(item.getDisplayTitle());
  if (display) {
    return display;
  }
  const shortTitle = normalizeVisibleText(
    item.getField("shortTitle", false, true),
  );
  if (shortTitle) {
    return shortTitle;
  }
  const attachmentTitle = getFirstAttachmentTitle(item);
  if (attachmentTitle) {
    return attachmentTitle;
  }
  const key = String((item as any).key || "").trim();
  if (key) {
    return `[${key}]`;
  }
  return `[item:${item.id}]`;
}

function getGuaranteedTitleByItemID(itemID: number): string {
  const item = Zotero.Items.get(itemID);
  if (item?.isRegularItem?.()) {
    return getPreferredItemTitle(item);
  }
  return `[item:${itemID}]`;
}

async function openPdfForItem(itemID: number) {
  try {
    const item = Zotero.Items.get(itemID);
    if (!item) {
      return;
    }
    let pdfAttachment: Zotero.Item | undefined;
    if (item.isPDFAttachment?.()) {
      pdfAttachment = item;
    } else if (item.isRegularItem?.()) {
      const best = (await item.getBestAttachment()) as Zotero.Item | false;
      if (best && best.isPDFAttachment?.()) {
        pdfAttachment = best;
      }
    }
    if (pdfAttachment) {
      await ztoolkit
        .getGlobal("ZoteroPane")
        .viewPDF(pdfAttachment.id, {} as _ZoteroTypes.Reader.Location);
      return;
    }
  } catch (e) {
    ztoolkit.log("openPdfForItem failed", e);
  }
  ztoolkit.getGlobal("Zotero_Tabs").select("zotero-pane");
  ztoolkit.getGlobal("ZoteroPane").selectItem(itemID);
}

function getFirstAttachmentTitle(item: Zotero.Item): string {
  const attachmentIDs = item.getAttachments?.() || [];
  if (!attachmentIDs.length) {
    return "";
  }
  const firstAttachment = Zotero.Items.get(attachmentIDs[0]);
  if (!firstAttachment) {
    return "";
  }
  const attTitle = normalizeVisibleText(
    firstAttachment.getField("title", false, true),
  );
  if (attTitle) {
    return attTitle;
  }
  const fileName = normalizeVisibleText(firstAttachment.attachmentFilename);
  return fileName || "";
}

function normalizeVisibleText(value: unknown): string {
  const input = String(value || "");
  let cleaned = "";
  for (let i = 0; i < input.length; i++) {
    const code = input.charCodeAt(i);
    cleaned += code <= 0x1f || (code >= 0x7f && code <= 0x9f) ? " " : input[i];
  }
  return cleaned
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function escapeHTML(input: string) {
  return String(input || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
