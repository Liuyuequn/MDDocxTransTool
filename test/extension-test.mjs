// VS Code 插件层测试
//
// 分三段：
//   A. 命令层：用桩宿主驱动 extension/lib/commands.js，覆盖两个 docx 入口与转换/预设流程
//   B. 接线层：用假 vscode 模块真实加载 extension/extension.cjs，跑通 activate() 与命令注册
//   C. 清单层：校验 package.json 的 commands / menus / activationEvents / main 自洽
//
// 不依赖 VS Code 安装，可在 CI（Node 18+）中运行。

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import Module, { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

import { defaultOptions, mergeOptions } from "../src/options.js";
import { convertMarkdownText } from "../src/converter.js";
import { loadCustomPreset } from "../src/preset-extract.js";
import { createCommands, suggestPresetName } from "../extension/lib/commands.js";
import { listAllPresets, resolvePresetPatch, WORKSPACE_PRESET_DIR } from "../extension/lib/presets-store.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..");
const require = createRequire(import.meta.url);

const tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), "mddtt-ext-"));
let passed = 0;
let failed = 0;

function check(name, cond) {
  console.log(`${cond ? "通过" : "失败"} | ${name}`);
  if (cond) passed++;
  else failed++;
}

const read = (file) => fs.promises.readFile(file, "utf-8");
const exists = (file) => fs.existsSync(file);

/** 生成测试用源 docx：带页边距、正文字号、页眉文字与页脚页码 */
async function makeSourceDocx(target) {
  const md = ["# 测试标题", "", "这是正文段落，用于提取正文版式。", "", "- 列表项一", "- 列表项二", ""].join("\n");
  const opts = mergeOptions(defaultOptions, {
    page: { margin: { top: 3, right: 2.5, bottom: 3, left: 2.5, header: 1.5 } },
    sizes: { body: 12 },
    header: { text: "测试页眉" },
    pageNumber: { pos: "bottom", format: "第X页/共Y页" },
  });
  const buffer = await convertMarkdownText(md, opts, tmpRoot);
  await fs.promises.writeFile(target, buffer);
}

/** 桩宿主：记录所有交互与日志，供断言使用 */
function createStubHost({ userDir, workspaceDir = null, answers = {} } = {}) {
  const lines = [];
  const calls = { pickFile: [], prompts: [], confirms: [], setDefaultPreset: [], openExternal: [], openInEditor: [] };
  const queue = (key) => {
    const answer = answers[key];
    return Array.isArray(answer) ? answer.shift() : answer;
  };
  const host = {
    settings: {
      defaultPreset: answers.defaultPreset ?? "",
      presetStorage: answers.presetStorage ?? "user",
      confirmOverwrite: answers.confirmOverwrite ?? true,
      openAfterConvert: answers.openAfterConvert ?? false,
    },
    log: (m) => lines.push(`LOG ${m}`),
    warn: (m) => lines.push(`WARN ${m}`),
    error: (m) => lines.push(`ERROR ${m}`),
    async reportError(m) {
      lines.push(`DIALOG-ERROR ${m}`);
    },
    async info(m) {
      lines.push(`INFO ${m}`);
      return queue("infoAction");
    },
    async pickFile(ext) {
      calls.pickFile.push(ext);
      return queue("pickFile") ?? null;
    },
    async promptPresetName(suggested) {
      calls.prompts.push(suggested);
      return queue("presetName") ?? null;
    },
    async confirm(message) {
      calls.confirms.push(message);
      return queue("confirm") ?? false;
    },
    async pickFromList(placeHolder, items) {
      calls.listItems = items;
      calls.listPlaceHolder = placeHolder;
      return queue("listPick");
    },
    async setDefaultPreset(name) {
      calls.setDefaultPreset.push(name);
    },
    async withProgress(title, task) {
      return task();
    },
    readFile: (p) => fs.promises.readFile(p),
    readText: (p) => fs.promises.readFile(p, "utf-8"),
    writeFile: (p, d) => fs.promises.writeFile(p, d),
    async exists(p) {
      try {
        await fs.promises.access(p);
        return true;
      } catch {
        return false;
      }
    },
    async openInEditor(p) {
      calls.openInEditor.push(p);
    },
    async openExternal(p) {
      calls.openExternal.push(p);
    },
    showOutput() {
      calls.showOutput = true;
    },
    presetDirs: () => ({ userDir, workspaceDir, cliDir: path.join(tmpRoot, "cli-none") }),
  };
  return { host, lines, calls };
}

// ==================== A. 命令层 ====================
console.log("—— A. 命令层（桩宿主） ——");

const caseDir = path.join(tmpRoot, "caseA");
const userPresets = path.join(caseDir, "userPresets");
const workspacePresets = path.join(caseDir, "ws", WORKSPACE_PRESET_DIR);
fs.mkdirSync(caseDir, { recursive: true });
const sourceDocx = path.join(caseDir, "模板 文档.docx");
await makeSourceDocx(sourceDocx);

check("A1 建议预设名：去空格与后缀", suggestPresetName(sourceDocx) === "模板文档");
check("A2 建议预设名：非法字符回退为 preset", suggestPresetName(path.join(caseDir, "***.docx")) === "preset");

// ---- 入口 1：右键菜单（命令收到 docx 路径） ----
{
  const source = path.join(caseDir, "entry1.docx");
  fs.copyFileSync(sourceDocx, source);
  const { host, lines, calls } = createStubHost({ userDir: userPresets, answers: { presetName: "entry1预设" } });
  await createCommands(host).extractPreset(source);

  const file = path.join(userPresets, "entry1预设.json");
  check("A3 入口1：预设文件已生成", exists(file));
  check("A4 入口1：未弹出文件选择对话框", calls.pickFile.length === 0);
  check("A5 入口1：输入框建议名取自文件名", calls.prompts[0] === "entry1");
  const saved = exists(file) ? JSON.parse(fs.readFileSync(file, "utf-8")) : {};
  check("A6 入口1：记录了来源文档", saved.source === source);
  check("A7 入口1：提取到页边距", saved.options?.page?.margin?.top === 3);
  check("A8 入口1：提取到正文字号", saved.options?.sizes?.body === 12);
  check("A9 入口1：提取到页眉文字", String(saved.options?.header?.text ?? "").includes("测试页眉"));
  check("A10 入口1：提取到页码格式", saved.options?.pageNumber?.format === "第X页/共Y页");
  check("A11 入口1：版式摘要写入输出通道", lines.some((l) => l.startsWith("LOG   页面")));
  check("A12 入口1：提示复用方式", lines.some((l) => l.includes("mddtt.defaultPreset")));
  check("A13 入口1：通知已保存", lines.some((l) => l.startsWith("INFO 预设「entry1预设」已保存")));
  check("A14 入口1：预设可被加载复用", resolvePresetPatch("entry1预设", { userDir: userPresets }).patch?.page?.margin?.top === 3);
}

// ---- 入口 2：命令面板（无路径 → 文件选择对话框） ----
{
  const source = path.join(caseDir, "entry2.docx");
  fs.copyFileSync(sourceDocx, source);
  const { host, calls } = createStubHost({
    userDir: userPresets,
    answers: { pickFile: source, presetName: "entry2预设" },
  });
  await createCommands(host).extractPreset(undefined);

  check("A15 入口2：弹出文件选择对话框", calls.pickFile.length === 1 && calls.pickFile[0] === ".docx");
  check("A16 入口2：使用对话框返回的文档提取", exists(path.join(userPresets, "entry2预设.json")));
}

// ---- 入口 2：用户取消对话框 ----
{
  const { host, lines, calls } = createStubHost({ userDir: userPresets, answers: { pickFile: null } });
  await createCommands(host).extractPreset(undefined);
  check("A17 入口2取消：不写预设且记录取消", calls.prompts.length === 0 && lines.some((l) => l.includes("已取消")));
}

// ---- 扩展名校验 ----
{
  const mdFile = path.join(caseDir, "note.md");
  fs.writeFileSync(mdFile, "# hi\n", "utf-8");
  const { host, lines, calls } = createStubHost({ userDir: userPresets });
  await createCommands(host).extractPreset(mdFile);
  check("A18 非 docx：报错且不进入命名流程", calls.prompts.length === 0 && lines.some((l) => l.startsWith("DIALOG-ERROR")));
}

// ---- 预设名不合法 ----
{
  const source = path.join(caseDir, "badname.docx");
  fs.copyFileSync(sourceDocx, source);
  const { host, lines } = createStubHost({ userDir: userPresets, answers: { presetName: "非法 名字" } });
  await createCommands(host).extractPreset(source);
  check("A19 非法预设名：报错且不写文件", lines.some((l) => l.includes("不合法")));
}

// ---- 覆盖确认：取消 ----
{
  const source = path.join(caseDir, "overwrite.docx");
  fs.copyFileSync(sourceDocx, source);
  const existing = path.join(userPresets, "已存在.json");
  fs.writeFileSync(existing, JSON.stringify({ name: "已存在", marker: "原始" }), "utf-8");
  const { host, lines } = createStubHost({
    userDir: userPresets,
    answers: { presetName: "已存在", confirm: false },
  });
  await createCommands(host).extractPreset(source);
  const after = JSON.parse(fs.readFileSync(existing, "utf-8"));
  check("A20 覆盖取消：原预设未被改写", after.marker === "原始");
  check("A21 覆盖取消：记录了取消", lines.some((l) => l.includes("已存在")));
}

// ---- 覆盖确认：确认 ----
{
  const source = path.join(caseDir, "overwrite2.docx");
  fs.copyFileSync(sourceDocx, source);
  const existing = path.join(userPresets, "待覆盖.json");
  fs.writeFileSync(existing, JSON.stringify({ name: "待覆盖", marker: "原始" }), "utf-8");
  const { host } = createStubHost({
    userDir: userPresets,
    answers: { presetName: "待覆盖", confirm: true },
  });
  await createCommands(host).extractPreset(source);
  const after = JSON.parse(fs.readFileSync(existing, "utf-8"));
  check("A22 覆盖确认：预设已更新", after.marker === undefined && after.options?.page?.margin?.top === 3);
}

// ---- 工作区级存储 ----
{
  const source = path.join(caseDir, "ws.docx");
  fs.copyFileSync(sourceDocx, source);
  const { host } = createStubHost({
    userDir: userPresets,
    workspaceDir: workspacePresets,
    answers: { presetName: "工作区预设", presetStorage: "workspace" },
  });
  await createCommands(host).extractPreset(source);
  check("A23 工作区级存储：写入 .vscode/mddtt-presets/", exists(path.join(workspacePresets, "工作区预设.json")));
  check("A24 工作区级存储：未写入用户级目录", !exists(path.join(userPresets, "工作区预设.json")));
}

// ---- md → docx ----
{
  const mdFile = path.join(caseDir, "doc.md");
  fs.writeFileSync(mdFile, "# 标题\n\n正文段落。\n", "utf-8");
  const docxOut = path.join(caseDir, "doc.docx");
  const { host, lines, calls } = createStubHost({ userDir: userPresets, answers: { openAfterConvert: true, infoAction: "打开" } });
  await createCommands(host).convertMarkdownToDocx(mdFile);
  check("A25 md→docx：生成输出文件", exists(docxOut));
  const zip = exists(docxOut) ? await JSZip.loadAsync(fs.readFileSync(docxOut)) : null;
  check("A26 md→docx：产物是合法 docx 包", Boolean(zip?.file("word/document.xml")));
  check("A27 md→docx：按设置打开结果", calls.openExternal.length === 1);
  check("A28 md→docx：日志记录默认格式", lines.some((l) => l.includes("默认格式")));
}

// ---- md → docx：使用预设 ----
{
  const mdFile = path.join(caseDir, "doc-sundy.md");
  fs.writeFileSync(mdFile, "# 标题\n\n正文。\n", "utf-8");
  const { host, lines } = createStubHost({ userDir: userPresets, answers: { defaultPreset: "sundy" } });
  await createCommands(host).convertMarkdownToDocx(mdFile);
  const zip = await JSZip.loadAsync(fs.readFileSync(path.join(caseDir, "doc-sundy.docx")));
  const headerFiles = Object.keys(zip.files).filter((f) => /^word\/header\d+\.xml$/.test(f));
  check("A29 md→docx：sundy 预设产生页眉", headerFiles.length > 0);
  check("A30 md→docx：日志记录所用预设", lines.some((l) => l.includes("预设：sundy")));
}

// ---- md → docx：未知预设 ----
{
  const mdFile = path.join(caseDir, "doc-bad.md");
  fs.writeFileSync(mdFile, "# 标题\n", "utf-8");
  const { host, lines } = createStubHost({ userDir: userPresets, answers: { defaultPreset: "不存在的预设" } });
  await createCommands(host).convertMarkdownToDocx(mdFile);
  check("A31 未知预设：报错且不产出文件", !exists(path.join(caseDir, "doc-bad.docx")) && lines.some((l) => l.includes("未知预设")));
}

// ---- 覆盖确认拒绝：不转换 ----
{
  const mdFile = path.join(caseDir, "doc-denied.md");
  fs.writeFileSync(mdFile, "# 标题\n", "utf-8");
  const out = path.join(caseDir, "doc-denied.docx");
  fs.writeFileSync(out, "占位", "utf-8");
  const { host, lines } = createStubHost({ userDir: userPresets, answers: { confirm: false } });
  await createCommands(host).convertMarkdownToDocx(mdFile);
  check("A32 拒绝覆盖：原文件保持不变", fs.readFileSync(out, "utf-8") === "占位");
  check("A33 拒绝覆盖：记录了取消", lines.some((l) => l.includes("已取消")));
}

// ---- docx → md ----
{
  const { host, calls } = createStubHost({ userDir: userPresets, answers: { openAfterConvert: true, infoAction: "打开" } });
  await createCommands(host).convertDocxToMarkdown(sourceDocx);
  const mdOut = path.join(caseDir, "模板 文档.md");
  const text = exists(mdOut) ? await read(mdOut) : "";
  check("A34 docx→md：生成 Markdown", text.includes("# 测试标题"));
  check("A35 docx→md：保留列表", text.includes("列表项一"));
  check("A36 docx→md：按设置打开结果", calls.openInEditor.length === 1);
}

// ---- 选择默认预设 ----
{
  const { host, calls, lines } = createStubHost({ userDir: userPresets, answers: { listPick: "sundy" } });
  const commands = createCommands(host);
  await commands.selectDefaultPreset();
  check("A37 选择默认预设：写入配置", calls.setDefaultPreset[0] === "sundy");
  const labels = calls.listItems.map((i) => i.label);
  check("A38 选择默认预设：列表含内置 sundy", labels.includes("sundy"));
  check("A39 选择默认预设：列表含默认格式项", labels.includes("（默认格式）"));
  check("A40 选择默认预设：日志记录结果", lines.some((l) => l.includes("默认预设已设为：sundy")));
}

// ---- 预设列表与搜索优先级 ----
{
  const all = listAllPresets({ userDir: userPresets, workspaceDir: workspacePresets });
  check("A41 预设列表：内置 sundy 在最前", all[0]?.name === "sundy" && all[0]?.kind === "builtin");
  check("A42 预设列表：含工作区级预设", all.some((p) => p.name === "工作区预设" && p.kind === "workspace"));
  check("A43 预设列表：含用户级预设", all.some((p) => p.name === "entry1预设" && p.kind === "user"));
  const patch = loadCustomPreset("工作区预设", workspacePresets);
  check("A44 自定义预设可被核心读取", patch?.page?.margin?.top === 3);
}

// ==================== B. 接线层（假 vscode） ====================
console.log("—— B. 接线层（假 vscode 真实加载 extension.cjs） ——");

function createFakeVscode(state) {
  const registered = {};
  const infoMessages = [];
  const outputLines = [];
  const quickPickItems = [];
  const configuration = {
    get: (key, fallback) => (key in state.config ? state.config[key] : fallback),
    update: async (key, value, target) => {
      state.config[key] = value;
      state.updates.push({ key, value, target });
    },
  };
  const Uri = { file: (p) => ({ fsPath: p, scheme: "file", toString: () => `file://${p}` }) };
  const window = {
    // 必须用 getter：activeTextEditor 会在测试过程中被切换
    get activeTextEditor() {
      return state.activeEditor ? { document: { uri: Uri.file(state.activeEditor) } } : undefined;
    },
    createOutputChannel: (name, options) => {
      void name;
      void options;
      const push = (level) => (message) => outputLines.push(`${level} ${message}`);
      return {
        info: push("INFO"),
        warn: push("WARN"),
        error: push("ERROR"),
        debug: push("DEBUG"),
        trace: push("TRACE"),
        show() {},
        dispose() {},
      };
    },
    showOpenDialog: async () => state.openDialogResult,
    showInputBox: async () => state.inputBoxResult,
    showWarningMessage: async () => state.warningResult,
    showInformationMessage: async (message) => {
      infoMessages.push(message);
      return state.infoResult;
    },
    showErrorMessage: async (message) => {
      infoMessages.push(`ERROR ${message}`);
      return undefined;
    },
    showQuickPick: async (items) => {
      quickPickItems.push(items);
      return state.quickPickResult === undefined ? undefined : items[state.quickPickResult];
    },
    showTextDocument: async (document) => {
      state.openedEditors.push(document.uri.fsPath);
      return document;
    },
    withProgress: async (options, task) => {
      void options;
      return task();
    },
  };
  return {
    registered,
    infoMessages,
    outputLines,
    quickPickItems,
    Uri,
    ProgressLocation: { Notification: 15 },
    ConfigurationTarget: { Global: 1, Workspace: 2 },
    commands: {
      registerCommand: (id, handler) => {
        registered[id] = handler;
        return { dispose() {} };
      },
    },
    workspace: {
      workspaceFolders: state.workspaceFolder ? [{ uri: Uri.file(state.workspaceFolder) }] : undefined,
      getConfiguration: () => configuration,
      openTextDocument: async (uri) => ({ uri }),
    },
    window,
    env: {
      openExternal: async (uri) => {
        state.externalOpened.push(uri.fsPath);
        return true;
      },
    },
  };
}

/** 用假 vscode 加载真实入口（拦截 require("vscode")） */
function loadExtensionWithFakeVscode(fakeVscode) {
  const original = Module._load;
  Module._load = function load(request, parent, isMain) {
    if (request === "vscode") return fakeVscode;
    return original.apply(this, arguments);
  };
  try {
    const entry = require.resolve("../extension/extension.cjs");
    delete require.cache[entry];
    return require("../extension/extension.cjs");
  } finally {
    Module._load = original;
  }
}

const hostCaseDir = path.join(tmpRoot, "caseB");
fs.mkdirSync(hostCaseDir, { recursive: true });
const hostDocx = path.join(hostCaseDir, "接线.docx");
await makeSourceDocx(hostDocx);

const state = {
  config: {},
  updates: [],
  workspaceFolder: path.join(hostCaseDir, "ws"),
  activeEditor: null,
  openDialogResult: undefined,
  inputBoxResult: null,
  warningResult: undefined,
  infoResult: undefined,
  quickPickResult: undefined,
  openedEditors: [],
  externalOpened: [],
};
const fakeVscode = createFakeVscode(state);
const extension = loadExtensionWithFakeVscode(fakeVscode);

const context = {
  subscriptions: [],
  globalStorageUri: fakeVscode.Uri.file(path.join(hostCaseDir, "globalStorage")),
  extensionUri: fakeVscode.Uri.file(root),
};
await extension.activate(context);

check("B1 activate 注册 5 个命令", Object.keys(fakeVscode.registered).length === 5);
check(
  "B2 注册的命令 ID 完整",
  ["mddtt.convertMarkdownToDocx", "mddtt.convertDocxToMarkdown", "mddtt.extractPreset", "mddtt.selectDefaultPreset", "mddtt.showOutput"].every(
    (id) => typeof fakeVscode.registered[id] === "function"
  )
);
check("B3 activate 写入输出通道", fakeVscode.outputLines.some((l) => l.includes("MDDTT 已激活")));
// 1 个输出通道 + 5 个命令注册
check("B4 activate 注册了订阅清理", context.subscriptions.length === 6);

// ---- 入口 1：右键 Uri 直达（真实接线） ----
{
  state.inputBoxResult = "接线预设";
  state.openDialogResult = undefined;
  await fakeVscode.registered["mddtt.extractPreset"](fakeVscode.Uri.file(hostDocx));
  const file = path.join(hostCaseDir, "globalStorage", "presets", "接线预设.json");
  check("B5 入口1（Uri）：预设写入 globalStorage", exists(file));
  check("B6 入口1（Uri）：Uri.fsPath 正确换算", exists(file) && JSON.parse(fs.readFileSync(file, "utf-8")).source === hostDocx);
  check("B7 入口1（Uri）：摘要写入输出通道", fakeVscode.outputLines.some((l) => l.startsWith("INFO   页面")));
}

// ---- 入口 2：命令面板（无参数 → showOpenDialog） ----
{
  state.inputBoxResult = "对话框预设";
  state.openDialogResult = [fakeVscode.Uri.file(hostDocx)];
  await fakeVscode.registered["mddtt.extractPreset"](undefined);
  check("B8 入口2（面板）：走文件选择对话框并保存", exists(path.join(hostCaseDir, "globalStorage", "presets", "对话框预设.json")));
}

// ---- 入口 2 取消 ----
{
  state.openDialogResult = undefined;
  state.inputBoxResult = "不应保存";
  await fakeVscode.registered["mddtt.extractPreset"](undefined);
  check("B9 入口2（面板）：取消不写文件", !exists(path.join(hostCaseDir, "globalStorage", "presets", "不应保存.json")));
}

// ---- 回退到活动编辑器 ----
{
  const activeDocx = path.join(hostCaseDir, "激活.docx");
  fs.copyFileSync(hostDocx, activeDocx);
  state.activeEditor = activeDocx;
  state.inputBoxResult = "活动编辑器预设";
  await fakeVscode.registered["mddtt.extractPreset"](undefined);
  check("B10 无参回退：使用活动编辑器的 docx", exists(path.join(hostCaseDir, "globalStorage", "presets", "活动编辑器预设.json")));
  state.activeEditor = null;
}

// ---- 转换命令接线 ----
{
  const mdFile = path.join(hostCaseDir, "接线.md");
  fs.writeFileSync(mdFile, "# 接线测试\n\n正文。\n", "utf-8");
  state.config.openAfterConvert = false;
  await fakeVscode.registered["mddtt.convertMarkdownToDocx"](fakeVscode.Uri.file(mdFile));
  check("B11 md→docx 接线：生成 docx", exists(path.join(hostCaseDir, "接线.docx")));

  const docx2md = path.join(hostCaseDir, "回转.docx");
  fs.copyFileSync(hostDocx, docx2md);
  await fakeVscode.registered["mddtt.convertDocxToMarkdown"](fakeVscode.Uri.file(docx2md));
  const mdOut = path.join(hostCaseDir, "回转.md");
  check("B12 docx→md 接线：生成 Markdown", exists(mdOut) && (await read(mdOut)).includes("测试标题"));
}

// ---- 选择默认预设接线（写入配置） ----
{
  state.quickPickResult = 1; // 列表下标 0 = （默认格式），1 = sundy
  state.config.defaultPreset = "";
  await fakeVscode.registered["mddtt.selectDefaultPreset"]();
  check("B13 选择默认预设接线：写入 mddtt.defaultPreset", state.updates.some((u) => u.key === "defaultPreset" && u.value === "sundy"));
  check("B14 选择默认预设接线：QuickPick 传出内置预设", fakeVscode.quickPickItems[0]?.some((i) => i.value === "sundy"));
}

// ---- 工作区级存储接线 ----
{
  state.config.presetStorage = "workspace";
  state.config.defaultPreset = "";
  state.inputBoxResult = "工作区接线预设";
  await fakeVscode.registered["mddtt.extractPreset"](fakeVscode.Uri.file(hostDocx));
  check(
    "B15 工作区级存储接线：写入工作区目录",
    exists(path.join(state.workspaceFolder, WORKSPACE_PRESET_DIR, "工作区接线预设.json"))
  );
}

// ==================== C. 清单层 ====================
console.log("—— C. 清单层（package.json 自洽性） ——");

const pkg = JSON.parse(await read(path.join(root, "package.json")));
const entrySource = await read(path.join(root, "extension", "extension.cjs"));
const contributes = pkg.contributes ?? {};
const commandIds = (contributes.commands ?? []).map((c) => c.command);
const menuCommandIds = Object.values(contributes.menus ?? {}).flat().map((m) => m.command);

check("C1 声明 main 且文件存在", Boolean(pkg.main) && exists(path.join(root, pkg.main)));
check("C2 声明 engines.vscode", Boolean(pkg.engines?.vscode));
check("C3 命令行入口保留", pkg.bin?.mddtt === "./src/cli.js");
check("C4 命令 ID 在入口中均有注册", commandIds.every((id) => entrySource.includes(`"${id}"`)));
check("C5 菜单引用的命令均已在 commands 中声明", menuCommandIds.every((id) => commandIds.includes(id)));
check("C6 activationEvents 覆盖全部命令", commandIds.every((id) => (pkg.activationEvents ?? []).includes(`onCommand:${id}`)));
check("C7 每个命令都有标题与分类", (contributes.commands ?? []).every((c) => c.title && c.category));
const explorerMenus = contributes.menus?.["explorer/context"] ?? [];
check(
  "C8 资源管理器右键：docx 上提供「转 Markdown」与「提取版式」",
  explorerMenus.some((m) => m.command === "mddtt.convertDocxToMarkdown" && m.when.includes(".docx")) &&
    explorerMenus.some((m) => m.command === "mddtt.extractPreset" && m.when.includes(".docx"))
);
check(
  "C9 资源管理器右键：md 上提供「转 docx」",
  explorerMenus.some((m) => m.command === "mddtt.convertMarkdownToDocx" && m.when.includes(".md"))
);
check("C10 配置项齐备", ["mddtt.defaultPreset", "mddtt.presetStorage", "mddtt.confirmOverwrite", "mddtt.openAfterConvert"].every((k) => contributes.configuration?.properties?.[k]));

// sundy 预设运行时依赖 reference/圣典律师.png，打包排除项不能误伤运行时目录
const ignoreRules = await read(path.join(root, ".vscodeignore"));
check("C11 vsce 排除项不误伤运行时文件", !/^\s*(src|extension|reference|node_modules)\//m.test(ignoreRules));
check("C12 vsce 排除项排除了测试目录", /^\s*test\/\*\*/m.test(ignoreRules));

// 未声明 untrustedWorkspaces 的插件会在工作区「受限模式」下被 VS Code 直接禁用
check("C13 声明工作区受限模式下可用", pkg.capabilities?.untrustedWorkspaces?.supported === true);
check("C14 声明不支持虚拟工作区（需真实文件系统）", pkg.capabilities?.virtualWorkspaces === false);

// ==================== 汇总 ====================
fs.rmSync(tmpRoot, { recursive: true, force: true });
console.log(`\n插件层测试：通过 ${passed}，失败 ${failed}`);
if (failed > 0) process.exit(1);
