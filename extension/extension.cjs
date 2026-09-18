"use strict";
// MDDTT VS Code 插件入口（CommonJS）
//
// 为什么是 .cjs：VS Code 以 require() 加载 main 入口，而本项目根 package.json 是
// "type": "module"（命令行与 src/ 均为 ESM）。用 .cjs 后缀让 Node 明确按 CommonJS
// 加载本文件，再用动态 import() 载入 ESM 的 src/ 与 extension/lib/，无需打包器。
//
// 命令入口约定：菜单调用时 VS Code 传入资源 Uri；命令面板调用时不传参数，
// 此时回退到「当前活动编辑器」再回退到文件选择对话框。

const vscode = require("vscode");
const fs = require("node:fs");
const path = require("node:path");

const MD_EXT = ".md";
const DOCX_EXT = ".docx";

/** LogOutputChannel：写入「输出 → MDDTT」面板 */
let output = null;

/** 支持本地与 Remote 场景；其余虚拟文件系统交由文件选择对话框兜底 */
const SUPPORTED_SCHEMES = new Set(["file", "vscode-remote"]);

function activate(context) {
  return (async () => {
    const [{ createCommands }, store, { validPresetName }] = await Promise.all([
      import("./lib/commands.js"),
      import("./lib/presets-store.js"),
      import("../src/preset-extract.js"),
    ]);

    output = vscode.window.createOutputChannel("MDDTT", { log: true });
    context.subscriptions.push(output);

    const host = createHost(context, store, validPresetName);
    const commands = createCommands(host);

    const register = (id, handler) => {
      context.subscriptions.push(vscode.commands.registerCommand(id, handler));
    };
    // 入口 1：资源管理器 / 编辑器标签右键（VS Code 传入 Uri）
    // 入口 2：命令面板（无参数 → handler 收到 undefined → 弹出文件选择对话框）
    register("mddtt.convertMarkdownToDocx", (uri) => commands.convertMarkdownToDocx(inputPath(uri, MD_EXT)));
    register("mddtt.convertDocxToMarkdown", (uri) => commands.convertDocxToMarkdown(inputPath(uri, DOCX_EXT)));
    register("mddtt.extractPreset", (uri) => commands.extractPreset(inputPath(uri, DOCX_EXT)));
    register("mddtt.selectDefaultPreset", () => commands.selectDefaultPreset());
    register("mddtt.showOutput", () => commands.showOutput());

    const dirs = host.presetDirs();
    output.info("MDDTT 已激活：Markdown ↔ docx 双向转换、docx 版式提取");
    output.info(`用户级预设目录：${dirs.userDir}`);
    if (dirs.workspaceDir) output.info(`工作区级预设目录：${dirs.workspaceDir}`);
    if (fs.existsSync(dirs.cliDir)) output.info(`命令行预设目录（只读）：${dirs.cliDir}`);
  })();
}

function deactivate() {
  output = null;
}

/**
 * 把命令入参（VS Code Uri / Uri 数组）换算为文件系统路径。
 * 拿不到可用路径时回退到「当前活动编辑器」中扩展名匹配的文件；
 * 仍拿不到时返回 undefined，由命令弹出文件选择对话框。
 */
function inputPath(uri, expectedExt) {
  const direct = toFsPath(uri);
  if (direct) return direct;
  const activePath = toFsPath(vscode.window.activeTextEditor?.document?.uri);
  if (activePath && activePath.toLowerCase().endsWith(expectedExt)) return activePath;
  return undefined;
}

function toFsPath(uri) {
  if (!uri) return undefined;
  if (Array.isArray(uri)) return toFsPath(uri[0]);
  if (typeof uri === "string") return uri;
  if (!uri.fsPath) return undefined;
  if (uri.scheme && !SUPPORTED_SCHEMES.has(uri.scheme)) {
    output?.warn(`暂不支持虚拟文件系统（${uri.scheme}），请改用文件选择对话框：${uri.toString()}`);
    return undefined;
  }
  return uri.fsPath;
}

/** 用 VS Code API 实现 extension/lib/commands.js 约定的宿主接口 */
function createHost(context, store, validPresetName) {
  const config = () => vscode.workspace.getConfiguration("mddtt");
  const workspaceDir = () => {
    const folder = vscode.workspace.workspaceFolders?.[0];
    return folder ? path.join(folder.uri.fsPath, store.WORKSPACE_PRESET_DIR) : null;
  };

  return {
    get settings() {
      return {
        defaultPreset: config().get("defaultPreset", ""),
        presetStorage: config().get("presetStorage", "user"),
        confirmOverwrite: config().get("confirmOverwrite", true),
        openAfterConvert: config().get("openAfterConvert", true),
      };
    },

    log: (message) => output.info(message),
    warn: (message) => output.warn(message),
    error: (message) => output.error(message),
    showOutput: () => output.show(),

    async reportError(message) {
      await vscode.window.showErrorMessage(message);
    },
    async info(message, ...actions) {
      return vscode.window.showInformationMessage(message, ...actions);
    },

    /** 文件选择对话框：命令面板入口的「提交 docx」方式 */
    async pickFile(ext) {
      const isMarkdown = ext === MD_EXT;
      const picked = await vscode.window.showOpenDialog({
        canSelectFiles: true,
        canSelectMany: false,
        filters: isMarkdown ? { Markdown: ["md"] } : { "Word 文档": ["docx"] },
        title: isMarkdown ? "选择要转换的 Markdown 文件" : "选择要转换 / 提取版式的 Word 文档",
        defaultUri: vscode.workspace.workspaceFolders?.[0]?.uri,
      });
      return picked?.[0]?.fsPath;
    },

    async promptPresetName(suggested) {
      return vscode.window.showInputBox({
        title: "保存为自定义预设",
        prompt: "预设名：中英文、数字、下划线、连字符（不以连字符开头）",
        value: suggested,
        valueSelection: [0, suggested.length],
        validateInput: (value) =>
          validPresetName(value.trim()) ? null : "预设名不合法（仅限中英文、数字、下划线、连字符，且不以连字符开头）",
      });
    },

    async confirm(message, confirmLabel) {
      const picked = await vscode.window.showWarningMessage(message, { modal: true }, confirmLabel);
      return picked === confirmLabel;
    },

    async pickFromList(placeHolder, items) {
      const picked = await vscode.window.showQuickPick(items, { placeHolder, matchOnDescription: true });
      return picked?.value;
    },

    async setDefaultPreset(name) {
      await config().update("defaultPreset", name, vscode.ConfigurationTarget.Global);
    },

    async withProgress(title, task) {
      return vscode.window.withProgress(
        { location: vscode.ProgressLocation.Notification, title, cancellable: false },
        () => task()
      );
    },

    readFile: (file) => fs.promises.readFile(file),
    readText: (file) => fs.promises.readFile(file, "utf-8"),
    writeFile: (file, data) => fs.promises.writeFile(file, data),
    async exists(file) {
      try {
        await fs.promises.access(file);
        return true;
      } catch {
        return false;
      }
    },

    async openInEditor(file) {
      const document = await vscode.workspace.openTextDocument(vscode.Uri.file(file));
      await vscode.window.showTextDocument(document, { preview: false });
    },
    async openExternal(file) {
      await vscode.env.openExternal(vscode.Uri.file(file));
    },

    presetDirs() {
      return {
        userDir: path.join(context.globalStorageUri.fsPath, store.USER_PRESET_SUBDIR),
        workspaceDir: workspaceDir(),
        cliDir: store.cliPresetsDir(),
      };
    },
  };
}

module.exports = { activate, deactivate };
