// MDDTT VS Code 命令实现
//
// 设计约束：本模块只依赖「宿主接口」(host) 与项目内核 (src/)，不 import vscode。
// 真实宿主由 extension/extension.cjs 用 VS Code API 实现，测试用桩宿主替代，
// 因此命令逻辑可以脱离 VS Code 运行验证。
//
// 宿主接口（createCommands 的入参）：
//   get settings()                      → { defaultPreset, presetStorage, confirmOverwrite, openAfterConvert }
//   log / warn / error(msg)             → 输出通道
//   reportError(msg)                    → 错误弹窗
//   info(msg, ...actions) → action      → 信息提示，返回用户点击的按钮
//   pickFile(ext) → path|null           → 文件选择对话框（".md" / ".docx"）
//   promptPresetName(suggested) → name  → 预设名输入框
//   confirm(msg, label) → bool          → 模态确认框
//   pickFromList(placeHolder, items) → value
//   setDefaultPreset(name)              → 写入配置
//   withProgress(title, task) → result  → 进度提示包裹
//   readFile(path) → Uint8Array         → 读二进制
//   readText(path) → string             → 读文本
//   writeFile(path, data)               → 写文件（Buffer / string）
//   exists(path) → bool
//   openInEditor(path) / openExternal(path) / showOutput()
//   presetDirs() → { userDir, workspaceDir, cliDir }
//
// 两个 docx 入口（本文件同时支持）：
//   1. 资源管理器右键 / 编辑器标签右键：命令收到 Uri，宿主换算为路径后传入
//   2. 命令面板：无路径参数 → inputPath 为 undefined → pickFile 弹文件选择对话框

import path from "node:path";
import { defaultOptions, mergeOptions } from "../../src/options.js";
import { convertMarkdownText } from "../../src/converter.js";
import { convertDocxBufferToMarkdown } from "../../src/docx-to-md.js";
import {
  customPresetPath,
  extractPresetOptionsFromBuffer,
  presetSummaryLines,
  savePresetFromExtraction,
  validPresetName,
} from "../../src/preset-extract.js";
import { PRESET_KIND_LABEL, listAllPresets, resolvePresetPatch } from "./presets-store.js";

const MD_EXT = ".md";
const DOCX_EXT = ".docx";

const extOf = (p) => path.extname(p).toLowerCase();
const swapExt = (p, ext) => p.slice(0, p.length - path.extname(p).length) + ext;

/** 统一失败出口：写日志 + 弹错误提示 */
async function fail(host, action, err) {
  const message = err?.message ?? String(err);
  host.error(`${action}失败：${message}`);
  await host.reportError(`${action}失败：${message}`);
}

/** 输出文件已存在时按配置决定是否覆盖 */
async function confirmWritable(host, output) {
  if (!(await host.exists(output))) return true;
  if (!host.settings.confirmOverwrite) {
    host.warn(`输出文件已存在，按设置直接覆盖：${output}`);
    return true;
  }
  const ok = await host.confirm(`输出文件已存在，是否覆盖？\n\n${output}`, "覆盖");
  if (!ok) host.log(`已取消：${output} 已存在`);
  return ok;
}

/** 由 docx 文件名推导建议预设名（去掉后缀与不合法的字符） */
export function suggestPresetName(inputPath) {
  const base = path.basename(inputPath, path.extname(inputPath));
  const cleaned = base.replace(/[^\w\u4e00-\u9fff-]/g, "").replace(/^-+/, "").slice(0, 64);
  return validPresetName(cleaned) ? cleaned : "preset";
}

/** 解析 md → docx 使用的预设补丁；返回 { patch, name, error } */
function resolvePresetForConvert(host) {
  const name = String(host.settings.defaultPreset ?? "").trim();
  if (!name) return { patch: null, name: "" };
  const resolved = resolvePresetPatch(name, host.presetDirs());
  if (resolved.error) {
    const available = listAllPresets(host.presetDirs()).map((p) => p.name).join(" / ");
    return { patch: null, name, error: `${resolved.error}（可用预设：${available}）` };
  }
  return { patch: resolved.patch, name };
}

export function createCommands(host) {
  /** Markdown → docx；inputPath 为空时弹出文件选择对话框 */
  async function convertMarkdownToDocx(inputPath) {
    const action = "Markdown 转 docx";
    const input = inputPath ?? (await host.pickFile(MD_EXT));
    if (!input) {
      host.log("已取消：未选择 Markdown 文件");
      return;
    }
    if (extOf(input) !== MD_EXT) {
      await fail(host, action, new Error(`仅支持 .md 文件（当前：${path.basename(input)}）`));
      return;
    }
    const output = swapExt(input, DOCX_EXT);
    if (!(await confirmWritable(host, output))) return;

    const preset = resolvePresetForConvert(host);
    if (preset.error) {
      await fail(host, action, new Error(preset.error));
      return;
    }
    const opts = mergeOptions(defaultOptions, preset.patch ?? {});

    let buffer;
    try {
      const mdText = await host.readText(input);
      buffer = await host.withProgress(
        `正在转换为 docx：${path.basename(input)}`,
        () => convertMarkdownText(mdText, opts, path.dirname(input))
      );
      await host.writeFile(output, buffer);
    } catch (err) {
      await fail(host, action, err);
      return;
    }

    const presetLabel = preset.name ? `预设：${preset.name}` : "默认格式";
    host.log(`转换完成：${output}（${presetLabel}）`);
    if (host.settings.openAfterConvert) {
      const choice = await host.info(`已生成 ${path.basename(output)}（${presetLabel}）`, "打开");
      if (choice === "打开") await host.openExternal(output);
    } else {
      await host.info(`已生成 ${path.basename(output)}（${presetLabel}）`);
    }
  }

  /** docx → Markdown；inputPath 为空时弹出文件选择对话框 */
  async function convertDocxToMarkdown(inputPath) {
    const action = "docx 转 Markdown";
    const input = inputPath ?? (await host.pickFile(DOCX_EXT));
    if (!input) {
      host.log("已取消：未选择 Word 文档");
      return;
    }
    if (extOf(input) !== DOCX_EXT) {
      await fail(host, action, new Error(`仅支持 .docx 文件（当前：${path.basename(input)}）`));
      return;
    }
    const output = swapExt(input, MD_EXT);
    if (!(await confirmWritable(host, output))) return;

    try {
      const buffer = await host.readFile(input);
      const markdown = await host.withProgress(
        `正在转换为 Markdown：${path.basename(input)}`,
        () => convertDocxBufferToMarkdown(buffer, path.dirname(output))
      );
      await host.writeFile(output, markdown);
    } catch (err) {
      await fail(host, action, err);
      return;
    }

    host.log(`转换完成：${output}`);
    if (host.settings.openAfterConvert) {
      const choice = await host.info(`已生成 ${path.basename(output)}`, "打开");
      if (choice === "打开") await host.openInEditor(output);
    } else {
      await host.info(`已生成 ${path.basename(output)}`);
    }
  }

  /**
   * 从 docx 提取版式并保存为自定义预设。
   * 入口 1：资源管理器/标签右键 → 传入 docx 路径。
   * 入口 2：命令面板 → 不传路径 → 弹出文件选择对话框。
   * VS Code 无法阅读 docx，因此这里呈现的是「版式摘要」而非文档内容。
   */
  async function extractPreset(inputPath) {
    const action = "版式提取";
    const input = inputPath ?? (await host.pickFile(DOCX_EXT));
    if (!input) {
      host.log("已取消：未选择 Word 文档");
      return;
    }
    if (extOf(input) !== DOCX_EXT) {
      await fail(host, action, new Error(`仅支持 .docx 文件（当前：${path.basename(input)}）`));
      return;
    }

    let extraction;
    try {
      const buffer = await host.readFile(input);
      extraction = await host.withProgress(
        `正在提取版式：${path.basename(input)}`,
        () => extractPresetOptionsFromBuffer(buffer)
      );
    } catch (err) {
      await fail(host, action, err);
      return;
    }

    // 版式摘要：VS Code 无法显示 docx 内容，这里把可映射的版式逐行写入输出通道
    host.log(`版式提取：${input}`);
    for (const line of presetSummaryLines(extraction.options)) host.log(`  ${line}`);
    for (const note of extraction.notes) host.warn(`  - ${note}`);

    const name = await host.promptPresetName(suggestPresetName(input));
    const presetName = name === undefined || name === null ? "" : String(name).trim();
    if (!presetName) {
      host.log("已取消：未输入预设名");
      return;
    }
    if (!validPresetName(presetName)) {
      await fail(host, action, new Error(`预设名「${presetName}」不合法（仅限中英文、数字、下划线、连字符，且不以连字符开头）`));
      return;
    }

    const dirs = host.presetDirs();
    const useWorkspace = host.settings.presetStorage === "workspace" && dirs.workspaceDir;
    const dir = useWorkspace ? dirs.workspaceDir : dirs.userDir;
    const file = customPresetPath(presetName, dir);
    const exists = await host.exists(file);
    if (exists) {
      const ok = await host.confirm(`预设「${presetName}」已存在，是否覆盖？\n\n${file}`, "覆盖");
      if (!ok) {
        host.log(`已取消：预设「${presetName}」已存在`);
        return;
      }
    }

    let saved;
    try {
      saved = await savePresetFromExtraction(extraction, presetName, {
        overwrite: exists,
        dir,
        source: input,
      });
    } catch (err) {
      await fail(host, action, err);
      return;
    }

    const scopeLabel = useWorkspace ? "工作区级" : "用户级";
    host.log(`预设已保存（${scopeLabel}）：${saved.file}`);
    host.log(`复用方式：将设置 mddtt.defaultPreset 设为「${presetName}」，或用命令「MDDTT: 选择默认预设」`);
    const choice = await host.info(`预设「${presetName}」已保存（${scopeLabel}）`, "打开预设文件");
    if (choice === "打开预设文件") await host.openInEditor(saved.file);
  }

  /** 选择 md → docx 使用的默认预设（写入 mddtt.defaultPreset） */
  async function selectDefaultPreset() {
    const dirs = host.presetDirs();
    const current = String(host.settings.defaultPreset ?? "").trim();
    const items = [
      {
        label: "（默认格式）",
        description: "与 sundy 排版相同，仅无页眉页脚页码",
        value: "",
      },
      ...listAllPresets(dirs).map((preset) => ({
        label: preset.name,
        description: `${PRESET_KIND_LABEL[preset.kind] ?? preset.kind}${preset.name === current ? "（当前）" : ""}`,
        value: preset.name,
      })),
    ];
    const picked = await host.pickFromList("选择 Markdown 转 docx 使用的预设", items);
    if (picked === undefined || picked === null) {
      host.log("已取消：未选择预设");
      return;
    }
    await host.setDefaultPreset(picked);
    host.log(`默认预设已设为：${picked || "（默认格式）"}`);
    await host.info(picked ? `默认预设：${picked}` : "已改用默认格式");
  }

  return {
    convertMarkdownToDocx,
    convertDocxToMarkdown,
    extractPreset,
    selectDefaultPreset,
    showOutput: () => host.showOutput(),
  };
}
