// 预设发现与加载：内置预设 + 自定义预设（工作区级 / 用户级 / 命令行遗留目录）
//
// 本模块不依赖 vscode 模块，宿主只提供三个目录路径，便于脱离 VS Code 做单元测试。
// 搜索优先级：内置 → 工作区级 → 用户级 → 命令行遗留目录（同名时高优先级者生效）。

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { presets as builtinPresets } from "../../src/presets.js";
import { loadCustomPreset, validPresetName } from "../../src/preset-extract.js";

/** 用户级预设目录名（位于 VS Code globalStorage 下） */
export const USER_PRESET_SUBDIR = "presets";

/** 工作区级预设目录（相对工作区根目录；可随仓库提交、团队共享） */
export const WORKSPACE_PRESET_DIR = path.join(".vscode", "mddtt-presets");

/** 内置预设名 */
export const builtinPresetNames = Object.keys(builtinPresets);

/** 预设来源标签（用于列表展示） */
export const PRESET_KIND_LABEL = {
  builtin: "内置",
  workspace: "工作区",
  user: "用户",
  cli: "命令行（只读）",
};

/** 命令行自定义预设目录：MDDTT_HOME 可重定向，默认 ~/.mddtt/presets/ */
export function cliPresetsDir() {
  return process.env.MDDTT_HOME
    ? path.join(process.env.MDDTT_HOME, "presets")
    : path.join(os.homedir(), ".mddtt", "presets");
}

/** 按优先级排列的自定义预设目录列表 */
function customDirs({ workspaceDir = null, userDir = null, cliDir = null } = {}) {
  return [
    ["workspace", workspaceDir],
    ["user", userDir],
    ["cli", cliDir],
  ].filter(([, dir]) => Boolean(dir));
}

/**
 * 列出全部可用预设：[{ name, kind, dir }]
 * 同名时只保留优先级最高的一项；内置预设排在最前，自定义预设按名称排序。
 */
export function listAllPresets(dirs) {
  const list = builtinPresetNames.map((name) => ({ name, kind: "builtin", dir: null }));
  const seen = new Set(list.map((p) => p.name));
  const custom = [];
  for (const [kind, dir] of customDirs(dirs)) {
    if (!fs.existsSync(dir)) continue;
    for (const file of fs.readdirSync(dir)) {
      if (!file.endsWith(".json")) continue;
      const name = file.slice(0, -5);
      if (seen.has(name)) continue; // 高优先级目录已提供同名预设
      seen.add(name);
      custom.push({ name, kind, dir });
    }
  }
  custom.sort((a, b) => a.name.localeCompare(b.name));
  return [...list, ...custom];
}

/**
 * 解析指定预设为待合并的补丁对象。
 * 返回 { patch } / { patch: null, file, error }：找不到时给出 error 文案。
 */
export function resolvePresetPatch(name, dirs) {
  if (builtinPresets[name]) return { patch: builtinPresets[name], file: null };
  for (const [, dir] of customDirs(dirs)) {
    const file = path.join(dir, `${name}.json`);
    if (fs.existsSync(file)) return { patch: loadCustomPreset(name, dir), file };
  }
  return { patch: null, file: null, error: `未知预设「${name}」` };
}

/** 复用核心的预设名校验（此处再导出，供宿主做输入校验） */
export { validPresetName };
