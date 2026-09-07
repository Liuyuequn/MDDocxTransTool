// 单元测试：纯函数层（单位换算、字号解析、深合并、参数解析、docx 导入注释）
// 运行：node test/unit-test.mjs（npm test 会先跑本文件再跑端到端校验）

import { test } from "node:test";
import assert from "node:assert/strict";
import {
  cmToTwip,
  ptToTwip,
  ptToHalfPoint,
  parseFontSize,
  mergeOptions,
  PAGE_SIZES,
  defaultOptions,
} from "../src/options.js";
import { parseArgs, parseMargin, parseHeadingSize } from "../src/args.js";
import { validPresetName } from "../src/preset-extract.js";
import { maxImageWidthPx } from "../src/inline.js";
import { linesToTwip } from "../src/styles.js";
import {
  commentAnnotation,
  revisionComment,
  summarizeCommentRange,
} from "../src/docx-import/document-model.js";

// ============ preset-extract.js：预设名校验 ============

test("validPresetName: 中英文/数字/下划线/连字符合法", () => {
  assert.equal(validPresetName("firm"), true);
  assert.equal(validPresetName("律所模板1"), true);
  assert.equal(validPresetName("my-preset_2"), true);
});

test("validPresetName: 拒绝路径分隔符、空格与空值", () => {
  assert.equal(validPresetName("a/b"), false);
  assert.equal(validPresetName("a b"), false);
  assert.equal(validPresetName(".."), false);
  assert.equal(validPresetName("-x"), false); // 不以连字符开头
  assert.equal(validPresetName(""), false);
  assert.equal(validPresetName(null), false);
});

// ============ options.js：单位换算 ============

test("cmToTwip: 2.54cm = 1440 twip（1 英寸）", () => {
  assert.equal(cmToTwip(2.54), 1440);
});

test("cmToTwip: 0 = 0", () => {
  assert.equal(cmToTwip(0), 0);
});

test("cmToTwip: 3.18cm 四舍五入为 1803", () => {
  assert.equal(cmToTwip(3.18), 1803);
});

test("ptToTwip: 1pt = 20 twip", () => {
  assert.equal(ptToTwip(12), 240);
});

test("ptToHalfPoint: 四号 14pt = 28 半磅", () => {
  assert.equal(ptToHalfPoint(14), 28);
});

test("ptToHalfPoint: 五号 10.5pt = 21 半磅", () => {
  assert.equal(ptToHalfPoint(10.5), 21);
});

// ============ options.js：字号解析 ============

test("parseFontSize: 数字字符串解析为 pt", () => {
  assert.equal(parseFontSize("14"), 14);
});

test("parseFontSize: 纯数字直接返回", () => {
  assert.equal(parseFontSize(12), 12);
});

test("parseFontSize: 中文字号名映射（四号=14）", () => {
  assert.equal(parseFontSize("四号"), 14);
});

test("parseFontSize: 中文字号名映射（五号=10.5）", () => {
  assert.equal(parseFontSize("五号"), 10.5);
});

test("parseFontSize: 容忍前后空白", () => {
  assert.equal(parseFontSize("  小五 "), 9);
});

test("parseFontSize: 无法识别时抛错", () => {
  assert.throws(() => parseFontSize("胡说"), /无法识别的字号/);
});

// ============ options.js：深合并 ============

test("mergeOptions: 嵌套对象递归合并", () => {
  const base = { page: { size: "A4", margin: { top: 1, left: 1 } } };
  const patch = { page: { margin: { top: 2 } } };
  const merged = mergeOptions(base, patch);
  assert.deepEqual(merged, { page: { size: "A4", margin: { top: 2, left: 1 } } });
});

test("mergeOptions: 数组整体覆盖而非逐项合并", () => {
  const base = { sizes: { heading: [22, 16, 14, 12, 11, 11] } };
  const patch = { sizes: { heading: [22, 16, 14, 14, 14, 14] } };
  const merged = mergeOptions(base, patch);
  assert.deepEqual(merged.sizes.heading, [22, 16, 14, 14, 14, 14]);
});

test("mergeOptions: undefined 字段不覆盖已有值", () => {
  const base = { header: { text: "A" } };
  const patch = { header: { text: undefined, align: "left" } };
  const merged = mergeOptions(base, patch);
  assert.equal(merged.header.text, "A");
  assert.equal(merged.header.align, "left");
});

test("mergeOptions: 多个补丁按顺序叠加（后者胜）", () => {
  const base = { fonts: { body: { eastAsia: "等线" } } };
  const p1 = { fonts: { body: { eastAsia: "仿宋" } } };
  const p2 = { fonts: { body: { eastAsia: "黑体" } } };
  assert.equal(mergeOptions(base, p1, p2).fonts.body.eastAsia, "黑体");
});

test("mergeOptions: null 值直接覆盖（用于显式清空语义）", () => {
  const base = { header: { text: "A" } };
  const patch = { header: { text: null } };
  assert.equal(mergeOptions(base, patch).header.text, null);
});

// ============ args.js：parseMargin ============

test("parseMargin: 单值四边统一", () => {
  const patch = { page: {} };
  parseMargin("2.5", patch);
  assert.deepEqual(patch.page.margin, { top: 2.5, right: 2.5, bottom: 2.5, left: 2.5 });
});

test("parseMargin: 四值按 上,右,下,左 顺序", () => {
  const patch = { page: {} };
  parseMargin("2.54,3.18,2.54,3.18", patch);
  assert.deepEqual(patch.page.margin, { top: 2.54, right: 3.18, bottom: 2.54, left: 3.18 });
});

test("parseMargin: 容忍数值间空白", () => {
  const patch = { page: {} };
  parseMargin("1, 2, 3, 4", patch);
  assert.deepEqual(patch.page.margin, { top: 1, right: 2, bottom: 3, left: 4 });
});

test("parseMargin: 非数字抛错", () => {
  assert.throws(() => parseMargin("abc", { page: {} }), /页边距格式错误/);
});

test("parseMargin: 负数抛错", () => {
  assert.throws(() => parseMargin("-1", { page: {} }), /页边距格式错误/);
});

test("parseMargin: 三个值抛错（应为 1 或 4 个）", () => {
  assert.throws(() => parseMargin("1,2,3", { page: {} }), /1 个或 4 个/);
});

// ============ args.js：parseHeadingSize ============

test("parseHeadingSize: auto 不写入字号", () => {
  const patch = { heading: {}, sizes: {} };
  parseHeadingSize("auto", patch);
  assert.equal(patch.sizes.heading, undefined);
});

test("parseHeadingSize: 单值扩展为六级相同", () => {
  const patch = { heading: {}, sizes: {} };
  parseHeadingSize("22", patch);
  assert.deepEqual(patch.sizes.heading, [22, 22, 22, 22, 22, 22]);
});

test("parseHeadingSize: 六级逗号分隔原样生效", () => {
  const patch = { heading: {}, sizes: {} };
  parseHeadingSize("22,16,14,14,14,14", patch);
  assert.deepEqual(patch.sizes.heading, [22, 16, 14, 14, 14, 14]);
});

test("parseHeadingSize: 中文字号名可用于标题", () => {
  const patch = { heading: {}, sizes: {} };
  parseHeadingSize("二号", patch);
  assert.deepEqual(patch.sizes.heading, [22, 22, 22, 22, 22, 22]);
});

test("parseHeadingSize: 错误级数抛错", () => {
  assert.throws(() => parseHeadingSize("1,2,3", { heading: {}, sizes: {} }), /单一值或逗号分隔的六级值/);
});

// ============ args.js：parseArgs ============

test("parseArgs: 别名与全名等价（-f 与 --font）", () => {
  const a = parseArgs(["-f", "宋体"]);
  const b = parseArgs(["--font", "宋体"]);
  assert.deepEqual(a.patch.fonts.body, b.patch.fonts.body);
  assert.equal(a.patch.fonts.body.eastAsia, "宋体");
});

test("parseArgs: --name=value 等价于 --name value", () => {
  const a = parseArgs(["--font=宋体"]);
  const b = parseArgs(["--font", "宋体"]);
  assert.deepEqual(a.patch, b.patch);
});

test("parseArgs: flag 参数无需取值", () => {
  const r = parseArgs(["--no-bold"]);
  assert.deepEqual(r.patch.heading.bold, [false, false, false, false, false, false]);
});

test("parseArgs: 输出控制参数单独归位（output/overwrite/preset）", () => {
  const r = parseArgs(["-o", "out.docx", "--overwrite", "--preset", "sundy"]);
  assert.equal(r.output, "out.docx");
  assert.equal(r.overwrite, true);
  assert.equal(r.preset, "sundy");
});

test("parseArgs: 未知参数记入 errors", () => {
  const r = parseArgs(["--no-such-param"]);
  assert.equal(r.errors.length, 1);
  assert.match(r.errors[0], /未知参数/);
});

test("parseArgs: 参数缺少取值记入 errors", () => {
  const r = parseArgs(["--font"]);
  assert.equal(r.errors.length, 1);
  assert.match(r.errors[0], /缺少取值/);
});

test("parseArgs: 取值非法记入 errors（校验函数）", () => {
  const r = parseArgs(["--orientation", "diagonal"]);
  assert.equal(r.errors.length, 1);
  assert.match(r.errors[0], /portrait 或 landscape/);
});

test("parseArgs: 互斥页眉参数记入 errors", () => {
  const r = parseArgs(["--header", "A", "--header-left", "B"]);
  assert.ok(r.errors.some((e) => e.includes("互斥")));
});

test("parseArgs: --total-pages 在无 Y 模板上追加 /共Y页", () => {
  const r = parseArgs(["-p", "bottom", "--total-pages"]);
  assert.equal(r.patch.pageNumber.format, "X/共Y页");
});

test("parseArgs: --total-pages 不改动已含 Y 的模板", () => {
  const r = parseArgs(["--page-num-format", "第X页/共Y页", "--total-pages"]);
  assert.equal(r.patch.pageNumber.format, "第X页/共Y页");
});

test("parseArgs: 非选项 token 记为多余位置参数", () => {
  const r = parseArgs(["to"]);
  assert.equal(r.errors.length, 1);
  assert.match(r.errors[0], /多余的位置参数/);
});

test("parseArgs: --line-rule 与 --line-height 组合", () => {
  const r = parseArgs(["--line-rule", "exact", "--line-height", "20"]);
  assert.equal(r.patch.paragraph.lineRule, "exact");
  assert.equal(r.patch.paragraph.line, 20);
});

test("parseArgs: 非法 --line-rule 记入 errors", () => {
  const r = parseArgs(["--line-rule", "foo"]);
  assert.equal(r.errors.length, 1);
  assert.match(r.errors[0], /auto、exact 或 atLeast/);
});

// ============ styles.js：行距换算 ============

test("linesToTwip: auto 倍数行距按 字号×倍数 换算", () => {
  assert.equal(linesToTwip(1, 14, { line: 1.5, lineRule: "auto" }), 420); // 14×1.5×20
});

test("linesToTwip: exact 固定行距与字号无关", () => {
  assert.equal(linesToTwip(1, 14, { line: 20, lineRule: "exact" }), 400);
  assert.equal(linesToTwip(1, 22, { line: 20, lineRule: "exact" }), 400);
});

test("linesToTwip: 无 lineCfg 时退回默认 1.28 倍", () => {
  assert.equal(linesToTwip(1, 14, null), Math.round(14 * 1.28 * 20));
});

// ============ inline.js：正文图片宽度上限 ============

test("maxImageWidthPx: 默认 A4 = 内容区宽度(twip)÷15 取整", () => {
  const contentTwip = PAGE_SIZES.A4.width - 2 * cmToTwip(defaultOptions.page.margin.left);
  assert.equal(maxImageWidthPx(defaultOptions), Math.floor(contentTwip / 15));
});

test("maxImageWidthPx: 边距变宽时上限随之缩小", () => {
  const wide = mergeOptions(defaultOptions, { page: { margin: { left: 5, right: 5 } } });
  const contentTwip = PAGE_SIZES.A4.width - 2 * cmToTwip(5);
  assert.equal(maxImageWidthPx(wide), Math.floor(contentTwip / 15));
});

test("maxImageWidthPx: 窄页面上限更小（A5 横向口径）", () => {
  const a5 = mergeOptions(defaultOptions, { page: { size: "A5" } });
  const contentTwip = PAGE_SIZES.A5.width - 2 * cmToTwip(defaultOptions.page.margin.left);
  assert.equal(maxImageWidthPx(a5), Math.floor(contentTwip / 15));
});

// ============ docx-import：修订与批注注释 ============

test("revisionComment: 修订前原文写入 Markdown 注释", () => {
  assert.equal(revisionComment("旧条款"), "<!-- 此处系修订；修订前原文：旧条款 -->");
});

test("summarizeCommentRange: 超过 20 字符时保留首尾各 10 字符", () => {
  assert.equal(
    summarizeCommentRange("ABCDEFGHIJKLMNOPQRSTUVWXYZ"),
    "ABCDEFGHIJ……QRSTUVWXYZ"
  );
});

test("commentAnnotation: 标注范围并包含批注正文", () => {
  assert.equal(
    commentAnnotation("被批注内容", "请核对此条款"),
    "<!-- 批注范围：“被批注内容”；此处有批注：请核对此条款 -->"
  );
});
