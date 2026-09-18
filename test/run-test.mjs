// 端到端校验：覆盖 md/docx 双向转换、预设、错误处理、Schema 前置输出与 docx 富结构降级
import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

// 期望值统一由“配置声明 + 换算函数”推导，避免魔法数与实现耦合；
// 换算函数本身由 unit-test.mjs 覆盖
import { cmToTwip, ptToHalfPoint, PAGE_SIZES, CHINESE_FONT_SIZES, defaultOptions } from "../src/options.js";
import { presets } from "../src/presets.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..");
const cli = path.join(root, "src", "cli.js");
const mdPath = path.join(__dirname, "sample.md");

let failed = 0;
function check(name, cond) {
  console.log(`${cond ? "通过" : "失败"} | ${name}`);
  if (!cond) failed++;
}

async function convert(outName, extraArgs = []) {
  const outPath = path.join(__dirname, outName);
  if (fs.existsSync(outPath)) fs.unlinkSync(outPath);
  execFileSync(process.execPath, [cli, mdPath, "-o", outPath, ...extraArgs], { stdio: "pipe" });
  const zip = await JSZip.loadAsync(fs.readFileSync(outPath));
  const headerFiles = Object.keys(zip.files).filter((f) => /^word\/header\d+\.xml$/.test(f));
  const footerFiles = Object.keys(zip.files).filter((f) => /^word\/footer\d+\.xml$/.test(f));
  const readAll = (files) => files.map((f) => zip.file(f).async("string"));
  return {
    document: await zip.file("word/document.xml").async("string"),
    styles: await zip.file("word/styles.xml").async("string"),
    numbering: await zip.file("word/numbering.xml").async("string"),
    header: headerFiles.length ? (await Promise.all(readAll(headerFiles))).join("\n") : null,
    footer: footerFiles.length ? (await Promise.all(readAll(footerFiles))).join("\n") : null,
  };
}

function styleBlock(stylesXml, styleId) {
  const re = new RegExp(`<w:style [^>]*w:styleId="${styleId}"[\\s\\S]*?</w:style>`);
  const m = stylesXml.match(re);
  return m ? m[0] : "";
}

function docDefaults(stylesXml) {
  const m = stylesXml.match(/<w:docDefaults>[\s\S]*?<\/w:docDefaults>/);
  return m ? m[0] : "";
}

function paragraphWithText(documentXml, text) {
  return (documentXml.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g) || [])
    .find((p) => p.includes(text)) || "";
}

// ============ 用例 1：默认转换 ============
console.log("—— 用例 1：默认转换 ——");
const d = await convert("sample-default.docx");
// 约定：默认输出 = sundy 排版 − 页眉页脚页码（期望值从 defaultOptions 推导）
const DF = defaultOptions;
const dfLineTwip = String(Math.round(DF.paragraph.line * 240));
const dfIndentChars = String(Math.round(DF.paragraph.firstLineChars * 100));
const dfAfterTwip = String(Math.round(DF.paragraph.afterLines * DF.sizes.body * DF.paragraph.line * 20));
const dH1 = styleBlock(d.styles, "Heading1");
const dBody = paragraphWithText(d.document, "中英文混排测试");
[
  ["标题样式 Heading1", d.document.includes('<w:pStyle w:val="Heading1"/>')],
  ["六级标题 Heading6", d.document.includes('<w:pStyle w:val="Heading6"/>')],
  ["加粗", d.document.includes("<w:b/>")],
  ["斜体", d.document.includes("<w:i/>")],
  ["删除线", d.document.includes("<w:strike/>")],
  ["行内代码底纹", d.document.includes("EFEFEF")],
  ["代码块底纹", d.document.includes("F2F2F2")],
  ["表格", d.document.includes("<w:tbl>")],
  ["表头底纹", d.document.includes("EFEFEF")],
  ["图片", d.document.includes("<w:drawing>")],
  ["超链接", d.document.includes("<w:hyperlink")],
  ["引用左边框", d.document.includes("w:left")],
  ["有序列表编号", d.numbering.includes("%1.")],
  ["无序列表编号", d.numbering.includes("\u2022")],
  ["A4 页面尺寸", d.document.includes('w:w="11906"')],
  [`默认首行缩进两字符（firstLineChars=${dfIndentChars}）`, d.document.includes(`w:firstLineChars="${dfIndentChars}"`)],
  [`默认行距 ${DF.paragraph.line}（w:line=${dfLineTwip}）`, docDefaults(d.styles).includes(`w:line="${dfLineTwip}"`)],
  [`正文段落直接写入 ${DF.paragraph.line} 倍行距`, dBody.includes(`w:line="${dfLineTwip}"`) && dBody.includes('w:lineRule="auto"')],
  [`默认段后 ${DF.paragraph.afterLines} 行（w:after=${dfAfterTwip}）`, docDefaults(d.styles).includes(`w:after="${dfAfterTwip}"`)],
  ["默认 H1 居中", dH1.includes('<w:jc w:val="center"/>')],
  ["默认无页眉", d.header === null],
  ["默认无页脚", d.footer === null],
  ["默认无页码字段", !d.document.includes("PAGE")],
  ["缺失图片容错提示", d.document.includes("图片缺失")],
].forEach(([n, c]) => check(n, c));

// ============ 用例 2：preset sundy ============
console.log("\n—— 用例 2：preset sundy ——");
const l = await convert("sample-sundy.docx", ["--preset", "sundy"]);
const h1 = styleBlock(l.styles, "Heading1");
const h4 = styleBlock(l.styles, "Heading4");
const lDefaults = docDefaults(l.styles);
// 期望值全部从 sundy 预设声明推导
const LG = presets.sundy;
const lgA4w = String(PAGE_SIZES[LG.page.size].width);
const lgMarginTop = String(cmToTwip(LG.page.margin.top));
const lgMarginLeft = String(cmToTwip(LG.page.margin.left));
const lgMarginHeader = String(cmToTwip(LG.page.margin.header));
const lgH1Half = String(ptToHalfPoint(LG.sizes.heading[0]));
const lgH4Half = String(ptToHalfPoint(LG.sizes.heading[3]));
const lgBodyHalf = String(ptToHalfPoint(LG.sizes.body));
const lgHeaderHalf = String(ptToHalfPoint(LG.sizes.header));
const lgFooterHalf = String(ptToHalfPoint(LG.sizes.footer));
const lgLineTwip = String(Math.round(LG.paragraph.line * 240));
const lgIndentChars = String(Math.round(LG.paragraph.firstLineChars * 100));
const lBody = paragraphWithText(l.document, "中英文混排测试");
[
  // 纸张
  ["A4 尺寸", l.document.includes(`w:w="${lgA4w}"`)],
  ["上下页边距按预设（twip）", l.document.includes(`w:top="${lgMarginTop}"`) && l.document.includes(`w:bottom="${lgMarginTop}"`)],
  ["左右页边距按预设（twip）", l.document.includes(`w:left="${lgMarginLeft}"`) && l.document.includes(`w:right="${lgMarginLeft}"`)],
  [`页眉距顶端 ${LG.page.margin.header}cm，整体上移并与正文留白`, l.document.includes(`w:header="${lgMarginHeader}"`)],
  // 标题
  ["标题宋体", h1.includes('w:eastAsia="宋体"')],
  ["标题西文 Times New Roman", h1.includes('w:ascii="Times New Roman"')],
  ["H1 字号按预设（半磅）", h1.includes(`w:val="${lgH1Half}"`)],
  ["H1 居中", h1.includes('<w:jc w:val="center"/>')],
  ["H1 加粗", h1.includes("<w:b/>")],
  ["H4 字号按预设且不加粗", h4.includes(`w:val="${lgH4Half}"`) && !h4.includes("<w:b/>")],
  // 正文
  ["正文仿宋", lDefaults.includes('w:eastAsia="仿宋"')],
  ["正文西文 Times New Roman", lDefaults.includes('w:ascii="Times New Roman"')],
  ["正文字号按预设（半磅）", lDefaults.includes(`<w:sz w:val="${lgBodyHalf}"/>`)],
  [`首行缩进两字符（firstLineChars=${lgIndentChars}）`, l.document.includes(`w:firstLineChars="${lgIndentChars}"`)],
  [`行距按预设（w:line=${lgLineTwip} auto）`, lDefaults.includes(`w:line="${lgLineTwip}"`) && lDefaults.includes('w:lineRule="auto"')],
  [`sundy 正文段落直接写入 ${LG.paragraph.line} 倍行距`, lBody.includes(`w:line="${lgLineTwip}"`) && lBody.includes('w:lineRule="auto"')],
  // 页眉页脚
  ["页眉三行文字（所名/官网/地址）", l.header && l.header.includes("圣典律师事务所") && l.header.includes("圣典官网") && l.header.includes("总所地址") && l.header.includes("新城科技园4A栋6楼、7楼")],
  ["页眉左对齐", l.header && l.header.includes('<w:jc w:val="left"/>')],
  ["页眉 logo 图片", l.header && l.header.includes("<w:drawing>")],
  ["页眉 URL 为超链接", l.header && l.header.includes("<w:hyperlink")],
  ["页眉字号字体按预设", l.header && l.header.includes(`w:val="${lgHeaderHalf}"`) && l.header.includes('w:eastAsia="仿宋"')],
  ["页眉渐变色带为原生渐变（a:gradFill）", l.header && l.header.includes("<a:gradFill")],
  ["渐变色带起点亮正红 E60012", l.header && l.header.includes('val="E60012"')],
  ["渐变色带终点金色 FFD700", l.header && l.header.includes('val="FFD700"')],
  ["渐变色带含 4 个颜色停靠点", l.header && (l.header.match(/<a:gs /g) || []).length >= 4],
  ["渐变色带橙金占比（45% 起橙红、70% 起纯橙）", l.header && l.header.includes('pos="45000"') && l.header.includes('pos="70000"')],
  ["页脚“第X页/共Y页”模板", l.footer && l.footer.includes("第") && l.footer.includes("页/共")],
  ["页脚页码字号按预设（半磅）", l.footer && l.footer.includes(`w:val="${lgFooterHalf}"`)],
  ["页脚页码仿宋", l.footer && l.footer.includes('w:eastAsia="仿宋"')],
  ["页脚含页码字段", l.footer && l.footer.includes("PAGE")],
  ["页脚含总页数字段", l.footer && l.footer.includes("NUMPAGES")],
  ["无 titlePage（首页不差异化）", !l.document.includes("<w:titlePg/>")],
].forEach(([n, c]) => check(n, c));

// ============ 用例 3：自定义参数 ============
console.log("\n—— 用例 3：自定义参数 ——");
const c = await convert("sample-custom.docx", [
  "--v-align", "center",
  "-p", "bottom", "--page-num-format", "第X页", "--page-num-start", "3",
  "--header-left", "MDDTT 测试", "--header-right", "2026-09-03",
  "--font", "黑体", "--font-size", "小四",
  "--indent", "2", "--align", "justify", "--no-italic",
  "--margin", "2,2,3,2",
]);
[
  ["页面垂直居中（v-align center）", c.document.includes('w:vAlign w:val="center"') || c.document.includes('w:val="center"')],
  [`自定义页边距（上 2cm=${cmToTwip(2)} twip）`, c.document.includes(`w:top="${cmToTwip(2)}"`)],
  ["左分布页眉（tab 右对齐）", c.header && c.header.includes("MDDTT 测试") && c.header.includes("2026-09-03")],
  ["页码格式“第X页”", c.footer && c.footer.includes("第") && c.footer.includes("PAGE")],
  ["起始页码 3", c.document.includes('w:start="3"')],
  ["正文黑体", c.styles.includes('w:eastAsia="黑体"')],
  [`正文小四（${ptToHalfPoint(CHINESE_FONT_SIZES["小四"])} 半磅）`, c.styles.includes(`w:val="${ptToHalfPoint(CHINESE_FONT_SIZES["小四"])}"`)],
  ["首行缩进两字符", c.document.includes('w:firstLineChars="200"')],
  ["正文两端对齐", c.document.includes('w:val="both"')],
  ["--no-italic 生效（无斜体）", !c.document.includes("<w:i/>")],
].forEach(([n, cond]) => check(n, cond));

// ============ 用例 4：错误处理 ============
console.log("\n—— 用例 4：错误处理 ——");
{
  const packageJson = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf-8"));
  const help = execFileSync(process.execPath, [cli, "--help"], { stdio: "pipe", encoding: "utf-8" });
  check("npm 包名和产品描述统一为 MDDTT", packageJson.name === "mddtt" && packageJson.description.startsWith("MDDTT -"));
  check("全局命令仅注册为 mddtt", JSON.stringify(packageJson.bin) === JSON.stringify({ mddtt: "./src/cli.js" }));
  check("帮助文本使用 MDDTT 品牌和 mddtt 命令", help.startsWith("MDDTT -") && help.includes("mddtt notes.md"));
}
function expectError(name, args) {
  try {
    execFileSync(process.execPath, [cli, ...args], { stdio: "pipe" });
    check(name, false);
  } catch (e) {
    check(name, e.status === 1 && e.stderr.length > 0);
  }
}
// 在 expectError 基础上，额外校验 stderr 含指定提示文本（用于验证错误反馈的引导信息）
function expectErrorWith(name, args, text) {
  try {
    execFileSync(process.execPath, [cli, ...args], { stdio: "pipe" });
    check(name, false);
  } catch (e) {
    check(name, e.status === 1 && e.stderr.toString().includes(text));
  }
}
expectError("未知参数报错", [mdPath, "--no-such-param"]);
expectError("互斥参数报错（--header 与 --header-left）", [mdPath, "--header", "A", "--header-left", "B"]);
expectError("未知预设报错", [mdPath, "--preset", "nope"]);
expectError("页边距格式错误报错", [mdPath, "-m", "abc"]);
expectErrorWith("旧语法 md to docx 提示直接使用", [mdPath, "to", "docx"], "请直接使用");
expectErrorWith("旧语法 docx to md 提示直接使用", [path.join(__dirname, "sample-default.docx"), "to", "md"], "请直接使用");
expectError("多余位置参数报错", [mdPath, "垃圾参数"]);
expectErrorWith("无参数报错并提示引号包裹", [], "英文引号");
expectErrorWith("仅选项参数无文件报错", ["--preset", "sundy"], "请指定要转换的文件");
expectErrorWith("找不到文件提示检查文件名", ["no-such-file.md"], "请检查文件名");
expectErrorWith(
  "文件名含空格未加引号提示整体包裹",
  ["no-such-dir", "我的", "文件.md"],
  'mddtt "no-such-dir 我的 文件.md"',
);
{
  // 不支持的扩展名：先造一个存在的 .txt 文件，确保走到扩展名校验分支而非“找不到文件”
  const txtPath = path.join(__dirname, "unsupported.txt");
  fs.writeFileSync(txtPath, "test", "utf-8");
  expectError("不支持的扩展名报错", [txtPath]);
  fs.unlinkSync(txtPath);
}
{
  // 无后缀文件存在：提示写明后缀（而非报“不支持的文件类型「」”）
  const noExtPath = path.join(__dirname, "no-extension-file");
  fs.writeFileSync(noExtPath, "test", "utf-8");
  expectErrorWith("无后缀文件提示写明后缀", [noExtPath], "请写明文件后缀");
  fs.unlinkSync(noExtPath);
}
{
  // 后缀自动匹配：仅 .md 存在 → 自动补全并转换成功；.md 与 .docx 共存 → 报错要求写明
  const stem = path.join(__dirname, "auto-ext-stem");
  const mdPath2 = `${stem}.md`;
  const docxPath2 = `${stem}.docx`;
  // 清理残留后仅创建 .md
  for (const p of [mdPath2, docxPath2, `${stem}-out.docx`]) { if (fs.existsSync(p)) fs.unlinkSync(p); }
  fs.writeFileSync(mdPath2, "# 自动匹配测试\n\n正文段落。\n", "utf-8");
  const stdout = execFileSync(process.execPath, [cli, stem, "-o", docxPath2], { stdio: "pipe", encoding: "utf-8" });
  check("省略后缀自动匹配 .md 并转换", stdout.includes("已自动匹配") && fs.existsSync(docxPath2));
  // 此时 .md 与 .docx 共存，再省略后缀应报错
  try {
    execFileSync(process.execPath, [cli, stem, "-o", `${stem}-out.docx`], { stdio: "pipe" });
    check("双后缀共存时报错要求写明", false);
  } catch (e) {
    check("双后缀共存时报错要求写明", e.status === 1 && e.stderr.toString().includes("写明后缀"));
  }
  // 无任何匹配 → 仍走“找不到文件”
  expectErrorWith("无匹配后缀提示找不到文件", [path.join(__dirname, "no-such-stem")], "找不到文件");
  fs.unlinkSync(mdPath2);
  fs.unlinkSync(docxPath2);
}

// ============ 用例 5：docx → Markdown ============
console.log("\n—— 用例 5：docx → Markdown ——");
{
  // 先用 sample.md 生成一个 docx，再转回 md
  const docxPath = path.join(__dirname, "docx2md-test.docx");
  const mdOutPath = path.join(__dirname, "docx2md-test.md");
  const picDir = path.join(__dirname, "MDPictures");
  // 清理残留
  for (const p of [docxPath, mdOutPath]) { if (fs.existsSync(p)) fs.unlinkSync(p); }
  if (fs.existsSync(picDir)) fs.rmSync(picDir, { recursive: true });
  // md → docx
  execFileSync(process.execPath, [cli, mdPath, "-o", docxPath], { stdio: "pipe" });
  check("docx→md 前置：生成测试 docx", fs.existsSync(docxPath));
  // docx → md
  execFileSync(process.execPath, [cli, docxPath, "-o", mdOutPath], { stdio: "pipe" });
  const mdOut = fs.readFileSync(mdOutPath, "utf-8");
  const pics = fs.existsSync(picDir) ? fs.readdirSync(picDir) : [];
  check("docx→md 输出文件存在", fs.existsSync(mdOutPath));
  check("docx→md 含一级标题", mdOut.includes("# "));
  check("docx→md 含二级标题", mdOut.includes("## "));
  check("docx→md 含加粗", mdOut.includes("**"));
  check("docx→md 含斜体", mdOut.includes("*"));
  check("docx→md 含超链接", mdOut.includes("](http"));
  check("docx→md 含图片引用", mdOut.includes("!["));
  check("docx→md 图片提取到 MDPictures 文件夹", pics.some((f) => f.startsWith("img-") && /\.(png|jpg|gif|bmp)$/.test(f)));
  check("docx→md 图片为相对路径引用", mdOut.includes("](MDPictures/img-"));
  check("docx→md 无 data URI 内嵌图片", !mdOut.includes("data:image"));
  check("docx→md 含列表项", mdOut.includes("- "));
  check("docx→md 含表格", mdOut.includes("| ---"));
  check("docx→md 无连续三个空行", !mdOut.includes("\n\n\n\n"));
  // 清理
  for (const p of [docxPath, mdOutPath]) { if (fs.existsSync(p)) fs.unlinkSync(p); }
  if (fs.existsSync(picDir)) fs.rmSync(picDir, { recursive: true });
}

// ============ 用例 6：docx → md 合并单元格（HTML 回退） ============
console.log("\n—— 用例 6：docx → md 合并单元格（HTML 回退）——");
{
  // 用 docx 库生成含横跨/纵跨合并单元格的测试文档
  const { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType } = await import("docx");
  const docxPath = path.join(__dirname, "merged-cells.docx");
  const mdOutPath = path.join(__dirname, "merged-cells.md");
  for (const p of [docxPath, mdOutPath]) { if (fs.existsSync(p)) fs.unlinkSync(p); }
  const doc = new Document({
    sections: [{
      children: [
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph("横跨两列")], columnSpan: 2 }),
                new TableCell({ children: [new Paragraph("C1")] }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph("纵跨两行")], rowSpan: 2 }),
                new TableCell({ children: [new Paragraph("B2")] }),
                new TableCell({ children: [new Paragraph("C2")] }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph("B3")] }),
                new TableCell({ children: [new Paragraph("C3")] }),
              ],
            }),
          ],
        }),
      ],
    }],
  });
  fs.writeFileSync(docxPath, await Packer.toBuffer(doc));
  execFileSync(process.execPath, [cli, docxPath, "-o", mdOutPath], { stdio: "pipe" });
  const out = fs.readFileSync(mdOutPath, "utf-8");
  check("合并表格：输出文件存在", fs.existsSync(mdOutPath));
  check("合并表格：保留 colspan 属性", out.includes('colspan="2"'));
  check("合并表格：保留 rowspan 属性", out.includes('rowspan="2"'));
  check("合并表格：HTML 回退含 <table>", out.includes("<table"));
  check("合并表格：含说明注释", out.includes("合并单元格"));
  check("合并表格：内容完整（横跨两列/纵跨两行/B3）", out.includes("横跨两列") && out.includes("纵跨两行") && out.includes("B3"));
  check("合并表格：未误生成管道表格", !out.includes("| ---"));
  for (const p of [docxPath, mdOutPath]) { if (fs.existsSync(p)) fs.unlinkSync(p); }
}

// ============ 用例 7：--save-preset 格式提取与复用 ============
console.log("\n—— 用例 7：--save-preset 格式提取与复用 ——");
{
  // MDDTT_HOME 重定向到临时目录，隔离用户真实的 ~/.mddtt/presets/
  const tmpHome = path.join(__dirname, "tmp-mddtt-home");
  fs.rmSync(tmpHome, { recursive: true, force: true });
  const env = { ...process.env, MDDTT_HOME: tmpHome };
  const run = (args) => execFileSync(process.execPath, [cli, ...args], { stdio: "pipe", encoding: "utf-8", env });

  // ---- 7.1 从自定义参数生成的 docx 提取 ----
  const srcDocx = path.join(__dirname, "preset-src.docx");
  execFileSync(process.execPath, [cli, mdPath, "-o", srcDocx,
    "--font", "黑体", "--font-size", "小四",
    "-p", "bottom", "--page-num-format", "第X页", "--page-num-start", "3",
    "--header", "测试页眉", "--align", "justify", "--line-height", "1.5",
  ], { stdio: "pipe" });
  const out1 = run([srcDocx, "--save-preset", "case7"]);
  check("提取保存成功并输出摘要", out1.includes("预设已保存") && out1.includes("提取结果"));
  const pj = JSON.parse(fs.readFileSync(path.join(tmpHome, "presets", "case7.json"), "utf-8"));
  check("提取-页面 A4 纵向", pj.options.page?.size === "A4" && pj.options.page?.orientation === "portrait");
  check("提取-页边距 2.54/3.18", pj.options.page?.margin?.top === 2.54 && pj.options.page?.margin?.left === 3.18);
  check("提取-正文黑体", pj.options.fonts?.body?.eastAsia === "黑体");
  check("提取-正文小四 12pt", pj.options.sizes?.body === 12);
  check("提取-行距 1.5", pj.options.paragraph?.line === 1.5);
  check("提取-首行缩进 2 字符", pj.options.paragraph?.firstLineChars === 2);
  check("提取-正文两端对齐", pj.options.paragraph?.align === "justify");
  check("提取-页眉文字与居中", pj.options.header?.text === "测试页眉" && pj.options.header?.align === "center");
  check("提取-页码格式「第X页」", pj.options.pageNumber?.format === "第X页");
  check("提取-页码位置与起始页", pj.options.pageNumber?.pos === "bottom" && pj.options.pageNumber?.start === 3);
  check("提取-预设文件含来源与时间元信息", pj.source === srcDocx && typeof pj.extractedAt === "string");

  // ---- 7.2 复用自定义预设转换 ----
  const reusedDocx = path.join(__dirname, "preset-reused.docx");
  run([mdPath, "-o", reusedDocx, "--preset", "case7"]);
  check("复用自定义预设转换成功", fs.existsSync(reusedDocx));
  const rz = await JSZip.loadAsync(fs.readFileSync(reusedDocx));
  const rDoc = await rz.file("word/document.xml").async("string");
  const rStyles = await rz.file("word/styles.xml").async("string");
  const rFooter = Object.keys(rz.files).find((f) => /^word\/footer\d+\.xml$/.test(f));
  const rFtr = rFooter ? await rz.file(rFooter).async("string") : "";
  check("复用-正文黑体", rStyles.includes('w:eastAsia="黑体"'));
  check("复用-正文小四（24 半磅）", rStyles.includes('w:val="24"'));
  check("复用-行距 1.5（line=360）", rStyles.includes('w:line="360"'));
  check("复用-两端对齐", rDoc.includes('w:val="both"'));
  check("复用-起始页码 3", rDoc.includes('w:start="3"'));
  check("复用-页码模板", rFtr.includes("第") && rFtr.includes("PAGE"));

  // ---- 7.3 sundy docx 全量往返（提取值与 sundy 预设逐项互逆） ----
  const sundyDocx = path.join(__dirname, "preset-sundy.docx");
  execFileSync(process.execPath, [cli, mdPath, "-o", sundyDocx, "--preset", "sundy"], { stdio: "pipe" });
  run([sundyDocx, "--save-preset", "sundycopy"]);
  const lj = JSON.parse(fs.readFileSync(path.join(tmpHome, "presets", "sundycopy.json"), "utf-8"));
  const lo = lj.options;
  check("sundy往返-正文仿宋/TNR 四号", lo.fonts?.body?.eastAsia === "仿宋" && lo.fonts?.body?.ascii === "Times New Roman" && lo.sizes?.body === 14);
  check("sundy往返-标题宋体 H1-H3 字号", lo.fonts?.heading?.eastAsia === "宋体" && JSON.stringify(lo.sizes?.heading) === JSON.stringify([22, 16, 14, 14, 14, 14]));
  check("sundy往返-H1 居中加粗 H4 不加粗", lo.heading?.align?.[0] === "center" && lo.heading?.bold?.[0] === true && lo.heading?.bold?.[3] === false);
  check("sundy往返-标题段前 0.5 行段后 0 行", lo.heading?.spacing?.beforeLines === 0.5 && lo.heading?.spacing?.afterLines === 0);
  check("sundy往返-段落三项（缩进2/行距1.28/段后1.5行）", lo.paragraph?.firstLineChars === 2 && lo.paragraph?.line === 1.28 && lo.paragraph?.afterLines === 1.5);
  check("sundy往返-页眉距顶端 0.85cm", lo.page?.margin?.header === 0.85);
  check("sundy往返-页眉三行左对齐", lo.header?.text?.split("\n").length === 3 && lo.header?.align === "left");
  check("sundy往返-页眉字号 9", lo.sizes?.header === 9);
  check("sundy往返-页脚字号 10.5", lo.sizes?.footer === 10.5);
  check("sundy往返-页码「第X页/共Y页」居中", lo.pageNumber?.format === "第X页/共Y页" && lo.pageNumber?.align === "center" && lo.pageNumber?.pos === "bottom");
  check("sundy往返-页眉图片已提取并落盘", typeof lo.header?.image?.path === "string" && fs.existsSync(lo.header.image.path));
  check("sundy往返-渐变图形已注明跳过", (lj.notes ?? []).some((n) => n.includes("图形")));

  // ---- 7.4 错误与边界 ----
  try { run([srcDocx, "--save-preset", "case7"]); check("预设重名报错提示 --overwrite", false); }
  catch (e) { check("预设重名报错提示 --overwrite", e.status === 1 && e.stderr.includes("--overwrite")); }
  const outOverwrite = run([srcDocx, "--save-preset", "case7", "--overwrite"]);
  check("--overwrite 可覆盖预设", outOverwrite.includes("预设已保存"));
  try { run([mdPath, "--save-preset", "bad"]); check(".md 文件禁止提取", false); }
  catch (e) { check(".md 文件禁止提取", e.status === 1 && e.stderr.includes("仅支持 .docx")); }
  try { run([srcDocx, "--save-preset", "a/b"]); check("非法预设名报错", false); }
  catch (e) { check("非法预设名报错", e.status === 1 && e.stderr.includes("不合法")); }
  try { run([mdPath, "--preset", "no-such-preset"]); check("未知预设列出自定义预设", false); }
  catch (e) { check("未知预设列出自定义预设", e.status === 1 && e.stderr.includes("自定义") && e.stderr.includes("case7")); }

  // ---- 7.5 非标准 styleId 的标题识别（w:name 兜底） ----
  const altDocx = path.join(__dirname, "preset-altstyleid.docx");
  execFileSync(process.execPath, [cli, mdPath, "-o", altDocx, "--preset", "sundy"], { stdio: "pipe" });
  {
    const z = await JSZip.loadAsync(fs.readFileSync(altDocx));
    const stylesXml = await z.file("word/styles.xml").async("string");
    // 把标准 styleId Heading1-6 改成非标准 ID，保留 w:name，验证提取按样式名兜底识别标题
    const altered = stylesXml.replace(/w:styleId="Heading([1-6])"/g, 'w:styleId="H$1Alt"');
    z.file("word/styles.xml", altered);
    fs.writeFileSync(altDocx, await z.generateAsync({ type: "nodebuffer" }));
  }
  run([altDocx, "--save-preset", "altstyle"]);
  const aj = JSON.parse(fs.readFileSync(path.join(tmpHome, "presets", "altstyle.json"), "utf-8"));
  check("非标准 styleId 兜底-标题字号已提取", Array.isArray(aj.options.sizes?.heading) && aj.options.sizes.heading[0] === 22 && aj.options.sizes.heading.length === 6);
  check("非标准 styleId 兜底-已注明按样式名识别", (aj.notes ?? []).some((n) => n.includes("样式名")));

  // 清理
  fs.rmSync(tmpHome, { recursive: true, force: true });
  for (const p of [srcDocx, reusedDocx, sundyDocx, altDocx]) fs.rmSync(p, { force: true });
}

// ============ 用例 8：图片超链接/标题、行距规则、格式二义性 ============
console.log("\n—— 用例 8：图片超链接/标题、行距规则、格式二义性 ——");
{
  const tmpHome8 = path.join(__dirname, "tmp-mddtt-home-8");
  fs.rmSync(tmpHome8, { recursive: true, force: true });
  const env8 = { ...process.env, MDDTT_HOME: tmpHome8 };
  const run8 = (args) => execFileSync(process.execPath, [cli, ...args], { stdio: "pipe", encoding: "utf-8", env: env8 });

  // ---- 8.1 链接图片 + 图片标题 ----
  const imgMd = path.join(__dirname, "img-link-title.md");
  const imgDocx = path.join(__dirname, "img-link-title.docx");
  fs.writeFileSync(imgMd, [
    "[![logo](assets/test-image.png)](https://example.com/x)",
    "",
    '![带标题](assets/test-image.png "这是图片标题")',
    "",
  ].join("\n"), "utf-8");
  execFileSync(process.execPath, [cli, imgMd, "-o", imgDocx], { stdio: "pipe" });
  const iz = await JSZip.loadAsync(fs.readFileSync(imgDocx));
  const iDoc = await iz.file("word/document.xml").async("string");
  check("图片超链接-超链接包裹图片", /<w:hyperlink[^>]*>\s*<w:r>\s*<w:drawing>/.test(iDoc));
  check("图片标题-标题文字已输出", iDoc.includes("这是图片标题"));
  const titleIdx = iDoc.indexOf("这是图片标题");
  check("图片标题-下方居中显示", iDoc.slice(iDoc.lastIndexOf("<w:p>", titleIdx), titleIdx).includes('w:val="center"'));

  // ---- 8.2 固定行距（exact）提取与还原 ----
  const exactDocx = path.join(__dirname, "exact-line.docx");
  execFileSync(process.execPath, [cli, mdPath, "-o", exactDocx, "--line-rule", "exact", "--line-height", "20"], { stdio: "pipe" });
  run8([exactDocx, "--save-preset", "exactline"]);
  const ej = JSON.parse(fs.readFileSync(path.join(tmpHome8, "presets", "exactline.json"), "utf-8"));
  check("固定行距-提取 line=20/lineRule=exact", ej.options.paragraph?.line === 20 && ej.options.paragraph?.lineRule === "exact");
  const exactReused = path.join(__dirname, "exact-line-reused.docx");
  run8([mdPath, "-o", exactReused, "--preset", "exactline"]);
  const ez = await JSZip.loadAsync(fs.readFileSync(exactReused));
  const eStyles = await ez.file("word/styles.xml").async("string");
  check("固定行距-复用还原 w:line=400/lineRule=exact", /w:line="400"[^>]*w:lineRule="exact"/.test(eStyles));

  // ---- 8.3 标题对齐二义性反馈 ----
  const { Document, Packer, Paragraph, HeadingLevel, AlignmentType, TextRun } = await import("docx");
  const ambDocx = path.join(__dirname, "ambiguous.docx");
  const ambDoc = new Document({
    sections: [{
      children: [
        new Paragraph({ heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER, children: [new TextRun("标题A")] }),
        new Paragraph({ heading: HeadingLevel.HEADING_1, alignment: AlignmentType.LEFT, children: [new TextRun("标题B")] }),
        new Paragraph("正文段落"),
      ],
    }],
  });
  fs.writeFileSync(ambDocx, await Packer.toBuffer(ambDoc));
  run8([ambDocx, "--save-preset", "amb"]);
  const ambJ = JSON.parse(fs.readFileSync(path.join(tmpHome8, "presets", "amb.json"), "utf-8"));
  check("二义性-检测到 H1 对齐不一致", (ambJ.notes ?? []).some((n) => n.includes("H1") && n.includes("对齐不一致") && n.includes("居中") && n.includes("左对齐")));

  // 清理
  fs.rmSync(tmpHome8, { recursive: true, force: true });
  for (const p of [imgMd, imgDocx, exactDocx, exactReused, ambDocx]) fs.rmSync(p, { force: true });
}

// ============ 用例 9：换行识别（breaks / --no-breaks） ============
console.log("\n—— 用例 9：换行识别 ——");
{
  const hbMd = path.join(__dirname, "hardbreak.md");
  const hbDocx = path.join(__dirname, "hardbreak.docx");
  const hbNoBreaksDocx = path.join(__dirname, "hardbreak-nobreaks.docx");
  const readDocumentXml = async (p) => {
    const z = await JSZip.loadAsync(fs.readFileSync(p));
    return z.file("word/document.xml").async("string");
  };
  const paragraphsOf = (xml) => xml.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g) || [];
  for (const p of [hbMd, hbDocx, hbNoBreaksDocx]) fs.rmSync(p, { force: true });

  // 单换行（中文文档每行一换行的常见写法）、双空格换行、空行分段
  fs.writeFileSync(hbMd, [
    "单换行甲行",
    "单换行乙行",
    "",
    "双空格甲行  ",
    "双空格乙行",
    "",
    "多行甲行",
    "多行乙行",
    "多行丙行",
    "",
    "新段落",
    "",
  ].join("\n"), "utf-8");

  execFileSync(process.execPath, [cli, hbMd, "-o", hbDocx], { stdio: "pipe" });
  const hbXml = await readDocumentXml(hbDocx);
  const hbParas = paragraphsOf(hbXml);
  check("换行-单换行生成段内换行（w:br）", (hbParas[0].match(/<w:br\/>/g) || []).length === 1 && hbParas[0].includes("单换行甲行") && hbParas[0].includes("单换行乙行"));
  check("换行-换行后补两字符缩进（全角空格×2）", /<w:br\/><\/w:r><w:r><w:t xml:space="preserve">\u3000\u3000<\/w:t><\/w:r>/.test(hbParas[0]));
  check("换行-双空格硬换行仍在同一段", (hbParas[1].match(/<w:br\/>/g) || []).length === 1 && hbParas[1].includes("双空格甲行") && hbParas[1].includes("双空格乙行"));
  check("换行-连续单换行逐行保留", (hbParas[2].match(/<w:br\/>/g) || []).length === 2 && hbParas[2].includes("多行甲行") && hbParas[2].includes("多行丙行"));
  check("换行-空行仍分为独立段落", hbParas.length === 4 && hbParas[3].includes("新段落"));

  execFileSync(process.execPath, [cli, hbMd, "-o", hbNoBreaksDocx, "--no-breaks"], { stdio: "pipe" });
  const nbXml = await readDocumentXml(hbNoBreaksDocx);
  check("换行---no-breaks 时单换行不生成 w:br", !paragraphsOf(nbXml)[0].includes("<w:br/>") && paragraphsOf(nbXml)[0].includes("单换行甲行") && paragraphsOf(nbXml)[0].includes("单换行乙行"));

  // settings.xml：关闭"扩展手动换行符结尾行的对齐"，避免两端对齐时把各短行撑满整行
  const hbZip = await JSZip.loadAsync(fs.readFileSync(hbDocx));
  const settings = await hbZip.file("word/settings.xml").async("string");
  const compat = /<w:compat>[\s\S]*?<\/w:compat>/.exec(settings);
  check("换行-写入 doNotExpandShiftReturn", settings.includes("<w:doNotExpandShiftReturn/>"));
  check("换行-compatSetting 顺序合规（兼容性开关在前）",
    compat && compat[0].indexOf("<w:doNotExpandShiftReturn/>") < compat[0].indexOf("<w:compatSetting"));

  for (const p of [hbMd, hbDocx, hbNoBreaksDocx]) fs.rmSync(p, { force: true });
}

// ============ 用例 10：docx 富结构降级（修订/批注/文本框/浮动对象/分节） ============
console.log("\n—— 用例 10：docx 富结构降级 ——");
{
  const richDocx = path.join(__dirname, "rich-import.docx");
  const richMd = path.join(__dirname, "rich-import.md");
  const zip = await JSZip.loadAsync(fs.readFileSync(path.join(__dirname, "sample-default.docx")));
  let documentXml = await zip.file("word/document.xml").async("string");

  const richBlocks = [
    '<w:p><w:del w:id="1" w:author="测试人员" w:date="2026-09-07T00:00:00Z"><w:r><w:delText>修订前文本</w:delText></w:r></w:del><w:ins w:id="2" w:author="测试人员" w:date="2026-09-07T00:01:00Z"><w:r><w:t>修订后文本</w:t></w:r></w:ins></w:p>',
    '<w:p><w:commentRangeStart w:id="7"/><w:r><w:t>ABCDEFGHIJKLMNOPQRSTUVWXYZ</w:t></w:r><w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:p>',
    '<w:p><w:r><w:pict><v:shape id="TextBox1"><v:textbox><w:txbxContent><w:p><w:r><w:t>文本框第一段</w:t></w:r></w:p><w:p><w:r><w:t>文本框第二段</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>',
    '<w:p><w:pPr><w:sectPr><w:cols w:num="2" w:space="720"/></w:sectPr></w:pPr><w:r><w:t>第一节末尾内容</w:t></w:r></w:p>',
    '<w:p><w:r><w:t>第二节线性内容</w:t></w:r></w:p>',
  ].join("");
  documentXml = documentXml.replace("<w:sectPr", `${richBlocks}<w:sectPr`);
  documentXml = documentXml
    .replace("<wp:inline ", "<wp:anchor ")
    .replace("</wp:inline>", "</wp:anchor>")
    .replace('descr="测试图片"', 'descr="浮动图片测试"');
  zip.file("word/document.xml", documentXml);
  zip.file("word/comments.xml", [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    '<w:comment w:id="7" w:author="批注人" w:date="2026-09-07T00:02:00Z">',
    '<w:p><w:r><w:t>请核对此处条款</w:t></w:r></w:p>',
    '</w:comment></w:comments>',
  ].join(""));
  fs.writeFileSync(richDocx, await zip.generateAsync({ type: "nodebuffer" }));

  execFileSync(process.execPath, [cli, richDocx, "-o", richMd, "--overwrite"], { stdio: "pipe" });
  const markdown = fs.readFileSync(richMd, "utf-8");
  check("修订-采用修订后文本", markdown.includes("修订后文本") && !markdown.includes("修订前文本修订后文本"));
  check("修订-紧接注释并记录修订前原文", markdown.includes("修订后文本<!-- 此处系修订；修订前原文：修订前文本 -->"));
  check("批注-注释紧接被批注内容", markdown.includes("ABCDEFGHIJKLMNOPQRSTUVWXYZ<!-- 批注范围："));
  check("批注-长范围保留首尾并省略中间", markdown.includes("ABCDEFGHIJ……QRSTUVWXYZ"));
  check("批注-注释包含批注正文", markdown.includes("此处有批注：请核对此处条款"));
  check("文本框-内容按普通段落线性输出", markdown.indexOf("文本框第一段") >= 0 && markdown.indexOf("文本框第一段") < markdown.indexOf("文本框第二段"));
  check("浮动图片-降级为普通 Markdown 图片", markdown.includes("![浮动图片测试](MDPictures/"));
  check("多栏多节-内容按源顺序线性排列", markdown.indexOf("第一节末尾内容") >= 0 && markdown.indexOf("第一节末尾内容") < markdown.indexOf("第二节线性内容"));

  for (const file of [richDocx, richMd]) fs.rmSync(file, { force: true });
}

console.log(failed ? `\n${failed} 项校验未通过` : "\n全部校验通过");
process.exit(failed ? 1 : 0);
