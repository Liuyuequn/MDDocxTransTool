// 预设方案：每个预设是与 defaultOptions 同形的补丁对象
// sundy 为圣典法律文书规范（已锁定）；默认转换（无预设）与其排版相同，仅无页眉页脚页码

import path from "node:path";
import { fileURLToPath } from "node:url";

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

export const presets = {
  // 圣典法律文书（原 legal，已锁定）
  sundy: {
    page: {
      size: "A4",
      orientation: "portrait",
      // 页眉距顶端 0.85cm：将整组页眉内容上移，为渐变分割线与正文留出少许空间
      margin: { top: 2.54, right: 3.18, bottom: 2.54, left: 3.18, header: 0.85 },
    },
    fonts: {
      body: { eastAsia: "仿宋", ascii: "Times New Roman" },
      heading: { eastAsia: "宋体", ascii: "Times New Roman" },
      header: { eastAsia: "仿宋", ascii: "Times New Roman" },
      footer: { eastAsia: "仿宋", ascii: "Times New Roman" },
    },
    sizes: {
      body: 14,                          // 四号
      heading: [22, 16, 14, 14, 14, 14], // 二号/三号/四号×4
      header: 9,                         // 页眉小五
      footer: 10.5,                       // 页脚页码五号
    },
    heading: {
      bold: [true, true, true, false, false, false],
      align: ["center", "left", "left", "left", "left", "left"],
      spacing: { beforeLines: 0.5, afterLines: 0 },
    },
    paragraph: {
      firstLineChars: 2,              // 首行缩进两字符
      afterLines: 0.5,                // 空行分段后的段后距 0.5 行
      line: 1.28,                     // 行距 1.28 行
    },
    header: {
      text: "圣典律师事务所\n圣典官网：https://www.sundylawyer.com/\n总所地址：南京市建邺区奥体大街68号新城科技园4A栋6楼、7楼",
      align: "left",
      image: { path: path.join(pkgRoot, "reference", "圣典律师.png"), heightCm: 1.2 },
      // 页眉底端渐变色带：亮正红 → 红橙 → 橙 → 金，橙金从 45% 起即明显（占比约 55%）
      gradientLine: {
        stops: [[0, "E60012"], [0.45, "FF5A00"], [0.7, "FFA500"], [1, "FFD700"]],
        heightCm: 0.06,
        gapCm: 0.2, // 色带与上方页眉文字之间的间距
      },
    },
    pageNumber: {
      pos: "bottom",
      align: "center",
      format: "第X页/共Y页",
    },
  },
};
