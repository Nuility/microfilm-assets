import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const { SpreadsheetFile, Workbook } = require("@oai/artifact-tool");

const here = path.dirname(fileURLToPath(import.meta.url));
const outDir = path.dirname(here);
const jsonPath = path.join(outDir, "《把答案写在大地上》镜头音乐数据.json");
const outputPath = path.join(outDir, "《把答案写在大地上》镜头-音乐对应关系表.xlsx");
const data = JSON.parse(await fs.readFile(jsonPath, "utf8"));

const wb = Workbook.create();
const fontFamily = "Microsoft YaHei";
const navy = "#1F4E78";
const midBlue = "#5B9BD5";
const paleBlue = "#EAF2F8";
const paleGold = "#FFF4CC";
const paleGreen = "#E2F0D9";
const paleRed = "#FCE4D6";
const border = "#D9D9D9";
const textColor = "#1F1F1F";

function colName(n) {
  let s = "";
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

function setupSheet(sheet, title, subtitle, columns, rows, widths, tabColor = null) {
  sheet.showGridLines = false;
  if (tabColor) sheet.tabColor = tabColor;
  const last = colName(columns.length);
  sheet.getRange(`A2:${last}2`).format.borders = { preset: "doubleBottom", style: "thin", color: navy };
  sheet.getRange("A2").values = [[title]];
  sheet.getRange("A2").format.font = { name: fontFamily, size: 15, bold: true, color: "#000000" };
  sheet.getRange("A3").values = [[subtitle]];
  sheet.getRange("A3").format.font = { name: fontFamily, size: 10, italic: true, color: "#555555" };
  sheet.getRange(`A5:${last}5`).values = [columns];
  sheet.getRange(`A5:${last}5`).format = {
    fill: navy,
    font: { name: fontFamily, size: 10, bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#FFFFFF" },
  };
  if (rows.length) {
    const endRow = 5 + rows.length;
    sheet.getRange(`A6:${last}${endRow}`).values = rows;
    sheet.getRange(`A6:${last}${endRow}`).format.font = { name: fontFamily, size: 9, color: textColor };
    sheet.getRange(`A6:${last}${endRow}`).format.verticalAlignment = "center";
    sheet.getRange(`A6:${last}${endRow}`).format.wrapText = true;
    sheet.getRange(`A6:${last}${endRow}`).format.borders = { preset: "all", style: "thin", color: border };
    for (let r = 6; r <= endRow; r++) {
      if ((r - 6) % 2 === 1) sheet.getRange(`A${r}:${last}${r}`).format.fill = "#F7F9FB";
    }
  }
  widths.forEach((w, i) => {
    sheet.getRange(`${colName(i + 1)}:${colName(i + 1)}`).format.columnWidth = w;
  });
  sheet.getRange("1:1").format.rowHeight = 8;
  sheet.getRange("2:2").format.rowHeight = 26;
  sheet.getRange("3:3").format.rowHeight = 22;
  sheet.getRange("5:5").format.rowHeight = 34;
  sheet.freezePanes.freezeRows(5);
  return 5 + rows.length;
}

const mapSheet = wb.worksheets.add("镜头音乐映射");
const mapCols = ["镜头", "时间", "画面", "视觉情绪", "人物状态", "剧情作用", "音乐名称", "来源", "是否需要音乐", "音乐类型", "音乐情绪", "进入点", "退出点", "音量", "音乐作用", "音效重点", "制作状态"];
const mapRows = data.shots.map(s => [
  s["镜头"], s["时间"], s["剧情内容"], s["视觉情绪"], s["人物状态"], s["剧情作用"], s["音乐名称"], s["来源"], s["是否需要音乐"], s["需要类型"], s["音乐情绪"], s["音乐进入时间"], s["音乐结束时间"], s["音量"], s["音乐作用"], s["音效重点"], "待制作"
]);
const mapEnd = setupSheet(
  mapSheet,
  "《把答案写在大地上》镜头音乐映射",
  "6分00秒 · 36镜头 · AI原创六Cue主方案 · 镜头33禁止叠加外部配乐",
  mapCols,
  mapRows,
  [7, 13, 29, 20, 22, 23, 12, 16, 10, 28, 17, 12, 12, 22, 30, 28, 12],
  navy
);
mapSheet.getRange(`A6:B${mapEnd}`).format.horizontalAlignment = "center";
mapSheet.getRange(`A6:A${mapEnd}`).format.numberFormat = "00";
mapSheet.getRange(`G6:I${mapEnd}`).format.horizontalAlignment = "center";
mapSheet.getRange(`L6:M${mapEnd}`).format.horizontalAlignment = "center";
mapSheet.getRange(`Q6:Q${mapEnd}`).format.fill = paleGold;
mapSheet.getRange(`Q6:Q${mapEnd}`).dataValidation = { rule: { type: "list", values: ["待制作", "已生成", "已剪辑", "已混音", "已锁定"] } };
mapSheet.getRange(`Q6:Q${mapEnd}`).conditionalFormats.add("containsText", { text: "已锁定", format: { fill: paleGreen, font: { color: "#375623", bold: true } } });
mapSheet.getRange(`Q6:Q${mapEnd}`).conditionalFormats.add("containsText", { text: "待制作", format: { fill: paleGold, font: { color: "#7F6000" } } });
mapSheet.getRange("A1").values = [["默认：先制作M01-M06；版权曲只作整段替代，不混搭。"]];
mapSheet.getRange("A1").format.font = { name: fontFamily, size: 9, italic: true, color: "#666666" };

const cueSheet = wb.worksheets.add("Cue总览");
const cueCols = ["编号", "名称", "适用镜头", "时长", "BPM", "调性", "结构", "制作状态", "文件名建议"];
const cueRows = data.cues.map(c => [c["编号"], c["名称"], c["适用镜头"], c["时长"], c["BPM"], c["调性"], c["结构"], "待生成", `${c["编号"]}_${c["名称"]}_v01_48k24.wav`]);
const cueEnd = setupSheet(cueSheet, "六首原创Cue总览", "六首共享同一四音问号动机；主混与分轨均为48kHz/24bit WAV", cueCols, cueRows, [9, 16, 13, 20, 14, 18, 55, 12, 30], midBlue);
cueSheet.getRange(`A6:F${cueEnd}`).format.horizontalAlignment = "center";
cueSheet.getRange(`H6:H${cueEnd}`).format.fill = paleGold;
cueSheet.getRange(`H6:H${cueEnd}`).dataValidation = { rule: { type: "list", values: ["待生成", "初版", "修改中", "通过", "锁定"] } };

const promptSheet = wb.worksheets.add("AI音乐Prompt");
const promptCols = ["编号", "名称", "适用镜头", "BPM", "调性", "结构", "完整生成Prompt"];
const promptRows = data.cues.map(c => [c["编号"], c["名称"], c["适用镜头"], c["BPM"], c["调性"], c["结构"], c["Prompt"]]);
const promptEnd = setupSheet(promptSheet, "AI原创音乐生成Prompt", "使用同一平台、同一音色参考和同一四音主题种子；不得模仿具体作品或作曲家", promptCols, promptRows, [9, 16, 14, 14, 20, 55, 110], "#70AD47");
promptSheet.getRange(`A6:E${promptEnd}`).format.horizontalAlignment = "center";
promptSheet.getRange(`F6:G${promptEnd}`).format.verticalAlignment = "top";

const rightsSheet = wb.worksheets.add("版权音乐资源");
const rightsCols = ["音乐名称", "作者", "推荐平台", "官方链接", "音乐类型", "适用镜头", "版权状态", "商业使用要求", "采用决定", "凭证归档"];
const rightsRows = data.resources.map(r => [r["名称"], r["作者"], r["平台"], r["链接"], r["类型"], r["适用镜头"], r["版权状态"], r["商业使用要求"], "备选", "未归档"]);
const rightsEnd = setupSheet(rightsSheet, "版权音乐资源推荐", "2026-09-18按作者官方曲目页核验 · CC BY 4.0 · 正式采用前保存网页、下载文件和署名文本", rightsCols, rightsRows, [20, 18, 34, 55, 38, 30, 32, 75, 12, 14], "#A5A5A5");
rightsSheet.getRange(`A6:B${rightsEnd}`).format.horizontalAlignment = "center";
rightsSheet.getRange(`I6:J${rightsEnd}`).format.fill = paleGold;
rightsSheet.getRange(`I6:I${rightsEnd}`).dataValidation = { rule: { type: "list", values: ["备选", "采用", "淘汰"] } };
rightsSheet.getRange(`J6:J${rightsEnd}`).dataValidation = { rule: { type: "list", values: ["未归档", "已归档"] } };
rightsSheet.getRange("A1").values = [["统一署名：'[Track Title]' by Scott Buckley – released under CC-BY 4.0. www.scottbuckley.com.au"]];
rightsSheet.getRange("A1").format.font = { name: fontFamily, size: 9, italic: true, color: "#666666" };

const sfxSheet = wb.worksheets.add("音效设计");
const sfxCols = ["音效名称", "对应镜头", "来源", "是否需要生成", "执行说明", "录制/素材状态", "文件名/素材ID"];
const sfxRows = data.sfx.map(x => [x["音效名称"], x["对应镜头"], x["来源"], x["是否需要生成"], x["执行"], "待录制", ""]);
const sfxEnd = setupSheet(sfxSheet, "音效设计方案", "生活实声优先；钥匙、笔记本、木门、稳定器和投影设备必须保持统一声音身份", sfxCols, sfxRows, [32, 18, 40, 16, 65, 16, 28], "#ED7D31");
sfxSheet.getRange(`B6:B${sfxEnd}`).format.horizontalAlignment = "center";
sfxSheet.getRange(`D6:D${sfxEnd}`).format.horizontalAlignment = "center";
sfxSheet.getRange(`F6:G${sfxEnd}`).format.fill = paleGold;
sfxSheet.getRange(`F6:F${sfxEnd}`).dataValidation = { rule: { type: "list", values: ["待录制", "已录制", "待补录", "已编辑", "已锁定"] } };

const checkSheet = wb.worksheets.add("制作检查");
const checkCols = ["检查项", "当前结论", "说明", "复核人", "复核日期"];
const checkRows = data.checks.map(x => [x[0], x[1], x[2], "", ""]);
const checkEnd = setupSheet(checkSheet, "音乐制作检查报告", "交付前逐项复核；版权与技术交付仍需在实际制作时关闭", checkCols, checkRows, [28, 18, 80, 18, 16], "#8064A2");
checkSheet.getRange(`B6:B${checkEnd}`).format.horizontalAlignment = "center";
checkSheet.getRange(`D6:E${checkEnd}`).format.fill = paleGold;
checkSheet.getRange(`B6:B${checkEnd}`).conditionalFormats.add("containsText", { text: "通过", format: { fill: paleGreen, font: { color: "#375623", bold: true } } });
checkSheet.getRange(`B6:B${checkEnd}`).conditionalFormats.add("containsText", { text: "待", format: { fill: paleGold, font: { color: "#7F6000" } } });
checkSheet.getRange(`B6:B${checkEnd}`).conditionalFormats.add("containsText", { text: "风险", format: { fill: paleRed, font: { color: "#9C0006", bold: true } } });

wb.recalculate();

const inspectMap = await wb.inspect({ kind: "table", sheetId: "镜头音乐映射", range: `A1:Q${mapEnd}`, include: "values,formulas", tableMaxRows: 40, tableMaxCols: 18, maxChars: 18000 });
console.log(inspectMap.ndjson);
const errors = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 100 }, summary: "final formula error scan" });
console.log(errors.ndjson);

for (const sheetName of ["镜头音乐映射", "Cue总览", "AI音乐Prompt", "版权音乐资源", "音效设计", "制作检查"]) {
  const preview = await wb.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(path.join(here, `preview_${sheetName}.png`), new Uint8Array(await preview.arrayBuffer()));
}

const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(outputPath);
console.log(outputPath);
