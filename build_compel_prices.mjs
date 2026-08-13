import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const csvPath = "compel_cart_1779642_prices.csv";
const htmlPath = "compel_cart_1779642.html";
const outputDir = "outputs/compel_cart_1779642";
const outputPath = `${outputDir}/compel_cart_1779642_prices.xlsx`;

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    const next = text[i + 1];
    if (quoted) {
      if (ch === '"' && next === '"') {
        cell += '"';
        i += 1;
      } else if (ch === '"') {
        quoted = false;
      } else {
        cell += ch;
      }
    } else if (ch === '"') {
      quoted = true;
    } else if (ch === ";") {
      row.push(cell);
      cell = "";
    } else if (ch === "\n") {
      row.push(cell.replace(/\r$/, ""));
      rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += ch;
    }
  }
  if (cell.length || row.length) {
    row.push(cell.replace(/\r$/, ""));
    rows.push(row);
  }
  return rows.filter((r) => r.some((v) => v !== ""));
}

function asNumber(value) {
  if (!value) return null;
  const n = Number(String(value).replace(",", "."));
  return Number.isFinite(n) ? n : null;
}

function price2Entries(html) {
  const entries = [];
  const re = /<price2\s+val="([^"]+)"\s+cur="([^"]+)"[^>]*>\s*<span\s+data-title="([^"]*)"/gs;
  for (const match of html.matchAll(re)) {
    const rur = /([\d\s.,]+)\s*RUR/.exec(match[3]);
    entries.push({
      usd: asNumber(match[1]),
      cur: match[2],
      rur: rur ? asNumber(rur[1].replace(/\s/g, "")) : null,
    });
  }
  return entries;
}

const csvText = await fs.readFile(csvPath, "utf8");
const htmlText = await fs.readFile(htmlPath, "utf8");
const csvRows = parseCsv(csvText.replace(/^\uFEFF/, ""));
const headers = csvRows[0];
const sourceRows = csvRows.slice(1);

const rows = sourceRows.map((r) => [
  asNumber(r[0]),
  r[1],
  r[2],
  r[3],
  asNumber(r[4]),
  r[5],
  asNumber(r[6]),
  asNumber(r[7]),
  r[8],
  asNumber(r[9]),
  asNumber(r[10]),
  r[11],
  asNumber(r[12]),
]);

const missingRows = rows.filter((r) => r[10] === null).map((r) => r[0]).join(", ");
const tailPrices = price2Entries(htmlText).slice(-2);
const cheapTotal = tailPrices[0] ?? { usd: null, rur: null };
const optimalTotal = tailPrices[1] ?? { usd: null, rur: null };

await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();
const summary = workbook.worksheets.add("Сводка");
const details = workbook.worksheets.add("Позиции");

details.getRange(`A1:M${rows.length + 1}`).values = [headers, ...rows];
details.getRange("A1:M1").format = {
  fill: "#1F4E79",
  font: { bold: true, color: "#FFFFFF" },
};
details.getRange(`A2:A${rows.length + 1}`).format.numberFormat = "0";
details.getRange(`E2:E${rows.length + 1}`).format.numberFormat = "0";
details.getRange(`G2:G${rows.length + 1}`).format.numberFormat = "0";
details.getRange(`H2:H${rows.length + 1}`).format.numberFormat = "0.00000";
details.getRange(`J2:J${rows.length + 1}`).format.numberFormat = "#,##0.00";
details.getRange(`K2:K${rows.length + 1}`).format.numberFormat = "#,##0.00";
details.getRange(`M2:M${rows.length + 1}`).format.numberFormat = "#,##0.00";
details.getRange("A:A").format.columnWidthPx = 46;
details.getRange("B:C").format.columnWidthPx = 190;
details.getRange("D:D").format.columnWidthPx = 130;
details.getRange("E:G").format.columnWidthPx = 90;
details.getRange("H:M").format.columnWidthPx = 118;
details.getRange("A1:M1").format.wrapText = true;

summary.getRange("A1:D1").values = [["Расчет SDS Compel", "", "", ""]];
summary.getRange("A1:D1").merge();
summary.getRange("A1:D1").format = {
  fill: "#1F4E79",
  font: { bold: true, color: "#FFFFFF", size: 14 },
};

summary.getRange("A3:D11").values = [
  ["Показатель", "USD", "RUR справочно", "Комментарий"],
  ["Итого строк с подобранной ценой", null, null, "Сумма по листу Позиции"],
  ["Итог сайта: оптимизировано по цене", cheapTotal.usd, cheapTotal.rur, "30 дн"],
  ["Итог сайта: оптимизировано по сроку", optimalTotal.usd, optimalTotal.rur, "24 дн"],
  ["Всего строк", rows.length, null, ""],
  ["Строк с ценой", null, null, ""],
  ["Строк без подобранного предложения", null, null, missingRows],
  ["Источник", "", "", "sds.compel.ru/user/carts/1779642"],
  ["Дата сохраненного HTML", "04.06.2026 16:53", "", ""],
];
summary.getRange("B4").formulas = [[`=SUM('Позиции'!K2:K${rows.length + 1})`]];
summary.getRange("C4").formulas = [[`=SUM('Позиции'!M2:M${rows.length + 1})`]];
summary.getRange("B8").formulas = [[`=COUNT('Позиции'!K2:K${rows.length + 1})`]];
summary.getRange("B9").formulas = [[`=B7-B8`]];

summary.getRange("A3:D3").format = {
  fill: "#D9EAF7",
  font: { bold: true },
};
summary.getRange("A4:A11").format = { font: { bold: true } };
summary.getRange("B4:C7").format.numberFormat = "#,##0.00";
summary.getRange("B8:B10").format.numberFormat = "0";
summary.getRange("A:A").format.columnWidthPx = 250;
summary.getRange("B:C").format.columnWidthPx = 125;
summary.getRange("D:D").format.columnWidthPx = 270;
summary.getRange("A3:D11").format.wrapText = true;
summary.getRange("A3:D11").format.rowHeightPx = 30;

const inspect = await workbook.inspect({
  kind: "table",
  sheetId: "Сводка",
  range: "A3:D11",
  include: "values,formulas",
  maxChars: 4000,
});
console.log(inspect);

for (const sheetName of ["Сводка", "Позиции"]) {
  const preview = await workbook.render({
    sheetName,
    autoCrop: "all",
    scale: 1,
    format: "png",
  });
  await fs.writeFile(
    `${outputDir}/${sheetName}.png`,
    new Uint8Array(await preview.arrayBuffer()),
  );
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`saved ${outputPath}`);
