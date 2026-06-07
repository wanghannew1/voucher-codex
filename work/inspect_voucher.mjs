import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "/Users/sunyitong/coding/voucher/代码资料/凭证信息_temp - 2026-06-02T140354.252.xlsx";
const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

console.log(JSON.stringify({ sheetNames: workbook.worksheets.map(ws => ws.name) }, null, 2));

for (const ws of workbook.worksheets) {
  const used = ws.usedRange?.address ?? "unknown";
  console.log(`SHEET ${ws.name} USED ${used}`);
  const previewRange = `${ws.name}!A1:Z12`;
  const preview = await workbook.inspect({
    kind: "table",
    range: previewRange,
    include: "values,formulas",
    tableMaxRows: 12,
    tableMaxCols: 26,
  });
  console.log(preview.ndjson);
}
