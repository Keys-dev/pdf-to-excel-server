import ExcelJS from "exceljs";
import type { ExtractedTable } from "../types.js";

export async function tablesToExcelBuffer(tables: ExtractedTable[]): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "PDF to Office";
  workbook.created = new Date();

  if (tables.length === 0) {
    workbook.addWorksheet("Sheet1");
  }

  tables.forEach((table, index) => {
    const sheet = workbook.addWorksheet(`Table ${index + 1} (p${table.page})`);
    table.rows.forEach((row) => sheet.addRow(row));

    sheet.getRow(1).font = { bold: true };
    sheet.columns.forEach((column) => {
      let maxLength = 10;
      column.eachCell?.({ includeEmpty: true }, (cell) => {
        const length = cell.value ? String(cell.value).length : 0;
        if (length > maxLength) maxLength = length;
      });
      column.width = Math.min(maxLength + 2, 60);
    });
  });

  const arrayBuffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}
