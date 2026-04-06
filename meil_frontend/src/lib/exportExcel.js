import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

/**
 * Export data to a fully styled Excel file using ExcelJS.
 * @param {string} sheetName - Sheet tab name
 * @param {string[]} headers  - Column header labels (will be uppercased)
 * @param {any[][]} rows      - Data rows (same order as headers)
 * @param {string} filename   - Filename without extension
 */
export async function exportToExcel(sheetName, headers, rows, filename) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "MEIL MDM Portal";
  wb.created = new Date();

  const ws = wb.addWorksheet(sheetName.slice(0, 31));

  const colCount = headers.length;

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  ws.mergeCells(1, 1, 1, colCount);
  const titleCell = ws.getCell("A1");
  titleCell.value = `${sheetName.toUpperCase()} — Exported on ${new Date().toLocaleDateString("en-IN")}`;
  titleCell.font  = { bold: true, size: 13, color: { argb: "FFFFFFFF" } };
  titleCell.fill  = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E3A5F" } };
  titleCell.alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(1).height = 26;

  // ── Row 2: Headers ────────────────────────────────────────────────────────
  const headerRow = ws.addRow(headers.map(h => h.toUpperCase()));
  headerRow.height = 20;
  headerRow.eachCell(cell => {
    cell.value    = String(cell.value).toUpperCase();
    cell.font     = { bold: true, size: 10, color: { argb: "FFFFFFFF" } };
    cell.fill     = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2563EB" } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: false };
    cell.border   = {
      bottom: { style: "medium", color: { argb: "FF1E40AF" } },
      right:  { style: "thin",   color: { argb: "FF93C5FD" } },
    };
  });

  // ── Rows 3+: Data ─────────────────────────────────────────────────────────
  rows.forEach((rowData, i) => {
    const row = ws.addRow(rowData);
    const isEven = i % 2 === 0;
    row.height = 16;
    row.eachCell({ includeEmpty: true }, cell => {
      cell.font      = { size: 9 };
      cell.fill      = {
        type: "pattern", pattern: "solid",
        fgColor: { argb: isEven ? "FFEFF6FF" : "FFFFFFFF" },
      };
      cell.alignment = { vertical: "middle" };
      cell.border    = {
        bottom: { style: "thin", color: { argb: "FFDBEAFE" } },
      };
    });
  });

  // ── Auto column widths ────────────────────────────────────────────────────
  ws.columns.forEach((col, i) => {
    const maxLen = Math.max(
      headers[i].length,
      ...rows.map(r => String(r[i] ?? "").length)
    );
    col.width = Math.min(Math.max(maxLen + 3, 14), 55);
  });

  // ── Freeze header rows ────────────────────────────────────────────────────
  ws.views = [{ state: "frozen", ySplit: 2 }];

  // ── Write & download ──────────────────────────────────────────────────────
  const buffer = await wb.xlsx.writeBuffer();
  const blob   = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const date = new Date().toISOString().slice(0, 10);
  saveAs(blob, `${filename}_${date}.xlsx`);
}
