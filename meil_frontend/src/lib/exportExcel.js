import * as XLSX from "xlsx";

/**
 * Export data to a styled Excel file.
 * @param {string} sheetName - Name of the sheet tab
 * @param {string[]} headers - Column header labels
 * @param {any[][]} rows - Array of row arrays (values in same order as headers)
 * @param {string} filename - Filename without extension
 */
export function exportToExcel(sheetName, headers, rows, filename) {
  const wb = XLSX.utils.book_new();

  // Title row + header row + data rows
  const titleRow = [`${sheetName} - Exported on ${new Date().toLocaleDateString("en-IN")}`];
  const wsData = [titleRow, headers, ...rows];

  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // Merge title row across all columns
  ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } }];

  // Auto column widths: measure max content length per column
  const colWidths = headers.map((h, colIdx) => {
    const maxLen = Math.max(
      String(h).length,
      ...rows.map((r) => String(r[colIdx] ?? "").length)
    );
    return { wch: Math.min(Math.max(maxLen + 2, 12), 60) };
  });
  ws["!cols"] = colWidths;

  // Apply styles to title, header, and data cells
  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cellAddr = XLSX.utils.encode_cell({ r: R, c: C });
      if (!ws[cellAddr]) continue;

      if (R === 0) {
        // Title row
        ws[cellAddr].s = {
          font: { bold: true, sz: 13, color: { rgb: "FFFFFF" } },
          fill: { fgColor: { rgb: "1E3A5F" } },
          alignment: { horizontal: "center", vertical: "center" },
        };
      } else if (R === 1) {
        // Header row
        ws[cellAddr].s = {
          font: { bold: true, sz: 10, color: { rgb: "FFFFFF" } },
          fill: { fgColor: { rgb: "2563EB" } },
          alignment: { horizontal: "center", vertical: "center", wrapText: true },
          border: {
            bottom: { style: "medium", color: { rgb: "1E40AF" } },
          },
        };
      } else {
        // Data rows — alternate row shading
        const isEven = R % 2 === 0;
        ws[cellAddr].s = {
          font: { sz: 9 },
          fill: { fgColor: { rgb: isEven ? "EFF6FF" : "FFFFFF" } },
          alignment: { vertical: "center", wrapText: false },
          border: {
            bottom: { style: "thin", color: { rgb: "DBEAFE" } },
          },
        };
      }
    }
  }

  // Row height: title=24, header=18, data=15
  ws["!rows"] = [
    { hpt: 24 },
    { hpt: 18 },
    ...rows.map(() => ({ hpt: 15 })),
  ];

  XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31));

  const date = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `${filename}_${date}.xlsx`);
}
