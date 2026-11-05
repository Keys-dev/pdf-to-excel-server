import ExcelJS from "exceljs";
/**
 * Export layout cells to Excel as a compact table.
 * Uses row/col indices from the PDF extractor when available, falling back to
 * PDF position-based grid mapping when row/col is not present.
 */
export async function layoutToExcelBuffer(cells) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "PDF to Office";
    workbook.created = new Date();
    if (cells.length === 0) {
        workbook.addWorksheet("Sheet1");
    }
    // Group cells by page
    const byPage = new Map();
    for (const cell of cells) {
        const existing = byPage.get(cell.page) || [];
        existing.push(cell);
        byPage.set(cell.page, existing);
    }
    for (const [page, pageCells] of byPage) {
        const sheet = workbook.addWorksheet(`Page ${page}`);
        // Check if cells have row/col indices
        const hasIndices = pageCells.every((c) => c.row !== undefined && c.col !== undefined);
        if (hasIndices) {
            // Use row/col indices to build a proper table
            const grid = new Map();
            const usedRows = new Set();
            const usedCols = new Set();
            for (const cell of pageCells) {
                const r = cell.row;
                const c = cell.col;
                if (!grid.has(r))
                    grid.set(r, new Map());
                grid.get(r).set(c, cell.text);
                usedRows.add(r);
                usedCols.add(c);
            }
            if (usedRows.size === 0)
                continue;
            // Remap sparse row/col indices to compact consecutive indices
            const rowMap = new Map();
            [...usedRows].sort((a, b) => a - b).forEach((r, i) => rowMap.set(r, i));
            const colMap = new Map();
            [...usedCols].sort((a, b) => a - b).forEach((c, i) => colMap.set(c, i));
            // Write compact grid to Excel (1-based)
            for (const [gridR, excelR] of rowMap) {
                for (const [gridC, excelC] of colMap) {
                    const excelCell = sheet.getCell(excelR + 1, excelC + 1);
                    excelCell.value = grid.get(gridR)?.get(gridC) ?? "";
                    excelCell.border = {
                        top: { style: "thin" },
                        left: { style: "thin" },
                        bottom: { style: "thin" },
                        right: { style: "thin" },
                    };
                }
            }
        }
        else {
            // Fallback: use PDF positions with a fine grid
            const step = 2;
            const gridCols = Math.ceil(100 / step);
            const gridRows = Math.ceil(100 / step);
            const grid = Array.from({ length: gridRows }, () => Array(gridCols).fill(null));
            for (const cell of pageCells) {
                const col = Math.floor(cell.x / step);
                const row = Math.floor(cell.y / step);
                if (row < gridRows && col < gridCols && grid[row][col] === null) {
                    grid[row][col] = cell.text;
                }
            }
            const usedRows = new Set();
            const usedCols = new Set();
            for (let r = 0; r < gridRows; r++) {
                for (let c = 0; c < gridCols; c++) {
                    if (grid[r][c] !== null) {
                        usedRows.add(r);
                        usedCols.add(c);
                    }
                }
            }
            if (usedRows.size === 0)
                continue;
            const rowMap = new Map();
            [...usedRows].sort((a, b) => a - b).forEach((r, i) => rowMap.set(r, i));
            const colMap = new Map();
            [...usedCols].sort((a, b) => a - b).forEach((c, i) => colMap.set(c, i));
            for (const [gridR, excelR] of rowMap) {
                for (const [gridC, excelC] of colMap) {
                    const excelCell = sheet.getCell(excelR + 1, excelC + 1);
                    excelCell.value = grid[gridR]?.[gridC] ?? "";
                    excelCell.border = {
                        top: { style: "thin" },
                        left: { style: "thin" },
                        bottom: { style: "thin" },
                        right: { style: "thin" },
                    };
                }
            }
        }
        // Set reasonable column widths
        sheet.columns.forEach((column) => {
            column.width = 18;
        });
    }
    const arrayBuffer = await workbook.xlsx.writeBuffer();
    return Buffer.from(arrayBuffer);
}
