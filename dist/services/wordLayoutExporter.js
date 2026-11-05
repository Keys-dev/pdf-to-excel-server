import { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType, HeadingLevel } from "docx";
/**
 * Export layout cells to Word, placing each cell at a position that approximates
 * its original PDF layout. Uses a grid-based table where cells are placed at
 * their percentage positions mapped to table rows/columns.
 */
export async function layoutToWordBuffer(cells, fileName) {
    const children = [
        new Paragraph({ text: fileName, heading: HeadingLevel.HEADING_1 }),
    ];
    if (cells.length === 0) {
        children.push(new Paragraph({ text: "No cells were selected." }));
    }
    // Group cells by page
    const byPage = new Map();
    for (const cell of cells) {
        const existing = byPage.get(cell.page) || [];
        existing.push(cell);
        byPage.set(cell.page, existing);
    }
    for (const [page, pageCells] of byPage) {
        children.push(new Paragraph({ text: `Page ${page}`, heading: HeadingLevel.HEADING_2 }));
        // Build a grid table from the layout cells
        const table = buildLayoutTable(pageCells);
        children.push(table);
        children.push(new Paragraph({ text: "" }));
    }
    const doc = new Document({
        sections: [{ children }],
    });
    return Packer.toBuffer(doc);
}
function buildLayoutTable(cells) {
    // Check if cells have row/col indices
    const hasIndices = cells.every((c) => c.row !== undefined && c.col !== undefined);
    if (hasIndices) {
        // Use row/col indices to build a proper table
        const grid = new Map();
        const usedRows = new Set();
        const usedCols = new Set();
        for (const cell of cells) {
            const r = cell.row;
            const c = cell.col;
            if (!grid.has(r))
                grid.set(r, new Map());
            grid.get(r).set(c, cell.text);
            usedRows.add(r);
            usedCols.add(c);
        }
        // Remap sparse row/col indices to compact consecutive indices
        const sortedRows = [...usedRows].sort((a, b) => a - b);
        const sortedCols = [...usedCols].sort((a, b) => a - b);
        const actualRows = sortedRows.length;
        const actualCols = sortedCols.length;
        const cellWidth = 100 / actualCols;
        const rows = Array.from({ length: actualRows }, (_, ri) => {
            const gridR = sortedRows[ri];
            const rowCells = Array.from({ length: actualCols }, (_, ci) => {
                const gridC = sortedCols[ci];
                const value = grid.get(gridR)?.get(gridC) ?? "";
                return new TableCell({
                    width: { size: cellWidth, type: WidthType.PERCENTAGE },
                    children: [new Paragraph({ text: value })],
                });
            });
            return new TableRow({ children: rowCells });
        });
        return new Table({
            rows,
            width: { size: 100, type: WidthType.PERCENTAGE },
        });
    }
    // Fallback: use PDF positions with a fine grid
    const step = 2;
    const gridCols = Math.ceil(100 / step);
    const gridRows = Math.ceil(100 / step);
    const grid = Array.from({ length: gridRows }, () => Array(gridCols).fill(null));
    for (const cell of cells) {
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
    // If no cells were placed, return an empty table
    if (usedRows.size === 0) {
        return new Table({
            rows: [new TableRow({ children: [new TableCell({ children: [new Paragraph({ text: "" })] })] })],
            width: { size: 100, type: WidthType.PERCENTAGE },
        });
    }
    // Remap to compact consecutive indices
    const sortedRows = [...usedRows].sort((a, b) => a - b);
    const sortedCols = [...usedCols].sort((a, b) => a - b);
    const actualRows = sortedRows.length;
    const actualCols = sortedCols.length;
    // Build docx table rows from compact grid
    const cellWidth = 100 / actualCols;
    const rows = Array.from({ length: actualRows }, (_, ri) => {
        const gridR = sortedRows[ri];
        const rowCells = Array.from({ length: actualCols }, (_, ci) => {
            const gridC = sortedCols[ci];
            const value = grid[gridR]?.[gridC] ?? "";
            return new TableCell({
                width: { size: cellWidth, type: WidthType.PERCENTAGE },
                children: [new Paragraph({ text: value })],
            });
        });
        return new TableRow({ children: rowCells });
    });
    return new Table({
        rows,
        width: { size: 100, type: WidthType.PERCENTAGE },
    });
}
