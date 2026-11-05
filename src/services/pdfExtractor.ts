import { getDocument } from "pdfjs-dist/legacy/build/pdf.mjs";
import type { ExtractedTable, ExtractionResult, TableRow, TableConfidence, LayoutCell, PageLayout } from "../types.js";

interface PositionedItem {
  text: string;
  x: number;
  y: number;
  /** Actual glyph-run width in PDF user-space units, from pdfjs TextItem.width. */
  width: number;
  /** Font size in PDF points, from pdfjs TextItem.fontSize. */
  fontSize: number;
  /** Mark Content Item ID — present when PDF has structural markup. */
  mcid?: number;
}

/**
 * The gap (in PDF points) between the end of one text item and the start of the
 * next that we treat as a "column break" (two separate table cells).
 * This is a base threshold; it is scaled by font size for adaptive behavior.
 */
const COLUMN_GAP_THRESHOLD_BASE = 16;

/**
 * Gaps *smaller* than this are assumed to be adjacent glyphs within the same
 * word — no space is inserted. Gaps between WORD_GAP_MIN and
 * COLUMN_GAP_THRESHOLD are treated as a single inter-word space.
 *
 * Why this matters: some PDFs (especially those produced by design tools or
 * certain scanners) store each glyph as a separate TextItem. The estimated
 * endpoint (x + width) of individual characters can differ from the actual
 * start of the next glyph by 1–3 pts due to kerning/tracking, so a small
 * tolerance is needed to avoid inserting spurious spaces between them.
 */
const WORD_GAP_MIN = 2;

/** Minimum row tolerance in PDF points (floor value). */
const ROW_TOLERANCE_MIN = 2;

/** Maximum row tolerance in PDF points (ceiling value). */
const ROW_TOLERANCE_MAX = 20;

/**
 * Compute an adaptive row tolerance based on the median font size of items on
 * the page. Larger fonts → larger tolerance, smaller fonts → smaller tolerance.
 * Clamped between ROW_TOLERANCE_MIN and ROW_TOLERANCE_MAX.
 */
function computeRowTolerance(items: PositionedItem[]): number {
  if (items.length === 0) return 4;

  const fontSizes = items
    .map((item) => item.fontSize)
    .filter((fs) => fs > 0)
    .sort((a, b) => a - b);

  if (fontSizes.length === 0) return 4;

  // Use median font size
  const medianFontSize =
    fontSizes.length % 2 === 0
      ? (fontSizes[fontSizes.length / 2 - 1] + fontSizes[fontSizes.length / 2]) / 2
      : fontSizes[Math.floor(fontSizes.length / 2)];

  // Tolerance is ~0.5× median font size, clamped
  return Math.max(ROW_TOLERANCE_MIN, Math.min(ROW_TOLERANCE_MAX, medianFontSize * 0.5));
}

/**
 * Collect all x-start positions across all rows, then cluster them to produce
 * a shared set of column boundaries. This ensures consistent column counts
 * across all rows in a table.
 *
 * Uses gap-based clustering with frequency weighting: positions that appear
 * in more rows are weighted higher when computing the cluster center. This
 * makes column boundaries more robust to outlier rows.
 */
function computeColumnBoundaries(
  rows: PositionedItem[][],
  fontSize: number
): number[] {
  // Count how many rows each x-position appears in
  const xFrequency: Map<number, number> = new Map();
  const allXPositions: number[] = [];

  for (const row of rows) {
    // Deduplicate x positions within a single row
    const rowXSet = new Set(row.map((item) => item.x));
    for (const x of rowXSet) {
      xFrequency.set(x, (xFrequency.get(x) ?? 0) + 1);
      allXPositions.push(x);
    }
  }

  if (allXPositions.length === 0) return [];

  // Sort unique x-values
  const sorted = [...new Set(allXPositions)].sort((a, b) => a - b);
  if (sorted.length === 0) return [];

  // Gap-based clustering
  const clusterGap = Math.max(8, fontSize * 0.5);
  const clusters: number[][] = [[sorted[0]]];

  for (let i = 1; i < sorted.length; i++) {
    const lastCluster = clusters[clusters.length - 1];
    const lastValue = lastCluster[lastCluster.length - 1];

    if (sorted[i] - lastValue <= clusterGap) {
      lastCluster.push(sorted[i]);
    } else {
      clusters.push([sorted[i]]);
    }
  }

  // Take the frequency-weighted mean of each cluster as the column boundary
  // Positions appearing in more rows pull the center more strongly
  return clusters.map((cluster) => {
    if (cluster.length === 1) return cluster[0];
    const totalWeight = cluster.reduce((sum, x) => sum + (xFrequency.get(x) ?? 1), 0);
    const weightedMean = cluster.reduce((sum, x) => sum + x * (xFrequency.get(x) ?? 1), 0) / totalWeight;
    return weightedMean;
  });
}

/**
 * Group raw text items into rows using an adaptive y-tolerance based on font size.
 * Items are sorted by y (descending) then x (ascending), and grouped when their
 * y-positions are within the computed tolerance of the row's median y.
 *
 * Using the row's median y (instead of the last item's y) prevents drift when
 * items within a row have slightly different y-positions — a common occurrence
 * in real-world PDFs where baseline alignment is imperfect.
 */
function groupIntoRows(items: PositionedItem[], rowTolerance: number): PositionedItem[][] {
  const sorted = [...items].sort((a, b) => b.y - a.y || a.x - b.x);
  const rows: PositionedItem[][] = [];

  for (const item of sorted) {
    const lastRow = rows[rows.length - 1];
    if (lastRow && lastRow.length > 0) {
      // Use the median y of the current row for comparison, not the last item
      const rowYs = lastRow.map((i) => i.y).sort((a, b) => a - b);
      const medianY =
        rowYs.length % 2 === 0
          ? (rowYs[rowYs.length / 2 - 1] + rowYs[rowYs.length / 2]) / 2
          : rowYs[Math.floor(rowYs.length / 2)];

      if (Math.abs(medianY - item.y) <= rowTolerance) {
        lastRow.push(item);
      } else {
        rows.push([item]);
      }
    } else {
      rows.push([item]);
    }
  }

  return rows.map((row) => row.sort((a, b) => a.x - b.x));
}

/**
 * Convert a row of positioned items into table cells using global column
 * boundaries. Items are snapped to the nearest column boundary, ensuring
 * consistent column alignment across all rows.
 *
 * Items that share the same column but differ slightly in y (multi-line cells)
 * are joined with a space.
 */
function rowToCellsAligned(
  row: PositionedItem[],
  colBoundaries: number[],
  columnGapThreshold: number
): TableRow {
  if (row.length === 0 || colBoundaries.length === 0) return [];

  // Build column bin edges: midpoint between consecutive boundaries
  const colEdges: number[] = [colBoundaries[0] - 50];

  for (let i = 1; i < colBoundaries.length; i++) {
    const mid = (colBoundaries[i - 1] + colBoundaries[i]) / 2;
    colEdges.push(mid);
  }
  colEdges.push(colBoundaries[colBoundaries.length - 1] + 500);

  // Collect items per column, preserving order
  const colItems: PositionedItem[][] = new Array(colBoundaries.length)
    .fill(null)
    .map(() => []);

  for (const item of row) {
    let colIndex = 0;
    for (let i = 0; i < colEdges.length - 1; i++) {
      if (item.x >= colEdges[i] && item.x < colEdges[i + 1]) {
        colIndex = i;
        break;
      }
    }
    colIndex = Math.min(colIndex, colBoundaries.length - 1);
    colItems[colIndex].push(item);
  }

  // Build cell text from each column's items
  const cells: string[] = colItems.map((itemsInCol) => {
    if (itemsInCol.length === 0) return "";
    // Sort by y first (for multi-line), then by x
    const sorted = [...itemsInCol].sort((a, b) => a.y - b.y || a.x - b.x);
    let text = sorted[0].text;
    for (let i = 1; i < sorted.length; i++) {
      const prev = sorted[i - 1];
      const gap = sorted[i].x - (prev.x + (prev.width > 0 ? prev.width : prev.text.length * 6));
      if (gap > columnGapThreshold) {
        text += " " + sorted[i].text;
      } else if (gap > WORD_GAP_MIN) {
        text += " " + sorted[i].text;
      } else {
        text += sorted[i].text;
      }
    }
    return text.trim();
  });

  // Remove trailing empty cells but preserve empty cells between non-empty ones
  let lastNonEmpty = cells.length - 1;
  while (lastNonEmpty >= 0 && cells[lastNonEmpty] === "") lastNonEmpty--;
  return cells.slice(0, lastNonEmpty + 1);
}

/**
 * Fallback: convert a row to cells using gap-based heuristic (original approach).
 * Used when global alignment is not applicable.
 */
function rowToCells(row: PositionedItem[], columnGapThreshold: number): TableRow {
  if (row.length === 0) return [];

  const cells: string[] = [];
  let current = row[0].text;
  // Use the actual glyph-run width from pdfjs; fall back to a per-character
  // estimate only when pdfjs reports 0 (e.g. some control characters).
  let lastX = row[0].x + (row[0].width > 0 ? row[0].width : row[0].text.length * 6);

  for (let i = 1; i < row.length; i++) {
    const item = row[i];
    const gap = item.x - lastX;

    if (gap > columnGapThreshold) {
      // Large gap → new column cell
      cells.push(current.trim());
      current = item.text;
    } else if (gap > WORD_GAP_MIN) {
      // Medium gap → inter-word space (same cell, separate words)
      current += " " + item.text;
    } else {
      // Small / negative gap → adjacent glyphs in the same word, no space
      current += item.text;
    }

    lastX = item.x + (item.width > 0 ? item.width : item.text.length * 6);
  }

  if (current.trim()) cells.push(current.trim());
  return cells;
}

function looksLikeTableRow(cells: TableRow): boolean {
  return cells.length >= 2;
}

/**
 * Normalize extracted text: collapse multiple whitespace characters, trim,
 * and remove zero-width or invisible Unicode characters that sometimes appear
 * in PDF text layers.
 */
function normalizeText(text: string): string {
  return text
    .replace(/[\u200B-\u200D\uFEFF]/g, "") // zero-width spaces and BOM
    .replace(/[\t\r\n]+/g, " ") // tabs and newlines to space
    .replace(/ +/g, " ") // collapse multiple spaces
    .trim();
}

/**
 * Split a list of positioned rows into separate table regions based on vertical
 * gaps between rows. If the gap between two consecutive rows is significantly
 * larger than the median row gap, they belong to different tables.
 *
 * Returns an array of row groups, each representing a distinct table region.
 */
function splitIntoTableRegions(
  positionedRows: PositionedItem[][],
  medianFontSize: number
): PositionedItem[][][] {
  if (positionedRows.length < 2) return [positionedRows];

  // Compute gaps between consecutive rows
  const gaps: number[] = [];
  for (let i = 1; i < positionedRows.length; i++) {
    const prevY = positionedRows[i - 1][0]?.y ?? 0;
    const currY = positionedRows[i][0]?.y ?? 0;
    gaps.push(Math.abs(prevY - currY));
  }

  if (gaps.length === 0) return [positionedRows];

  // Use median gap as the baseline row spacing
  const sortedGaps = [...gaps].sort((a, b) => a - b);
  const medianGap =
    sortedGaps.length % 2 === 0
      ? (sortedGaps[sortedGaps.length / 2 - 1] + sortedGaps[sortedGaps.length / 2]) / 2
      : sortedGaps[Math.floor(sortedGaps.length / 2)];

  // A gap is a table boundary if it's significantly larger than the median
  // Use a threshold of 2× median gap or at least 30 points (whichever is smaller)
  const boundaryThreshold = Math.min(medianGap * 2, 30 + medianFontSize);

  const regions: PositionedItem[][][] = [[positionedRows[0]]];

  for (let i = 1; i < positionedRows.length; i++) {
    const prevY = positionedRows[i - 1][0]?.y ?? 0;
    const currY = positionedRows[i][0]?.y ?? 0;
    const gap = Math.abs(prevY - currY);

    if (gap > boundaryThreshold) {
      regions.push([positionedRows[i]]);
    } else {
      regions[regions.length - 1].push(positionedRows[i]);
    }
  }

  return regions;
}

/**
 * Compute confidence metrics for a detected table based on structural regularity.
 */
function computeConfidence(
  rows: PositionedItem[][],
  cellRows: TableRow[],
  fromMarkup: boolean
): TableConfidence {
  // Column consistency: what fraction of rows have the same column count as the mode
  const colCounts = cellRows.map((r) => r.length);
  const colCountFreq: Record<number, number> = {};
  for (const c of colCounts) {
    colCountFreq[c] = (colCountFreq[c] ?? 0) + 1;
  }
  const modeColCount = Number(
    Object.entries(colCountFreq).sort((a, b) => Number(b[1]) - Number(a[1]))[0]?.[0] ?? 0
  );
  const columnConsistency =
    colCounts.length > 0
      ? Math.round((colCountFreq[modeColCount] / colCounts.length) * 100)
      : 0;

  // Row spacing regularity: how uniform are the y-gaps between consecutive rows
  if (rows.length >= 3) {
    const rowYs = rows.map((r) => r[0]?.y ?? 0);
    const gaps: number[] = [];
    for (let i = 1; i < rowYs.length; i++) {
      gaps.push(Math.abs(rowYs[i] - rowYs[i - 1]));
    }
    const meanGap = gaps.reduce((s, g) => s + g, 0) / gaps.length;
    const variance = gaps.reduce((s, g) => s + Math.pow(g - meanGap, 2), 0) / gaps.length;
    const stdDev = Math.sqrt(variance);
    // Coefficient of variation → regularity score
    const cv = meanGap > 0 ? stdDev / meanGap : 0;
    const rowSpacingRegularity = Math.max(0, Math.round(100 * (1 - cv)));

    // Overall score: weighted average, with markup bonus
    let score = Math.round(
      columnConsistency * 0.6 + rowSpacingRegularity * 0.4
    );
    if (fromMarkup) score = Math.min(100, score + 15);

    return {
      score,
      columnConsistency,
      rowSpacingRegularity,
      fromMarkup,
    };
  }

  // Few rows — give a baseline score
  const baseScore = fromMarkup ? 85 : 60;
  return {
    score: baseScore,
    columnConsistency,
    rowSpacingRegularity: 100,
    fromMarkup,
  };
}

/**
 * Check if the PDF page has structural markup (MCID or annotations) that can
 * be used for more accurate table extraction.
 */
function hasStructuralMarkup(contentItems: any[]): boolean {
  return contentItems.some(
    (item) =>
      typeof item.mcid === "number" ||
      (item.attrs && typeof item.attrs.MCID === "number")
  );
}

/**
 * Extract table structure from PDF markup if available.
 * Uses MCID (Mark Content Item ID) to identify table cells.
 */
function extractFromMarkup(
  items: PositionedItem[]
): { rows: TableRow[]; fromMarkup: boolean } | null {
  // Filter items that have MCID
  const markupItems = items.filter((item) => item.mcid !== undefined && item.mcid !== -1);

  if (markupItems.length < 4) {
    return null; // Not enough markup items to form a table
  }

  // Group by MCID to identify cells
  const mcidGroups: Record<number, PositionedItem[]> = {};
  for (const item of markupItems) {
    const mcid = item.mcid!;
    if (!mcidGroups[mcid]) mcidGroups[mcid] = [];
    mcidGroups[mcid].push(item);
  }

  // Each MCID group represents a cell — combine text items within each group
  const cells = Object.entries(mcidGroups).map(([mcid, groupItems]) => {
    const sorted = groupItems.sort((a, b) => a.y - b.y || a.x - b.x);
    const text = sorted.map((i) => i.text).join(" ").trim();
    return {
      mcid: Number(mcid),
      text,
      x: Math.min(...groupItems.map((i) => i.x)),
      y: Math.min(...groupItems.map((i) => i.y)),
    };
  });

  if (cells.length < 4) return null;

  // Group cells into rows by y-position
  const rowTolerance = computeRowTolerance(items);
  const sortedByY = cells.sort((a, b) => b.y - a.y || a.x - b.x);
  const rows: { y: number; cells: typeof cells }[] = [];

  for (const cell of sortedByY) {
    const lastRow = rows[rows.length - 1];
    if (lastRow && Math.abs(lastRow.y - cell.y) <= rowTolerance) {
      lastRow.cells.push(cell);
    } else {
      rows.push({ y: cell.y, cells: [cell] });
    }
  }

  // Sort cells within each row by x
  for (const row of rows) {
    row.cells.sort((a, b) => a.x - b.x);
  }

  // Convert to string[][] with normalized text
  const tableRows = rows
    .map((row) => row.cells.map((c) => normalizeText(c.text)))
    .filter((r) => r.length >= 2);

  if (tableRows.length < 2) return null;

  return { rows: tableRows, fromMarkup: true };
}

/**
 * Build per-page layout data from positioned items. Converts absolute PDF
 * coordinates to page-relative percentages and assigns row/column indices
 * based on the grouped rows structure.
 */
function buildPageLayout(
  items: PositionedItem[],
  positionedRows: PositionedItem[][],
  pageWidth: number,
  pageHeight: number
): PageLayout {
  // Compute row/col indices from the grouped rows
  const cells: LayoutCell[] = [];
  for (let ri = 0; ri < positionedRows.length; ri++) {
    const row = positionedRows[ri];
    for (let ci = 0; ci < row.length; ci++) {
      const item = row[ci];
      cells.push({
        text: item.text,
        x: Math.round((item.x / pageWidth) * 10000) / 100,
        y: Math.round(((pageHeight - item.y) / pageHeight) * 10000) / 100,
        w: Math.round((item.width / pageWidth) * 10000) / 100,
        h: Math.round((item.fontSize / pageHeight) * 10000) / 100,
        row: ri,
        col: ci,
      });
    }
  }

  return {
    page: 0, // filled in by caller
    cells,
    rowCount: positionedRows.length,
    colCount: positionedRows.length > 0 ? Math.max(...positionedRows.map((r) => r.length)) : 0,
  };
}

export async function extractFromPdf(buffer: Buffer, fileName: string): Promise<ExtractionResult> {
  const doc = await getDocument({ data: new Uint8Array(buffer) }).promise;
  const tables: ExtractedTable[] = [];
  const layouts: PageLayout[] = [];
  let rawText = "";

  for (let pageNum = 1; pageNum <= doc.numPages; pageNum++) {
    const page = await doc.getPage(pageNum);
    const content = await page.getTextContent();

    // Check for structural markup before processing
    const rawContentItems = content.items as any[];
    const markupAvailable = hasStructuralMarkup(rawContentItems);

    const items: PositionedItem[] = rawContentItems
      .filter((item) => typeof item.str === "string" && item.str.trim().length > 0)
      .map((item) => ({
        text: item.str as string,
        x: item.transform[4] as number,
        y: item.transform[5] as number,
        width: typeof item.width === "number" ? item.width : 0,
        fontSize: typeof item.fontSize === "number" ? item.fontSize : 12,
        mcid:
          typeof item.mcid === "number"
            ? item.mcid
            : item.attrs && typeof item.attrs.MCID === "number"
            ? item.attrs.MCID
            : undefined,
      }));

    rawText += items.map((i) => i.text).join(" ") + "\n";

    // Compute page dimensions for layout positioning
    const viewport = page.getViewport({ scale: 1 });
    const pageWidth = viewport.width;
    const pageHeight = viewport.height;

    // Heuristic extraction with adaptive tolerance and global column alignment
    const rowTolerance = computeRowTolerance(items);

    // Compute median font size for adaptive thresholds
    const fontSizes = items
      .map((item) => item.fontSize)
      .filter((fs) => fs > 0)
      .sort((a, b) => a - b);
    const medianFontSize =
      fontSizes.length > 0
        ? fontSizes.length % 2 === 0
          ? (fontSizes[fontSizes.length / 2 - 1] + fontSizes[fontSizes.length / 2]) / 2
          : fontSizes[Math.floor(fontSizes.length / 2)]
        : 12;

    // Group items into rows (always needed for layout)
    const positionedRows = groupIntoRows(items, rowTolerance);

    // Build layout data for this page
    const pageLayout = buildPageLayout(items, positionedRows, pageWidth, pageHeight);
    pageLayout.page = pageNum;
    layouts.push(pageLayout);

    // Try markup-based extraction first if available
    if (markupAvailable) {
      const markupResult = extractFromMarkup(items);
      if (markupResult && markupResult.rows.length >= 2) {
        const confidence = computeConfidence(
          [items], // Single row group for markup — confidence is high by default
          markupResult.rows,
          true
        );
        tables.push({
          page: pageNum,
          rows: markupResult.rows,
          confidence,
        });
        continue; // Skip heuristic table extraction for this page
      }
    }

    // Adaptive column gap threshold based on font size
    const columnGapThreshold = COLUMN_GAP_THRESHOLD_BASE * (medianFontSize / 12);

    // Split rows into separate table regions based on vertical gaps
    const tableRegions = splitIntoTableRegions(positionedRows, medianFontSize);

    for (const regionRows of tableRegions) {
      // Compute column boundaries per region for better alignment
      const colBoundaries = computeColumnBoundaries(regionRows, medianFontSize);

      let cellRows: TableRow[];
      if (colBoundaries.length >= 2) {
        // Use global alignment with per-region column boundaries
        cellRows = regionRows.map((row) =>
          rowToCellsAligned(row, colBoundaries, columnGapThreshold)
        );
      } else {
        // Fallback to gap-based heuristic
        cellRows = regionRows.map((row) => rowToCells(row, columnGapThreshold));
      }

      // Normalize text in each cell and filter to table-like rows
      const filteredRows = cellRows
        .map((row) => row.map(normalizeText))
        .filter(looksLikeTableRow);

      if (filteredRows.length >= 2) {
        const confidence = computeConfidence(regionRows, filteredRows, false);
        tables.push({
          page: pageNum,
          rows: filteredRows,
          confidence,
        });
      }
    }
  }

  return {
    fileName,
    pageCount: doc.numPages,
    tables,
    rawText: rawText.trim(),
    layout: layouts,
  };
}
