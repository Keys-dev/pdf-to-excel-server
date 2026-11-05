export type TableRow = string[];

/** Confidence metrics for a detected table */
export interface TableConfidence {
  /** Overall confidence score 0–100 */
  score: number;
  /** How consistent column counts are across rows (0–100) */
  columnConsistency: number;
  /** How regular row spacing is (0–100) */
  rowSpacingRegularity: number;
  /** Whether the table was extracted from PDF markup (higher confidence) vs heuristic */
  fromMarkup: boolean;
}

export interface ExtractedTable {
  page: number;
  rows: TableRow[];
  /** Optional confidence metrics for this table */
  confidence?: TableConfidence;
}

/** A single text item with its position on the page (percentages of page width/height) */
export interface LayoutCell {
  /** The extracted text */
  text: string;
  /** Left edge as % of page width (0–100) */
  x: number;
  /** Top edge as % of page height (0–100) */
  y: number;
  /** Width as % of page width (0–100) */
  w: number;
  /** Height as % of page height (0–100) */
  h: number;
  /** Row index assigned by layout grouping (for coordinate display) */
  row: number;
  /** Column index assigned by layout grouping (for coordinate display) */
  col: number;
}

/** One page worth of layout cells */
export interface PageLayout {
  /** 1-based page number */
  page: number;
  /** Cells positioned on this page */
  cells: LayoutCell[];
  /** Number of rows detected on this page */
  rowCount: number;
  /** Number of columns detected on this page */
  colCount: number;
}

export interface ExtractionResult {
  fileName: string;
  pageCount: number;
  tables: ExtractedTable[];
  rawText: string;
  /** Per-page layout data preserving original text positions */
  layout?: PageLayout[];
}

/** A cell with its layout position for layout-preserving export */
export interface LayoutExportCell {
  text: string;
  /** Left edge as % of page width (0–100) */
  x: number;
  /** Top edge as % of page height (0–100) */
  y: number;
  /** Width as % of page width (0–100) */
  w: number;
  /** Height as % of page height (0–100) */
  h: number;
  /** 1-based page number */
  page: number;
  /** Row index in the detected table grid */
  row?: number;
  /** Column index in the detected table grid */
  col?: number;
}
