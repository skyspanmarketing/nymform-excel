// Wires the Excel host (src/office/adapter.ts, the only file that touches Office.js) into the flow.
// Named imports only, so a release bundle leaves out the adapter's bench-only functions.
import {
  ExcelHostError,
  SelectionTooLargeError,
  firstEmptyColumn,
  insertFormula,
  isRangeEmpty,
  listNames,
  listSheets,
  listTables,
  officeTheme,
  readRange,
  readSelection,
  readWorkbookInfo,
  ready,
  selectCells,
} from "../office/adapter";
import type { HostAdapter } from "./flow";

/** What the task pane needs to pick a theme: Office's own theme, when the host says. */
export interface ThemeInfo {
  isDarkTheme?: boolean;
  bodyBackgroundColor?: string;
}

export interface OfficeHost extends HostAdapter {
  /** Resolves once Office.js is ready inside Excel; rejects elsewhere. */
  ready(): Promise<unknown>;
  /** Office's theme, or null when the host doesn't say. */
  theme(): ThemeInfo | null;
}

/** True for a dark Office theme: its own flag, or else a dark body background. */
export function isDarkTheme(theme: ThemeInfo | null): boolean {
  if (!theme) return false;
  if (typeof theme.isDarkTheme === "boolean") return theme.isDarkTheme;
  const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(theme.bodyBackgroundColor?.trim() ?? "");
  if (!m) return false;
  const [r, g, b] = [m[1], m[2], m[3]].map((h) => parseInt(h ?? "0", 16) / 255) as [number, number, number];
  return 0.2126 * r + 0.7152 * g + 0.0722 * b < 0.4;
}

export const officeHost: OfficeHost = {
  ready: () => ready(),
  readSelection: (opts) => readSelection(opts),
  listSheets: () => listSheets(),
  listTables: () => listTables(),
  listNames: () => listNames(),
  readWorkbookInfo: () => readWorkbookInfo(),
  isRangeEmpty: (sheet, address) => isRangeEmpty(sheet, address),
  selectCells: (sheet, cells) => selectCells(sheet, cells),
  readRange: (fullAddress) => readRange(fullAddress),
  firstEmptyColumn: (sheet, fromColumn, firstRow, lastRow) => firstEmptyColumn(sheet, fromColumn, firstRow, lastRow),
  insertFormula: (sheet, cell, formula, fillDown, lastRow) => insertFormula(sheet, cell, formula, fillDown, lastRow),
  isTooLarge: (e) => e instanceof SelectionTooLargeError,
  // The adapter's messages are plain sentences without values or stack traces; a too-large one
  // names at most an address and a cell count ("The data around E4 has 28,000 cells. ...").
  plainMessage: (e) => (e instanceof ExcelHostError || e instanceof SelectionTooLargeError ? e.message : null),
  theme: () => {
    try {
      return officeTheme();
    } catch {
      return null;
    }
  },
};
