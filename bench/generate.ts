// Synthetic bench workbooks and tasks (spec §11). Run with `npx tsx bench/generate.ts`.
// Every value is made up and seeded (42): names from faker, emails at example.com / example.org,
// phones in the fictional 555-01xx range. Never put real data in this repo (spec §0).
import { mkdirSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { join, resolve } from "node:path";
import ExcelJS from "exceljs";
import { Faker, base, en } from "@faker-js/faker";
import type { CellValue, RangeData } from "../src/core/types";
import { validateTask, type BenchTask, type Category, type Check } from "./compare";

export const SEED = 42;
export const BENCH_MARKER = "NYMFORM_BENCH_V1";
export const BENCH_SHEET = "_nymform_bench";
export const CREATOR = "Nymform bench";
/** Fixed workbook created/modified time and zip entry time, so regenerating gives identical files. */
export const FIXED_TIME = Date.UTC(2026, 8, 24, 12, 0, 0);

export const FMT = {
  date: "m/d/yyyy",
  time: "h:mm AM/PM",
  amount: "#,##0.00",
  whole: "#,##0",
  general: "General",
} as const;

export type WorkbookKey = "orders" | "roster" | "sections" | "appointments";
export const WORKBOOK_KEYS: readonly WorkbookKey[] = ["orders", "roster", "sections", "appointments"];

export interface ColumnDef {
  header: string;
  numFmt: string;
  width: number;
}

export interface BenchWorkbook {
  key: WorkbookKey;
  file: string;
  sheet: string;
  /** Local address of the header row plus data rows, e.g. "A1:F501". */
  range: string;
  columns: ColumnDef[];
  /** Data rows (no header). Dates are Excel serials, times day fractions. */
  rows: CellValue[][];
  privateLetters: string[];
  canaries: string[];
}

export interface BenchData {
  workbooks: BenchWorkbook[];
  tasks: BenchTask[];
}

export interface PrivateValuesEntry {
  sheet: string;
  range: string;
  privateLetters: string[];
  /** Every non-empty display text and raw value of the private data cells, deduplicated. */
  values: string[];
  canaries: string[];
}

// ---------------------------------------------------------------------------------------------
// Dates, times and display text

const MS_PER_DAY = 86_400_000;
const EPOCH_OFFSET = 25_569; // serial of 1970-01-01 with Excel's 1899-12-30 epoch

export function serialOf(y: number, m: number, d: number): number {
  return Date.UTC(y, m - 1, d) / MS_PER_DAY + EPOCH_OFFSET;
}

export function partsOf(serial: number): { y: number; m: number; d: number; weekday: number } {
  const dt = new Date((Math.floor(serial) - EPOCH_OFFSET) * MS_PER_DAY);
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), weekday: dt.getUTCDay() };
}

function grouped(n: number, decimals: number): string {
  const fixed = Math.abs(n).toFixed(decimals);
  const [int = "0", frac] = fixed.split(".");
  const withCommas = int.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return (n < 0 ? "-" : "") + withCommas + (frac !== undefined ? `.${frac}` : "");
}

/** The text Excel shows for a value in one of the bench number formats. */
export function displayText(value: CellValue, numFmt: string): string {
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "string") return value;
  switch (numFmt) {
    case FMT.date: {
      const p = partsOf(value);
      return `${p.m}/${p.d}/${p.y}`;
    }
    case FMT.time: {
      const minutes = Math.round((value - Math.floor(value)) * 1440) % 1440;
      const h = Math.floor(minutes / 60);
      const mm = String(minutes % 60).padStart(2, "0");
      return `${h % 12 === 0 ? 12 : h % 12}:${mm} ${h < 12 ? "AM" : "PM"}`;
    }
    case FMT.amount:
      return grouped(value, 2);
    case FMT.whole:
      return grouped(value, 0);
    default:
      return String(value);
  }
}

function letter(i: number): string {
  return String.fromCharCode(65 + i);
}

/** The workbook's selection as the Office adapter would hand it over (header row included). */
export function toRangeData(wb: BenchWorkbook): RangeData {
  const headers = wb.columns.map((c) => c.header);
  const valueType = (v: CellValue) =>
    v === null ? "Empty" : typeof v === "string" ? "String" : typeof v === "boolean" ? "Boolean" : "Double";
  return {
    sheet: wb.sheet,
    address: wb.range,
    rowIndex: 0,
    columnIndex: 0,
    values: [headers, ...wb.rows.map((r) => [...r])],
    text: [headers, ...wb.rows.map((r) => r.map((v, c) => displayText(v, wb.columns[c]!.numFmt)))],
    formulas: [headers, ...wb.rows.map((r) => r.map((v) => (v === null ? "" : v)))],
    valueTypes: [headers.map(() => "String"), ...wb.rows.map((r) => r.map(valueType))],
    numberFormat: [headers.map(() => FMT.general), ...wb.rows.map((r) => r.map((_, c) => wb.columns[c]!.numFmt))],
    table: null,
  };
}

/** Every non-empty display text and raw value in the private data cells, first-seen order. */
export function privateValuesOf(wb: BenchWorkbook): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  const add = (s: string) => {
    if (s !== "" && !seen.has(s)) {
      seen.add(s);
      out.push(s);
    }
  };
  const cols = wb.privateLetters.map((l) => l.charCodeAt(0) - 65);
  for (const row of wb.rows) {
    for (const c of cols) {
      const v = row[c] ?? null;
      if (v === null) continue;
      add(displayText(v, wb.columns[c]!.numFmt));
      add(String(v));
    }
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// Prompts

/** Task prompts: answerable from the structure alone, and free of any private value. */
export const PROMPTS: Readonly<Record<string, string>> = {
  "orders-01": "Which Product was in the order with the largest Amount?",
  "orders-02": "How many orders from the West region had an Amount of 1,000 or more?",
  "orders-03": "Total Amount per Region, one row per region",
  "orders-04": "What was the total Amount of orders placed in March 2025?",
  "orders-05": "For each order, the domain of the customer's Email (the part after the @)",
  "orders-06": "Rank each order's Amount within its Region, where 1 is the largest (ties share a rank)",
  "orders-07":
    "Flag orders that repeat an earlier row with the same Customer, Order date, Product and Amount: TRUE on the repeat, FALSE otherwise",
  "roster-01": "Average Salary per Department, one row per department",
  "roster-02": "Which Department does the employee with the earliest Start date work in?",
  "roster-03": "Completed years of service for each employee as of June 30, 2025",
  "roster-04": "Each employee's last name (the last word of Name)",
  "roster-05": "The three highest-paid employees, highest first: Name and Salary",
  "sections-01": "Number of sections for each Days pattern, one row per pattern",
  "sections-02": "Who is the Instructor of the section with the highest Enrolled?",
  "sections-03": "A label for each section: the Course, a hyphen, then the Section, like BIO 110-02",
  "sections-04":
    "The three sections with the lowest fill rate (Enrolled divided by Cap): Course, Section and fill rate",
  "appointments-01":
    "Total Duration (min) per Location, counting only appointments that were not no-shows, one row per location",
  "appointments-02": "Number of no-shows per Location, one row per location",
  "appointments-03": "How many appointments fell on a Saturday or Sunday?",
  "appointments-04":
    "Flag double bookings: TRUE when the same Client name has another appointment on the same Appointment date, otherwise FALSE",
};

/** Every column header across the four workbooks. */
const HEADERS = [
  ...["Customer", "Email", "Region", "Order date", "Product", "Amount"],
  ...["Name", "Employee ID", "Department", "Rank", "Start date", "Salary"],
  ...["Course", "Section", "Instructor", "Days", "Start time", "Units", "Enrolled", "Cap"],
  ...["Client name", "Phone", "Location", "Appointment date", "Duration (min)", "No-show"],
];

// ---------------------------------------------------------------------------------------------
// Canaries: made-up names that exist nowhere else, planted once each in private columns.

const CANARY_FIRST = [
  "Quorbel", "Zanthe", "Vendrix", "Oskaru", "Thelvani", "Mirrabel", "Kestrun", "Ulvanne", "Brixomar", "Yselde",
  "Dravenko", "Pellissa", "Corvantha", "Ismerel", "Taldric", "Wennomar", "Fiorxa", "Gavrinne", "Hollast", "Jorvanne",
  "Lustrena", "Maelthor", "Nixabel", "Orrindel", "Pruvella", "Quennric", "Rhosvald", "Sylthra", "Tovrenna", "Ulbrexa",
  "Vassoly", "Wistrel", "Xandrell", "Yorvessa", "Zerrund", "Abrisande", "Belqunor", "Cindravel", "Dostrevin", "Elvquist",
];
const CANARY_LAST = [
  "Vantrisk", "Pellowin", "Drastmoor", "Quillaby", "Fennrow", "Marbrecht", "Oxtavel", "Brindlewock", "Corrander", "Thistlevane",
  "Ulmsforde", "Varnhollow", "Wexcombe", "Yarrowind", "Zelthorpe", "Astravell", "Bexmoor", "Crindall", "Dovrasque", "Esterquill",
  "Fallowmere", "Grenvasse", "Hovrand", "Istelwick", "Jastrowen", "Kelvanox", "Lorrimaze", "Morvessant", "Norquell", "Orrendash",
  "Pexhollow", "Quarnvelt", "Rostivane", "Selkmarr", "Tromvelde", "Uskavell", "Vorthane", "Wendrisk", "Xelmoor", "Yavrenko",
];

function canaryName(i: number): { first: string; last: string; full: string } {
  const first = CANARY_FIRST[i]!;
  const last = CANARY_LAST[i]!;
  return { first, last, full: `${first} ${last}` };
}

// ---------------------------------------------------------------------------------------------
// Helpers over the seeded faker instance

const NAME_PART_RE = /^[A-Z][A-Za-z'-]{1,}$/;

/** Words of 4+ characters, lowercased, split on anything that isn't a letter or digit. */
export function words4(text: string): string[] {
  return text
    .toLowerCase()
    .split(/[^\p{L}\p{N}]+/u)
    .filter((w) => w.length >= 4);
}

/**
 * Words a generated person name must not contain: the workbooks' fixed vocabulary, headers and task
 * prompts. The auditor's word variant would rightly block a request where a customer's surname
 * ("West") is also a region, so such a name would measure a collision, not the model.
 */
function reservedWords(): Set<string> {
  const texts = [
    ...REGIONS,
    ...PRODUCTS.map((p) => p.name),
    ...DEPARTMENTS,
    ...RANKS.map((r) => r.name),
    ...SUBJECTS,
    ...DAYS.map(([d]) => d),
    ...LOCATIONS,
    ...HEADERS,
    ...Object.values(PROMPTS),
  ];
  return new Set(texts.flatMap(words4));
}

function round2(x: number): number {
  return Math.round(x * 100) / 100;
}

function roundTo(x: number, digits: number): number {
  const f = 10 ** digits;
  return Math.round(x * f) / f;
}

function weighted<T>(f: Faker, items: readonly (readonly [T, number])[]): T {
  return f.helpers.weightedArrayElement(items.map(([value, weight]) => ({ value, weight })));
}

let reserved: Set<string> | null = null;

function uniqueName(f: Faker, used: Set<string>): { first: string; last: string; full: string } {
  reserved ??= reservedWords();
  for (;;) {
    const first = f.person.firstName();
    const last = f.person.lastName();
    if (!NAME_PART_RE.test(first) || !NAME_PART_RE.test(last)) continue;
    if (words4(`${first} ${last}`).some((w) => reserved!.has(w))) continue;
    const full = `${first} ${last}`;
    if (used.has(full.toLowerCase())) continue;
    used.add(full.toLowerCase());
    return { first, last, full };
  }
}

function emailFor(f: Faker, first: string, last: string, used: Set<string>): string {
  const local = `${first}.${last}`.toLowerCase().replace(/[^a-z0-9.]/g, "");
  const domain = f.datatype.boolean({ probability: 0.3 }) ? "example.org" : "example.com";
  for (let n = 1; ; n++) {
    const email = `${local}${n === 1 ? "" : n}@${domain}`;
    if (!used.has(email)) {
      used.add(email);
      return email;
    }
  }
}

function randomSerial(f: Faker, from: number, to: number): number {
  return f.number.int({ min: from, max: to });
}

function assert(cond: unknown, message: string): asserts cond {
  if (!cond) throw new Error(`Bench data check failed: ${message}`);
}

function sortBy<T>(items: T[], ...keys: ((t: T) => string | number)[]): T[] {
  return items
    .map((t, i) => ({ t, i }))
    .sort((a, b) => {
      for (const k of keys) {
        const x = k(a.t);
        const y = k(b.t);
        if (x < y) return -1;
        if (x > y) return 1;
      }
      return a.i - b.i;
    })
    .map((e) => e.t);
}

/** Groups rows by a key, sorted by key, reducing each group to a value. */
function groupRows<K extends string>(
  rows: readonly CellValue[][],
  key: (r: CellValue[]) => K,
  reduce: (group: CellValue[][]) => number,
): [K, number][] {
  const groups = new Map<K, CellValue[][]>();
  for (const r of rows) {
    const k = key(r);
    const g = groups.get(k);
    if (g) g.push(r);
    else groups.set(k, [r]);
  }
  return [...groups.keys()].sort().map((k) => [k, reduce(groups.get(k)!)]);
}

const num = (v: CellValue | undefined): number => (typeof v === "number" ? v : NaN);
const str = (v: CellValue | undefined): string => (typeof v === "string" ? v : "");

/** The index of the single largest (or smallest) value; throws on a tie so the task stays well defined. */
function uniqueExtreme(values: readonly number[], kind: "max" | "min", what: string): number {
  const best = kind === "max" ? Math.max(...values) : Math.min(...values);
  const at = values.flatMap((v, i) => (v === best ? [i] : []));
  assert(at.length === 1, `${what} is not unique`);
  return at[0]!;
}

// ---------------------------------------------------------------------------------------------
// Orders: Customer, Email, Region, Order date, Product, Amount (500 rows)

const REGIONS = ["Central", "East", "North", "South", "West"] as const;
const PRODUCTS = [
  { name: "Desk Lamp", min: 24, max: 49 },
  { name: "Office Chair", min: 149, max: 329 },
  { name: "Standing Desk", min: 399, max: 749 },
  { name: "Monitor Arm", min: 59, max: 129 },
  { name: "Keyboard", min: 39, max: 119 },
  { name: "Webcam", min: 49, max: 139 },
  { name: "Headset", min: 29, max: 179 },
  { name: "Notebook Pack", min: 9, max: 24 },
] as const;

function makeOrders(f: Faker, usedNames: Set<string>): BenchWorkbook {
  const usedEmails = new Set<string>();
  const customers = Array.from({ length: 180 }, () => {
    const n = uniqueName(f, usedNames);
    return { name: n.full, email: emailFor(f, n.first, n.last, usedEmails), region: f.helpers.arrayElement(REGIONS) };
  });
  const from = serialOf(2025, 1, 1);
  const to = serialOf(2025, 12, 31);
  type Order = { row: CellValue[]; special: boolean };
  const order = (c: { name: string; email: string; region: string }, special = false): Order => {
    const p = f.helpers.arrayElement(PRODUCTS);
    const price = f.number.float({ min: p.min, max: p.max, fractionDigits: 2 });
    const qty = weighted(f, [[1, 40], [2, 22], [3, 14], [4, 9], [5, 6], [6, 4], [8, 3], [10, 2]] as const);
    const discount = f.datatype.boolean({ probability: 0.15 }) ? 0.9 : 1;
    return { row: [c.name, c.email, c.region, randomSerial(f, from, to), p.name, round2(price * qty * discount)], special };
  };

  const orders: Order[] = [];
  for (let i = 0; i < 488; i++) orders.push(order(f.helpers.arrayElement(customers)));
  // Five made-up customers whose names and emails are canaries, one order each.
  const canaries: string[] = [];
  const canaryEmails: string[] = [];
  for (let i = 0; i < 5; i++) {
    const n = canaryName(i);
    const e = canaryName(35 + i);
    const email = `${e.first}.${e.last}@example.org`.toLowerCase();
    canaries.push(n.full);
    canaryEmails.push(email);
    orders.push(order({ name: n.full, email, region: f.helpers.arrayElement(REGIONS) }, true));
  }
  // One bulk order, the single largest Amount.
  const bulkBuyer = f.helpers.arrayElement(customers);
  orders.push({
    row: [bulkBuyer.name, bulkBuyer.email, bulkBuyer.region, randomSerial(f, from, to), "Standing Desk", 9874.5],
    special: true,
  });

  let sorted = sortBy(orders, (o) => num(o.row[3]));
  // Six double entries: an exact repeat of an earlier order, one to three rows later.
  const candidates = sorted.flatMap((o, i) => (o.special ? [] : [i]));
  const picks = f.helpers.arrayElements(candidates, 6).sort((a, b) => b - a);
  for (const i of picks) {
    const at = Math.min(sorted.length, i + f.number.int({ min: 1, max: 3 }));
    sorted = [...sorted.slice(0, at), { row: [...sorted[i]!.row], special: true }, ...sorted.slice(at)];
  }
  const rows = sorted.map((o) => o.row);
  assert(rows.length === 500, "orders has 500 rows");
  return {
    key: "orders",
    file: "orders.xlsx",
    sheet: "Orders",
    range: "A1:F501",
    columns: [
      { header: "Customer", numFmt: FMT.general, width: 22 },
      { header: "Email", numFmt: FMT.general, width: 32 },
      { header: "Region", numFmt: FMT.general, width: 10 },
      { header: "Order date", numFmt: FMT.date, width: 12 },
      { header: "Product", numFmt: FMT.general, width: 16 },
      { header: "Amount", numFmt: FMT.amount, width: 12 },
    ],
    rows,
    privateLetters: ["A", "B"],
    canaries: [...canaries, ...canaryEmails],
  };
}

function ordersTasks(wb: BenchWorkbook): BenchTask[] {
  const rows = wb.rows;
  const amounts = rows.map((r) => num(r[5]));
  const largest = uniqueExtreme(amounts, "max", "largest order");
  const march = [serialOf(2025, 3, 1), serialOf(2025, 4, 1)] as const;
  const t = (id: string, category: Category, check: Check) => task(wb, id, category, check);
  return [
    t("orders-01", "lookup", {
      type: "cell",
      expected: str(rows[largest]![4]),
    }),
    t("orders-02", "counting", {
      type: "cell",
      expected: rows.filter((r) => r[2] === "West" && num(r[5]) >= 1000).length,
      tolerance: 0,
    }),
    t("orders-03", "conditional_totals", {
      type: "table",
      expected: groupRows(rows, (r) => str(r[2]), (g) => round2(g.reduce((s, r) => s + num(r[5]), 0))),
      tolerance: 0.01,
    }),
    t("orders-04", "dates", {
      type: "cell",
      expected: round2(rows.filter((r) => num(r[3]) >= march[0] && num(r[3]) < march[1]).reduce((s, r) => s + num(r[5]), 0)),
      tolerance: 0.01,
    }),
    t("orders-05", "text", {
      type: "column",
      expected: rows.map((r) => str(r[1]).slice(str(r[1]).indexOf("@") + 1)),
    }),
    t(
      "orders-06",
      "ranking",
      {
        type: "column",
        expected: rows.map((r) => 1 + rows.filter((o) => o[2] === r[2] && num(o[5]) > num(r[5])).length),
        tolerance: 0,
      },
    ),
    t(
      "orders-07",
      "row_anomalies",
      {
        type: "column",
        expected: rows.map((r, i) =>
          rows.slice(0, i).some((o) => o[0] === r[0] && o[3] === r[3] && o[4] === r[4] && o[5] === r[5]),
        ),
      },
    ),
  ];
}

// ---------------------------------------------------------------------------------------------
// Roster: Name, Employee ID, Department, Rank, Start date, Salary (120 rows)

const DEPARTMENTS = ["Engineering", "Finance", "Marketing", "Operations", "Sales", "Support"] as const;
const RANKS = [
  { name: "Associate", weight: 40, min: 52_000, max: 68_000 },
  { name: "Senior", weight: 30, min: 70_000, max: 92_000 },
  { name: "Lead", weight: 15, min: 90_000, max: 115_000 },
  { name: "Manager", weight: 10, min: 105_000, max: 135_000 },
  { name: "Director", weight: 5, min: 140_000, max: 185_000 },
] as const;

function makeRoster(f: Faker, usedNames: Set<string>): BenchWorkbook {
  const usedIds = new Set<string>();
  const newId = () => {
    for (;;) {
      const id = `E${f.number.int({ min: 10_000, max: 99_999 })}`;
      if (!usedIds.has(id)) {
        usedIds.add(id);
        return id;
      }
    }
  };
  const from = serialOf(2004, 1, 5);
  const to = serialOf(2025, 6, 1);
  const employee = (name: string, start?: number): CellValue[] => {
    const rank = weighted(f, RANKS.map((r) => [r, r.weight] as const));
    const salary = Math.round(f.number.int({ min: rank.min, max: rank.max }) / 100) * 100;
    return [name, newId(), f.helpers.arrayElement(DEPARTMENTS), rank.name, start ?? randomSerial(f, from, to), salary];
  };
  const rows: CellValue[][] = [];
  for (let i = 0; i < 108; i++) rows.push(employee(uniqueName(f, usedNames).full));
  // A round salary, as payroll data often has: its range rendering "90000-99999" starts with it.
  const round = employee(uniqueName(f, usedNames).full);
  round[3] = "Senior";
  round[5] = 90_000;
  rows.push(round);
  // The longest-serving employee, the single earliest Start date.
  rows.push(employee(uniqueName(f, usedNames).full, serialOf(2002, 3, 18)));
  const canaries: string[] = [];
  for (let i = 5; i < 15; i++) {
    const n = canaryName(i);
    canaries.push(n.full);
    rows.push(employee(n.full));
  }
  return {
    key: "roster",
    file: "roster.xlsx",
    sheet: "Roster",
    range: "A1:F121",
    columns: [
      { header: "Name", numFmt: FMT.general, width: 22 },
      { header: "Employee ID", numFmt: FMT.general, width: 13 },
      { header: "Department", numFmt: FMT.general, width: 14 },
      { header: "Rank", numFmt: FMT.general, width: 11 },
      { header: "Start date", numFmt: FMT.date, width: 12 },
      { header: "Salary", numFmt: FMT.whole, width: 11 },
    ],
    rows: sortBy(rows, (r) => str(r[1])),
    privateLetters: ["A", "B", "F"],
    canaries,
  };
}

function rosterTasks(wb: BenchWorkbook): BenchTask[] {
  const rows = wb.rows;
  const earliest = uniqueExtreme(rows.map((r) => num(r[4])), "min", "earliest start date");
  const bySalary = sortBy([...rows], (r) => -num(r[5]));
  assert(num(bySalary[2]![5]) > num(bySalary[3]![5]), "third-highest salary is not unique");
  const t = (id: string, category: Category, check: Check) => task(wb, id, category, check);
  return [
    t("roster-01", "conditional_totals", {
      type: "table",
      expected: groupRows(rows, (r) => str(r[2]), (g) => roundTo(g.reduce((s, r) => s + num(r[5]), 0) / g.length, 4)),
      tolerance: 0.01,
    }),
    t("roster-02", "lookup", {
      type: "cell",
      expected: str(rows[earliest]![2]),
    }),
    t("roster-03", "dates", {
      type: "column",
      expected: rows.map((r) => {
        const p = partsOf(num(r[4]));
        return 2025 - p.y - (p.m > 6 || (p.m === 6 && p.d > 30) ? 1 : 0);
      }),
      tolerance: 0,
    }),
    t("roster-04", "text", {
      type: "column",
      expected: rows.map((r) => str(r[0]).split(" ").pop() ?? ""),
    }),
    t("roster-05", "ranking", {
      type: "table",
      expected: bySalary.slice(0, 3).map((r) => [str(r[0]), num(r[5])]),
      tolerance: 0,
    }),
  ];
}

// ---------------------------------------------------------------------------------------------
// Sections: Course, Section, Instructor, Days, Start time, Units, Enrolled, Cap (200 rows)

const SUBJECTS = ["ANTH", "BIO", "CHEM", "CS", "ECON", "ENGL", "HIST", "MATH", "PHYS", "PSYC", "SOC", "SPAN"] as const;
const COURSE_NUMBERS = [101, 110, 120, 150, 201, 210, 220, 301, 310, 350, 401] as const;
const DAYS = [
  ["MWF", 30],
  ["TTh", 32],
  ["MW", 20],
  ["TW", 8],
  ["F", 10],
] as const;

function makeSections(f: Faker, usedNames: Set<string>): BenchWorkbook {
  const instructors = Array.from({ length: 55 }, () => uniqueName(f, usedNames).full);
  const courses = f.helpers.shuffle(SUBJECTS.flatMap((s) => COURSE_NUMBERS.map((n) => `${s} ${n}`)));
  const section = (course: string, n: number): CellValue[] => {
    const cap = weighted(f, [[24, 15], [30, 20], [35, 15], [40, 20], [45, 10], [60, 10], [80, 6], [120, 4]] as const);
    const enrolled = Math.max(1, Math.round(cap * f.number.float({ min: 0.55, max: 1.08 })));
    const start = (f.number.int({ min: 16, max: 38 }) * 30) / 1440; // 8:00 AM to 7:00 PM
    return [
      course,
      String(n).padStart(2, "0"),
      f.helpers.arrayElement(instructors),
      weighted(f, DAYS),
      start,
      weighted(f, [[1, 5], [2, 10], [3, 45], [4, 35], [5, 5]] as const),
      enrolled,
      cap,
    ];
  };
  const rows: CellValue[][] = [];
  for (const course of courses) {
    const count = weighted(f, [[1, 30], [2, 30], [3, 20], [4, 12], [5, 8]] as const);
    for (let n = 1; n <= count && rows.length < 199; n++) rows.push(section(course, n));
    if (rows.length >= 199) break;
  }
  assert(rows.length === 199, "sections has 199 generated rows");
  // One large lecture: the single highest Enrolled.
  rows.push(["PSYC 100", "01", f.helpers.arrayElement(instructors), "MWF", 600 / 1440, 3, 243, 250]);
  // Three nearly empty sections: the three lowest fill rates.
  const low = f.helpers.arrayElements(
    rows.map((_, i) => i).slice(0, 199),
    3,
  );
  const lowFill: [number, number][] = [
    [3, 40],
    [6, 45],
    [6, 35],
  ];
  low.forEach((i, k) => {
    const r = rows[i]!;
    r[6] = lowFill[k]![0];
    r[7] = lowFill[k]![1];
  });
  // Ten made-up instructors, each teaching one section.
  const canaries: string[] = [];
  const taken = new Set(low);
  const canaryRows = f.helpers.arrayElements(
    rows.map((_, i) => i).filter((i) => i < 199 && !taken.has(i)),
    10,
  );
  canaryRows.forEach((i, k) => {
    const n = canaryName(15 + k);
    canaries.push(n.full);
    rows[i]![2] = n.full;
  });
  return {
    key: "sections",
    file: "sections.xlsx",
    sheet: "Sections",
    range: "A1:H201",
    columns: [
      { header: "Course", numFmt: FMT.general, width: 11 },
      { header: "Section", numFmt: FMT.general, width: 9 },
      { header: "Instructor", numFmt: FMT.general, width: 22 },
      { header: "Days", numFmt: FMT.general, width: 7 },
      { header: "Start time", numFmt: FMT.time, width: 11 },
      { header: "Units", numFmt: FMT.general, width: 7 },
      { header: "Enrolled", numFmt: FMT.general, width: 9 },
      { header: "Cap", numFmt: FMT.general, width: 7 },
    ],
    rows: sortBy(rows, (r) => str(r[0]).split(" ")[0]!, (r) => Number(str(r[0]).split(" ")[1]), (r) => str(r[1])),
    privateLetters: ["C"],
    canaries,
  };
}

function sectionsTasks(wb: BenchWorkbook): BenchTask[] {
  const rows = wb.rows;
  const biggest = uniqueExtreme(rows.map((r) => num(r[6])), "max", "highest Enrolled");
  const rate = (r: CellValue[]) => num(r[6]) / num(r[7]);
  const byRate = sortBy([...rows], rate);
  assert(rate(byRate[2]!) < rate(byRate[3]!), "third-lowest fill rate is not unique");
  const t = (id: string, category: Category, check: Check) => task(wb, id, category, check);
  return [
    t("sections-01", "counting", {
      type: "table",
      expected: groupRows(rows, (r) => str(r[3]), (g) => g.length),
      tolerance: 0,
    }),
    t("sections-02", "lookup", {
      type: "cell",
      expected: str(rows[biggest]![2]),
    }),
    t("sections-03", "text", {
      type: "column",
      expected: rows.map((r) => `${str(r[0])}-${str(r[1])}`),
    }),
    t(
      "sections-04",
      "ranking",
      {
        type: "table",
        expected: byRate.slice(0, 3).map((r) => [str(r[0]), str(r[1]), roundTo(rate(r), 6)]),
        tolerance: 0.0005,
      },
    ),
  ];
}

// ---------------------------------------------------------------------------------------------
// Appointments: Client name, Phone, Location, Appointment date, Duration (min), No-show (300 rows)

const LOCATIONS = ["Downtown", "Eastside", "Harbor View", "Northgate"] as const;
const AREA_CODES = ["831", "408", "650", "510", "415"] as const;

function makeAppointments(f: Faker, usedNames: Set<string>): BenchWorkbook {
  const usedPhones = new Set<string>();
  const phone = () => {
    for (;;) {
      // 555-0100 to 555-0199 is reserved for fictional use.
      const p = `(${f.helpers.arrayElement(AREA_CODES)}) 555-01${String(f.number.int({ min: 0, max: 99 })).padStart(2, "0")}`;
      if (!usedPhones.has(p)) {
        usedPhones.add(p);
        return p;
      }
    }
  };
  const clients = Array.from({ length: 140 }, () => ({ name: uniqueName(f, usedNames).full, phone: phone() }));
  const from = serialOf(2025, 1, 2);
  const to = serialOf(2025, 12, 30);
  const day = () => {
    for (;;) {
      const s = randomSerial(f, from, to);
      const wd = partsOf(s).weekday;
      if (wd === 0 && !f.datatype.boolean({ probability: 0.12 })) continue;
      if (wd === 6 && !f.datatype.boolean({ probability: 0.45 })) continue;
      return s;
    }
  };
  const appointment = (c: { name: string; phone: string }, date = day()): CellValue[] => [
    c.name,
    c.phone,
    f.helpers.arrayElement(LOCATIONS),
    date,
    weighted(f, [[15, 10], [30, 40], [45, 20], [60, 25], [90, 5]] as const),
    f.datatype.boolean({ probability: 0.12 }),
  ];
  const rows: CellValue[][] = [];
  for (let i = 0; i < 286; i++) rows.push(appointment(f.helpers.arrayElement(clients)));
  const canaries: string[] = [];
  for (let i = 25; i < 35; i++) {
    const n = canaryName(i);
    canaries.push(n.full);
    rows.push(appointment({ name: n.full, phone: phone() }));
  }
  // Four double bookings: the same client twice on the same day.
  for (const i of f.helpers.arrayElements(rows.map((_, k) => k).slice(0, 286), 4)) {
    const r = rows[i]!;
    rows.push(appointment({ name: str(r[0]), phone: str(r[1]) }, num(r[3])));
  }
  assert(rows.length === 300, "appointments has 300 rows");
  return {
    key: "appointments",
    file: "appointments.xlsx",
    sheet: "Appointments",
    range: "A1:F301",
    columns: [
      { header: "Client name", numFmt: FMT.general, width: 22 },
      { header: "Phone", numFmt: FMT.general, width: 16 },
      { header: "Location", numFmt: FMT.general, width: 13 },
      { header: "Appointment date", numFmt: FMT.date, width: 17 },
      { header: "Duration (min)", numFmt: FMT.general, width: 14 },
      { header: "No-show", numFmt: FMT.general, width: 9 },
    ],
    rows: sortBy(rows, (r) => num(r[3])),
    privateLetters: ["A", "B"],
    canaries,
  };
}

function appointmentsTasks(wb: BenchWorkbook): BenchTask[] {
  const rows = wb.rows;
  const noShows = groupRows(rows, (r) => str(r[2]), (g) => g.filter((r) => r[5] === true).length);
  assert(noShows.length === LOCATIONS.length && noShows.every(([, n]) => n > 0), "every location has a no-show");
  const t = (id: string, category: Category, check: Check) => task(wb, id, category, check);
  return [
    t(
      "appointments-01",
      "conditional_totals",
      {
        type: "table",
        expected: groupRows(rows, (r) => str(r[2]), (g) => g.filter((r) => r[5] === false).reduce((s, r) => s + num(r[4]), 0)),
        tolerance: 0,
      },
    ),
    t("appointments-02", "counting", {
      type: "table",
      expected: noShows,
      tolerance: 0,
    }),
    t("appointments-03", "dates", {
      type: "cell",
      expected: rows.filter((r) => [0, 6].includes(partsOf(num(r[3])).weekday)).length,
      tolerance: 0,
    }),
    t(
      "appointments-04",
      "row_anomalies",
      {
        type: "column",
        expected: rows.map((r, i) => rows.some((o, k) => k !== i && o[0] === r[0] && o[3] === r[3])),
      },
    ),
  ];
}

// ---------------------------------------------------------------------------------------------

function task(wb: BenchWorkbook, id: string, category: Category, check: Check): BenchTask {
  const prompt = PROMPTS[id];
  assert(prompt !== undefined, `task ${id} has a prompt`);
  return { id, workbook: wb.file, sheet: wb.sheet, range: wb.range, category, prompt, private: [...wb.privateLetters], check };
}

function checkWorkbook(wb: BenchWorkbook, tasks: readonly BenchTask[]): void {
  for (const c of wb.columns) assert(HEADERS.includes(c.header), `${wb.key} header is listed in HEADERS`);
  const [start, end] = wb.range.split(":");
  assert(start === "A1", `${wb.key} range starts at A1`);
  assert(end === `${letter(wb.columns.length - 1)}${wb.rows.length + 1}`, `${wb.key} range matches its data`);
  assert(wb.rows.every((r) => r.length === wb.columns.length), `${wb.key} rows are rectangular`);
  assert(wb.canaries.length === 10 && new Set(wb.canaries).size === 10, `${wb.key} has 10 distinct canaries`);
  const privateCols = new Set(wb.privateLetters.map((l) => l.charCodeAt(0) - 65));
  for (const canary of wb.canaries) {
    assert(canary.length >= 6, `${wb.key} canary length`);
    const lower = canary.toLowerCase();
    let exact = 0;
    wb.rows.forEach((r) =>
      r.forEach((v, c) => {
        const text = displayText(v, wb.columns[c]!.numFmt);
        if (text === canary) {
          assert(privateCols.has(c), `${wb.key} canary sits in a private column`);
          exact++;
        } else {
          assert(!text.toLowerCase().includes(lower), `${wb.key} canary appears inside another cell`);
        }
      }),
    );
    assert(exact === 1, `${wb.key} canary planted exactly once`);
  }
  // No word of a private multi-word value may also be a word of the headers, prompts or other cells.
  const vocab = new Set([
    ...wb.columns.flatMap((c) => words4(c.header)),
    ...tasks.filter((t) => t.workbook === wb.file).flatMap((t) => words4(t.prompt)),
    ...wb.rows.flatMap((r) => r.flatMap((v, c) => (privateCols.has(c) ? [] : words4(displayText(v, wb.columns[c]!.numFmt))))),
  ]);
  for (const r of wb.rows) {
    for (const c of privateCols) {
      const text = displayText(r[c] ?? null, wb.columns[c]!.numFmt);
      if (/\s/.test(text.trim())) assert(!words4(text).some((w) => vocab.has(w)), `${wb.key} private value shares a word with other text`);
    }
  }
}

/** Builds all four workbooks and the 20 tasks from one seeded faker instance. Pure and deterministic. */
export function generateData(seed = SEED): BenchData {
  const f = new Faker({ locale: [en, base] });
  f.seed(seed);
  const usedNames = new Set<string>(CANARY_FIRST.map((_, i) => canaryName(i).full.toLowerCase()));
  const orders = makeOrders(f, usedNames);
  const roster = makeRoster(f, usedNames);
  const sections = makeSections(f, usedNames);
  const appointments = makeAppointments(f, usedNames);
  const workbooks = [orders, roster, sections, appointments];
  const tasks = [...ordersTasks(orders), ...rosterTasks(roster), ...sectionsTasks(sections), ...appointmentsTasks(appointments)];
  for (const wb of workbooks) checkWorkbook(wb, tasks);
  const all = workbooks.flatMap((w) => w.canaries);
  assert(new Set(all).size === all.length, "canaries are unique across workbooks");

  for (const t of tasks) {
    const problems = validateTask(t);
    assert(problems.length === 0, `task ${t.id}: ${problems.join(" ")}`);
    const wb = workbooks.find((w) => w.file === t.workbook)!;
    if (t.check.type === "column") assert(t.check.expected.length === wb.rows.length, `task ${t.id} has one value per row`);
    const values = privateValuesOf(wb).filter((v) => v.length >= 4);
    const prompt = t.prompt.toLowerCase();
    assert(!values.some((v) => prompt.includes(v.toLowerCase())), `task ${t.id} prompt holds no private value`);
  }
  assert(tasks.length === 20 && new Set(tasks.map((t) => t.id)).size === 20, "20 distinct tasks");
  return { workbooks, tasks };
}

export function privateValuesFile(data: BenchData): Record<string, PrivateValuesEntry> {
  const out: Record<string, PrivateValuesEntry> = {};
  for (const wb of data.workbooks) {
    out[wb.key] = {
      sheet: wb.sheet,
      range: wb.range,
      privateLetters: [...wb.privateLetters],
      values: privateValuesOf(wb),
      canaries: [...wb.canaries],
    };
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// Output

/** JSON with two-space indentation, but arrays of plain values on one line, so task files stay readable. */
export function formatJson(value: unknown, indent = ""): string {
  const inner = indent + "  ";
  if (Array.isArray(value)) {
    if (value.length === 0) return "[]";
    if (value.every((v) => v === null || typeof v !== "object")) return `[${value.map((v) => JSON.stringify(v)).join(", ")}]`;
    return `[\n${value.map((v) => inner + formatJson(v, inner)).join(",\n")}\n${indent}]`;
  }
  if (value !== null && typeof value === "object") {
    const entries = Object.entries(value).filter(([, v]) => v !== undefined);
    if (entries.length === 0) return "{}";
    return `{\n${entries.map(([k, v]) => `${inner}${JSON.stringify(k)}: ${formatJson(v, inner)}`).join(",\n")}\n${indent}}`;
  }
  return JSON.stringify(value);
}

/** Runs `fn` with `new Date()` and `Date.now()` pinned, so zip entries carry a fixed time. */
async function withFixedClock<T>(ms: number, fn: () => Promise<T>): Promise<T> {
  const RealDate = Date;
  class FixedDate extends RealDate {
    constructor(...args: unknown[]) {
      if (args.length === 0) super(ms);
      else super(...(args as [number, number]));
    }
    static override now(): number {
      return ms;
    }
  }
  globalThis.Date = FixedDate as DateConstructor;
  try {
    return await fn();
  } finally {
    globalThis.Date = RealDate;
  }
}

export async function workbookBuffer(wb: BenchWorkbook): Promise<Buffer> {
  const book = new ExcelJS.Workbook();
  book.creator = CREATOR;
  book.lastModifiedBy = CREATOR;
  book.created = new Date(FIXED_TIME);
  book.modified = new Date(FIXED_TIME);

  const ws = book.addWorksheet(wb.sheet, { views: [{ state: "frozen", ySplit: 1 }] });
  ws.columns = wb.columns.map((c) => ({ width: c.width }));
  const header = ws.addRow(wb.columns.map((c) => c.header));
  header.font = { bold: true };
  for (const r of wb.rows) {
    const row = ws.addRow(r);
    wb.columns.forEach((c, i) => {
      if (c.numFmt !== FMT.general) row.getCell(i + 1).numFmt = c.numFmt;
    });
  }

  const marker = book.addWorksheet(BENCH_SHEET, { state: "hidden" });
  marker.addRow([BENCH_MARKER]);
  marker.addRow(["private", ...wb.privateLetters]);
  marker.addRow(["canaries", ...wb.canaries]);

  const out = await withFixedClock(FIXED_TIME, () => book.xlsx.writeBuffer());
  return Buffer.from(out as ArrayBuffer);
}

export async function writeAll(data: BenchData, root: string): Promise<string[]> {
  const workbooksDir = join(root, "workbooks");
  const tasksDir = join(root, "tasks");
  mkdirSync(workbooksDir, { recursive: true });
  mkdirSync(tasksDir, { recursive: true });
  const written: string[] = [];
  for (const wb of data.workbooks) {
    const p = join(workbooksDir, wb.file);
    writeFileSync(p, await workbookBuffer(wb));
    written.push(p);
  }
  const pv = join(workbooksDir, "private-values.json");
  writeFileSync(pv, formatJson(privateValuesFile(data)) + "\n");
  written.push(pv);
  const files = new Set(data.tasks.map((t) => `${t.id}.json`));
  for (const name of readdirSync(tasksDir)) {
    if (name.endsWith(".json") && !files.has(name)) rmSync(join(tasksDir, name));
  }
  for (const t of data.tasks) {
    const p = join(tasksDir, `${t.id}.json`);
    writeFileSync(p, formatJson(t) + "\n");
    written.push(p);
  }
  // The bench panel imports the tasks through this index (bench builds only).
  const ids = data.tasks.map((t) => t.id);
  const index = [
    "// Generated by bench/generate.ts. The bench panel's task list (bench builds only).",
    'import type { BenchTask } from "../compare";',
    ...ids.map((id, i) => `import t${i} from "./${id}.json";`),
    "",
    `export const TASKS = [${ids.map((_, i) => `t${i}`).join(", ")}] as unknown as BenchTask[];`,
    "",
  ].join("\n");
  const indexPath = join(tasksDir, "index.ts");
  writeFileSync(indexPath, index);
  written.push(indexPath);
  return written;
}

async function main(): Promise<void> {
  const root = resolve(process.argv[1] ?? ".", "..");
  const data = generateData(SEED);
  const written = await writeAll(data, root);
  console.log(`Wrote ${written.length} files under ${root}:`);
  for (const w of data.workbooks) console.log(`  ${w.file}: ${w.sheet}!${w.range}, ${w.rows.length} rows, private ${w.privateLetters.join(",")}`);
  console.log(`  ${data.tasks.length} tasks, private-values.json`);
}

if (process.argv[1] && /generate\.ts$/.test(process.argv[1])) {
  main().catch((e: unknown) => {
    console.error(e instanceof Error ? e.message : e);
    process.exit(1);
  });
}
