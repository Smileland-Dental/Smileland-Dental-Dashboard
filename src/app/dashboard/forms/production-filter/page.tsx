'use client';

import React, { useCallback, useEffect, useMemo, useState } from 'react';
import { db } from '@/lib/firebase.config';
import { collection, getDocs, query, where } from 'firebase/firestore';

const LOCATION_OPTIONS = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia', 'Crowns', 'Endo'] as const;

type TableRow = {
  position: string;
  name: string;
  sales: string;
  coffeeNew: string;
  coffeeReturn: string;
  coffeeTotal: string;
  coffeeNo: string;
  renderedCoffee: string;
  coffeeYes: string;
  orangeJuiceNew: string;
  orangeJuiceReturn: string;
  orangeJuiceTotal: string;
};

type SugarRow = {
  position: string;
  name: string;
  sugar: string;
  sugarGood: string;
  sugarBad: string;
  paper: string;
};

type TableTotals = {
  sales: number;
  coffeeNew: number;
  coffeeReturn: number;
  coffeeTotal: number;
  coffeeNo: number;
  renderedCoffee: number;
  coffeeYes: number;
  orangeJuiceNew: number;
  orangeJuiceReturn: number;
  orangeJuiceTotal: number;
};

type SugarTotals = {
  sugar: number;
  sugarGood: number;
  sugarBad: number;
  paper: number;
};

type ReasonRow = {
  reason: string;
  orangeJuice: string;
  paper: string;
  coffee: string;
};

type LocationSummary = {
  pineapple?: string;
  rose?: string;
  total?: string;
  mailedProduction?: string;
};

type ProductionSideMetrics = {
  add?: string;
  noShow?: string;
  scheduled?: string;
  seen?: string;
  seenPercent?: string;
};

type ExtraInputRow = {
  position?: string;
  name?: string;
  customer?: string;
  icecream?: string;
  cake?: string;
  donut?: string;
  tart?: string;
  peach?: string;
  peppermint?: string;
  doctorPreventative?: string;
  doctorRestorative?: string;
  doctorCraProduction?: string;
};

type SimpleFormDoc = {
  id: string;
  date?: string;
  location?: string;
  reasonIfLate?: string;
  submittedDateTime?: string;
  grandTotal?: string;
  coffeeSales?: string;
  salesWithoutCoffee?: string;
  paperAtOrangeJuice?: string;
  paperAtTea?: string;
  justPaper?: string;
  prophyTotal?: string;
  notes?: string;
  notDue?: string;
  /** When true, form is finalized and should appear in reports. */
  pdfSaved?: boolean;
  tableRows?: TableRow[];
  sugarRows?: SugarRow[];
  reasonRows?: ReasonRow[];
  tableTotals?: Partial<TableTotals> & Record<string, unknown>;
  sugarTotals?: Partial<SugarTotals> & Record<string, unknown>;
  locationSummary?: LocationSummary;
  productionSideMetrics?: ProductionSideMetrics;
  coffeeActualTotals?: Partial<Record<keyof TableTotals, string>>;
  extraInputRows?: ExtraInputRow[];
  checkIn?: string;
  checkOut?: string;
  closer?: string;
  [key: string]: unknown;
};

type ColumnFieldId =
  | ''
  | 'main.grandTotal'
  | 'main.coffeeSales'
  | 'main.salesWithoutCoffee'
  | 'main.paperAtOrangeJuice'
  | 'main.paperAtTea'
  | 'main.justPaper'
  | 'main.prophyTotal'
  | 'table.sales'
  | 'table.coffeeNew'
  | 'table.coffeeReturn'
  | 'table.coffeeTotal'
  | 'table.coffeeNo'
  | 'table.renderedCoffee'
  | 'table.coffeeYes'
  | 'table.orangeJuiceNew'
  | 'table.orangeJuiceReturn'
  | 'table.orangeJuiceTotal'
  | 'sugar.sugar'
  | 'sugar.sugarGood'
  | 'sugar.sugarBad'
  | 'sugar.paper'
  | 'location.pineapple'
  | 'location.rose'
  | 'location.total'
  | 'location.mailedProduction'
  | 'side.add'
  | 'side.noShow'
  | 'side.scheduled'
  | 'side.seen'
  | 'side.seenPercent'
  | 'actual.orangeJuiceNew'
  | 'actual.orangeJuiceReturn'
  | 'docperf.customer'
  | 'docperf.icecream'
  | 'docperf.cake'
  | 'docperf.donut'
  | 'docperf.tart'
  | 'docperf.peach'
  | 'docperf.peppermint'
  | 'docperf.doctorPreventative'
  | 'docperf.doctorRestorative'
  | 'docperf.doctorCraProduction';

type ShortProcBundleId = 'bundle.spOE' | 'bundle.spPro' | 'bundle.spCRA';

const OPERATING_BUNDLE_ID = 'bundle.operating' as const;

type OperatingBundleId = typeof OPERATING_BUNDLE_ID;

const OPERATING_COL_COUNT = 4;

const OPERATING_SUB_LABELS = ['Check In', 'Check Out', 'Hours Open', 'Closer'] as const;

const OPERATING_HOURS_OPEN_COL_INDEX = 2;

const OPERATING_HOURS_OPEN_MIN_WIDTH = 168;

const NAME_BUNDLE_ID = 'bundle.nameRow' as const;

const POSITION_BUNDLE_ID = 'bundle.positionRow' as const;

type SlotValue =
  | ColumnFieldId
  | ShortProcBundleId
  | OperatingBundleId
  | typeof NAME_BUNDLE_ID
  | typeof POSITION_BUNDLE_ID;

const FIXED_REASON_OPTIONS = [
  'Declined/DDP/Pt Left',
  'Not Due/Freq',
  'Medical Clearance',
  'Furture TX/FMS',
  'Courtesy, Not Billable',
  'Mistakenly Done, Not Billable',
  'Uncooperative/Re-eval',
  'Not Documented',
  'Age Limit/No teeth',
  'Not Complete/Filled out',
] as const;

const BUNDLE_COLSPAN = FIXED_REASON_OPTIONS.length;

const SHORT_PROC_BUNDLE_OPTIONS: { id: ShortProcBundleId; label: string }[] = [
  { id: 'bundle.spOE', label: 'Short Procedures — OE' },
  { id: 'bundle.spPro', label: 'Short Procedures — Pro' },
  { id: 'bundle.spCRA', label: 'Short Procedures — CRA' },
];

const FIELD_GROUPS: { label: string; options: { id: ColumnFieldId; label: string }[] }[] = [
  {
    label: 'Summary',
    options: [
      { id: 'main.grandTotal', label: 'Submitted Production' },
      { id: 'location.total', label: '1st Review Production' },
      { id: 'location.mailedProduction', label: 'Mailed Production' },
      { id: 'main.coffeeSales', label: 'CRA Production Production' },
      { id: 'main.salesWithoutCoffee', label: 'Production W/Out CRA' },
      { id: 'main.paperAtOrangeJuice', label: 'Prophy @ OE' },
      { id: 'main.paperAtTea', label: 'Prophy @ TX' },
      { id: 'main.justPaper', label: 'Just Prophy' },
      { id: 'main.prophyTotal', label: 'Actual Prophy' },
    ],
  },
  {
    label: 'CRA / OE',
    options: [
      { id: 'table.coffeeNew', label: 'CRA (New)' },
      { id: 'table.coffeeReturn', label: 'CRA (Return)' },
      { id: 'table.coffeeTotal', label: 'CRA Total' },
      { id: 'table.coffeeNo', label: 'CRA (Not Billable)' },
      { id: 'table.renderedCoffee', label: 'Rendered CRA' },
      { id: 'table.coffeeYes', label: 'CRA (Billable)' },
      { id: 'table.orangeJuiceNew', label: 'OE (NP)' },
      { id: 'table.orangeJuiceReturn', label: 'OE (RC)' },
      { id: 'table.orangeJuiceTotal', label: 'OE Total' },
      { id: 'actual.orangeJuiceNew', label: 'Actual OE (NP)' },
      { id: 'actual.orangeJuiceReturn', label: 'Actual OE (RC)' },
    ],
  },
  {
    label: 'Sealant / Prophy',
    options: [
      { id: 'sugar.sugar', label: 'Sealant' },
      { id: 'sugar.sugarGood', label: 'Sealant (Billable)' },
      { id: 'sugar.sugarBad', label: 'Sealant (Redo)' },
      { id: 'sugar.paper', label: 'Prophy Documented' },
    ],
  },
  {
    label: 'Production / Visits',
    options: [
      { id: 'location.pineapple', label: 'Preventative Production' },
      { id: 'location.rose', label: 'Restorative Production' },
      { id: 'side.add', label: 'Add On' },
      { id: 'side.noShow', label: 'No Shows' },
      { id: 'side.scheduled', label: 'Scheduled' },
      { id: 'side.seen', label: 'Seen' },
      { id: 'side.seenPercent', label: 'Seen %' },
    ],
  },
  {
    label: 'Doctors Performance',
    options: [
      { id: 'docperf.doctorPreventative', label: 'Preventative Doctor' },
      { id: 'docperf.doctorRestorative', label: 'Restorative Doctor' },
      { id: 'docperf.doctorCraProduction', label: 'CRA Production Doctor' },
      { id: 'table.sales', label: 'Production Doctor' },
      { id: 'docperf.customer', label: 'Patient Seen' },
      { id: 'docperf.icecream', label: 'Insurance' },
      { id: 'docperf.cake', label: 'Cash' },
      { id: 'docperf.donut', label: 'Dentical' },
      { id: 'docperf.tart', label: 'Treatment' },
      { id: 'docperf.peach', label: 'Primary Teeth' },
      { id: 'docperf.peppermint', label: 'Permanent Teeth' },
    ],
  },
];

function collectNameBundleMetricOptions(): { id: ColumnFieldId; label: string }[] {
  const out: { id: ColumnFieldId; label: string }[] = [];
  for (const g of FIELD_GROUPS) {
    for (const opt of g.options) {
      const { id } = opt;
      if (id && (id.startsWith('table.') || id.startsWith('sugar.') || id.startsWith('docperf.'))) {
        out.push(opt);
      }
    }
  }
  return out;
}

function dedupeFieldOptions(
  options: { id: ColumnFieldId; label: string }[]
): { id: ColumnFieldId; label: string }[] {
  const seen = new Set<ColumnFieldId>();
  const out: { id: ColumnFieldId; label: string }[] = [];
  for (const opt of options) {
    if (!opt.id || seen.has(opt.id)) continue;
    seen.add(opt.id);
    out.push(opt);
  }
  return out;
}

const NAME_BUNDLE_METRIC_OPTIONS = dedupeFieldOptions(collectNameBundleMetricOptions());
const NAME_BUNDLE_MAX_METRICS = 12;

const POSITION_BUNDLE_METRIC_OPTIONS = NAME_BUNDLE_METRIC_OPTIONS;

const DOCPERF_EXTRA_INPUT_SUM_KEYS = new Set([
  'doctorPreventative',
  'doctorRestorative',
  'doctorCraProduction',
]);

const POSITION_BUNDLE_MAX_METRICS = 12;

const NUM_SLOTS = 7;

const EMPTY_COLUMN_SLOTS: SlotValue[] = Array.from({ length: NUM_SLOTS }, () => '' as SlotValue);

type NameBundleSlotConfig = { personName: string; metrics: ColumnFieldId[] };

type PositionBundleSlotConfig = { position: string; metrics: ColumnFieldId[] };

function createEmptyNameBundleSlotConfigs(): NameBundleSlotConfig[] {
  return Array.from({ length: NUM_SLOTS }, () => ({
    personName: '',
    metrics: ['docperf.customer' as ColumnFieldId],
  }));
}

function createEmptyPositionBundleSlotConfigs(): PositionBundleSlotConfig[] {
  return Array.from({ length: NUM_SLOTS }, () => ({
    position: '',
    metrics: ['docperf.customer' as ColumnFieldId],
  }));
}

function isShortProcBundle(slot: SlotValue): slot is ShortProcBundleId {
  return slot === 'bundle.spOE' || slot === 'bundle.spPro' || slot === 'bundle.spCRA';
}

function isOperatingBundle(slot: SlotValue): slot is OperatingBundleId {
  return slot === OPERATING_BUNDLE_ID;
}

/** Parse "8:30 AM", "17:00", etc. to minutes from midnight. */
function parseTimeToMinutes(value: string): number | null {
  const s = String(value ?? '').trim();
  if (!s) return null;
  const m12 = /^(\d{1,2}):(\d{2})\s*(AM|PM)$/i.exec(s);
  if (m12) {
    let h = Number(m12[1]);
    const min = Number(m12[2]);
    const ap = m12[3].toUpperCase();
    if (!Number.isFinite(h) || !Number.isFinite(min) || min < 0 || min > 59) return null;
    if (ap === 'PM' && h !== 12) h += 12;
    if (ap === 'AM' && h === 12) h = 0;
    if (h < 0 || h > 23) return null;
    return h * 60 + min;
  }
  const m24 = /^(\d{1,2}):(\d{2})$/.exec(s);
  if (m24) {
    const h = Number(m24[1]);
    const min = Number(m24[2]);
    if (!Number.isFinite(h) || !Number.isFinite(min) || h < 0 || h > 23 || min < 0 || min > 59) return null;
    return h * 60 + min;
  }
  return null;
}

function formatMinutesAsHrsMins(totalMinutes: number): string {
  if (!Number.isFinite(totalMinutes) || totalMinutes <= 0) return '';
  const wholeHours = Math.floor(totalMinutes / 60);
  const mins = totalMinutes % 60;
  const parts: string[] = [];
  if (wholeHours > 0) parts.push(`${wholeHours} hrs`);
  if (mins > 0) parts.push(`${mins} mins`);
  return parts.join(' ');
}

function computeHoursOpen(checkIn: string, checkOut: string): string {
  const inMin = parseTimeToMinutes(checkIn);
  const outMin = parseTimeToMinutes(checkOut);
  if (inMin === null || outMin === null) return '';
  let diff = outMin - inMin;
  if (diff < 0) diff += 24 * 60;
  if (diff <= 0) return '';
  return formatMinutesAsHrsMins(diff);
}

function parseHoursOpenToMinutes(hoursOpen: string): number | null {
  const s = String(hoursOpen ?? '').trim();
  if (!s) return null;
  const hrsMins = /(\d+)\s*hrs?\s+(\d+)\s*mins?/i.exec(s);
  if (hrsMins) {
    const h = Number(hrsMins[1]);
    const m = Number(hrsMins[2]);
    if (Number.isFinite(h) && Number.isFinite(m)) return h * 60 + m;
  }
  const hrsOnly = /^(\d+)\s*hrs?$/i.exec(s);
  if (hrsOnly) {
    const h = Number(hrsOnly[1]);
    if (Number.isFinite(h)) return h * 60;
  }
  const minsOnly = /^(\d+)\s*mins?$/i.exec(s);
  if (minsOnly) {
    const m = Number(minsOnly[1]);
    if (Number.isFinite(m)) return m;
  }
  const hm = /^(\d+):(\d{2})$/.exec(s);
  if (hm) {
    const h = Number(hm[1]);
    const m = Number(hm[2]);
    if (!Number.isFinite(h) || !Number.isFinite(m)) return null;
    return h * 60 + m;
  }
  const n = Number(s);
  if (Number.isFinite(n) && n > 0) return Math.round(n * 60);
  return null;
}

function getOperatingRowValues(doc: SimpleFormDoc): [string, string, string, string] {
  const checkIn = String(doc.checkIn ?? '').trim();
  const checkOut = String(doc.checkOut ?? '').trim();
  const hoursOpen = computeHoursOpen(checkIn, checkOut);
  const closer = String(doc.closer ?? '').trim();
  return [checkIn, checkOut, hoursOpen, closer];
}

function isNameBundle(slot: SlotValue): boolean {
  return slot === NAME_BUNDLE_ID;
}

function isPositionBundle(slot: SlotValue): boolean {
  return slot === POSITION_BUNDLE_ID;
}

const FIELD_SELECT_MIN_WIDTH_PX = 280;
const METRIC_SELECT_MIN_WIDTH_PX = 240;
const SINGLE_COLUMN_HEADER_MIN_WIDTH_PX = 300;

const FIELD_OPTION_LABEL_BY_ID: Map<string, string> = (() => {
  const map = new Map<string, string>();
  for (const g of FIELD_GROUPS) {
    for (const opt of g.options) {
      if (opt.id) map.set(opt.id, opt.label);
    }
  }
  return map;
})();

function getFieldOptionLabel(id: ColumnFieldId): string {
  return FIELD_OPTION_LABEL_BY_ID.get(id) ?? id;
}

function getSlotDisplayLabel(slot: SlotValue): string {
  if (!slot) return 'Select';
  if (isShortProcBundle(slot)) {
    return SHORT_PROC_BUNDLE_OPTIONS.find((o) => o.id === slot)?.label ?? slot;
  }
  if (isOperatingBundle(slot)) return 'Operating';
  if (isNameBundle(slot)) return 'Name';
  if (isPositionBundle(slot)) return 'Position';
  return getFieldOptionLabel(slot as ColumnFieldId);
}

const PRODUCTION_FILTER_SELECT_CSS = `
.production-filter-root .production-filter-toolbar select,
.production-filter-root .production-filter-toolbar input[type="month"] {
  width: 100%;
  min-width: 0;
  max-width: 100%;
  box-sizing: border-box;
  field-sizing: fixed;
}
.production-filter-root .production-filter-toolbar > * {
  min-width: 0;
}
.production-filter-root .production-filter-table select {
  field-sizing: content;
  min-width: ${FIELD_SELECT_MIN_WIDTH_PX}px;
  max-width: 100%;
  line-height: 1.35;
}
.production-filter-root .production-filter-table select.metric-field-select {
  min-width: ${METRIC_SELECT_MIN_WIDTH_PX}px;
}
`;

function bundleToMetric(bundle: ShortProcBundleId): 'orangeJuice' | 'paper' | 'coffee' {
  if (bundle === 'bundle.spOE') return 'orangeJuice';
  if (bundle === 'bundle.spPro') return 'paper';
  return 'coffee';
}

function getReasonCellValue(
  doc: SimpleFormDoc,
  reasonLabel: string,
  metric: 'orangeJuice' | 'paper' | 'coffee'
): string {
  const rows = Array.isArray(doc.reasonRows) ? doc.reasonRows : [];
  const row = rows.find((r) => String(r.reason).trim() === reasonLabel.trim());
  if (!row) return '';
  return String(row[metric] ?? '');
}

const subHeaderThStyle: React.CSSProperties = {
  fontSize: 10,
  fontWeight: 600,
  textAlign: 'left' as const,
  whiteSpace: 'normal' as const,
  wordBreak: 'break-word' as const,
  verticalAlign: 'bottom' as const,
  minWidth: METRIC_SELECT_MIN_WIDTH_PX,
  lineHeight: 1.25,
};

function operatingColumnThStyle(colIndex: number): React.CSSProperties {
  if (colIndex === OPERATING_HOURS_OPEN_COL_INDEX) {
    return {
      ...subHeaderThStyle,
      minWidth: OPERATING_HOURS_OPEN_MIN_WIDTH,
      maxWidth: 260,
      whiteSpace: 'nowrap',
    };
  }
  return subHeaderThStyle;
}

function operatingColumnTdStyle(colIndex: number, base: React.CSSProperties): React.CSSProperties {
  if (colIndex === OPERATING_HOURS_OPEN_COL_INDEX) {
    return {
      ...base,
      minWidth: OPERATING_HOURS_OPEN_MIN_WIDTH,
      whiteSpace: 'nowrap',
    };
  }
  return base;
}

function monthRange(ym: string): { start: string; end: string } {
  const [y, m] = ym.split('-').map((v) => Number(v));
  if (!y || !m) return { start: '', end: '' };
  const last = new Date(y, m, 0).getDate();
  const mm = String(m).padStart(2, '0');
  return {
    start: `${y}-${mm}-01`,
    end: `${y}-${mm}-${String(last).padStart(2, '0')}`,
  };
}

function filterDocsByLocationAndMonth(
  docs: SimpleFormDoc[],
  loc: string,
  start: string,
  end: string
): SimpleFormDoc[] {
  const trimmed = loc.trim();
  return docs.filter((d) => {
    const date = String(d.date ?? '');
    return String(d.location ?? '') === trimmed && date >= start && date <= end;
  });
}

function parseNumber(value: string | number | undefined): number {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function computeSeenPercentRounded(
  scheduledRaw: string | number | undefined,
  seenRaw: string | number | undefined
): string {
  const sch = parseNumber(scheduledRaw);
  if (sch === 0) return '-';
  const seen = parseNumber(seenRaw);
  return String(Math.round((seen / sch) * 100));
}

const CURRENCY_FIELD_IDS = new Set<ColumnFieldId>([
  'main.grandTotal',
  'main.coffeeSales',
  'main.salesWithoutCoffee',
  'location.pineapple',
  'location.rose',
  'location.total',
  'location.mailedProduction',
  'docperf.doctorPreventative',
  'docperf.doctorRestorative',
  'docperf.doctorCraProduction',
  'table.sales',
]);

function isCurrencyField(id: ColumnFieldId): boolean {
  return CURRENCY_FIELD_IDS.has(id);
}

const PERCENT_FIELD_IDS = new Set<ColumnFieldId>(['side.seenPercent']);

function isPercentField(id: ColumnFieldId): boolean {
  return PERCENT_FIELD_IDS.has(id);
}

function formatCurrencyDisplay(raw: string): string {
  const s = String(raw ?? '').trim();
  if (!s) return '';
  if (s.startsWith('$')) return s;
  return `$${s}`;
}

function formatPercentDisplay(raw: string): string {
  const s = String(raw ?? '').trim();
  if (!s || s === '-') return s;
  if (s.endsWith('%')) return s;
  return `${s}%`;
}

function formatMetricCellDisplay(metric: ColumnFieldId, raw: string): string {
  const s = String(raw ?? '').trim();
  if (!s) return '';
  if (isCurrencyField(metric)) return formatCurrencyDisplay(s);
  if (isPercentField(metric)) return formatPercentDisplay(s);
  return s;
}

function formatTotalCell(n: number, currency = false, percent = false): string {
  if (!Number.isFinite(n) || n === 0) return '';
  let out: string;
  if (Number.isInteger(n)) out = String(n);
  else {
    const rounded = Math.round(n * 100) / 100;
    out = String(rounded);
  }
  if (currency) return `$${out}`;
  if (percent) return `${out}%`;
  return out;
}

function computeRow(row: TableRow): TableRow {
  const coffeeHasInput = row.coffeeNew.trim() !== '' || row.coffeeReturn.trim() !== '';
  const orangeHasInput = row.orangeJuiceNew.trim() !== '' || row.orangeJuiceReturn.trim() !== '';
  const coffeeTotal = coffeeHasInput
    ? String(parseNumber(row.coffeeNew) + parseNumber(row.coffeeReturn))
    : '';
  const orangeJuiceTotal = orangeHasInput
    ? String(parseNumber(row.orangeJuiceNew) + parseNumber(row.orangeJuiceReturn))
    : '';
  return { ...row, coffeeTotal, orangeJuiceTotal };
}

function hasAnyTableRowValue(row: TableRow): boolean {
  return Object.values(row).some((value) => String(value).trim() !== '');
}

function hasAnySugarRowValue(row: SugarRow): boolean {
  return Object.values(row).some((value) => String(value).trim() !== '');
}

function getTableTotalsFromRows(rows: TableRow[]): TableTotals {
  return {
    sales: rows.reduce((sum, row) => sum + parseNumber(row.sales), 0),
    coffeeNew: rows.reduce((sum, row) => sum + parseNumber(row.coffeeNew), 0),
    coffeeReturn: rows.reduce((sum, row) => sum + parseNumber(row.coffeeReturn), 0),
    coffeeTotal: rows.reduce((sum, row) => sum + parseNumber(row.coffeeTotal), 0),
    coffeeNo: rows.reduce((sum, row) => sum + parseNumber(row.coffeeNo), 0),
    renderedCoffee: rows.reduce((sum, row) => sum + parseNumber(row.renderedCoffee), 0),
    coffeeYes: rows.reduce((sum, row) => sum + parseNumber(row.coffeeYes), 0),
    orangeJuiceNew: rows.reduce((sum, row) => sum + parseNumber(row.orangeJuiceNew), 0),
    orangeJuiceReturn: rows.reduce((sum, row) => sum + parseNumber(row.orangeJuiceReturn), 0),
    orangeJuiceTotal: rows.reduce((sum, row) => sum + parseNumber(row.orangeJuiceTotal), 0),
  };
}

function getSugarTotalsFromRows(rows: SugarRow[]): SugarTotals {
  return {
    sugar: rows.reduce((sum, row) => sum + parseNumber(row.sugar), 0),
    sugarGood: rows.reduce((sum, row) => sum + parseNumber(row.sugarGood), 0),
    sugarBad: rows.reduce((sum, row) => sum + parseNumber(row.sugarBad), 0),
    paper: rows.reduce((sum, row) => sum + parseNumber(row.paper), 0),
  };
}

const DOCPERF_ROW_KEYS = ['customer', 'icecream', 'cake', 'donut', 'tart', 'peach', 'peppermint'] as const;

type DocPerfRowKey = (typeof DOCPERF_ROW_KEYS)[number];

function getDoctorsPerfCellDisplay(doc: SimpleFormDoc, key: DocPerfRowKey): string {
  const rows = Array.isArray(doc.extraInputRows) ? doc.extraInputRows : [];
  if (rows.length === 0) return '';

  const hasAny = rows.some((r) => String((r as Record<string, unknown>)[key] ?? '').trim() !== '');
  if (!hasAny) return '';
  const sum = rows.reduce(
    (s, r) => s + parseNumber((r as Record<string, unknown>)[key] as string | number | undefined),
    0
  );
  return String(sum);
}

function coerceTableRow(row: Partial<TableRow> | undefined): TableRow {
  return {
    position: String(row?.position ?? ''),
    name: String(row?.name ?? ''),
    sales: String(row?.sales ?? ''),
    coffeeNew: String(row?.coffeeNew ?? ''),
    coffeeReturn: String(row?.coffeeReturn ?? ''),
    coffeeTotal: String(row?.coffeeTotal ?? ''),
    coffeeNo: String(row?.coffeeNo ?? ''),
    renderedCoffee: String(row?.renderedCoffee ?? ''),
    coffeeYes: String(row?.coffeeYes ?? ''),
    orangeJuiceNew: String(row?.orangeJuiceNew ?? ''),
    orangeJuiceReturn: String(row?.orangeJuiceReturn ?? ''),
    orangeJuiceTotal: String(row?.orangeJuiceTotal ?? ''),
  };
}

function findRowByTrimmedName<T extends { name?: string }>(rows: T[] | undefined, personName: string): T | undefined {
  const t = personName.trim();
  if (!t) return undefined;
  return (rows || []).find((r) => String(r.name ?? '').trim() === t);
}

function findRowByTrimmedNameAndPosition<T extends { name?: string; position?: string }>(
  rows: T[] | undefined,
  personName: string,
  position: string
): T | undefined {
  const t = personName.trim();
  const p = position.trim();
  if (!t || !p) return undefined;
  return (rows || []).find(
    (r) => String(r.name ?? '').trim() === t && String(r.position ?? '').trim() === p
  );
}

/** One cell: named row in CRA / sealant / doctors table for the chosen metric. */
function getNamedPersonRowMetricValue(doc: SimpleFormDoc, personName: string, metric: ColumnFieldId): string {
  const pn = personName.trim();
  if (!pn || !metric) return '';
  if (metric.startsWith('table.')) {
    const raw = findRowByTrimmedName(doc.tableRows, pn);
    if (!raw) return '';
    const key = metric.replace('table.', '') as keyof TableRow;
    const normalized = computeRow(coerceTableRow(raw as Partial<TableRow>));
    return String(normalized[key] ?? '');
  }
  if (metric.startsWith('sugar.')) {
    const raw = findRowByTrimmedName(doc.sugarRows, pn);
    if (!raw) return '';
    const key = metric.replace('sugar.', '') as keyof SugarRow;
    return String((raw as Record<string, unknown>)[key] ?? '');
  }
  if (metric.startsWith('docperf.')) {
    const raw = findRowByTrimmedName(doc.extraInputRows, pn);
    if (!raw) return '';
    const key = metric.replace('docperf.', '');
    return String((raw as Record<string, unknown>)[key] ?? '');
  }
  return '';
}

function getPositionPersonRowMetricValue(
  doc: SimpleFormDoc,
  position: string,
  personName: string,
  metric: ColumnFieldId
): string {
  const pos = position.trim();
  const pn = personName.trim();
  if (!pos || !pn || !metric) return '';
  if (metric.startsWith('table.')) {
    const raw = findRowByTrimmedNameAndPosition(doc.tableRows, pn, pos);
    if (!raw) return '';
    const key = metric.replace('table.', '') as keyof TableRow;
    const normalized = computeRow(coerceTableRow(raw as Partial<TableRow>));
    return String(normalized[key] ?? '');
  }
  if (metric.startsWith('sugar.')) {
    const raw = findRowByTrimmedNameAndPosition(doc.sugarRows, pn, pos);
    if (!raw) return '';
    const key = metric.replace('sugar.', '') as keyof SugarRow;
    return String((raw as Record<string, unknown>)[key] ?? '');
  }
  if (metric.startsWith('docperf.')) {
    const raw = findRowByTrimmedNameAndPosition(doc.extraInputRows, pn, pos);
    if (!raw) return '';
    const key = metric.replace('docperf.', '');
    return String((raw as Record<string, unknown>)[key] ?? '');
  }
  return '';
}

/** Sorted unique names that appear with `position` in any row (table / sealant / doctors) across `docs`. */
function collectFixedNamesForPosition(docs: SimpleFormDoc[], position: string): string[] {
  const p = position.trim();
  if (!p) return [];
  const set = new Set<string>();
  for (const doc of docs) {
    for (const r of doc.tableRows || []) {
      if (String(r.position ?? '').trim() === p) {
        const n = String(r.name ?? '').trim();
        if (n) set.add(n);
      }
    }
    for (const r of doc.sugarRows || []) {
      if (String(r.position ?? '').trim() === p) {
        const n = String(r.name ?? '').trim();
        if (n) set.add(n);
      }
    }
    for (const r of doc.extraInputRows || []) {
      if (String(r.position ?? '').trim() === p) {
        const n = String(r.name ?? '').trim();
        if (n) set.add(n);
      }
    }
  }
  return Array.from(set).sort((a, b) => a.localeCompare(b));
}

function resolveTotals(doc: SimpleFormDoc): { table: TableTotals; sugar: SugarTotals } {
  let table: TableTotals;
  if (doc.tableTotals && typeof doc.tableTotals === 'object') {
    const t = doc.tableTotals as Record<string, unknown>;
    table = {
      sales: parseNumber(t.sales as string | number | undefined),
      coffeeNew: parseNumber(t.coffeeNew as string | number | undefined),
      coffeeReturn: parseNumber(t.coffeeReturn as string | number | undefined),
      coffeeTotal: parseNumber(t.coffeeTotal as string | number | undefined),
      coffeeNo: parseNumber(t.coffeeNo as string | number | undefined),
      renderedCoffee: parseNumber(t.renderedCoffee as string | number | undefined),
      coffeeYes: parseNumber(t.coffeeYes as string | number | undefined),
      orangeJuiceNew: parseNumber(t.orangeJuiceNew as string | number | undefined),
      orangeJuiceReturn: parseNumber(t.orangeJuiceReturn as string | number | undefined),
      orangeJuiceTotal: parseNumber(t.orangeJuiceTotal as string | number | undefined),
    };
  } else if (Array.isArray(doc.tableRows) && doc.tableRows.length > 0) {
    const normalized = doc.tableRows.map((r) => computeRow({ ...r })).filter(hasAnyTableRowValue);
    table = getTableTotalsFromRows(normalized);
  } else {
    table = getTableTotalsFromRows([]);
  }

  let sugar: SugarTotals;
  if (doc.sugarTotals && typeof doc.sugarTotals === 'object') {
    const s = doc.sugarTotals as Record<string, unknown>;
    sugar = {
      sugar: parseNumber(s.sugar as string | number | undefined),
      sugarGood: parseNumber(s.sugarGood as string | number | undefined),
      sugarBad: parseNumber(s.sugarBad as string | number | undefined),
      paper: parseNumber(s.paper as string | number | undefined),
    };
  } else if (Array.isArray(doc.sugarRows) && doc.sugarRows.length > 0) {
    const normalized = doc.sugarRows.filter(hasAnySugarRowValue);
    sugar = getSugarTotalsFromRows(normalized);
  } else {
    sugar = getSugarTotalsFromRows([]);
  }

  return { table, sugar };
}

function formatDayColumn(dateStr: string): string {
  if (!dateStr || !/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return '—';
  const d = new Date(`${dateStr}T12:00:00`);
  if (Number.isNaN(d.getTime())) return '—';
  return new Intl.DateTimeFormat('en-US', { weekday: 'short' }).format(d);
}

function formatDateMDY(iso: string): string {
  const s = iso.trim();
  if (!s) return '—';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const [y, m, d] = s.split('-');
  if (!y || !m || !d) return s;
  return `${Number(m)}/${Number(d)}/${y}`;
}

function getFieldDisplay(
  doc: SimpleFormDoc,
  id: ColumnFieldId,
  totals: { table: TableTotals; sugar: SugarTotals }
): string {
  if (!id) return '';
  if (id.startsWith('location.')) {
    const key = id.replace('location.', '');
    const loc = doc.locationSummary;
    if (!loc || typeof loc !== 'object') return '';
    const v = (loc as Record<string, unknown>)[key];
    if (v === undefined || v === null) return '';
    return String(v);
  }
  if (id.startsWith('side.')) {
    const key = id.replace('side.', '');
    const side = doc.productionSideMetrics;
    if (!side || typeof side !== 'object') return '';
    if (key === 'seenPercent') {
      return computeSeenPercentRounded(side.scheduled, side.seen);
    }
    const v = (side as Record<string, unknown>)[key];
    if (v === undefined || v === null) return '';
    return String(v);
  }
  if (id.startsWith('actual.')) {
    const key = id.replace('actual.', '');
    const act = doc.coffeeActualTotals;
    if (!act || typeof act !== 'object') return '';
    const v = (act as Record<string, unknown>)[key];
    if (v === undefined || v === null) return '';
    return String(v);
  }
  if (id.startsWith('docperf.')) {
    const sub = id.replace('docperf.', '');
    if (DOCPERF_EXTRA_INPUT_SUM_KEYS.has(sub)) {
      const rows = Array.isArray(doc.extraInputRows) ? doc.extraInputRows : [];
      if (rows.length === 0) return '';
      const hasAny = rows.some((r) => String((r as Record<string, unknown>)[sub] ?? '').trim() !== '');
      if (!hasAny) return '';
      const sum = rows.reduce(
        (s, r) => s + parseNumber((r as Record<string, unknown>)[sub] as string | number | undefined),
        0
      );
      return String(sum);
    }
    if (!DOCPERF_ROW_KEYS.includes(sub as DocPerfRowKey)) return '';
    return getDoctorsPerfCellDisplay(doc, sub as DocPerfRowKey);
  }
  if (id.startsWith('main.')) {
    const key = id.replace('main.', '') as keyof SimpleFormDoc;
    const v = doc[key];
    if (v === undefined || v === null) return '';
    return String(v);
  }
  if (id.startsWith('table.')) {
    const key = id.replace('table.', '') as keyof TableTotals;
    return String(totals.table[key]);
  }
  if (id.startsWith('sugar.')) {
    const key = id.replace('sugar.', '') as keyof SugarTotals;
    return String(totals.sugar[key]);
  }
  return '';
}

export default function SimpleFormsDropdownViewPage() {
  const [location, setLocation] = useState('');
  const [month, setMonth] = useState(() => {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  });
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [rows, setRows] = useState<SimpleFormDoc[]>([]);
  const [columnSlots, setColumnSlots] = useState<SlotValue[]>(() => [...EMPTY_COLUMN_SLOTS]);
  const [nameBundleBySlot, setNameBundleBySlot] = useState<NameBundleSlotConfig[]>(() => createEmptyNameBundleSlotConfigs());
  const [positionBundleBySlot, setPositionBundleBySlot] = useState<PositionBundleSlotConfig[]>(
    () => createEmptyPositionBundleSlotConfigs()
  );

  const resetColumnConfiguration = useCallback(() => {
    setColumnSlots([...EMPTY_COLUMN_SLOTS]);
    setNameBundleBySlot(createEmptyNameBundleSlotConfigs());
    setPositionBundleBySlot(createEmptyPositionBundleSlotConfigs());
  }, []);

  const uniquePersonNames = useMemo(() => {
    const set = new Set<string>();
    for (const doc of rows) {
      for (const r of doc.tableRows || []) {
        const n = String(r.name ?? '').trim();
        if (n) set.add(n);
      }
      for (const r of doc.sugarRows || []) {
        const n = String(r.name ?? '').trim();
        if (n) set.add(n);
      }
      for (const r of doc.extraInputRows || []) {
        const n = String(r.name ?? '').trim();
        if (n) set.add(n);
      }
    }
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [rows]);

  const uniquePositions = useMemo(() => {
    const set = new Set<string>();
    for (const doc of rows) {
      for (const r of doc.tableRows || []) {
        const p = String(r.position ?? '').trim();
        if (p) set.add(p);
      }
      for (const r of doc.sugarRows || []) {
        const p = String(r.position ?? '').trim();
        if (p) set.add(p);
      }
      for (const r of doc.extraInputRows || []) {
        const p = String(r.position ?? '').trim();
        if (p) set.add(p);
      }
    }
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [rows]);

  const load = useCallback(async () => {
    if (!location.trim() || !month) {
      setError('Please select Office and month.');
      return;
    }
    const { start, end } = monthRange(month);
    if (!start || !end) {
      setError('Invalid month format.');
      return;
    }

    setLoading(true);
    setError('');
    try {
      let list: SimpleFormDoc[] = [];
      try {
        const q = query(
          collection(db, 'simple-forms'),
          where('location', '==', location.trim()),
          where('submittedDateTime', '!=', ''),
          where('date', '>=', start),
          where('date', '<=', end)
        );
        const snap = await getDocs(q);
        snap.forEach((docSnap) => {
          list.push({ id: docSnap.id, ...(docSnap.data() as object) } as SimpleFormDoc);
        });
      } catch {
        const allSnap = await getDocs(collection(db, 'simple-forms'));
        const all: SimpleFormDoc[] = [];
        allSnap.forEach((docSnap) => {
          const data = docSnap.data() as Record<string, unknown>;
          if (!data.submittedDateTime) return;
          all.push({ id: docSnap.id, ...data } as SimpleFormDoc);
        });
        list = filterDocsByLocationAndMonth(all, location, start, end);
      }
      list.sort((a, b) => String(a.date ?? '').localeCompare(String(b.date ?? '')));
      setRows(list);
      if (list.length === 0) {
        setError('No saved data matches these filters.');
      }
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : 'Unknown error';
      setError(`Failed to load: ${msg}`);
      setRows([]);
    } finally {
      setLoading(false);
    }
  }, [location, month]);

  useEffect(() => {
    if (!location.trim()) {
      setRows([]);
      setError('');
      return;
    }
    if (month) {
      void load();
    }
  }, [location, month, load]);

  const setSlot = (index: number, id: SlotValue) => {
    setColumnSlots((prev) => {
      const next = [...prev];
      next[index] = id;
      return next;
    });
    if (id === NAME_BUNDLE_ID) {
      setNameBundleBySlot((prev) => {
        const n = [...prev];
        const cur = n[index];
        if (!cur.metrics.length) {
          n[index] = { ...cur, metrics: ['docperf.customer' as ColumnFieldId] };
        }
        return n;
      });
    }
    if (id === POSITION_BUNDLE_ID) {
      setPositionBundleBySlot((prev) => {
        const n = [...prev];
        const cur = n[index];
        if (!cur.metrics.length) {
          n[index] = { ...cur, metrics: ['docperf.customer' as ColumnFieldId] };
        }
        return n;
      });
    }
  };

  /** Short proc / Name: 1 sub-header row. Position: 2 (metric groups, then names like Short Procedures). */
  const maxSubHeaderRows = useMemo(() => {
    let m = 0;
    for (const slot of columnSlots) {
      if (isShortProcBundle(slot) || isOperatingBundle(slot) || isNameBundle(slot)) m = Math.max(m, 1);
      if (isPositionBundle(slot)) m = Math.max(m, 2);
    }
    return m;
  }, [columnSlots]);

  const dayDateRowSpan = maxSubHeaderRows > 0 ? 1 + maxSubHeaderRows : 1;

  const footerSlotTotals = useMemo(() => {
    return columnSlots.map((slotId, slotIndex) => {
      if (!slotId) return { type: 'empty' as const };
      if (isShortProcBundle(slotId)) {
        const metric = bundleToMetric(slotId);
        return {
          type: 'bundle' as const,
          values: FIXED_REASON_OPTIONS.map((label) =>
            rows.reduce((sum, doc) => sum + parseNumber(getReasonCellValue(doc, label, metric)), 0)
          ),
        };
      }
      if (isOperatingBundle(slotId)) {
        const totalMinutes = rows.reduce((sum, doc) => {
          const [, , hoursOpen] = getOperatingRowValues(doc);
          const mins = parseHoursOpenToMinutes(hoursOpen);
          return sum + (mins ?? 0);
        }, 0);
        return {
          type: 'operating' as const,
          hoursOpenTotal: formatMinutesAsHrsMins(totalMinutes),
        };
      }
      if (isNameBundle(slotId)) {
        const cfg = nameBundleBySlot[slotIndex] ?? { personName: '', metrics: ['docperf.customer' as ColumnFieldId] };
        const metrics = cfg.metrics.length > 0 ? cfg.metrics : (['docperf.customer' as ColumnFieldId]);
        return {
          type: 'nameBundle' as const,
          values: metrics.map((metric) =>
            rows.reduce(
              (sum, doc) => sum + parseNumber(getNamedPersonRowMetricValue(doc, cfg.personName, metric)),
              0
            )
          ),
        };
      }
      if (isPositionBundle(slotId)) {
        const cfg = positionBundleBySlot[slotIndex] ?? {
          position: '',
          metrics: ['docperf.customer' as ColumnFieldId],
        };
        const metrics = cfg.metrics.length > 0 ? cfg.metrics : (['docperf.customer' as ColumnFieldId]);
        const people = collectFixedNamesForPosition(rows, cfg.position);
        const cols = people.length > 0 ? people : [''];
        const valuesByMetric = metrics.map((metric) =>
          cols.map((person) =>
            people.length === 0
              ? 0
              : rows.reduce(
                  (sum, doc) =>
                    sum + parseNumber(getPositionPersonRowMetricValue(doc, cfg.position, person, metric)),
                  0
                )
          )
        );
        return { type: 'positionBundle' as const, people, valuesByMetric };
      }
      const total = rows.reduce((sum, doc) => {
        const t = resolveTotals(doc);
        return sum + parseNumber(getFieldDisplay(doc, slotId as ColumnFieldId, t));
      }, 0);
      return { type: 'single' as const, total };
    });
  }, [rows, columnSlots, nameBundleBySlot, positionBundleBySlot]);

  const footerRowStyle: React.CSSProperties = {
    background: '#e8eef7',
    fontWeight: 700,
    color: '#0f172a',
  };

  const thSelectStyle: React.CSSProperties = {
    width: '100%',
    minWidth: FIELD_SELECT_MIN_WIDTH_PX,
    minHeight: 36,
    padding: '6px 8px',
    border: '1px solid #d1d5db',
    borderRadius: 6,
    background: '#fff',
    fontSize: 13,
    fontWeight: 600,
  };

  const metricSelectStyle: React.CSSProperties = {
    ...thSelectStyle,
    minWidth: METRIC_SELECT_MIN_WIDTH_PX,
  };

  const slotFieldSelect = (slotIndex: number, slotId: SlotValue) => (
    <select
      aria-label={`Column ${slotIndex + 1} field`}
      title={getSlotDisplayLabel(slotId)}
      value={slotId}
      onChange={(e) => setSlot(slotIndex, e.target.value as SlotValue)}
      style={thSelectStyle}
    >
      <option value="">Select</option>
      {FIELD_GROUPS.map((g, gi) => (
        <React.Fragment key={g.label}>
          <optgroup label={g.label}>
            {g.options.map((opt) => (
              <option key={`${g.label}-${opt.id}`} value={opt.id}>
                {opt.label}
              </option>
            ))}
          </optgroup>
          {gi === 0 ? (
            <optgroup label="Operating">
              <option value={OPERATING_BUNDLE_ID}>Operating</option>
            </optgroup>
          ) : null}
        </React.Fragment>
      ))}
      <optgroup label="Short Procedures (bundles)">
        {SHORT_PROC_BUNDLE_OPTIONS.map((opt) => (
          <option key={opt.id} value={opt.id}>
            {opt.label}
          </option>
        ))}
      </optgroup>
      <optgroup label="By name (bundle)">
        <option value={NAME_BUNDLE_ID}>Name</option>
      </optgroup>
      <optgroup label="By position (bundle)">
        <option value={POSITION_BUNDLE_ID}>Position</option>
      </optgroup>
    </select>
  );

  const cellStyle: React.CSSProperties = {
    border: '1px solid #e5e7eb',
    padding: '8px',
    fontSize: 14,
    verticalAlign: 'top',
  };

  return (
    <main
      className="production-filter-root"
      style={{
        minHeight: '100vh',
        display: 'grid',
        placeItems: 'start center',
        background: '#ffffff',
        padding: '48px 24px',
      }}
    >
      <style>{PRODUCTION_FILTER_SELECT_CSS}</style>
      <section
        style={{
          width: '100%',
          maxWidth: 2600,
          background: '#fff',
          padding: 20,
        }}
      >
        <h1 style={{ margin: '0 0 8px', fontSize: 28, fontWeight: 800, color: '#111827' }}>
          Production Filter
        </h1>

        <div
          className="production-filter-toolbar"
          style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))',
            gap: 12,
            marginBottom: 20,
            alignItems: 'end',
          }}
        >
          <div>
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Office</label>
            <select
              value={location}
              onChange={(e) => {
                resetColumnConfiguration();
                setLocation(e.target.value);
              }}
              style={{
                width: '100%',
                height: 40,
                padding: '0 10px',
                border: '1px solid #d1d5db',
                borderRadius: 8,
                background: '#fff',
              }}
            >
              <option value="">Select</option>
              {LOCATION_OPTIONS.map((loc) => (
                <option key={loc} value={loc}>
                  {loc}
                </option>
              ))}
            </select>
          </div>
          <div>
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Month</label>
            <input
              type="month"
              value={month}
              onChange={(e) => {
                resetColumnConfiguration();
                setMonth(e.target.value);
              }}
              style={{
                width: '100%',
                height: 40,
                padding: '0 10px',
                border: '1px solid #d1d5db',
                borderRadius: 8,
              }}
            />
          </div>
          <div>
            <button
              type="button"
              onClick={() => void load()}
              disabled={loading}
              style={{
                height: 40,
                padding: '0 16px',
                borderRadius: 8,
                border: '1px solid #2563eb',
                background: loading ? '#93c5fd' : '#2563eb',
                color: '#fff',
                fontWeight: 600,
                cursor: loading ? 'not-allowed' : 'pointer',
              }}
            >
              {loading ? 'Loading…' : 'Reload'}
            </button>
          </div>
        </div>

        {error ? <p style={{ color: '#b45309', marginBottom: 12, fontSize: 14 }}>{error}</p> : null}

        {!location.trim() ? (
          <p style={{ color: '#6b7280', fontSize: 14 }}>Select an office to load the table.</p>
        ) : null}

        {location.trim() && rows.length > 0 && (
          <div className="production-filter-table" style={{ overflowX: 'auto' }}>
            <table
              style={{
                width: '100%',
                borderCollapse: 'collapse',
                fontSize: 14,
              }}
            >
              <thead>
                <tr style={{ background: '#f3f4f6' }}>
                  <th
                    rowSpan={dayDateRowSpan}
                    style={{
                      ...cellStyle,
                      whiteSpace: 'nowrap',
                      minWidth: 72,
                    }}
                  >
                    Day
                  </th>
                  <th
                    rowSpan={dayDateRowSpan}
                    style={{
                      ...cellStyle,
                      whiteSpace: 'nowrap',
                      minWidth: 110,
                    }}
                  >
                    Date
                  </th>
                  {columnSlots.map((slotId, slotIndex) => {
                    if (isShortProcBundle(slotId)) {
                      return (
                        <th
                          key={`h-${slotIndex}`}
                          colSpan={BUNDLE_COLSPAN}
                          style={{ ...cellStyle, minWidth: 100 * BUNDLE_COLSPAN }}
                        >
                          {slotFieldSelect(slotIndex, slotId)}
                        </th>
                      );
                    }
                    if (isOperatingBundle(slotId)) {
                      return (
                        <th
                          key={`h-${slotIndex}`}
                          colSpan={OPERATING_COL_COUNT}
                          style={{
                            ...cellStyle,
                            minWidth: 100 * (OPERATING_COL_COUNT - 1) + OPERATING_HOURS_OPEN_MIN_WIDTH,
                          }}
                        >
                          {slotFieldSelect(slotIndex, slotId)}
                        </th>
                      );
                    }
                    if (isNameBundle(slotId)) {
                      const nbMetrics = nameBundleBySlot[slotIndex]?.metrics?.length
                        ? nameBundleBySlot[slotIndex].metrics.length
                        : 1;
                      const colSpan = Math.max(1, nbMetrics);
                      return (
                        <th
                          key={`h-${slotIndex}`}
                          colSpan={colSpan}
                          style={{ ...cellStyle, minWidth: 112 * colSpan }}
                        >
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                            {slotFieldSelect(slotIndex, slotId)}
                            <label style={{ fontSize: 11, fontWeight: 600, color: '#64748b' }}>Name</label>
                            <select
                              aria-label={`Column ${slotIndex + 1} person name`}
                              title={nameBundleBySlot[slotIndex]?.personName?.trim() || 'Select name…'}
                              value={nameBundleBySlot[slotIndex]?.personName ?? ''}
                              onChange={(e) => {
                                const v = e.target.value;
                                setNameBundleBySlot((prev) => {
                                  const n = [...prev];
                                  n[slotIndex] = { ...n[slotIndex], personName: v };
                                  return n;
                                });
                              }}
                              style={thSelectStyle}
                            >
                              <option value="">Select name…</option>
                              {uniquePersonNames.map((nm) => (
                                <option key={nm} value={nm}>
                                  {nm}
                                </option>
                              ))}
                            </select>
                            <div>
                              <button
                                type="button"
                                onClick={() => {
                                  setNameBundleBySlot((prev) => {
                                    const n = [...prev];
                                    const cur = n[slotIndex];
                                    if (cur.metrics.length >= NAME_BUNDLE_MAX_METRICS) return prev;
                                    const last = cur.metrics[cur.metrics.length - 1] ?? 'docperf.customer';
                                    n[slotIndex] = { ...cur, metrics: [...cur.metrics, last] };
                                    return n;
                                  });
                                }}
                                style={{
                                  fontSize: 12,
                                  padding: '4px 8px',
                                  borderRadius: 6,
                                  border: '1px solid #cbd5e1',
                                  background: '#fff',
                                  cursor: 'pointer',
                                  fontWeight: 600,
                                }}
                              >
                                + column
                              </button>
                            </div>
                          </div>
                        </th>
                      );
                    }
                    if (isPositionBundle(slotId)) {
                      const pbCfg = positionBundleBySlot[slotIndex] ?? {
                        position: '',
                        metrics: ['docperf.customer' as ColumnFieldId],
                      };
                      const pbMetrics = pbCfg.metrics.length > 0 ? pbCfg.metrics.length : 1;
                      const pbPeople = collectFixedNamesForPosition(rows, pbCfg.position);
                      const personColCount = Math.max(1, pbPeople.length);
                      const colSpan = Math.max(1, pbMetrics) * personColCount;
                      const fixedCount = pbPeople.length;
                      return (
                        <th
                          key={`h-${slotIndex}`}
                          colSpan={colSpan}
                          style={{ ...cellStyle, minWidth: 96 * colSpan }}
                        >
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                            {slotFieldSelect(slotIndex, slotId)}
                            <label style={{ fontSize: 11, fontWeight: 600, color: '#64748b' }}>Position</label>
                            <select
                              aria-label={`Column ${slotIndex + 1} position`}
                              title={pbCfg.position.trim() || 'Select position…'}
                              value={pbCfg.position}
                              onChange={(e) => {
                                const v = e.target.value;
                                setPositionBundleBySlot((prev) => {
                                  const n = [...prev];
                                  n[slotIndex] = { ...n[slotIndex], position: v };
                                  return n;
                                });
                              }}
                              style={thSelectStyle}
                            >
                              <option value="">Select position…</option>
                              {uniquePositions.map((pos) => (
                                <option key={pos} value={pos}>
                                  {pos}
                                </option>
                              ))}
                            </select>
                            <span style={{ fontSize: 11, color: '#94a3b8' }}>
                              {pbCfg.position.trim()
                                ? `${fixedCount} name${fixedCount === 1 ? '' : 's'}`
                                : 'Choose a position to list names'}
                            </span>
                            <div>
                              <button
                                type="button"
                                onClick={() => {
                                  setPositionBundleBySlot((prev) => {
                                    const n = [...prev];
                                    const cur = n[slotIndex];
                                    if (cur.metrics.length >= POSITION_BUNDLE_MAX_METRICS) return prev;
                                    const last = cur.metrics[cur.metrics.length - 1] ?? 'docperf.customer';
                                    n[slotIndex] = { ...cur, metrics: [...cur.metrics, last] };
                                    return n;
                                  });
                                }}
                                style={{
                                  fontSize: 12,
                                  padding: '4px 8px',
                                  borderRadius: 6,
                                  border: '1px solid #cbd5e1',
                                  background: '#fff',
                                  cursor: 'pointer',
                                  fontWeight: 600,
                                }}
                              >
                                + column
                              </button>
                            </div>
                          </div>
                        </th>
                      );
                    }
                    return (
                      <th
                        key={`h-${slotIndex}`}
                        rowSpan={dayDateRowSpan}
                        style={{ ...cellStyle, minWidth: SINGLE_COLUMN_HEADER_MIN_WIDTH_PX }}
                      >
                        {slotFieldSelect(slotIndex, slotId)}
                      </th>
                    );
                  })}
                </tr>
                {maxSubHeaderRows >= 1 ? (
                  <tr style={{ background: '#f9fafb' }}>
                    {columnSlots.flatMap((slotId, slotIndex) => {
                      if (isShortProcBundle(slotId)) {
                        return FIXED_REASON_OPTIONS.map((label, ri) => (
                          <th key={`h2-${slotIndex}-${ri}`} style={{ ...cellStyle, ...subHeaderThStyle }}>
                            {label}
                          </th>
                        ));
                      }
                      if (isOperatingBundle(slotId)) {
                        return OPERATING_SUB_LABELS.map((label, ri) => (
                          <th
                            key={`h2-op-${slotIndex}-${ri}`}
                            style={{ ...cellStyle, ...operatingColumnThStyle(ri) }}
                          >
                            {label}
                          </th>
                        ));
                      }
                      if (isNameBundle(slotId)) {
                        const cfg = nameBundleBySlot[slotIndex] ?? {
                          personName: '',
                          metrics: ['docperf.customer' as ColumnFieldId],
                        };
                        const metrics = cfg.metrics.length > 0 ? cfg.metrics : (['docperf.customer' as ColumnFieldId]);
                        return metrics.map((metric, mi) => (
                          <th key={`h2-${slotIndex}-${mi}`} style={{ ...cellStyle, ...subHeaderThStyle }}>
                            <select
                              className="metric-field-select"
                              aria-label={`Column ${slotIndex + 1} metric ${mi + 1}`}
                              title={getFieldOptionLabel(metric)}
                              value={metric}
                              onChange={(e) => {
                                const nextM = e.target.value as ColumnFieldId;
                                setNameBundleBySlot((prev) => {
                                  const n = [...prev];
                                  const row = { ...n[slotIndex] };
                                  const nextMetrics = [...row.metrics];
                                  nextMetrics[mi] = nextM;
                                  n[slotIndex] = { ...row, metrics: nextMetrics };
                                  return n;
                                });
                              }}
                              style={metricSelectStyle}
                            >
                              {NAME_BUNDLE_METRIC_OPTIONS.map((opt) => (
                                <option key={`nb-${slotIndex}-${mi}-${opt.id}`} value={opt.id}>
                                  {opt.label}
                                </option>
                              ))}
                            </select>
                            {metrics.length > 1 ? (
                              <button
                                type="button"
                                onClick={() => {
                                  setNameBundleBySlot((prev) => {
                                    const n = [...prev];
                                    const cur = n[slotIndex];
                                    if (cur.metrics.length <= 1) return prev;
                                    const nextMetrics = cur.metrics.filter((_, idx) => idx !== mi);
                                    n[slotIndex] = { ...cur, metrics: nextMetrics };
                                    return n;
                                  });
                                }}
                                style={{
                                  display: 'block',
                                  marginTop: 6,
                                  fontSize: 11,
                                  padding: '2px 6px',
                                  borderRadius: 4,
                                  border: '1px solid #e5e7eb',
                                  background: '#fff',
                                  cursor: 'pointer',
                                  color: '#64748b',
                                }}
                              >
                                Remove column
                              </button>
                            ) : null}
                          </th>
                        ));
                      }
                      if (isPositionBundle(slotId)) {
                        const cfg = positionBundleBySlot[slotIndex] ?? {
                          position: '',
                          metrics: ['docperf.customer' as ColumnFieldId],
                        };
                        const metrics = cfg.metrics.length > 0 ? cfg.metrics : (['docperf.customer' as ColumnFieldId]);
                        const people = collectFixedNamesForPosition(rows, cfg.position);
                        const span = Math.max(1, people.length);
                        return metrics.map((metric, mi) => (
                          <th
                            key={`h2-pb-m-${slotIndex}-${mi}`}
                            colSpan={span}
                            style={{ ...cellStyle, ...subHeaderThStyle, minWidth: 92 * span }}
                          >
                            <select
                              className="metric-field-select"
                              aria-label={`Column ${slotIndex + 1} position bundle metric ${mi + 1}`}
                              title={getFieldOptionLabel(metric)}
                              value={metric}
                              onChange={(e) => {
                                const nextM = e.target.value as ColumnFieldId;
                                setPositionBundleBySlot((prev) => {
                                  const n = [...prev];
                                  const row = { ...n[slotIndex] };
                                  const nextMetrics = [...row.metrics];
                                  nextMetrics[mi] = nextM;
                                  n[slotIndex] = { ...row, metrics: nextMetrics };
                                  return n;
                                });
                              }}
                              style={metricSelectStyle}
                            >
                              {POSITION_BUNDLE_METRIC_OPTIONS.map((opt) => (
                                <option key={`pb-${slotIndex}-${mi}-${opt.id}`} value={opt.id}>
                                  {opt.label}
                                </option>
                              ))}
                            </select>
                            {metrics.length > 1 ? (
                              <button
                                type="button"
                                onClick={() => {
                                  setPositionBundleBySlot((prev) => {
                                    const n = [...prev];
                                    const cur = n[slotIndex];
                                    if (cur.metrics.length <= 1) return prev;
                                    const nextMetrics = cur.metrics.filter((_, idx) => idx !== mi);
                                    n[slotIndex] = { ...cur, metrics: nextMetrics };
                                    return n;
                                  });
                                }}
                                style={{
                                  display: 'block',
                                  marginTop: 6,
                                  fontSize: 11,
                                  padding: '2px 6px',
                                  borderRadius: 4,
                                  border: '1px solid #e5e7eb',
                                  background: '#fff',
                                  cursor: 'pointer',
                                  color: '#64748b',
                                }}
                              >
                                Remove column
                              </button>
                            ) : null}
                          </th>
                        ));
                      }
                      return [];
                    })}
                  </tr>
                ) : null}
                {maxSubHeaderRows >= 2 ? (
                  <tr style={{ background: '#f9fafb' }}>
                    {columnSlots.flatMap((slotId, slotIndex) => {
                      if (isShortProcBundle(slotId)) {
                        return FIXED_REASON_OPTIONS.map((label, ri) => (
                          <th
                            key={`h3-sp-${slotIndex}-${ri}`}
                            style={{ ...cellStyle, ...subHeaderThStyle, color: '#cbd5e1', fontSize: 10 }}
                          >
                            {'\u00a0'}
                          </th>
                        ));
                      }
                      if (isOperatingBundle(slotId)) {
                        return OPERATING_SUB_LABELS.map((label, ri) => (
                          <th
                            key={`h3-op-${slotIndex}-${ri}`}
                            style={{
                              ...cellStyle,
                              ...operatingColumnThStyle(ri),
                              color: '#cbd5e1',
                              fontSize: 10,
                            }}
                          >
                            {'\u00a0'}
                          </th>
                        ));
                      }
                      if (isNameBundle(slotId)) {
                        const cfg = nameBundleBySlot[slotIndex] ?? {
                          personName: '',
                          metrics: ['docperf.customer' as ColumnFieldId],
                        };
                        const metrics = cfg.metrics.length > 0 ? cfg.metrics : (['docperf.customer' as ColumnFieldId]);
                        return metrics.map((_, mi) => (
                          <th
                            key={`h3-nb-${slotIndex}-${mi}`}
                            style={{ ...cellStyle, ...subHeaderThStyle, color: '#cbd5e1', fontSize: 10 }}
                          >
                            {'\u00a0'}
                          </th>
                        ));
                      }
                      if (isPositionBundle(slotId)) {
                        const cfg = positionBundleBySlot[slotIndex] ?? {
                          position: '',
                          metrics: ['docperf.customer' as ColumnFieldId],
                        };
                        const metrics = cfg.metrics.length > 0 ? cfg.metrics : (['docperf.customer' as ColumnFieldId]);
                        const people = collectFixedNamesForPosition(rows, cfg.position);
                        if (people.length === 0) {
                          return metrics.map((_, mi) => (
                            <th key={`h3-pb-${slotIndex}-${mi}-0`} style={{ ...cellStyle, ...subHeaderThStyle, minWidth: 88 }}>
                              —
                            </th>
                          ));
                        }
                        return metrics.flatMap((_, mi) =>
                          people.map((person, pi) => (
                            <th key={`h3-pb-${slotIndex}-${mi}-${pi}`} style={{ ...cellStyle, ...subHeaderThStyle, minWidth: 88 }}>
                              {person}
                            </th>
                          ))
                        );
                      }
                      return [];
                    })}
                  </tr>
                ) : null}
              </thead>
              <tbody>
                {rows.map((doc) => {
                  const totals = resolveTotals(doc);
                  const dateStr = String(doc.date ?? '');
                  return (
                    <tr key={doc.id}>
                      <td style={{ ...cellStyle, fontWeight: 600, color: '#374151' }}>
                        {formatDayColumn(dateStr)}
                      </td>
                      <td style={cellStyle}>{formatDateMDY(dateStr)}</td>
                      {columnSlots.flatMap((slotId, slotIndex) => {
                        if (isShortProcBundle(slotId)) {
                          return FIXED_REASON_OPTIONS.map((label) => (
                            <td key={`c-${doc.id}-${slotIndex}-${label}`} style={cellStyle}>
                              {getReasonCellValue(doc, label, bundleToMetric(slotId))}
                            </td>
                          ));
                        }
                        if (isOperatingBundle(slotId)) {
                          const [checkIn, checkOut, hoursOpen, closer] = getOperatingRowValues(doc);
                          const values = [checkIn, checkOut, hoursOpen, closer];
                          return values.map((val, ci) => (
                            <td
                              key={`c-op-${doc.id}-${slotIndex}-${ci}`}
                              style={operatingColumnTdStyle(ci, cellStyle)}
                            >
                              {String(val).trim() !== '' ? val : '—'}
                            </td>
                          ));
                        }
                        if (isNameBundle(slotId)) {
                          const cfg = nameBundleBySlot[slotIndex] ?? {
                            personName: '',
                            metrics: ['docperf.customer' as ColumnFieldId],
                          };
                          const metrics = cfg.metrics.length > 0 ? cfg.metrics : (['docperf.customer' as ColumnFieldId]);
                          return metrics.map((metric, mi) => {
                            const raw = getNamedPersonRowMetricValue(doc, cfg.personName, metric);
                            const show = String(raw).trim() !== '';
                            return (
                              <td key={`c-${doc.id}-${slotIndex}-${mi}`} style={cellStyle}>
                                {show ? formatMetricCellDisplay(metric, raw) : '—'}
                              </td>
                            );
                          });
                        }
                        if (isPositionBundle(slotId)) {
                          const cfg = positionBundleBySlot[slotIndex] ?? {
                            position: '',
                            metrics: ['docperf.customer' as ColumnFieldId],
                          };
                          const metrics = cfg.metrics.length > 0 ? cfg.metrics : (['docperf.customer' as ColumnFieldId]);
                          const people = collectFixedNamesForPosition(rows, cfg.position);
                          const cols = people.length > 0 ? people : [''];
                          return metrics.flatMap((metric, mi) =>
                            cols.map((person, pi) => {
                              const emptySlot = people.length === 0;
                              const raw = emptySlot
                                ? ''
                                : getPositionPersonRowMetricValue(doc, cfg.position, person, metric);
                              const show = String(raw).trim() !== '';
                              return (
                                <td key={`c-pb-${doc.id}-${slotIndex}-${mi}-${pi}`} style={cellStyle}>
                                  {emptySlot ? '—' : show ? formatMetricCellDisplay(metric, raw) : '—'}
                                </td>
                              );
                            })
                          );
                        }
                        const fieldId = slotId as ColumnFieldId;
                        const raw = getFieldDisplay(doc, fieldId, totals);
                        const show = String(raw).trim() !== '';
                        return [
                          <td key={`c-${doc.id}-${slotIndex}`} style={cellStyle}>
                            {show ? formatMetricCellDisplay(fieldId, raw) : '—'}
                          </td>,
                        ];
                      })}
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr style={footerRowStyle}>
                  <td style={{ ...cellStyle, whiteSpace: 'nowrap' }}>Total</td>
                  <td style={{ ...cellStyle, color: '#64748b' }}>—</td>
                  {footerSlotTotals.flatMap((fv, slotIndex) => {
                    const slotId = columnSlots[slotIndex] as ColumnFieldId;
                    if (fv.type === 'empty') {
                      return [<td key={`ft-${slotIndex}`} style={cellStyle} />];
                    }
                    if (fv.type === 'bundle') {
                      return fv.values.map((n, ri) => (
                        <td key={`ft-${slotIndex}-${ri}`} style={cellStyle}>
                          {formatTotalCell(n)}
                        </td>
                      ));
                    }
                    if (fv.type === 'operating') {
                      return OPERATING_SUB_LABELS.map((_, ri) => (
                        <td key={`ft-op-${slotIndex}-${ri}`} style={operatingColumnTdStyle(ri, cellStyle)}>
                          {ri === OPERATING_HOURS_OPEN_COL_INDEX ? fv.hoursOpenTotal || '—' : '—'}
                        </td>
                      ));
                    }
                    if (fv.type === 'nameBundle') {
                      const nbCfg = nameBundleBySlot[slotIndex];
                      const nbMetrics =
                        nbCfg && nbCfg.metrics.length > 0
                          ? nbCfg.metrics
                          : (['docperf.customer' as ColumnFieldId]);
                      return fv.values.map((n, ri) => (
                        <td key={`ft-${slotIndex}-nb-${ri}`} style={cellStyle}>
                          {formatTotalCell(n, isCurrencyField(nbMetrics[ri] ?? ''))}
                        </td>
                      ));
                    }
                    if (fv.type === 'positionBundle') {
                      const pbCfg = positionBundleBySlot[slotIndex];
                      const pbMetrics =
                        pbCfg && pbCfg.metrics.length > 0
                          ? pbCfg.metrics
                          : (['docperf.customer' as ColumnFieldId]);
                      return fv.valuesByMetric.flatMap((perPerson, mi) =>
                        perPerson.map((n, pi) => {
                          const formatted = formatTotalCell(n, isCurrencyField(pbMetrics[mi] ?? ''));
                          return (
                            <td key={`ft-${slotIndex}-pb-${mi}-${pi}`} style={cellStyle}>
                              {fv.people.length === 0 ? '—' : formatted || '—'}
                            </td>
                          );
                        })
                      );
                    }
                    return [
                      <td key={`ft-${slotIndex}`} style={cellStyle}>
                        {formatTotalCell(
                          fv.total,
                          isCurrencyField(slotId),
                          isPercentField(slotId)
                        )}
                      </td>,
                    ];
                  })}
                </tr>
              </tfoot>
            </table>
          </div>
        )}

        {location.trim() && rows.length > 0 && (
          <div
            style={{
              marginTop: 16,
              display: 'flex',
              flexWrap: 'wrap',
              alignItems: 'center',
              gap: 12,
            }}
          >
            <span style={{ fontSize: 14, color: '#6b7280' }}>
              {rows.length} day{rows.length === 1 ? '' : 's'} · {NUM_SLOTS} column slots
            </span>
            <button
              type="button"
              onClick={resetColumnConfiguration}
              style={{
                height: 34,
                padding: '0 12px',
                borderRadius: 8,
                border: '1px solid #cbd5e1',
                background: '#fff',
                fontWeight: 600,
                cursor: 'pointer',
                fontSize: 13,
              }}
            >
              Reset columns
            </button>
          </div>
        )}

      </section>
    </main>
  );
}