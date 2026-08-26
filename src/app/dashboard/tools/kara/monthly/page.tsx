'use client';

import React, { Suspense, useEffect, useMemo, useState } from 'react';
import { useSearchParams } from 'next/navigation';
import { auth, db } from '@/lib/firebase.config';
import { onAuthStateChanged } from 'firebase/auth';
import { collection, doc, getDoc, getDocs, serverTimestamp, setDoc } from 'firebase/firestore';
import { Document, Page, StyleSheet, Text, View, pdf } from '@react-pdf/renderer';

type TableRow = {
  position?: string;
  name?: string;
  sales?: string;
  coffeeNew?: string;
  coffeeReturn?: string;
  coffeeTotal?: string;
  coffeeNo?: string;
  renderedCoffee?: string;
  coffeeYes?: string;
  orangeJuiceNew?: string;
  orangeJuiceReturn?: string;
  orangeJuiceTotal?: string;
};

type SugarRow = {
  position?: string;
  name?: string;
  sugar?: string;
  sugarGood?: string;
  sugarBad?: string;
  paper?: string;
};

type DiagnoseRow = {
  sName?: string;
  diagnose?: string;
};

type ReasonRow = {
  reason?: string;
  orangeJuice?: string;
  paper?: string;
  coffee?: string;
};

type ExtraInputRow = {
  position?: string;
  name?: string;
  doctorPreventative?: string;
  doctorRestorative?: string;
  doctorCraProduction?: string;
  customer?: string;
  icecream?: string;
  cake?: string;
  donut?: string;
  tart?: string;
  peach?: string;
  peppermint?: string;
  total?: string;
};

type FormDoc = {
  id: string;
  date?: string;
  location?: string;
  submittedDateTime?: string;
  grandTotal?: string;
  coffeeSales?: string;
  salesWithoutCoffee?: string;
  paperAtOrangeJuice?: string;
  paperAtTea?: string;
  justPaper?: string;
  prophyTotal?: string;
  tableRows?: TableRow[];
  sugarRows?: SugarRow[];
  diagnoseRows?: DiagnoseRow[];
  reasonRows?: ReasonRow[];
  extraInputRows?: ExtraInputRow[];
  locationSummary?: { pineapple?: string; rose?: string; mailedProduction?: string };
  coffeeActualTotals?: { orangeJuiceNew?: string; orangeJuiceReturn?: string };
  productionSideMetrics?: {
    add?: unknown;
    noShow?: unknown;
    scheduled?: unknown;
    seen?: unknown;
    referral?: unknown;
    postcard?: unknown;
  };
};

type Aggregate = {
  key: string;
  month: string;
  location: string;
  docCount: number;
  grandTotal: number;
  coffeeSales: number;
  salesWithoutCoffee: number;
  paperAtOrangeJuice: number;
  paperAtTea: number;
  justPaper: number;
  prophyTotal: number;
  actualOrangeJuiceNew: number;
  actualOrangeJuiceReturn: number;
  pineapple: number;
  rose: number;
  mailedProduction: number;
  visitsAdd: number;
  visitsNoShow: number;
  visitsScheduled: number;
  visitsSeen: number;
  visitsReferral: number;
  visitsPostcard: number;
  coffeeRows: Array<{ position: string; name: string; doctorPreventative: number; doctorRestorative: number; doctorCraProduction: number; sales: number; coffeeNew: number; coffeeReturn: number; coffeeTotal: number; coffeeNo: number; renderedCoffee: number; coffeeYes: number; orangeJuiceNew: number; orangeJuiceReturn: number; orangeJuiceTotal: number }>;
  additionalRows: Array<{ position: string; name: string; customer: number; icecream: number; cake: number; donut: number; tart: number; peach: number; peppermint: number; total: number }>;
  sugarRows: Array<{ position: string; name: string; sugar: number; sugarGood: number; sugarBad: number; paper: number }>;
  diagnoseRows: Array<{ sName: string; diagnose: number }>;
  reasonRows: Array<{ reason: string; orangeJuice: number; paper: number; coffee: number }>;
};

type AggregateCoffeeRow = Aggregate['coffeeRows'][number];
type AggregateAdditionalRow = Aggregate['additionalRows'][number];
type AggregateSugarRow = Aggregate['sugarRows'][number];
type AggregateDiagnoseRow = Aggregate['diagnoseRows'][number];
type AggregateReasonRow = Aggregate['reasonRows'][number];

const tableCellStyle: React.CSSProperties = { border: '1px solid #e5e7eb', padding: 8 };
const tableHeaderStyle: React.CSSProperties = { ...tableCellStyle, textAlign: 'left', background: '#f3f4f6' };
const tableFooterRowStyle: React.CSSProperties = { background: '#f9fafb', fontWeight: 700 };
const tableSubtotalRowStyle: React.CSSProperties = { background: '#f9fafb', fontWeight: 600 };
const tableEmptyCellStyle: React.CSSProperties = { ...tableCellStyle, padding: 10, color: '#6b7280' };

const monthlyPdfStyles = StyleSheet.create({
  page: {
    padding: 20,
    fontSize: 8,
    color: '#111827',
  },
  title: {
    fontSize: 14,
    fontWeight: 700,
    marginBottom: 6,
  },
  meta: {
    marginBottom: 10,
  },
  sectionTitle: {
    fontSize: 10,
    fontWeight: 700,
    marginTop: 8,
    marginBottom: 4,
  },
  table: {
    width: '100%',
    borderTopWidth: 1,
    borderLeftWidth: 1,
    borderColor: '#d1d5db',
    borderStyle: 'solid',
    marginBottom: 8,
  },
  row: {
    flexDirection: 'row',
    width: '100%',
  },
  headerRow: {
    backgroundColor: '#f3f4f6',
  },
  totalRow: {
    backgroundColor: '#f9fafb',
  },
  positionGroupSeparatorCell: {
    borderTopWidth: 2,
    borderTopColor: '#9ca3af',
  },
  cell: {
    borderRightWidth: 1,
    borderBottomWidth: 1,
    borderRightColor: '#d1d5db',
    borderBottomColor: '#d1d5db',
    borderStyle: 'solid',
    paddingVertical: 3,
    paddingHorizontal: 2,
    minWidth: 0,
    flexShrink: 1,
  },
  cellText: {
    fontSize: 8,
  },
  cellTextCompact: {
    fontSize: 7,
  },
});

function pdfTableCellStyle(flex: number) {
  return [monthlyPdfStyles.cell, { flex }];
}

type PdfExtraCellStyle = typeof monthlyPdfStyles.cell | typeof monthlyPdfStyles.positionGroupSeparatorCell;

function renderPdfTableCell(
  content: string,
  flex: number,
  key: string,
  compact?: boolean,
  extraCellStyle?: PdfExtraCellStyle
) {
  const cellStyle = extraCellStyle
    ? [...pdfTableCellStyle(flex), extraCellStyle]
    : pdfTableCellStyle(flex);
  const textStyle = compact ? monthlyPdfStyles.cellTextCompact : monthlyPdfStyles.cellText;

  return (
    <View key={key} style={cellStyle}>
      <Text style={textStyle}>{content}</Text>
    </View>
  );
}

type PdfTableDataRowMeta = { positionGroupStart?: boolean };
function renderPdfTableInner(
  headers: string[],
  rows: string[][],
  widths: number[],
  summaryRows?: string[][],
  compact?: boolean,
  dataRowMeta?: PdfTableDataRowMeta[]
) {
  const columnFlex = widths.length > 0 ? widths : headers.map(() => 1);

  return (
    <View style={monthlyPdfStyles.table}>
      <View style={[monthlyPdfStyles.row, monthlyPdfStyles.headerRow]}>
        {headers.map((h, idx) => renderPdfTableCell(h, columnFlex[idx] ?? 1, `h_${h}_${idx}`, compact))}
      </View>
      {rows.length === 0 ? (
        <View style={monthlyPdfStyles.row}>
          {renderPdfTableCell('No data', columnFlex.reduce((sum, w) => sum + w, 0) || 1, 'no_data', compact)}
        </View>
      ) : (
        rows.map((row, rowIdx) => {
          const positionGroupStart = dataRowMeta?.[rowIdx]?.positionGroupStart;
          const cellSeparatorStyle = positionGroupStart ? monthlyPdfStyles.positionGroupSeparatorCell : undefined;
          return (
            <View key={`r_${rowIdx}`} style={monthlyPdfStyles.row}>
              {row.map((c, idx) =>
                renderPdfTableCell(c, columnFlex[idx] ?? 1, `c_${rowIdx}_${idx}`, compact, cellSeparatorStyle)
              )}
            </View>
          );
        })
      )}
      {(summaryRows || []).map((summaryRow, rowIdx) => (
          <View key={`s_${rowIdx}`} style={[monthlyPdfStyles.row, monthlyPdfStyles.totalRow]}>
            {summaryRow.map((c, idx) =>
              renderPdfTableCell(c, columnFlex[idx] ?? 1, `s_${rowIdx}_${idx}`, compact)
            )}
          </View>
        ))}
    </View>
  );
}

const PDF_TABLE_MAX_DATA_ROWS = 18;

function renderPdfTable(
  headers: string[],
  rows: string[][],
  widths: number[],
  summaryRows?: string[][],
  sectionTitle?: string | null,
  compact?: boolean,
  dataRowMeta?: PdfTableDataRowMeta[]
) {
  const rowBlocks: string[][][] = [];
  if (rows.length === 0) {
    rowBlocks.push([]);
  } else {
    for (let i = 0; i < rows.length; i += PDF_TABLE_MAX_DATA_ROWS) {
      rowBlocks.push(rows.slice(i, i + PDF_TABLE_MAX_DATA_ROWS));
    }
  }

  return (
    <View>
      {rowBlocks.map((blockRows, blockIdx) => (
        <View key={`pdf_table_${sectionTitle ?? 'table'}_${blockIdx}`} wrap={false}>
          {blockIdx === 0 && sectionTitle ? (
            <Text style={monthlyPdfStyles.sectionTitle}>{sectionTitle}</Text>
          ) : null}
          {renderPdfTableInner(
            headers,
            blockRows,
            widths,
            blockIdx === rowBlocks.length - 1 ? summaryRows : undefined,
            compact,
            blockIdx === 0 ? dataRowMeta?.slice(0, blockRows.length) : dataRowMeta?.slice(
              rowBlocks.slice(0, blockIdx).reduce((sum, block) => sum + block.length, 0),
              rowBlocks.slice(0, blockIdx).reduce((sum, block) => sum + block.length, 0) + blockRows.length
            )
          )}
        </View>
      ))}
    </View>
  );
}

function parseNumber(value: unknown): number {
  const n = Number(value);
  if (!Number.isFinite(n)) return 0;
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

function addAmount(base: number, value: unknown): number {
  return Math.round((base + parseNumber(value) + Number.EPSILON) * 100) / 100;
}

function formatAmount(value: number): string {
  const rounded = Math.round((value + Number.EPSILON) * 100) / 100;
  return rounded.toLocaleString('en-US', {
    minimumFractionDigits: Number.isInteger(rounded) ? 0 : 2,
    maximumFractionDigits: Number.isInteger(rounded) ? 0 : 2,
  });
}

function sumRowsAmount<T>(rows: readonly T[], pick: (row: T) => number): number {
  return rows.reduce((sum, row) => addAmount(sum, pick(row)), 0);
}

function sumCoffeeRows(rows: AggregateCoffeeRow[], pick: (row: AggregateCoffeeRow) => number): number {
  return sumRowsAmount(rows, pick);
}

function sumAdditionalRows(rows: AggregateAdditionalRow[], pick: (row: AggregateAdditionalRow) => number): number {
  return sumRowsAmount(rows, pick);
}

function sumSugarRows(rows: AggregateSugarRow[], pick: (row: AggregateSugarRow) => number): number {
  return sumRowsAmount(rows, pick);
}

const EMPTY_COFFEE_ROWS: AggregateCoffeeRow[] = [];
const EMPTY_ADDITIONAL_ROWS: AggregateAdditionalRow[] = [];
const EMPTY_SUGAR_ROWS: AggregateSugarRow[] = [];

function sumCoffeeRowFields(rows: AggregateCoffeeRow[]) {
  return {
    sales: sumCoffeeRows(rows, (r) => r.sales),
    coffeeNew: sumCoffeeRows(rows, (r) => r.coffeeNew),
    coffeeReturn: sumCoffeeRows(rows, (r) => r.coffeeReturn),
    coffeeTotal: sumCoffeeRows(rows, (r) => r.coffeeTotal),
    coffeeNo: sumCoffeeRows(rows, (r) => r.coffeeNo),
    renderedCoffee: sumCoffeeRows(rows, (r) => r.renderedCoffee),
    coffeeYes: sumCoffeeRows(rows, (r) => r.coffeeYes),
    orangeJuiceNew: sumCoffeeRows(rows, (r) => r.orangeJuiceNew),
    orangeJuiceReturn: sumCoffeeRows(rows, (r) => r.orangeJuiceReturn),
    orangeJuiceTotal: sumCoffeeRows(rows, (r) => r.orangeJuiceTotal),
    doctorPreventative: sumCoffeeRows(rows, (r) => r.doctorPreventative),
    doctorRestorative: sumCoffeeRows(rows, (r) => r.doctorRestorative),
    doctorCraProduction: sumCoffeeRows(rows, (r) => r.doctorCraProduction),
  };
}

function sumAdditionalRowFields(rows: AggregateAdditionalRow[]) {
  return {
    customer: sumAdditionalRows(rows, (r) => r.customer),
    icecream: sumAdditionalRows(rows, (r) => r.icecream),
    cake: sumAdditionalRows(rows, (r) => r.cake),
    donut: sumAdditionalRows(rows, (r) => r.donut),
    tart: sumAdditionalRows(rows, (r) => r.tart),
    peach: sumAdditionalRows(rows, (r) => r.peach),
    peppermint: sumAdditionalRows(rows, (r) => r.peppermint),
    total: sumAdditionalRows(rows, (r) => r.total),
  };
}

function sumSugarRowFields(rows: AggregateSugarRow[]) {
  return {
    sugar: sumSugarRows(rows, (r) => r.sugar),
    sugarGood: sumSugarRows(rows, (r) => r.sugarGood),
    sugarBad: sumSugarRows(rows, (r) => r.sugarBad),
    paper: sumSugarRows(rows, (r) => r.paper),
  };
}

function formatActionError(prefix: string, error: unknown, fallback = 'Error'): string {
  const message = (error as { message?: string })?.message || fallback;
  return prefix ? `${prefix}: ${message}` : message;
}

function formatDollarAmount(value: number): string {
  const amount = formatAmount(value);
  return amount.startsWith('-') ? `-$${amount.slice(1)}` : `$${amount}`;
}

function formatDollarDailyAverage(total: number, days: number): string {
  if (days <= 0) return '$0';
  const avg = Math.round(((total / days) + Number.EPSILON) * 100) / 100;
  return formatDollarAmount(avg);
}

function formatDailyAverage(total: number, days: number): string {
  if (days <= 0) return '0';
  return Math.round((total / days) + Number.EPSILON).toLocaleString('en-US');
}

function formatPercentage(value: number, total: number): string {
  if (total <= 0) return '0%';
  return `${Math.round((value / total) * 100)}%`;
}

function personKey(position: unknown, name: unknown): string {
  return `${String(position ?? '').trim() || '-'}__${String(name ?? '').trim() || '-'}`;
}

function personKeyParts(pKey: string): { position: string; name: string } {
  const [position, name] = pKey.split('__');
  return { position, name };
}

function emptyCoffeeAggregateRow(pKey: string): AggregateCoffeeRow {
  const { position, name } = personKeyParts(pKey);
  return {
    position,
    name,
    doctorPreventative: 0,
    doctorRestorative: 0,
    doctorCraProduction: 0,
    sales: 0,
    coffeeNew: 0,
    coffeeReturn: 0,
    coffeeTotal: 0,
    coffeeNo: 0,
    renderedCoffee: 0,
    coffeeYes: 0,
    orangeJuiceNew: 0,
    orangeJuiceReturn: 0,
    orangeJuiceTotal: 0,
  };
}

function emptyAdditionalAggregateRow(pKey: string): AggregateAdditionalRow {
  const { position, name } = personKeyParts(pKey);
  return {
    position,
    name,
    customer: 0,
    icecream: 0,
    cake: 0,
    donut: 0,
    tart: 0,
    peach: 0,
    peppermint: 0,
    total: 0,
  };
}

function emptySugarAggregateRow(pKey: string): AggregateSugarRow {
  const { position, name } = personKeyParts(pKey);
  return {
    position,
    name,
    sugar: 0,
    sugarGood: 0,
    sugarBad: 0,
    paper: 0,
  };
}

const POSITION_SORT_ORDER = ['doctor', 'rda', 'da'] as const;

function positionSortRank(position: unknown): number {
  const normalized = String(position ?? '').trim().toLowerCase();
  const idx = POSITION_SORT_ORDER.indexOf(normalized as (typeof POSITION_SORT_ORDER)[number]);
  return idx >= 0 ? idx : POSITION_SORT_ORDER.length;
}

function compareByPositionThenName(
  a: { position: string; name: string },
  b: { position: string; name: string }
): number {
  const rankDiff = positionSortRank(a.position) - positionSortRank(b.position);
  if (rankDiff !== 0) return rankDiff;
  const posDiff = a.position.localeCompare(b.position);
  if (posDiff !== 0) return posDiff;
  return a.name.localeCompare(b.name);
}

function countRowsByPosition<T extends { position: string }>(rows: T[], position: string): number {
  const target = position.trim().toLowerCase();
  return rows.filter((r) => String(r.position ?? '').trim().toLowerCase() === target).length;
}

function safeStr(value: unknown): string {
  return String(value ?? '').trim();
}

function sanitizeFilename(name: string): string {
  return name.replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, '_');
}

function downloadPdfBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function normalizeDocMonth(dateValue: unknown): string {
  const raw = String(dateValue ?? '').trim();
  const match = raw.match(/^(\d{4})-(\d{1,2})/);
  if (!match) return '';
  return `${match[1]}-${String(Number(match[2])).padStart(2, '0')}`;
}

function whiteBearMonthDocIdVariants(month: string): string[] {
  const normalized = normalizeDocMonth(month) || String(month ?? '').trim();
  const variants: string[] = [];
  const add = (value: string) => {
    const trimmed = String(value ?? '').trim();
    if (trimmed && !variants.includes(trimmed)) variants.push(trimmed);
  };

  add(normalized);
  const match = normalized.match(/^(\d{4})-(\d{1,2})$/);
  if (match) {
    add(`${match[1]}-${Number(match[2])}`);
    add(`${match[1]}-${String(Number(match[2])).padStart(2, '0')}`);
  }
  return variants;
}

function getWhiteBearMonthDocId(month: string): string {
  return normalizeDocMonth(month) || String(month ?? '').trim();
}

function getMonthlyProductionDocId(month: string, selectedOffice: string): string {
  return `${getWhiteBearMonthDocId(month)}_${getMonthlyLocationFieldKey(selectedOffice)}`;
}

function monthlyProductionDocIdVariants(month: string, selectedOffice: string): string[] {
  const locationKeys = [
    getMonthlyLocationFieldKey(selectedOffice),
    String(selectedOffice ?? '').trim(),
  ];
  const seen = new Set<string>();
  const out: string[] = [];

  for (const monthId of whiteBearMonthDocIdVariants(month)) {
    for (const locationKey of locationKeys) {
      if (!locationKey) continue;
      const id = `${monthId}_${locationKey}`;
      if (!seen.has(id)) {
        seen.add(id);
        out.push(id);
      }
    }
  }
  return out;
}

function normalizeOfficeName(value: unknown): string {
  return String(value ?? '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

function parseWhiteBearLocations(raw: unknown): Record<string, unknown> {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return {};
  return { ...(raw as Record<string, unknown>) };
}

function officeNamesMatch(a: unknown, b: unknown): boolean {
  const left = normalizeOfficeName(a);
  const right = normalizeOfficeName(b);
  if (!left || !right) return false;
  if (left === right) return true;
  return left.replace(/\s+/g, '') === right.replace(/\s+/g, '');
}

function toFirestoreKey(value: string): string {
  return value.replace(/[.#$/\[\]]/g, '_');
}

function getMonthlyLocationFieldKey(location: string): string {
  const raw = String(location ?? '').trim();
  return /[.#$/\[\]]/.test(raw) ? toFirestoreKey(raw) : raw;
}

type DaysInOfficeOverrides = Record<string, number>;

function emptyDaysInOfficeOverrides(): DaysInOfficeOverrides {
  return {};
}

function normalizeDaysInOfficeOverrides(raw: unknown): DaysInOfficeOverrides {
  if (!raw || typeof raw !== 'object') return {};
  const data = raw as Record<string, unknown>;
  const out: DaysInOfficeOverrides = {};

  const isLegacySectioned = 'coffee' in data || 'additional' in data || 'sugar' in data;
  if (isLegacySectioned) {
    (['coffee', 'additional', 'sugar'] as const).forEach((section) => {
      const sectionRaw = data[section];
      if (!sectionRaw || typeof sectionRaw !== 'object') return;
      Object.entries(sectionRaw as Record<string, unknown>).forEach(([key, value]) => {
        const n = parseNumber(value);
        if (Number.isFinite(n) && n >= 0) out[key] = n;
      });
    });
    return out;
  }

  Object.entries(data).forEach(([key, value]) => {
    const n = parseNumber(value);
    if (Number.isFinite(n) && n >= 0) out[key] = n;
  });
  return out;
}

function DaysInOfficeCell({
  personKeyValue,
  value,
  isEditing,
  onChange,
  cellStyle = tableCellStyle,
}: {
  personKeyValue: string;
  value: number | undefined;
  isEditing: boolean;
  onChange: (personKeyValue: string, rawValue: string) => void;
  cellStyle?: React.CSSProperties;
}) {
  if (!isEditing) {
    return <td style={cellStyle}>{value === undefined ? '' : value}</td>;
  }
  return (
    <td style={cellStyle}>
      <input
        type="number"
        min={0}
        step={1}
        value={value === undefined ? '' : value}
        onChange={(e) => onChange(personKeyValue, e.target.value)}
        style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
      />
    </td>
  );
}

function renderDaysInOfficeTotal(
  overrides: DaysInOfficeOverrides,
  rows: Array<{ position: string; name: string }>
): string {
  const hasAny = rows.some((row) => overrides[personKey(row.position, row.name)] !== undefined);
  if (!hasAny) return '';
  const total = rows.reduce((sum, row) => {
    const val = overrides[personKey(row.position, row.name)];
    return val === undefined ? sum : sum + val;
  }, 0);
  return String(total.toFixed(2));
}

function getDaysInOfficeValue(
  overrides: DaysInOfficeOverrides,
  position: string,
  name: string
): number | undefined {
  return overrides[personKey(position, name)];
}

function sumDaysInOfficeFromOverrides(
  overrides: DaysInOfficeOverrides,
  rows: Array<{ position: string; name: string }>
): number {
  return rows.reduce((sum, row) => {
    const val = getDaysInOfficeValue(overrides, row.position, row.name);
    return val === undefined ? sum : sum + val;
  }, 0);
}

function formatDailyAverageWithDays(total: number, days: number | undefined): string {
  if (days === undefined || days <= 0) return '';
  return formatDailyAverage(total, days);
}

function formatDollarDailyAverageWithDays(total: number, days: number | undefined): string {
  if (days === undefined || days <= 0) return '';
  return formatDollarDailyAverage(total, days);
}

function totalDaysForAverage(totalDays: number): number | undefined {
  return totalDays > 0 ? totalDays : undefined;
}

type WhiteBearLocationSettings = {
  locationName?: string;
  daysInOfficeOverrides?: unknown;
};

function collectWhiteBearLocationEntries(
  locations: Record<string, unknown>,
  selectedOffice: string
): WhiteBearLocationSettings[] {
  const locationKey = getMonthlyLocationFieldKey(selectedOffice);
  const normalizedOffice = selectedOffice.replace(/\s+/g, '').toLowerCase();
  const seen = new Set<string>();
  const out: WhiteBearLocationSettings[] = [];

  const add = (key: string) => {
    const trimmed = String(key ?? '').trim();
    if (!trimmed || seen.has(trimmed)) return;
    seen.add(trimmed);
    const value = locations[trimmed];
    if (value && typeof value === 'object') out.push(value as WhiteBearLocationSettings);
  };

  add(selectedOffice);
  add(locationKey);
  Object.entries(locations).forEach(([key, value]) => {
    const keyRaw = String(key ?? '').trim();
    const nameNormalized = String((value as WhiteBearLocationSettings)?.locationName ?? '')
      .trim()
      .replace(/\s+/g, '')
      .toLowerCase();
    if (
      keyRaw === selectedOffice ||
      getMonthlyLocationFieldKey(keyRaw) === locationKey ||
      keyRaw.replace(/\s+/g, '').toLowerCase() === normalizedOffice ||
      nameNormalized === normalizedOffice ||
      officeNamesMatch(keyRaw, selectedOffice) ||
      officeNamesMatch((value as WhiteBearLocationSettings)?.locationName, selectedOffice)
    ) {
      add(keyRaw);
    }
  });

  return out;
}

function collectWhiteBearLocationEntriesWithFallback(
  locations: Record<string, unknown>,
  selectedOffice: string
): WhiteBearLocationSettings[] {
  const matched = collectWhiteBearLocationEntries(locations, selectedOffice);
  if (matched.length > 0) return matched;

  const candidates = Object.values(locations)
    .filter((value) => value && typeof value === 'object')
    .map((value) => value as WhiteBearLocationSettings)
    .filter((value) => value.daysInOfficeOverrides !== undefined);

  return candidates.length === 1 ? candidates : [];
}

function mergeWhiteBearLocationSettings(entries: WhiteBearLocationSettings[]) {
  let daysInOffice: DaysInOfficeOverrides = {};

  entries.forEach((entry) => {
    daysInOffice = {
      ...daysInOffice,
      ...normalizeDaysInOfficeOverrides(entry.daysInOfficeOverrides),
    };
  });

  return { daysInOffice };
}

function daysInOfficeFromProductionDoc(data: Record<string, unknown> | undefined): DaysInOfficeOverrides | null {
  if (!data || data.daysInOfficeOverrides === undefined) return null;
  return normalizeDaysInOfficeOverrides(data.daysInOfficeOverrides);
}

async function loadWhiteBearSettingsForOffice(month: string, selectedOffice: string) {
  let daysInOffice: DaysInOfficeOverrides = {};

  for (const docId of monthlyProductionDocIdVariants(month, selectedOffice)) {
    const snap = await getDoc(doc(db, 'monthly production', docId));
    if (!snap.exists()) continue;
    const data = snap.data() as Record<string, unknown>;
    const fromDoc = daysInOfficeFromProductionDoc(data);
    if (fromDoc) {
      daysInOffice = { ...daysInOffice, ...fromDoc };
    }
  }

  const entries: WhiteBearLocationSettings[] = [];
  for (const monthId of whiteBearMonthDocIdVariants(month)) {
    const snap = await getDoc(doc(db, 'monthly production', monthId));
    if (!snap.exists()) continue;
    const data = snap.data() as Record<string, unknown> | undefined;
    const locations = parseWhiteBearLocations(data?.locations);
    entries.push(...collectWhiteBearLocationEntriesWithFallback(locations, selectedOffice));

    const hasRootSettings = data && data.daysInOfficeOverrides !== undefined;
    if (hasRootSettings && (!data.locationName || officeNamesMatch(data.locationName, selectedOffice))) {
      entries.push({
        locationName: String(data.locationName ?? selectedOffice),
        daysInOfficeOverrides: data.daysInOfficeOverrides,
      });
    }
  }

  if (entries.length > 0) {
    daysInOffice = { ...daysInOffice, ...mergeWhiteBearLocationSettings(entries).daysInOffice };
  }

  if (Object.keys(daysInOffice).length === 0) return null;
  return { daysInOffice };
}

function normalizePositionLabel(position: unknown): string {
  return String(position ?? '').trim().toLowerCase();
}

function isPositionGroupStart<T extends { position: string }>(rows: readonly T[], index: number): boolean {
  if (index <= 0) return false;
  return normalizePositionLabel(rows[index]?.position) !== normalizePositionLabel(rows[index - 1]?.position);
}

function tableCellStyleWithPositionSeparator(isPositionGroupStart: boolean): React.CSSProperties {
  if (!isPositionGroupStart) return tableCellStyle;
  return { ...tableCellStyle, borderTop: '2px solid #9ca3af' };
}

function buildPdfPersonRowCells<T extends { position: string; name: string }>(
  rows: T[],
  daysInOfficeOverrides: DaysInOfficeOverrides,
  buildCells: (row: T, days: number | undefined) => string[]
): string[][] {
  return rows.map((row) => {
    const days = getDaysInOfficeValue(daysInOfficeOverrides, row.position, row.name);
    return buildCells(row, days);
  });
}

function pdfPersonRowPrefix(row: { position: string; name: string }, days: number | undefined): string[] {
  return [safeStr(row.position), safeStr(row.name), days === undefined ? '' : String(days)];
}

function SummaryTable({
  headers,
  totalRow,
  dailyAverageRow,
  minWidth,
}: {
  headers: string[];
  totalRow: React.ReactNode[];
  dailyAverageRow: React.ReactNode[];
  minWidth: number;
}) {
  return (
    <div style={{ overflowX: 'auto', marginBottom: 16 }}>
      <table style={{ width: '100%', minWidth, borderCollapse: 'collapse', fontSize: 14 }}>
        <thead>
          <tr>
            {headers.map((header, idx) => (
              <th key={`${header}_${idx}`} style={tableHeaderStyle}>
                {header}
              </th>
            ))}
          </tr>
        </thead>
        <tfoot>
          <tr style={tableFooterRowStyle}>
            {totalRow.map((cell, idx) => (
              <td key={`total_${idx}`} style={tableCellStyle}>
                {cell}
              </td>
            ))}
          </tr>
          <tr style={tableFooterRowStyle}>
            {dailyAverageRow.map((cell, idx) => (
              <td key={`avg_${idx}`} style={tableCellStyle}>
                {cell}
              </td>
            ))}
          </tr>
        </tfoot>
      </table>
    </div>
  );
}

function buildSugarPositionSubtotalPdfSummary(sugarRows: AggregateSugarRow[]): string[][] {
  const rows: string[][] = [];

  (['doctor', 'rda'] as const).forEach((position) => {
    const positionRows = sugarRows.filter(
      (row) => normalizePositionLabel(row.position) === position
    );
    if (positionRows.length === 0) return;
    rows.push([
      `${position} Total`,
      '-',
      '-',
      formatAmount(sumSugarRows(positionRows, (row) => row.sugar)),
      formatAmount(sumSugarRows(positionRows, (row) => row.sugarGood)),
      '',
      '',
      '',
      '',
      '',
      '',
    ]);
  });

  return rows;
}

function createMonthlyReportPdfDocument({
  aggregate,
  generatedDate,
  daysInOfficeOverrides,
}: {
  aggregate: Aggregate;
  generatedDate: string;
  daysInOfficeOverrides: DaysInOfficeOverrides;
}) {
  const coffeeTotal = sumCoffeeRowFields(aggregate.coffeeRows);
  const additionalTotal = sumAdditionalRowFields(aggregate.additionalRows);
  const sugarTotal = sumSugarRowFields(aggregate.sugarRows);
  const dailyDivisor = aggregate.docCount;
  const coffeeSalesTotal = coffeeTotal.sales;
  const coffeeDaysInOfficeTotal = sumDaysInOfficeFromOverrides(daysInOfficeOverrides, aggregate.coffeeRows);
  const coffeeTotalAmount = coffeeTotal.coffeeTotal;
  const coffeeOjTotalAmount = coffeeTotal.orangeJuiceTotal;
  const sugarTotalAmount = sugarTotal.sugar;
  const sugarBillableTotalAmount = sugarTotal.sugarGood;
  const paperTotalAmount = sugarTotal.paper;
  const sugarDaysInOfficeTotal = sumDaysInOfficeFromOverrides(daysInOfficeOverrides, aggregate.sugarRows);

  const coffeeSalesRows = buildPdfPersonRowCells(aggregate.coffeeRows, daysInOfficeOverrides, (r, days) => [
    ...pdfPersonRowPrefix(r, days),
    formatDollarAmount(r.doctorPreventative),
    formatDollarAmount(r.doctorRestorative),
    formatDollarAmount(r.doctorCraProduction),
    formatDollarAmount(r.sales),
    formatDollarDailyAverageWithDays(r.sales, days),
    formatPercentage(r.sales, coffeeSalesTotal),
  ]);
  const coffeeStatusRows = buildPdfPersonRowCells(aggregate.coffeeRows, daysInOfficeOverrides, (r, days) => [
    ...pdfPersonRowPrefix(r, days),
    formatAmount(r.coffeeNew),
    formatAmount(r.coffeeReturn),
    formatAmount(r.coffeeTotal),
    formatDailyAverageWithDays(r.coffeeTotal, days),
    formatPercentage(r.coffeeTotal, coffeeTotalAmount),
    formatAmount(r.coffeeNo),
    formatAmount(r.renderedCoffee),
    formatAmount(r.coffeeYes),
  ]);
  const coffeeOjRows = buildPdfPersonRowCells(aggregate.coffeeRows, daysInOfficeOverrides, (r, days) => [
    ...pdfPersonRowPrefix(r, days),
    formatAmount(r.orangeJuiceNew),
    formatAmount(r.orangeJuiceReturn),
    formatAmount(r.orangeJuiceTotal),
    formatDailyAverageWithDays(r.orangeJuiceTotal, days),
    formatPercentage(r.orangeJuiceTotal, coffeeOjTotalAmount),
    '-',
    '-',
  ]);
  const additionalRows = buildPdfPersonRowCells(aggregate.additionalRows, daysInOfficeOverrides, (r, days) => [
    ...pdfPersonRowPrefix(r, days),
    formatAmount(r.customer),
    formatAmount(r.icecream),
    formatAmount(r.cake),
    formatAmount(r.donut),
    formatAmount(r.tart),
    formatAmount(r.peach),
    formatAmount(r.peppermint),
    formatAmount(r.total),
  ]);
  const sugarRows = buildPdfPersonRowCells(aggregate.sugarRows, daysInOfficeOverrides, (r, days) => [
    ...pdfPersonRowPrefix(r, days),
    formatAmount(r.sugar),
    formatAmount(r.sugarGood),
    formatDailyAverageWithDays(r.sugarGood, days),
    formatPercentage(r.sugarGood, sugarBillableTotalAmount),
    formatAmount(r.sugarBad),
    formatAmount(r.paper),
    formatDailyAverageWithDays(r.paper, days),
    formatPercentage(r.paper, paperTotalAmount),
  ]);
  const sugarRowMeta: PdfTableDataRowMeta[] = aggregate.sugarRows.map((_, index) => ({
    positionGroupStart: isPositionGroupStart(aggregate.sugarRows, index),
  }));
  const sugarPositionSubtotalRows = buildSugarPositionSubtotalPdfSummary(aggregate.sugarRows);
  const diagnoseRows: string[][] = aggregate.diagnoseRows.map((r: AggregateDiagnoseRow) => [
    safeStr(r.sName),
    formatAmount(r.diagnose),
  ]);
  const reasonRows: string[][] = aggregate.reasonRows.map((r: AggregateReasonRow) => [
    safeStr(r.reason),
    formatAmount(r.orangeJuice),
    formatAmount(r.paper),
    formatAmount(r.coffee),
  ]);

  return (
    <Document>
      <Page size="A4" orientation="landscape" style={monthlyPdfStyles.page}>
        <Text style={monthlyPdfStyles.title}>{aggregate.month} {aggregate.location} Monthly Report</Text>
        <Text style={monthlyPdfStyles.meta}>Days: {aggregate.docCount} | Generated: {generatedDate}</Text>

        {renderPdfTable(
          ['', 'Submitted Production', 'CRA Production', 'Production W/Out CRA', 'Prophy @ OE', 'Prophy @ TX', 'Just Prophy', 'Actual Prophy'],
          [[
            'Total',
            formatDollarAmount(aggregate.grandTotal),
            formatDollarAmount(aggregate.coffeeSales),
            formatDollarAmount(aggregate.salesWithoutCoffee),
            formatAmount(aggregate.paperAtOrangeJuice),
            formatAmount(aggregate.paperAtTea),
            formatAmount(aggregate.justPaper),
            formatAmount(aggregate.prophyTotal),
          ]],
          [9, 13, 13, 13, 13, 13, 13, 13],
          [[
            'Daily Average',
            formatDollarDailyAverage(aggregate.grandTotal, dailyDivisor),
            formatDollarDailyAverage(aggregate.coffeeSales, dailyDivisor),
            formatDollarDailyAverage(aggregate.salesWithoutCoffee, dailyDivisor),
            formatDailyAverage(aggregate.paperAtOrangeJuice, dailyDivisor),
            formatDailyAverage(aggregate.paperAtTea, dailyDivisor),
            formatDailyAverage(aggregate.justPaper, dailyDivisor),
            formatDailyAverage(aggregate.prophyTotal, dailyDivisor),
          ]],
          'Production'
        )}

        {renderPdfTable(
          ['', 'Preventative', 'Restorative', 'CRA Production', '1st Review Production', 'Mailed Production'],
          [[
            'Total',
            formatDollarAmount(aggregate.pineapple),
            formatDollarAmount(aggregate.rose),
            formatDollarAmount(aggregate.coffeeSales),
            formatDollarAmount(aggregate.pineapple + aggregate.rose + aggregate.coffeeSales),
            formatDollarAmount(aggregate.mailedProduction),
          ]],
          [14, 17, 17, 17, 17, 18],
          [[
            'Daily Average',
            formatDollarDailyAverage(aggregate.pineapple, dailyDivisor),
            formatDollarDailyAverage(aggregate.rose, dailyDivisor),
            formatDollarDailyAverage(aggregate.coffeeSales, dailyDivisor),
            formatDollarDailyAverage(aggregate.pineapple + aggregate.rose + aggregate.coffeeSales, dailyDivisor),
            formatDollarDailyAverage(aggregate.mailedProduction, dailyDivisor),
          ]]
        )}

        {renderPdfTable(
          ['Add On', 'No Shows', 'Scheduled', 'Seen', 'Seen %', 'Referral', 'Postcard'],
          [[
            formatAmount(aggregate.visitsAdd),
            formatAmount(aggregate.visitsNoShow),
            formatAmount(aggregate.visitsScheduled),
            formatAmount(aggregate.visitsSeen),
            formatPercentage(aggregate.visitsSeen, aggregate.visitsScheduled),
            formatAmount(aggregate.visitsReferral),
            formatAmount(aggregate.visitsPostcard),
          ]],
          [14, 14, 14, 12, 12, 16, 16],
          undefined,
          'Visits'
        )}

        {renderPdfTable(
          ['Position', 'Name', 'Days in Office', 'Preventative', 'Restorative', 'CRA Production', 'Production', 'Average', '%'],
          coffeeSalesRows,
          [10, 10, 9, 11, 11, 11, 11, 11, 15],
          [[
            'Total',
            String(countRowsByPosition(aggregate.coffeeRows, 'doctor')),
            String(coffeeDaysInOfficeTotal || ''),
            formatDollarAmount(sumCoffeeRows(aggregate.coffeeRows, (r) => r.doctorPreventative)),
            formatDollarAmount(sumCoffeeRows(aggregate.coffeeRows, (r) => r.doctorRestorative)),
            formatDollarAmount(sumCoffeeRows(aggregate.coffeeRows, (r) => r.doctorCraProduction)),
            formatDollarAmount(coffeeSalesTotal),
            formatDollarDailyAverageWithDays(coffeeSalesTotal, totalDaysForAverage(coffeeDaysInOfficeTotal)),
            '100%',
          ], [
            'Daily Average',
            '-',
            '-',
            formatDollarDailyAverage(sumCoffeeRows(aggregate.coffeeRows, (r) => r.doctorPreventative), dailyDivisor),
            formatDollarDailyAverage(sumCoffeeRows(aggregate.coffeeRows, (r) => r.doctorRestorative), dailyDivisor),
            formatDollarDailyAverage(sumCoffeeRows(aggregate.coffeeRows, (r) => r.doctorCraProduction), dailyDivisor),
            formatDollarDailyAverage(coffeeSalesTotal, dailyDivisor),
            '-',
            '-',
          ]],
          'Doctors'
        )}
        {renderPdfTable(
          ['Position', 'Name', 'Days in Office', 'CRA (New)', 'CRA (Return)', 'CRA Total', 'Average', '%', 'CRA (Not Billable)', 'Rendered CRA', 'CRA (Billable)'],
          coffeeStatusRows,
          [10, 10, 10, 9, 9, 9, 9, 9, 8, 9, 9],
          [[
            'Total',
            String(countRowsByPosition(aggregate.coffeeRows, 'doctor')),
            String(coffeeDaysInOfficeTotal || ''),
            formatAmount(coffeeTotal.coffeeNew),
            formatAmount(coffeeTotal.coffeeReturn),
            formatAmount(coffeeTotal.coffeeTotal),
            formatDailyAverageWithDays(coffeeTotal.coffeeTotal, totalDaysForAverage(coffeeDaysInOfficeTotal)),
            '100%',
            formatAmount(coffeeTotal.coffeeNo),
            formatAmount(coffeeTotal.renderedCoffee),
            formatAmount(coffeeTotal.coffeeYes),
          ], [
            'Daily Average',
            '-',
            '-',
            formatDailyAverage(coffeeTotal.coffeeNew, dailyDivisor),
            formatDailyAverage(coffeeTotal.coffeeReturn, dailyDivisor),
            formatDailyAverage(coffeeTotal.coffeeTotal, dailyDivisor),
            '-',
            '-',
            formatDailyAverage(coffeeTotal.coffeeNo, dailyDivisor),
            formatDailyAverage(coffeeTotal.renderedCoffee, dailyDivisor),
            formatDailyAverage(coffeeTotal.coffeeYes, dailyDivisor),
          ]]
        )}
        {renderPdfTable(
          ['Position', 'Name', 'Days in Office', 'OE (NP)', 'OE (RC)', 'OE Total', 'Average', '%', 'Actual OE (NP)', 'Actual OE (RC)'],
          coffeeOjRows,
          [10, 10, 10, 10, 10, 10, 10, 8, 11, 11],
          [[
            'Total',
            String(countRowsByPosition(aggregate.coffeeRows, 'doctor')),
            String(coffeeDaysInOfficeTotal || ''),
            formatAmount(coffeeTotal.orangeJuiceNew),
            formatAmount(coffeeTotal.orangeJuiceReturn),
            formatAmount(coffeeTotal.orangeJuiceTotal),
            formatDailyAverageWithDays(coffeeTotal.orangeJuiceTotal, totalDaysForAverage(coffeeDaysInOfficeTotal)),
            '100%',
            formatAmount(aggregate.actualOrangeJuiceNew),
            formatAmount(aggregate.actualOrangeJuiceReturn),
          ], [
            'Daily Average',
            '-',
            '-',
            formatDailyAverage(coffeeTotal.orangeJuiceNew, dailyDivisor),
            formatDailyAverage(coffeeTotal.orangeJuiceReturn, dailyDivisor),
            formatDailyAverage(coffeeTotal.orangeJuiceTotal, dailyDivisor),
            '-',
            '-',
            formatDailyAverage(aggregate.actualOrangeJuiceNew, dailyDivisor),
            formatDailyAverage(aggregate.actualOrangeJuiceReturn, dailyDivisor),
          ]]
        )}

        {renderPdfTable(
          ['Position', 'Name', 'Days', 'Patient Seen', 'Insurance', 'Cash', 'Dentical', 'Treatment', 'Primary Teeth', 'Permanent Teeth', 'Total'],
          additionalRows,
          [10, 10, 6, 8, 8, 8, 8, 8, 8, 10, 16],
          [[
            'Total',
            String(countRowsByPosition(aggregate.additionalRows, 'doctor')),
            '-',
            formatAmount(additionalTotal.customer),
            formatAmount(additionalTotal.icecream),
            formatAmount(additionalTotal.cake),
            formatAmount(additionalTotal.donut),
            formatAmount(additionalTotal.tart),
            formatAmount(additionalTotal.peach),
            formatAmount(additionalTotal.peppermint),
            formatAmount(additionalTotal.total),
          ], [
            'Daily Average',
            '-',
            '-',
            formatDailyAverage(additionalTotal.customer, dailyDivisor),
            formatDailyAverage(additionalTotal.icecream, dailyDivisor),
            formatDailyAverage(additionalTotal.cake, dailyDivisor),
            formatDailyAverage(additionalTotal.donut, dailyDivisor),
            formatDailyAverage(additionalTotal.tart, dailyDivisor),
            formatDailyAverage(additionalTotal.peach, dailyDivisor),
            formatDailyAverage(additionalTotal.peppermint, dailyDivisor),
            formatDailyAverage(additionalTotal.total, dailyDivisor),
          ]]
        )}

        {renderPdfTable(
          ['Position', 'Name', 'Days', 'Sealant', 'Sealant (Billable)', 'Sealant Average', '%', 'Sealant (Redo)', 'Prophy', 'Prophy Average', '%'],
          sugarRows,
          [11, 11, 8, 9, 9, 9, 8, 9, 9, 9, 8],
          [
            ...sugarPositionSubtotalRows,
            [
              'Total',
              String(aggregate.sugarRows.length),
              String(sugarDaysInOfficeTotal || ''),
              formatAmount(sugarTotal.sugar),
              formatAmount(sugarTotal.sugarGood),
              '-',
              '100%',
              formatAmount(sugarTotal.sugarBad),
              formatAmount(sugarTotal.paper),
              '-',
              '100%',
            ],
            [
              'Daily Average',
              '-',
              '-',
              formatDailyAverage(sugarTotal.sugar, dailyDivisor),
              formatDailyAverage(sugarTotal.sugarGood, dailyDivisor),
              '-',
              '-',
              formatDailyAverage(sugarTotal.sugarBad, dailyDivisor),
              formatDailyAverage(sugarTotal.paper, dailyDivisor),
              '-',
              '-',
            ],
          ],
          'Sealant & Prophy',
          undefined,
          sugarRowMeta
        )}

        {renderPdfTable(
          ['Name', 'Sealant Diagnose'],
          diagnoseRows,
          [40, 30],
          undefined,
          'Sealant Diagnose'
        )}
        
        {renderPdfTable(
          ['Reasoning', 'OE', 'Pro', 'CRA'],
          reasonRows,
          [40, 20, 20, 20],
          undefined,
          'Short Procedures'
        )}
      </Page>
    </Document>
  );
}

function MonthlyReportPageContent() {
  const searchParams = useSearchParams();
  const selectedMonth = searchParams.get('month') ?? '';
  const selectedOffice = searchParams.get('office') ?? '';
  const [pageReady, setPageReady] = useState(false);
  const [docs, setDocs] = useState<FormDoc[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [isPdfDownloading, setIsPdfDownloading] = useState(false);
  const [saveMessage, setSaveMessage] = useState('');
  const [isEditing, setIsEditing] = useState(false);
  const [daysOverridesSaved, setDaysOverridesSaved] = useState<DaysInOfficeOverrides>(emptyDaysInOfficeOverrides);
  const [daysOverridesDraft, setDaysOverridesDraft] = useState<DaysInOfficeOverrides>(emptyDaysInOfficeOverrides);
  const hasSelection = selectedMonth !== '' && selectedOffice !== '';
  const activeDaysOverrides = isEditing ? daysOverridesDraft : daysOverridesSaved;

  useEffect(() => {
    let cancelled = false;
    const goHome = () => {
      if (typeof window !== 'undefined') {
        window.location.replace('/');
      }
    };

    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      try {
        if (!currentUser) {
          goHome();
          return;
        }

        const userDoc = await getDoc(doc(db, 'users', currentUser.uid));
        if (!userDoc.exists()) {
          goHome();
          return;
        }

        const userData = userDoc.data();
        if (
          userData?.role !== 'HR' &&
          userData?.role !== 'Director'
        ) {
          goHome();
          return;
        }

        if (!cancelled) {
          setPageReady(true);
        }
      } catch {
        goHome();
      }
    });

    if (
      process.env.NODE_ENV === 'production' &&
      typeof window !== 'undefined' &&
      window.location.protocol !== 'https:'
    ) {
      window.location.href = window.location.href.replace('http:', 'https:');
    }

    return () => {
      cancelled = true;
      unsubscribe();
    };
  }, []);

  useEffect(() => {
    if (!pageReady) return;

    const load = async () => {
      try {
        const snap = await getDocs(collection(db, 'simple-forms'));
        const loaded = snap.docs.map((d) => ({ id: d.id, ...(d.data() as Omit<FormDoc, 'id'>) }));
        setDocs(loaded);
      } catch (e: unknown) {
        setError(formatActionError('', e, 'Error, Please try again.'));
      } finally {
        setLoading(false);
      }
    };
    load();
  }, [pageReady]);

  const monthlyDocs = useMemo(
    () => {
      if (!hasSelection) return [];
      return docs.filter((doc) =>
        !!doc.submittedDateTime &&
        normalizeDocMonth(doc.date) === selectedMonth &&
        String(doc.location ?? '').trim() === selectedOffice
      );
    },
    [docs, hasSelection, selectedMonth, selectedOffice]
  );

  const aggregates = useMemo(() => {
    if (!hasSelection) return [];

    const map = new Map<string, Aggregate>();

    monthlyDocs.forEach((doc) => {
      const month = normalizeDocMonth(doc.date) || selectedMonth;
      const location = String(doc.location ?? '').trim() || 'Unknown';
      const key = `${month}_${location}`;

      let agg = map.get(key);
      if (!agg) {
        agg = {
          key,
          month,
          location,
          docCount: 0,
          grandTotal: 0,
          coffeeSales: 0,
          salesWithoutCoffee: 0,
          paperAtOrangeJuice: 0,
          paperAtTea: 0,
          justPaper: 0,
          prophyTotal: 0,
          actualOrangeJuiceNew: 0,
          actualOrangeJuiceReturn: 0,
          pineapple: 0,
          rose: 0,
          mailedProduction: 0,
          visitsAdd: 0,
          visitsNoShow: 0,
          visitsScheduled: 0,
          visitsSeen: 0,
          visitsReferral: 0,
          visitsPostcard: 0,
          coffeeRows: [],
          additionalRows: [],
          sugarRows: [],
          diagnoseRows: [],
          reasonRows: [],
        };
        map.set(key, agg);
      }

      agg.docCount += 1;
      agg.grandTotal = addAmount(agg.grandTotal, doc.grandTotal);
      agg.coffeeSales = addAmount(agg.coffeeSales, doc.coffeeSales);
      agg.salesWithoutCoffee = addAmount(agg.salesWithoutCoffee, doc.salesWithoutCoffee);
      agg.paperAtOrangeJuice = addAmount(agg.paperAtOrangeJuice, doc.paperAtOrangeJuice);
      agg.paperAtTea = addAmount(agg.paperAtTea, doc.paperAtTea);
      agg.justPaper = addAmount(agg.justPaper, doc.justPaper);
      agg.prophyTotal = addAmount(agg.prophyTotal, doc.prophyTotal);
      agg.actualOrangeJuiceNew = addAmount(agg.actualOrangeJuiceNew, doc.coffeeActualTotals?.orangeJuiceNew);
      agg.actualOrangeJuiceReturn = addAmount(agg.actualOrangeJuiceReturn, doc.coffeeActualTotals?.orangeJuiceReturn);
      agg.pineapple = addAmount(agg.pineapple, doc.locationSummary?.pineapple);
      agg.rose = addAmount(agg.rose, doc.locationSummary?.rose);
      agg.mailedProduction = addAmount(agg.mailedProduction, doc.locationSummary?.mailedProduction);

      const sideMetrics = doc.productionSideMetrics;
      if (sideMetrics) {
        agg.visitsAdd = addAmount(agg.visitsAdd, sideMetrics.add);
        agg.visitsNoShow = addAmount(agg.visitsNoShow, sideMetrics.noShow);
        agg.visitsScheduled = addAmount(agg.visitsScheduled, sideMetrics.scheduled);
        agg.visitsSeen = addAmount(agg.visitsSeen, sideMetrics.seen);
        agg.visitsReferral = addAmount(agg.visitsReferral, sideMetrics.referral);
        agg.visitsPostcard = addAmount(agg.visitsPostcard, sideMetrics.postcard);
      }

      const coffeeMap = new Map(agg.coffeeRows.map((r) => [personKey(r.position, r.name), r]));
      (doc.tableRows || []).forEach((row) => {
        const pKey = personKey(row.position, row.name);
        const prev = coffeeMap.get(pKey) || emptyCoffeeAggregateRow(pKey);
        prev.sales = addAmount(prev.sales, row.sales);
        prev.coffeeNew = addAmount(prev.coffeeNew, row.coffeeNew);
        prev.coffeeReturn = addAmount(prev.coffeeReturn, row.coffeeReturn);
        prev.coffeeTotal = addAmount(prev.coffeeTotal, row.coffeeTotal);
        prev.coffeeNo = addAmount(prev.coffeeNo, row.coffeeNo);
        prev.renderedCoffee = addAmount(prev.renderedCoffee, row.renderedCoffee);
        prev.coffeeYes = addAmount(prev.coffeeYes, row.coffeeYes);
        prev.orangeJuiceNew = addAmount(prev.orangeJuiceNew, row.orangeJuiceNew);
        prev.orangeJuiceReturn = addAmount(prev.orangeJuiceReturn, row.orangeJuiceReturn);
        prev.orangeJuiceTotal = addAmount(prev.orangeJuiceTotal, row.orangeJuiceTotal);
        coffeeMap.set(pKey, prev);
      });

      const extraMap = new Map(agg.additionalRows.map((r) => [personKey(r.position, r.name), r]));
      (doc.extraInputRows || []).forEach((row) => {
        const pKey = personKey(row.position, row.name);
        const coffeePrev = coffeeMap.get(pKey) || emptyCoffeeAggregateRow(pKey);
        coffeePrev.doctorPreventative = addAmount(coffeePrev.doctorPreventative, row.doctorPreventative);
        coffeePrev.doctorRestorative = addAmount(coffeePrev.doctorRestorative, row.doctorRestorative);
        coffeePrev.doctorCraProduction = addAmount(coffeePrev.doctorCraProduction, row.doctorCraProduction);
        coffeeMap.set(pKey, coffeePrev);

        const prev = extraMap.get(pKey) || emptyAdditionalAggregateRow(pKey);
        prev.customer = addAmount(prev.customer, row.customer);
        prev.icecream = addAmount(prev.icecream, row.icecream);
        prev.cake = addAmount(prev.cake, row.cake);
        prev.donut = addAmount(prev.donut, row.donut);
        prev.tart = addAmount(prev.tart, row.tart);
        prev.peach = addAmount(prev.peach, row.peach);
        prev.peppermint = addAmount(prev.peppermint, row.peppermint);
        prev.total = addAmount(prev.total, row.total);
        extraMap.set(pKey, prev);
      });
      agg.coffeeRows = Array.from(coffeeMap.values()).sort(compareByPositionThenName);
      agg.additionalRows = Array.from(extraMap.values()).sort(compareByPositionThenName);

      const sugarMap = new Map(agg.sugarRows.map((r) => [personKey(r.position, r.name), r]));
      (doc.sugarRows || []).forEach((row) => {
        const pKey = personKey(row.position, row.name);
        const prev = sugarMap.get(pKey) || emptySugarAggregateRow(pKey);
        prev.sugar = addAmount(prev.sugar, row.sugar);
        prev.sugarGood = addAmount(prev.sugarGood, row.sugarGood);
        prev.sugarBad = addAmount(prev.sugarBad, row.sugarBad);
        prev.paper = addAmount(prev.paper, row.paper);
        sugarMap.set(pKey, prev);
      });
      agg.sugarRows = Array.from(sugarMap.values()).sort(compareByPositionThenName);

      const diagnoseMap = new Map(agg.diagnoseRows.map((r) => [r.sName, r]));
      (doc.diagnoseRows || []).forEach((row) => {
        const sName = String(row.sName ?? '').trim() || '-';
        const prev = diagnoseMap.get(sName) || { sName, diagnose: 0 };
        prev.diagnose = addAmount(prev.diagnose, row.diagnose);
        diagnoseMap.set(sName, prev);
      });
      agg.diagnoseRows = Array.from(diagnoseMap.values()).sort((a, b) => a.sName.localeCompare(b.sName));
      
      const reasonMap = new Map(agg.reasonRows.map((r) => [r.reason, r]));
      (doc.reasonRows || []).forEach((row) => {
        const reason = String(row.reason ?? '').trim() || '-';
        const prev = reasonMap.get(reason) || { reason, orangeJuice: 0, paper: 0, coffee: 0 };
        prev.orangeJuice = addAmount(prev.orangeJuice, row.orangeJuice);
        prev.paper = addAmount(prev.paper, row.paper);
        prev.coffee = addAmount(prev.coffee, row.coffee);
        reasonMap.set(reason, prev);
      });
      agg.reasonRows = Array.from(reasonMap.values()).sort((a, b) => a.reason.localeCompare(b.reason));
    });

    return Array.from(map.values()).sort((a, b) => b.key.localeCompare(a.key));
  }, [hasSelection, monthlyDocs, selectedMonth]);

  const selectedAggregate = aggregates[0] ?? null;

  useEffect(() => {
    let cancelled = false;

    const applyEmpty = () => {
      if (cancelled) return;
      setDaysOverridesSaved(emptyDaysInOfficeOverrides());
      setDaysOverridesDraft(emptyDaysInOfficeOverrides());
      setIsEditing(false);
    };

    const loadDaysInOfficeOverrides = async () => {
      if (!pageReady || !hasSelection) {
        applyEmpty();
        return;
      }

      try {
        const merged = await loadWhiteBearSettingsForOffice(selectedMonth, selectedOffice);
        if (cancelled) return;
        if (!merged) {
          applyEmpty();
          return;
        }

        setDaysOverridesSaved(merged.daysInOffice);
        setDaysOverridesDraft(merged.daysInOffice);
        setIsEditing(false);
      } catch {
        applyEmpty();
      }
    };

    loadDaysInOfficeOverrides();
    return () => {
      cancelled = true;
    };
  }, [pageReady, hasSelection, selectedMonth, selectedOffice]);

  const coffeeTotals = sumCoffeeRowFields(selectedAggregate?.coffeeRows ?? EMPTY_COFFEE_ROWS);
  const additionalTotals = sumAdditionalRowFields(selectedAggregate?.additionalRows ?? EMPTY_ADDITIONAL_ROWS);
  const sugarTotals = sumSugarRowFields(selectedAggregate?.sugarRows ?? EMPTY_SUGAR_ROWS);
  const selectedCoffeeTotalAmount = coffeeTotals.coffeeTotal;
  const selectedOjTotalAmount = coffeeTotals.orangeJuiceTotal;
  const selectedSugarTotalAmount = sugarTotals.sugar;
  const selectedSugarBillableTotalAmount = sugarTotals.sugarGood;
  const selectedPaperTotalAmount = sugarTotals.paper;
  const selectedCoffeeSalesTotal = coffeeTotals.sales;
  const selectedDoctorPreventativeTotal = coffeeTotals.doctorPreventative;
  const selectedDoctorRestorativeTotal = coffeeTotals.doctorRestorative;
  const selectedDoctorCraProductionTotal = coffeeTotals.doctorCraProduction;

  const handleStartEdit = () => {
    setSaveMessage('');
    setDaysOverridesDraft(daysOverridesSaved);
    setIsEditing(true);
  };

  const handleCancelEdit = () => {
    setDaysOverridesDraft(daysOverridesSaved);
    setIsEditing(false);
  };

  const handleDaysInOfficeChange = (key: string, rawValue: string) => {
    const trimmed = rawValue.trim();
    if (trimmed === '') {
      setDaysOverridesDraft((prev) => {
        const next = { ...prev };
        delete next[key];
        return next;
      });
      return;
    }
    const parsed = parseNumber(trimmed);
    if (Number.isFinite(parsed) && parsed >= 0) {
      setDaysOverridesDraft((prev) => ({
        ...prev,
        [key]: parsed,
      }));
    }
  };

  const handleSaveEdit = async () => {
    if (!selectedAggregate) return;
    try {
      const docId = getMonthlyProductionDocId(selectedMonth, selectedOffice);
      const docRef = doc(db, 'monthly production', docId);
      const existingSnap = await getDoc(docRef);
      const existing = existingSnap.exists() ? (existingSnap.data() as Record<string, unknown>) : {};

      await setDoc(
        docRef,
        {
          month: getWhiteBearMonthDocId(selectedMonth),
          location: getMonthlyLocationFieldKey(selectedOffice),
          locationName: selectedOffice,
          updatedAt: serverTimestamp(),
          daysInOfficeOverrides: daysOverridesDraft,
          ...(existing.officeSummaryGoals !== undefined
            ? { officeSummaryGoals: existing.officeSummaryGoals }
            : {}),
          ...(existing.oeTableGoals !== undefined ? { oeTableGoals: existing.oeTableGoals } : {}),
        },
        { merge: true }
      );
      setDaysOverridesSaved(daysOverridesDraft);
      setIsEditing(false);
      setSaveMessage('Saved!');
    } catch (e: unknown) {
      setSaveMessage(formatActionError('Failed to save', e));
    }
  };

  const handleDownloadPdf = async () => {
    if (!selectedAggregate) return;
    try {
      setIsPdfDownloading(true);
      setSaveMessage('');

      const generatedDate = new Date().toLocaleDateString('en-US', {
        timeZone: 'America/Los_Angeles',
      });
      const pdfDoc = createMonthlyReportPdfDocument({
        aggregate: selectedAggregate,
        generatedDate,
        daysInOfficeOverrides: daysOverridesSaved,
      });
      const blob = await pdf(pdfDoc).toBlob();
      const filename = sanitizeFilename(
        `${selectedAggregate.month}_${selectedAggregate.location}_monthly-report.pdf`
      );
      downloadPdfBlob(blob, filename);
      setSaveMessage('PDF downloaded!');
    } catch (e: unknown) {
      setSaveMessage(formatActionError('Failed to download the pdf', e));
    } finally {
      setIsPdfDownloading(false);
    }
  };

  if (!pageReady) {
    return <main style={{ minHeight: '100vh', background: '#fff' }} />;
  }

  return (
    <main style={{ minHeight: '100vh', background: '#fff', padding: 24 }}>
      <section style={{ maxWidth: 2800, margin: '0 auto', padding: 20 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 10, marginBottom: 14 }}>
          <div>
            <h2 style={{ margin: 0, fontSize: 22, fontWeight: 700 }}>Monthly Production</h2>
            {hasSelection && (
              <p style={{ margin: '6px 0 0', color: '#475569', fontSize: 14 }}>
                {selectedMonth} · {selectedOffice}
              </p>
            )}
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            {isEditing ? (
              <>
                <button
                  type="button"
                  onClick={handleSaveEdit}
                  style={{
                    height: 36,
                    padding: '0 14px',
                    borderRadius: 8,
                    border: '1px solid #16a34a',
                    background: '#16a34a',
                    color: '#fff',
                    fontWeight: 700,
                    cursor: 'pointer',
                  }}
                >
                  Save
                </button>
                <button
                  type="button"
                  onClick={handleCancelEdit}
                  style={{
                    height: 36,
                    padding: '0 14px',
                    borderRadius: 8,
                    border: '1px solid #cbd5e1',
                    background: '#fff',
                    color: '#111827',
                    fontWeight: 700,
                    cursor: 'pointer',
                  }}
                >
                  Cancel
                </button>
              </>
            ) : (
              <button
                type="button"
                onClick={handleStartEdit}
                disabled={!selectedAggregate}
                style={{
                  height: 36,
                  padding: '0 14px',
                  borderRadius: 8,
                  border: '1px solid #334155',
                  background: !selectedAggregate ? '#e2e8f0' : '#475569',
                  color: '#fff',
                  fontWeight: 700,
                  cursor: !selectedAggregate ? 'not-allowed' : 'pointer',
                }}
              >
                Edit
              </button>
            )}
            <button
              type="button"
              onClick={handleDownloadPdf}
              disabled={!selectedAggregate || isPdfDownloading || isEditing}
              style={{
                height: 36,
                padding: '0 14px',
                borderRadius: 8,
                border: '1px solid #1d4ed8',
                background: !selectedAggregate || isPdfDownloading || isEditing ? '#bfdbfe' : '#2563eb',
                color: '#fff',
                fontWeight: 700,
                cursor: !selectedAggregate || isPdfDownloading || isEditing ? 'not-allowed' : 'pointer',
              }}
            >
              {isPdfDownloading ? 'Downloading...' : 'Download PDF'}
            </button>
          </div>
        </div>
        {saveMessage && <p style={{ margin: '0 0 10px', color: saveMessage.includes('Failed') ? '#b91c1c' : '#166534' }}>{saveMessage}</p>}
        {!hasSelection && (
          <p style={{ margin: 0, color: '#6b7280' }}>Month and office details need to be included in the URL.</p>
        )}
        {hasSelection && error && <p style={{ margin: 0, color: '#b91c1c' }}>{error}</p>}
        {hasSelection && !loading && !error && !selectedAggregate && (
          <p style={{ margin: 0, color: '#6b7280' }}>No completed monthly data is available for the selected Month/Office.</p>
        )}
        {hasSelection && !loading && !error && selectedAggregate && (
          <div style={{ padding: 12 }}>
                  <SummaryTable
                    minWidth={900}
                    headers={[
                      '',
                      'Submitted Production',
                      'CRA Production',
                      'Production W/Out CRA',
                      'Prophy @ OE',
                      'Prophy @ TX',
                      'Just Prophy',
                      'Actual Prophy',
                    ]}
                    totalRow={[
                      'Total',
                      formatDollarAmount(selectedAggregate.grandTotal),
                      formatDollarAmount(selectedAggregate.coffeeSales),
                      formatDollarAmount(selectedAggregate.salesWithoutCoffee),
                      formatAmount(selectedAggregate.paperAtOrangeJuice),
                      formatAmount(selectedAggregate.paperAtTea),
                      formatAmount(selectedAggregate.justPaper),
                      formatAmount(selectedAggregate.prophyTotal),
                    ]}
                    dailyAverageRow={[
                      'Daily Average',
                      formatDollarDailyAverage(selectedAggregate.grandTotal, selectedAggregate.docCount),
                      formatDollarDailyAverage(selectedAggregate.coffeeSales, selectedAggregate.docCount),
                      formatDollarDailyAverage(selectedAggregate.salesWithoutCoffee, selectedAggregate.docCount),
                      formatDailyAverage(selectedAggregate.paperAtOrangeJuice, selectedAggregate.docCount),
                      formatDailyAverage(selectedAggregate.paperAtTea, selectedAggregate.docCount),
                      formatDailyAverage(selectedAggregate.justPaper, selectedAggregate.docCount),
                      formatDailyAverage(selectedAggregate.prophyTotal, selectedAggregate.docCount),
                    ]}
                  />

                  <SummaryTable
                    minWidth={700}
                    headers={['', 'Preventative', 'Restorative', 'CRA Production', '1st Review Production', 'Mailed Production']}
                    totalRow={[
                      'Total',
                      formatDollarAmount(selectedAggregate.pineapple),
                      formatDollarAmount(selectedAggregate.rose),
                      formatDollarAmount(selectedAggregate.coffeeSales),
                      formatDollarAmount(selectedAggregate.pineapple + selectedAggregate.rose + selectedAggregate.coffeeSales),
                      formatDollarAmount(selectedAggregate.mailedProduction),
                    ]}
                    dailyAverageRow={[
                      'Daily Average',
                      formatDollarDailyAverage(selectedAggregate.pineapple, selectedAggregate.docCount),
                      formatDollarDailyAverage(selectedAggregate.rose, selectedAggregate.docCount),
                      formatDollarDailyAverage(selectedAggregate.coffeeSales, selectedAggregate.docCount),
                      formatDollarDailyAverage(
                        selectedAggregate.pineapple + selectedAggregate.rose + selectedAggregate.coffeeSales,
                        selectedAggregate.docCount
                      ),
                      formatDollarDailyAverage(selectedAggregate.mailedProduction, selectedAggregate.docCount),
                    ]}
                  />

                  <div style={{ overflowX: 'auto', marginBottom: 16 }}>
                    <table style={{ width: '100%', minWidth: 500, borderCollapse: 'collapse', fontSize: 14 }}>
                      <thead>
                        <tr>
                          {['Add On', 'No Shows', 'Scheduled', 'Seen', 'Seen %', 'Referral', 'Postcard Count'].map((h, idx) => (
                            <th key={h} style={tableHeaderStyle}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        <tr style={tableFooterRowStyle}>
                          <td style={tableCellStyle}>{formatAmount(selectedAggregate.visitsAdd)}</td>
                          <td style={tableCellStyle}>{formatAmount(selectedAggregate.visitsNoShow)}</td>
                          <td style={tableCellStyle}>{formatAmount(selectedAggregate.visitsScheduled)}</td>
                          <td style={tableCellStyle}>{formatAmount(selectedAggregate.visitsSeen)}</td>
                          <td style={tableCellStyle}>
                            {formatPercentage(selectedAggregate.visitsSeen, selectedAggregate.visitsScheduled)}
                          </td>
                          <td style={tableCellStyle}>{formatAmount(selectedAggregate.visitsReferral)}</td>
                          <td style={tableCellStyle}>{formatAmount(selectedAggregate.visitsPostcard)}</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>

                  <div style={{ overflowX: 'auto', marginBottom: 12 }}>
                    <h4 style={{ margin: '0 0 8px' }}>Production</h4>
                    <table style={{ width: '100%', minWidth: 1200, borderCollapse: 'collapse', fontSize: 14 }}>
                      <thead>
                        <tr>
                          {['Position', 'Name', 'Days in Office', 'Preventative', 'Restorative', 'CRA Production', 'Production', 'Production Average', 'Production %'].map((h, idx) => (
                            <th key={`${h}_${idx}`} style={tableHeaderStyle}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {selectedAggregate.coffeeRows.length === 0 ? (
                          <tr><td colSpan={9} style={tableEmptyCellStyle}>No data is available.</td></tr>
                        ) : (
                          selectedAggregate.coffeeRows.map((r) => (
                            <tr key={`${r.position}_${r.name}`}>
                              <td style={tableCellStyle}>{r.position}</td>
                              <td style={tableCellStyle}>{r.name}</td>
                              <DaysInOfficeCell
                                personKeyValue={personKey(r.position, r.name)}
                                value={activeDaysOverrides[personKey(r.position, r.name)]}
                                isEditing={isEditing}
                                onChange={handleDaysInOfficeChange}
                              />
                              <td style={tableCellStyle}>{formatDollarAmount(r.doctorPreventative)}</td>
                              <td style={tableCellStyle}>{formatDollarAmount(r.doctorRestorative)}</td>
                              <td style={tableCellStyle}>{formatDollarAmount(r.doctorCraProduction)}</td>
                              <td style={tableCellStyle}>{formatDollarAmount(r.sales)}</td>
                              <td style={tableCellStyle}>{formatDollarDailyAverageWithDays(r.sales, getDaysInOfficeValue(activeDaysOverrides, r.position, r.name))}</td>
                              <td style={tableCellStyle}>{formatPercentage(r.sales, selectedCoffeeSalesTotal)}</td>
                            </tr>
                          ))
                        )}
                      </tbody>
                      <tfoot>
                        <tr style={tableFooterRowStyle}>
                          <td style={tableCellStyle}>Total</td>
                          <td style={tableCellStyle}>
                            {countRowsByPosition(selectedAggregate.coffeeRows, 'doctor')}
                          </td>
                          <td style={tableCellStyle}>
                            {renderDaysInOfficeTotal(activeDaysOverrides, selectedAggregate.coffeeRows)}
                          </td>
                          <td style={tableCellStyle}>{formatDollarAmount(selectedDoctorPreventativeTotal)}</td>
                          <td style={tableCellStyle}>{formatDollarAmount(selectedDoctorRestorativeTotal)}</td>
                          <td style={tableCellStyle}>{formatDollarAmount(selectedDoctorCraProductionTotal)}</td>
                          <td style={tableCellStyle}>{formatDollarAmount(selectedCoffeeSalesTotal)}</td>
                          <td style={tableCellStyle}>{formatDollarDailyAverageWithDays(selectedCoffeeSalesTotal, totalDaysForAverage(sumDaysInOfficeFromOverrides(activeDaysOverrides, selectedAggregate.coffeeRows)))}</td>
                          <td style={tableCellStyle}>100%</td>
                        </tr>
                        <tr style={tableFooterRowStyle}>
                          <td style={tableCellStyle}>Daily Average</td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>{formatDollarDailyAverage(selectedDoctorPreventativeTotal, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>{formatDollarDailyAverage(selectedDoctorRestorativeTotal, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>{formatDollarDailyAverage(selectedDoctorCraProductionTotal, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>{formatDollarDailyAverage(selectedCoffeeSalesTotal, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>-</td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>

                  <div style={{ overflowX: 'auto', marginBottom: 12 }}>
                    <h4 style={{ margin: '0 0 8px' }}>CRA</h4>
                    <table style={{ width: '100%', minWidth: 1300, borderCollapse: 'collapse', fontSize: 14 }}>
                      <thead>
                        <tr>
                          {['Position', 'Name', 'Days in Office', 'CRA (New)', 'CRA (Return)', 'CRA Total', 'CRA Average', 'CRA %', 'CRA (Not Billable)', 'Rendered CRA', 'CRA (Billable)'].map((h, idx) => (
                            <th key={`${h}_${idx}`} style={tableHeaderStyle}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {selectedAggregate.coffeeRows.length === 0 ? (
                          <tr><td colSpan={11} style={tableEmptyCellStyle}>No data is available.</td></tr>
                        ) : (
                          selectedAggregate.coffeeRows.map((r) => (
                            <tr key={`${r.position}_${r.name}_coffee`}>
                              <td style={tableCellStyle}>{r.position}</td>
                              <td style={tableCellStyle}>{r.name}</td>
                              <DaysInOfficeCell
                                personKeyValue={personKey(r.position, r.name)}
                                value={activeDaysOverrides[personKey(r.position, r.name)]}
                                isEditing={isEditing}
                                onChange={handleDaysInOfficeChange}
                              />
                              <td style={tableCellStyle}>{formatAmount(r.coffeeNew)}</td>
                              <td style={tableCellStyle}>{formatAmount(r.coffeeReturn)}</td>
                              <td style={tableCellStyle}>{formatAmount(r.coffeeTotal)}</td>
                              <td style={tableCellStyle}>{formatDailyAverageWithDays(r.coffeeTotal, getDaysInOfficeValue(activeDaysOverrides, r.position, r.name))}</td>
                              <td style={tableCellStyle}>{formatPercentage(r.coffeeTotal, selectedCoffeeTotalAmount)}</td>
                              <td style={tableCellStyle}>{formatAmount(r.coffeeNo)}</td>
                              <td style={tableCellStyle}>{formatAmount(r.renderedCoffee)}</td>
                              <td style={tableCellStyle}>{formatAmount(r.coffeeYes)}</td>
                            </tr>
                          ))
                        )}
                      </tbody>
                      <tfoot>
                        <tr style={tableFooterRowStyle}>
                          <td style={tableCellStyle}>Total</td>
                          <td style={tableCellStyle}>
                            {countRowsByPosition(selectedAggregate.coffeeRows, 'doctor')}
                          </td>
                          <td style={tableCellStyle}>
                            {renderDaysInOfficeTotal(activeDaysOverrides, selectedAggregate.coffeeRows)}
                          </td>
                          <td style={tableCellStyle}>{formatAmount(coffeeTotals.coffeeNew)}</td>
                          <td style={tableCellStyle}>{formatAmount(coffeeTotals.coffeeReturn)}</td>
                          <td style={tableCellStyle}>{formatAmount(selectedCoffeeTotalAmount)}</td>
                          <td style={tableCellStyle}>{formatDailyAverageWithDays(selectedCoffeeTotalAmount, totalDaysForAverage(sumDaysInOfficeFromOverrides(activeDaysOverrides, selectedAggregate.coffeeRows)))}</td>
                          <td style={tableCellStyle}>100%</td>
                          <td style={tableCellStyle}>{formatAmount(coffeeTotals.coffeeNo)}</td>
                          <td style={tableCellStyle}>{formatAmount(coffeeTotals.renderedCoffee)}</td>
                          <td style={tableCellStyle}>{formatAmount(coffeeTotals.coffeeYes)}</td>
                        </tr>
                        <tr style={tableFooterRowStyle}>
                          <td style={tableCellStyle}>Daily Average</td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>{formatDailyAverage(coffeeTotals.coffeeNew, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>{formatDailyAverage(coffeeTotals.coffeeReturn, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>{formatDailyAverage(selectedCoffeeTotalAmount, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>{formatDailyAverage(coffeeTotals.coffeeNo, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>{formatDailyAverage(coffeeTotals.renderedCoffee, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>{formatDailyAverage(coffeeTotals.coffeeYes, selectedAggregate.docCount)}</td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>

                  <div style={{ overflowX: 'auto' }}>
                    <h4 style={{ margin: '0 0 8px' }}>OE</h4>
                    <table style={{ width: '100%', minWidth: 1150, borderCollapse: 'collapse', fontSize: 14 }}>
                      <thead>
                        <tr>
                          {['Position', 'Name', 'Days in Office', 'OE (NP)', 'OE (RC)', 'OE Total', 'OE Average', 'OE %', 'Actual OE (NP)', 'Actual OE (RC)'].map((h, idx) => (
                            <th key={`${h}_${idx}`} style={tableHeaderStyle}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {selectedAggregate.coffeeRows.length === 0 ? (
                          <tr><td colSpan={10} style={tableEmptyCellStyle}>No data is available.</td></tr>
                        ) : (
                          selectedAggregate.coffeeRows.map((r) => (
                            <tr key={`${r.position}_${r.name}_oj`}>
                              <td style={tableCellStyle}>{r.position}</td>
                              <td style={tableCellStyle}>{r.name}</td>
                              <DaysInOfficeCell
                                personKeyValue={personKey(r.position, r.name)}
                                value={activeDaysOverrides[personKey(r.position, r.name)]}
                                isEditing={isEditing}
                                onChange={handleDaysInOfficeChange}
                              />
                              <td style={tableCellStyle}>{formatAmount(r.orangeJuiceNew)}</td>
                              <td style={tableCellStyle}>{formatAmount(r.orangeJuiceReturn)}</td>
                              <td style={tableCellStyle}>{formatAmount(r.orangeJuiceTotal)}</td>
                              <td style={tableCellStyle}>{formatDailyAverageWithDays(r.orangeJuiceTotal, getDaysInOfficeValue(activeDaysOverrides, r.position, r.name))}</td>
                              <td style={tableCellStyle}>{formatPercentage(r.orangeJuiceTotal, selectedOjTotalAmount)}</td>
                              <td style={tableCellStyle}>-</td>
                              <td style={tableCellStyle}>-</td>
                            </tr>
                          ))
                        )}
                      </tbody>
                      <tfoot>
                        <tr style={tableFooterRowStyle}>
                          <td style={tableCellStyle}>Total</td>
                          <td style={tableCellStyle}>
                            {countRowsByPosition(selectedAggregate.coffeeRows, 'doctor')}
                          </td>
                          <td style={tableCellStyle}>
                            {renderDaysInOfficeTotal(activeDaysOverrides, selectedAggregate.coffeeRows)}
                          </td>
                          <td style={tableCellStyle}>{formatAmount(coffeeTotals.orangeJuiceNew)}</td>
                          <td style={tableCellStyle}>{formatAmount(coffeeTotals.orangeJuiceReturn)}</td>
                          <td style={tableCellStyle}>{formatAmount(selectedOjTotalAmount)}</td>
                          <td style={tableCellStyle}>{formatDailyAverageWithDays(selectedOjTotalAmount, totalDaysForAverage(sumDaysInOfficeFromOverrides(activeDaysOverrides, selectedAggregate.coffeeRows)))}</td>
                          <td style={tableCellStyle}>100%</td>
                          <td style={tableCellStyle}>{formatAmount(selectedAggregate.actualOrangeJuiceNew)}</td>
                          <td style={tableCellStyle}>{formatAmount(selectedAggregate.actualOrangeJuiceReturn)}</td>
                        </tr>
                        <tr style={tableFooterRowStyle}>
                          <td style={tableCellStyle}>Daily Average</td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>{formatDailyAverage(coffeeTotals.orangeJuiceNew, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>{formatDailyAverage(coffeeTotals.orangeJuiceReturn, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>{formatDailyAverage(selectedOjTotalAmount, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>{formatDailyAverage(selectedAggregate.actualOrangeJuiceNew, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>{formatDailyAverage(selectedAggregate.actualOrangeJuiceReturn, selectedAggregate.docCount)}</td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>

                  <h3 style={{ margin: '16px 0 10px' }}>Doctors</h3>
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', minWidth: 1100, borderCollapse: 'collapse', fontSize: 14 }}>
                      <thead>
                        <tr>
                          {['Position', 'Name', 'Days in Office', 'Patient Seen', 'Insurance', 'Cash', 'Dentical', 'Treatment', 'Primary Teeth', 'Permanent Teeth', 'Total'].map((h, idx) => (
                            <th key={`${h}_${idx}`} style={tableHeaderStyle}>
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {selectedAggregate.additionalRows.length === 0 ? (
                          <tr><td colSpan={11} style={tableEmptyCellStyle}>No data is available.</td></tr>
                        ) : (
                          selectedAggregate.additionalRows.map((r) => (
                            <tr key={`${r.position}_${r.name}`}>
                              <td style={tableCellStyle}>{r.position}</td>
                              <td style={tableCellStyle}>{r.name}</td>
                              <DaysInOfficeCell
                                personKeyValue={personKey(r.position, r.name)}
                                value={activeDaysOverrides[personKey(r.position, r.name)]}
                                isEditing={isEditing}
                                onChange={handleDaysInOfficeChange}
                              />
                              <td style={tableCellStyle}>{formatAmount(r.customer)}</td>
                              <td style={tableCellStyle}>{formatAmount(r.icecream)}</td>
                              <td style={tableCellStyle}>{formatAmount(r.cake)}</td>
                              <td style={tableCellStyle}>{formatAmount(r.donut)}</td>
                              <td style={tableCellStyle}>{formatAmount(r.tart)}</td>
                              <td style={tableCellStyle}>{formatAmount(r.peach)}</td>
                              <td style={tableCellStyle}>{formatAmount(r.peppermint)}</td>
                              <td style={tableCellStyle}>{formatAmount(r.total)}</td>
                            </tr>
                          ))
                        )}
                      </tbody>
                      <tfoot>
                        <tr style={tableFooterRowStyle}>
                          <td style={tableCellStyle}>Total</td>
                          <td style={tableCellStyle}>
                            {countRowsByPosition(selectedAggregate.additionalRows, 'doctor')}
                          </td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>{formatAmount(additionalTotals.customer)}</td>
                          <td style={tableCellStyle}>{formatAmount(additionalTotals.icecream)}</td>
                          <td style={tableCellStyle}>{formatAmount(additionalTotals.cake)}</td>
                          <td style={tableCellStyle}>{formatAmount(additionalTotals.donut)}</td>
                          <td style={tableCellStyle}>{formatAmount(additionalTotals.tart)}</td>
                          <td style={tableCellStyle}>{formatAmount(additionalTotals.peach)}</td>
                          <td style={tableCellStyle}>{formatAmount(additionalTotals.peppermint)}</td>
                          <td style={tableCellStyle}>{formatAmount(additionalTotals.total)}</td>
                        </tr>
                        <tr style={tableFooterRowStyle}>
                          <td style={tableCellStyle}>Daily Average</td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>{formatDailyAverage(additionalTotals.customer, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>{formatDailyAverage(additionalTotals.icecream, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>{formatDailyAverage(additionalTotals.cake, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>{formatDailyAverage(additionalTotals.donut, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>{formatDailyAverage(additionalTotals.tart, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>{formatDailyAverage(additionalTotals.peach, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>{formatDailyAverage(additionalTotals.peppermint, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>{formatDailyAverage(additionalTotals.total, selectedAggregate.docCount)}</td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>

                  <h3 style={{ margin: '16px 0 10px' }}>Sealant & Prophy</h3>
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', minWidth: 1300, borderCollapse: 'collapse', fontSize: 14 }}>
                      <thead>
                        <tr>
                          {['Position', 'Name', 'Days in Office', 'Sealant', 'Sealant (Billable)', 'Sealant Average', 'Sealant %', 'Sealant (Redo)', 'Prophy', 'Prophy Average', 'Prophy %'].map((h, idx) => (
                            <th key={`${h}_${idx}`} style={tableHeaderStyle}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {selectedAggregate.sugarRows.length === 0 ? (
                          <tr><td colSpan={11} style={tableEmptyCellStyle}>No data is available.</td></tr>
                        ) : (
                          selectedAggregate.sugarRows.map((r, index) => {
                            const rowCellStyle = tableCellStyleWithPositionSeparator(
                              isPositionGroupStart(selectedAggregate.sugarRows, index)
                            );
                            return (
                            <tr key={`${r.position}_${r.name}`}>
                              <td style={rowCellStyle}>{r.position}</td>
                              <td style={rowCellStyle}>{r.name}</td>
                              <DaysInOfficeCell
                                personKeyValue={personKey(r.position, r.name)}
                                value={activeDaysOverrides[personKey(r.position, r.name)]}
                                isEditing={isEditing}
                                onChange={handleDaysInOfficeChange}
                                cellStyle={rowCellStyle}
                              />
                              <td style={rowCellStyle}>{formatAmount(r.sugar)}</td>
                              <td style={rowCellStyle}>{formatAmount(r.sugarGood)}</td>
                              <td style={rowCellStyle}>{formatDailyAverageWithDays(r.sugarGood, getDaysInOfficeValue(activeDaysOverrides, r.position, r.name))}</td>
                              <td style={rowCellStyle}>{formatPercentage(r.sugarGood, selectedSugarBillableTotalAmount)}</td>
                              <td style={rowCellStyle}>{formatAmount(r.sugarBad)}</td>
                              <td style={rowCellStyle}>{formatAmount(r.paper)}</td>
                              <td style={rowCellStyle}>{formatDailyAverageWithDays(r.paper, getDaysInOfficeValue(activeDaysOverrides, r.position, r.name))}</td>
                              <td style={rowCellStyle}>{formatPercentage(r.paper, selectedPaperTotalAmount)}</td>
                            </tr>
                            );
                          })
                        )}
                      </tbody>
                      <tfoot>
                        {(['doctor', 'rda'] as const).map((position) => {
                          const positionRows = selectedAggregate.sugarRows.filter(
                            (row) => normalizePositionLabel(row.position) === position
                          );
                          if (positionRows.length === 0) return null;
                          const sealantTotal = sumSugarRows(positionRows, (row) => row.sugar);
                          const billableTotal = sumSugarRows(positionRows, (row) => row.sugarGood);
                          return (
                            <tr key={`${position}_sealant_total`} style={tableSubtotalRowStyle}>
                              <td style={tableCellStyle}>{position} Total</td>
                              <td style={tableCellStyle}>-</td>
                              <td style={tableCellStyle}>-</td>
                              <td style={tableCellStyle}>{formatAmount(sealantTotal)}</td>
                              <td style={tableCellStyle}>{formatAmount(billableTotal)}</td>
                              <td colSpan={6} style={tableCellStyle} />
                            </tr>
                          );
                        })}
                        <tr style={tableFooterRowStyle}>
                          <td style={tableCellStyle}>Total</td>
                          <td style={tableCellStyle}>{selectedAggregate.sugarRows.length}</td>
                          <td style={tableCellStyle}>
                            {renderDaysInOfficeTotal(activeDaysOverrides, selectedAggregate.sugarRows)}
                          </td>
                          <td style={tableCellStyle}>{formatAmount(selectedSugarTotalAmount)}</td>
                          <td style={tableCellStyle}>{formatAmount(selectedSugarBillableTotalAmount)}</td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>100%</td>
                          <td style={tableCellStyle}>{formatAmount(sugarTotals.sugarBad)}</td>
                          <td style={tableCellStyle}>{formatAmount(selectedPaperTotalAmount)}</td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>100%</td>
                        </tr>
                        <tr style={tableFooterRowStyle}>
                          <td style={tableCellStyle}>Daily Average</td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>{formatDailyAverage(selectedSugarTotalAmount, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>{formatDailyAverage(selectedSugarBillableTotalAmount, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>{formatDailyAverage(sugarTotals.sugarBad, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>{formatDailyAverage(selectedPaperTotalAmount, selectedAggregate.docCount)}</td>
                          <td style={tableCellStyle}>-</td>
                          <td style={tableCellStyle}>-</td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>

                  <h3 style={{ margin: '16px 0 10px' }}>Sealant Diagnose</h3>
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '50%', minWidth: 700, borderCollapse: 'collapse', fontSize: 14 }}>
                      <thead>
                        <tr>
                          {['Name', 'Sealant Diagnose'].map((h, idx) => (
                            <th key={`${h}_${idx}`} style={tableHeaderStyle}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {selectedAggregate.diagnoseRows.length === 0 ? (
                          <tr><td colSpan={4} style={tableEmptyCellStyle}>No data is available.</td></tr>
                        ) : (
                          selectedAggregate.diagnoseRows.map((r) => (
                            <tr key={r.sName}>
                              <td style={tableCellStyle}>{r.sName}</td>
                              <td style={tableCellStyle}>{formatAmount(r.diagnose)}</td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  </div>

                  <h3 style={{ margin: '16px 0 10px' }}>Short Procedure</h3>
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', minWidth: 700, borderCollapse: 'collapse', fontSize: 14 }}>
                      <thead>
                        <tr>
                          {['Short Procedures', 'OE', 'Pro', 'CRA'].map((h, idx) => (
                            <th key={`${h}_${idx}`} style={tableHeaderStyle}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {selectedAggregate.reasonRows.length === 0 ? (
                          <tr><td colSpan={4} style={tableEmptyCellStyle}>No data is available.</td></tr>
                        ) : (
                          selectedAggregate.reasonRows.map((r) => (
                            <tr key={r.reason}>
                              <td style={tableCellStyle}>{r.reason}</td>
                              <td style={tableCellStyle}>{formatAmount(r.orangeJuice)}</td>
                              <td style={tableCellStyle}>{formatAmount(r.paper)}</td>
                              <td style={tableCellStyle}>{formatAmount(r.coffee)}</td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  </div>
          </div>
        )}
      </section>
    </main>
  );
}

export default function MonthlyPage() {
  return (
    <Suspense fallback={<main style={{ minHeight: '100vh', background: '#fff' }} />}>
      <MonthlyReportPageContent />
    </Suspense>
  );
}
