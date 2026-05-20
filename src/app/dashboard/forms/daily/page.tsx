'use client';

import React, { useEffect, useMemo, useState } from 'react';
import { db } from '@/lib/firebase.config';
import { collection, doc, getDocs, increment, serverTimestamp, setDoc } from 'firebase/firestore';
import { pdf, Document, Page, View, Text, StyleSheet } from '@react-pdf/renderer';

type ReportRow = {
  name?: string;
  timeStart?: string;
  timeEnd?: string;
  chartCount?: string;
  amount?: string;
};

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

type ReasonRow = {
  reason?: string;
  orangeJuice?: string;
  paper?: string;
  coffee?: string;
};

type ExtraInputRow = {
  position?: string;
  name?: string;
  /** Per-doctor row values (office-wide totals live in `locationSummary`). */
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
  pineapple?: string;
  rose?: string;
  total?: string;
};

type LocationSummary = {
  pineapple?: string;
  rose?: string;
  total?: string;
};

/** Visits-style metrics beside Production 1 (labels in first column). */
type ProductionSideMetrics = {
  add?: string;
  noShow?: string;
  scheduled?: string;
  seen?: string;
  seenPercent?: string;
};

const PRODUCTION_SIDE_METRIC_ROWS: { key: keyof ProductionSideMetrics; label: string }[] = [
  { key: 'add', label: 'Add On' },
  { key: 'noShow', label: 'No Shows' },
  { key: 'scheduled', label: 'Scheduled' },
  { key: 'seen', label: 'Seen' },
  { key: 'seenPercent', label: 'Seen %' },
];

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

type FormDoc = {
  id: string;
  date?: string;
  location?: string;
  reasonIfLate?: string;
  checkIn?: string;
  checkOut?: string;
  hoursOpen?: string;
  closer?: string;
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
  notesLines?: string[];
  notDueLines?: string[];
  name?: string;
  timeStart?: string;
  timeEnd?: string;
  chartCount?: string;
  amount?: string;
  reportRows?: ReportRow[];
  tableRows?: TableRow[];
  tableTotals?: Partial<TableTotals>;
  coffeeActualTotals?: Partial<Record<keyof TableTotals, string>>;
  sugarRows?: SugarRow[];
  sugarTotals?: Partial<SugarTotals>;
  reasonRows?: ReasonRow[];
  extraInputRows?: ExtraInputRow[];
  locationSummary?: LocationSummary;
  productionSideMetrics?: ProductionSideMetrics;
  edited?: boolean;
  editedAt?: unknown;
  pdfSaved?: boolean;
  pdfSavedAt?: unknown;
};

const TABLE_HEADERS = [
  'Position',
  'Name',
  'CRA (New)',
  'CRA (Return)',
  'CRA Total',
  'CRA (Not Billable)',
  'Rendered CRA',
  'CRA (Billable)',
  'OE (NP)',
  'OE (RC)',
  'OE Total',
  'Actual OE (NP)',
  'Actual OE (RC)',
];

const SUGAR_HEADERS = ['Position', 'Name', 'Sealant', 'Sealant (Billable)', 'Sealant (Redo)', 'Prophy'];
const REASON_HEADERS = ['Reasoning', 'OE', 'Pro', 'CRA'];

const NOTES_MAX_LENGTH = 300;

/** A4 가로 PDF 본문 너비(pt) — page padding 20×2 기준. */
const PDF_LANDSCAPE_CONTENT_WIDTH_PT = 841.89 - 40;
const PDF_NOTES_COLUMN_WIDTH_PT = Math.floor(PDF_LANDSCAPE_CONTENT_WIDTH_PT / 2) - 1;
const PDF_NOTES_TEXT_WIDTH_PT = PDF_NOTES_COLUMN_WIDTH_PT - 10;

/** 브라우저 number 스텝 화살표 제거 — 직접 입력만 (WebKit / Firefox). */
const D_PAGE_NUMBER_INPUT_SPINNER_RESET_CSS = `
.d-page-main input[type="number"]::-webkit-outer-spin-button,
.d-page-main input[type="number"]::-webkit-inner-spin-button {
  -webkit-appearance: none;
  margin: 0;
}
.d-page-main input[type="number"] {
  -moz-appearance: textfield;
  appearance: textfield;
}
`;

/** 오른쪽 상세 카드의 좌우 `padding`만큼 테이블 영역을 넓혀, 카드 안쪽 면과 표 격자 선이 이어지게 함 */
const D_PAGE_DETAIL_CARD_PADDING_PX = 12;
const dPageDetailTableBleedScroll: React.CSSProperties = {
  marginLeft: -D_PAGE_DETAIL_CARD_PADDING_PX,
  marginRight: -D_PAGE_DETAIL_CARD_PADDING_PX,
  width: `calc(100% + ${D_PAGE_DETAIL_CARD_PADDING_PX * 2}px)`,
  overflowX: 'auto',
};

const reportPdfStyles = StyleSheet.create({
  page: { padding: 20, fontFamily: 'Helvetica', fontSize: 8 },
  title: { fontSize: 14, fontWeight: 'bold', marginBottom: 5 },
  subtitle: { fontSize: 8, marginBottom: 10, color: '#374151' },
  section: { marginBottom: 10 },
  sectionTitle: { fontSize: 9, fontWeight: 'bold', marginBottom: 4 },
  table: { borderWidth: 0.6, borderColor: '#d1d5db' },
  tableRow: { flexDirection: 'row', borderBottomWidth: 0.5, borderColor: '#e5e7eb' },
  tableCell: { flex: 1, paddingVertical: 3, paddingHorizontal: 4, borderRightWidth: 0.5, borderColor: '#e5e7eb' },
  tableHeaderCell: { backgroundColor: '#f3f4f6' },
  tableHeaderText: { fontSize: 7, fontWeight: 'bold' },
  tableCellText: { fontSize: 7 },
  notesTableCell: {
    width: PDF_NOTES_COLUMN_WIDTH_PT,
    maxWidth: PDF_NOTES_COLUMN_WIDTH_PT,
    flexGrow: 0,
    flexShrink: 0,
    paddingVertical: 6,
    paddingHorizontal: 5,
    borderRightWidth: 0.5,
    borderColor: '#e5e7eb',
    overflow: 'hidden',
  },
  notesTableCellLast: {
    width: PDF_NOTES_COLUMN_WIDTH_PT,
    maxWidth: PDF_NOTES_COLUMN_WIDTH_PT,
    flexGrow: 0,
    flexShrink: 0,
    paddingVertical: 6,
    paddingHorizontal: 5,
    overflow: 'hidden',
  },
  notesTableText: { fontSize: 7, lineHeight: 1.35, width: PDF_NOTES_TEXT_WIDTH_PT },
});

function safeStr(value: unknown, maxLength = 80): string {
  if (value == null) return '';
  return String(value).trim().slice(0, maxLength).replace(/[<>]/g, '');
}

function sanitizeFilename(filename: string): string {
  return filename.replace(/[^a-zA-Z0-9._-]/g, '_').replace(/\.\./g, '_').slice(0, 255);
}

/** Notes / Not Due — 줄바꿈 유지, 칸 너비 안에서 줄 wrap (전체 trim 하지 않음). */
function safeNotesPdfText(value: unknown, maxLength = NOTES_MAX_LENGTH): string {
  if (value == null) return '';
  return String(value).slice(0, maxLength).replace(/[<>]/g, '');
}

function normalizeMultilineText(value: string): string {
  return value.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
}

function loadMultilineField(text: unknown, lines: unknown): string {
  if (Array.isArray(lines) && lines.length > 0) {
    return lines.map((line) => String(line)).join('\n');
  }
  return normalizeMultilineText(String(text ?? ''));
}

function multilineToLines(value: string): string[] {
  const normalized = normalizeMultilineText(value);
  if (normalized === '') return [];
  return normalized.split('\n');
}

function getMultilineFirestoreFields(notes: string, notDue: string) {
  const normalizedNotes = normalizeMultilineText(notes).slice(0, NOTES_MAX_LENGTH);
  const normalizedNotDue = normalizeMultilineText(notDue).slice(0, NOTES_MAX_LENGTH);
  return {
    notes: normalizedNotes,
    notesLines: multilineToLines(normalizedNotes),
    notDue: normalizedNotDue,
    notDueLines: multilineToLines(normalizedNotDue),
  };
}

/** react-pdf에서 긴 한 줄이 옆 칸으로 넘칠 때를 대비해 폭에 맞게 줄 분할. */
function wrapNotesLineForPdf(line: string, maxCharsPerLine = 72): string[] {
  if (line.length <= maxCharsPerLine) return [line];
  const chunks: string[] = [];
  let rest = line;
  while (rest.length > maxCharsPerLine) {
    let breakAt = rest.lastIndexOf(' ', maxCharsPerLine);
    if (breakAt <= 0) breakAt = maxCharsPerLine;
    chunks.push(rest.slice(0, breakAt));
    rest = rest.slice(breakAt).trimStart();
  }
  if (rest.length > 0) chunks.push(rest);
  return chunks;
}

function expandNotesTextToPdfLines(text: string): string[] {
  return text.split(/\r?\n/).flatMap((line) => wrapNotesLineForPdf(line));
}

function createPdfMultilineCellContent(
  styles: ReturnType<typeof StyleSheet.create>,
  raw: string
): React.ReactElement {
  const text = safeNotesPdfText(raw);
  if (text.trim() === '') {
    return React.createElement(Text, { style: styles.notesTableText }, '-');
  }
  const lines = expandNotesTextToPdfLines(text);
  return React.createElement(
    View,
    { style: { width: PDF_NOTES_TEXT_WIDTH_PT } },
    ...lines.map((line, i) =>
      React.createElement(
        Text,
        { key: `line-${i}`, style: styles.notesTableText },
        line === '' ? ' ' : line
      )
    )
  );
}

function createPdfNotesColumn(
  styles: ReturnType<typeof StyleSheet.create>,
  title: string,
  body: string,
  options: { isLast?: boolean }
) {
  const columnStyle = options.isLast ? styles.notesTableCellLast : styles.notesTableCell;
  return React.createElement(
    View,
    { style: columnStyle },
    React.createElement(
      View,
      { style: [styles.tableHeaderCell, { paddingVertical: 4, paddingHorizontal: 5 }] },
      React.createElement(Text, { style: styles.tableHeaderText }, title)
    ),
    React.createElement(
      View,
      { style: { paddingVertical: 6, paddingHorizontal: 5, width: PDF_NOTES_COLUMN_WIDTH_PT } },
      createPdfMultilineCellContent(styles, body)
    )
  );
}

function createPdfNotesTable(
  styles: ReturnType<typeof StyleSheet.create>,
  notDue: string,
  notes: string
) {
  return React.createElement(
    View,
    {
      style: {
        flexDirection: 'row',
        flexWrap: 'nowrap',
        width: PDF_LANDSCAPE_CONTENT_WIDTH_PT,
        borderWidth: 0.6,
        borderColor: '#d1d5db',
      },
    },
    createPdfNotesColumn(styles, 'Not Due', notDue, { isLast: false }),
    createPdfNotesColumn(styles, 'Notes', notes, { isLast: true })
  );
}

function createPdfTable(
  styles: ReturnType<typeof StyleSheet.create>,
  headers: string[],
  rows: string[][],
  keyPrefix: string
) {
  return React.createElement(
    View,
    { style: styles.table },
    React.createElement(
      View,
      { style: styles.tableRow, key: `${keyPrefix}-header` },
      ...headers.map((header, idx) =>
        React.createElement(
          View,
          { key: `${keyPrefix}-h-${idx}`, style: [styles.tableCell, styles.tableHeaderCell] },
          React.createElement(Text, { style: styles.tableHeaderText }, safeStr(header, 60) || '-')
        )
      )
    ),
    ...rows.map((row, rowIdx) =>
      React.createElement(
        View,
        { style: styles.tableRow, key: `${keyPrefix}-r-${rowIdx}` },
        ...headers.map((_, colIdx) =>
          React.createElement(
            View,
            { key: `${keyPrefix}-c-${rowIdx}-${colIdx}`, style: styles.tableCell },
            React.createElement(Text, { style: styles.tableCellText }, safeStr(row[colIdx] ?? '', 100) || '-')
          )
        )
      )
    )
  );
}

function createSubmittedReportPDFDocument(props: {
  date: string;
  location: string;
  generatedDate: string;
  reasonIfLate: string;
  checkIn: string;
  checkOut: string;
  hoursOpen: string;
  closer: string;
  submittedAt: string;
  grandTotal: string;
  coffeeSales: string;
  salesWithoutCoffee: string;
  paperAtOrangeJuice: string;
  paperAtTea: string;
  justPaper: string;
  prophyTotal: string;
  reportRows: ReportRow[];
  tableRows: TableRow[];
  tableTotals: TableTotals | null;
  actualOrangeNew: string;
  actualOrangeReturn: string;
  extraInputRows: ExtraInputRow[];
  extraInputTotals: ReturnType<typeof getExtraInputTotals>;
  locationSummary: {
    pineapple: string;
    rose: string;
    coffeeSales: string;
    total: string;
  };
  productionSideMetrics: ProductionSideMetrics;
  sugarRows: SugarRow[];
  sugarTotals: SugarTotals | null;
  reasonRows: ReasonRow[];
  notes: string;
  notDue: string;
}) {
  const s = reportPdfStyles;
  const {
    date,
    location,
    generatedDate,
    reasonIfLate,
    checkIn,
    checkOut,
    hoursOpen,
    closer,
    submittedAt,
    grandTotal,
    coffeeSales,
    salesWithoutCoffee,
    paperAtOrangeJuice,
    paperAtTea,
    justPaper,
    prophyTotal,
    reportRows,
    tableRows,
    tableTotals,
    actualOrangeNew,
    actualOrangeReturn,
    extraInputRows,
    extraInputTotals,
    locationSummary,
    productionSideMetrics,
    sugarRows,
    sugarTotals,
    reasonRows,
    notes,
    notDue,
  } = props;
  const sanitizedGeneratedDate = safeStr(generatedDate, 80).replace(/\$r/gi, ' ').replace(/\$/g, '').replace(/[\r\n]/g, ' ');

  const reportTable = createPdfTable(
    s,
    ['Name', 'Start', 'End', 'Chart', 'Amount'],
    reportRows.map((row) => [
      safeStr(row.name, 30),
      safeStr(row.timeStart, 20),
      safeStr(row.timeEnd, 20),
      safeStr(row.chartCount, 12),
      safeStr(formatCurrencyLabel(row.amount), 12),
    ]),
    'report'
  );

  const salesPaperSummaryTable = createPdfTable(
    s,
    [
      'Grand Total',
      'CRA Production',
      'Production W/Out CRA',
      'Prophy @ OE',
      'Prophy @ TX',
      'Just Prophy',
      'Prophy Total',
    ],
    [
      [
        formatCurrencyLabel(grandTotal || '-'),
        formatCurrencyLabel(coffeeSales || '-'),
        formatCurrencyLabel(salesWithoutCoffee || '-'),
        paperAtOrangeJuice || '-',
        paperAtTea || '-',
        justPaper || '-',
        prophyTotal || '-',
      ],
    ],
    'sales-paper-summary'
  );
  const submissionInfoTable = createPdfTable(
    s,
    ['Reason if Late', 'Submitted by office at:'],
    [[reasonIfLate || '-', submittedAt || '-']],
    'submission-info'
  );

  const officeHoursTable = createPdfTable(
    s,
    ['Check In', 'Check Out', 'Hours Open', 'Closer'],
    [[checkIn || '-', checkOut || '-', formatHoursOpenLabel(hoursOpen), closer || '-']],
    'office-hours'
  );

  const locationSummaryTable = createPdfTable(
    s,
    ['Preventative', 'Restorative', 'CRA Production', 'Total'],
    [
      [
        formatCurrencyLabel(locationSummary.pineapple || '-'),
        formatCurrencyLabel(locationSummary.rose || '-'),
        formatCurrencyLabel(locationSummary.coffeeSales || '-'),
        formatCurrencyLabel(locationSummary.total || '-'),
      ],
    ],
    'location-summary'
  );

  const sideMetrics = {
    ...createEmptyProductionSideMetrics(),
    ...productionSideMetrics,
  };
  const visitsPdfHeaders = PRODUCTION_SIDE_METRIC_ROWS.map(({ label }) => label);
  const visitsPdfValues = PRODUCTION_SIDE_METRIC_ROWS.map(({ key }) =>
    key === 'seenPercent'
      ? safeStr(formatSeenPercentDisplay(computeSeenPercentRounded(sideMetrics.scheduled, sideMetrics.seen)), 40) || '-'
      : safeStr(sideMetrics[key], 40) || '-'
  );
  const productionSideMetricsTable = createPdfTable(s, visitsPdfHeaders, [visitsPdfValues], 'production-side-metrics');

  const productionTotalForPdf = String(
    roundToCents(
      extraInputRows.reduce((acc, row, idx) => acc + getDoctorPerformanceRowProductionValue(row, tableRows[idx]?.sales), 0)
    )
  );
  const additionalInputsTable = createPdfTable(
    s,
    [
      'Position',
      'Name',
      'Preventative',
      'Restorative',
      'CRA Production',
      'Production',
      'Patient Seen',
      'Insurance',
      'Cash',
      'Dentical',
      'Treatment',
      'Primary Teeth',
      'Permanent Teeth',
    ],
    [
      ...extraInputRows.map((row, idx) => [
        safeStr(row.position, 20),
        safeStr(row.name, 24),
        safeStr(formatCurrencyLabel(row.doctorPreventative), 12),
        safeStr(formatCurrencyLabel(row.doctorRestorative), 12),
        safeStr(formatCurrencyLabel(row.doctorCraProduction), 12),
        safeStr(formatCurrencyLabel(formatDoctorPerformanceProductionCell(row, tableRows[idx]?.sales)), 12),
        safeStr(row.customer, 10),
        safeStr(row.icecream, 10),
        safeStr(row.cake, 10),
        safeStr(row.donut, 10),
        safeStr(row.tart, 10),
        safeStr(row.peach, 10),
        safeStr(row.peppermint, 10),
      ]),
      [
        'Total',
        '-',
        formatCurrencyLabel(String(extraInputTotals.doctorPreventative)),
        formatCurrencyLabel(String(extraInputTotals.doctorRestorative)),
        formatCurrencyLabel(String(extraInputTotals.doctorCraProduction)),
        formatCurrencyLabel(productionTotalForPdf),
        String(extraInputTotals.customer),
        String(extraInputTotals.icecream),
        String(extraInputTotals.cake),
        String(extraInputTotals.donut),
        String(extraInputTotals.tart),
        String(extraInputTotals.peach),
        String(extraInputTotals.peppermint),
      ],
    ],
    'additional-inputs'
  );

  const coffeeTable = createPdfTable(
    s,
    TABLE_HEADERS,
    [
      ...tableRows.map((row) => [
        safeStr(row.position, 16),
        safeStr(row.name, 20),
        safeStr(row.coffeeNew, 12),
        safeStr(row.coffeeReturn, 12),
        safeStr(row.coffeeTotal, 12),
        safeStr(row.coffeeNo, 12),
        safeStr(row.renderedCoffee, 12),
        safeStr(row.coffeeYes, 12),
        safeStr(row.orangeJuiceNew, 12),
        safeStr(row.orangeJuiceReturn, 12),
        safeStr(row.orangeJuiceTotal, 12),
        '-',
        '-',
      ]),
      [
        'Total',
        '-',
        String(tableTotals?.coffeeNew ?? 0),
        String(tableTotals?.coffeeReturn ?? 0),
        String(tableTotals?.coffeeTotal ?? 0),
        String(tableTotals?.coffeeNo ?? 0),
        String(tableTotals?.renderedCoffee ?? 0),
        String(tableTotals?.coffeeYes ?? 0),
        String(tableTotals?.orangeJuiceNew ?? 0),
        String(tableTotals?.orangeJuiceReturn ?? 0),
        String(tableTotals?.orangeJuiceTotal ?? 0),
        actualOrangeNew || '-',
        actualOrangeReturn || '-',
      ],
    ],
    'coffee-table'
  );

  const sugarTable = createPdfTable(
    s,
    SUGAR_HEADERS,
    [
      ...sugarRows.map((row) => [
        safeStr(row.position, 16),
        safeStr(row.name, 20),
        safeStr(row.sugar, 10),
        safeStr(row.sugarGood, 10),
        safeStr(row.sugarBad, 10),
        safeStr(row.paper, 10),
      ]),
      [
        'Total',
        '-',
        String(sugarTotals?.sugar ?? 0),
        String(sugarTotals?.sugarGood ?? 0),
        String(sugarTotals?.sugarBad ?? 0),
        String(sugarTotals?.paper ?? 0),
      ],
    ],
    'sugar-table'
  );

  const reasonTable = createPdfTable(
    s,
    REASON_HEADERS,
    reasonRows.map((row) => [
      safeStr(row.reason, 40),
      safeStr(row.orangeJuice, 12),
      safeStr(row.paper, 12),
      safeStr(row.coffee, 12),
    ]),
    'reason-table'
  );

  const notesTable = createPdfNotesTable(s, notDue, notes);

  return React.createElement(
    Document,
    null,
    React.createElement(
      Page,
      { size: 'A4', orientation: 'landscape', style: s.page },
      React.createElement(Text, { style: s.title }, `${location || '-'} Finalized Production - ${date || '/'} `),
      React.createElement(
        Text,
        { style: s.subtitle },
        `Generated: ${sanitizedGeneratedDate || '-'}`
      ),
      React.createElement(View, { style: s.section }, React.createElement(Text, { style: s.sectionTitle }, 'Submission Info'), submissionInfoTable),
      React.createElement(View, { style: s.section }, React.createElement(Text, { style: s.sectionTitle }, 'Office Hours'), officeHoursTable),
      React.createElement(View, { style: s.section }, React.createElement(Text, { style: s.sectionTitle }, 'Billers'), reportTable),
      React.createElement(
        View,
        { style: s.section },
        React.createElement(Text, { style: s.sectionTitle }, 'Production'),
        salesPaperSummaryTable
      ),
      React.createElement(
        View,
        { style: s.section },
        locationSummaryTable,
        React.createElement(Text, { style: { ...s.sectionTitle, marginTop: 10 } }, 'Visits'),
        productionSideMetricsTable
      ),
      React.createElement(View, { style: s.section }, React.createElement(Text, { style: s.sectionTitle }, 'Doctors'), additionalInputsTable),
      React.createElement(View, { style: s.section }, React.createElement(Text, { style: s.sectionTitle }, 'CRA / OE'), coffeeTable),
      React.createElement(View, { style: s.section }, React.createElement(Text, { style: s.sectionTitle }, 'Sealant / Prophy'), sugarTable),
      React.createElement(View, { style: s.section }, React.createElement(Text, { style: s.sectionTitle }, 'Short Procedures'), reasonTable),
      React.createElement(View, { style: s.section }, React.createElement(Text, { style: s.sectionTitle }, 'Notes'), notesTable)
    )
  );
}

function parseNumber(value: unknown): number {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

/** 달러 입력: 쉼표·$·공백 제거 후 숫자. */
function parseMoney(value: unknown): number {
  const s = String(value ?? '')
    .trim()
    .replace(/^\$/, '')
    .replace(/,/g, '')
    .trim();
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function roundToCents(n: number): number {
  return Math.round(n * 100) / 100;
}

function formatTime12h(hours24: number, minutes: number): string {
  let hours = hours24 % 24;
  const ampm = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12;
  if (hours === 0) hours = 12;
  return `${hours}:${String(minutes).padStart(2, '0')} ${ampm}`;
}

function parseTimeToMinutes(timeStr: unknown): number | null {
  const t = String(timeStr ?? '').trim();
  if (!t) return null;

  const ampmMatch = t.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)$/i);
  if (ampmMatch) {
    let hours = Number(ampmMatch[1]);
    const minutes = Number(ampmMatch[2]);
    const ampm = ampmMatch[3].toUpperCase();
    if (!Number.isFinite(hours) || !Number.isFinite(minutes) || hours < 1 || hours > 12 || minutes > 59) {
      return null;
    }
    if (ampm === 'AM') {
      if (hours === 12) hours = 0;
    } else if (hours !== 12) {
      hours += 12;
    }
    return hours * 60 + minutes;
  }

  const match = t.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!match) return null;
  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  if (!Number.isFinite(hours) || !Number.isFinite(minutes) || hours > 23 || minutes > 59) return null;
  return hours * 60 + minutes;
}

/** Firestore(AM/PM 등) → type="time" 입력용 HH:mm */
function toTimeInputValue(stored: unknown): string {
  const minutes = parseTimeToMinutes(stored);
  if (minutes == null) return '';
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}

/** type="time" 입력값 → Firestore 저장용 AM/PM */
function toStoredCheckTime12h(localValue: string): string {
  const trimmed = localValue.trim();
  if (!trimmed) return '';
  const minutes = parseTimeToMinutes(trimmed);
  if (minutes == null) return trimmed;
  return formatTime12h(Math.floor(minutes / 60), minutes % 60);
}

function formatMinutesAsHoursLabel(totalMinutes: number): string {
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  const parts: string[] = [];
  if (hours > 0) parts.push(`${hours} hour${hours !== 1 ? 's' : ''}`);
  if (minutes > 0) parts.push(`${minutes} minute${minutes !== 1 ? 's' : ''}`);
  return parts.length > 0 ? parts.join(' ') : '0 minutes';
}

/** Hours Open = Check Out − Check In → "5 hours 30 minutes" (익일 퇴근은 +24h). */
function computeHoursOpen(checkIn: unknown, checkOut: unknown): string {
  const inMin = parseTimeToMinutes(checkIn);
  const outMin = parseTimeToMinutes(checkOut);
  if (inMin == null || outMin == null) return '';
  let diffMin = outMin - inMin;
  if (diffMin < 0) diffMin += 24 * 60;
  return formatMinutesAsHoursLabel(diffMin);
}

/** UI/PDF 표시 — 저장값이 옛 소수(5.5) 형식이면 변환. */
function formatHoursOpenLabel(hoursOpen: unknown): string {
  const raw = String(hoursOpen ?? '').trim();
  if (!raw) return '-';
  if (/hour|minute/i.test(raw)) return raw;
  const n = parseNumber(raw);
  if (Number.isFinite(n)) return formatMinutesAsHoursLabel(Math.round(n * 60));
  return raw;
}

/** Prophy Total = Prophy @ OE + Prophy @ TX + Just Prophy (셋 다 비어 있으면 ''). */
function computeProphyTotal(oe: unknown, tea: unknown, just: unknown): string {
  const hasBasis =
    String(oe ?? '').trim() !== '' || String(tea ?? '').trim() !== '' || String(just ?? '').trim() !== '';
  if (!hasBasis) return '';
  return String(Math.round(parseNumber(oe) + parseNumber(tea) + parseNumber(just)));
}

/** Read-only UI/PDF: leading $ when a value is present; '-' unchanged. */
function formatCurrencyLabel(value: unknown): string {
  const s = String(value ?? '').trim();
  if (!s || s === '-') return '-';
  const core = s.startsWith('$') ? s.slice(1).trim() : s;
  if (!core) return '-';
  return `$${core}`;
}

/** Append % for Seen % display; '-' stays '-'. */
function formatSeenPercentDisplay(computed: string): string {
  const t = String(computed ?? '').trim();
  if (!t || t === '-') return '-';
  const n = t.replace(/%$/, '').trim();
  return `${n}%`;
}

/** Seen % = round(Seen ÷ Scheduled × 100); '-' when Scheduled is 0. */
function computeSeenPercentRounded(scheduledRaw: unknown, seenRaw: unknown): string {
  const sch = parseNumber(scheduledRaw);
  if (sch === 0) return '-';
  const seen = parseNumber(seenRaw);
  return String(Math.round((seen / sch) * 100));
}

function seenPercentReadOnlyInputValueFromMetrics(m: ProductionSideMetrics | undefined): string {
  const s = computeSeenPercentRounded(m?.scheduled, m?.seen);
  return s === '-' ? '' : `${s}%`;
}

function getDoctorPerformanceProductionSum(row: ExtraInputRow | undefined): number {
  if (!row) return 0;
  return roundToCents(
    parseMoney(row.doctorPreventative) + parseMoney(row.doctorRestorative) + parseMoney(row.doctorCraProduction)
  );
}

function doctorPerformanceProductionRowHasInput(row: ExtraInputRow | undefined): boolean {
  if (!row) return false;
  return [row.doctorPreventative, row.doctorRestorative, row.doctorCraProduction].some((v) => String(v ?? '').trim() !== '');
}

/** Sum of Preventative + Restorative + CRA Production; if all three empty, keep saved Production (`sales`). */
function getDoctorPerformanceRowProductionValue(row: ExtraInputRow | undefined, fallbackSales?: string): number {
  const sum = getDoctorPerformanceProductionSum(row);
  if (sum !== 0 || doctorPerformanceProductionRowHasInput(row)) return sum;
  return roundToCents(parseMoney(fallbackSales));
}

function formatDoctorPerformanceProductionCell(row: ExtraInputRow | undefined, fallbackSales?: string): string {
  const v = getDoctorPerformanceRowProductionValue(row, fallbackSales);
  if (v === 0 && !doctorPerformanceProductionRowHasInput(row) && String(fallbackSales ?? '').trim() === '') return '-';
  return String(v);
}

/** Value for read-only number input (matches Dentical); empty string when the cell shows '-'. */
function doctorPerformanceProductionReadOnlyInputValue(row: ExtraInputRow | undefined, fallbackSales?: string): string {
  const s = formatDoctorPerformanceProductionCell(row, fallbackSales);
  return s === '-' ? '' : s;
}

function getMonthKey(dateValue: string | undefined): string {
  const raw = String(dateValue ?? '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw.slice(0, 7);
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
}

function toFirestoreKey(value: string | undefined): string {
  const raw = String(value ?? '').trim();
  if (!raw) return 'Unknown';
  return raw.replace(/[.#$/\[\]]/g, '_');
}

function getDurationLabel(start: string | undefined, end: string | undefined): string {
  if (!start || !end) return '-';
  const startDate = new Date(start);
  const endDate = new Date(end);
  if (Number.isNaN(startDate.getTime()) || Number.isNaN(endDate.getTime())) return '-';

  const diffMs = endDate.getTime() - startDate.getTime();
  if (diffMs < 0) return '-';

  const totalMinutes = Math.floor(diffMs / 60000);
  const days = Math.floor(totalMinutes / (24 * 60));
  const hours = Math.floor((totalMinutes % (24 * 60)) / 60);
  const minutes = totalMinutes % 60;

  const parts: string[] = [];
  if (days > 0) parts.push(`${days}d`);
  if (hours > 0) parts.push(`${hours}h`);
  if (minutes > 0) parts.push(`${minutes}m`);
  return parts.length > 0 ? parts.join(' ') : '0m';
}

function createEmptyReportRow(): ReportRow {
  return {
    name: '',
    timeStart: '',
    timeEnd: '',
    chartCount: '',
    amount: '',
  };
}

function createEmptyTableRow(): TableRow {
  return {
    position: '',
    name: '',
    sales: '',
    coffeeNew: '',
    coffeeReturn: '',
    coffeeTotal: '',
    coffeeNo: '',
    renderedCoffee: '',
    coffeeYes: '',
    orangeJuiceNew: '',
    orangeJuiceReturn: '',
    orangeJuiceTotal: '',
  };
}

function createEmptySugarRow(): SugarRow {
  return {
    position: '',
    name: '',
    sugar: '',
    sugarGood: '',
    sugarBad: '',
    paper: '',
  };
}

function createEmptyReasonRow(): ReasonRow {
  return {
    reason: '',
    orangeJuice: '',
    paper: '',
    coffee: '',
  };
}

function createEmptyExtraInputRow(): ExtraInputRow {
  return {
    position: '',
    name: '',
    doctorPreventative: '',
    doctorRestorative: '',
    doctorCraProduction: '',
    customer: '',
    icecream: '',
    cake: '',
    donut: '',
    tart: '',
    peach: '',
    peppermint: '',
    pineapple: '',
    rose: '',
    total: '',
  };
}

function createEmptyLocationSummary(): LocationSummary {
  return {
    pineapple: '',
    rose: '',
    total: '',
  };
}

function createEmptyProductionSideMetrics(): ProductionSideMetrics {
  return {
    add: '',
    noShow: '',
    scheduled: '',
    seen: '',
    seenPercent: '',
  };
}

function normalizeExtraInputRows(rows: ExtraInputRow[] | undefined, targetLength: number): ExtraInputRow[] {
  return Array.from({ length: targetLength }, (_, idx) => ({
    ...createEmptyExtraInputRow(),
    ...(rows?.[idx] || {}),
  }));
}

function computeExtraInputRows(rows: ExtraInputRow[]): ExtraInputRow[] {
  return rows.map((row) => {
    const hasBaseInput =
      String(row.customer ?? '').trim() !== '' ||
      String(row.icecream ?? '').trim() !== '' ||
      String(row.cake ?? '').trim() !== '';
    const hasFlavorInput =
      String(row.tart ?? '').trim() !== '' ||
      String(row.peach ?? '').trim() !== '' ||
      String(row.peppermint ?? '').trim() !== '';
    return {
      ...row,
      donut: hasBaseInput ? String(parseNumber(row.customer) - parseNumber(row.icecream) - parseNumber(row.cake)) : '',
      total: hasFlavorInput
        ? String(parseNumber(row.tart) + parseNumber(row.peach) + parseNumber(row.peppermint))
        : '',
    };
  });
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

function computeTableRows(rows: TableRow[]): TableRow[] {
  return rows.map((row) => {
    const coffeeHasInput = String(row.coffeeNew ?? '').trim() !== '' || String(row.coffeeReturn ?? '').trim() !== '';
    const orangeHasInput =
      String(row.orangeJuiceNew ?? '').trim() !== '' || String(row.orangeJuiceReturn ?? '').trim() !== '';
    const coffeeTotal = coffeeHasInput ? String(parseNumber(row.coffeeNew) + parseNumber(row.coffeeReturn)) : '';
    const orangeJuiceTotal = orangeHasInput ? String(parseNumber(row.orangeJuiceNew) + parseNumber(row.orangeJuiceReturn)) : '';
    /** CRA (Billable) = CRA Total − CRA (Not Billable); CRA Total은 New+Return 합. */
    const coffeeYes = coffeeTotal !== '' ? String(parseNumber(coffeeTotal) - parseNumber(row.coffeeNo)) : '';
    return {
      ...row,
      coffeeTotal,
      orangeJuiceTotal,
      coffeeYes,
    };
  });
}

/** Sealant (Billable) = Sealant − Sealant (Redo). */
function computeSugarRow(row: SugarRow): SugarRow {
  const sealantHasBasis = String(row.sugar ?? '').trim() !== '' || String(row.sugarBad ?? '').trim() !== '';
  const sugarGood = sealantHasBasis ? String(parseNumber(row.sugar) - parseNumber(row.sugarBad)) : '';
  return { ...row, sugarGood };
}

function computeSugarRows(rows: SugarRow[]): SugarRow[] {
  return rows.map((row) => computeSugarRow(row));
}

function getSugarTotalsFromRows(rows: SugarRow[]): SugarTotals {
  return {
    sugar: rows.reduce((sum, row) => sum + parseNumber(row.sugar), 0),
    sugarGood: rows.reduce((sum, row) => sum + parseNumber(row.sugarGood), 0),
    sugarBad: rows.reduce((sum, row) => sum + parseNumber(row.sugarBad), 0),
    paper: rows.reduce((sum, row) => sum + parseNumber(row.paper), 0),
  };
}

function getExtraInputTotals(rows: ExtraInputRow[]) {
  return {
    doctorPreventative: rows.reduce((sum, row) => sum + parseNumber(row.doctorPreventative), 0),
    doctorRestorative: rows.reduce((sum, row) => sum + parseNumber(row.doctorRestorative), 0),
    doctorCraProduction: rows.reduce((sum, row) => sum + parseNumber(row.doctorCraProduction), 0),
    customer: rows.reduce((sum, row) => sum + parseNumber(row.customer), 0),
    icecream: rows.reduce((sum, row) => sum + parseNumber(row.icecream), 0),
    cake: rows.reduce((sum, row) => sum + parseNumber(row.cake), 0),
    donut: rows.reduce((sum, row) => sum + parseNumber(row.donut), 0),
    tart: rows.reduce((sum, row) => sum + parseNumber(row.tart), 0),
    peach: rows.reduce((sum, row) => sum + parseNumber(row.peach), 0),
    peppermint: rows.reduce((sum, row) => sum + parseNumber(row.peppermint), 0),
    total: rows.reduce((sum, row) => sum + parseNumber(row.total), 0),
  };
}

function getLegacyLocationSummaryFromExtraRows(rows: ExtraInputRow[] | undefined): LocationSummary {
  const source = rows || [];
  if (source.length === 0) return createEmptyLocationSummary();
  const hasPineappleInput = source.some((row) => String(row.pineapple ?? '').trim() !== '');
  const hasRoseInput = source.some((row) => String(row.rose ?? '').trim() !== '');
  const pineappleTotal = source.reduce((sum, row) => sum + parseNumber(row.pineapple), 0);
  const roseTotal = source.reduce((sum, row) => sum + parseNumber(row.rose), 0);
  return {
    pineapple: hasPineappleInput ? String(pineappleTotal) : '',
    rose: hasRoseInput ? String(roseTotal) : '',
    total: '',
  };
}

function computeCoffeeActualTotals(values: Partial<Record<keyof TableTotals, string>> | undefined) {
  const coffeeNewRaw = String(values?.coffeeNew ?? '');
  const coffeeReturnRaw = String(values?.coffeeReturn ?? '');
  const orangeJuiceNewRaw = String(values?.orangeJuiceNew ?? '');
  const orangeJuiceReturnRaw = String(values?.orangeJuiceReturn ?? '');

  const hasCoffeeInput = coffeeNewRaw.trim() !== '' || coffeeReturnRaw.trim() !== '';
  const hasOrangeInput = orangeJuiceNewRaw.trim() !== '' || orangeJuiceReturnRaw.trim() !== '';

  return {
    coffeeNew: coffeeNewRaw,
    coffeeReturn: coffeeReturnRaw,
    coffeeTotal: hasCoffeeInput ? String(parseNumber(coffeeNewRaw) + parseNumber(coffeeReturnRaw)) : '',
    orangeJuiceNew: orangeJuiceNewRaw,
    orangeJuiceReturn: orangeJuiceReturnRaw,
    orangeJuiceTotal: hasOrangeInput ? String(parseNumber(orangeJuiceNewRaw) + parseNumber(orangeJuiceReturnRaw)) : '',
  } as Partial<Record<keyof TableTotals, string>>;
}

export default function ViewPage() {
  const [docs, setDocs] = useState<FormDoc[]>([]);
  const [selectedId, setSelectedId] = useState<string>('');
  const [dateFilter, setDateFilter] = useState<string[]>([]);
  const [locationFilter, setLocationFilter] = useState<string[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [isEditing, setIsEditing] = useState(false);
  const [isPdfSaving, setIsPdfSaving] = useState(false);
  const [draft, setDraft] = useState<FormDoc | null>(null);
  const [saveMessage, setSaveMessage] = useState('');

  useEffect(() => {
    const load = async () => {
      try {
        const snap = await getDocs(collection(db, 'simple-forms'));
        const loaded = snap.docs
          .map((d) => {
            const data = d.data() as Omit<FormDoc, 'id'>;
            const { notesLines: _nl, notDueLines: _ndl, ...rest } = data;
            return {
              id: d.id,
              ...rest,
              notes: loadMultilineField(data.notes, data.notesLines).slice(0, NOTES_MAX_LENGTH),
              notDue: loadMultilineField(data.notDue, data.notDueLines).slice(0, NOTES_MAX_LENGTH),
            };
          })
          .filter((doc) => String(doc.submittedDateTime ?? '').trim() !== '')
          .sort((a, b) => `${b.date ?? ''}_${b.location ?? ''}`.localeCompare(`${a.date ?? ''}_${a.location ?? ''}`));
        setDocs(loaded);
        if (loaded.length > 0) {
          setSelectedId(loaded[0].id);
        }
      } catch (e: any) {
        setError(e?.message || '조회 중 오류가 발생했습니다.');
      } finally {
        setLoading(false);
      }
    };

    load();
  }, []);

  const filteredDocs = useMemo(() => {
    return docs.filter((doc) => {
      if (String(doc.submittedDateTime ?? '').trim() === '') return false;
      const dateValue = String(doc.date ?? '');
      const locationValue = String(doc.location ?? '');
      const dateMatched = dateFilter.length === 0 || dateFilter.includes(dateValue);
      const locationMatched = locationFilter.length === 0 || locationFilter.includes(locationValue);
      return dateMatched && locationMatched;
    });
  }, [docs, dateFilter, locationFilter]);
  const dateOptions = useMemo(
    () => Array.from(new Set(docs.map((doc) => String(doc.date ?? '').trim()).filter((v) => v !== ''))),
    [docs]
  );
  const locationOptions = useMemo(
    () => Array.from(new Set(docs.map((doc) => String(doc.location ?? '').trim()).filter((v) => v !== ''))),
    [docs]
  );
  const dateFilterLabel = dateFilter.length === 0 ? 'Date (All)' : `Date (${dateFilter.length})`;
  const locationFilterLabel = locationFilter.length === 0 ? 'Office (All)' : `Location (${locationFilter.length})`;
  const toggleDateFilterValue = (value: string) => {
    setDateFilter((prev) => (prev.includes(value) ? prev.filter((item) => item !== value) : [...prev, value]));
  };
  const toggleLocationFilterValue = (value: string) => {
    setLocationFilter((prev) => (prev.includes(value) ? prev.filter((item) => item !== value) : [...prev, value]));
  };

  useEffect(() => {
    if (filteredDocs.length === 0) {
      setSelectedId('');
      return;
    }
    const existsInFiltered = filteredDocs.some((doc) => doc.id === selectedId);
    if (!existsInFiltered) {
      setSelectedId(filteredDocs[0].id);
    }
  }, [filteredDocs, selectedId]);

  const selectedDoc = useMemo(() => filteredDocs.find((d) => d.id === selectedId), [filteredDocs, selectedId]);
  const reportRows = useMemo(() => {
    if (!selectedDoc) return [];
    if (Array.isArray(selectedDoc.reportRows) && selectedDoc.reportRows.length > 0) return selectedDoc.reportRows;
    return [
      {
        name: selectedDoc.name ?? '',
        timeStart: selectedDoc.timeStart ?? '',
        timeEnd: selectedDoc.timeEnd ?? '',
        chartCount: selectedDoc.chartCount ?? '',
        amount: selectedDoc.amount ?? '',
      },
    ];
  }, [selectedDoc]);
  useEffect(() => {
    if (!selectedDoc) {
      setDraft(null);
      return;
    }
    const normalizedTableRows = (selectedDoc.tableRows || []).map((row) => ({ ...createEmptyTableRow(), ...row }));
    const normalizedReportRows =
      Array.isArray(selectedDoc.reportRows) && selectedDoc.reportRows.length > 0
        ? selectedDoc.reportRows.map((row) => ({ ...createEmptyReportRow(), ...row }))
        : [
            {
              name: selectedDoc.name ?? '',
              timeStart: selectedDoc.timeStart ?? '',
              timeEnd: selectedDoc.timeEnd ?? '',
              chartCount: selectedDoc.chartCount ?? '',
              amount: selectedDoc.amount ?? '',
            },
          ];
    const normalizedExtraInputRows = computeExtraInputRows(
      normalizeExtraInputRows(selectedDoc.extraInputRows, (selectedDoc.tableRows || []).length)
    ).map((row, idx) => ({
      ...row,
      position: row.position || normalizedTableRows[idx]?.position || '',
      name: row.name || normalizedTableRows[idx]?.name || '',
    }));
    const normalizedLocationSummary = {
      ...createEmptyLocationSummary(),
      ...(selectedDoc.locationSummary || getLegacyLocationSummaryFromExtraRows(selectedDoc.extraInputRows)),
    };
    const normalizedProductionSideMetrics = {
      ...createEmptyProductionSideMetrics(),
      ...(selectedDoc.productionSideMetrics || {}),
    };
    const { notesLines: _draftNl, notDueLines: _draftNdl, ...selectedWithoutLineArrays } = selectedDoc;
    setDraft({
      ...selectedWithoutLineArrays,
      checkIn: toTimeInputValue(selectedDoc.checkIn),
      checkOut: toTimeInputValue(selectedDoc.checkOut),
      notes: loadMultilineField(selectedDoc.notes, selectedDoc.notesLines).slice(0, NOTES_MAX_LENGTH),
      notDue: loadMultilineField(selectedDoc.notDue, selectedDoc.notDueLines).slice(0, NOTES_MAX_LENGTH),
      reportRows: normalizedReportRows,
      tableRows: normalizedTableRows,
      coffeeActualTotals: selectedDoc.coffeeActualTotals || {},
      sugarRows: (selectedDoc.sugarRows || []).map((row) => ({ ...createEmptySugarRow(), ...row })),
      reasonRows: (selectedDoc.reasonRows || []).map((row) => ({ ...createEmptyReasonRow(), ...row })),
      extraInputRows: normalizedExtraInputRows,
      locationSummary: normalizedLocationSummary,
      productionSideMetrics: normalizedProductionSideMetrics,
    });
    setIsEditing(false);
    setSaveMessage('');
  }, [selectedDoc]);
  useEffect(() => {
    if (selectedDoc?.pdfSaved && isEditing) {
      setIsEditing(false);
      setSaveMessage('PDF 저장 완료 문서는 수정할 수 없습니다.');
    }
  }, [selectedDoc?.pdfSaved, isEditing]);

  const updateDraftField = (field: keyof FormDoc, value: string) => {
    const nextValue =
      field === 'notes' || field === 'notDue' ? value.slice(0, NOTES_MAX_LENGTH) : value;
    setDraft((prev) => {
      if (!prev) return prev;
      if (field === 'checkIn' || field === 'checkOut') {
        const checkIn = field === 'checkIn' ? nextValue : String(prev.checkIn ?? '');
        const checkOut = field === 'checkOut' ? nextValue : String(prev.checkOut ?? '');
        return {
          ...prev,
          [field]: nextValue,
          hoursOpen: computeHoursOpen(checkIn, checkOut),
        };
      }
      return { ...prev, [field]: nextValue };
    });
  };

  const updateCoffeeActualTotalField = (
    field: 'orangeJuiceNew' | 'orangeJuiceReturn',
    value: string
  ) => {
    setDraft((prev) =>
      prev
        ? {
            ...prev,
            coffeeActualTotals: {
              ...(prev.coffeeActualTotals || {}),
              ...(String(value).trim() === ''
                ? { [field]: value }
                : (() => {
                    const tableRows = computeTableRows(
                      (prev.tableRows || []).map((row) => ({
                        ...createEmptyTableRow(),
                        ...row,
                      }))
                    );
                    const orangeJuiceTotal = getTableTotalsFromRows(tableRows).orangeJuiceTotal;
                    const targetField = field === 'orangeJuiceNew' ? 'orangeJuiceReturn' : 'orangeJuiceNew';

                    return {
                      [field]: value,
                      [targetField]: String(orangeJuiceTotal - parseNumber(value)),
                    };
                  })()),
            },
          }
        : prev
    );
  };

  const updateRowField = <T extends ReportRow | TableRow | SugarRow | ReasonRow | ExtraInputRow>(
    section: 'reportRows' | 'tableRows' | 'sugarRows' | 'reasonRows' | 'extraInputRows',
    rowIndex: number,
    field: keyof T,
    value: string
  ) => {
    setDraft((prev) => {
      if (!prev) return prev;
      const rows = [...((prev[section] as T[]) || [])];
      rows[rowIndex] = { ...rows[rowIndex], [field]: value } as T;
      return { ...prev, [section]: rows };
    });
  };

  const updateDoctorPerformanceProductionField = (
    rowIndex: number,
    field: 'doctorPreventative' | 'doctorRestorative' | 'doctorCraProduction',
    value: string
  ) => {
    setDraft((prev) => {
      if (!prev) return prev;
      const tableLen = (prev.tableRows || []).length;
      const normalized = normalizeExtraInputRows(prev.extraInputRows, tableLen);
      const nextRow = { ...normalized[rowIndex], [field]: value };
      const nextNormalized = normalized.map((r, i) => (i === rowIndex ? nextRow : r));
      const extraInputRows = computeExtraInputRows(nextNormalized);

      const table = (prev.tableRows || []).map((r) => ({ ...createEmptyTableRow(), ...r }));
      const prevTr = { ...createEmptyTableRow(), ...table[rowIndex] };
      const newSales = String(getDoctorPerformanceRowProductionValue(extraInputRows[rowIndex], prevTr.sales));
      table[rowIndex] = { ...prevTr, sales: newSales };

      return {
        ...prev,
        extraInputRows,
        tableRows: computeTableRows(table),
      };
    });
  };

  const updateLocationSummaryField = (field: 'pineapple' | 'rose', value: string) => {
    setDraft((prev) => {
      if (!prev) return prev;
      return {
        ...prev,
        locationSummary: {
          ...createEmptyLocationSummary(),
          ...(prev.locationSummary || {}),
          [field]: value,
        },
      };
    });
  };

  const updateProductionSideMetricField = (field: keyof ProductionSideMetrics, value: string) => {
    if (field === 'seenPercent') return;
    setDraft((prev) => {
      if (!prev) return prev;
      return {
        ...prev,
        productionSideMetrics: {
          ...createEmptyProductionSideMetrics(),
          ...(prev.productionSideMetrics || {}),
          [field]: value,
        },
      };
    });
  };

  const handleSaveEdit = async () => {
    if (!draft) return;
    try {
      const tableRows = computeTableRows(
        (draft.tableRows || []).map((row) => ({
          ...createEmptyTableRow(),
          ...row,
        }))
      );
      const sugarRows = computeSugarRows(
        (draft.sugarRows || []).map((row) => ({
          ...createEmptySugarRow(),
          ...row,
        }))
      );
      const reportRowsToSave = (draft.reportRows || []).map((row) => ({
        ...createEmptyReportRow(),
        ...row,
      }));
      const reasonRows = (draft.reasonRows || []).map((row) => ({
        ...createEmptyReasonRow(),
        ...row,
      }));
      const extraInputRows = computeExtraInputRows(normalizeExtraInputRows(draft.extraInputRows, tableRows.length)).map((row, idx) => ({
        ...row,
        position: row.position || tableRows[idx]?.position || '',
        name: row.name || tableRows[idx]?.name || '',
        pineapple: '',
        rose: '',
      }));

      const tableRowsWithSyncedSales = computeTableRows(
        tableRows.map((tr, idx) => ({
          ...tr,
          sales: String(getDoctorPerformanceRowProductionValue(extraInputRows[idx], tr.sales)),
        }))
      );

      const tableTotals = getTableTotalsFromRows(tableRowsWithSyncedSales);
      const sugarTotals = getSugarTotalsFromRows(sugarRows);
      const hasCoffeeYesInput = tableRowsWithSyncedSales.some((row) => String(row.coffeeYes ?? '').trim() !== '');
      const coffeeSales = hasCoffeeYesInput ? String(tableTotals.coffeeYes * 61) : '';
      const draftLocationSummary = {
        ...createEmptyLocationSummary(),
        ...(draft.locationSummary || {}),
      };
      const locationSummary = {
        pineapple: String(draftLocationSummary.pineapple ?? ''),
        rose: String(draftLocationSummary.rose ?? ''),
        total: String(parseNumber(draftLocationSummary.pineapple) + parseNumber(draftLocationSummary.rose) + parseNumber(coffeeSales)),
      };
      const draftSideMetrics = {
        ...createEmptyProductionSideMetrics(),
        ...(draft.productionSideMetrics || {}),
      };
      const seenPercentComputed = computeSeenPercentRounded(draftSideMetrics.scheduled, draftSideMetrics.seen);
      const productionSideMetrics = {
        add: String(draftSideMetrics.add ?? ''),
        noShow: String(draftSideMetrics.noShow ?? ''),
        scheduled: String(draftSideMetrics.scheduled ?? ''),
        seen: String(draftSideMetrics.seen ?? ''),
        seenPercent: seenPercentComputed === '-' ? '' : seenPercentComputed,
      };
      const grandTotal = String(draft.grandTotal ?? '');
      const bothEmpty = grandTotal.trim() === '' && coffeeSales.trim() === '';
      const salesWithoutCoffee = bothEmpty ? '' : String(parseNumber(grandTotal) - parseNumber(coffeeSales));
      const prophyTotal = computeProphyTotal(
        draft.paperAtOrangeJuice,
        draft.paperAtTea,
        draft.justPaper
      );
      const hoursOpen = computeHoursOpen(draft.checkIn, draft.checkOut);
      const firstReport = reportRowsToSave[0] ?? createEmptyReportRow();
      const coffeeActualTotals = computeCoffeeActualTotals(draft.coffeeActualTotals || {});

      const { notesLines: _saveNl, notDueLines: _saveNdl, ...draftWithoutLineArrays } = draft;
      const multilineFields = getMultilineFirestoreFields(
        String(draft.notes ?? ''),
        String(draft.notDue ?? '')
      );
      const payload: Omit<FormDoc, 'id'> & { updatedAt: unknown } = {
        ...draftWithoutLineArrays,
        ...multilineFields,
        edited: true,
        editedAt: serverTimestamp(),
        reportRows: reportRowsToSave,
        tableRows: tableRowsWithSyncedSales,
        tableTotals,
        coffeeActualTotals,
        sugarRows,
        sugarTotals,
        reasonRows,
        extraInputRows,
        locationSummary,
        productionSideMetrics,
        name: firstReport.name ?? '',
        timeStart: firstReport.timeStart ?? '',
        timeEnd: firstReport.timeEnd ?? '',
        chartCount: firstReport.chartCount ?? '',
        amount: firstReport.amount ?? '',
        coffeeSales,
        salesWithoutCoffee,
        prophyTotal,
        checkIn: toStoredCheckTime12h(String(draft.checkIn ?? '')),
        checkOut: toStoredCheckTime12h(String(draft.checkOut ?? '')),
        hoursOpen,
        closer: String(draft.closer ?? ''),
        updatedAt: serverTimestamp(),
      };

      await setDoc(doc(db, 'simple-forms', draft.id), payload, { merge: true });
      setDocs((prev) => prev.map((item) => (item.id === draft.id ? { id: draft.id, ...payload } : item)));
      setIsEditing(false);
      setSaveMessage('수정 내용이 저장되었습니다.');
    } catch (e: any) {
      setSaveMessage(`저장 실패: ${e?.message || '알 수 없는 오류'}`);
    }
  };
  const tableTotals = useMemo(() => {
    if (!selectedDoc) return null;
    if (selectedDoc.tableTotals) {
      return {
        sales: parseNumber(selectedDoc.tableTotals.sales),
        coffeeNew: parseNumber(selectedDoc.tableTotals.coffeeNew),
        coffeeReturn: parseNumber(selectedDoc.tableTotals.coffeeReturn),
        coffeeTotal: parseNumber(selectedDoc.tableTotals.coffeeTotal),
        coffeeNo: parseNumber(selectedDoc.tableTotals.coffeeNo),
        renderedCoffee: parseNumber(selectedDoc.tableTotals.renderedCoffee),
        coffeeYes: parseNumber(selectedDoc.tableTotals.coffeeYes),
        orangeJuiceNew: parseNumber(selectedDoc.tableTotals.orangeJuiceNew),
        orangeJuiceReturn: parseNumber(selectedDoc.tableTotals.orangeJuiceReturn),
        orangeJuiceTotal: parseNumber(selectedDoc.tableTotals.orangeJuiceTotal),
      };
    }
    return getTableTotalsFromRows(selectedDoc.tableRows || []);
  }, [selectedDoc]);
  const sugarTotals = useMemo(() => {
    if (!selectedDoc) return null;
    if (selectedDoc.sugarTotals) {
      return {
        sugar: parseNumber(selectedDoc.sugarTotals.sugar),
        sugarGood: parseNumber(selectedDoc.sugarTotals.sugarGood),
        sugarBad: parseNumber(selectedDoc.sugarTotals.sugarBad),
        paper: parseNumber(selectedDoc.sugarTotals.paper),
      };
    }
    return getSugarTotalsFromRows(selectedDoc.sugarRows || []);
  }, [selectedDoc]);
  const visibleReportRows = isEditing ? draft?.reportRows || [] : reportRows;
  const visibleTableRows = isEditing ? computeTableRows(draft?.tableRows || []) : selectedDoc?.tableRows || [];
  const visibleSugarRows = isEditing
    ? computeSugarRows(
        (draft?.sugarRows || []).map((row) => ({
          ...createEmptySugarRow(),
          ...row,
        }))
      )
    : selectedDoc?.sugarRows || [];
  const visibleReasonRows = isEditing ? draft?.reasonRows || [] : selectedDoc?.reasonRows || [];
  const visibleExtraInputRows = isEditing
    ? computeExtraInputRows(normalizeExtraInputRows(draft?.extraInputRows, visibleTableRows.length))
    : computeExtraInputRows(normalizeExtraInputRows(selectedDoc?.extraInputRows, (selectedDoc?.tableRows || []).length));
  const visibleExtraInputRowsWithIdentity = visibleExtraInputRows.map((row, idx) => ({
    ...row,
    position: row.position || visibleTableRows[idx]?.position || '',
    name: row.name || visibleTableRows[idx]?.name || '',
  }));
  const visibleTableTotals = isEditing ? getTableTotalsFromRows(visibleTableRows) : tableTotals;
  const visibleCoffeeActualTotals = computeCoffeeActualTotals(isEditing ? draft?.coffeeActualTotals || {} : selectedDoc?.coffeeActualTotals || {});
  const visibleSugarTotals = isEditing ? getSugarTotalsFromRows(visibleSugarRows) : sugarTotals;
  const visibleExtraInputTotals = getExtraInputTotals(visibleExtraInputRowsWithIdentity);
  const legacyLocationSummary = getLegacyLocationSummaryFromExtraRows(selectedDoc?.extraInputRows);
  const visibleLocationSummary = {
    ...createEmptyLocationSummary(),
    ...(isEditing ? draft?.locationSummary || {} : selectedDoc?.locationSummary || legacyLocationSummary),
  };
  const hasCoffeeYesInputVisible = visibleTableRows.some((row) => String(row.coffeeYes ?? '').trim() !== '');
  const visibleCoffeeSales = hasCoffeeYesInputVisible
    ? String((visibleTableTotals?.coffeeYes ?? 0) * 61)
    : selectedDoc?.coffeeSales || '';
  const visiblePineappleValue = String(visibleLocationSummary.pineapple ?? '');
  const visibleRoseValue = String(visibleLocationSummary.rose ?? '');
  const visibleLocationTotal =
    visiblePineappleValue.trim() === '' && visibleRoseValue.trim() === '' && visibleCoffeeSales.trim() === ''
      ? ''
      : String(parseNumber(visiblePineappleValue) + parseNumber(visibleRoseValue) + parseNumber(visibleCoffeeSales));
  const visibleGrandTotal = isEditing ? String(draft?.grandTotal ?? '') : selectedDoc?.grandTotal || '';
  const visibleSalesWithoutCoffee =
    visibleGrandTotal.trim() === '' && visibleCoffeeSales.trim() === ''
      ? ''
      : String(parseNumber(visibleGrandTotal) - parseNumber(visibleCoffeeSales));
  const visiblePaperAtOrangeJuice = isEditing ? String(draft?.paperAtOrangeJuice ?? '') : selectedDoc?.paperAtOrangeJuice || '';
  const visiblePaperAtTea = isEditing ? String(draft?.paperAtTea ?? '') : selectedDoc?.paperAtTea || '';
  const visibleJustPaper = isEditing ? String(draft?.justPaper ?? '') : selectedDoc?.justPaper || '';
  const visibleProphyTotal = computeProphyTotal(
    visiblePaperAtOrangeJuice,
    visiblePaperAtTea,
    visibleJustPaper
  );
  const visibleCheckIn = isEditing ? String(draft?.checkIn ?? '') : selectedDoc?.checkIn || '';
  const visibleCheckOut = isEditing ? String(draft?.checkOut ?? '') : selectedDoc?.checkOut || '';
  const visibleCloser = isEditing ? String(draft?.closer ?? '') : selectedDoc?.closer || '';
  const visibleHoursOpen = (() => {
    const checkIn = isEditing ? draft?.checkIn : selectedDoc?.checkIn;
    const checkOut = isEditing ? draft?.checkOut : selectedDoc?.checkOut;
    const computed = computeHoursOpen(checkIn, checkOut);
    if (computed) return computed;
    const stored = isEditing ? draft?.hoursOpen : selectedDoc?.hoursOpen;
    const formatted = formatHoursOpenLabel(stored);
    return formatted === '-' ? '' : formatted;
  })();
  const visibleProductionSideMetrics = {
    ...createEmptyProductionSideMetrics(),
    ...(isEditing ? draft?.productionSideMetrics || {} : selectedDoc?.productionSideMetrics || {}),
  };

  const handleAddReportRow = () => {
    setDraft((prev) => (prev ? { ...prev, reportRows: [...(prev.reportRows || []), createEmptyReportRow()] } : prev));
  };
  const handleAddTableRow = () => {
    setDraft((prev) =>
      prev
        ? {
            ...prev,
            tableRows: [...(prev.tableRows || []), createEmptyTableRow()],
            extraInputRows: [...(prev.extraInputRows || []), createEmptyExtraInputRow()],
          }
        : prev
    );
  };
  const handleAddSugarRow = () => {
    setDraft((prev) => (prev ? { ...prev, sugarRows: [...(prev.sugarRows || []), createEmptySugarRow()] } : prev));
  };
  const handleAddReasonRow = () => {
    setDraft((prev) => (prev ? { ...prev, reasonRows: [...(prev.reasonRows || []), createEmptyReasonRow()] } : prev));
  };
  const handleDownloadPdf = async () => {
    if (!selectedDoc) return;
    try {
      setIsPdfSaving(true);
      setSaveMessage('');
      const generatedDate = new Date().toLocaleDateString('ko-KR');
      const pdfDoc = createSubmittedReportPDFDocument({
        date: selectedDoc.date || '',
        location: selectedDoc.location || '',
        generatedDate,
        reasonIfLate: selectedDoc.reasonIfLate || '',
        checkIn: visibleCheckIn,
        checkOut: visibleCheckOut,
        hoursOpen: visibleHoursOpen,
        closer: visibleCloser,
        submittedAt: selectedDoc.submittedDateTime || '',
        grandTotal: visibleGrandTotal,
        coffeeSales: visibleCoffeeSales,
        salesWithoutCoffee: visibleSalesWithoutCoffee,
        paperAtOrangeJuice: visiblePaperAtOrangeJuice,
        paperAtTea: visiblePaperAtTea,
        justPaper: visibleJustPaper,
        prophyTotal: visibleProphyTotal,
        reportRows: visibleReportRows,
        tableRows: visibleTableRows,
        tableTotals: visibleTableTotals,
        actualOrangeNew: String(visibleCoffeeActualTotals.orangeJuiceNew ?? ''),
        actualOrangeReturn: String(visibleCoffeeActualTotals.orangeJuiceReturn ?? ''),
        extraInputRows: visibleExtraInputRowsWithIdentity,
        extraInputTotals: visibleExtraInputTotals,
        locationSummary: {
          pineapple: visiblePineappleValue,
          rose: visibleRoseValue,
          coffeeSales: visibleCoffeeSales,
          total: visibleLocationTotal,
        },
        productionSideMetrics: visibleProductionSideMetrics,
        sugarRows: visibleSugarRows,
        sugarTotals: visibleSugarTotals,
        reasonRows: visibleReasonRows,
        notes: isEditing ? String(draft?.notes ?? '') : selectedDoc.notes || '',
        notDue: isEditing ? String(draft?.notDue ?? '') : selectedDoc.notDue || '',
      });
      const blob = await pdf(pdfDoc).toBlob();
      const filename = sanitizeFilename(`${selectedDoc.date || 'no-date'}_${selectedDoc.location || 'no-location'}_daily production.pdf`);
      const downloadUrl = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = filename;
      link.click();
      URL.revokeObjectURL(downloadUrl);

      // PDF 생성 완료 시점에 월/지점별 누적 집계도 함께 저장
      const monthKey = getMonthKey(selectedDoc.date);
      const locationKey = toFirestoreKey(selectedDoc.location);
      const monthlyDocId = `${monthKey}-${locationKey}`;
      const monthlyRef = doc(db, 'simple-forms-monthly', monthlyDocId);
      const grandTotalValue = parseNumber(visibleGrandTotal);
      const coffeeSalesValue = parseNumber(visibleCoffeeSales);
      const salesWithoutCoffeeValue = parseNumber(visibleSalesWithoutCoffee);
      const paperAtOrangeJuiceValue = parseNumber(visiblePaperAtOrangeJuice);
      const paperAtTeaValue = parseNumber(visiblePaperAtTea);
      const justPaperValue = parseNumber(visibleJustPaper);
      const prophyTotalValue = parseNumber(visibleProphyTotal);
      const actualOrangeNewValue = parseNumber(visibleCoffeeActualTotals.orangeJuiceNew);
      const actualOrangeReturnValue = parseNumber(visibleCoffeeActualTotals.orangeJuiceReturn);
      const pineappleValue = parseNumber(visiblePineappleValue);
      const roseValue = parseNumber(visibleRoseValue);

      await setDoc(
        monthlyRef,
        {
          docId: monthlyDocId,
          month: monthKey,
          location: String(selectedDoc.location ?? '').trim() || 'Unknown',
          locationKey,
          updatedAt: serverTimestamp(),
          docCount: increment(1),
          grandTotal: increment(grandTotalValue),
          coffeeSales: increment(coffeeSalesValue),
          salesWithoutCoffee: increment(salesWithoutCoffeeValue),
          paperAtOrangeJuice: increment(paperAtOrangeJuiceValue),
          paperAtTea: increment(paperAtTeaValue),
          justPaper: increment(justPaperValue),
          prophyTotal: increment(prophyTotalValue),
          actualOrangeJuiceNew: increment(actualOrangeNewValue),
          actualOrangeJuiceReturn: increment(actualOrangeReturnValue),
          pineapple: increment(pineappleValue),
          rose: increment(roseValue),
        },
        { merge: true }
      );

      await setDoc(
        doc(db, 'simple-forms', selectedDoc.id),
        {
          pdfSaved: true,
          pdfSavedAt: serverTimestamp(),
        },
        { merge: true }
      );
      setDocs((prev) =>
        prev.map((item) =>
          item.id === selectedDoc.id
            ? {
                ...item,
                pdfSaved: true,
                pdfSavedAt: new Date().toISOString(),
              }
            : item
        )
      );
      setSaveMessage('PDF가 다운로드되었습니다.');
    } catch (e: any) {
      setSaveMessage(`PDF 생성 실패: ${e?.message || '알 수 없는 오류'}`);
    } finally {
      setIsPdfSaving(false);
    }
  };

  return (
    <main className="d-page-main" style={{ minHeight: '100vh', background: '#fff', padding: 24 }}>
      <style>{D_PAGE_NUMBER_INPUT_SPINNER_RESET_CSS}</style>
      <section
        style={{
          maxWidth: 2600,
          margin: '0 auto',
          padding: 20,
        }}
      >
        <div style={{ marginBottom: 14, display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8 }}>
          <h2 style={{ margin: '0 0 8px', fontSize: 28, fontWeight: 800, color: '#111827' }}>Submitted Finalized Production</h2>
          {selectedDoc && (
            <div style={{ display: 'flex', gap: 8 }}>
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
                    onClick={() => {
                      setIsEditing(false);
                      setSaveMessage('');
                    }}
                    style={{
                      height: 36,
                      padding: '0 14px',
                      borderRadius: 8,
                      border: '1px solid #cbd5e1',
                      background: '#fff',
                      color: '#111827',
                      fontWeight: 600,
                      cursor: 'pointer',
                    }}
                  >
                    Cancel
                  </button>
                </>
              ) : (
                <>
                  {!selectedDoc.pdfSaved && (
                    <button
                      type="button"
                      onClick={() => {
                        if (selectedDoc.pdfSaved) return;
                        setIsEditing(true);
                        setSaveMessage('');
                      }}
                      style={{
                        height: 36,
                        padding: '0 14px',
                        borderRadius: 8,
                        border: '1px solid #2563eb',
                        background: '#2563eb',
                        color: '#fff',
                        fontWeight: 700,
                        cursor: 'pointer',
                      }}
                    >
                      Edit
                    </button>
                  )}
                  {!selectedDoc.pdfSaved && (
                    <button
                      type="button"
                      onClick={handleDownloadPdf}
                      disabled={isPdfSaving}
                      style={{
                        height: 36,
                        padding: '0 14px',
                        borderRadius: 8,
                        border: '1px solid #7c3aed',
                        background: isPdfSaving ? '#a78bfa' : '#7c3aed',
                        color: '#fff',
                        fontWeight: 700,
                        cursor: isPdfSaving ? 'not-allowed' : 'pointer',
                      }}
                    >
                      {isPdfSaving ? 'Generating...' : 'Generate PDF'}
                    </button>
                  )}
                </>
              )}
            </div>
          )}
        </div>
        {saveMessage && <p style={{ margin: '0 0 10px', color: saveMessage.startsWith('저장 실패') ? '#b91c1c' : '#166534' }}>{saveMessage}</p>}

        {loading && <p style={{ margin: 0, color: '#6b7280' }}>Loading...</p>}
        {error && <p style={{ margin: 0, color: '#b91c1c' }}>{error}</p>}

        {!loading && !error && (
          <div style={{ display: 'grid', gridTemplateColumns: '380px 1fr', gap: 16 }}>
            <div style={{ border: '1px solid #e5e7eb', borderRadius: 10, overflow: 'hidden' }}>
              <div style={{ padding: 10, background: '#f9fafb', borderBottom: '1px solid #e5e7eb', fontWeight: 700 }}>
                Date / Office
              </div>
              <div style={{ padding: 10, borderBottom: '1px solid #e5e7eb', display: 'grid', gap: 8 }}>
                <details style={{ width: '72%', border: '1px solid #d1d5db', borderRadius: 6, background: '#fff' }}>
                  <summary style={{ cursor: 'pointer', listStyle: 'none', padding: '8px 10px', fontWeight: 600 }}>{dateFilterLabel}</summary>
                  <div style={{ borderTop: '1px solid #e5e7eb', padding: '8px 10px', display: 'grid', gap: 6, maxHeight: 180, overflowY: 'auto' }}>
                    <label style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer' }}>
                      <input type="checkbox" checked={dateFilter.length === 0} onChange={() => setDateFilter([])} />
                      All
                    </label>
                    {dateOptions.map((date) => (
                      <label key={date} style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer' }}>
                        <input type="checkbox" checked={dateFilter.includes(date)} onChange={() => toggleDateFilterValue(date)} />
                        {date}
                      </label>
                    ))}
                  </div>
                </details>
                <details style={{ width: '72%', border: '1px solid #d1d5db', borderRadius: 6, background: '#fff' }}>
                  <summary style={{ cursor: 'pointer', listStyle: 'none', padding: '8px 10px', fontWeight: 600 }}>{locationFilterLabel}</summary>
                  <div style={{ borderTop: '1px solid #e5e7eb', padding: '8px 10px', display: 'grid', gap: 6, maxHeight: 180, overflowY: 'auto' }}>
                    <label style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer' }}>
                      <input type="checkbox" checked={locationFilter.length === 0} onChange={() => setLocationFilter([])} />
                      All
                    </label>
                    {locationOptions.map((location) => (
                      <label key={location} style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer' }}>
                        <input
                          type="checkbox"
                          checked={locationFilter.includes(location)}
                          onChange={() => toggleLocationFilterValue(location)}
                        />
                        {location}
                      </label>
                    ))}
                  </div>
                </details>
              </div>
              {filteredDocs.length === 0 ? (
                <div style={{ padding: 12, color: '#6b7280' }}>저장된 데이터가 없습니다.</div>
              ) : (
                filteredDocs.map((d, docIdx) => {
                  const active = d.id === selectedId;
                  const statusLabel = d.pdfSaved ? 'Completed' : d.edited ? 'Editing' : '';
                  const statusStyle = d.pdfSaved
                    ? { background: '#dcfce7', color: '#166534', border: '1px solid #86efac' }
                    : { background: '#dbeafe', color: '#1d4ed8', border: '1px solid #93c5fd' };
                  const isLastDocRow = docIdx === filteredDocs.length - 1;
                  return (
                    <button
                      key={d.id}
                      type="button"
                      onClick={() => setSelectedId(d.id)}
                      style={{
                        width: '100%',
                        textAlign: 'left',
                        padding: '10px 12px',
                        border: 0,
                        borderBottom: isLastDocRow ? 'none' : '1px solid #e5e7eb',
                        background: active ? '#eef2ff' : '#fff',
                        cursor: 'pointer',
                      }}
                    >
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
                        <div style={{ fontWeight: 700 }}>{d.date || '-'}</div>
                        {statusLabel && (
                          <span
                            style={{
                              fontSize: 11,
                              fontWeight: 700,
                              borderRadius: 999,
                              padding: '2px 8px',
                              ...statusStyle,
                            }}
                          >
                            {statusLabel}
                          </span>
                        )}
                      </div>
                      <div style={{ color: '#475569', fontSize: 13 }}>{d.location || '-'}</div>
                    </button>
                  );
                })
              )}
            </div>

            <div
              style={{
                border: '1px solid #e5e7eb',
                borderRadius: 10,
                padding: D_PAGE_DETAIL_CARD_PADDING_PX,
                boxSizing: 'border-box',
              }}
            >
              {!selectedDoc ? (
                <p style={{ margin: 0, color: '#6b7280' }}>왼쪽에서 항목을 선택해 주세요.</p>
              ) : (
                <>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, minmax(220px, 1fr))', gap: 10, marginBottom: 14 }}>
                    <div>
                      <strong>Date:</strong> {selectedDoc.date || '-'}
                    </div>
                    <div>
                      <strong>Office:</strong> {selectedDoc.location || '-'}
                    </div>
                    <div>
                      <strong>Reason if Late:</strong>{' '}
                      {isEditing ? (
                        <input
                          type="text"
                          value={String(draft?.reasonIfLate ?? '')}
                          onChange={(e) => updateDraftField('reasonIfLate', e.target.value)}
                          style={{ width: '100%', maxWidth: 420, height: 30, marginLeft: 6, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                        />
                      ) : (
                        selectedDoc.reasonIfLate || '-'
                      )}
                    </div>
                    <div>
                      <strong>Submitted at:</strong> {selectedDoc.submittedDateTime || '-'}
                    </div>
                  </div>

                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, minmax(200px, 1fr))', gap: 10, marginBottom: 14 }}>
                    <div>
                      <strong>Check In:</strong>{' '}
                      {isEditing ? (
                        <input
                          type="time"
                          value={visibleCheckIn}
                          onChange={(e) => updateDraftField('checkIn', e.target.value)}
                          style={{ height: 30, marginLeft: 6, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                        />
                      ) : (
                        visibleCheckIn || '-'
                      )}
                    </div>
                    <div>
                      <strong>Check Out:</strong>{' '}
                      {isEditing ? (
                        <input
                          type="time"
                          value={visibleCheckOut}
                          onChange={(e) => updateDraftField('checkOut', e.target.value)}
                          style={{ height: 30, marginLeft: 6, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                        />
                      ) : (
                        visibleCheckOut || '-'
                      )}
                    </div>
                    <div>
                      <strong>Hours Open:</strong>{' '}
                      <span style={{ marginLeft: 6, color: '#374151' }}>
                        {formatHoursOpenLabel(visibleHoursOpen)}
                      </span>
                    </div>
                    <div>
                      <strong>Closer:</strong>{' '}
                      {isEditing ? (
                        <input
                          type="text"
                          value={visibleCloser}
                          onChange={(e) => updateDraftField('closer', e.target.value)}
                          style={{ width: '100%', maxWidth: 220, height: 30, marginLeft: 6, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                        />
                      ) : (
                        visibleCloser || '-'
                      )}
                    </div>
                  </div>

                  <h3 style={{ margin: '10px 0' }}>Billers</h3>
                  <div style={dPageDetailTableBleedScroll}>
                    <table style={{ width: '100%', minWidth: 920, borderCollapse: 'collapse', fontSize: 14 }}>
                      <thead>
                        <tr style={{ background: '#f3f4f6' }}>
                          {['Name', 'Date/Time Start', 'Date/Time End', 'Duration', '# Chart', '$ Amount'].map((h) => (
                            <th key={h} style={{ border: '1px solid #e5e7eb', padding: 8, textAlign: 'left' }}>
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {visibleReportRows.map((r, idx) => (
                          <tr key={`report-${idx}`}>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={r.name || ''}
                                  onChange={(e) => updateRowField<ReportRow>('reportRows', idx, 'name', e.target.value)}
                                  style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                                />
                              ) : (
                                r.name || '-'
                              )}
                            </td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                              {isEditing ? (
                                <input
                                  type="datetime-local"
                                  value={r.timeStart || ''}
                                  onChange={(e) => updateRowField<ReportRow>('reportRows', idx, 'timeStart', e.target.value)}
                                  style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                                />
                              ) : (
                                r.timeStart || '-'
                              )}
                            </td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                              {isEditing ? (
                                <input
                                  type="datetime-local"
                                  value={r.timeEnd || ''}
                                  onChange={(e) => updateRowField<ReportRow>('reportRows', idx, 'timeEnd', e.target.value)}
                                  style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                                />
                              ) : (
                                r.timeEnd || '-'
                              )}
                            </td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8, color: '#374151' }}>
                              {getDurationLabel(r.timeStart, r.timeEnd)}
                            </td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                              {isEditing ? (
                                <input
                                  type="number"
                                  value={r.chartCount || ''}
                                  onChange={(e) => updateRowField<ReportRow>('reportRows', idx, 'chartCount', e.target.value)}
                                  style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                                />
                              ) : (
                                r.chartCount || '-'
                              )}
                            </td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                              {isEditing ? (
                                <input
                                  type="number"
                                  value={r.amount || ''}
                                  onChange={(e) => updateRowField<ReportRow>('reportRows', idx, 'amount', e.target.value)}
                                  style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                                />
                              ) : (
                                formatCurrencyLabel(r.amount || '-')
                              )}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  {isEditing && (
                    <div style={{ marginTop: 8 }}>
                      <button
                        type="button"
                        onClick={handleAddReportRow}
                        style={{
                          height: 34,
                          padding: '0 12px',
                          borderRadius: 8,
                          border: '1px solid #cbd5e1',
                          background: '#fff',
                          color: '#111827',
                          fontWeight: 600,
                          cursor: 'pointer',
                        }}
                      >
                        Add Row
                      </button>
                    </div>
                  )}

                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, minmax(220px, 1fr))', gap: 10, marginTop: 14 }}>
                    <div>
                      <strong>Grand Total:</strong>{' '}
                      {isEditing ? (
                        <input
                          type="number"
                          value={String(draft?.grandTotal ?? '')}
                          onChange={(e) => updateDraftField('grandTotal', e.target.value)}
                          style={{ width: 140, height: 30, marginLeft: 6, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                        />
                      ) : (
                        formatCurrencyLabel(selectedDoc.grandTotal || '-')
                      )}
                    </div>
                    <div>
                      <strong>CRA Production:</strong> {formatCurrencyLabel(visibleCoffeeSales || '-')}
                    </div>
                    <div>
                      <strong>Production W/Out CRA:</strong> {formatCurrencyLabel(visibleSalesWithoutCoffee || '-')}
                    </div>
                  </div>

                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, minmax(200px, 1fr))', gap: 10, marginTop: 10 }}>
                    <div>
                      <strong>Prophy @ OE:</strong>{' '}
                      {isEditing ? (
                        <input
                          type="number"
                          value={String(draft?.paperAtOrangeJuice ?? '')}
                          onChange={(e) => updateDraftField('paperAtOrangeJuice', e.target.value)}
                          style={{ width: 120, height: 30, marginLeft: 6, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                        />
                      ) : (
                        selectedDoc.paperAtOrangeJuice || '-'
                      )}
                    </div>
                    <div>
                      <strong>Prophy @ TX:</strong>{' '}
                      {isEditing ? (
                        <input
                          type="number"
                          value={String(draft?.paperAtTea ?? '')}
                          onChange={(e) => updateDraftField('paperAtTea', e.target.value)}
                          style={{ width: 120, height: 30, marginLeft: 6, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                        />
                      ) : (
                        selectedDoc.paperAtTea || '-'
                      )}
                    </div>
                    <div>
                      <strong>Just Prophy:</strong>{' '}
                      {isEditing ? (
                        <input
                          type="number"
                          value={String(draft?.justPaper ?? '')}
                          onChange={(e) => updateDraftField('justPaper', e.target.value)}
                          style={{ width: 120, height: 30, marginLeft: 6, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                        />
                      ) : (
                        selectedDoc.justPaper || '-'
                      )}
                    </div>
                    <div>
                      <strong>Prophy Total:</strong>{' '}
                      <span style={{ marginLeft: 6, color: '#374151' }}>{visibleProphyTotal || '-'}</span>
                    </div>
                  </div>

                  <h3 style={{ margin: '16px 0 10px' }}>Production</h3>
                  <div
                    style={{
                      ...dPageDetailTableBleedScroll,
                      display: 'flex',
                      flexWrap: 'wrap',
                      gap: 16,
                      alignItems: 'flex-start',
                    }}
                  >
                    <div>
                      <table style={{ width: '100%', maxWidth: 480, borderCollapse: 'collapse', fontSize: 14 }}>
                        <tbody>
                          <tr>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8, background: '#f9fafb', fontWeight: 700 }}>Preventative</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                              {isEditing ? (
                                <input
                                  type="number"
                                  value={String(draft?.locationSummary?.pineapple ?? '')}
                                  onChange={(e) => updateLocationSummaryField('pineapple', e.target.value)}
                                  style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                                />
                              ) : (
                                formatCurrencyLabel(visiblePineappleValue || '-')
                              )}
                            </td>
                          </tr>
                          <tr>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8, background: '#f9fafb', fontWeight: 700 }}>Restorative</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                              {isEditing ? (
                                <input
                                  type="number"
                                  value={String(draft?.locationSummary?.rose ?? '')}
                                  onChange={(e) => updateLocationSummaryField('rose', e.target.value)}
                                  style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                                />
                              ) : (
                                formatCurrencyLabel(visibleRoseValue || '-')
                              )}
                            </td>
                          </tr>
                          <tr>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8, background: '#f9fafb', fontWeight: 700 }}>CRA Production</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{formatCurrencyLabel(visibleCoffeeSales || '-')}</td>
                          </tr>
                          <tr>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8, background: '#f3f4f6', fontWeight: 800 }}>Total</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8, fontWeight: 800 }}>{formatCurrencyLabel(visibleLocationTotal || '-')}</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                    <div>
                      <table style={{ width: '100%', minWidth: 220, maxWidth: 320, borderCollapse: 'collapse', fontSize: 14 }}>
                        <tbody>
                          {(() => {
                            const draftSideMetricsMerged = {
                              ...createEmptyProductionSideMetrics(),
                              ...(draft?.productionSideMetrics || {}),
                            };
                            return PRODUCTION_SIDE_METRIC_ROWS.map(({ key, label }) => (
                              <tr key={key}>
                                <td style={{ border: '1px solid #e5e7eb', padding: 8, background: '#f9fafb', fontWeight: 700, whiteSpace: 'nowrap' }}>
                                  {label}
                                </td>
                                <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                                  {key === 'seenPercent' ? (
                                    isEditing ? (
                                      <input
                                        type="text"
                                        readOnly
                                        value={seenPercentReadOnlyInputValueFromMetrics(draftSideMetricsMerged)}
                                        style={{
                                          width: '100%',
                                          minWidth: 100,
                                          height: 30,
                                          border: '1px solid #d1d5db',
                                          borderRadius: 6,
                                          padding: '0 8px',
                                          background: '#f3f4f6',
                                          color: '#6b7280',
                                        }}
                                      />
                                    ) : (
                                      formatSeenPercentDisplay(
                                        computeSeenPercentRounded(visibleProductionSideMetrics.scheduled, visibleProductionSideMetrics.seen)
                                      )
                                    )
                                  ) : isEditing ? (
                                    <input
                                      type="text"
                                      value={String(draft?.productionSideMetrics?.[key] ?? '')}
                                      onChange={(e) => updateProductionSideMetricField(key, e.target.value)}
                                      style={{ width: '100%', minWidth: 100, height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                                    />
                                  ) : (
                                    String(visibleProductionSideMetrics[key] ?? '').trim() || '-'
                                  )}
                                </td>
                              </tr>
                            ));
                          })()}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  <>
                    <h3 style={{ margin: '16px 0 10px' }}>Doctors</h3>
                      <div style={dPageDetailTableBleedScroll}>
                        <table style={{ width: '100%', minWidth: 1400, borderCollapse: 'collapse', fontSize: 14 }}>
                          <thead>
                            <tr style={{ background: '#f3f4f6' }}>
                              {[
                                'Position',
                                'Name',
                                'Preventative',
                                'Restorative',
                                'CRA Production',
                                'Production',
                                'Patient Seen',
                                'Insurance',
                                'Cash',
                                'Dentical',
                                'Treatment',
                                'Primary Teeth',
                                'Permanent Teeth',
                              ].map((h) => (
                                <th key={h} style={{ border: '1px solid #e5e7eb', padding: 8, textAlign: 'left' }}>
                                  {h}
                                </th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {visibleExtraInputRowsWithIdentity.map((row, idx) => (
                              <tr key={`extra-input-${idx}`}>
                                <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                                  {isEditing ? (
                                    <input
                                      type="text"
                                      value={row.position || ''}
                                      onChange={(e) => updateRowField<ExtraInputRow>('extraInputRows', idx, 'position', e.target.value)}
                                      style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                                    />
                                  ) : (
                                    row.position || '-'
                                  )}
                                </td>
                                <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                                  {isEditing ? (
                                    <input
                                      type="text"
                                      value={row.name || ''}
                                      onChange={(e) => updateRowField<ExtraInputRow>('extraInputRows', idx, 'name', e.target.value)}
                                      style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                                    />
                                  ) : (
                                    row.name || '-'
                                  )}
                                </td>
                                <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                                  {isEditing ? (
                                    <input
                                      type="number"
                                      value={row.doctorPreventative || ''}
                                      onChange={(e) => updateDoctorPerformanceProductionField(idx, 'doctorPreventative', e.target.value)}
                                      style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                                    />
                                  ) : (
                                    formatCurrencyLabel(row.doctorPreventative || '-')
                                  )}
                                </td>
                                <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                                  {isEditing ? (
                                    <input
                                      type="number"
                                      value={row.doctorRestorative || ''}
                                      onChange={(e) => updateDoctorPerformanceProductionField(idx, 'doctorRestorative', e.target.value)}
                                      style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                                    />
                                  ) : (
                                    formatCurrencyLabel(row.doctorRestorative || '-')
                                  )}
                                </td>
                                <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                                  {isEditing ? (
                                    <input
                                      type="number"
                                      value={row.doctorCraProduction || ''}
                                      onChange={(e) => updateDoctorPerformanceProductionField(idx, 'doctorCraProduction', e.target.value)}
                                      style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                                    />
                                  ) : (
                                    formatCurrencyLabel(row.doctorCraProduction || '-')
                                  )}
                                </td>
                                <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                                  {isEditing ? (
                                    <input
                                      type="text"
                                      readOnly
                                      value={(() => {
                                        const v = doctorPerformanceProductionReadOnlyInputValue(row, visibleTableRows[idx]?.sales);
                                        return v === '' ? '' : formatCurrencyLabel(v);
                                      })()}
                                      style={{
                                        width: '100%',
                                        height: 30,
                                        border: '1px solid #d1d5db',
                                        borderRadius: 6,
                                        padding: '0 8px',
                                        background: '#f3f4f6',
                                        color: '#6b7280',
                                      }}
                                    />
                                  ) : (
                                    formatCurrencyLabel(formatDoctorPerformanceProductionCell(row, visibleTableRows[idx]?.sales))
                                  )}
                                </td>
                                <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                                  {isEditing ? (
                                    <input
                                      type="number"
                                      value={row.customer || ''}
                                      onChange={(e) => updateRowField<ExtraInputRow>('extraInputRows', idx, 'customer', e.target.value)}
                                      style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                                    />
                                  ) : (
                                    row.customer || '-'
                                  )}
                                </td>
                                <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                                  {isEditing ? (
                                    <input
                                      type="number"
                                      value={row.icecream || ''}
                                      onChange={(e) => updateRowField<ExtraInputRow>('extraInputRows', idx, 'icecream', e.target.value)}
                                      style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                                    />
                                  ) : (
                                    row.icecream || '-'
                                  )}
                                </td>
                                <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                                  {isEditing ? (
                                    <input
                                      type="number"
                                      value={row.cake || ''}
                                      onChange={(e) => updateRowField<ExtraInputRow>('extraInputRows', idx, 'cake', e.target.value)}
                                      style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                                    />
                                  ) : (
                                    row.cake || '-'
                                  )}
                                </td>
                                <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                                  {isEditing ? (
                                    <input
                                      type="number"
                                      value={row.donut || ''}
                                      readOnly
                                      style={{
                                        width: '100%',
                                        height: 30,
                                        border: '1px solid #d1d5db',
                                        borderRadius: 6,
                                        padding: '0 8px',
                                        background: '#f3f4f6',
                                        color: '#6b7280',
                                      }}
                                    />
                                  ) : (
                                    row.donut || '-'
                                  )}
                                </td>
                                <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                                  {isEditing ? (
                                    <input
                                      type="number"
                                      value={row.tart || ''}
                                      onChange={(e) => updateRowField<ExtraInputRow>('extraInputRows', idx, 'tart', e.target.value)}
                                      style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                                    />
                                  ) : (
                                    row.tart || '-'
                                  )}
                                </td>
                                <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                                  {isEditing ? (
                                    <input
                                      type="number"
                                      value={row.peach || ''}
                                      onChange={(e) => updateRowField<ExtraInputRow>('extraInputRows', idx, 'peach', e.target.value)}
                                      style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                                    />
                                  ) : (
                                    row.peach || '-'
                                  )}
                                </td>
                                <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                                  {isEditing ? (
                                    <input
                                      type="number"
                                      value={row.peppermint || ''}
                                      onChange={(e) => updateRowField<ExtraInputRow>('extraInputRows', idx, 'peppermint', e.target.value)}
                                      style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }}
                                    />
                                  ) : (
                                    row.peppermint || '-'
                                  )}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                          <tfoot>
                            <tr style={{ background: '#f9fafb', fontWeight: 700 }}>
                              <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>Total</td>
                              <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>-</td>
                              <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                                {formatCurrencyLabel(String(visibleExtraInputTotals.doctorPreventative))}
                              </td>
                              <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                                {formatCurrencyLabel(String(visibleExtraInputTotals.doctorRestorative))}
                              </td>
                              <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                                {formatCurrencyLabel(String(visibleExtraInputTotals.doctorCraProduction))}
                              </td>
                              <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                                {formatCurrencyLabel(
                                  String(
                                    roundToCents(
                                      visibleExtraInputRowsWithIdentity.reduce(
                                        (acc, r, i) => acc + getDoctorPerformanceRowProductionValue(r, visibleTableRows[i]?.sales),
                                        0
                                      )
                                    )
                                  )
                                )}
                              </td>
                              <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleExtraInputTotals.customer}</td>
                              <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleExtraInputTotals.icecream}</td>
                              <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleExtraInputTotals.cake}</td>
                              <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleExtraInputTotals.donut}</td>
                              <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleExtraInputTotals.tart}</td>
                              <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleExtraInputTotals.peach}</td>
                              <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleExtraInputTotals.peppermint}</td>
                            </tr>
                          </tfoot>
                        </table>
                      </div>
                    </>
                  <h3 style={{ margin: '16px 0 10px' }}>CRA / OE</h3>
                  <div style={dPageDetailTableBleedScroll}>
                    <table style={{ width: '100%', minWidth: 1500, borderCollapse: 'collapse', fontSize: 14 }}>
                      <thead>
                        <tr style={{ background: '#f3f4f6' }}>
                          {TABLE_HEADERS.map((h) => (
                            <th key={h} style={{ border: '1px solid #e5e7eb', padding: 8, textAlign: 'left' }}>
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {visibleTableRows.map((r, idx) => (
                          <tr key={`table-${idx}`}>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                              {isEditing ? (
                                <input type="text" value={r.position || ''} onChange={(e) => updateRowField<TableRow>('tableRows', idx, 'position', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} />
                              ) : (
                                r.position || '-'
                              )}
                            </td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                              {isEditing ? (
                                <input type="text" value={r.name || ''} onChange={(e) => updateRowField<TableRow>('tableRows', idx, 'name', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} />
                              ) : (
                                r.name || '-'
                              )}
                            </td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="number" value={r.coffeeNew || ''} onChange={(e) => updateRowField<TableRow>('tableRows', idx, 'coffeeNew', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.coffeeNew || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="number" value={r.coffeeReturn || ''} onChange={(e) => updateRowField<TableRow>('tableRows', idx, 'coffeeReturn', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.coffeeReturn || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{r.coffeeTotal || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="number" value={r.coffeeNo || ''} onChange={(e) => updateRowField<TableRow>('tableRows', idx, 'coffeeNo', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.coffeeNo || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="number" value={r.renderedCoffee || ''} onChange={(e) => updateRowField<TableRow>('tableRows', idx, 'renderedCoffee', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.renderedCoffee || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8, color: isEditing ? '#6b7280' : undefined }}>{r.coffeeYes || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="number" value={r.orangeJuiceNew || ''} onChange={(e) => updateRowField<TableRow>('tableRows', idx, 'orangeJuiceNew', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.orangeJuiceNew || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="number" value={r.orangeJuiceReturn || ''} onChange={(e) => updateRowField<TableRow>('tableRows', idx, 'orangeJuiceReturn', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.orangeJuiceReturn || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{r.orangeJuiceTotal || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8, color: '#6b7280' }}>-</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8, color: '#6b7280' }}>-</td>
                          </tr>
                        ))}
                      </tbody>
                      <tfoot>
                        <tr style={{ background: '#f9fafb', fontWeight: 700 }}>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>Total</td>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>-</td>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleTableTotals ? visibleTableTotals.coffeeNew : 0}</td>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleTableTotals ? visibleTableTotals.coffeeReturn : 0}</td>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleTableTotals ? visibleTableTotals.coffeeTotal : 0}</td>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleTableTotals ? visibleTableTotals.coffeeNo : 0}</td>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleTableTotals ? visibleTableTotals.renderedCoffee : 0}</td>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleTableTotals ? visibleTableTotals.coffeeYes : 0}</td>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleTableTotals ? visibleTableTotals.orangeJuiceNew : 0}</td>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleTableTotals ? visibleTableTotals.orangeJuiceReturn : 0}</td>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                            {visibleTableTotals ? visibleTableTotals.orangeJuiceTotal : 0}
                          </td>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                            {isEditing ? (
                              <input type="number" value={String(visibleCoffeeActualTotals.orangeJuiceNew ?? '')} onChange={(e) => updateCoffeeActualTotalField('orangeJuiceNew', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} />
                            ) : (
                              visibleCoffeeActualTotals.orangeJuiceNew || '-'
                            )}
                          </td>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>
                            {isEditing ? (
                              <input type="number" value={String(visibleCoffeeActualTotals.orangeJuiceReturn ?? '')} onChange={(e) => updateCoffeeActualTotalField('orangeJuiceReturn', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} />
                            ) : (
                              visibleCoffeeActualTotals.orangeJuiceReturn || '-'
                            )}
                          </td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>
                  {isEditing && (
                    <div style={{ marginTop: 8 }}>
                      <button
                        type="button"
                        onClick={handleAddTableRow}
                        style={{
                          height: 34,
                          padding: '0 12px',
                          borderRadius: 8,
                          border: '1px solid #cbd5e1',
                          background: '#fff',
                          color: '#111827',
                          fontWeight: 600,
                          cursor: 'pointer',
                        }}
                      >
                        Add Row
                      </button>
                    </div>
                  )}
                  <h3 style={{ margin: '16px 0 10px' }}>Sealant / Prophy</h3>
                  <div style={dPageDetailTableBleedScroll}>
                    <table style={{ width: '100%', minWidth: 760, borderCollapse: 'collapse', fontSize: 14 }}>
                      <thead>
                        <tr style={{ background: '#f3f4f6' }}>
                          {SUGAR_HEADERS.map((h) => (
                            <th key={h} style={{ border: '1px solid #e5e7eb', padding: 8, textAlign: 'left' }}>
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {visibleSugarRows.map((r, idx) => (
                          <tr key={`sugar-${idx}`}>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="text" value={r.position || ''} onChange={(e) => updateRowField<SugarRow>('sugarRows', idx, 'position', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.position || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="text" value={r.name || ''} onChange={(e) => updateRowField<SugarRow>('sugarRows', idx, 'name', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.name || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="number" value={r.sugar || ''} onChange={(e) => updateRowField<SugarRow>('sugarRows', idx, 'sugar', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.sugar || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8, color: isEditing ? '#6b7280' : undefined }}>{r.sugarGood || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="number" value={r.sugarBad || ''} onChange={(e) => updateRowField<SugarRow>('sugarRows', idx, 'sugarBad', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.sugarBad || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="number" value={r.paper || ''} onChange={(e) => updateRowField<SugarRow>('sugarRows', idx, 'paper', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.paper || '-'}</td>
                          </tr>
                        ))}
                      </tbody>
                      <tfoot>
                        <tr style={{ background: '#f9fafb', fontWeight: 700 }}>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>Total</td>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>-</td>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleSugarTotals ? visibleSugarTotals.sugar : 0}</td>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleSugarTotals ? visibleSugarTotals.sugarGood : 0}</td>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleSugarTotals ? visibleSugarTotals.sugarBad : 0}</td>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleSugarTotals ? visibleSugarTotals.paper : 0}</td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>
                  {isEditing && (
                    <div style={{ marginTop: 8 }}>
                      <button
                        type="button"
                        onClick={handleAddSugarRow}
                        style={{
                          height: 34,
                          padding: '0 12px',
                          borderRadius: 8,
                          border: '1px solid #cbd5e1',
                          background: '#fff',
                          color: '#111827',
                          fontWeight: 600,
                          cursor: 'pointer',
                        }}
                      >
                        Add Row
                      </button>
                    </div>
                  )}
                  <h3 style={{ margin: '16px 0 10px' }}>Short Procedures</h3>
                  <div style={dPageDetailTableBleedScroll}>
                    <table style={{ width: '100%', minWidth: 560, borderCollapse: 'collapse', fontSize: 14 }}>
                      <thead>
                        <tr style={{ background: '#f3f4f6' }}>
                          {REASON_HEADERS.map((h) => (
                            <th key={h} style={{ border: '1px solid #e5e7eb', padding: 8, textAlign: 'left' }}>
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {visibleReasonRows.map((r, idx) => (
                          <tr key={`reason-${idx}`}>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="text" value={r.reason || ''} onChange={(e) => updateRowField<ReasonRow>('reasonRows', idx, 'reason', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.reason || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="number" value={r.orangeJuice || ''} onChange={(e) => updateRowField<ReasonRow>('reasonRows', idx, 'orangeJuice', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.orangeJuice || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="number" value={r.paper || ''} onChange={(e) => updateRowField<ReasonRow>('reasonRows', idx, 'paper', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.paper || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="number" value={r.coffee || ''} onChange={(e) => updateRowField<ReasonRow>('reasonRows', idx, 'coffee', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.coffee || '-'}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  {isEditing && (
                    <div style={{ marginTop: 8 }}>
                      <button
                        type="button"
                        onClick={handleAddReasonRow}
                        style={{
                          height: 34,
                          padding: '0 12px',
                          borderRadius: 8,
                          border: '1px solid #cbd5e1',
                          background: '#fff',
                          color: '#111827',
                          fontWeight: 600,
                          cursor: 'pointer',
                        }}
                      >
                        Add Row
                      </button>
                    </div>
                  )}

                  <div style={{ marginTop: 14, display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
                    <div>
                      <strong>Not Due</strong>
                      {isEditing ? (
                        <textarea
                          value={String(draft?.notDue ?? '')}
                          maxLength={NOTES_MAX_LENGTH}
                          onChange={(e) => updateDraftField('notDue', e.target.value)}
                          rows={4}
                          style={{
                            marginTop: 6,
                            width: '100%',
                            minHeight: 110,
                            padding: 8,
                            border: '1px solid #d1d5db',
                            borderRadius: 8,
                            resize: 'vertical',
                            whiteSpace: 'pre-wrap',
                          }}
                        />
                      ) : (
                        <p
                          style={{
                            margin: '6px 0 0',
                            minHeight: 110,
                            color: '#374151',
                            whiteSpace: 'pre-wrap',
                            wordBreak: 'break-word',
                          }}
                        >
                          {String(selectedDoc.notDue ?? '') === '' ? '-' : selectedDoc.notDue}
                        </p>
                      )}
                    </div>
                    <div>
                      <strong>Notes</strong>
                      {isEditing ? (
                        <textarea
                          value={String(draft?.notes ?? '')}
                          maxLength={NOTES_MAX_LENGTH}
                          onChange={(e) => updateDraftField('notes', e.target.value)}
                          rows={4}
                          style={{
                            marginTop: 6,
                            width: '100%',
                            minHeight: 110,
                            padding: 8,
                            border: '1px solid #d1d5db',
                            borderRadius: 8,
                            resize: 'vertical',
                            whiteSpace: 'pre-wrap',
                          }}
                        />
                      ) : (
                        <p
                          style={{
                            margin: '6px 0 0',
                            minHeight: 110,
                            color: '#374151',
                            whiteSpace: 'pre-wrap',
                            wordBreak: 'break-word',
                          }}
                        >
                          {String(selectedDoc.notes ?? '') === '' ? '-' : selectedDoc.notes}
                        </p>
                      )}
                    </div>
                  </div>
                </>
              )}
            </div>
          </div>
        )}
      </section>
    </main>
  );
}

