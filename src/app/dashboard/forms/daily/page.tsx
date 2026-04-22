'use client';

import React, { useEffect, useMemo, useState } from 'react';
import { db } from '@/lib/firebase.config';
import { collection, doc, getDocs, increment, serverTimestamp, setDoc } from 'firebase/firestore';
import { getStorage, ref, uploadBytes } from 'firebase/storage';
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
  submittedDateTime?: string;
  grandTotal?: string;
  coffeeSales?: string;
  salesWithoutCoffee?: string;
  paperAtOrangeJuice?: string;
  paperAtTea?: string;
  justPaper?: string;
  notes?: string;
  notDue?: string;
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
  edited?: boolean;
  editedAt?: unknown;
  pdfSaved?: boolean;
  pdfSavedAt?: unknown;
};

const TABLE_HEADERS = [
  'Position',
  'Name',
  'Production',
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
});

function safeStr(value: unknown, maxLength = 80): string {
  if (value == null) return '';
  return String(value).trim().slice(0, maxLength).replace(/[<>]/g, '');
}

function sanitizeFilename(filename: string): string {
  return filename.replace(/[^a-zA-Z0-9._-]/g, '_').replace(/\.\./g, '_').slice(0, 255);
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
  submittedAt: string;
  grandTotal: string;
  coffeeSales: string;
  salesWithoutCoffee: string;
  paperAtOrangeJuice: string;
  paperAtTea: string;
  justPaper: string;
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
    submittedAt,
    grandTotal,
    coffeeSales,
    salesWithoutCoffee,
    paperAtOrangeJuice,
    paperAtTea,
    justPaper,
    reportRows,
    tableRows,
    tableTotals,
    actualOrangeNew,
    actualOrangeReturn,
    extraInputRows,
    extraInputTotals,
    locationSummary,
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
      safeStr(row.amount, 12),
    ]),
    'report'
  );

  const salesPaperSummaryTable = createPdfTable(
    s,
    ['Grand Total', 'CRA Production', 'Production W/Out CRA', 'Prophy @ OE', 'Prophy @ TX', 'Just Prophy'],
    [[grandTotal || '-', coffeeSales || '-', salesWithoutCoffee || '-', paperAtOrangeJuice || '-', paperAtTea || '-', justPaper || '-']],
    'sales-paper-summary'
  );
  const submissionInfoTable = createPdfTable(
    s,
    ['Reason if Late', 'Submitted by office at:'],
    [[reasonIfLate || '-', submittedAt || '-']],
    'submission-info'
  );

  const locationSummaryTable = createPdfTable(
    s,
    ['Preventative', 'Restorative', 'CRA Production', 'Total'],
    [[locationSummary.pineapple || '-', locationSummary.rose || '-', locationSummary.coffeeSales || '-', locationSummary.total || '-']],
    'location-summary'
  );

  const additionalInputsTable = createPdfTable(
    s,
    ['Position', 'Name', 'Patient Seen', 'Insurance', 'Cash', 'Dentical', 'Treatment', 'Primary Teeth', 'Permanent Teeth'],
    [
      ...extraInputRows.map((row) => [
        safeStr(row.position, 20),
        safeStr(row.name, 24),
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
        safeStr(row.sales, 12),
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
        String(tableTotals?.sales ?? 0),
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

  const notesTable = createPdfTable(
    s,
    ['Not Due', 'Notes'],
    [[safeStr(notDue, 300) || '-', safeStr(notes, 300) || '-']],
    'notes-table'
  );

  return React.createElement(
    Document,
    null,
    React.createElement(
      Page,
      { size: 'A4', orientation: 'landscape', style: s.page },
      React.createElement(Text, { style: s.title }, 'Submitted Report'),
      React.createElement(
        Text,
        { style: s.subtitle },
        `Production Date: ${date || '-'} | Office: ${location || '-'} | Generated: ${sanitizedGeneratedDate || '-'}`
      ),
      React.createElement(View, { style: s.section }, React.createElement(Text, { style: s.sectionTitle }, 'Submission Info'), submissionInfoTable),
      React.createElement(View, { style: s.section }, React.createElement(Text, { style: s.sectionTitle }, 'Billers'), reportTable),
      React.createElement(
        View,
        { style: s.section },
        React.createElement(Text, { style: s.sectionTitle }, 'Production 2'),
        salesPaperSummaryTable
      ),
      React.createElement(View, { style: s.section }, React.createElement(Text, { style: s.sectionTitle }, 'Production 1'), locationSummaryTable),
      React.createElement(View, { style: s.section }, React.createElement(Text, { style: s.sectionTitle }, 'Doctors Performance'), additionalInputsTable),
      React.createElement(View, { style: s.section }, React.createElement(Text, { style: s.sectionTitle }, 'CRA'), coffeeTable),
      React.createElement(View, { style: s.section }, React.createElement(Text, { style: s.sectionTitle }, 'Sealant'), sugarTable),
      React.createElement(View, { style: s.section }, React.createElement(Text, { style: s.sectionTitle }, 'Short Procedures'), reasonTable),
      React.createElement(View, { style: s.section }, React.createElement(Text, { style: s.sectionTitle }, 'Notes'), notesTable)
    )
  );
}

function parseNumber(value: unknown): number {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
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
    position: 'Barista',
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
    position: 'Barista',
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
    return {
      ...row,
      coffeeTotal: coffeeHasInput ? String(parseNumber(row.coffeeNew) + parseNumber(row.coffeeReturn)) : '',
      orangeJuiceTotal: orangeHasInput ? String(parseNumber(row.orangeJuiceNew) + parseNumber(row.orangeJuiceReturn)) : '',
    };
  });
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
          .map((d) => ({ id: d.id, ...(d.data() as Omit<FormDoc, 'id'>) }))
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
    setDraft({
      ...selectedDoc,
      reportRows: normalizedReportRows,
      tableRows: normalizedTableRows,
      coffeeActualTotals: selectedDoc.coffeeActualTotals || {},
      sugarRows: (selectedDoc.sugarRows || []).map((row) => ({ ...createEmptySugarRow(), ...row })),
      reasonRows: (selectedDoc.reasonRows || []).map((row) => ({ ...createEmptyReasonRow(), ...row })),
      extraInputRows: normalizedExtraInputRows,
      locationSummary: normalizedLocationSummary,
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
    setDraft((prev) => (prev ? { ...prev, [field]: value } : prev));
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

  const handleSaveEdit = async () => {
    if (!draft) return;
    try {
      const tableRows = computeTableRows(
        (draft.tableRows || []).map((row) => ({
          ...createEmptyTableRow(),
          ...row,
        }))
      );
      const sugarRows = (draft.sugarRows || []).map((row) => ({
        ...createEmptySugarRow(),
        ...row,
      }));
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

      const tableTotals = getTableTotalsFromRows(tableRows);
      const sugarTotals = getSugarTotalsFromRows(sugarRows);
      const hasCoffeeYesInput = tableRows.some((row) => String(row.coffeeYes ?? '').trim() !== '');
      const coffeeSales = hasCoffeeYesInput ? String(tableTotals.coffeeYes * 100) : '';
      const draftLocationSummary = {
        ...createEmptyLocationSummary(),
        ...(draft.locationSummary || {}),
      };
      const locationSummary = {
        pineapple: String(draftLocationSummary.pineapple ?? ''),
        rose: String(draftLocationSummary.rose ?? ''),
        total: String(parseNumber(draftLocationSummary.pineapple) + parseNumber(draftLocationSummary.rose) + parseNumber(coffeeSales)),
      };
      const grandTotal = String(draft.grandTotal ?? '');
      const bothEmpty = grandTotal.trim() === '' && coffeeSales.trim() === '';
      const salesWithoutCoffee = bothEmpty ? '' : String(parseNumber(grandTotal) - parseNumber(coffeeSales));
      const firstReport = reportRowsToSave[0] ?? createEmptyReportRow();
      const coffeeActualTotals = computeCoffeeActualTotals(draft.coffeeActualTotals || {});

      const payload: Omit<FormDoc, 'id'> & { updatedAt: unknown } = {
        ...draft,
        edited: true,
        editedAt: serverTimestamp(),
        reportRows: reportRowsToSave,
        tableRows,
        tableTotals,
        coffeeActualTotals,
        sugarRows,
        sugarTotals,
        reasonRows,
        extraInputRows,
        locationSummary,
        name: firstReport.name ?? '',
        timeStart: firstReport.timeStart ?? '',
        timeEnd: firstReport.timeEnd ?? '',
        chartCount: firstReport.chartCount ?? '',
        amount: firstReport.amount ?? '',
        coffeeSales,
        salesWithoutCoffee,
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
  const visibleSugarRows = isEditing ? draft?.sugarRows || [] : selectedDoc?.sugarRows || [];
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
    ? String((visibleTableTotals?.coffeeYes ?? 0) * 100)
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
        submittedAt: selectedDoc.submittedDateTime || '',
        grandTotal: visibleGrandTotal,
        coffeeSales: visibleCoffeeSales,
        salesWithoutCoffee: visibleSalesWithoutCoffee,
        paperAtOrangeJuice: visiblePaperAtOrangeJuice,
        paperAtTea: visiblePaperAtTea,
        justPaper: visibleJustPaper,
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
        sugarRows: visibleSugarRows,
        sugarTotals: visibleSugarTotals,
        reasonRows: visibleReasonRows,
        notes: isEditing ? String(draft?.notes ?? '') : selectedDoc.notes || '',
        notDue: isEditing ? String(draft?.notDue ?? '') : selectedDoc.notDue || '',
      });
      const blob = await pdf(pdfDoc).toBlob();
      const filename = sanitizeFilename(`${selectedDoc.date || 'no-date'}_${selectedDoc.location || 'no-location'}_submitted-report.pdf`);
      const safeLocation = sanitizeFilename(selectedDoc.location || 'no-location');
      const safeDate = sanitizeFilename(selectedDoc.date || 'no-date');
      const storage = getStorage();
      const targetRef = ref(storage, `Coffee/${safeLocation}/${safeDate}/${filename}`);
      await uploadBytes(targetRef, blob);

      // PDF 저장 완료 시점에 월/지점별 누적 집계도 함께 저장
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
      setSaveMessage('PDF가 Firebase Storage(Coffee)에 저장되었습니다.');
    } catch (e: any) {
      setSaveMessage(`PDF 저장 실패: ${e?.message || '알 수 없는 오류'}`);
    } finally {
      setIsPdfSaving(false);
    }
  };

  return (
    <main style={{ minHeight: '100vh', background: '#fff', padding: 24 }}>
      <section
        style={{
          maxWidth: 2600,
          margin: '0 auto',
          border: '1px solid #e5e7eb',
          borderRadius: 12,
          padding: 20,
          boxShadow: '0 8px 24px rgba(15, 23, 42, 0.06)',
        }}
      >
        <div style={{ marginBottom: 14, display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8 }}>
          <h2 style={{ margin: 0, fontSize: 20 }}>Submitted Finalized Production</h2>
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
                filteredDocs.map((d) => {
                  const active = d.id === selectedId;
                  const statusLabel = d.pdfSaved ? 'Completed' : d.edited ? 'Editing' : '';
                  const statusStyle = d.pdfSaved
                    ? { background: '#dcfce7', color: '#166534', border: '1px solid #86efac' }
                    : { background: '#dbeafe', color: '#1d4ed8', border: '1px solid #93c5fd' };
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
                        borderBottom: '1px solid #f1f5f9',
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

            <div style={{ border: '1px solid #e5e7eb', borderRadius: 10, padding: 12 }}>
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

                  <h3 style={{ margin: '10px 0' }}>Billers</h3>
                  <div style={{ overflowX: 'auto' }}>
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
                                r.amount || '-'
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
                        Add Reporting Row
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
                        selectedDoc.grandTotal || '-'
                      )}
                    </div>
                    <div>
                      <strong>CRA Production:</strong> {visibleCoffeeSales || '-'}
                    </div>
                    <div>
                      <strong>Production W/Out CRA:</strong> {visibleSalesWithoutCoffee || '-'}
                    </div>
                  </div>

                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, minmax(220px, 1fr))', gap: 10, marginTop: 10 }}>
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
                  </div>

                  <h3 style={{ margin: '16px 0 10px' }}>Production 1</h3>
                  <div style={{ overflowX: 'auto' }}>
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
                              visiblePineappleValue || '-'
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
                              visibleRoseValue || '-'
                            )}
                          </td>
                        </tr>
                        <tr>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8, background: '#f9fafb', fontWeight: 700 }}>CRA Production</td>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleCoffeeSales || '-'}</td>
                        </tr>
                        <tr>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8, background: '#f3f4f6', fontWeight: 800 }}>Total</td>
                          <td style={{ border: '1px solid #e5e7eb', padding: 8, fontWeight: 800 }}>{visibleLocationTotal || '-'}</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>

                  <>
                    <h3 style={{ margin: '16px 0 10px' }}>Doctors Performance</h3>
                      <div style={{ overflowX: 'auto' }}>
                        <table style={{ width: '100%', minWidth: 1100, borderCollapse: 'collapse', fontSize: 14 }}>
                          <thead>
                            <tr style={{ background: '#f3f4f6' }}>
                              {[
                                'Position',
                                'Name',
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
                  <h3 style={{ margin: '16px 0 10px' }}>CRA</h3>
                  <div style={{ overflowX: 'auto' }}>
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
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="number" value={r.sales || ''} onChange={(e) => updateRowField<TableRow>('tableRows', idx, 'sales', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.sales || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="number" value={r.coffeeNew || ''} onChange={(e) => updateRowField<TableRow>('tableRows', idx, 'coffeeNew', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.coffeeNew || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="number" value={r.coffeeReturn || ''} onChange={(e) => updateRowField<TableRow>('tableRows', idx, 'coffeeReturn', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.coffeeReturn || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{r.coffeeTotal || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="number" value={r.coffeeNo || ''} onChange={(e) => updateRowField<TableRow>('tableRows', idx, 'coffeeNo', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.coffeeNo || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="number" value={r.renderedCoffee || ''} onChange={(e) => updateRowField<TableRow>('tableRows', idx, 'renderedCoffee', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.renderedCoffee || '-'}</td>
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="number" value={r.coffeeYes || ''} onChange={(e) => updateRowField<TableRow>('tableRows', idx, 'coffeeYes', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.coffeeYes || '-'}</td>
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
                          <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{visibleTableTotals ? visibleTableTotals.sales : 0}</td>
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
                        Add Coffee Row
                      </button>
                    </div>
                  )}
                  <h3 style={{ margin: '16px 0 10px' }}>Sealant</h3>
                  <div style={{ overflowX: 'auto' }}>
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
                            <td style={{ border: '1px solid #e5e7eb', padding: 8 }}>{isEditing ? <input type="number" value={r.sugarGood || ''} onChange={(e) => updateRowField<SugarRow>('sugarRows', idx, 'sugarGood', e.target.value)} style={{ width: '100%', height: 30, border: '1px solid #d1d5db', borderRadius: 6, padding: '0 8px' }} /> : r.sugarGood || '-'}</td>
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
                        Add Sugar Row
                      </button>
                    </div>
                  )}
                  <h3 style={{ margin: '16px 0 10px' }}>Short Procedures</h3>
                  <div style={{ overflowX: 'auto' }}>
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
                        Add Reason Row
                      </button>
                    </div>
                  )}

                  <div style={{ marginTop: 14, display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
                    <div>
                      <strong>Not Due</strong>
                      {isEditing ? (
                        <textarea
                          value={String(draft?.notDue ?? '')}
                          onChange={(e) => updateDraftField('notDue', e.target.value)}
                          rows={4}
                          style={{ marginTop: 6, width: '100%', minHeight: 110, padding: 8, border: '1px solid #d1d5db', borderRadius: 8, resize: 'vertical' }}
                        />
                      ) : (
                        <p style={{ margin: '6px 0 0', minHeight: 110, color: '#374151' }}>{selectedDoc.notDue || '-'}</p>
                      )}
                    </div>
                    <div>
                      <strong>Notes</strong>
                      {isEditing ? (
                        <textarea
                          value={String(draft?.notes ?? '')}
                          onChange={(e) => updateDraftField('notes', e.target.value)}
                          rows={4}
                          style={{ marginTop: 6, width: '100%', minHeight: 110, padding: 8, border: '1px solid #d1d5db', borderRadius: 8, resize: 'vertical' }}
                        />
                      ) : (
                        <p style={{ margin: '6px 0 0', minHeight: 110, color: '#374151' }}>{selectedDoc.notes || '-'}</p>
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


