'use client';

import React, { useEffect, useRef, useState } from 'react';
import { onAuthStateChanged } from 'firebase/auth';
import { auth, db } from '@/lib/firebase.config';
import { doc, getDoc, serverTimestamp, setDoc, updateDoc } from 'firebase/firestore';

type SimpleFormData = {
  location: string;
  date: string;
  reasonIfLate: string;
  submittedDateTime: string;
  grandTotal: string;
  coffeeSales: string;
  salesWithoutCoffee: string;
  paperAtOrangeJuice: string;
  paperAtTea: string;
  justPaper: string;
  prophyTotal: string;
  notes: string;
  notDue: string;
};

type ReportRow = {
  name: string;
  timeStart: string;
  timeEnd: string;
  chartCount: string;
  amount: string;
};

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

const TABLE_COLUMNS: Array<{ key: keyof TableRow; label: string }> = [
  { key: 'position', label: 'Position' },
  { key: 'name', label: 'Name' },
  { key: 'sales', label: '$ Production' },
  { key: 'coffeeNew', label: 'CRA (New)' },
  { key: 'coffeeReturn', label: 'CRA (Return)' },
  { key: 'coffeeTotal', label: 'CRA Total' },
  { key: 'coffeeNo', label: 'CRA (Not Billable)' },
  { key: 'renderedCoffee', label: 'Rendered CRA' },
  { key: 'coffeeYes', label: 'CRA (Billable)' },
  { key: 'orangeJuiceNew', label: 'OE (NP)' },
  { key: 'orangeJuiceReturn', label: 'OE (RC)' },
  { key: 'orangeJuiceTotal', label: 'OE Total' },
];

const SUGAR_COLUMNS: Array<{ key: keyof SugarRow; label: string }> = [
  { key: 'position', label: 'Position' },
  { key: 'name', label: 'Name' },
  { key: 'sugar', label: 'Sealant' },
  { key: 'sugarBad', label: 'Sealant (Redo)' },
  { key: 'sugarGood', label: 'Sealant (Billable)' },
  { key: 'paper', label: 'Prophy' },
];

const REASON_COLUMNS: Array<{ key: keyof ReasonRow; label: string }> = [
  { key: 'reason', label: 'Short Procedures (Reasoning)' },
  { key: 'orangeJuice', label: 'OE' },
  { key: 'paper', label: 'Pro' },
  { key: 'coffee', label: 'CRA' },
];

const POSITION_OPTIONS = ['Doctor', 'RDA', 'DA', 'Extern', 'Working Interview'];
const COFFEE_POSITION_OPTIONS = ['Doctor'];
const FIXED_REASON_OPTIONS = ['Declined/DDP/Pt Left', 'Not Due/Freq', 'Medical Clearance', 'Furture TX/FMS', 'Courtesy, Not Billable', 'Mistakenly Done, Not Billable', 'Uncooperative/Re-eval', 'Not Documented', 'Age Limit/No teeth', 'Not Complete/Filled out'];
const LOCATION_OPTIONS = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia', 'Crowns', 'Endo'];
const REQUIRED_ACCESS = 'production';

const READONLY_FIELDS: Array<keyof TableRow> = ['coffeeTotal', 'orangeJuiceTotal', 'coffeeYes'];
const READONLY_SUGAR_FIELDS: Array<keyof SugarRow> = ['sugarGood'];
const TOTAL_TARGET_FIELDS: Array<keyof TableRow> = [
  'sales',
  'coffeeNew',
  'coffeeReturn',
  'coffeeTotal',
  'coffeeNo',
  'renderedCoffee',
  'coffeeYes',
  'orangeJuiceNew',
  'orangeJuiceReturn',
  'orangeJuiceTotal',
];
const SUGAR_TOTAL_TARGET_FIELDS: Array<keyof SugarRow> = ['sugar', 'sugarGood', 'sugarBad', 'paper'];

function parseNumber(value: string): number {
  const n = Number(String(value).replace(/,/g, ''));
  return Number.isFinite(n) ? n : 0;
}

function formatWithCommas(value: string | number): string {
  const raw = String(value ?? '').replace(/,/g, '').trim();
  if (raw === '') return '';

  const isNegative = raw.startsWith('-');
  const body = isNegative ? raw.slice(1) : raw;
  if (body === '' || body === '.') {
    return isNegative ? `-${body}` : body;
  }

  const [intPartRaw, ...decParts] = body.split('.');
  const intPart = intPartRaw.replace(/\D/g, '');
  const formattedInt = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const hasDecimal = decParts.length > 0;
  const decPart = hasDecimal ? decParts.join('').replace(/\D/g, '') : '';
  const result = hasDecimal ? `${formattedInt}.${decPart}` : formattedInt;
  return isNegative ? `-${result}` : result;
}

function sanitizeNumberInput(raw: string, allowDecimal: boolean): string {
  let s = raw.replace(/,/g, '');
  if (allowDecimal) {
    s = s.replace(/[^\d.]/g, '');
    const firstDot = s.indexOf('.');
    if (firstDot !== -1) {
      s = s.slice(0, firstDot + 1) + s.slice(firstDot + 1).replace(/\./g, '');
    }
  } else {
    s = s.replace(/\D/g, '');
  }
  return s;
}

function roundToCents(n: number): number {
  return Math.round(n * 100) / 100;
}

function toMoneyString(n: number): string {
  return roundToCents(n).toFixed(2);
}

function formatMoneyWithCommas(value: string | number): string {
  const raw = String(value ?? '').replace(/,/g, '').trim();
  if (raw === '') return '';
  const n = Number(raw);
  if (!Number.isFinite(n)) return formatWithCommas(value);

  const fixed = toMoneyString(n);
  const isNegative = fixed.startsWith('-');
  const body = isNegative ? fixed.slice(1) : fixed;
  const [intPart, decPart] = body.split('.');
  const formattedInt = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return `${isNegative ? '-' : ''}${formattedInt}.${decPart}`;
}

function roundToWhole(n: number): number {
  return Math.round(n);
}

function computeRow(row: TableRow): TableRow {
  const coffeeHasInput = row.coffeeNew.trim() !== '' || row.coffeeReturn.trim() !== '';
  const orangeHasInput = row.orangeJuiceNew.trim() !== '' || row.orangeJuiceReturn.trim() !== '';
  const coffeeTotalNum = coffeeHasInput
    ? roundToWhole(parseNumber(row.coffeeNew) + parseNumber(row.coffeeReturn))
    : NaN;
  const coffeeTotal = coffeeHasInput ? String(coffeeTotalNum) : '';
  const orangeJuiceTotalNum = orangeHasInput
    ? roundToWhole(parseNumber(row.orangeJuiceNew) + parseNumber(row.orangeJuiceReturn))
    : NaN;
  const orangeJuiceTotal = orangeHasInput ? String(orangeJuiceTotalNum) : '';

  const coffeeYes =
    coffeeTotal !== '' ? String(roundToWhole(coffeeTotalNum - parseNumber(row.coffeeNo))) : '';

  return {
    ...row,
    coffeeTotal,
    orangeJuiceTotal,
    coffeeYes,
  };
}

function computeProphyTotal(oe: string, tea: string, just: string): string {
  const hasBasis = oe.trim() !== '' || tea.trim() !== '' || just.trim() !== '';
  if (!hasBasis) return '';
  return String(Math.round(parseNumber(oe) + parseNumber(tea) + parseNumber(just)));
}

function computeSugarRow(row: SugarRow): SugarRow {
  const sealantHasBasis = row.sugar.trim() !== '' || row.sugarBad.trim() !== '';
  const sugarGood = sealantHasBasis
    ? String(roundToWhole(parseNumber(row.sugar) - parseNumber(row.sugarBad)))
    : '';
  return { ...row, sugarGood };
}

function createEmptyRow(): TableRow {
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

function getFixedReason(index: number) {
  return FIXED_REASON_OPTIONS[index % FIXED_REASON_OPTIONS.length];
}

function createEmptyReasonRow(index: number): ReasonRow {
  return {
    reason: getFixedReason(index),
    orangeJuice: '',
    paper: '',
    coffee: '',
  };
}

function getTableTotals(rows: TableRow[]): TableTotals {
  const sum = (field: keyof TableRow, round: (n: number) => number) =>
    round(rows.reduce((acc, row) => acc + parseNumber(row[field]), 0));
  return {
    sales: sum('sales', roundToCents),
    coffeeNew: sum('coffeeNew', roundToWhole),
    coffeeReturn: sum('coffeeReturn', roundToWhole),
    coffeeTotal: sum('coffeeTotal', roundToWhole),
    coffeeNo: sum('coffeeNo', roundToWhole),
    renderedCoffee: sum('renderedCoffee', roundToWhole),
    coffeeYes: sum('coffeeYes', roundToWhole),
    orangeJuiceNew: sum('orangeJuiceNew', roundToWhole),
    orangeJuiceReturn: sum('orangeJuiceReturn', roundToWhole),
    orangeJuiceTotal: sum('orangeJuiceTotal', roundToWhole),
  };
}

function getSugarTotals(rows: SugarRow[]): SugarTotals {
  const sum = (field: keyof SugarRow) =>
    roundToWhole(rows.reduce((acc, row) => acc + parseNumber(row[field]), 0));
  return {
    sugar: sum('sugar'),
    sugarGood: sum('sugarGood'),
    sugarBad: sum('sugarBad'),
    paper: sum('paper'),
  };
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

function formatDateTime12h(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  let hours = date.getHours();
  const ampm = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12;
  if (hours === 0) hours = 12;
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds} ${ampm}`;
}

function getNowDateTimeString() {
  return formatDateTime12h(new Date());
}

function normalizeMultilineText(value: string): string {
  return value.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
}

function multilineToLines(value: string): string[] {
  const normalized = normalizeMultilineText(value);
  if (normalized === '') return [];
  return normalized.split('\n');
}

function loadMultilineField(text: unknown, lines: unknown): string {
  if (Array.isArray(lines) && lines.length > 0) {
    return lines.map((line) => String(line)).join('\n');
  }
  return normalizeMultilineText(String(text ?? ''));
}

function getMultilineFirestoreFields(notes: string, notDue: string) {
  const normalizedNotes = normalizeMultilineText(notes);
  const normalizedNotDue = normalizeMultilineText(notDue);
  return {
    notes: normalizedNotes,
    notesLines: multilineToLines(normalizedNotes),
    notDue: normalizedNotDue,
    notDueLines: multilineToLines(normalizedNotDue),
  };
}

function parseDateTimeValue(value: string): Date | null {
  const trimmed = value.trim();
  if (!trimmed) return null;
  const parsed = new Date(trimmed);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function toDateTimeLocalValue(stored: string): string {
  const trimmed = stored.trim();
  if (!trimmed) return '';
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(trimmed)) {
    return trimmed.slice(0, 16);
  }
  const parsed = parseDateTimeValue(trimmed);
  if (!parsed) return '';
  const year = parsed.getFullYear();
  const month = String(parsed.getMonth() + 1).padStart(2, '0');
  const day = String(parsed.getDate()).padStart(2, '0');
  const hours = String(parsed.getHours()).padStart(2, '0');
  const minutes = String(parsed.getMinutes()).padStart(2, '0');
  return `${year}-${month}-${day}T${hours}:${minutes}`;
}

function toStoredDateTime12h(localValue: string): string {
  const trimmed = localValue.trim();
  if (!trimmed) return '';
  const parsed = parseDateTimeValue(trimmed);
  if (!parsed) return trimmed;
  return formatDateTime12h(parsed);
}

function normalizeReportRowsForLoad(rows: Partial<ReportRow>[]): ReportRow[] {
  return rows.map((row) => {
    const merged = { ...createEmptyReportRow(), ...row };
    return {
      ...merged,
      timeStart: toDateTimeLocalValue(String(merged.timeStart ?? '')),
      timeEnd: toDateTimeLocalValue(String(merged.timeEnd ?? '')),
    };
  });
}

function normalizeReportRowsForSave(rows: ReportRow[]): ReportRow[] {
  return rows.map((row) => ({
    ...row,
    timeStart: toStoredDateTime12h(row.timeStart),
    timeEnd: toStoredDateTime12h(row.timeEnd),
  }));
}

function getDurationLabel(start: string, end: string): string {
  if (!start || !end) return '-';
  const startDate = parseDateTimeValue(start);
  const endDate = parseDateTimeValue(end);
  if (!startDate || !endDate) return '-';

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

function hasAnyDataToSave(
  form: SimpleFormData,
  reportRows: ReportRow[],
  tableRows: TableRow[],
  sugarRows: SugarRow[],
  reasonRows: ReasonRow[]
) {
  const hasMainFormValue =
    form.reasonIfLate.trim() !== '' ||
    form.grandTotal.trim() !== '' ||
    form.coffeeSales.trim() !== '' ||
    form.salesWithoutCoffee.trim() !== '' ||
    form.paperAtOrangeJuice.trim() !== '' ||
    form.paperAtTea.trim() !== '' ||
    form.justPaper.trim() !== '' ||
    form.notes.trim() !== '' ||
    form.notDue.trim() !== '';

  const hasReportRowValue = reportRows.some(
    (row) =>
      row.name.trim() !== '' ||
      row.timeStart.trim() !== '' ||
      row.timeEnd.trim() !== '' ||
      row.chartCount.trim() !== '' ||
      row.amount.trim() !== ''
  );

  const hasTableValue = tableRows.some(
    (row) =>
      row.name.trim() !== '' ||
      row.sales.trim() !== '' ||
      row.coffeeNew.trim() !== '' ||
      row.coffeeReturn.trim() !== '' ||
      row.coffeeNo.trim() !== '' ||
      row.renderedCoffee.trim() !== '' ||
      row.coffeeYes.trim() !== '' ||
      row.orangeJuiceNew.trim() !== '' ||
      row.orangeJuiceReturn.trim() !== ''
  );

  const hasSugarValue = sugarRows.some(
    (row) =>
      row.name.trim() !== '' ||
      row.sugar.trim() !== '' ||
      row.sugarGood.trim() !== '' ||
      row.sugarBad.trim() !== '' ||
      row.paper.trim() !== ''
  );

  const hasReasonValue = reasonRows.some(
    (row) => row.orangeJuice.trim() !== '' || row.paper.trim() !== '' || row.coffee.trim() !== ''
  );

  return hasMainFormValue || hasReportRowValue || hasTableValue || hasSugarValue || hasReasonValue;
}

function hasAnyTableRowValue(row: TableRow): boolean {
  return Object.values(row).some((value) => String(value).trim() !== '');
}

function hasAnySugarRowValue(row: SugarRow): boolean {
  return Object.values(row).some((value) => String(value).trim() !== '');
}

function DollarPrefixedNumberInput({
  value,
  onChange,
  readOnly,
  height = 40,
  borderRadius = 8,
}: {
  value: string;
  onChange?: (value: string) => void;
  readOnly?: boolean;
  height?: number;
  borderRadius?: number;
}) {
  const [focused, setFocused] = useState(false);
  const readOnlyBg = readOnly ? '#f3f4f6' : '#fff';
  const displayValue = readOnly || !focused ? formatMoneyWithCommas(value) : value;

  return (
    <div
      style={{
        display: 'flex',
        alignItems: 'center',
        width: '100%',
        height,
        border: '1px solid #d1d5db',
        borderRadius,
        paddingLeft: 8,
        gap: 4,
        background: readOnlyBg,
        boxSizing: 'border-box',
      }}
    >
      <span style={{ color: '#6b7280', fontWeight: 600, fontSize: 13, flexShrink: 0, userSelect: 'none' }}>$</span>
      <input
        type="text"
        inputMode="decimal"
        value={displayValue}
        onChange={(e) => {
          let next = sanitizeNumberInput(e.target.value, true);
          const dot = next.indexOf('.');
          if (dot !== -1) {
            next = next.slice(0, dot + 1) + next.slice(dot + 1, dot + 3);
          }
          onChange?.(next);
        }}
        onFocus={() => setFocused(true)}
        onBlur={() => {
          setFocused(false);
          if (!onChange || value.trim() === '') return;
          const normalized = toMoneyString(parseNumber(value));
          if (normalized !== value.replace(/,/g, '')) {
            onChange(normalized);
          }
        }}
        readOnly={readOnly}
        style={{
          flex: 1,
          minWidth: 0,
          height: Math.max(height - 4, 24),
          border: 'none',
          outline: 'none',
          background: 'transparent',
          color: readOnly ? '#6b7280' : '#111827',
          fontSize: 14,
        }}
      />
    </div>
  );
}

function FormattedNumberInput({
  value,
  onChange,
  readOnly,
  height = 40,
  borderRadius = 8,
  style,
}: {
  value: string;
  onChange?: (value: string) => void;
  readOnly?: boolean;
  height?: number;
  borderRadius?: number;
  style?: React.CSSProperties;
}) {
  const [focused, setFocused] = useState(false);
  const displayValue = readOnly || !focused ? formatWithCommas(value) : value;

  return (
    <input
      type="text"
      inputMode="numeric"
      value={displayValue}
      readOnly={readOnly}
      onChange={(e) => onChange?.(sanitizeNumberInput(e.target.value, false))}
      onFocus={() => setFocused(true)}
      onBlur={() => setFocused(false)}
      style={{
        width: '100%',
        height,
        padding: '0 10px',
        border: '1px solid #d1d5db',
        borderRadius,
        background: readOnly ? '#f3f4f6' : '#fff',
        color: readOnly ? '#6b7280' : '#111827',
        ...style,
      }}
    />
  );
}

export default function Page() {
  const [form, setForm] = useState<SimpleFormData>({
    location: '',
    date: '',
    reasonIfLate: '',
    submittedDateTime: '',
    grandTotal: '',
    coffeeSales: '',
    salesWithoutCoffee: '',
    paperAtOrangeJuice: '',
    paperAtTea: '',
    justPaper: '',
    prophyTotal: '',
    notes: '',
    notDue: '',
  });
  const [reportRows, setReportRows] = useState<ReportRow[]>([createEmptyReportRow()]);
  const [status, setStatus] = useState('');
  const [submitFeedback, setSubmitFeedback] = useState<{
    type: 'success' | 'error';
    message: string;
  } | null>(null);
  const lastSavedKeyRef = useRef('');
  const [tableRows, setTableRows] = useState<TableRow[]>(
    Array.from({ length: 4 }, () => createEmptyRow())
  );
  const [sugarRows, setSugarRows] = useState<SugarRow[]>(
    Array.from({ length: 6 }, () => createEmptySugarRow())
  );
  const [reasonRows, setReasonRows] = useState<ReasonRow[]>(
    Array.from({ length: 10 }, (_, i) => createEmptyReasonRow(i))
  );
  const [docSubmitLock, setDocSubmitLock] = useState(false);
  const [locationOptions, setLocationOptions] = useState<string[]>([]);
  const [pageAccess, setPageAccess] = useState<'loading' | 'allowed'>('loading');

  useEffect(() => {
    const goHome = () => {
      if (typeof window !== 'undefined') {
        window.location.replace('/');
      }
    };

    const unsubscribe = onAuthStateChanged(auth, async (user) => {
      try {
        if (!user) {
          goHome();
          return;
        }

        setPageAccess('loading');
        const snap = await getDoc(doc(db, 'users', user.uid));
        const data = snap.exists() ? snap.data() : undefined;
        const accessRaw = data?.forms;
        const accessList = Array.isArray(accessRaw)
          ? accessRaw.map((item) => String(item).trim())
          : [];
        const hasAccess = accessList.includes(REQUIRED_ACCESS);

        if (!hasAccess) {
          goHome();
          return;
        }

        const officesRaw = data?.offices;
        const userOffices = new Set(
          Array.isArray(officesRaw)
            ? officesRaw.map((office) => String(office).trim()).filter((office) => office !== '')
            : []
        );
        const allowedLocations = LOCATION_OPTIONS.filter((location) => userOffices.has(location));

        setLocationOptions(allowedLocations);
        setForm((prev) => {
          if (allowedLocations.length === 1) {
            return prev.location === allowedLocations[0]
              ? prev
              : { ...prev, location: allowedLocations[0] };
          }
          if (prev.location && !allowedLocations.includes(prev.location)) {
            return { ...prev, location: '' };
          }
          return prev;
        });
        setPageAccess('allowed');
      } catch {
        goHome();
      }
    });

    return () => unsubscribe();
  }, []);

  useEffect(() => {
    const hasCoffeeTotalBasis = tableRows.some(
      (row) => row.coffeeNew.trim() !== '' || row.coffeeReturn.trim() !== ''
    );
    const coffeeYesSum = tableRows.reduce(
      (sum, row) => sum + parseNumber(computeRow(row).coffeeYes),
      0
    );
    const calculatedCoffeeSales = hasCoffeeTotalBasis
      ? toMoneyString(coffeeYesSum * 61)
      : '';

    setForm((prev) =>
      prev.coffeeSales === calculatedCoffeeSales ? prev : { ...prev, coffeeSales: calculatedCoffeeSales }
    );
  }, [tableRows]);

  useEffect(() => {
    const bothEmpty = form.grandTotal.trim() === '' && form.coffeeSales.trim() === '';
    const calculatedSalesWithoutCoffee = bothEmpty
      ? ''
      : toMoneyString(parseNumber(form.grandTotal) - parseNumber(form.coffeeSales));

    setForm((prev) =>
      prev.salesWithoutCoffee === calculatedSalesWithoutCoffee
        ? prev
        : { ...prev, salesWithoutCoffee: calculatedSalesWithoutCoffee }
    );
  }, [form.grandTotal, form.coffeeSales]);

  useEffect(() => {
    const calculatedProphyTotal = computeProphyTotal(
      form.paperAtOrangeJuice,
      form.paperAtTea,
      form.justPaper
    );
    setForm((prev) =>
      prev.prophyTotal === calculatedProphyTotal ? prev : { ...prev, prophyTotal: calculatedProphyTotal }
    );
  }, [form.paperAtOrangeJuice, form.paperAtTea, form.justPaper]);

  const getDocId = () => {
    const safeLocation = form.location.trim().replace(/\s+/g, '_');
    return `${form.date}_${safeLocation}`;
  };

  useEffect(() => {
    if (!form.date || !form.location.trim()) {
      setDocSubmitLock(false);
      return;
    }

    setDocSubmitLock(false);
    let cancelled = false;
    const loadByDateAndLocation = async () => {
      try {
        const snap = await getDoc(doc(db, 'simple-forms', getDocId()));
        if (cancelled) return;

        if (!snap.exists()) {
          if (!cancelled) setDocSubmitLock(false);
          setReportRows([createEmptyReportRow()]);
          setTableRows(Array.from({ length: 4 }, () => createEmptyRow()));
          setSugarRows(Array.from({ length: 6 }, () => createEmptySugarRow()));
          setReasonRows(Array.from({ length: 10 }, (_, i) => createEmptyReasonRow(i)));
          setForm((prev) => ({
            ...prev,
            reasonIfLate: '',
            submittedDateTime: '',
            grandTotal: '',
            coffeeSales: '',
            salesWithoutCoffee: '',
            paperAtOrangeJuice: '',
            paperAtTea: '',
            justPaper: '',
            prophyTotal: '',
            notes: '',
            notDue: '',
          }));
          lastSavedKeyRef.current = '';
          return;
        }

        const data = snap.data() as any;
        const existingSubmitted = String(data.submittedDateTime ?? '').trim();
        if (existingSubmitted !== '') {
          if (!cancelled) setDocSubmitLock(true);
          setReportRows([createEmptyReportRow()]);
          setTableRows(Array.from({ length: 4 }, () => createEmptyRow()));
          setSugarRows(Array.from({ length: 6 }, () => createEmptySugarRow()));
          setReasonRows(Array.from({ length: 10 }, (_, i) => createEmptyReasonRow(i)));
          setForm((prev) => ({
            ...prev,
            reasonIfLate: '',
            submittedDateTime: '',
            grandTotal: '',
            coffeeSales: '',
            salesWithoutCoffee: '',
            paperAtOrangeJuice: '',
            paperAtTea: '',
            justPaper: '',
            prophyTotal: '',
            notes: '',
            notDue: '',
          }));
          lastSavedKeyRef.current = '';
          setStatus('');
          return;
        }

        const loadedReportRows =
          Array.isArray(data.reportRows) && data.reportRows.length > 0
            ? normalizeReportRowsForLoad(data.reportRows)
            : [createEmptyReportRow()];
        const loadedTableRows =
          Array.isArray(data.tableRows) && data.tableRows.length > 0
            ? data.tableRows.map((row: Partial<TableRow>) => computeRow({ ...createEmptyRow(), ...row }))
            : Array.from({ length: 4 }, () => createEmptyRow());
        const loadedSugarRows =
          Array.isArray(data.sugarRows) && data.sugarRows.length > 0
            ? data.sugarRows.map((row: Partial<SugarRow>) => computeSugarRow({ ...createEmptySugarRow(), ...row }))
            : Array.from({ length: 6 }, () => createEmptySugarRow());
        const loadedReasonRows =
          Array.isArray(data.reasonRows) && data.reasonRows.length > 0
            ? data.reasonRows.map((row: Partial<ReasonRow>, index: number) => ({ ...createEmptyReasonRow(index), ...row }))
            : Array.from({ length: 10 }, (_, i) => createEmptyReasonRow(i));

        setReportRows(loadedReportRows);
        setTableRows(loadedTableRows);
        setSugarRows(loadedSugarRows);
        setReasonRows(loadedReasonRows);
        setForm((prev) => ({
          ...prev,
          reasonIfLate: String(data.reasonIfLate ?? ''),
          submittedDateTime: String(data.submittedDateTime ?? ''),
          grandTotal: String(data.grandTotal ?? ''),
          coffeeSales: String(data.coffeeSales ?? ''),
          salesWithoutCoffee: String(data.salesWithoutCoffee ?? ''),
          paperAtOrangeJuice: String(data.paperAtOrangeJuice ?? ''),
          paperAtTea: String(data.paperAtTea ?? ''),
          justPaper: String(data.justPaper ?? ''),
          prophyTotal: String(data.prophyTotal ?? ''),
          notes: loadMultilineField(data.notes, data.notesLines),
          notDue: loadMultilineField(data.notDue, data.notDueLines),
        }));
        lastSavedKeyRef.current = '';
        if (!cancelled) setDocSubmitLock(false);
      } catch {
        if (cancelled) return;
        setDocSubmitLock(false);
        setStatus('Failed to load');
      }
    };

    loadByDateAndLocation();
    return () => {
      cancelled = true;
    };
  }, [form.date, form.location]);

  useEffect(() => {
    if (!form.date || !form.location.trim()) {
      return;
    }
    if (docSubmitLock) {
      return;
    }
    if (form.submittedDateTime.trim() !== '') {
      return;
    }
    if (!hasAnyDataToSave(form, reportRows, tableRows, sugarRows, reasonRows)) {
      return;
    }

    const normalizedRows = tableRows.map(computeRow).filter(hasAnyTableRowValue);
    const normalizedSugarRows = sugarRows.map(computeSugarRow).filter(hasAnySugarRowValue);
    const saveKey = `${JSON.stringify(form)}_${JSON.stringify(reportRows)}_${JSON.stringify(normalizedRows)}_${JSON.stringify(normalizedSugarRows)}_${JSON.stringify(reasonRows)}`;
    if (saveKey === lastSavedKeyRef.current) {
      return;
    }

    const timer = setTimeout(async () => {
      try {
        const docId = getDocId();
        const docRef = doc(db, 'simple-forms', docId);
        const snap = await getDoc(docRef);
        if (snap.exists()) {
          const submitted = String(snap.data()?.submittedDateTime ?? '').trim();
          if (submitted !== '') {
            return;
          }
        }

        const storedReportRows = normalizeReportRowsForSave(reportRows);
        const firstReportRow = storedReportRows[0] ?? createEmptyReportRow();
        const tableTotals = getTableTotals(normalizedRows);
        const sugarTotals = getSugarTotals(normalizedSugarRows);
        const { submittedDateTime: _submittedDateTime, notes, notDue, ...formWithoutSubmittedDateTime } = form;
        await setDoc(
          docRef,
          {
            ...formWithoutSubmittedDateTime,
            ...getMultilineFirestoreFields(notes, notDue),
            reportRows: storedReportRows,
            name: firstReportRow.name.trim(),
            timeStart: firstReportRow.timeStart,
            timeEnd: firstReportRow.timeEnd,
            chartCount: firstReportRow.chartCount,
            amount: firstReportRow.amount,
            tableRows: normalizedRows,
            tableTotals,
            sugarRows: normalizedSugarRows,
            sugarTotals,
            reasonRows,
            updatedAt: serverTimestamp(),
          },
          { merge: true }
        );
        lastSavedKeyRef.current = saveKey;
      } catch {
      }
    }, 700);

    return () => clearTimeout(timer);
  }, [form, reportRows, tableRows, sugarRows, reasonRows, docSubmitLock]);

  const handleReportRowChange = (rowIndex: number, field: keyof ReportRow, value: string) => {
    setReportRows((prev) => {
      const next = [...prev];
      next[rowIndex] = { ...next[rowIndex], [field]: value };
      return next;
    });
  };

  const handleAddReportRow = () => {
    setReportRows((prev) => [...prev, createEmptyReportRow()]);
  };

  const handleCellChange = (rowIndex: number, field: keyof TableRow, value: string) => {
    if (READONLY_FIELDS.includes(field)) {
      return;
    }

    setTableRows((prev) => {
      const next = [...prev];
      const updatedRow = { ...next[rowIndex], [field]: value };
      next[rowIndex] = computeRow(updatedRow);
      return next;
    });
  };

  const handleAddTableRow = () => {
    setTableRows((prev) => [...prev, createEmptyRow()]);
  };

  const getColumnTotal = (field: keyof TableRow) => {
    const sum = tableRows.reduce((acc, row) => acc + parseNumber(computeRow(row)[field]), 0);
    return field === 'sales' ? roundToCents(sum) : roundToWhole(sum);
  };

  const handleSugarCellChange = (rowIndex: number, field: keyof SugarRow, value: string) => {
    if (READONLY_SUGAR_FIELDS.includes(field)) return;
    setSugarRows((prev) => {
      const next = [...prev];
      next[rowIndex] = computeSugarRow({ ...next[rowIndex], [field]: value });
      return next;
    });
  };

  const handleAddSugarRow = () => {
    setSugarRows((prev) => [...prev, createEmptySugarRow()]);
  };

  const getSugarColumnTotal = (field: keyof SugarRow) => {
    const sum = sugarRows.reduce((acc, row) => acc + parseNumber(computeSugarRow(row)[field]), 0);
    return roundToWhole(sum);
  };
  const handleReasonCellChange = (rowIndex: number, field: keyof ReasonRow, value: string) => {
    if (field === 'reason') return;

    setReasonRows((prev) => {
      const next = [...prev];
      next[rowIndex] = { ...next[rowIndex], [field]: value };
      return next;
    });
  };

  const handleSubmit = async () => {
    if (!form.location.trim()) {
      setStatus('Please select Office.');
      return;
    }

    if (!form.date) {
      setStatus('Please select Date of Production.');
      return;
    }

    if (docSubmitLock) {
      setStatus('The office production for this date has already been submitted.');
      return;
    }

    try {
      const docId = getDocId();
      const existingSnap = await getDoc(doc(db, 'simple-forms', docId));
      const alreadySubmitted = String(existingSnap.data()?.submittedDateTime ?? '').trim();
      if (alreadySubmitted !== '') {
        setStatus('The office production for this date has already been submitted. You cannot submit again');
        return;
      }

      const normalizedRows = tableRows.map(computeRow).filter(hasAnyTableRowValue);
      const normalizedSugarRows = sugarRows.map(computeSugarRow).filter(hasAnySugarRowValue);
      const storedReportRows = normalizeReportRowsForSave(reportRows);
      const firstReportRow = storedReportRows[0] ?? createEmptyReportRow();
      const tableTotals = getTableTotals(normalizedRows);
      const sugarTotals = getSugarTotals(normalizedSugarRows);
      const submittedDateTime = getNowDateTimeString();

      const { notes, notDue, ...formWithoutMultiline } = form;
      const docRef = doc(db, 'simple-forms', docId);
      const payload = {
        ...formWithoutMultiline,
        ...getMultilineFirestoreFields(notes, notDue),
        submittedDateTime,
        reportRows: storedReportRows,
        name: firstReportRow.name.trim(),
        timeStart: firstReportRow.timeStart,
        timeEnd: firstReportRow.timeEnd,
        chartCount: firstReportRow.chartCount,
        amount: firstReportRow.amount,
        tableRows: normalizedRows,
        tableTotals,
        sugarRows: normalizedSugarRows,
        sugarTotals,
        reasonRows,
        updatedAt: serverTimestamp(),
      };

      // Keep fields written by other pages; only patch this page's fields.
      if (existingSnap.exists()) {
        await updateDoc(docRef, payload);
      } else {
        await setDoc(docRef, payload);
      }

      lastSavedKeyRef.current = '';
      setForm({
        location: locationOptions.length === 1 ? locationOptions[0] : '',
        date: '',
        reasonIfLate: '',
        submittedDateTime: '',
        grandTotal: '',
        coffeeSales: '',
        salesWithoutCoffee: '',
        paperAtOrangeJuice: '',
        paperAtTea: '',
        justPaper: '',
        prophyTotal: '',
        notes: '',
        notDue: '',
      });
      setReportRows([createEmptyReportRow()]);
      setTableRows(Array.from({ length: 4 }, () => createEmptyRow()));
      setSugarRows(Array.from({ length: 6 }, () => createEmptySugarRow()));
      setReasonRows(Array.from({ length: 10 }, (_, i) => createEmptyReasonRow(i)));
      setStatus('');
      setSubmitFeedback({ type: 'success', message: 'Submitted Successfully' });
    } catch {
      setSubmitFeedback({ type: 'error', message: 'Failed to Submit' });
    }
  };

  if (pageAccess !== 'allowed') {
    return null;
  }

  return (
    <main
      style={{
        minHeight: '100vh',
        display: 'grid',
        placeItems: 'center',
        background: '#ffffff',
        padding: '48px 24px',
      }}
    >
      <section
        style={{
          width: '100%',
          maxWidth: 1500,
          background: '#fff',
          border: '1px solid #e5e7eb',
          borderRadius: 12,
          padding: 20,
          boxShadow: '0 8px 24px rgba(15, 23, 42, 0.06)',
        }}
      >
        <h1 style={{ margin: '0 0 16px', fontSize: 28, fontWeight: 800, color: '#111827' }}>Production</h1>
        <div
          style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(2, minmax(160px, 1fr))',
            gap: 12,
            marginBottom: 12,
          }}
        >
          <div>
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Date of Production</label>
            <input
              type="date"
              value={form.date}
              onChange={(e) => setForm((prev) => ({ ...prev, date: e.target.value }))}
              style={{
                width: '100%',
                height: 40,
                padding: '0 10px',
                border: '1px solid #d1d5db',
                borderRadius: 8,
                background: '#fff',
              }}
            />
          </div>
          <div>
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Office</label>
            {locationOptions.length === 1 ? (
              <input
                type="text"
                readOnly
                value={locationOptions[0]}
                style={{
                  width: '100%',
                  height: 40,
                  padding: '0 10px',
                  border: '1px solid #d1d5db',
                  borderRadius: 8,
                  background: '#f3f4f6',
                  color: '#6b7280',
                }}
              />
            ) : (
              <select
                value={form.location}
                onChange={(e) => setForm((prev) => ({ ...prev, location: e.target.value }))}
                disabled={locationOptions.length === 0}
                style={{
                  width: '100%',
                  height: 40,
                  padding: '0 10px',
                  border: '1px solid #d1d5db',
                  borderRadius: 8,
                  background: locationOptions.length === 0 ? '#f3f4f6' : '#fff',
                  color: locationOptions.length === 0 ? '#6b7280' : '#111827',
                }}
              >
                <option value="">Select</option>
                {locationOptions.map((location) => (
                  <option key={location} value={location}>
                    {location}
                  </option>
                ))}
              </select>
            )}
          </div>
        </div>
        {docSubmitLock ? (
          <p
            role="alert"
            style={{
              margin: '0 0 12px',
              padding: '12px 14px',
              borderRadius: 8,
              background: '#fef2f2',
              border: '1px solid #fecaca',
              color: '#991b1b',
              fontSize: 14,
              fontWeight: 600,
            }}
          >
            {`(${getDocId()}) has already been submitted. Please choose different date or office.`}
          </p>
        ) : null}

        <fieldset
          disabled={docSubmitLock}
          style={{
            border: 'none',
            margin: 0,
            padding: 0,
            minWidth: 0,
          }}
        >
          <legend style={{ position: 'absolute', width: 1, height: 1, padding: 0, margin: -1, overflow: 'hidden', clip: 'rect(0,0,0,0)', whiteSpace: 'nowrap', border: 0 }}>
            Production form
          </legend>
          <div style={{ marginBottom: 12 }}>
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Reason if Late</label>
            <input
              type="text"
              value={form.reasonIfLate}
              onChange={(e) => setForm((prev) => ({ ...prev, reasonIfLate: e.target.value }))}
              style={{
                width: '100%',
                maxWidth: 480,
                height: 40,
                padding: '0 10px',
                border: '1px solid #d1d5db',
                borderRadius: 8,
              }}
            />
          </div>

        {reportRows.map((row, rowIndex) => (
          <div
            key={`report-row-${rowIndex}`}
            style={{
              display: 'grid',
              gridTemplateColumns: 'repeat(6, minmax(140px, 1fr))',
              gap: 12,
              marginBottom: 12,
            }}
          >
            <div>
              <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Biller Name</label>
              <input
                type="text"
                value={row.name}
                onChange={(e) => handleReportRowChange(rowIndex, 'name', e.target.value)}
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
              <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Date/Time Start</label>
              <input
                type="datetime-local"
                value={row.timeStart}
                onChange={(e) => handleReportRowChange(rowIndex, 'timeStart', e.target.value)}
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
              <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Date/Time End</label>
              <input
                type="datetime-local"
                value={row.timeEnd}
                onChange={(e) => handleReportRowChange(rowIndex, 'timeEnd', e.target.value)}
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
              <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Duration</label>
              <input
                type="text"
                readOnly
                value={getDurationLabel(row.timeStart, row.timeEnd)}
                style={{
                  width: '100%',
                  height: 40,
                  padding: '0 10px',
                  border: '1px solid #d1d5db',
                  borderRadius: 8,
                  background: '#f3f4f6',
                  color: '#6b7280',
                }}
              />
            </div>
            <div>
              <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}># Chart</label>
              <FormattedNumberInput
                value={row.chartCount}
                onChange={(value) => handleReportRowChange(rowIndex, 'chartCount', value)}
              />
            </div>
            <div>
              <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>$ Amount</label>
              <DollarPrefixedNumberInput
                value={row.amount}
                onChange={(value) => handleReportRowChange(rowIndex, 'amount', value)}
              />
            </div>
          </div>
        ))}
        <div style={{ marginBottom: 16 }}>
          <button
            type="button"
            onClick={handleAddReportRow}
            style={{
              height: 40,
              padding: '0 16px',
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

        <div
          style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(3, minmax(180px, 1fr))',
            gap: 12,
            marginTop: 20,
            marginBottom: 8,
          }}
        >
          <div>
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Grand Total</label>
            <DollarPrefixedNumberInput
              value={form.grandTotal}
              onChange={(value) => setForm((prev) => ({ ...prev, grandTotal: value }))}
            />
          </div>
          <div>
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>CRA Production</label>
            <DollarPrefixedNumberInput value={form.coffeeSales} readOnly />
          </div>
          <div>
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Production W/Out CRA</label>
            <DollarPrefixedNumberInput value={form.salesWithoutCoffee} readOnly />
          </div>
        </div>
        <div
          style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(4, minmax(160px, 1fr))',
            gap: 12,
            marginBottom: 8,
          }}
        >
          <div>
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Prophy @ OE</label>
            <FormattedNumberInput
              value={form.paperAtOrangeJuice}
              onChange={(value) => setForm((prev) => ({ ...prev, paperAtOrangeJuice: value }))}
            />
          </div>
          <div>
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Prophy @ TX</label>
            <FormattedNumberInput
              value={form.paperAtTea}
              onChange={(value) => setForm((prev) => ({ ...prev, paperAtTea: value }))}
            />
          </div>
          <div>
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Just Prophy</label>
            <FormattedNumberInput
              value={form.justPaper}
              onChange={(value) => setForm((prev) => ({ ...prev, justPaper: value }))}
            />
          </div>
          <div>
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Actual Prophy</label>
            <FormattedNumberInput value={form.prophyTotal} readOnly />
          </div>
        </div>

        <div style={{ marginTop: 20, marginBottom: 8 }}>
          <button
            type="button"
            onClick={handleAddTableRow}
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
            Add Row
          </button>
        </div>

        <div style={{ overflowX: 'auto' }}>
          <table
            style={{
              width: '100%',
              minWidth: 1300,
              borderCollapse: 'collapse',
              border: '1px solid #e5e7eb',
              fontSize: 14,
            }}
          >
            <thead>
              <tr style={{ background: '#f3f4f6' }}>
                {TABLE_COLUMNS.map((column) => (
                  <th
                    key={column.key}
                    style={{
                      border: '1px solid #e5e7eb',
                      padding: '8px',
                      textAlign: 'left',
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {column.label}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {tableRows.map((row, rowIndex) => {
                const derivedRow = computeRow(row);
                return (
                <tr key={`row-${rowIndex}`}>
                  {TABLE_COLUMNS.map((column) => (
                    <td key={`${rowIndex}-${column.key}`} style={{ border: '1px solid #e5e7eb', padding: '6px' }}>
                      {column.key === 'position' ? (
                        <select
                          value={row.position}
                          onChange={(e) => handleCellChange(rowIndex, 'position', e.target.value)}
                          style={{
                            width: '100%',
                            height: 34,
                            border: '1px solid #d1d5db',
                            borderRadius: 6,
                            padding: '0 8px',
                            background: '#fff',
                          }}
                        >
                          <option value="">Select</option>
                          {COFFEE_POSITION_OPTIONS.map((option) => (
                            <option key={option} value={option}>
                              {option}
                            </option>
                          ))}
                        </select>
                      ) : column.key === 'sales' ? (
                        <DollarPrefixedNumberInput
                          height={34}
                          borderRadius={6}
                          value={row.sales}
                          onChange={(value) => handleCellChange(rowIndex, 'sales', value)}
                        />
                      ) : TOTAL_TARGET_FIELDS.includes(column.key) ? (
                        <FormattedNumberInput
                          height={34}
                          borderRadius={6}
                          value={
                            READONLY_FIELDS.includes(column.key)
                              ? derivedRow[column.key]
                              : row[column.key]
                          }
                          readOnly={READONLY_FIELDS.includes(column.key)}
                          onChange={(value) => handleCellChange(rowIndex, column.key, value)}
                          style={{ padding: '0 8px' }}
                        />
                      ) : (
                        <input
                          type="text"
                          value={row[column.key]}
                          onChange={(e) => handleCellChange(rowIndex, column.key, e.target.value)}
                          style={{
                            width: '100%',
                            height: 34,
                            border: '1px solid #d1d5db',
                            borderRadius: 6,
                            padding: '0 8px',
                            background: '#fff',
                            color: '#111827',
                          }}
                        />
                      )}
                    </td>
                  ))}
                </tr>
              );
              })}
            </tbody>
            <tfoot>
              <tr style={{ background: '#f9fafb', fontWeight: 700 }}>
                {TABLE_COLUMNS.map((column) => {
                  let value = '';
                  if (column.key === 'position') value = 'Total';
                  else if (column.key === 'name') value = '-';
                  else if (TOTAL_TARGET_FIELDS.includes(column.key)) {
                    value =
                      column.key === 'sales'
                        ? formatMoneyWithCommas(getColumnTotal(column.key))
                        : formatWithCommas(getColumnTotal(column.key));
                  }

                  const displayValue = column.key === 'sales' && value !== '' ? `$${value}` : value;

                  return (
                    <td key={`total-${column.key}`} style={{ border: '1px solid #e5e7eb', padding: '8px' }}>
                      {displayValue}
                    </td>
                  );
                })}
              </tr>
            </tfoot>
          </table>
        </div>

        <div style={{ marginTop: 24, marginBottom: 8 }}>
          <button
            type="button"
            onClick={handleAddSugarRow}
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
            Add Row
          </button>
        </div>

        <div style={{ overflowX: 'auto' }}>
          <table
            style={{
              width: '100%',
              minWidth: 860,
              borderCollapse: 'collapse',
              border: '1px solid #e5e7eb',
              fontSize: 14,
            }}
          >
            <thead>
              <tr style={{ background: '#f3f4f6' }}>
                {SUGAR_COLUMNS.map((column) => (
                  <th
                    key={column.key}
                    style={{
                      border: '1px solid #e5e7eb',
                      padding: '8px',
                      textAlign: 'left',
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {column.label}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {sugarRows.map((row, rowIndex) => {
                const derivedSugar = computeSugarRow(row);
                return (
                <tr key={`sugar-row-${rowIndex}`}>
                  {SUGAR_COLUMNS.map((column) => (
                    <td key={`sugar-${rowIndex}-${column.key}`} style={{ border: '1px solid #e5e7eb', padding: '6px' }}>
                      {column.key === 'position' ? (
                        <select
                          value={row.position}
                          onChange={(e) => handleSugarCellChange(rowIndex, 'position', e.target.value)}
                          style={{
                            width: '100%',
                            height: 34,
                            border: '1px solid #d1d5db',
                            borderRadius: 6,
                            padding: '0 8px',
                            background: '#fff',
                          }}
                        >
                          <option value="">Select</option>
                          {POSITION_OPTIONS.map((option) => (
                            <option key={option} value={option}>
                              {option}
                            </option>
                          ))}
                        </select>
                      ) : column.key === 'name' ? (
                        <input
                          type="text"
                          value={row[column.key]}
                          onChange={(e) => handleSugarCellChange(rowIndex, column.key, e.target.value)}
                          style={{
                            width: '100%',
                            height: 34,
                            border: '1px solid #d1d5db',
                            borderRadius: 6,
                            padding: '0 8px',
                            background: '#fff',
                            color: '#111827',
                          }}
                        />
                      ) : (
                        <FormattedNumberInput
                          height={34}
                          borderRadius={6}
                          value={
                            READONLY_SUGAR_FIELDS.includes(column.key)
                              ? derivedSugar[column.key]
                              : row[column.key]
                          }
                          readOnly={READONLY_SUGAR_FIELDS.includes(column.key)}
                          onChange={(value) => handleSugarCellChange(rowIndex, column.key, value)}
                          style={{ padding: '0 8px' }}
                        />
                      )}
                    </td>
                  ))}
                </tr>
              );
              })}
            </tbody>
            <tfoot>
              <tr style={{ background: '#f9fafb', fontWeight: 700 }}>
                {SUGAR_COLUMNS.map((column) => {
                  let value = '';
                  if (column.key === 'position') value = 'Total';
                  else if (column.key === 'name') value = '-';
                  else if (SUGAR_TOTAL_TARGET_FIELDS.includes(column.key)) value = formatWithCommas(getSugarColumnTotal(column.key));

                  return (
                    <td key={`sugar-total-${column.key}`} style={{ border: '1px solid #e5e7eb', padding: '8px' }}>
                      {value}
                    </td>
                  );
                })}
              </tr>
            </tfoot>
          </table>
        </div>

        <div style={{ marginTop: 24, overflowX: 'auto' }}>
          <table
            style={{
              width: '100%',
              minWidth: 640,
              borderCollapse: 'collapse',
              border: '1px solid #e5e7eb',
              fontSize: 14,
            }}
          >
            <thead>
              <tr style={{ background: '#f3f4f6' }}>
                {REASON_COLUMNS.map((column) => (
                  <th
                    key={column.key}
                    style={{
                      border: '1px solid #e5e7eb',
                      padding: '8px',
                      textAlign: 'left',
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {column.label}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {reasonRows.map((row, rowIndex) => (
                <tr key={`reason-row-${rowIndex}`}>
                  {REASON_COLUMNS.map((column) => (
                    <td key={`reason-${rowIndex}-${column.key}`} style={{ border: '1px solid #e5e7eb', padding: '6px' }}>
                      {column.key === 'reason' ? (
                        <input
                          type="text"
                          value={row[column.key]}
                          readOnly
                          style={{
                            width: '100%',
                            height: 34,
                            border: '1px solid #d1d5db',
                            borderRadius: 6,
                            padding: '0 8px',
                            background: '#f3f4f6',
                            color: '#6b7280',
                          }}
                        />
                      ) : (
                        <FormattedNumberInput
                          height={34}
                          borderRadius={6}
                          value={row[column.key]}
                          onChange={(value) => handleReasonCellChange(rowIndex, column.key, value)}
                          style={{ padding: '0 8px' }}
                        />
                      )}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <div
          style={{
            marginTop: 24,
            display: 'grid',
            gridTemplateColumns: '1fr 1fr',
            gap: 12,
          }}
        >
          <div>
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Notes</label>
            <textarea
              value={form.notes}
              maxLength={300}
              onChange={(e) =>
                setForm((prev) => ({
                  ...prev,
                  notes: normalizeMultilineText(e.target.value).slice(0, 300),
                }))
              }
              rows={4}
              style={{
                width: '100%',
                resize: 'vertical',
                padding: '10px',
                border: '1px solid #d1d5db',
                borderRadius: 8,
                fontFamily: 'inherit',
                fontSize: 14,
                whiteSpace: 'pre-wrap',
              }}
            />
          </div>
          <div>
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Not Due</label>
            <textarea
              value={form.notDue}
              maxLength={300}
              onChange={(e) =>
                setForm((prev) => ({
                  ...prev,
                  notDue: normalizeMultilineText(e.target.value).slice(0, 300),
                }))
              }
              rows={4}
              style={{
                width: '100%',
                resize: 'vertical',
                padding: '10px',
                border: '1px solid #d1d5db',
                borderRadius: 8,
                fontFamily: 'inherit',
                fontSize: 14,
                whiteSpace: 'pre-wrap',
              }}
            />
          </div>
        </div>

        <p style={{ marginTop: 14, minHeight: 22, color: '#374151', fontSize: 14 }}>{status}</p>
        <div style={{ marginTop: 10 }}>
          <button
            type="button"
            disabled={docSubmitLock}
            onClick={handleSubmit}
            style={{
              width: '100%',
              height: 44,
              borderRadius: 8,
              border: '1px solid #9fe0f4',
              background: '#9fe0f4',
              color: '#fff',
              fontWeight: 700,
              cursor: docSubmitLock ? 'not-allowed' : 'pointer',
              opacity: docSubmitLock ? 0.55 : 1,
            }}
          >
            Submit
          </button>
        </div>
        </fieldset>
      </section>

      {submitFeedback && (
        <div
          role="dialog"
          aria-modal="true"
          aria-labelledby="submit-feedback-title"
          onClick={() => setSubmitFeedback(null)}
          style={{
            position: 'fixed',
            inset: 0,
            zIndex: 1000,
            display: 'grid',
            placeItems: 'center',
            background: 'rgba(15, 23, 42, 0.45)',
            padding: 24,
          }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{
              width: '100%',
              maxWidth: 420,
              background: '#fff',
              borderRadius: 16,
              padding: '32px 28px 24px',
              boxShadow: '0 20px 50px rgba(15, 23, 42, 0.25)',
              textAlign: 'center',
            }}
          >
            <div
              aria-hidden
              style={{
                width: 64,
                height: 64,
                margin: '0 auto 16px',
                borderRadius: '50%',
                display: 'grid',
                placeItems: 'center',
                fontSize: 32,
                fontWeight: 700,
                color: '#fff',
                background: submitFeedback.type === 'success' ? '#9fe0f4' : '#dc2626',
              }}
            >
              {submitFeedback.type === 'success' ? '✓' : '!'}
            </div>
            <h2
              id="submit-feedback-title"
              style={{
                margin: '0 0 8px',
                fontSize: 22,
                fontWeight: 800,
                color: submitFeedback.type === 'success' ? '#9fe0f4' : '#991b1b',
              }}
            >
              {submitFeedback.message}
            </h2>
            <button
              type="button"
              onClick={() => setSubmitFeedback(null)}
              style={{
                width: '100%',
                height: 44,
                borderRadius: 8,
                border: 'none',
                background: submitFeedback.type === 'success' ? '#9fe0f4' : '#dc2626',
                color: '#fff',
                fontWeight: 700,
                fontSize: 15,
                cursor: 'pointer',
              }}
            >
              OK
            </button>
          </div>
        </div>
      )}
    </main>
  );
}
