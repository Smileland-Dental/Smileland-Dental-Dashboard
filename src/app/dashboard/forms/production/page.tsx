'use client';

import React, { useEffect, useRef, useState } from 'react';
import { db } from '@/lib/firebase.config';
import { doc, getDoc, serverTimestamp, setDoc } from 'firebase/firestore';

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
  { key: 'sales', label: 'Production' },
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
  { key: 'sugarGood', label: 'Sealant (Billable)' },
  { key: 'sugarBad', label: 'Sealant (Redo)' },
  { key: 'paper', label: 'Prophy' },
];

const REASON_COLUMNS: Array<{ key: keyof ReasonRow; label: string }> = [
  { key: 'reason', label: 'Short Procedures (Reasoning)' },
  { key: 'orangeJuice', label: 'OE' },
  { key: 'paper', label: 'Pro' },
  { key: 'coffee', label: 'CRA' },
];

const POSITION_OPTIONS = ['Doctor', 'RDA', 'DA'];
const COFFEE_POSITION_OPTIONS = ['Doctor'];
const FIXED_REASON_OPTIONS = ['Declined/DDP/Pt Left', 'Not Due/Freq', 'Medical Clearance', 'Furture TX/FMS', 'Courtesy, Not Billable', 'Mistakenly Done, Not Billable', 'Uncooperative/Re-eval', 'Not Documented', 'Age Limit/No teeth', 'Not Complete/Filled out'];
const LOCATION_OPTIONS = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];
const READONLY_FIELDS: Array<keyof TableRow> = ['coffeeTotal', 'orangeJuiceTotal'];
const NUMERIC_FIELDS: Array<keyof TableRow> = [
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
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
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

  return {
    ...row,
    coffeeTotal,
    orangeJuiceTotal,
  };
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

function getSugarTotals(rows: SugarRow[]): SugarTotals {
  return {
    sugar: rows.reduce((sum, row) => sum + parseNumber(row.sugar), 0),
    sugarGood: rows.reduce((sum, row) => sum + parseNumber(row.sugarGood), 0),
    sugarBad: rows.reduce((sum, row) => sum + parseNumber(row.sugarBad), 0),
    paper: rows.reduce((sum, row) => sum + parseNumber(row.paper), 0),
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

function getNowDateTimeString() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  const seconds = String(now.getSeconds()).padStart(2, '0');
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

function getDurationLabel(start: string, end: string): string {
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

function hasAnyDataToSave(
  form: SimpleFormData,
  reportRows: ReportRow[],
  tableRows: TableRow[],
  sugarRows: SugarRow[],
  reasonRows: ReasonRow[]
) {
  const hasMainFormValue =
    form.reasonIfLate.trim() !== '' ||
    form.submittedDateTime.trim() !== '' ||
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
    notes: '',
    notDue: '',
  });
  const [reportRows, setReportRows] = useState<ReportRow[]>([createEmptyReportRow()]);
  const [status, setStatus] = useState('');
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

  useEffect(() => {
    const hasCoffeeYesInput = tableRows.some((row) => row.coffeeYes.trim() !== '');
    const coffeeYesTotal = tableRows.reduce((sum, row) => sum + parseNumber(row.coffeeYes), 0);
    const calculatedCoffeeSales = hasCoffeeYesInput ? String(coffeeYesTotal * 61) : '';

    setForm((prev) =>
      prev.coffeeSales === calculatedCoffeeSales ? prev : { ...prev, coffeeSales: calculatedCoffeeSales }
    );
  }, [tableRows]);

  useEffect(() => {
    const bothEmpty = form.grandTotal.trim() === '' && form.coffeeSales.trim() === '';
    const calculatedSalesWithoutCoffee = bothEmpty
      ? ''
      : String(parseNumber(form.grandTotal) - parseNumber(form.coffeeSales));

    setForm((prev) =>
      prev.salesWithoutCoffee === calculatedSalesWithoutCoffee
        ? prev
        : { ...prev, salesWithoutCoffee: calculatedSalesWithoutCoffee }
    );
  }, [form.grandTotal, form.coffeeSales]);

  const getDocId = () => {
    const safeLocation = form.location.trim().replace(/\s+/g, '_');
    return `${form.date}_${safeLocation}`;
  };

  useEffect(() => {
    if (!form.date || !form.location.trim()) {
      return;
    }

    let cancelled = false;
    const loadByDateAndLocation = async () => {
      try {
        const snap = await getDoc(doc(db, 'simple-forms', getDocId()));
        if (cancelled) return;

        if (!snap.exists()) {
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
            notes: '',
            notDue: '',
          }));
          lastSavedKeyRef.current = '';
          return;
        }

        const data = snap.data() as any;
        const loadedReportRows =
          Array.isArray(data.reportRows) && data.reportRows.length > 0
            ? data.reportRows.map((row: Partial<ReportRow>) => ({ ...createEmptyReportRow(), ...row }))
            : [createEmptyReportRow()];
        const loadedTableRows =
          Array.isArray(data.tableRows) && data.tableRows.length > 0
            ? data.tableRows.map((row: Partial<TableRow>) => ({ ...createEmptyRow(), ...row }))
            : Array.from({ length: 4 }, () => createEmptyRow());
        const loadedSugarRows =
          Array.isArray(data.sugarRows) && data.sugarRows.length > 0
            ? data.sugarRows.map((row: Partial<SugarRow>) => ({ ...createEmptySugarRow(), ...row }))
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
          notes: String(data.notes ?? ''),
          notDue: String(data.notDue ?? ''),
        }));
        lastSavedKeyRef.current = '';
      } catch (error: any) {
        if (cancelled) return;
        setStatus(`불러오기 실패: ${error?.message || '알 수 없는 오류'}`);
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
    if (!hasAnyDataToSave(form, reportRows, tableRows, sugarRows, reasonRows)) {
      return;
    }

    const normalizedRows = tableRows.map(computeRow).filter(hasAnyTableRowValue);
    const normalizedSugarRows = sugarRows.filter(hasAnySugarRowValue);
    const saveKey = `${JSON.stringify(form)}_${JSON.stringify(reportRows)}_${JSON.stringify(normalizedRows)}_${JSON.stringify(normalizedSugarRows)}_${JSON.stringify(reasonRows)}`;
    if (saveKey === lastSavedKeyRef.current) {
      return;
    }

    const timer = setTimeout(async () => {
      try {
        const docId = getDocId();
        const firstReportRow = reportRows[0] ?? createEmptyReportRow();
        const tableTotals = getTableTotals(normalizedRows);
        const sugarTotals = getSugarTotals(normalizedSugarRows);
        const { submittedDateTime: _submittedDateTime, ...formWithoutSubmittedDateTime } = form;
        await setDoc(doc(db, 'simple-forms', docId), {
          ...formWithoutSubmittedDateTime,
          reportRows,
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
        });
        lastSavedKeyRef.current = saveKey;
        setStatus('Saved Successfully');
      } catch (error: any) {
        setStatus('Failed to Save');
      }
    }, 700);

    return () => clearTimeout(timer);
  }, [form, reportRows, tableRows, sugarRows, reasonRows]);

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
    return tableRows.reduce((sum, row) => sum + parseNumber(row[field]), 0);
  };

  const handleSugarCellChange = (rowIndex: number, field: keyof SugarRow, value: string) => {
    setSugarRows((prev) => {
      const next = [...prev];
      next[rowIndex] = { ...next[rowIndex], [field]: value };
      return next;
    });
  };

  const handleAddSugarRow = () => {
    setSugarRows((prev) => [...prev, createEmptySugarRow()]);
  };

  const getSugarColumnTotal = (field: keyof SugarRow) => {
    return sugarRows.reduce((sum, row) => sum + parseNumber(row[field]), 0);
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
      setStatus('제출하려면 Location을 먼저 선택해 주세요.');
      return;
    }

    if (!form.date) {
      setStatus('제출하려면 Date가 필요합니다.');
      return;
    }

    try {
      const docId = getDocId();
      const normalizedRows = tableRows.map(computeRow).filter(hasAnyTableRowValue);
      const normalizedSugarRows = sugarRows.filter(hasAnySugarRowValue);
      const firstReportRow = reportRows[0] ?? createEmptyReportRow();
      const tableTotals = getTableTotals(normalizedRows);
      const sugarTotals = getSugarTotals(normalizedSugarRows);
      const submittedDateTime = getNowDateTimeString();

      await setDoc(doc(db, 'simple-forms', docId), {
        ...form,
        submittedDateTime,
        reportRows,
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
      });

      lastSavedKeyRef.current = '';
      setForm({
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
        notes: '',
        notDue: '',
      });
      setReportRows([createEmptyReportRow()]);
      setTableRows(Array.from({ length: 4 }, () => createEmptyRow()));
      setSugarRows(Array.from({ length: 6 }, () => createEmptySugarRow()));
      setReasonRows(Array.from({ length: 10 }, (_, i) => createEmptyReasonRow(i)));
      setStatus('Submitted Successfully');
    } catch (error: any) {
      setStatus('Failed to Submit');
    }
  };

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
        <h1 style={{ margin: '0 0 16px', fontSize: 28, fontWeight: 800, color: '#111827' }}>Finalized Production</h1>
        <div
          style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(5, minmax(160px, 1fr))',
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
            <select
              value={form.location}
              onChange={(e) => setForm((prev) => ({ ...prev, location: e.target.value }))}
              style={{
                width: '100%',
                height: 40,
                padding: '0 10px',
                border: '1px solid #d1d5db',
                borderRadius: 8,
                background: '#fff',
              }}
            >
              <option value="">Select Office</option>
              {LOCATION_OPTIONS.map((location) => (
                <option key={location} value={location}>
                  {location}
                </option>
              ))}
            </select>
          </div>
          <div>
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Reason if Late</label>
            <input
              type="text"
              value={form.reasonIfLate}
              onChange={(e) => setForm((prev) => ({ ...prev, reasonIfLate: e.target.value }))}
              style={{
                width: '100%',
                height: 40,
                padding: '0 10px',
                border: '1px solid #d1d5db',
                borderRadius: 8,
              }}
            />
          </div>
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
              <input
                type="number"
                inputMode="numeric"
                value={row.chartCount}
                onChange={(e) => handleReportRowChange(rowIndex, 'chartCount', e.target.value)}
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
              <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>$ Amount</label>
              <input
                type="number"
                inputMode="decimal"
                step="0.01"
                value={row.amount}
                onChange={(e) => handleReportRowChange(rowIndex, 'amount', e.target.value)}
                style={{
                  width: '100%',
                  height: 40,
                  padding: '0 10px',
                  border: '1px solid #d1d5db',
                  borderRadius: 8,
                }}
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
            <input
              type="number"
              inputMode="decimal"
              value={form.grandTotal}
              onChange={(e) => setForm((prev) => ({ ...prev, grandTotal: e.target.value }))}
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
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>CRA Production</label>
            <input
              type="number"
              inputMode="decimal"
              value={form.coffeeSales}
              readOnly
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
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Production W/Out CRA</label>
            <input
              type="number"
              inputMode="decimal"
              value={form.salesWithoutCoffee}
              readOnly
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
        </div>
        <div
          style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(3, minmax(180px, 1fr))',
            gap: 12,
            marginBottom: 8,
          }}
        >
          <div>
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Prophy @ OE</label>
            <input
              type="number"
              inputMode="decimal"
              value={form.paperAtOrangeJuice}
              onChange={(e) => setForm((prev) => ({ ...prev, paperAtOrangeJuice: e.target.value }))}
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
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Prophy @ TX</label>
            <input
              type="number"
              inputMode="decimal"
              value={form.paperAtTea}
              onChange={(e) => setForm((prev) => ({ ...prev, paperAtTea: e.target.value }))}
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
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Just Prophy</label>
            <input
              type="number"
              inputMode="decimal"
              value={form.justPaper}
              onChange={(e) => setForm((prev) => ({ ...prev, justPaper: e.target.value }))}
              style={{
                width: '100%',
                height: 40,
                padding: '0 10px',
                border: '1px solid #d1d5db',
                borderRadius: 8,
              }}
            />
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
              {tableRows.map((row, rowIndex) => (
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
                          <option value="">Select Position</option>
                          {COFFEE_POSITION_OPTIONS.map((option) => (
                            <option key={option} value={option}>
                              {option}
                            </option>
                          ))}
                        </select>
                      ) : (
                        <input
                          type={NUMERIC_FIELDS.includes(column.key) ? 'number' : 'text'}
                          inputMode={NUMERIC_FIELDS.includes(column.key) ? 'numeric' : 'text'}
                          value={row[column.key]}
                          readOnly={READONLY_FIELDS.includes(column.key)}
                          onChange={(e) => handleCellChange(rowIndex, column.key, e.target.value)}
                          style={{
                            width: '100%',
                            height: 34,
                            border: '1px solid #d1d5db',
                            borderRadius: 6,
                            padding: '0 8px',
                            background: READONLY_FIELDS.includes(column.key) ? '#f3f4f6' : '#fff',
                            color: READONLY_FIELDS.includes(column.key) ? '#6b7280' : '#111827',
                          }}
                        />
                      )}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr style={{ background: '#f9fafb', fontWeight: 700 }}>
                {TABLE_COLUMNS.map((column) => {
                  let value = '';
                  if (column.key === 'position') value = 'Total';
                  else if (column.key === 'name') value = '-';
                  else if (TOTAL_TARGET_FIELDS.includes(column.key)) value = String(getColumnTotal(column.key));

                  return (
                    <td key={`total-${column.key}`} style={{ border: '1px solid #e5e7eb', padding: '8px' }}>
                      {value}
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
              {sugarRows.map((row, rowIndex) => (
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
                          <option value="">Select Position</option>
                          {POSITION_OPTIONS.map((option) => (
                            <option key={option} value={option}>
                              {option}
                            </option>
                          ))}
                        </select>
                      ) : (
                        <input
                          type={column.key === 'name' ? 'text' : 'number'}
                          inputMode={column.key === 'name' ? 'text' : 'numeric'}
                          value={row[column.key]}
                          onChange={(e) => handleSugarCellChange(rowIndex, column.key, e.target.value)}
                          style={{
                            width: '100%',
                            height: 34,
                            border: '1px solid #d1d5db',
                            borderRadius: 6,
                            padding: '0 8px',
                          }}
                        />
                      )}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr style={{ background: '#f9fafb', fontWeight: 700 }}>
                {SUGAR_COLUMNS.map((column) => {
                  let value = '';
                  if (column.key === 'position') value = 'Total';
                  else if (column.key === 'name') value = '-';
                  else if (SUGAR_TOTAL_TARGET_FIELDS.includes(column.key)) value = String(getSugarColumnTotal(column.key));

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
                      <input
                        type={column.key === 'reason' ? 'text' : 'number'}
                        inputMode={column.key === 'reason' ? 'text' : 'numeric'}
                        value={row[column.key]}
                        readOnly={column.key === 'reason'}
                        onChange={(e) => handleReasonCellChange(rowIndex, column.key, e.target.value)}
                        style={{
                          width: '100%',
                          height: 34,
                          border: '1px solid #d1d5db',
                          borderRadius: 6,
                          padding: '0 8px',
                          background: column.key === 'reason' ? '#f3f4f6' : '#fff',
                          color: column.key === 'reason' ? '#6b7280' : '#111827',
                        }}
                      />
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
              onChange={(e) => setForm((prev) => ({ ...prev, notes: e.target.value }))}
              rows={4}
              style={{
                width: '100%',
                resize: 'vertical',
                padding: '10px',
                border: '1px solid #d1d5db',
                borderRadius: 8,
                fontFamily: 'inherit',
                fontSize: 14,
              }}
            />
          </div>
          <div>
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600 }}>Not Due</label>
            <textarea
              value={form.notDue}
              onChange={(e) => setForm((prev) => ({ ...prev, notDue: e.target.value }))}
              rows={4}
              style={{
                width: '100%',
                resize: 'vertical',
                padding: '10px',
                border: '1px solid #d1d5db',
                borderRadius: 8,
                fontFamily: 'inherit',
                fontSize: 14,
              }}
            />
          </div>
        </div>

        <p style={{ marginTop: 14, minHeight: 22, color: '#374151', fontSize: 14 }}>{status}</p>
        <div style={{ marginTop: 10 }}>
          <button
            type="button"
            onClick={handleSubmit}
            style={{
              width: '100%',
              height: 44,
              borderRadius: 8,
              border: '1px solid #16a34a',
              background: '#16a34a',
              color: '#fff',
              fontWeight: 700,
              cursor: 'pointer',
            }}
          >
            Submit
          </button>
        </div>
      </section>
    </main>
  );
}

