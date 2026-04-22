'use client';

import React, { useCallback, useEffect, useState } from 'react';
import { db } from '@/lib/firebase.config';
import { collection, getDocs, query, where } from 'firebase/firestore';

const LOCATION_OPTIONS = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'] as const;

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

type ReasonTotals = {
  orangeJuice: number;
  paper: number;
  coffee: number;
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
  notes?: string;
  notDue?: string;
  tableRows?: TableRow[];
  sugarRows?: SugarRow[];
  reasonRows?: ReasonRow[];
  tableTotals?: Partial<TableTotals> & Record<string, unknown>;
  sugarTotals?: Partial<SugarTotals> & Record<string, unknown>;
  [key: string]: unknown;
};

/** Fields selectable in column dropdowns (Biller / report rows excluded). */
type ColumnFieldId =
  | ''
  | 'main.grandTotal'
  | 'main.coffeeSales'
  | 'main.salesWithoutCoffee'
  | 'main.paperAtOrangeJuice'
  | 'main.paperAtTea'
  | 'main.justPaper'
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
  | 'reason.orangeJuice'
  | 'reason.paper'
  | 'reason.coffee';

/** One slot picks this → table shows one column per fixed reason (OE, Pro, or CRA only). */
type ShortProcBundleId = 'bundle.spOE' | 'bundle.spPro' | 'bundle.spCRA';

type SlotValue = ColumnFieldId | ShortProcBundleId;

/** Same order/labels as p_page.tsx reason rows (match `reason` text when saving). */
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
      { id: 'main.grandTotal', label: 'Grand Total' },
      { id: 'main.coffeeSales', label: 'CRA Production' },
      { id: 'main.salesWithoutCoffee', label: 'Production W/Out CRA' },
      { id: 'main.paperAtOrangeJuice', label: 'Prophy @ OE' },
      { id: 'main.paperAtTea', label: 'Prophy @ TX' },
      { id: 'main.justPaper', label: 'Just Prophy' },
    ],
  },
  {
    label: 'Production totals',
    options: [
      { id: 'table.sales', label: 'Production' },
      { id: 'table.coffeeNew', label: 'CRA (New)' },
      { id: 'table.coffeeReturn', label: 'CRA (Return)' },
      { id: 'table.coffeeTotal', label: 'CRA Total' },
      { id: 'table.coffeeNo', label: 'CRA (Not Billable)' },
      { id: 'table.renderedCoffee', label: 'Rendered CRA' },
      { id: 'table.coffeeYes', label: 'CRA (Billable)' },
      { id: 'table.orangeJuiceNew', label: 'OE (NP)' },
      { id: 'table.orangeJuiceReturn', label: 'OE (RC)' },
      { id: 'table.orangeJuiceTotal', label: 'OE Total' },
    ],
  },
  {
    label: 'Sealant totals',
    options: [
      { id: 'sugar.sugar', label: 'Sealant' },
      { id: 'sugar.sugarGood', label: 'Sealant (Billable)' },
      { id: 'sugar.sugarBad', label: 'Sealant (Redo)' },
      { id: 'sugar.paper', label: 'Prophy' },
    ],
  },
  {
    label: 'Short Procedures (totals)',
    options: [
      { id: 'reason.orangeJuice', label: 'OE' },
      { id: 'reason.paper', label: 'Pro' },
      { id: 'reason.coffee', label: 'CRA' },
    ],
  },
];

const NUM_SLOTS = 12;

const EMPTY_COLUMN_SLOTS: SlotValue[] = Array.from({ length: NUM_SLOTS }, () => '' as SlotValue);

function isShortProcBundle(slot: SlotValue): slot is ShortProcBundleId {
  return slot === 'bundle.spOE' || slot === 'bundle.spPro' || slot === 'bundle.spCRA';
}

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
  minWidth: 104,
  maxWidth: 140,
  lineHeight: 1.25,
};

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

function getReasonTotalsFromRows(rows: ReasonRow[]): ReasonTotals {
  return {
    orangeJuice: rows.reduce((sum, row) => sum + parseNumber(row.orangeJuice), 0),
    paper: rows.reduce((sum, row) => sum + parseNumber(row.paper), 0),
    coffee: rows.reduce((sum, row) => sum + parseNumber(row.coffee), 0),
  };
}

function resolveTotals(doc: SimpleFormDoc): { table: TableTotals; sugar: SugarTotals; reason: ReasonTotals } {
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

  const reason =
    Array.isArray(doc.reasonRows) && doc.reasonRows.length > 0
      ? getReasonTotalsFromRows(doc.reasonRows)
      : getReasonTotalsFromRows([]);

  return { table, sugar, reason };
}

/** First column Day: weekday (English short, e.g. Mon). */
function formatDayColumn(dateStr: string): string {
  if (!dateStr || !/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return '—';
  const d = new Date(`${dateStr}T12:00:00`);
  if (Number.isNaN(d.getTime())) return '—';
  return new Intl.DateTimeFormat('en-US', { weekday: 'short' }).format(d);
}

/** Date column: stored YYYY-MM-DD → display as M/D/YYYY (no leading zeros). */
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
  totals: { table: TableTotals; sugar: SugarTotals; reason: ReasonTotals }
): string {
  if (!id) return '';
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
  if (id.startsWith('reason.')) {
    const key = id.replace('reason.', '') as keyof ReasonTotals;
    return String(totals.reason[key]);
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
          all.push({ id: docSnap.id, ...(docSnap.data() as object) } as SimpleFormDoc);
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
  };

  const hasProcedureSubHeader = columnSlots.some(isShortProcBundle);

  const thSelectStyle: React.CSSProperties = {
    width: '100%',
    minWidth: 120,
    height: 36,
    padding: '0 8px',
    border: '1px solid #d1d5db',
    borderRadius: 6,
    background: '#fff',
    fontSize: 13,
    fontWeight: 600,
  };

  const slotFieldSelect = (slotIndex: number, slotId: SlotValue) => (
    <select
      aria-label={`Column ${slotIndex + 1} field`}
      value={slotId}
      onChange={(e) => setSlot(slotIndex, e.target.value as SlotValue)}
      style={thSelectStyle}
    >
      <option value="">Select field…</option>
      {FIELD_GROUPS.map((g) => (
        <optgroup key={g.label} label={g.label}>
          {g.options.map((opt) => (
            <option key={opt.id} value={opt.id}>
              {opt.label}
            </option>
          ))}
        </optgroup>
      ))}
      <optgroup label="Short Procedures (bundles)">
        {SHORT_PROC_BUNDLE_OPTIONS.map((opt) => (
          <option key={opt.id} value={opt.id}>
            {opt.label}
          </option>
        ))}
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
      style={{
        minHeight: '100vh',
        display: 'grid',
        placeItems: 'start center',
        background: '#ffffff',
        padding: '48px 24px',
      }}
    >
      <section
        style={{
          width: '100%',
          maxWidth: 2600,
          background: '#fff',
          border: '1px solid #e5e7eb',
          borderRadius: 12,
          padding: 20,
          boxShadow: '0 8px 24px rgba(15, 23, 42, 0.06)',
        }}
      >
        <h1 style={{ margin: '0 0 8px', fontSize: 28, fontWeight: 800, color: '#111827' }}>
          Production Filter
        </h1>

        <div
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
              onChange={(e) => setLocation(e.target.value)}
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
              onChange={(e) => setMonth(e.target.value)}
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
          <div style={{ overflowX: 'auto' }}>
            <table
              style={{
                width: '100%',
                borderCollapse: 'collapse',
                border: '1px solid #e5e7eb',
                fontSize: 14,
              }}
            >
              <thead>
                <tr style={{ background: '#f3f4f6' }}>
                  <th
                    rowSpan={hasProcedureSubHeader ? 2 : 1}
                    style={{
                      ...cellStyle,
                      whiteSpace: 'nowrap',
                      minWidth: 72,
                    }}
                  >
                    Day
                  </th>
                  <th
                    rowSpan={hasProcedureSubHeader ? 2 : 1}
                    style={{
                      ...cellStyle,
                      whiteSpace: 'nowrap',
                      minWidth: 110,
                    }}
                  >
                    Date
                  </th>
                  {columnSlots.map((slotId, slotIndex) =>
                    isShortProcBundle(slotId) ? (
                      <th
                        key={`h-${slotIndex}`}
                        colSpan={BUNDLE_COLSPAN}
                        style={{ ...cellStyle, minWidth: 100 * BUNDLE_COLSPAN }}
                      >
                        {slotFieldSelect(slotIndex, slotId)}
                      </th>
                    ) : (
                      <th
                        key={`h-${slotIndex}`}
                        rowSpan={hasProcedureSubHeader ? 2 : 1}
                        style={{ ...cellStyle, minWidth: 140 }}
                      >
                        {slotFieldSelect(slotIndex, slotId)}
                      </th>
                    )
                  )}
                </tr>
                {hasProcedureSubHeader ? (
                  <tr style={{ background: '#f9fafb' }}>
                    {columnSlots.flatMap((slotId, slotIndex) =>
                      isShortProcBundle(slotId)
                        ? FIXED_REASON_OPTIONS.map((label, ri) => (
                            <th key={`h2-${slotIndex}-${ri}`} style={{ ...cellStyle, ...subHeaderThStyle }}>
                              {label}
                            </th>
                          ))
                        : []
                    )}
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
                      {columnSlots.flatMap((slotId, slotIndex) =>
                        isShortProcBundle(slotId)
                          ? FIXED_REASON_OPTIONS.map((label) => (
                              <td key={`c-${doc.id}-${slotIndex}-${label}`} style={cellStyle}>
                                {getReasonCellValue(doc, label, bundleToMetric(slotId))}
                              </td>
                            ))
                          : [
                              <td key={`c-${doc.id}-${slotIndex}`} style={cellStyle}>
                                {getFieldDisplay(doc, slotId as ColumnFieldId, totals)}
                              </td>,
                            ]
                      )}
                    </tr>
                  );
                })}
              </tbody>
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
              {rows.length} row{rows.length === 1 ? '' : 's'} · {NUM_SLOTS} column slots
            </span>
            <button
              type="button"
              onClick={() => setColumnSlots([...EMPTY_COLUMN_SLOTS])}
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

