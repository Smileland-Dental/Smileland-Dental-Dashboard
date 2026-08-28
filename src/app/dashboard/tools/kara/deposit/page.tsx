'use client'

import React, { useEffect, useMemo, useState } from 'react';
import { collection, deleteDoc, doc, getDoc, getDocs, query, setDoc, where } from 'firebase/firestore';
import { onAuthStateChanged } from 'firebase/auth';
import { db, auth } from '@/lib/firebase.config';

const WORK_OFFICE_OPTIONS = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

const COLUMNS = ['Date', 'Dentical', 'Insurance Check', 'Personal Check', 'Credit Card', 'Ins. Electronic', 'Care Credit', 'Cash', 'Totals'] as const;
const VALUE_KEYS = ['d', 'ic', 'pc', 'cc', 'ie', 'cac', 'c'] as const;
const DATA_ROW_COUNT = 6;

type ValueKey = (typeof VALUE_KEYS)[number];

type TableRow = {
  date: string;
} & Record<ValueKey, string>;

function emptyRow(): TableRow {
  return {
    date: '',
    d: '',
    ic: '',
    pc: '',
    cc: '',
    ie: '',
    cac: '',
    c: '',
  };
}

function toCents(value: string): number {
  const n = Number(value);
  if (!Number.isFinite(n)) return 0;
  return Math.round(n * 100);
}

function formatCents(cents: number): string {
  const sign = cents < 0 ? '-' : '';
  const abs = Math.abs(cents);
  const whole = Math.floor(abs / 100);
  const frac = String(abs % 100).padStart(2, '0');
  return `${sign}${whole}.${frac}`;
}

type SavedTotals = Record<ValueKey, string> & {
  grandTotal: string;
};

type SavedSubmission = {
  id: string;
  date: string;
  office: string;
  rows: TableRow[];
  submittedAt: string;
  totals: SavedTotals;
};

function yearMonthFromDate(date: string): string {
  const matched = /^(\d{4})-(\d{2})/.exec(date || '');
  return matched ? `${matched[1]}-${matched[2]}` : '';
}

function californiaSubmittedAt(now = new Date()): string {
  const dtf = new Intl.DateTimeFormat('en-US', {
    timeZone: 'America/Los_Angeles',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hourCycle: 'h23',
  });
  const parts = Object.fromEntries(
    dtf
      .formatToParts(now)
      .filter((p) => p.type !== 'literal')
      .map((p) => [p.type, p.value])
  );
  const hour = parts.hour === '24' ? '00' : parts.hour;
  const asUtc = Date.UTC(
    Number(parts.year),
    Number(parts.month) - 1,
    Number(parts.day),
    Number(hour),
    Number(parts.minute),
    Number(parts.second)
  );
  const offsetMin = Math.round((asUtc - now.getTime()) / 60000);
  const sign = offsetMin >= 0 ? '+' : '-';
  const abs = Math.abs(offsetMin);
  const offsetHours = String(Math.floor(abs / 60)).padStart(2, '0');
  const offsetMinutes = String(abs % 60).padStart(2, '0');
  return `${parts.year}-${parts.month}-${parts.day}T${hour}:${parts.minute}:${parts.second}${sign}${offsetHours}:${offsetMinutes}`;
}

function totalsFromRows(tableRows: TableRow[]): SavedTotals {
  const cents = {} as Record<ValueKey, number>;
  for (const key of VALUE_KEYS) {
    cents[key] = tableRows.reduce((sum, row) => sum + toCents(row[key]), 0);
  }
  const grand = VALUE_KEYS.reduce((sum, key) => sum + cents[key], 0);
  return {
    d: formatCents(cents.d),
    ic: formatCents(cents.ic),
    pc: formatCents(cents.pc),
    cc: formatCents(cents.cc),
    ie: formatCents(cents.ie),
    cac: formatCents(cents.cac),
    c: formatCents(cents.c),
    grandTotal: formatCents(grand),
  };
}

function normalizeTotals(raw: any, tableRows: TableRow[]): SavedTotals {
  const source = raw && typeof raw === 'object' ? raw : {};
  const hasAny = VALUE_KEYS.some((key) => typeof source[key] === 'string' && source[key] !== '');
  if (!hasAny) return totalsFromRows(tableRows);
  return {
    d: typeof source.d === 'string' ? source.d : '',
    ic: typeof source.ic === 'string' ? source.ic : '',
    pc: typeof source.pc === 'string' ? source.pc : '',
    cc: typeof source.cc === 'string' ? source.cc : '',
    ie: typeof source.ie === 'string' ? source.ie : '',
    cac: typeof source.cac === 'string' ? source.cac : '',
    c: typeof source.c === 'string' ? source.c : '',
    grandTotal: typeof source.grandTotal === 'string' ? source.grandTotal : '',
  };
}

function normalizeRows(raw: any): TableRow[] {
  const source = Array.isArray(raw) ? raw : [];
  return Array.from({ length: DATA_ROW_COUNT }, (_, i) => {
    const row = source[i] || {};
    return {
      date: typeof row.date === 'string' ? row.date : '',      
      d: typeof row.d === 'string' ? row.d : '',
      ic: typeof row.ic === 'string' ? row.ic : '',
      pc: typeof row.pc === 'string' ? row.pc : '',
      cc: typeof row.cc === 'string' ? row.cc : '',
      ie: typeof row.ie === 'string' ? row.ie : '',
      cac: typeof row.cac === 'string' ? row.cac : '',
      c: typeof row.c === 'string' ? row.c : '',
    };
  });
}

export default function Deposit() {
  const [headerDate, setHeaderDate] = useState(() =>
    new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' })
  );
  const [office, setOffice] = useState('');
  const [rows, setRows] = useState<TableRow[]>(() =>
    Array.from({ length: DATA_ROW_COUNT }, emptyRow)
  );
  const [submitStatus, setSubmitStatus] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isDeleting, setIsDeleting] = useState(false);
  const [submissions, setSubmissions] = useState<SavedSubmission[]>([]);
  const [selectedId, setSelectedId] = useState('');
  const [filterYear, setFilterYear] = useState('');
  const [filterMonth, setFilterMonth] = useState('');
  const [isMonthlyView, setIsMonthlyView] = useState(false);
  const [isYearlyView, setIsYearlyView] = useState(false);
  const [pageReady, setPageReady] = useState(false);

  const loadSubmissions = async (officeName: string) => {
    if (!officeName) {
      setSubmissions([]);
      setSelectedId('');
      setIsMonthlyView(false);
      setIsYearlyView(false);
      return;
    }

    try {
      const snap = await getDocs(query(collection(db, 'deposit'), where('office', '==', officeName)));
      const list = snap.docs
        .map((item) => {
          const data = item.data();
          const rows = normalizeRows(data.rows);
          return {
            id: item.id,
            date: typeof data.date === 'string' ? data.date : item.id.split('_')[0],
            office: typeof data.office === 'string' ? data.office : officeName,
            rows,
            submittedAt: typeof data.submittedAt === 'string' ? data.submittedAt : '',
            totals: normalizeTotals(data.totals, rows),
          };
        })
        .sort((a, b) => {
          const dateCmp = (b.date || '').localeCompare(a.date || '');
          if (dateCmp !== 0) return dateCmp;
          return (b.submittedAt || '').localeCompare(a.submittedAt || '');
        });
      setSubmissions(list);
    } catch {
      setSubmissions([]);
    }
  };

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
    loadSubmissions(office);
  }, [office]);

  const availableYears = useMemo(() => {
    const values = new Set<string>();
    for (const item of submissions) {
      const ym = yearMonthFromDate(item.date);
      if (ym) values.add(ym.slice(0, 4));
    }
    return Array.from(values).sort((a, b) => b.localeCompare(a));
  }, [submissions]);

  const availableMonths = useMemo(() => {
    const values = new Set<string>();
    for (const item of submissions) {
      const ym = yearMonthFromDate(item.date);
      if (!ym) continue;
      const [year, month] = ym.split('-');
      if (year === filterYear && month) values.add(month);
    }
    return Array.from(values).sort((a, b) => b.localeCompare(a));
  }, [submissions, filterYear]);

  const filteredSubmissions = useMemo(() => {
    if (!filterYear || !filterMonth) return [];
    const ym = `${filterYear}-${filterMonth}`;
    return submissions.filter((item) => yearMonthFromDate(item.date) === ym);
  }, [submissions, filterYear, filterMonth]);

  useEffect(() => {
    if (availableYears.length === 0) {
      if (filterYear) setFilterYear('');
      return;
    }
    if (!availableYears.includes(filterYear)) {
      setFilterYear(availableYears[0]);
    }
  }, [availableYears, filterYear]);

  useEffect(() => {
    if (availableMonths.length === 0) {
      if (filterMonth) setFilterMonth('');
      return;
    }
    if (!availableMonths.includes(filterMonth)) {
      setFilterMonth(availableMonths[0]);
    }
  }, [availableMonths, filterMonth]);

  const rowTotals = useMemo(
    () => rows.map((row) => VALUE_KEYS.reduce((sum, key) => sum + toCents(row[key]), 0)),
    [rows]
  );

  const columnTotals = useMemo(() => {
    const totals = {} as Record<ValueKey, number>;
    for (const key of VALUE_KEYS) {
      totals[key] = rows.reduce((sum, row) => sum + toCents(row[key]), 0);
    }
    return totals;
  }, [rows]);

  const grandTotal = useMemo(
    () => rowTotals.reduce((sum, n) => sum + n, 0),
    [rowTotals]
  );

  const monthlyColumnTotals = useMemo(() => {
    const totals = {} as Record<ValueKey, number>;
    for (const key of VALUE_KEYS) {
      totals[key] = filteredSubmissions.reduce((sum, item) => sum + toCents(item.totals[key]), 0);
    }
    return totals;
  }, [filteredSubmissions]);

  const monthlyGrandTotal = useMemo(
    () =>
      filteredSubmissions.reduce((sum, item) => {
        if (item.totals.grandTotal) return sum + toCents(item.totals.grandTotal);
        return sum + VALUE_KEYS.reduce((inner, key) => inner + toCents(item.totals[key]), 0);
      }, 0),
    [filteredSubmissions]
  );

  const yearlyRows = useMemo(() => {
    if (!filterYear) return [];
    const byMonth = new Map<string, Record<ValueKey, number> & { grand: number }>();
    for (const item of submissions) {
      const ym = yearMonthFromDate(item.date);
      if (!ym || ym.slice(0, 4) !== filterYear) continue;
      const month = ym.slice(5, 7);
      const current = byMonth.get(month) || { d: 0, ic: 0, pc: 0, cc: 0, ie: 0, cac: 0, c: 0, grand: 0 };
      for (const key of VALUE_KEYS) {
        current[key] += toCents(item.totals[key]);
      }
      current.grand += item.totals.grandTotal
        ? toCents(item.totals.grandTotal)
        : VALUE_KEYS.reduce((sum, key) => sum + toCents(item.totals[key]), 0);
      byMonth.set(month, current);
    }
    return Array.from(byMonth.entries())
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([month, cents]) => ({
        month,
        totals: {          
          d: formatCents(cents.d),
          ic: formatCents(cents.ic),
          pc: formatCents(cents.pc),
          cc: formatCents(cents.cc),
          ie: formatCents(cents.ie),
          cac: formatCents(cents.cac),
          c: formatCents(cents.c),
          grandTotal: formatCents(cents.grand),
        } as SavedTotals,
      }));
  }, [submissions, filterYear]);

  const yearlyColumnTotals = useMemo(() => {
    const totals = {} as Record<ValueKey, number>;
    for (const key of VALUE_KEYS) {
      totals[key] = yearlyRows.reduce((sum, item) => sum + toCents(item.totals[key]), 0);
    }
    return totals;
  }, [yearlyRows]);

  const yearlyGrandTotal = useMemo(
    () => yearlyRows.reduce((sum, item) => sum + toCents(item.totals.grandTotal), 0),
    [yearlyRows]
  );

  const updateCell = (index: number, field: keyof TableRow, value: string) => {
    setRows((prev) =>
      prev.map((row, i) => (i === index ? { ...row, [field]: value } : row))
    );
  };

  const handleHeaderDateChange = (nextDate: string) => {
    setHeaderDate(nextDate);
    if (isMonthlyView || isYearlyView) {
      setIsMonthlyView(false);
      setIsYearlyView(false);
      setRows(Array.from({ length: DATA_ROW_COUNT }, emptyRow));
      return;
    }
    if (!selectedId) return;
    setSelectedId('');
    setRows(Array.from({ length: DATA_ROW_COUNT }, emptyRow));
  };

  const handleSubmit = async () => {
    if (!headerDate || !office) {
      alert('Please fill in Date and Office.');
      return;
    }

    const submittedAt = californiaSubmittedAt();
    const docId = `${headerDate}_${office}`.replace(/[\/\s]/g, '_');
    setIsSubmitting(true);
    setSubmitStatus('');

    try {
      await setDoc(doc(db, 'deposit', docId), {
        date: headerDate,
        office,
        submittedAt,
        rows,
        totals: {          
          d: formatCents(columnTotals.d),
          ic: formatCents(columnTotals.ic),
          pc: formatCents(columnTotals.pc),
          cc: formatCents(columnTotals.cc),
          ie: formatCents(columnTotals.ie),
          cac: formatCents(columnTotals.cac),
          c: formatCents(columnTotals.c),
          grandTotal: formatCents(grandTotal),
        },
      });
      setRows(Array.from({ length: DATA_ROW_COUNT }, emptyRow));
      setHeaderDate(new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' }));
      setSelectedId('');
      await loadSubmissions(office);
      setSubmitStatus('Submitted');
      setTimeout(() => setSubmitStatus(''), 2000);
    } catch {
      setSubmitStatus('Submit failed');
      alert('Submit failed. Please try again.');
    } finally {
      setIsSubmitting(false);
    }
  };

  const openSubmission = (item: SavedSubmission) => {
    setIsMonthlyView(false);
    setIsYearlyView(false);
    setSelectedId(item.id);
    setHeaderDate(item.date || '');
    setOffice(item.office);
    setRows(item.rows);
  };

  const openMonthlyView = () => {
    setIsMonthlyView(true);
    setIsYearlyView(false);
    setSelectedId('');
  };

  const openYearlyView = () => {
    setIsYearlyView(true);
    setIsMonthlyView(false);
    setSelectedId('');
  };

  const handleDelete = async () => {
    if (!selectedId) {
      alert('Please select a previous submission to delete.');
      return;
    }
    if (!confirm('Delete this submission?')) return;

    setIsDeleting(true);
    try {
      await deleteDoc(doc(db, 'deposit', selectedId));
      setRows(Array.from({ length: DATA_ROW_COUNT }, emptyRow));
      setHeaderDate(new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' }));
      setSelectedId('');
      await loadSubmissions(office);
    } catch {
      alert('Delete failed. Please try again.');
    } finally {
      setIsDeleting(false);
    }
  };

  const bodyStyle: React.CSSProperties = {
    padding: '20px',
    background: 'linear-gradient(to bottom, #ffffff, #f5f7fa)',
    minHeight: '100vh',
  };

  const containerStyle: React.CSSProperties = {
    maxWidth: '1400px',
    margin: '0 auto',
    backgroundColor: '#ffffff',
    padding: '30px',
    borderRadius: '12px',
    boxShadow: '0 8px 24px rgba(15, 23, 42, 0.06)',
    fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
    color: '#3b4252',
  };

  const cellStyle: React.CSSProperties = {
    padding: '10px 8px',
    textAlign: 'center',
    border: '1px solid #e6e8eb',
  };

  const inputStyle: React.CSSProperties = {
    padding: '8px 10px',
    border: '1px solid #e6e8eb',
    borderRadius: '4px',
    fontSize: '1em',
    backgroundColor: '#ffffff',
    color: '#3b4252',
    width: '100%',
    boxSizing: 'border-box',
    textAlign: 'center',
  };

  const headerInputStyle: React.CSSProperties = {
    ...inputStyle,
    textAlign: 'left',
  };

  const labelStyle: React.CSSProperties = {
    display: 'block',
    marginBottom: '5px',
    fontWeight: 'bold',
  };

  const buttonStyle: React.CSSProperties = {
    backgroundColor: '#7b8794',
    color: 'white',
    border: 'none',
    padding: '12px 24px',
    borderRadius: '6px',
    fontSize: '1em',
    fontWeight: 'bold',
    cursor: isSubmitting ? 'not-allowed' : 'pointer',
    opacity: isSubmitting ? 0.7 : 1,
  };

  const isViewingSubmission = Boolean(selectedId);
  const isReadOnly = isViewingSubmission || isMonthlyView || isYearlyView;

  const readOnlyCellStyle: React.CSSProperties = {
    ...inputStyle,
    backgroundColor: '#f8f9fb',
    cursor: 'default',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    minHeight: '38px',
  };

  if (!pageReady) {
    return null;
  }

  return (
    <div style={bodyStyle}>
      <div style={containerStyle}>
        <h1
          style={{
            color: '#4b5563',
            textAlign: 'center',
            marginBottom: '24px',
            fontSize: '2.5rem',
            fontWeight: 'bold',
          }}
        >
          Weekly Deposits
        </h1>

        <div style={{ display: 'flex', gap: '24px', alignItems: 'flex-start' }}>
        <aside
          style={{
            width: '240px',
            flexShrink: 0,
            backgroundColor: '#f8f9fb',
            border: '1px solid #e6e8eb',
            borderRadius: '8px',
            padding: '16px',
            minHeight: '320px',
          }}
        >

          {!office ? (
            <div style={{ fontSize: '14px', color: '#8b93a0' }}>Select an office to see submissions.</div>
          ) : submissions.length === 0 ? (
            <div style={{ fontSize: '14px', color: '#8b93a0' }}>No submissions yet.</div>
          ) : (
            <>
            <div style={{ display: 'flex', gap: '8px', marginBottom: '12px' }}>
              <div style={{ flex: 1, minWidth: 0 }}>
                <label style={{ ...labelStyle, fontSize: '13px' }}>Year</label>
                <select
                  value={filterYear}
                  onChange={(e) => setFilterYear(e.target.value)}
                  style={headerInputStyle}
                >
                  {availableYears.map((year) => (
                    <option key={year} value={year}>{year}</option>
                  ))}
                </select>
              </div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <label style={{ ...labelStyle, fontSize: '13px' }}>Month</label>
                <select
                  value={filterMonth}
                  onChange={(e) => setFilterMonth(e.target.value)}
                  style={headerInputStyle}
                >
                  {availableMonths.map((month) => (
                    <option key={month} value={month}>{month}</option>
                  ))}
                </select>
              </div>
            </div>
            <button
              onClick={openYearlyView}
              style={{
                width: '100%',
                textAlign: 'center',
                padding: '10px 12px',
                borderRadius: '6px',
                border: isYearlyView ? '2px solid #c5ccd6' : '1px solid #e6e8eb',
                backgroundColor: isYearlyView ? '#f0f2f5' : '#ffffff',
                color: '#3b4252',
                fontWeight: 'bold',
                cursor: 'pointer',
                marginBottom: '8px',
              }}
            >
              Yearly
            </button>
            <button
              onClick={openMonthlyView}
              style={{
                width: '100%',
                textAlign: 'center',
                padding: '10px 12px',
                borderRadius: '6px',
                border: isMonthlyView ? '2px solid #c5ccd6' : '1px solid #e6e8eb',
                backgroundColor: isMonthlyView ? '#f0f2f5' : '#ffffff',
                color: '#3b4252',
                fontWeight: 'bold',
                cursor: 'pointer',
                marginBottom: '8px',
              }}
            >
              Monthly
            </button>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
              {filteredSubmissions.length === 0 ? (
                <div style={{ fontSize: '14px', color: '#8b93a0' }}>No submissions in this month.</div>
              ) : filteredSubmissions.map((item) => (
                <button
                  key={item.id}
                  onClick={() => openSubmission(item)}
                  style={{
                    textAlign: 'left',
                    padding: '10px 12px',
                    borderRadius: '6px',
                    border: selectedId === item.id ? '2px solid #c5ccd6' : '1px solid #e6e8eb',
                    backgroundColor: selectedId === item.id ? '#f0f2f5' : '#ffffff',
                    color: '#3b4252',
                    fontWeight: selectedId === item.id ? 'bold' : 'normal',
                    cursor: 'pointer',
                  }}
                >
                  {item.date || item.id}
                </button>
              ))}
            </div>
            </>
          )}
        </aside>

        <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap', marginBottom: '24px' }}>
          <div style={{ flex: '1', minWidth: '200px' }}>
            <label style={labelStyle}>Office:</label>
            <select
              value={office}
              onChange={(e) => setOffice(e.target.value)}
              style={headerInputStyle}
            >
              <option value="">Select Office</option>
              {WORK_OFFICE_OPTIONS.map((option) => (
                <option key={option} value={option}>
                  {option}
                </option>
              ))}
            </select>
          </div>
          <div style={{ flex: '1', minWidth: '200px' }}>
            <label style={labelStyle}>Date:</label>
            <input
              type="date"
              value={headerDate}
              onChange={(e) => handleHeaderDateChange(e.target.value)}
              style={headerInputStyle}
            />
          </div>
        </div>

        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead style={{ backgroundColor: '#f8f9fb', color: '#3b4252' }}>
              <tr>
                {COLUMNS.map((col) => (
                  <th
                    key={col}
                    style={{
                      ...cellStyle,
                      border: '1px solid #e6e8eb',
                      minWidth: col === 'Date' ? '150px' : '90px',
                    }}
                  >
                    {col === 'Date' && isYearlyView ? 'Month' : col === 'Date' && isMonthlyView ? 'Submitted Date' : col}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {isYearlyView ? (
                <>
                  {yearlyRows.length === 0 ? (
                    <tr>
                      <td colSpan={COLUMNS.length} style={{ ...cellStyle, color: '#8b93a0' }}>
                        No submissions in this year.
                      </td>
                    </tr>
                  ) : (
                    yearlyRows.map((item, index) => (
                      <tr key={item.month} style={{ backgroundColor: index % 2 === 0 ? '#fbfcfd' : '#ffffff' }}>
                        <td style={cellStyle}>
                          <div style={readOnlyCellStyle}>{item.month}</div>
                        </td>
                        {VALUE_KEYS.map((key) => (
                          <td key={key} style={cellStyle}>
                            <div style={readOnlyCellStyle}>
                              {item.totals[key] === '' ? '' : `$${item.totals[key]}`}
                            </div>
                          </td>
                        ))}
                        <td style={{ ...cellStyle, fontWeight: 'bold' }}>${item.totals.grandTotal}</td>
                      </tr>
                    ))
                  )}
                  <tr style={{ backgroundColor: '#f5f7fa', fontWeight: 'bold' }}>
                    <td style={cellStyle}>Total</td>
                    {VALUE_KEYS.map((key) => (
                      <td key={key} style={cellStyle}>
                        ${formatCents(yearlyColumnTotals[key])}
                      </td>
                    ))}
                    <td style={cellStyle}>${formatCents(yearlyGrandTotal)}</td>
                  </tr>
                </>
              ) : isMonthlyView ? (
                <>
                  {filteredSubmissions.length === 0 ? (
                    <tr>
                      <td colSpan={COLUMNS.length} style={{ ...cellStyle, color: '#8b93a0' }}>
                        No submissions in this month.
                      </td>
                    </tr>
                  ) : (
                    filteredSubmissions.map((item, index) => {
                      const rowGrand = item.totals.grandTotal
                        || formatCents(VALUE_KEYS.reduce((sum, key) => sum + toCents(item.totals[key]), 0));
                      return (
                        <tr key={item.id} style={{ backgroundColor: index % 2 === 0 ? '#fbfcfd' : '#ffffff' }}>
                          <td style={cellStyle}>
                            <div style={readOnlyCellStyle}>{item.date}</div>
                          </td>
                          {VALUE_KEYS.map((key) => (
                            <td key={key} style={cellStyle}>
                              <div style={readOnlyCellStyle}>
                                {item.totals[key] === '' ? '' : `$${item.totals[key]}`}
                              </div>
                            </td>
                          ))}
                          <td style={{ ...cellStyle, fontWeight: 'bold' }}>${rowGrand}</td>
                        </tr>
                      );
                    })
                  )}
                  <tr style={{ backgroundColor: '#f5f7fa', fontWeight: 'bold' }}>
                    <td style={cellStyle}>Total</td>
                    {VALUE_KEYS.map((key) => (
                      <td key={key} style={cellStyle}>
                        ${formatCents(monthlyColumnTotals[key])}
                      </td>
                    ))}
                    <td style={cellStyle}>${formatCents(monthlyGrandTotal)}</td>
                  </tr>
                </>
              ) : (
                <>
              {rows.map((row, index) => (
                <tr key={index} style={{ backgroundColor: index % 2 === 0 ? '#fbfcfd' : '#ffffff' }}>
                  <td style={cellStyle}>
                    {isReadOnly ? (
                      <div style={readOnlyCellStyle}>{row.date}</div>
                    ) : (
                      <input
                        type="date"
                        value={row.date}
                        onChange={(e) => updateCell(index, 'date', e.target.value)}
                        style={inputStyle}
                      />
                    )}
                  </td>
                  {VALUE_KEYS.map((key) => (
                    <td key={key} style={cellStyle}>
                      {isReadOnly ? (
                        <div style={readOnlyCellStyle}>
                          {row[key] === '' ? '' : `$${row[key]}`}
                        </div>
                      ) : (
                        <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                          <span>$</span>
                          <input
                            type="text"
                            inputMode="decimal"
                            value={row[key]}
                            onChange={(e) => {
                              const v = e.target.value.replace(/[$,]/g, '');
                              if (v === '' || /^-?\d*\.?\d*$/.test(v)) {
                                updateCell(index, key, v);
                              }
                            }}
                            onBlur={() => {
                              if (row[key] === '') return;
                              updateCell(index, key, formatCents(toCents(row[key])));
                            }}
                            style={inputStyle}
                          />
                        </div>
                      )}
                    </td>
                  ))}
                  <td style={{ ...cellStyle, fontWeight: 'bold' }}>${formatCents(rowTotals[index])}</td>
                </tr>
              ))}
              <tr style={{ backgroundColor: '#f5f7fa', fontWeight: 'bold' }}>
                <td style={cellStyle}>Total</td>
                {VALUE_KEYS.map((key) => (
                  <td key={key} style={cellStyle}>
                    ${formatCents(columnTotals[key])}
                  </td>
                ))}
                <td style={cellStyle}>${formatCents(grandTotal)}</td>
              </tr>
                </>
              )}
            </tbody>
          </table>
        </div>

        <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', gap: '12px', marginTop: '24px' }}>
          <button onClick={handleSubmit} disabled={isSubmitting || isDeleting || isMonthlyView || isYearlyView} style={{
            ...buttonStyle,
            cursor: isSubmitting || isDeleting || isMonthlyView || isYearlyView ? 'not-allowed' : 'pointer',
            opacity: isSubmitting || isDeleting || isMonthlyView || isYearlyView ? 0.6 : 1,
          }}>
            {isSubmitting ? 'Submitting...' : 'Submit'}
          </button>
          <button
            onClick={handleDelete}
            disabled={!selectedId || isDeleting || isSubmitting}
            style={{
              ...buttonStyle,
              backgroundColor: '#dc3545',
              cursor: !selectedId || isDeleting || isSubmitting ? 'not-allowed' : 'pointer',
              opacity: !selectedId || isDeleting || isSubmitting ? 0.6 : 1,
            }}
          >
            {isDeleting ? 'Deleting...' : 'Delete'}
          </button>
          {submitStatus && (
            <span style={{ fontWeight: 'bold', color: submitStatus === 'Submitted' ? '#28a745' : '#dc3545' }}>
              {submitStatus}
            </span>
          )}
        </div>
        </div>
        </div>
      </div>
    </div>
  );
}
