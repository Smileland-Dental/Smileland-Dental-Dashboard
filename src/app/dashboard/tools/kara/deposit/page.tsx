'use client'

import React, { useEffect, useMemo, useState } from 'react';
import { collection, deleteDoc, doc, getDoc, getDocs, query, setDoc, where } from 'firebase/firestore';
import { onAuthStateChanged } from 'firebase/auth';
import { db, auth } from '@/lib/firebase.config';

const WORK_OFFICE_OPTIONS = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

const COLUMNS = ['Date', 'Insurance Check', 'Personal Check', 'Credit Card', 'Ins. Electronic', 'Care Credit', 'Cash', 'Totals'] as const;
const VALUE_KEYS = ['ic', 'pc', 'cc', 'ie', 'cac', 'c'] as const;
const DATA_ROW_COUNT = 6;

type ValueKey = (typeof VALUE_KEYS)[number];

type TableRow = {
  date: string;
} & Record<ValueKey, string>;

function emptyRow(): TableRow {
  return {
    date: '',
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

type SavedSubmission = {
  id: string;
  date: string;
  office: string;
  rows: TableRow[];
  submittedAt: string;
};

function yearMonthFromDate(date: string): string {
  const matched = /^(\d{4})-(\d{2})/.exec(date || '');
  return matched ? `${matched[1]}-${matched[2]}` : '';
}

function submissionStamp(now = new Date()): { idTime: string; iso: string } {
  const time = now
    .toLocaleTimeString('en-GB', { timeZone: 'America/Los_Angeles', hour12: false })
    .replace(/:/g, '');
  return { idTime: time, iso: now.toISOString() };
}

function formatSubmittedAt(iso: string): string {
  if (!iso) return '';
  const parsed = new Date(iso);
  if (!Number.isFinite(parsed.getTime())) return '';
  return parsed.toLocaleTimeString('en-GB', { timeZone: 'America/Los_Angeles', hour12: false });
}

function normalizeRows(raw: any): TableRow[] {
  const source = Array.isArray(raw) ? raw : [];
  return Array.from({ length: DATA_ROW_COUNT }, (_, i) => {
    const row = source[i] || {};
    return {
      date: typeof row.date === 'string' ? row.date : '',
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
  const [userOfficesOptions, setUserOfficesOptions] = useState<string[]>([]);
  const [rows, setRows] = useState<TableRow[]>(() =>
    Array.from({ length: DATA_ROW_COUNT }, emptyRow)
  );
  const [submitStatus, setSubmitStatus] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isDeleting, setIsDeleting] = useState(false);
  const [submissions, setSubmissions] = useState<SavedSubmission[]>([]);
  const [selectedId, setSelectedId] = useState('');
  const [filterYearMonth, setFilterYearMonth] = useState('');
  const [pageReady, setPageReady] = useState(false);

  const loadSubmissions = async (officeName: string) => {
    if (!officeName) {
      setSubmissions([]);
      setSelectedId('');
      return;
    }

    try {
      const snap = await getDocs(query(collection(db, 'deposit'), where('office', '==', officeName)));
      const list = snap.docs
        .map((item) => {
          const data = item.data();
          return {
            id: item.id,
            date: typeof data.date === 'string' ? data.date : item.id.split('_')[0],
            office: typeof data.office === 'string' ? data.office : officeName,
            rows: normalizeRows(data.rows),
            submittedAt: typeof data.submittedAt === 'string' ? data.submittedAt : '',
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
          userData?.role !== 'Manager' &&
          userData?.role !== 'HR' &&
          userData?.role !== 'Director'
        ) {
          goHome();
          return;
        }

        if (userData?.offices) {
          const officesArray = Array.isArray(userData.offices)
            ? userData.offices
            : [userData.offices];

          const validOptions = officesArray.filter((g: string) => WORK_OFFICE_OPTIONS.includes(g));

          if (!cancelled) {
            if (validOptions.length > 0) {
              setUserOfficesOptions(validOptions);
              if (validOptions.length === 1) {
                setOffice(validOptions[0]);
              }
            } else {
              setUserOfficesOptions([]);
            }
          }
        } else if (!cancelled) {
          setUserOfficesOptions([]);
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

  const availableYearMonths = useMemo(() => {
    const values = new Set<string>();
    for (const item of submissions) {
      const ym = yearMonthFromDate(item.date);
      if (ym) values.add(ym);
    }
    return Array.from(values).sort((a, b) => b.localeCompare(a));
  }, [submissions]);

  const filteredSubmissions = useMemo(() => {
    if (!filterYearMonth) return [];
    return submissions.filter((item) => yearMonthFromDate(item.date) === filterYearMonth);
  }, [submissions, filterYearMonth]);

  useEffect(() => {
    if (availableYearMonths.length === 0) {
      if (filterYearMonth) setFilterYearMonth('');
      return;
    }
    if (!availableYearMonths.includes(filterYearMonth)) {
      setFilterYearMonth(availableYearMonths[0]);
    }
  }, [availableYearMonths, filterYearMonth]);

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

  const updateCell = (index: number, field: keyof TableRow, value: string) => {
    setRows((prev) =>
      prev.map((row, i) => (i === index ? { ...row, [field]: value } : row))
    );
  };

  const handleHeaderDateChange = (nextDate: string) => {
    setHeaderDate(nextDate);
    if (!selectedId) return;
    setSelectedId('');
    setRows(Array.from({ length: DATA_ROW_COUNT }, emptyRow));
  };

  const handleSubmit = async () => {
    if (!headerDate || !office) {
      alert('Please fill in Date and Office.');
      return;
    }

    const { idTime, iso: submittedAt } = submissionStamp();
    const docId = `${headerDate}_${office}_${idTime}`.replace(/[\/\s]/g, '_');
    setIsSubmitting(true);
    setSubmitStatus('');

    try {
      await setDoc(doc(db, 'deposit', docId), {
        date: headerDate,
        office,
        submittedAt,
        rows,
        totals: {
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
    setSelectedId(item.id);
    setHeaderDate(item.date || '');
    setOffice(item.office);
    setRows(item.rows);
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
          Dentical Deposit
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
            <div style={{ marginBottom: '12px' }}>
              <label style={{ ...labelStyle, fontSize: '13px' }}>Year / Month</label>
              <select
                value={filterYearMonth}
                onChange={(e) => setFilterYearMonth(e.target.value)}
                style={headerInputStyle}
              >
                {availableYearMonths.map((ym) => (
                  <option key={ym} value={ym}>{ym}</option>
                ))}
              </select>
            </div>
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
                  {item.submittedAt ? ` ${formatSubmittedAt(item.submittedAt)}` : ''}
                </button>
              ))}
            </div>
            </>
          )}
        </aside>

        <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap', marginBottom: '24px' }}>
          <div style={{ flex: '1', minWidth: '200px' }}>
            <label style={labelStyle}>Date:</label>
            <input
              type="date"
              value={headerDate}
              onChange={(e) => handleHeaderDateChange(e.target.value)}
              style={headerInputStyle}
            />
          </div>
          {userOfficesOptions.length > 0 && (
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={labelStyle}>Office:</label>
              {userOfficesOptions.length === 1 ? (
                <div
                  style={{
                    ...headerInputStyle,
                    display: 'flex',
                    alignItems: 'center',
                    backgroundColor: '#f8f9fb',
                    fontWeight: '600',
                  }}
                >
                  {office}
                </div>
              ) : (
                <select
                  value={office}
                  onChange={(e) => setOffice(e.target.value)}
                  style={headerInputStyle}
                >
                  <option value="">Select Office</option>
                  {userOfficesOptions.map((option) => (
                    <option key={option} value={option}>
                      {option}
                    </option>
                  ))}
                </select>
              )}
            </div>
          )}
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
                    {col}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((row, index) => (
                <tr key={index} style={{ backgroundColor: index % 2 === 0 ? '#fbfcfd' : '#ffffff' }}>
                  <td style={cellStyle}>
                    {isViewingSubmission ? (
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
                      {isViewingSubmission ? (
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
            </tbody>
          </table>
        </div>

        <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', gap: '12px', marginTop: '24px' }}>
          <button onClick={handleSubmit} disabled={isSubmitting || isDeleting} style={buttonStyle}>
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
