'use client';

import React, { useEffect, useState } from 'react';
import { auth, db } from '@/lib/firebase.config';
import { onAuthStateChanged } from 'firebase/auth';
import { doc, getDoc, setDoc } from 'firebase/firestore';

const PAPERWORK_COLLECTION = 'corporate-faxcover';

const colors = {
  pageBg: '#f8fafc',
  cardBg: '#ffffff',
  cardBorder: '#e5e7eb',
  title: '#111827',
  text: '#111827',
  placeholder: '#9ca3af',
  label: '#374151',
  inputBorder: '#d1d5db',
  inputDisabledBg: '#f9fafb',
  hint: '#6b7280',
  error: '#ef4444',
};

type TableRowKind = 'section' | 'item' | 'input';

type TableRow = {
  kind: TableRowKind;
  list: string;
  qty: string;
};

const INITIAL_TABLE_ROWS: TableRow[] = [
  { kind: 'section', list: 'Executive Department', qty: '' },
  { kind: 'item', list: '1) Daily Office Duties', qty: '' },
  { kind: 'item', list: '2) Attachments', qty: '' },
  { kind: 'section', list: 'AR Department', qty: '' },
  { kind: 'item', list: '1) Daily Office Duties', qty: '' },
  { kind: 'item', list: '2) Attachments', qty: '' },
  { kind: 'section', list: 'Call Center', qty: '' },
  { kind: 'item', list: '1) Daily Office Duties', qty: '' },
  { kind: 'item', list: '2) Attachments', qty: '' },
  { kind: 'section', list: 'Miscellaneous Papers', qty: '' },
  { kind: 'input', list: '', qty: '' },
  { kind: 'input', list: '', qty: '' },
  { kind: 'input', list: '', qty: '' },
];

const tableCellStyle: React.CSSProperties = {
  padding: 0,
  border: `1px solid ${colors.cardBorder}`,
};

const tableInputStyle = (disabled = false): React.CSSProperties => ({
  width: '100%',
  height: 40,
  border: 0,
  borderRadius: 0,
  padding: '0 12px',
  background: disabled ? colors.inputDisabledBg : colors.cardBg,
  color: colors.text,
  fontSize: 14,
  fontWeight: 400,
  boxSizing: 'border-box',
  cursor: disabled ? 'default' : 'text',
});

const tableHeaderStyle: React.CSSProperties = {
  padding: '10px 12px',
  border: `1px solid ${colors.cardBorder}`,
  background: colors.inputDisabledBg,
  color: colors.label,
  fontSize: 13,
  fontWeight: 600,
  textAlign: 'left',
};

type ShiftTableRow = {
  office: string;
  opener: string;
  openerTime: string;
  open: string;
  closer: string;
  closerTime: string;
};

const SHIFT_OFFICES = ['Ming', 'Bernard', 'Delano', 'Tulare', 'Visalia', 'Ortho', 'California', 'Fresno'];
const SHIFT_TABLE_ROW_COUNT = 8;
const SHIFT_TABLE_COLUMNS: { key: keyof ShiftTableRow; label: string; type?: string }[] = [
  { key: 'office', label: 'Office' },
  { key: 'opener', label: 'Opener' },
  { key: 'openerTime', label: 'Check In Time', type: 'time' },
  { key: 'open', label: 'Doors Open Time', type: 'time' },
  { key: 'closer', label: 'Closer' },
  { key: 'closerTime', label: 'Check Out Time', type: 'time' },
];

function createEmptyShiftRows(): ShiftTableRow[] {
  return Array.from({ length: SHIFT_TABLE_ROW_COUNT }, (_, index) => ({
    office: SHIFT_OFFICES[index] ?? '',
    opener: '',
    openerTime: '',
    open: '',
    closer: '',
    closerTime: '',
  }));
}

const CHECKLIST_ITEMS = ['All Offices Checked Out', 'Corporate Cars Pulled In & Locked', 'Janitorial Daily List Review', 'Office Lights Off', 'Code Entered', 'East Front Door Locked/Bin'];

type ChecklistRow = {
  label: string;
  done: string;
};

function createChecklistRows(): ChecklistRow[] {
  return CHECKLIST_ITEMS.map((label) => ({ label, done: '' }));
}

function cloneTableRows(rows: TableRow[] = INITIAL_TABLE_ROWS): TableRow[] {
  return rows.map((row) => ({ ...row }));
}

function mergeTableRows(saved?: TableRow[]): TableRow[] {
  const base = cloneTableRows();
  if (!Array.isArray(saved)) return base;
  return base.map((row, index) => {
    const incoming = saved[index];
    if (!incoming) return row;
    if (row.kind === 'input') {
      return {
        ...row,
        list: String(incoming.list ?? ''),
        qty: String(incoming.qty ?? ''),
      };
    }
    return { ...row, qty: String(incoming.qty ?? '') };
  });
}

function mergeShiftRows(saved?: ShiftTableRow[]): ShiftTableRow[] {
  const base = createEmptyShiftRows();
  if (!Array.isArray(saved)) return base;
  return base.map((row, index) => {
    const incoming = saved[index];
    if (!incoming) return row;
    return {
      ...row,
      opener: String(incoming.opener ?? ''),
      openerTime: String(incoming.openerTime ?? ''),
      open: String(incoming.open ?? ''),
      closer: String(incoming.closer ?? ''),
      closerTime: String(incoming.closerTime ?? ''),
    };
  });
}

function mergeChecklistRows(saved?: ChecklistRow[]): ChecklistRow[] {
  const base = createChecklistRows();
  if (!Array.isArray(saved)) return base;
  return base.map((row, index) => {
    const incoming = saved[index];
    if (!incoming) return row;
    return { ...row, done: String(incoming.done ?? '') };
  });
}

function parseQtyValue(qty: string): number {
  const parsed = parseFloat(String(qty).replace(/,/g, '').trim());
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatQtyTotal(value: number): string {
  if (Number.isInteger(value)) return String(value);
  return String(Math.round(value * 1000) / 1000);
}

function getTodayDate(): string {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

const fieldLabelStyle: React.CSSProperties = {
  display: 'block',
  marginBottom: 6,
  fontWeight: 600,
  fontSize: 13,
  color: colors.label,
};

const inputStyle = (disabled = false): React.CSSProperties => ({
  width: '100%',
  height: 40,
  border: `1px solid ${colors.inputBorder}`,
  borderRadius: 10,
  padding: '0 12px',
  background: disabled ? colors.inputDisabledBg : colors.cardBg,
  color: colors.text,
  fontSize: 14,
  fontWeight: 600,
  boxSizing: 'border-box',
});

type TextFieldProps = {
  label: string;
  value: string;
  onChange?: (value: string) => void;
  placeholder?: string;
  type?: string;
  readOnly?: boolean;
  disabled?: boolean;
};

function TextField({
  label,
  value,
  onChange,
  placeholder,
  type = 'text',
  readOnly = false,
  disabled = false,
}: TextFieldProps) {
  const locked = readOnly || disabled;
  return (
    <div>
      <label style={fieldLabelStyle}>{label}</label>
      <input
        type={type}
        value={value}
        readOnly={readOnly}
        disabled={disabled}
        placeholder={placeholder}
        onChange={(event) => {
          if (locked) return;
          onChange?.(event.target.value);
        }}
        style={inputStyle(locked)}
      />
    </div>
  );
}

export default function SignLogPage() {
  const [pageReady, setPageReady] = useState(false);
  const [date, setDate] = useState(getTodayDate);
  const [cameInName, setCameInName] = useState('');
  const [timeCameIn, setTimeCameIn] = useState('');
  const [leftName, setLeftName] = useState('');
  const [timeLeft, setTimeLeft] = useState('');
  const [nextDayOpener, setNextDayOpener] = useState('');
  const [nextDayCloser, setNextDayCloser] = useState('');
  const [tableRows, setTableRows] = useState<TableRow[]>(() => cloneTableRows());
  const [shiftRows, setShiftRows] = useState<ShiftTableRow[]>(createEmptyShiftRows);
  const [checklistRows, setChecklistRows] = useState<ChecklistRow[]>(createChecklistRows);
  const [recordLoading, setRecordLoading] = useState(false);
  const [submitting, setSubmitting] = useState(false);
  const [isSubmitted, setIsSubmitted] = useState(false);
  const [status, setStatus] = useState('');
  const [error, setError] = useState('');

  const resetForm = () => {
    setCameInName('');
    setTimeCameIn('');
    setLeftName('');
    setTimeLeft('');
    setNextDayOpener('');
    setNextDayCloser('');
    setTableRows(cloneTableRows());
    setShiftRows(createEmptyShiftRows());
    setChecklistRows(createChecklistRows());
    setIsSubmitted(false);
  };

  const applyPaperworkData = (data: Record<string, unknown>) => {
    setCameInName(String(data.cameInName ?? ''));
    setTimeCameIn(String(data.timeCameIn ?? ''));
    setLeftName(String(data.leftName ?? ''));
    setTimeLeft(String(data.timeLeft ?? ''));
    setNextDayOpener(String(data.nextDayOpener ?? ''));
    setNextDayCloser(String(data.nextDayCloser ?? ''));
    setTableRows(mergeTableRows(data.tableRows as TableRow[] | undefined));
    setShiftRows(mergeShiftRows(data.shiftRows as ShiftTableRow[] | undefined));
    setChecklistRows(mergeChecklistRows(data.checklistRows as ChecklistRow[] | undefined));
  };

  const updateTableRow = (index: number, field: keyof TableRow, value: string) => {
    if (isSubmitted) return;
    setTableRows((prev) =>
      prev.map((row, rowIndex) => (rowIndex === index ? { ...row, [field]: value } : row))
    );
  };

  const updateShiftRow = (index: number, field: keyof ShiftTableRow, value: string) => {
    if (isSubmitted) return;
    setShiftRows((prev) =>
      prev.map((row, rowIndex) => (rowIndex === index ? { ...row, [field]: value } : row))
    );
  };

  const updateChecklistDone = (index: number, done: string) => {
    if (isSubmitted) return;
    setChecklistRows((prev) =>
      prev.map((row, rowIndex) => (rowIndex === index ? { ...row, done } : row))
    );
  };

  const qtyTotal = tableRows.reduce((sum, row) => sum + parseQtyValue(row.qty), 0);

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
        if (userData?.role !== 'HR' && userData?.role !== 'Corporate') {
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
    if (!pageReady || !date) return;

    let cancelled = false;
    const load = async () => {
      setRecordLoading(true);
      setStatus('');
      setError('');
      try {
        const snap = await getDoc(doc(db, PAPERWORK_COLLECTION, date));
        if (cancelled) return;
        if (!snap.exists()) {
          resetForm();
          setStatus('No submitted data for this date.');
          return;
        }
        applyPaperworkData(snap.data() as Record<string, unknown>);
        setIsSubmitted(true);
        setStatus('This date has been submitted.');
      } catch {
        if (!cancelled) {
          resetForm();
          setError('Failed to load.');
        }
      } finally {
        if (!cancelled) {
          setRecordLoading(false);
        }
      }
    };

    load();
    return () => {
      cancelled = true;
    };
  }, [pageReady, date]);

  const handleSubmit = async () => {
    if (!date || submitting || isSubmitted) return;
    setSubmitting(true);
    setStatus('');
    setError('');
    try {
      await setDoc(doc(db, PAPERWORK_COLLECTION, date), {
        date,
        cameInName,
        timeCameIn,
        leftName,
        timeLeft,
        nextDayOpener,
        nextDayCloser,
        tableRows,
        shiftRows,
        checklistRows,
        qtyTotal: formatQtyTotal(qtyTotal),
        submitted: true,
        submittedAt: new Date().toISOString(),
      });
      setIsSubmitted(true);
      setStatus('Submitted.');
    } catch {
      setError('Failed to submit.');
    } finally {
      setSubmitting(false);
    }
  };

  if (!pageReady) {
    return <main style={{ minHeight: '100vh', background: colors.pageBg }} />;
  }

  return (
    <main style={{ minHeight: '100vh', background: colors.pageBg, padding: 32 }}>
      <section
        style={{
          maxWidth: 1500,
          margin: '32px auto',
          border: `1px solid ${colors.cardBorder}`,
          borderRadius: 14,
          padding: '28px 26px',
          background: colors.cardBg,
          boxShadow: '0 2px 12px rgba(15, 23, 42, 0.05)',
        }}
      >
        <h1 style={{ margin: '0 0 10px', fontSize: 30, fontWeight: 700, color: colors.title, textAlign: 'center' }}>
          Corporate End of Day Cover
        </h1>
        <h3 style={{ margin: '0 0 35px', fontSize: 16, fontWeight: 700, color: colors.title, textAlign: 'center' }}>
          (Closer Collect and Attach Paperwork Provided)
        </h3>

        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1.4fr 1fr', gap: 14 }}>
          <TextField
            label="Date"
            value={date}
            onChange={setDate}
            type="date"
          />
          <TextField
            label="Opener Name"
            value={cameInName}
            onChange={setCameInName}
            placeholder="Enter name"
            disabled={isSubmitted}
          />
          <TextField
            label="Time Came In"
            value={timeCameIn}
            onChange={setTimeCameIn}
            type="time"
            disabled={isSubmitted}
          />
        </div>

        <div
          style={{
            display: 'grid',
            gridTemplateColumns: '1fr 1.4fr 1fr',
            gap: 14,
            marginTop: 14,
          }}
        >
          <div />
          <TextField
            label="Closer Name"
            value={leftName}
            onChange={setLeftName}
            placeholder="Enter name"
            disabled={isSubmitted}
          />
          <TextField
            label="Time Left"
            value={timeLeft}
            onChange={setTimeLeft}
            type="time"
            disabled={isSubmitted}
          />
        </div>

        <div
          style={{
            display: 'grid',
            gridTemplateColumns: '1fr 1.4fr 1fr',
            gap: 14,
            marginTop: 14,
          }}
        >
          <div />
          <TextField
            label="Next Day Opener"
            value={nextDayOpener}
            onChange={setNextDayOpener}
            placeholder="Enter name"
            disabled={isSubmitted}
          />
          <TextField
            label="Next Day Closer"
            value={nextDayCloser}
            onChange={setNextDayCloser}
            placeholder="Enter name"
            disabled={isSubmitted}
          />
        </div>

        <table
          style={{
            width: '100%',
            marginTop: 24,
            borderCollapse: 'collapse',
            tableLayout: 'fixed',
          }}
        >
          <thead>
            <tr>
              <th style={{ ...tableHeaderStyle, width: '70%' }}>
                Departments
              </th>
              <th style={{ ...tableHeaderStyle, width: '30%' }}>
                Qty
              </th>
            </tr>
          </thead>
          <tbody>
            {tableRows.map((row, index) => {
              if (row.kind === 'section') {
                return (
                  <tr key={index}>
                    <td
                      colSpan={2}
                      style={{
                        padding: '10px 12px',
                        border: `1px solid ${colors.cardBorder}`,
                        background: colors.inputDisabledBg,
                        color: colors.text,
                        fontSize: 14,
                        fontWeight: 600,
                        textAlign: 'center',
                      }}
                    >
                      {row.list}
                    </td>
                  </tr>
                );
              }

              return (
                <tr key={index}>
                  <td style={tableCellStyle}>
                    {row.kind === 'item' ? (
                      <div
                        style={{
                          height: 40,
                          padding: '0 12px',
                          display: 'flex',
                          alignItems: 'center',
                          color: colors.text,
                          fontSize: 14,
                        }}
                      >
                        {row.list}
                      </div>
                    ) : (
                      <input
                        type="text"
                        value={row.list}
                        readOnly={isSubmitted}
                        disabled={isSubmitted}
                        onChange={(event) => updateTableRow(index, 'list', event.target.value)}
                        style={tableInputStyle(isSubmitted)}
                      />
                    )}
                  </td>
                  <td style={tableCellStyle}>
                    <input
                      type="text"
                      value={row.qty}
                      readOnly={isSubmitted}
                      disabled={isSubmitted}
                      onChange={(event) => updateTableRow(index, 'qty', event.target.value)}
                      style={tableInputStyle(isSubmitted)}
                    />
                  </td>
                </tr>
              );
            })}
            <tr>
              <td
                style={{
                  height: 40,
                  padding: '0 12px',
                  border: `1px solid ${colors.cardBorder}`,
                  background: colors.inputDisabledBg,
                  color: colors.text,
                  fontSize: 14,
                  fontWeight: 600,
                  verticalAlign: 'middle',
                }}
              >
                Total
              </td>
              <td
                style={{
                  height: 40,
                  padding: '0 12px',
                  border: `1px solid ${colors.cardBorder}`,
                  background: colors.inputDisabledBg,
                  color: colors.text,
                  fontSize: 14,
                  fontWeight: 600,
                  verticalAlign: 'middle',
                }}
              >
                {formatQtyTotal(qtyTotal)}
              </td>
            </tr>
          </tbody>
        </table>

        <table
          style={{
            width: '100%',
            marginTop: 24,
            borderCollapse: 'collapse',
            tableLayout: 'fixed',
          }}
        >
          <thead>
            <tr>
              {SHIFT_TABLE_COLUMNS.map((column) => (
                <th key={column.key} style={tableHeaderStyle}>
                  {column.label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {shiftRows.map((row, index) => (
              <tr key={index}>
                {SHIFT_TABLE_COLUMNS.map((column) => {
                  const isFixedOffice = column.key === 'office' && index < SHIFT_OFFICES.length;
                  return (
                    <td key={column.key} style={tableCellStyle}>
                      {isFixedOffice ? (
                        <div
                          style={{
                            height: 40,
                            padding: '0 12px',
                            display: 'flex',
                            alignItems: 'center',
                            color: colors.text,
                            fontSize: 14,
                          }}
                        >
                          {row.office}
                        </div>
                      ) : (
                        <input
                          type={column.type ?? 'text'}
                          value={row[column.key]}
                          readOnly={isSubmitted}
                          disabled={isSubmitted}
                          onChange={(event) => updateShiftRow(index, column.key, event.target.value)}
                          style={tableInputStyle(isSubmitted)}
                        />
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>

        <table
          style={{
            width: '100%',
            marginTop: 24,
            borderCollapse: 'collapse',
            tableLayout: 'fixed',
          }}
        >
          <thead>
            <tr>
              <th style={{ ...tableHeaderStyle, width: '70%' }}>Keyholder Checklist</th>
              <th style={{ ...tableHeaderStyle, width: '30%' }}>Done</th>
            </tr>
          </thead>
          <tbody>
            {checklistRows.map((row, index) => (
              <tr key={row.label}>
                <td
                  style={{
                    height: 40,
                    padding: '0 12px',
                    border: `1px solid ${colors.cardBorder}`,
                    color: colors.text,
                    fontSize: 14,
                    verticalAlign: 'middle',
                  }}
                >
                  {row.label}
                </td>
                <td style={tableCellStyle}>
                  <input
                    type="text"
                    value={row.done}
                    readOnly={isSubmitted}
                    disabled={isSubmitted}
                    onChange={(event) => updateChecklistDone(index, event.target.value)}
                    style={tableInputStyle(isSubmitted)}
                  />
                </td>
              </tr>
            ))}
          </tbody>
        </table>

        {(status || error || recordLoading) && (
          <p
            style={{
              margin: '16px 0 0',
              fontSize: 13,
              color: error ? colors.error : colors.hint,
            }}
          >
            {recordLoading ? 'Loading...' : error || status}
          </p>
        )}

        {!isSubmitted && (
          <button
            type="button"
            onClick={handleSubmit}
            disabled={submitting || recordLoading || !date}
            style={{
              marginTop: 16,
              width: '100%',
              height: 44,
              border: `1px solid ${colors.cardBorder}`,
              borderRadius: 10,
              background: colors.cardBg,
              color: colors.text,
              fontWeight: 600,
              fontSize: 14,
              cursor: submitting || recordLoading || !date ? 'not-allowed' : 'pointer',
            }}
          >
            {submitting ? 'Submitting...' : 'Submit'}
          </button>
        )}
      </section>
    </main>
  );
}

