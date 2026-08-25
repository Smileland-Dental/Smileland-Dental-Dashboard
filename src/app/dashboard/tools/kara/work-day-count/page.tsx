'use client';

import React, { useEffect, useState } from 'react';
import { auth, db } from '@/lib/firebase.config';
import { onAuthStateChanged } from 'firebase/auth';
import { doc, getDoc, setDoc } from 'firebase/firestore';

const PURPLE_BEAR_COLLECTION = 'doctor-work-day';
const DAY_COUNT_DOC_ID = 'term';

const COLUMNS = [
  'Name',
  'Current Date',
  'Contract Length',
  'Days Missed / Not Completed Days',
  'Total Days Completed',
  'Total Days Left to Complete',
  'Estimated Term Date',
] as const;

type Column = (typeof COLUMNS)[number];

const HIGHLIGHT_VALUES = ['Less Than 90 Days Left To Complete', 'Contract Terms In Current Year', 'Contract Terminated', '90 Days Completed', '2 Year Contract', 'Pending Hire'] as const;
type HighlightValue = (typeof HIGHLIGHT_VALUES)[number];

type RowData = Record<Column, string> & {
  highlight: HighlightValue[];
};

const ROW_COUNT = 20;

const colors = {
  pageBg: '#f8fafc',
  cardBg: '#ffffff',
  cardBorder: '#e5e7eb',
  title: '#111827',
  text: '#111827',
  headerBg: '#f3f4f6',
  headerText: '#374151',
  rowBorder: '#e5e7eb',
  rowHover: '#f9fafb',
  inputBorder: '#d1d5db',
  accent: '#111827',
  accentSoft: '#f3f4f6',
  yellow: '#fef08a',
  pink: '#fbcfe8',
  red: '#fecaca',
  green: '#bbf7d0',
  orange: '#fdba74',
  purple: '#e9d5ff',
};

function getCaliforniaDate(): string {
  return new Intl.DateTimeFormat('en-US', {
    timeZone: 'America/Los_Angeles',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).format(new Date());
}

const AUTO_COLUMNS: Column[] = ['Current Date', 'Total Days Left to Complete', 'Estimated Term Date'];

function createEmptyRow(): RowData {
  const today = getCaliforniaDate();
  const row = COLUMNS.reduce(
    (acc, column) => {
      acc[column] = column === 'Current Date' ? today : '';
      return acc;
    },
    {} as Record<Column, string>
  );
  return { ...row, highlight: [] };
}

function createEmptyRows(): RowData[] {
  return Array.from({ length: ROW_COUNT }, () => createEmptyRow());
}

function isHighlightValue(value: unknown): value is HighlightValue {
  return HIGHLIGHT_VALUES.includes(value as HighlightValue);
}

function normalizeHighlights(raw: unknown): HighlightValue[] {
  if (Array.isArray(raw)) {
    return raw.filter(isHighlightValue);
  }
  if (isHighlightValue(raw)) {
    return [raw];
  }
  return [];
}

function normalizeRow(raw: unknown): RowData {
  const source = raw && typeof raw === 'object' ? (raw as Record<string, unknown>) : {};
  const base = createEmptyRow();
  for (const column of COLUMNS) {
    if (typeof source[column] === 'string') {
      base[column] = source[column];
    }
  }
  base.highlight = normalizeHighlights(source.highlight);
  return base;
}

function normalizeRows(raw: unknown): RowData[] {
  if (!Array.isArray(raw) || raw.length === 0) return createEmptyRows();
  return raw.map((item) => normalizeRow(item));
}

function serializeRows(rows: RowData[]) {
  return rows.map((row) => ({
    Name: row.Name,
    'Curent Date': row['Current Date'],
    'Contract Length': row['Contract Length'],
    'Days Missed / Not Completed Days': row['Days Missed / Not Completed Days'],
    'Total Days Completed': row['Total Days Completed'],
    'Total Days Left to Complete': row['Total Days Left to Complete'],
    'Estimated Term Date': row['Estimated Term Date'],
    highlight: row.highlight,
  }));
}

function calcDaysLeft(length: string, daysCompleted: string): string {
  const lengthNum = Number(length);
  const completedNum = Number(daysCompleted);
  if (length.trim() === '' || daysCompleted.trim() === '') return '';
  if (!Number.isFinite(lengthNum) || !Number.isFinite(completedNum)) return '';
  return String(lengthNum - completedNum);
}

function parseUsDate(dateStr: string): { year: number; month: number; day: number } | null {
  const match = dateStr.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!match) return null;
  const month = Number(match[1]);
  const day = Number(match[2]);
  const year = Number(match[3]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  return { year, month, day };
}

function formatUsDate(date: Date): string {
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const day = String(date.getUTCDate()).padStart(2, '0');
  const year = date.getUTCFullYear();
  return `${month}/${day}/${year}`;
}

function calcTermDate(dateStr: string, daysLeftStr: string): string {
  if (daysLeftStr.trim() === '') return '';
  const daysLeft = Number(daysLeftStr);
  if (!Number.isFinite(daysLeft) || !Number.isInteger(daysLeft)) return '';

  const parsed = parseUsDate(dateStr);
  if (!parsed) return '';

  const date = new Date(Date.UTC(parsed.year, parsed.month - 1, parsed.day, 12));
  let remaining = Math.abs(daysLeft);
  const direction = daysLeft >= 0 ? 1 : -1;

  while (remaining > 0) {
    date.setUTCDate(date.getUTCDate() + direction);
    const weekday = date.getUTCDay();
    if (weekday !== 0 && weekday !== 6) {
      remaining -= 1;
    }
  }

  date.setUTCDate(date.getUTCDate() + direction);

  return formatUsDate(date);
}

function countWeekdaysBetween(fromStr: string, toStr: string): number | null {
  const from = parseUsDate(fromStr);
  const to = parseUsDate(toStr);
  if (!from || !to) return null;

  const start = new Date(Date.UTC(from.year, from.month - 1, from.day, 12));
  const end = new Date(Date.UTC(to.year, to.month - 1, to.day, 12));
  if (start.getTime() === end.getTime()) return 0;

  const direction = end.getTime() > start.getTime() ? 1 : -1;
  let count = 0;
  const cursor = new Date(start.getTime());
  while (cursor.getTime() !== end.getTime()) {
    cursor.setUTCDate(cursor.getUTCDate() + direction);
    const weekday = cursor.getUTCDay();
    if (weekday !== 0 && weekday !== 6) {
      count += 1;
    }
  }

  return direction * count;
}

function calcDaysLeftForDate(previousDate: string, nextDate: string, daysLeftStr: string): string {
  if (daysLeftStr.trim() === '') return '';
  const daysLeft = Number(daysLeftStr);
  if (!Number.isFinite(daysLeft) || !Number.isInteger(daysLeft)) return daysLeftStr;

  const elapsed = countWeekdaysBetween(previousDate, nextDate);
  if (elapsed === null) return daysLeftStr;

  return String(daysLeft - elapsed);
}

function getCellHighlightBg(highlights: HighlightValue[], column: Column): string | undefined {
  let bg: string | undefined;
  if (highlights.includes('Contract Terminated')) bg = colors.red;
  if (highlights.includes('Less Than 90 Days Left To Complete') && column === 'Total Days Left to Complete') bg = colors.yellow;
  if (highlights.includes('Contract Terms In Current Year') && column === 'Estimated Term Date') bg = colors.pink;
  if (highlights.includes('90 Days Completed') && column === 'Total Days Completed') bg = colors.green;
  if (highlights.includes('2 Year Contract') && (column === 'Name' || column === 'Contract Length')) bg = colors.orange;
  if (highlights.includes('Pending Hire') && column === 'Name') bg = colors.purple;
  return bg;
}

const HIGHLIGHT_LEGEND: { label: HighlightValue; color: string }[] = [
  { label: 'Less Than 90 Days Left To Complete', color: colors.yellow },
  { label: 'Contract Terms In Current Year', color: colors.pink },
  { label: 'Contract Terminated', color: colors.red },
  { label: '90 Days Completed', color: colors.green },
  { label: '2 Year Contract', color: colors.orange },
  { label: 'Pending Hire', color: colors.purple },
];

const COLUMN_WIDTHS: Partial<Record<Column, string>> = {
  Name: '16%',
  'Current Date': '10%',
  'Contract Length': '12%',
  'Days Missed / Not Completed Days': '15%',
  'Total Days Completed': '15%',
  'Total Days Left to Complete': '18%',
  'Estimated Term Date': '14%',
};

const headerCellStyle: React.CSSProperties = {
  padding: '12px 10px',
  textAlign: 'left',
  fontWeight: 600,
  fontSize: 13,
  color: colors.headerText,
  background: colors.headerBg,
  borderBottom: `1px solid ${colors.rowBorder}`,
};

export default function PageW() {
  const [pageReady, setPageReady] = useState(false);
  const [isEditing, setIsEditing] = useState(false);
  const [rows, setRows] = useState(createEmptyRows);
  const [countAsOf, setCountAsOf] = useState('');
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [statusMessage, setStatusMessage] = useState('');
  const [error, setError] = useState('');

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
        if (userData?.role !== 'HR' && userData?.role !== 'Director') {
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

    let cancelled = false;

    const load = async () => {
      setLoading(true);
      setError('');
      try {
        const snap = await getDoc(doc(db, PURPLE_BEAR_COLLECTION, DAY_COUNT_DOC_ID));
        if (cancelled) return;
        if (snap.exists()) {
          const data = snap.data();
          setRows(normalizeRows(data?.rows));
          setCountAsOf(typeof data?.countAsOf === 'string' ? data.countAsOf : '');
        } else {
          setRows(createEmptyRows());
          setCountAsOf('');
        }
      } catch (e: unknown) {
        if (!cancelled) {
          const message = e instanceof Error ? e.message : 'Failed to load';
          setError(message);
        }
      } finally {
        if (!cancelled) setLoading(false);
      }
    };

    load();
    return () => {
      cancelled = true;
    };
  }, [pageReady]);

  const handleCellChange = (rowIndex: number, column: Column, value: string) => {
    setRows((prev) =>
      prev.map((row, index) => {
        if (index !== rowIndex) return row;
        const nextRow = { ...row, [column]: value };
        if (column === 'Contract Length' || column === 'Total Days Completed') {
          nextRow['Total Days Left to Complete'] = calcDaysLeft(nextRow['Contract Length'], nextRow['Total Days Completed']);
          nextRow['Estimated Term Date'] = calcTermDate(nextRow['Current Date'], nextRow['Total Days Left to Complete']);
        }
        return nextRow;
      })
    );
  };

  const handleHighlightToggle = (rowIndex: number, value: HighlightValue) => {
    setRows((prev) =>
      prev.map((row, index) => {
        if (index !== rowIndex) return row;
        const has = row.highlight.includes(value);
        const nextHighlight = has
          ? row.highlight.filter((item) => item !== value)
          : [...row.highlight, value];
        return { ...row, highlight: nextHighlight };
      })
    );
  };

  const handleAddRow = () => {
    setRows((prev) => [...prev, createEmptyRow()]);
  };

  const handleDeleteRow = (rowIndex: number) => {
    setRows((prev) => {
      if (prev.length <= 1) return prev;
      return prev.filter((_, index) => index !== rowIndex);
    });
  };

  const handleEditToggle = async () => {
    if (!isEditing) {
      const today = getCaliforniaDate();
      setRows((prev) =>
        prev.map((row) => {
          const nextRow = { ...row, Date: today };
          nextRow['Total Days Left to Complete'] = calcDaysLeftForDate(row['Current Date'], today, row['Total Days Left to Complete']);
          nextRow['Estimated Term Date'] = calcTermDate(nextRow['Current Date'], nextRow['Total Days Left to Complete']);
          return nextRow;
        })
      );
      setStatusMessage('');
      setError('');
      setIsEditing(true);
      return;
    }

    setSaving(true);
    setError('');
    setStatusMessage('');
    try {
      await setDoc(doc(db, PURPLE_BEAR_COLLECTION, DAY_COUNT_DOC_ID), {
        rows: serializeRows(rows),
        countAsOf,
        updatedAt: new Date().toISOString(),
      });
      setIsEditing(false);
      setStatusMessage('Saved');
    } catch (e: unknown) {
      const message = e instanceof Error ? e.message : 'Failed to save';
      setError(message);
    } finally {
      setSaving(false);
    }
  };

  if (!pageReady) {
    return <main style={{ minHeight: '100vh', background: colors.pageBg }} />;
  }

  return (
    <main style={{ minHeight: '100vh', background: colors.pageBg, padding: '24px 20px' }}>
      <section
        style={{
          width: '100%',
          maxWidth: 1200,
          margin: '0 auto',
          border: `1px solid ${colors.cardBorder}`,
          borderRadius: 14,
          padding: '24px 22px',
          background: colors.cardBg,
          boxShadow: '0 2px 12px rgba(15, 23, 42, 0.05)',
          overflow: 'visible',
          boxSizing: 'border-box',
        }}
      >
        <div
          style={{
            display: 'flex',
            alignItems: 'flex-start',
            justifyContent: 'space-between',
            gap: 12,
            marginBottom: 20,
          }}
        >
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8, minWidth: 0 }}>
            <h1 style={{ margin: 0, fontSize: 24, fontWeight: 700, color: colors.title }}>
              Work Day Count / Estimated Term Date
            </h1>
            <div
              style={{
                display: 'flex',
                flexWrap: 'wrap',
                gap: '6px 12px',
                alignItems: 'center',
              }}
            >
              {HIGHLIGHT_LEGEND.map((item) => (
                <span
                  key={item.label}
                  style={{
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: 5,
                    fontSize: 11,
                    color: colors.headerText,
                    lineHeight: 1.2,
                  }}
                >
                  <span
                    style={{
                      width: 10,
                      height: 10,
                      borderRadius: 2,
                      background: item.color,
                      border: `1px solid ${colors.cardBorder}`,
                      flexShrink: 0,
                    }}
                  />
                  <span style={{ fontWeight: 600 }}>{item.label}</span>
                </span>
              ))}
            </div>
          </div>
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: 6 }}>
            <button
              type="button"
              onClick={handleEditToggle}
              disabled={loading || saving}
              style={{
                border: `1px solid ${colors.cardBorder}`,
                borderRadius: 10,
                padding: '8px 16px',
                background: isEditing ? colors.accent : colors.accentSoft,
                color: isEditing ? '#ffffff' : colors.text,
                fontWeight: 600,
                fontSize: 14,
                cursor: loading || saving ? 'not-allowed' : 'pointer',
                flexShrink: 0,
                opacity: loading || saving ? 0.7 : 1,
              }}
            >
              {saving ? 'Saving...' : isEditing ? 'Done' : 'Edit'}
            </button>
            {statusMessage && (
              <span style={{ fontSize: 12, color: colors.headerText }}>{statusMessage}</span>
            )}
            {error && <span style={{ fontSize: 12, color: '#ef4444' }}>{error}</span>}
          </div>
        </div>

        {loading && (
          <p style={{ margin: '0 0 14px', fontSize: 14, color: colors.headerText }}>Loading...</p>
        )}

        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            gap: 8,
            marginBottom: 14,
            fontSize: 14,
            color: colors.text,
          }}
        >
          <span style={{ fontWeight: 600, whiteSpace: 'nowrap' }}>Work Day Count As Of:</span>
          {isEditing ? (
            <input
              type="text"
              value={countAsOf}
              onChange={(e) => setCountAsOf(e.target.value)}
              placeholder="MM/DD/YYYY"
              style={{
                width: 160,
                height: 32,
                border: `1px solid ${colors.inputBorder}`,
                borderRadius: 8,
                padding: '0 10px',
                fontSize: 14,
                color: colors.text,
                background: colors.cardBg,
                outline: 'none',
                boxSizing: 'border-box',
              }}
            />
          ) : (
            <span>{countAsOf || '—'}</span>
          )}
        </div>

        <table
          style={{
            width: '100%',
            tableLayout: 'fixed',
            borderCollapse: 'collapse',
            fontSize: 14,
            color: colors.text,
          }}
        >
          <colgroup>
            {COLUMNS.map((column) => (
              <col key={column} style={{ width: COLUMN_WIDTHS[column] }} />
            ))}
            {isEditing && (
              <>
                <col style={{ width: '180px' }} />
                <col style={{ width: '80px' }} />
              </>
            )}
          </colgroup>
          <thead>
            <tr>
              {COLUMNS.map((column) => (
                <th key={column} style={headerCellStyle}>
                  {column}
                </th>
              ))}
              {isEditing && (
                <>
                  <th style={{ ...headerCellStyle, textAlign: 'center' }}>Highlight</th>
                  <th style={{ ...headerCellStyle, textAlign: 'center' }}>Action</th>
                </>
              )}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, rowIndex) => {
              const rowBg = row.highlight.includes('Contract Terminated') ? colors.red : undefined;

              return (
                <tr key={rowIndex} style={{ background: rowBg }}>
                  {COLUMNS.map((column) => {
                    const cellBg = getCellHighlightBg(row.highlight, column);
                    return (
                      <td
                        key={`${rowIndex}-${column}`}
                        style={{
                          padding: isEditing ? '8px 8px' : '12px 10px',
                          borderBottom: `1px solid ${colors.rowBorder}`,
                          height: 44,
                          background: cellBg,
                          overflowWrap: 'anywhere',
                          wordBreak: 'break-word',
                        }}
                      >
                        {isEditing && !AUTO_COLUMNS.includes(column) ? (
                          <input
                            type="text"
                            value={row[column]}
                            onChange={(e) => handleCellChange(rowIndex, column, e.target.value)}
                            style={{
                              width: '100%',
                              height: 32,
                              border: `1px solid ${colors.inputBorder}`,
                              borderRadius: 8,
                              padding: '0 10px',
                              fontSize: 14,
                              color: colors.text,
                              background: cellBg ?? colors.cardBg,
                              outline: 'none',
                              boxSizing: 'border-box',
                            }}
                          />
                        ) : (
                          row[column]
                        )}
                      </td>
                    );
                  })}
                  {isEditing && (
                    <>
                      <td
                        style={{
                          padding: '8px 8px',
                          borderBottom: `1px solid ${colors.rowBorder}`,
                          background: rowBg,
                          verticalAlign: 'middle',
                        }}
                      >
                        <div
                          style={{
                            display: 'flex',
                            flexWrap: 'wrap',
                            gap: '4px 8px',
                          }}
                        >
                          {HIGHLIGHT_VALUES.map((value) => {
                            const checked = row.highlight.includes(value);
                            return (
                              <label
                                key={value}
                                style={{
                                  display: 'inline-flex',
                                  alignItems: 'center',
                                  gap: 4,
                                  fontSize: 11,
                                  color: colors.text,
                                  cursor: 'pointer',
                                  whiteSpace: 'nowrap',
                                }}
                              >
                                <input
                                  type="checkbox"
                                  checked={checked}
                                  onChange={() => handleHighlightToggle(rowIndex, value)}
                                  style={{ margin: 0, cursor: 'pointer' }}
                                />
                                {value}
                              </label>
                            );
                          })}
                        </div>
                      </td>
                      <td
                        style={{
                          padding: '8px 10px',
                          borderBottom: `1px solid ${colors.rowBorder}`,
                          textAlign: 'center',
                          background: rowBg,
                        }}
                      >
                        <button
                          type="button"
                          onClick={() => handleDeleteRow(rowIndex)}
                          disabled={rows.length <= 1}
                          style={{
                            border: `1px solid ${colors.cardBorder}`,
                            borderRadius: 8,
                            padding: '6px 10px',
                            background: rows.length <= 1 ? colors.accentSoft : '#fff',
                            color: rows.length <= 1 ? '#9ca3af' : '#b91c1c',
                            fontWeight: 600,
                            fontSize: 12,
                            cursor: rows.length <= 1 ? 'not-allowed' : 'pointer',
                          }}
                        >
                          Delete
                        </button>
                      </td>
                    </>
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>

        {isEditing && (
          <button
            type="button"
            onClick={handleAddRow}
            style={{
              marginTop: 14,
              border: `1px dashed ${colors.cardBorder}`,
              borderRadius: 10,
              padding: '10px 16px',
              background: colors.accentSoft,
              color: colors.text,
              fontWeight: 600,
              fontSize: 14,
              cursor: 'pointer',
              width: '100%',
            }}
          >
            + Add Row
          </button>
        )}
      </section>
    </main>
  );
}

