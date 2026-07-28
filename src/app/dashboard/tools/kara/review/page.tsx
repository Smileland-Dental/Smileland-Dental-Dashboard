'use client';

import React, { Suspense, useEffect, useState } from 'react';
import { useSearchParams } from 'next/navigation';
import { db } from '@/lib/firebase.config';
import { collection, doc, getDoc, getDocs, serverTimestamp, setDoc } from 'firebase/firestore';

const PINK_BEAR_COLLECTION = 'monthly report';
const BLACK_BEAR_COLLECTION = 'simple-forms';
const REVIEW_DAILY_GOAL = 2;
const COLUMN_COUNT = 6;
const LEGACY_REVIEW_COLUMN_COUNT = 5;
const COLUMN_HEADERS = ['Source', 'Rating', 'Date', 'Name', 'Review', 'Response'];
const DATA_COLUMN_MIN_WIDTHS = [90, 70, 100, 160, 560, 700];
const REVIEWS_DATA_MIN_WIDTH = DATA_COLUMN_MIN_WIDTHS.reduce((sum, width) => sum + width, 0);
const REVIEW_ACTION_COLUMN_WIDTH = 300;
const REVIEWS_TABLE_MIN_WIDTH = REVIEWS_DATA_MIN_WIDTH + REVIEW_ACTION_COLUMN_WIDTH;
const REVIEWS_SECTION_MAX_WIDTH = REVIEWS_TABLE_MIN_WIDTH;
const ISSUE_COLUMN_COUNT = 2;
const ISSUE_COLUMN_HEADERS = ['Date', 'Issue'];
const ISSUE_COLUMN_MIN_WIDTHS = [100, 480];
const ISSUES_TABLE_MIN_WIDTH = ISSUE_COLUMN_MIN_WIDTHS.reduce((sum, width) => sum + width, 0);
const EQUIPMENT_COLUMN_COUNT = 3;
const EQUIPMENT_COLUMN_HEADERS = ['Category', 'Description', 'Status'];
const EQUIPMENT_COLUMN_MIN_WIDTHS = [395, 790, 395];
const EQUIPMENT_TABLE_MIN_WIDTH = EQUIPMENT_COLUMN_MIN_WIDTHS.reduce((sum, width) => sum + width, 0);
const EQUIPMENT_ITEM_ROW_COLORS: Record<string, string> = {
  'DEM': '#dcfce7',
  'IT': '#e0f2fe',
  'GM': '#fce7f3',
  'J': '#ede9fe',
};

type ReportSection = 'reviews' | 'issues' | 'equipment';

type TableRow = {
  id: string;
  cells: string[];
  highlighted?: boolean;
  deletedReview?: boolean;
};

function getEquipmentRowBackground(row: TableRow): string | undefined {
  const item = String(row.cells[0] ?? '').trim().toUpperCase();
  return EQUIPMENT_ITEM_ROW_COLORS[item];
}

const SECTION_OPTIONS: { value: ReportSection; label: string }[] = [
  { value: 'reviews', label: 'Review' },
  { value: 'issues', label: 'Injury' },
  { value: 'equipment', label: 'Broken Equipment' },
];

const dataCellInputStyle: React.CSSProperties = {
  width: '100%',
  height: 32,
  border: '1px solid #d1d5db',
  borderRadius: 6,
  padding: '0 8px',
  fontSize: 14,
};

const centeredDataCellInputStyle: React.CSSProperties = {
  ...dataCellInputStyle,
  textAlign: 'center',
};

const flexSectionRowStyle: React.CSSProperties = {
  display: 'flex',
  gap: 16,
  marginBottom: 20,
  alignItems: 'stretch',
  width: '100%',
};

const addRowButtonStyle: React.CSSProperties = {
  border: '1px solid #e5e7eb',
  borderRadius: 8,
  padding: '8px 14px',
  background: '#fff',
  color: '#111827',
  fontWeight: 600,
  cursor: 'pointer',
};

type StoredTableRow = {
  cells?: string[];
  highlighted?: boolean;
  deletedReview?: boolean;
};

type ReviewSummarySnapshot = {
  dailyGoal: number;
  monthlyGoal: number | null;
  reviewDayCount: number | null;
  actual: {
    google: number;
    yelp: number;
    repugen: number;
  };
  dailyAverage: {
    google: number | null;
    yelp: number | null;
    repugen: number | null;
  };
  ratingCounts: {
    google: Record<'1' | '2' | '3' | '4' | '5', number>;
    yelp: Record<'1' | '2' | '3' | '4' | '5', number>;
    repugen: {
      '1-2': number;
      '3': number;
      '4-5': number;
    };
  };
  facebook: {
    recommended: string;
    notRecommended: string;
  };
};

type ReportDoc = {
  reviews?: {
    rows?: StoredTableRow[];
    facebookRecommended?: string;
    facebookNotRecommended?: string;
    note?: string;
    summary?: ReviewSummarySnapshot;
  };
  issues?: {
    rows?: StoredTableRow[];
    monthly?: string;
    yearly?: string;
    freeDays?: string;
    note?: string;
  };
  equipment?: {
    rows?: StoredTableRow[];
    note?: string;
  };
};

function getFieldString(value: unknown): string {
  if (value === undefined || value === null) return '';
  return String(value).trim();
}

const REVIEW_ROW_HIGHLIGHT_COLOR = '#fef9c3';
const REVIEW_ROW_DELETED_COLOR = '#fee2e2';

function getReviewRowBackground(row: TableRow): string | undefined {
  if (row.deletedReview) return REVIEW_ROW_DELETED_COLOR;
  if (row.highlighted) return REVIEW_ROW_HIGHLIGHT_COLOR;
  return undefined;
}

function normalizeRows(raw: unknown, columnCount: number): TableRow[] {
  if (!Array.isArray(raw)) return [];
  return raw
    .map((item) => {
      const rawCells = Array.isArray((item as StoredTableRow).cells)
        ? (item as StoredTableRow).cells!
        : [];
      const cells = Array.from({ length: columnCount }, (_, index) => getFieldString(rawCells[index]));
      const highlighted = !!(item as StoredTableRow).highlighted;
      return {
        id: crypto.randomUUID(),
        cells,
        ...(highlighted ? { highlighted: true } : {}),
      };
    })
    .filter((row) => row.cells.some((cell) => cell.trim() !== ''));
}

function migrateLegacyReviewRowCells(rawCells: string[]): string[] {
  if (rawCells.length !== LEGACY_REVIEW_COLUMN_COUNT) return rawCells;
  return [
    rawCells[0] ?? '',
    rawCells[1] ?? '',
    '',
    rawCells[2] ?? '',
    rawCells[3] ?? '',
    rawCells[4] ?? '',
  ];
}

function normalizeReviewRows(raw: unknown): TableRow[] {
  if (!Array.isArray(raw)) return [];
  return raw
    .map((item) => {
      const rawCells = Array.isArray((item as StoredTableRow).cells)
        ? migrateLegacyReviewRowCells((item as StoredTableRow).cells!.map((cell) => getFieldString(cell)))
        : [];
      const cells = Array.from({ length: COLUMN_COUNT }, (_, index) => getFieldString(rawCells[index]));
      const highlighted = !!(item as StoredTableRow).highlighted;
      const deletedReview = !!(item as StoredTableRow).deletedReview;
      return {
        id: crypto.randomUUID(),
        cells,
        ...(highlighted ? { highlighted: true } : {}),
        ...(deletedReview ? { deletedReview: true } : {}),
      };
    })
    .filter((row) => row.cells.some((cell) => cell.trim() !== ''));
}

function serializeRows(rows: TableRow[]): StoredTableRow[] {
  return rows.map(({ cells, highlighted, deletedReview }) => {
    const row: StoredTableRow = { cells: [...cells] };
    if (highlighted) row.highlighted = true;
    if (deletedReview) row.deletedReview = true;
    return row;
  });
}

function cloneRows(rows: TableRow[]): TableRow[] {
  return rows.map((row) => ({ ...row, cells: [...row.cells] }));
}

function filterNonEmptyRows(rows: TableRow[]): TableRow[] {
  return rows.filter((row) => row.cells.some((cell) => cell.trim() !== ''));
}

function normalizeDocMonth(dateValue: unknown): string {
  const match = String(dateValue ?? '').trim().match(/^(\d{4})-(\d{1,2})/);
  if (!match) return '';
  return `${match[1]}-${String(Number(match[2])).padStart(2, '0')}`;
}

type BlackBearDoc = {
  date?: string;
  location?: string;
  submittedDateTime?: string;
};

function parseDocDate(dateValue: unknown): Date | null {
  const raw = String(dateValue ?? '').trim();
  const match = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]) - 1;
  const day = Number(match[3]);
  const parsed = new Date(year, month, day);
  if (
    parsed.getFullYear() !== year ||
    parsed.getMonth() !== month ||
    parsed.getDate() !== day
  ) {
    return null;
  }
  return parsed;
}

function formatDisplayDate(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function countBlackBearDays(docs: BlackBearDoc[], month: string, office: string): number {
  const dateDocMap = new Map<string, string>();

  docs
    .filter(
      (doc) =>
        !!doc.submittedDateTime &&
        normalizeDocMonth(doc.date) === month &&
        String(doc.location ?? '').trim() === office
    )
    .forEach((doc) => {
      const parsed = parseDocDate(doc.date);
      if (!parsed) return;
      const key = formatDisplayDate(parsed);
      const existing = dateDocMap.get(key);
      if (!existing || (doc.submittedDateTime ?? '') > existing) {
        dateDocMap.set(key, doc.submittedDateTime ?? '');
      }
    });

  return dateDocMap.size;
}

function getMonthlyLocationFieldKey(location: string): string {
  const raw = String(location ?? '').trim();
  return /[.#$/\[\]]/.test(raw) ? raw.replace(/[.#$/\[\]]/g, '_') : raw;
}

function getReportDocId(month: string, office: string): string {
  const monthId = normalizeDocMonth(month) || String(month ?? '').trim();
  return `${monthId}_${getMonthlyLocationFieldKey(office)}`;
}

async function loadReportFromDb(month: string, office: string): Promise<ReportDoc | null> {
  const docIds = [getReportDocId(month, office), `${normalizeDocMonth(month) || month}_${office.trim()}`];
  const seen = new Set<string>();

  for (const docId of docIds) {
    if (!docId || seen.has(docId)) continue;
    seen.add(docId);
    const snap = await getDoc(doc(db, PINK_BEAR_COLLECTION, docId));
    if (!snap.exists()) continue;
    return snap.data() as ReportDoc;
  }

  return null;
}

function createEmptyRow(columnCount: number): TableRow {
  return {
    id: crypto.randomUUID(),
    cells: Array.from({ length: columnCount }, () => ''),
  };
}

function parseClipboardTable(text: string): string[][] {
  const normalized = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const lines = normalized.split('\n');
  if (lines.length > 0 && lines[lines.length - 1] === '') {
    lines.pop();
  }
  return lines.map((line) => line.split('\t'));
}

function applyPasteToRows(
  rows: TableRow[],
  startRowIndex: number,
  startColIndex: number,
  pastedRows: string[][],
  columnCount: number
): TableRow[] {
  const rowCountNeeded = startRowIndex + pastedRows.length;
  const nextRows = rows.map((row) => ({ ...row, cells: [...row.cells] }));

  while (nextRows.length < rowCountNeeded) {
    nextRows.push(createEmptyRow(columnCount));
  }

  pastedRows.forEach((pastedRow, rowOffset) => {
    const targetRowIndex = startRowIndex + rowOffset;
    pastedRow.forEach((value, colOffset) => {
      const targetColIndex = startColIndex + colOffset;
      if (targetColIndex < columnCount) {
        nextRows[targetRowIndex].cells[targetColIndex] = value;
      }
    });
  });

  return nextRows;
}

function updateRowCell(
  rows: TableRow[],
  rowIndex: number,
  colIndex: number,
  value: string
): TableRow[] {
  return rows.map((row, index) =>
    index === rowIndex
      ? { ...row, cells: row.cells.map((cell, cellIndex) => (cellIndex === colIndex ? value : cell)) }
      : row
  );
}

function deleteRowAt(rows: TableRow[], rowIndex: number): TableRow[] {
  return rows.filter((_, index) => index !== rowIndex);
}

function createCellChangeHandler(setRows: React.Dispatch<React.SetStateAction<TableRow[]>>) {
  return (rowIndex: number, colIndex: number, value: string) => {
    setRows((prev) => updateRowCell(prev, rowIndex, colIndex, value));
  };
}

function createDeleteRowHandler(setRows: React.Dispatch<React.SetStateAction<TableRow[]>>) {
  return (rowIndex: number) => {
    setRows((prev) => deleteRowAt(prev, rowIndex));
  };
}

function createAddRowHandler(
  columnCount: number,
  setRows: React.Dispatch<React.SetStateAction<TableRow[]>>
) {
  return () => {
    setRows((prev) => [...prev, createEmptyRow(columnCount)]);
  };
}

function createPasteHandler(
  columnCount: number,
  setRows: React.Dispatch<React.SetStateAction<TableRow[]>>
) {
  return (event: React.ClipboardEvent<HTMLInputElement>, rowIndex: number, colIndex: number) => {
    const text = event.clipboardData.getData('text/plain');
    const isMultiCell = text.includes('\t') || text.includes('\n') || text.includes('\r');
    if (!isMultiCell) return;

    event.preventDefault();
    const pastedRows = parseClipboardTable(text);
    if (pastedRows.length === 0) return;

    setRows((prev) => applyPasteToRows(prev, rowIndex, colIndex, pastedRows, columnCount));
  };
}

async function saveReportFields(
  month: string,
  office: string,
  fields: Record<string, unknown>
): Promise<void> {
  await setDoc(doc(db, PINK_BEAR_COLLECTION, getReportDocId(month, office)), fields, { merge: true });
}

type EditableDataTableProps = {
  columnCount: number;
  columnHeaders: string[];
  columnMinWidths?: (number | undefined)[];
  rows: TableRow[];
  isEditing: boolean;
  onCellChange: (rowIndex: number, colIndex: number, value: string) => void;
  onPaste: (event: React.ClipboardEvent<HTMLInputElement>, rowIndex: number, colIndex: number) => void;
  onDeleteRow: (rowIndex: number) => void;
  emptyEditMessage: string;
  getRowBackground?: (row: TableRow) => string | undefined;
  enableRowHighlight?: boolean;
  onToggleHighlight?: (rowIndex: number) => void;
  onToggleDeletedReview?: (rowIndex: number) => void;
  getColumnInputType?: (colIndex: number) => 'text' | 'date';
  actionColumnWidth?: number;
  disableHorizontalScroll?: boolean;
};

function EditableDataTable({
  columnCount,
  columnHeaders,
  columnMinWidths,
  rows,
  isEditing,
  onCellChange,
  onPaste,
  onDeleteRow,
  emptyEditMessage,
  getRowBackground,
  enableRowHighlight = false,
  onToggleHighlight,
  onToggleDeletedReview,
  getColumnInputType,
  actionColumnWidth = REVIEW_ACTION_COLUMN_WIDTH,
  disableHorizontalScroll = false,
}: EditableDataTableProps) {
  if (!isEditing && rows.length === 0) return null;

  const dataMinWidth = columnMinWidths
    ? columnMinWidths.reduce<number>((sum, width) => sum + (width ?? 0), 0)
    : 480;
  const tableMinWidth = dataMinWidth + (isEditing && enableRowHighlight ? actionColumnWidth : 0);

  return (
    <div
      style={{
        overflowX: disableHorizontalScroll ? 'visible' : 'auto',
        border: '1px solid #e5e7eb',
        borderRadius: 8,
      }}
    >
      <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: tableMinWidth || 480 }}>
        <thead>
          <tr>
            {columnHeaders.map((header, index) => (
              <th
                key={header}
                style={{
                  borderBottom: '1px solid #e5e7eb',
                  borderRight: index < columnCount - 1 ? '1px solid #e5e7eb' : undefined,
                  padding: '10px 12px',
                  background: '#f8fafc',
                  textAlign: 'left',
                  fontWeight: 700,
                  fontSize: 14,
                  minWidth: columnMinWidths?.[index],
                }}
              >
                {header}
              </th>
            ))}
            {isEditing && (
              <th
                style={{
                  borderBottom: '1px solid #e5e7eb',
                  padding: '10px 12px',
                  background: '#f8fafc',
                  textAlign: 'left',
                  fontWeight: 700,
                  fontSize: 14,
                  width: enableRowHighlight ? actionColumnWidth : 80,
                  minWidth: enableRowHighlight ? actionColumnWidth : 80,
                }}
              >
                Action
              </th>
            )}
          </tr>
        </thead>
        <tbody>
          {rows.length === 0 && isEditing && (
            <tr>
              <td
                colSpan={columnCount + 1}
                style={{
                  padding: '16px 12px',
                  color: '#64748b',
                  textAlign: 'center',
                }}
              >
                {emptyEditMessage}
              </td>
            </tr>
          )}
          {rows.map((row, rowIndex) => {
            const rowBackground = getRowBackground?.(row)
              ?? (enableRowHighlight ? getReviewRowBackground(row) : undefined);

            return (
            <tr key={row.id}>
              {row.cells.map((cell, colIndex) => (
                <td
                  key={colIndex}
                  style={{
                    borderBottom: '1px solid #e5e7eb',
                    borderRight: colIndex < columnCount - 1 ? '1px solid #e5e7eb' : undefined,
                    padding: isEditing ? '6px 8px' : '8px 12px',
                    color: cell ? '#111827' : '#cbd5e1',
                    minWidth: columnMinWidths?.[colIndex],
                    backgroundColor: rowBackground,
                  }}
                >
                  {isEditing ? (
                    <input
                      type={getColumnInputType?.(colIndex) ?? 'text'}
                      value={cell}
                      onChange={(e) => onCellChange(rowIndex, colIndex, e.target.value)}
                      onPaste={(e) => onPaste(e, rowIndex, colIndex)}
                      style={{
                        ...dataCellInputStyle,
                        backgroundColor: rowBackground ? 'transparent' : '#fff',
                      }}
                    />
                  ) : (
                    cell || '—'
                  )}
                </td>
              ))}
              {isEditing && (
                <td
                  style={{
                    borderBottom: '1px solid #e5e7eb',
                    padding: '6px 8px',
                    backgroundColor: rowBackground,
                  }}
                >
                  <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                    {enableRowHighlight && (
                      <button
                        type="button"
                        onClick={() => onToggleHighlight?.(rowIndex)}
                        style={{
                          border: row.highlighted ? '1px solid #facc15' : '1px solid #e5e7eb',
                          borderRadius: 6,
                          padding: '6px 10px',
                          background: row.highlighted ? '#fef08a' : '#fff',
                          color: '#111827',
                          fontSize: 13,
                          cursor: 'pointer',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        {row.highlighted ? 'Unmark' : 'Mentioned'}
                      </button>
                    )}
                    {enableRowHighlight && (
                      <button
                        type="button"
                        onClick={() => onToggleDeletedReview?.(rowIndex)}
                        style={{
                          border: row.deletedReview ? '1px solid #f87171' : '1px solid #fecaca',
                          borderRadius: 6,
                          padding: '6px 10px',
                          background: row.deletedReview ? '#fca5a5' : '#fff',
                          color: row.deletedReview ? '#7f1d1d' : '#b91c1c',
                          fontSize: 13,
                          cursor: 'pointer',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        {row.deletedReview ? 'Unmark' : 'Deleted Review'}
                      </button>
                    )}
                    <button
                      type="button"
                      onClick={() => onDeleteRow(rowIndex)}
                      style={{
                        border: '1px solid #fecaca',
                        borderRadius: 6,
                        padding: '6px 10px',
                        background: '#fff',
                        color: '#b91c1c',
                        fontSize: 13,
                        cursor: 'pointer',
                        whiteSpace: 'nowrap',
                      }}
                    >
                      Delete
                    </button>
                  </div>
                </td>
              )}
            </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

const summaryCellStyle: React.CSSProperties = {
  border: '1px solid #e5e7eb',
  padding: '10px 12px',
  fontSize: 14,
};

function countActualByPlatform(rows: TableRow[]) {
  let google = 0;
  let yelp = 0;
  let repugen = 0;

  rows.forEach((row) => {
    const value = String(row.cells[0] ?? '').trim().toLowerCase();
    if (value === 'google') google += 1;
    else if (value === 'yelp') yelp += 1;
    else if (value === 'repugen') repugen += 1;
  });

  return { google, yelp, repugen };
}

type ReviewPlatform = 'google' | 'yelp' | 'repugen';

function normalizeCellValue(value: unknown): string {
  const raw = String(value ?? '').trim();
  if (!raw) return '';
  const num = Number(raw);
  if (Number.isFinite(num)) return String(Math.trunc(num));
  return raw;
}

function countByPlatformAndRating(rows: TableRow[], platform: ReviewPlatform, rating: string): number {
  return rows.filter((row) => {
    if (row.deletedReview) return false;
    const column1 = String(row.cells[0] ?? '').trim().toLowerCase();
    const column2 = normalizeCellValue(row.cells[1]);
    return column1 === platform && column2 === rating;
  }).length;
}

function countRepugenByColumn2Label(rows: TableRow[], label: string): number {
  return rows.filter((row) => {
    if (row.deletedReview) return false;
    const column1 = String(row.cells[0] ?? '').trim().toLowerCase();
    const column2 = String(row.cells[1] ?? '').trim();
    return column1 === 'repugen' && column2 === label;
  }).length;
}

const countCellStyle: React.CSSProperties = {
  ...summaryCellStyle,
  textAlign: 'center',
  fontWeight: 600,
  color: '#111827',
};

const ratingLabelCellStyle: React.CSSProperties = {
  ...summaryCellStyle,
  fontWeight: 600,
  textAlign: 'center',
};

function RatingRows({ rows }: { rows: TableRow[] }) {
  const repugenLowCount = countRepugenByColumn2Label(rows, '1-2');
  const repugenMidCount = countByPlatformAndRating(rows, 'repugen', '3');
  const repugenHighCount = countRepugenByColumn2Label(rows, '4-5');

  return (
    <>
      <tr>
        <td style={ratingLabelCellStyle}>1</td>
        <td style={countCellStyle}>{countByPlatformAndRating(rows, 'google', '1')}</td>
        <td style={countCellStyle}>{countByPlatformAndRating(rows, 'yelp', '1')}</td>
        <td rowSpan={2} style={countCellStyle}>
          {repugenLowCount}
        </td>
      </tr>
      <tr>
        <td style={ratingLabelCellStyle}>2</td>
        <td style={countCellStyle}>{countByPlatformAndRating(rows, 'google', '2')}</td>
        <td style={countCellStyle}>{countByPlatformAndRating(rows, 'yelp', '2')}</td>
      </tr>
      <tr>
        <td style={ratingLabelCellStyle}>3</td>
        <td style={countCellStyle}>{countByPlatformAndRating(rows, 'google', '3')}</td>
        <td style={countCellStyle}>{countByPlatformAndRating(rows, 'yelp', '3')}</td>
        <td style={countCellStyle}>{repugenMidCount}</td>
      </tr>
      <tr>
        <td style={ratingLabelCellStyle}>4</td>
        <td style={countCellStyle}>{countByPlatformAndRating(rows, 'google', '4')}</td>
        <td style={countCellStyle}>{countByPlatformAndRating(rows, 'yelp', '4')}</td>
        <td rowSpan={2} style={countCellStyle}>
          {repugenHighCount}
        </td>
      </tr>
      <tr>
        <td style={ratingLabelCellStyle}>5</td>
        <td style={countCellStyle}>{countByPlatformAndRating(rows, 'google', '5')}</td>
        <td style={countCellStyle}>{countByPlatformAndRating(rows, 'yelp', '5')}</td>
      </tr>
    </>
  );
}

function computeDailyAverage(actual: number, dayCount: number | null): string {
  const value = computeDailyAverageNumber(actual, dayCount);
  if (value === null) return '—';
  return Number.isInteger(value) ? String(value) : value.toString();
}

function computeDailyAverageNumber(actual: number, dayCount: number | null): number | null {
  if (dayCount === null || dayCount === 0) return null;
  return Math.round((actual / dayCount) * 10) / 10;
}

function buildReviewSummary(
  rows: TableRow[],
  reviewDayCount: number | null,
  facebookRecommended: string,
  facebookNotRecommended: string
): ReviewSummarySnapshot {
  const actual = countActualByPlatform(rows);
  const monthlyGoal = reviewDayCount !== null ? REVIEW_DAILY_GOAL * reviewDayCount : null;

  return {
    dailyGoal: REVIEW_DAILY_GOAL,
    monthlyGoal,
    reviewDayCount,
    actual,
    dailyAverage: {
      google: computeDailyAverageNumber(actual.google, reviewDayCount),
      yelp: computeDailyAverageNumber(actual.yelp, reviewDayCount),
      repugen: computeDailyAverageNumber(actual.repugen, reviewDayCount),
    },
    ratingCounts: {
      google: {
        '1': countByPlatformAndRating(rows, 'google', '1'),
        '2': countByPlatformAndRating(rows, 'google', '2'),
        '3': countByPlatformAndRating(rows, 'google', '3'),
        '4': countByPlatformAndRating(rows, 'google', '4'),
        '5': countByPlatformAndRating(rows, 'google', '5'),
      },
      yelp: {
        '1': countByPlatformAndRating(rows, 'yelp', '1'),
        '2': countByPlatformAndRating(rows, 'yelp', '2'),
        '3': countByPlatformAndRating(rows, 'yelp', '3'),
        '4': countByPlatformAndRating(rows, 'yelp', '4'),
        '5': countByPlatformAndRating(rows, 'yelp', '5'),
      },
      repugen: {
        '1-2': countRepugenByColumn2Label(rows, '1-2'),
        '3': countByPlatformAndRating(rows, 'repugen', '3'),
        '4-5': countRepugenByColumn2Label(rows, '4-5'),
      },
    },
    facebook: {
      recommended: facebookRecommended,
      notRecommended: facebookNotRecommended,
    },
  };
}

function prepareEditRows(rows: TableRow[], columnCount: number): TableRow[] {
  return rows.length > 0 ? cloneRows(rows) : [createEmptyRow(columnCount)];
}

function SummaryTable({
  rows,
  isEditing,
  facebookRecommended,
  facebookNotRecommended,
  onFacebookRecommendedChange,
  onFacebookNotRecommendedChange,
  reviewDayCount,
}: {
  rows: TableRow[];
  isEditing: boolean;
  facebookRecommended: string;
  facebookNotRecommended: string;
  onFacebookRecommendedChange: (value: string) => void;
  onFacebookNotRecommendedChange: (value: string) => void;
  reviewDayCount: number | null;
}) {
  const actual = countActualByPlatform(rows);
  const monthlyGoal = reviewDayCount !== null ? REVIEW_DAILY_GOAL * reviewDayCount : null;

  return (
    <div style={{ overflowX: 'auto', border: '1px solid #e5e7eb', borderRadius: 8, flexShrink: 0, minWidth: 480 }}>
      <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 480 }}>
        <tbody>
          <tr>
            <td style={{ ...summaryCellStyle, background: '#f8fafc' }} />
            <td style={{ ...summaryCellStyle, background: '#f8fafc', fontWeight: 700, textAlign: 'center' }}>
              Google
            </td>
            <td style={{ ...summaryCellStyle, background: '#f8fafc', fontWeight: 700, textAlign: 'center' }}>
              Yelp
            </td>
            <td style={{ ...summaryCellStyle, background: '#f8fafc', fontWeight: 700, textAlign: 'center' }}>
              RepuGen
            </td>
          </tr>
          <tr>
            <td style={{ ...summaryCellStyle, fontWeight: 700 }}>Daily Goal</td>
            <td
              colSpan={3}
              style={{ ...summaryCellStyle, textAlign: 'center', fontWeight: 600, color: '#111827' }}
            >
              {REVIEW_DAILY_GOAL}
            </td>
          </tr>
          <tr>
            <td style={{ ...summaryCellStyle, fontWeight: 700 }}>Monthly Goal</td>
            <td
              colSpan={3}
              style={{
                ...summaryCellStyle,
                textAlign: 'center',
                fontWeight: monthlyGoal !== null ? 600 : 400,
                color: monthlyGoal !== null ? '#111827' : '#cbd5e1',
              }}
            >
              {monthlyGoal !== null ? monthlyGoal : '—'}
            </td>
          </tr>
          <tr>
            <td style={{ ...summaryCellStyle, fontWeight: 700 }}>Actual</td>
            <td style={{ ...summaryCellStyle, textAlign: 'center', fontWeight: 600, color: '#111827' }}>
              {actual.google}
            </td>
            <td style={{ ...summaryCellStyle, textAlign: 'center', fontWeight: 600, color: '#111827' }}>
              {actual.yelp}
            </td>
            <td style={{ ...summaryCellStyle, textAlign: 'center', fontWeight: 600, color: '#111827' }}>
              {actual.repugen}
            </td>
          </tr>
          <tr>
            <td style={{ ...summaryCellStyle, fontWeight: 700 }}>Daily Average</td>
            <td
              style={{
                ...summaryCellStyle,
                textAlign: 'center',
                fontWeight: reviewDayCount ? 600 : 400,
                color: reviewDayCount ? '#111827' : '#cbd5e1',
              }}
            >
              {computeDailyAverage(actual.google, reviewDayCount)}
            </td>
            <td
              style={{
                ...summaryCellStyle,
                textAlign: 'center',
                fontWeight: reviewDayCount ? 600 : 400,
                color: reviewDayCount ? '#111827' : '#cbd5e1',
              }}
            >
              {computeDailyAverage(actual.yelp, reviewDayCount)}
            </td>
            <td
              style={{
                ...summaryCellStyle,
                textAlign: 'center',
                fontWeight: reviewDayCount ? 600 : 400,
                color: reviewDayCount ? '#111827' : '#cbd5e1',
              }}
            >
              {computeDailyAverage(actual.repugen, reviewDayCount)}
            </td>
          </tr>
          <RatingRows rows={rows} />
          <tr>
            <td
              colSpan={4}
              style={{
                ...summaryCellStyle,
                textAlign: 'center',
                fontWeight: 700,
                background: '#f8fafc',
              }}
            >
              Facebook
            </td>
          </tr>
          <tr>
            <td style={{ ...summaryCellStyle, fontWeight: 700 }}>Recommended</td>
            <td style={{ ...summaryCellStyle, padding: isEditing ? '6px 8px' : '10px 12px' }}>
              {isEditing ? (
                <input
                  type="text"
                  value={facebookRecommended}
                  onChange={(e) => onFacebookRecommendedChange(e.target.value)}
                  style={dataCellInputStyle}
                />
              ) : (
                <span style={{ color: facebookRecommended ? '#111827' : '#cbd5e1' }}>
                  {facebookRecommended || '—'}
                </span>
              )}
            </td>
            <td style={{ ...summaryCellStyle, fontWeight: 700, textAlign: 'center' }}>Not Recommended</td>
            <td style={{ ...summaryCellStyle, padding: isEditing ? '6px 8px' : '10px 12px' }}>
              {isEditing ? (
                <input
                  type="text"
                  value={facebookNotRecommended}
                  onChange={(e) => onFacebookNotRecommendedChange(e.target.value)}
                  style={dataCellInputStyle}
                />
              ) : (
                <span style={{ color: facebookNotRecommended ? '#111827' : '#cbd5e1' }}>
                  {facebookNotRecommended || '—'}
                </span>
              )}
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

function IssueSideTable({
  issueTableRows,
  isEditing,
  yearly,
  freeDays,
  onYearlyChange,
  onFreeDaysChange,
}: {
  issueTableRows: TableRow[];
  isEditing: boolean;
  yearly: string;
  freeDays: string;
  onYearlyChange: (value: string) => void;
  onFreeDaysChange: (value: string) => void;
}) {
  const monthlyIssueCount = filterNonEmptyRows(issueTableRows).length;

  const renderValueCell = (label: string, value: string, onChange?: (value: string) => void) => {
    const isEditable = isEditing && onChange;

    return (
      <td
        style={{
          ...summaryCellStyle,
          minWidth: 80,
          padding: isEditable ? '6px 8px' : '10px 12px',
          textAlign: 'center',
          fontWeight: value ? 600 : 400,
          color: value || label === 'Monthly' ? '#111827' : '#cbd5e1',
        }}
      >
        {isEditable ? (
          <input type="text" value={value} onChange={(e) => onChange(e.target.value)} style={centeredDataCellInputStyle} />
        ) : (
          value || (label === 'Monthly' ? '0' : '—')
        )}
      </td>
    );
  };

  return (
    <div
      style={{
        overflowX: 'auto',
        border: '1px solid #e5e7eb',
        borderRadius: 8,
        flexShrink: 0,
        alignSelf: 'flex-start',
        width: 'fit-content',
      }}
    >
      <table style={{ borderCollapse: 'collapse', width: '100%' }}>
        <tbody>
          <tr>
            <td
              colSpan={2}
              style={{
                ...summaryCellStyle,
                textAlign: 'center',
                fontWeight: 700,
                background: '#f8fafc',
              }}
            >
              Injury
            </td>
          </tr>
          <tr>
            <td style={{ ...summaryCellStyle, fontWeight: 700, whiteSpace: 'nowrap' }}>Monthly</td>
            {renderValueCell('Monthly', String(monthlyIssueCount))}
          </tr>
          <tr>
            <td style={{ ...summaryCellStyle, fontWeight: 700, whiteSpace: 'nowrap' }}>Yearly</td>
            {renderValueCell('Yearly', yearly, onYearlyChange)}
          </tr>
          <tr>
            <td style={{ ...summaryCellStyle, fontWeight: 700, whiteSpace: 'nowrap' }}>Injury Free Days</td>
            {renderValueCell('Injury Free Days', freeDays, onFreeDaysChange)}
          </tr>
        </tbody>
      </table>
    </div>
  );
}

function NotePanel({
  title,
  note,
  isEditing,
  onChange,
  fillHeight = false,
}: {
  title: string;
  note: string;
  isEditing: boolean;
  onChange: (value: string) => void;
  fillHeight?: boolean;
}) {
  return (
    <div
      style={{
        width: '100%',
        ...(fillHeight ? { display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0 } : {}),
      }}
    >
      <div
        style={{
          border: '1px solid #e5e7eb',
          borderRadius: 8,
          overflow: 'hidden',
          ...(fillHeight
            ? { flex: 1, display: 'flex', flexDirection: 'column', minHeight: 0 }
            : { height: '100%' }),
        }}
      >
        <div
          style={{
            padding: '10px 12px',
            background: '#f8fafc',
            fontWeight: 700,
            fontSize: 14,
            borderBottom: '1px solid #e5e7eb',
          }}
        >
          {title}
        </div>
        <div
          style={{
            padding: isEditing ? 8 : '12px',
            ...(fillHeight ? { flex: 1, display: 'flex', flexDirection: 'column', minHeight: 0 } : {}),
          }}
        >
          {isEditing ? (
            <textarea
              value={note}
              onChange={(e) => onChange(e.target.value)}
              style={{
                width: '100%',
                minHeight: fillHeight ? '100%' : 120,
                height: fillHeight ? '100%' : undefined,
                border: '1px solid #d1d5db',
                borderRadius: 6,
                padding: '8px 10px',
                fontSize: 14,
                resize: fillHeight ? 'none' : 'vertical',
                fontFamily: 'inherit',
                lineHeight: 1.5,
                boxSizing: 'border-box',
                flex: fillHeight ? 1 : undefined,
              }}
            />
          ) : (
            <div
              style={{
                fontSize: 14,
                color: note ? '#111827' : '#cbd5e1',
                whiteSpace: 'pre-wrap',
                lineHeight: 1.5,
                minHeight: fillHeight ? '100%' : 24,
                flex: fillHeight ? 1 : undefined,
              }}
            >
              {note || '—'}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

export default function ReportTablePage() {
  return (
    <Suspense fallback={<main style={{ minHeight: '100vh', background: '#fff', padding: 24 }}>Loading...</main>}>
      <ReportTablePageContent />
    </Suspense>
  );
}

function ReportTablePageContent() {
  const searchParams = useSearchParams();
  const selectedMonth = searchParams.get('month') ?? '';
  const selectedOffice = searchParams.get('office') ?? '';
  const hasSelection = selectedMonth !== '' && selectedOffice !== '';
  const [rows, setRows] = useState<TableRow[]>([]);
  const [draftRows, setDraftRows] = useState<TableRow[]>([]);
  const [issueRows, setIssueRows] = useState<TableRow[]>([]);
  const [draftIssueRows, setDraftIssueRows] = useState<TableRow[]>([]);
  const [equipmentRows, setEquipmentRows] = useState<TableRow[]>([]);
  const [draftEquipmentRows, setDraftEquipmentRows] = useState<TableRow[]>([]);
  const [facebookRecommended, setFacebookRecommended] = useState('');
  const [facebookNotRecommended, setFacebookNotRecommended] = useState('');
  const [draftFacebookRecommended, setDraftFacebookRecommended] = useState('');
  const [draftFacebookNotRecommended, setDraftFacebookNotRecommended] = useState('');
  const [issueYearly, setIssueYearly] = useState('');
  const [issueFreeDays, setIssueFreeDays] = useState('');
  const [draftIssueYearly, setDraftIssueYearly] = useState('');
  const [draftIssueFreeDays, setDraftIssueFreeDays] = useState('');
  const [reviewNote, setReviewNote] = useState('');
  const [draftReviewNote, setDraftReviewNote] = useState('');
  const [issueNote, setIssueNote] = useState('');
  const [draftIssueNote, setDraftIssueNote] = useState('');
  const [equipmentNote, setEquipmentNote] = useState('');
  const [draftEquipmentNote, setDraftEquipmentNote] = useState('');
  const [isEditing, setIsEditing] = useState(false);
  const [activeSection, setActiveSection] = useState<ReportSection>('reviews');
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [saveMessage, setSaveMessage] = useState('');
  const [isSaving, setIsSaving] = useState(false);
  const [reviewDayCount, setReviewDayCount] = useState<number | null>(null);

  const resetDraftState = () => {
    setDraftRows([]);
    setDraftIssueRows([]);
    setDraftEquipmentRows([]);
    setDraftFacebookRecommended('');
    setDraftFacebookNotRecommended('');
    setDraftIssueYearly('');
    setDraftIssueFreeDays('');
    setDraftReviewNote('');
    setDraftIssueNote('');
    setDraftEquipmentNote('');
    setIsEditing(false);
  };

  const applyReportData = (data: ReportDoc | null) => {
    const nextRows = normalizeReviewRows(data?.reviews?.rows);
    const nextIssueRows = normalizeRows(data?.issues?.rows, ISSUE_COLUMN_COUNT);
    const nextEquipmentRows = normalizeRows(data?.equipment?.rows, EQUIPMENT_COLUMN_COUNT);
    const nextFacebookRecommended = getFieldString(data?.reviews?.facebookRecommended);
    const nextFacebookNotRecommended = getFieldString(data?.reviews?.facebookNotRecommended);
    const nextIssueYearly = getFieldString(data?.issues?.yearly);
    const nextIssueFreeDays = getFieldString(data?.issues?.freeDays);
    const nextReviewNote = getFieldString(data?.reviews?.note);
    const nextIssueNote = getFieldString(data?.issues?.note);
    const nextEquipmentNote = getFieldString(data?.equipment?.note);

    setRows(nextRows);
    setIssueRows(nextIssueRows);
    setEquipmentRows(nextEquipmentRows);
    setFacebookRecommended(nextFacebookRecommended);
    setFacebookNotRecommended(nextFacebookNotRecommended);
    setIssueYearly(nextIssueYearly);
    setIssueFreeDays(nextIssueFreeDays);
    setReviewNote(nextReviewNote);
    setIssueNote(nextIssueNote);
    setEquipmentNote(nextEquipmentNote);
    resetDraftState();
  };

  useEffect(() => {
    let cancelled = false;

    if (!hasSelection) {
      applyReportData(null);
      setReviewDayCount(null);
      setLoading(false);
      setError('');
      setSaveMessage('');
      return () => {
        cancelled = true;
      };
    }

    const load = async () => {
      try {
        setLoading(true);
        setError('');
        setSaveMessage('');
        setReviewDayCount(null);
        const normalizedMonth = normalizeDocMonth(selectedMonth) || selectedMonth;
        const [data, blackBearSnap] = await Promise.all([
          loadReportFromDb(selectedMonth, selectedOffice),
          getDocs(collection(db, BLACK_BEAR_COLLECTION)),
        ]);
        if (cancelled) return;
        applyReportData(data);
        const blackBearDocs = blackBearSnap.docs.map((d) => d.data() as BlackBearDoc);
        const dayCount = countBlackBearDays(blackBearDocs, normalizedMonth, selectedOffice);
        setReviewDayCount(dayCount);
      } catch (e: unknown) {
        if (cancelled) return;
        const message = e instanceof Error ? e.message : 'Error, please try again.';
        setError(message);
      } finally {
        if (!cancelled) setLoading(false);
      }
    };

    load();
    return () => {
      cancelled = true;
    };
  }, [hasSelection, selectedMonth, selectedOffice]);

  useEffect(() => {
    if (!hasSelection || loading || error || isEditing) return;

    const autoSaveDerivedFields = async () => {
      try {
        await saveReportFields(selectedMonth, selectedOffice, {
          reviews: {
            summary: buildReviewSummary(
              rows,
              reviewDayCount,
              facebookRecommended,
              facebookNotRecommended
            ),
          },
          issues: {
            monthly: String(filterNonEmptyRows(issueRows).length),
          },
        });
      } catch {
      }
    };

    autoSaveDerivedFields();
  }, [
    hasSelection,
    loading,
    error,
    isEditing,
    selectedMonth,
    selectedOffice,
    rows,
    issueRows,
    reviewDayCount,
    facebookRecommended,
    facebookNotRecommended,
  ]);

  const handleEdit = () => {
    setSaveMessage('');
    setDraftRows(prepareEditRows(rows, COLUMN_COUNT));
    setDraftIssueRows(prepareEditRows(issueRows, ISSUE_COLUMN_COUNT));
    setDraftEquipmentRows(prepareEditRows(equipmentRows, EQUIPMENT_COLUMN_COUNT));
    setDraftFacebookRecommended(facebookRecommended);
    setDraftFacebookNotRecommended(facebookNotRecommended);
    setDraftIssueYearly(issueYearly);
    setDraftIssueFreeDays(issueFreeDays);
    setDraftReviewNote(reviewNote);
    setDraftIssueNote(issueNote);
    setDraftEquipmentNote(equipmentNote);
    setIsEditing(true);
  };

  const handleSave = async () => {
    if (!hasSelection) return;

    const nextRows = filterNonEmptyRows(draftRows);
    const nextIssueRows = filterNonEmptyRows(draftIssueRows);
    const nextEquipmentRows = filterNonEmptyRows(draftEquipmentRows);
    const reviewSummary = buildReviewSummary(
      nextRows,
      reviewDayCount,
      draftFacebookRecommended,
      draftFacebookNotRecommended
    );

    try {
      setIsSaving(true);
      setSaveMessage('');
      await saveReportFields(selectedMonth, selectedOffice, {
          month: normalizeDocMonth(selectedMonth) || selectedMonth,
          location: getMonthlyLocationFieldKey(selectedOffice),
          locationName: selectedOffice,
          reviews: {
            rows: serializeRows(nextRows),
            facebookRecommended: draftFacebookRecommended,
            facebookNotRecommended: draftFacebookNotRecommended,
            note: draftReviewNote,
            summary: reviewSummary,
          },
          issues: {
            rows: serializeRows(nextIssueRows),
            monthly: String(filterNonEmptyRows(nextIssueRows).length),
            yearly: draftIssueYearly,
            freeDays: draftIssueFreeDays,
            note: draftIssueNote,
          },
          equipment: {
            rows: serializeRows(nextEquipmentRows),
            note: draftEquipmentNote,
          },
          updatedAt: serverTimestamp(),
        });

      setRows(nextRows);
      setIssueRows(nextIssueRows);
      setEquipmentRows(nextEquipmentRows);
      setFacebookRecommended(draftFacebookRecommended);
      setFacebookNotRecommended(draftFacebookNotRecommended);
      setIssueYearly(draftIssueYearly);
      setIssueFreeDays(draftIssueFreeDays);
      setReviewNote(draftReviewNote);
      setIssueNote(draftIssueNote);
      setEquipmentNote(draftEquipmentNote);
      setIsEditing(false);
      setSaveMessage('Saved!');
    } catch (e: unknown) {
      const message = e instanceof Error ? e.message : 'Error, please save again.';
      setSaveMessage(`Failed saving.: ${message}`);
    } finally {
      setIsSaving(false);
    }
  };

  const handleCancel = () => {
    setSaveMessage('');
    setDraftRows([]);
    setDraftIssueRows([]);
    setDraftEquipmentRows([]);
    setIsEditing(false);
  };

  const handleCellChange = createCellChangeHandler(setDraftRows);
  const handleIssueCellChange = createCellChangeHandler(setDraftIssueRows);
  const handleEquipmentCellChange = createCellChangeHandler(setDraftEquipmentRows);
  const handleDeleteRow = createDeleteRowHandler(setDraftRows);
  const handleDeleteIssueRow = createDeleteRowHandler(setDraftIssueRows);
  const handleDeleteEquipmentRow = createDeleteRowHandler(setDraftEquipmentRows);

  const handleToggleReviewHighlight = (rowIndex: number) => {
    setDraftRows((prev) =>
      prev.map((row, index) => {
        if (index !== rowIndex) return row;
        const highlighted = !row.highlighted;
        return {
          ...row,
          highlighted,
          ...(highlighted ? { deletedReview: false } : {}),
        };
      })
    );
  };

  const handleToggleDeletedReview = (rowIndex: number) => {
    setDraftRows((prev) =>
      prev.map((row, index) => {
        if (index !== rowIndex) return row;
        const deletedReview = !row.deletedReview;
        return {
          ...row,
          deletedReview,
          ...(deletedReview ? { highlighted: false } : {}),
        };
      })
    );
  };

  const handleAddRow = createAddRowHandler(COLUMN_COUNT, setDraftRows);
  const handleAddIssueRow = createAddRowHandler(ISSUE_COLUMN_COUNT, setDraftIssueRows);
  const handleAddEquipmentRow = createAddRowHandler(EQUIPMENT_COLUMN_COUNT, setDraftEquipmentRows);
  const handlePaste = createPasteHandler(COLUMN_COUNT, setDraftRows);
  const handleIssuePaste = createPasteHandler(ISSUE_COLUMN_COUNT, setDraftIssueRows);
  const handleEquipmentPaste = createPasteHandler(EQUIPMENT_COLUMN_COUNT, setDraftEquipmentRows);

  const addRowBySection: Record<ReportSection, () => void> = {
    reviews: handleAddRow,
    issues: handleAddIssueRow,
    equipment: handleAddEquipmentRow,
  };

  const displayRows = isEditing ? draftRows : rows;
  const displayIssueRows = isEditing ? draftIssueRows : issueRows;
  const displayEquipmentRows = isEditing ? draftEquipmentRows : equipmentRows;

  const sectionTitle =
    SECTION_OPTIONS.find((option) => option.value === activeSection)?.label ?? 'Review';

  const sectionMaxWidth =
    activeSection === 'reviews'
      ? REVIEWS_SECTION_MAX_WIDTH
      : activeSection === 'equipment'
        ? 1700
        : 1200;

  return (
    <main style={{ minHeight: '100vh', background: '#fff', padding: 24 }}>
      <section style={{ maxWidth: sectionMaxWidth, margin: '0 auto' }}>
        <div style={{ maxWidth: 1200 }}>
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            marginBottom: 20,
            gap: 12,
            flexWrap: 'wrap',
          }}
        >
          <div>
            <h1 style={{ margin: 0, fontSize: 24, fontWeight: 700 }}>{sectionTitle}</h1>
            {hasSelection && (
              <p style={{ margin: '6px 0 0', color: '#475569', fontSize: 14 }}>
                {selectedMonth} · {selectedOffice}
              </p>
            )}
          </div>
          <div style={{ display: 'flex', gap: 8 }}>
            {!isEditing ? (
              <button
                type="button"
                onClick={handleEdit}
                disabled={!hasSelection || loading || isSaving}
                style={{
                  border: '1px solid #2563eb',
                  borderRadius: 8,
                  padding: '8px 14px',
                  background: !hasSelection || loading || isSaving ? '#93c5fd' : '#2563eb',
                  color: '#fff',
                  fontWeight: 600,
                  cursor: !hasSelection || loading || isSaving ? 'not-allowed' : 'pointer',
                  opacity: !hasSelection || loading || isSaving ? 0.7 : 1,
                }}
              >
                Edit
              </button>
            ) : (
              <>
                <button
                  type="button"
                  onClick={() => addRowBySection[activeSection]()}
                  style={addRowButtonStyle}
                >
                  Add Row
                </button>
                <button
                  type="button"
                  onClick={handleCancel}
                  style={{
                    border: '1px solid #e5e7eb',
                    borderRadius: 8,
                    padding: '8px 14px',
                    background: '#fff',
                    color: '#64748b',
                    fontWeight: 600,
                    cursor: 'pointer',
                  }}
                >
                  Cancel
                </button>
                <button
                  type="button"
                  onClick={handleSave}
                  disabled={isSaving}
                  style={{
                    border: '1px solid #2563eb',
                    borderRadius: 8,
                    padding: '8px 14px',
                    background: '#2563eb',
                    color: '#fff',
                    fontWeight: 600,
                    cursor: isSaving ? 'not-allowed' : 'pointer',
                    opacity: isSaving ? 0.7 : 1,
                  }}
                >
                  {isSaving ? 'Saving...' : 'Save'}
                </button>
              </>
            )}
          </div>
        </div>

        <div
          role="radiogroup"
          aria-label="Report section"
          style={{
            display: 'flex',
            gap: 20,
            marginBottom: 20,
            padding: '10px 14px',
            border: '1px solid #e5e7eb',
            borderRadius: 8,
            background: '#f8fafc',
            width: 'fit-content',
          }}
        >
          {SECTION_OPTIONS.map((option) => (
            <label
              key={option.value}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: 8,
                cursor: 'pointer',
                fontSize: 14,
                fontWeight: activeSection === option.value ? 700 : 500,
                color: activeSection === option.value ? '#2563eb' : '#475569',
              }}
            >
              <input
                type="radio"
                name="report-section"
                value={option.value}
                checked={activeSection === option.value}
                onChange={() => setActiveSection(option.value)}
                style={{ accentColor: '#2563eb' }}
              />
              {option.label}
            </label>
          ))}
        </div>
        </div>

        {saveMessage && (
          <p style={{ margin: '0 0 16px', color: saveMessage.includes('실패') ? '#b91c1c' : '#166534' }}>
            {saveMessage}
          </p>
        )}
        {!hasSelection && <p style={{ margin: '0 0 16px', color: '#6b7280' }}>Month and office details need to be included in the URL.</p>}
        {hasSelection && loading && <p style={{ margin: '0 0 16px', color: '#6b7280' }}>Loading...</p>}
        {hasSelection && error && <p style={{ margin: '0 0 16px', color: '#b91c1c' }}>{error}</p>}

        {hasSelection && !loading && !error && (
        <>
        <div style={{ marginBottom: 20, maxWidth: 1700 }}>
          {activeSection === 'equipment' && (
            <NotePanel
              title="Broken Equipment Note"
              note={isEditing ? draftEquipmentNote : equipmentNote}
              isEditing={isEditing}
              onChange={setDraftEquipmentNote}
            />
          )}
        </div>

        {activeSection === 'reviews' && (
          <div style={{ width: '100%', minWidth: REVIEWS_TABLE_MIN_WIDTH }}>
            <div style={flexSectionRowStyle}>
              <SummaryTable
                rows={displayRows}
                isEditing={isEditing}
                facebookRecommended={isEditing ? draftFacebookRecommended : facebookRecommended}
                facebookNotRecommended={isEditing ? draftFacebookNotRecommended : facebookNotRecommended}
                onFacebookRecommendedChange={setDraftFacebookRecommended}
                onFacebookNotRecommendedChange={setDraftFacebookNotRecommended}
                reviewDayCount={reviewDayCount}
              />
              <div style={{ flex: 1, minWidth: 240, display: 'flex' }}>
                <NotePanel
                  title="Review Note"
                  note={isEditing ? draftReviewNote : reviewNote}
                  isEditing={isEditing}
                  onChange={setDraftReviewNote}
                  fillHeight
                />
              </div>
            </div>

            {(isEditing || rows.length > 0) && (
              <EditableDataTable
                columnCount={COLUMN_COUNT}
                columnHeaders={COLUMN_HEADERS}
                columnMinWidths={DATA_COLUMN_MIN_WIDTHS}
                rows={displayRows}
                isEditing={isEditing}
                onCellChange={handleCellChange}
                onPaste={handlePaste}
                onDeleteRow={handleDeleteRow}
                emptyEditMessage="Please add a row using the Add Row button."
                enableRowHighlight
                onToggleHighlight={handleToggleReviewHighlight}
                onToggleDeletedReview={handleToggleDeletedReview}
                getRowBackground={getReviewRowBackground}
                disableHorizontalScroll
              />
            )}
          </div>
        )}

        {activeSection === 'issues' && (
          <div style={{ width: '100%', minWidth: ISSUES_TABLE_MIN_WIDTH }}>
            <div style={flexSectionRowStyle}>
              <IssueSideTable
                issueTableRows={displayIssueRows}
                isEditing={isEditing}
                yearly={isEditing ? draftIssueYearly : issueYearly}
                freeDays={isEditing ? draftIssueFreeDays : issueFreeDays}
                onYearlyChange={setDraftIssueYearly}
                onFreeDaysChange={setDraftIssueFreeDays}
              />
              <div style={{ flex: 1, minWidth: 240, display: 'flex' }}>
                <NotePanel
                  title="Injury Note"
                  note={isEditing ? draftIssueNote : issueNote}
                  isEditing={isEditing}
                  onChange={setDraftIssueNote}
                  fillHeight
                />
              </div>
            </div>

            <EditableDataTable
              columnCount={ISSUE_COLUMN_COUNT}
              columnHeaders={ISSUE_COLUMN_HEADERS}
              columnMinWidths={ISSUE_COLUMN_MIN_WIDTHS}
              rows={displayIssueRows}
              isEditing={isEditing}
              onCellChange={handleIssueCellChange}
              onPaste={handleIssuePaste}
              onDeleteRow={handleDeleteIssueRow}
              emptyEditMessage="Please add a row using the Add Row button."
            />
          </div>
        )}

        {activeSection === 'equipment' && (
          <div style={{ width: '100%', minWidth: EQUIPMENT_TABLE_MIN_WIDTH }}>
            <EditableDataTable
              columnCount={EQUIPMENT_COLUMN_COUNT}
              columnHeaders={EQUIPMENT_COLUMN_HEADERS}
              columnMinWidths={EQUIPMENT_COLUMN_MIN_WIDTHS}
              rows={displayEquipmentRows}
              isEditing={isEditing}
              onCellChange={handleEquipmentCellChange}
              onPaste={handleEquipmentPaste}
              onDeleteRow={handleDeleteEquipmentRow}
              emptyEditMessage="Please add a row using the Add Row button."
              getRowBackground={getEquipmentRowBackground}
            />
          </div>
        )}
        </>
        )}
      </section>
    </main>
  );
}