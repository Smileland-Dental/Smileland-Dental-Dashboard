'use client';

import React, { useEffect, useMemo, useState } from 'react';
import { useSearchParams } from 'next/navigation';
import { db } from '@/lib/firebase.config';
import { collection, doc, getDoc, getDocs, setDoc } from 'firebase/firestore';

const S_GOAL_COLUMN_HEADERS = ['Day', 'Date', 'OE Done', 'Sealant Goal', 'Sealants Done', 'Difference'];
const ALT_GOAL_COLUMN_HEADERS = ['Day', 'Date', 'OE Goal', 'OE Done', 'Difference'];

type ViewMode = 's-goal' | 'alt-goal';

type CoffeeActualTotals = {
  orangeJuiceTotal?: string;
};

type SugarTotals = {
  sugarGood?: string | number;
};

type FormDoc = {
  id: string;
  date?: string;
  location?: string;
  submittedDateTime?: string;
  coffeeActualTotals?: CoffeeActualTotals;
  sugarTotals?: SugarTotals;
};

type DailyRow = {
  day: string;
  date: string;
  orangeJuiceTotal: string;
  column4: number | null;
  sugarGood: number | null;
  column6: number | null;
};

function normalizeDocMonth(dateValue: unknown): string {
  const raw = String(dateValue ?? '').trim();
  const match = raw.match(/^(\d{4})-(\d{1,2})/);
  if (!match) return '';
  return `${match[1]}-${String(Number(match[2])).padStart(2, '0')}`;
}

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

function formatDayOfWeek(date: Date): string {
  return date.toLocaleDateString('en-US', { weekday: 'short' });
}

function formatDisplayDate(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function parseIntegerValue(value: unknown): number | null {
  const raw = String(value ?? '').trim();
  if (!raw) return null;
  const num = Number(raw);
  if (!Number.isFinite(num)) return null;
  return Math.round(num);
}

function computeColumn4(column3Value: number): number {
  return Math.round(column3Value * 20 * 0.03);
}

function computeColumn6(column4Value: number, sugarGoodValue: number): number {
  return Math.round(sugarGoodValue - column4Value);
}

type OGoalRow = {
  day: string;
  date: string;
  done: string;
};

function getLatestDocsByDate(
  docs: FormDoc[],
  month: string,
  office: string
): { date: Date; doc: FormDoc }[] {
  const dateDocMap = new Map<string, { date: Date; doc: FormDoc }>();

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
      if (!existing || (doc.submittedDateTime ?? '') > (existing.doc.submittedDateTime ?? '')) {
        dateDocMap.set(key, { date: parsed, doc });
      }
    });

  return Array.from(dateDocMap.values()).sort((a, b) => a.date.getTime() - b.date.getTime());
}

function buildOGoalRows(docs: FormDoc[], month: string, office: string): OGoalRow[] {
  return getLatestDocsByDate(docs, month, office).map(({ date, doc }) => {
    const doneValue = parseIntegerValue(doc.coffeeActualTotals?.orangeJuiceTotal);
    return {
      day: formatDayOfWeek(date),
      date: formatDisplayDate(date),
      done: doneValue !== null ? String(doneValue) : '',
    };
  });
}

function computeOGGoalDifference(doneValue: string, goalValue: string): number | null {
  const done = parseIntegerValue(doneValue);
  const goal = parseIntegerValue(goalValue);
  if (done === null || goal === null) return null;
  return done - goal;
}

function sumPositiveSGoalDifferences(rows: DailyRow[]): number {
  return rows.reduce((sum, row) => {
    if (row.column6 !== null && row.column6 > 0) {
      return sum + row.column6;
    }
    return sum;
  }, 0);
}

function sumPositiveOGGoalDifferences(rows: OGoalRow[], goalValue: string): number {
  return rows.reduce((sum, row) => {
    const diff = computeOGGoalDifference(row.done, goalValue);
    if (diff !== null && diff > 0) {
      return sum + diff;
    }
    return sum;
  }, 0);
}

function getWhiteBearDocId(month: string, office: string): string {
  return `${month}_${office}`;
}

async function saveWhiteBearFields(
  month: string,
  office: string,
  fields: Record<string, unknown>
): Promise<void> {
  await setDoc(doc(db, 'monthly production', getWhiteBearDocId(month, office)), fields, { merge: true });
}

function buildGoalColumnStyles(
  thBase: React.CSSProperties,
  tdBase: React.CSSProperties,
  columnCount: number
): { thStyle: React.CSSProperties; cellStyle: React.CSSProperties } {
  const width = `${100 / columnCount}%`;
  return {
    thStyle: { ...thBase, width },
    cellStyle: { ...tdBase, width },
  };
}

function buildGoalFooterStyles(cellStyle: React.CSSProperties): {
  labelStyle: React.CSSProperties;
  valueStyle: React.CSSProperties;
} {
  const baseStyle: React.CSSProperties = {
    ...cellStyle,
    borderTop: '1px solid #e5e7eb',
    fontWeight: 700,
    background: '#f8fafc',
  };
  return {
    labelStyle: baseStyle,
    valueStyle: { ...baseStyle, borderRight: 'none' },
  };
}

function buildDailyRows(docs: FormDoc[], month: string, office: string): DailyRow[] {
  return getLatestDocsByDate(docs, month, office).map(({ date, doc }) => {
    const column3Value = parseIntegerValue(doc.coffeeActualTotals?.orangeJuiceTotal);
    const column4Value = column3Value !== null ? computeColumn4(column3Value) : null;
    const sugarGoodValue = parseIntegerValue(doc.sugarTotals?.sugarGood);
    const column6Value =
      column4Value !== null && sugarGoodValue !== null
        ? computeColumn6(column4Value, sugarGoodValue)
        : null;

    return {
      day: formatDayOfWeek(date),
      date: formatDisplayDate(date),
      orangeJuiceTotal: column3Value !== null ? String(column3Value) : '',
      column4: column4Value,
      sugarGood: sugarGoodValue,
      column6: column6Value,
    };
  });
}

export default function DailyReportPage() {
  const searchParams = useSearchParams();
  const month = searchParams.get('month') ?? '';
  const office = searchParams.get('office') ?? '';

  const [docs, setDocs] = useState<FormDoc[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [saveMessage, setSaveMessage] = useState('');
  const [isEditing, setIsEditing] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [noteSaved, setNoteSaved] = useState('');
  const [noteDraft, setNoteDraft] = useState('');
  const [oGoalSaved, setOGoalSaved] = useState('');
  const [oGoalDraft, setOGoalDraft] = useState('');
  const [oGoalNoteSaved, setOGoalNoteSaved] = useState('');
  const [oGoalNoteDraft, setOGoalNoteDraft] = useState('');
  const [viewMode, setViewMode] = useState<ViewMode>('s-goal');

  useEffect(() => {
    if (!month || !office) {
      setLoading(false);
      setError('');
      return;
    }

    const load = async () => {
      setLoading(true);
      setError('');
      try {
        const snap = await getDocs(collection(db, 'simple-forms'));
        const loaded = snap.docs.map((d) => ({ id: d.id, ...(d.data() as Omit<FormDoc, 'id'>) }));
        setDocs(loaded);
      } catch (e: any) {
        setError(e?.message || '일별 데이터 조회 중 오류가 발생했습니다.');
      } finally {
        setLoading(false);
      }
    };
    load();
  }, [month, office]);

  useEffect(() => {
    if (!month || !office) {
      setNoteSaved('');
      setNoteDraft('');
      setOGoalSaved('');
      setOGoalDraft('');
      setOGoalNoteSaved('');
      setOGoalNoteDraft('');
      return;
    }

    const loadWhiteBear = async () => {
      try {
        const snap = await getDoc(doc(db, 'monthly production', getWhiteBearDocId(month, office)));
        const data = snap.data();
        const loadedNote = String(data?.['sealant note'] ?? '');
        const loadedOGoal = String(data?.['oe goal'] ?? '');
        const loadedOGoalNote = String(data?.['oe goal note'] ?? '');
        setNoteSaved(loadedNote);
        setNoteDraft(loadedNote);
        setOGoalSaved(loadedOGoal);
        setOGoalDraft(loadedOGoal);
        setOGoalNoteSaved(loadedOGoalNote);
        setOGoalNoteDraft(loadedOGoalNote);
      } catch {
        setNoteSaved('');
        setNoteDraft('');
        setOGoalSaved('');
        setOGoalDraft('');
        setOGoalNoteSaved('');
        setOGoalNoteDraft('');
      }
    };

    loadWhiteBear();
  }, [month, office]);

  const syncDraftsFromSaved = () => {
    setNoteDraft(noteSaved);
    setOGoalDraft(oGoalSaved);
    setOGoalNoteDraft(oGoalNoteSaved);
  };

  const handleViewModeChange = (nextMode: ViewMode) => {
    if (isEditing) {
      syncDraftsFromSaved();
      setIsEditing(false);
    }
    setViewMode(nextMode);
  };

  const handleCancelEdit = () => {
    syncDraftsFromSaved();
    setIsEditing(false);
    setSaveMessage('');
  };

  const handleStartEdit = () => {
    syncDraftsFromSaved();
    setIsEditing(true);
    setSaveMessage('');
  };

  const handleSave = async () => {
    if (!month || !office) return;

    setIsSaving(true);
    setSaveMessage('');
    setError('');
    try {
      if (viewMode === 's-goal') {
        await saveWhiteBearFields(month, office, { 'sealant note': noteDraft });
        setNoteSaved(noteDraft);
      } else {
        await saveWhiteBearFields(month, office, {
          'oe goal': oGoalDraft,
          'oe goal note': oGoalNoteDraft,
        });
        setOGoalSaved(oGoalDraft);
        setOGoalNoteSaved(oGoalNoteDraft);
      }
      setIsEditing(false);
      setSaveMessage('Saved!');
    } catch (e: any) {
      setError(e?.message || 'Error. Please try again.');
    } finally {
      setIsSaving(false);
    }
  };

  const rows = useMemo(() => buildDailyRows(docs, month, office), [docs, month, office]);

  const oGoalRows = useMemo(() => buildOGoalRows(docs, month, office), [docs, month, office]);

  const activeNote = isEditing ? noteDraft : noteSaved;
  const activeOGoal = isEditing ? oGoalDraft : oGoalSaved;
  const activeOGoalNote = isEditing ? oGoalNoteDraft : oGoalNoteSaved;

  const differencePositiveTotal = useMemo(
    () => sumPositiveSGoalDifferences(rows),
    [rows]
  );

  const oGoalDifferencePositiveTotal = useMemo(
    () => sumPositiveOGGoalDifferences(oGoalRows, activeOGoal),
    [oGoalRows, activeOGoal]
  );

  const oGoalDifferencePositiveTotalSaved = useMemo(
    () => sumPositiveOGGoalDifferences(oGoalRows, oGoalSaved),
    [oGoalRows, oGoalSaved]
  );

  useEffect(() => {
    if (!month || !office || loading || error) return;
    if (rows.length === 0 && oGoalRows.length === 0) return;

    const saveTotals = async () => {
      try {
        const fields: Record<string, number> = {};
        if (rows.length > 0) {
          fields['sealant goal total'] = differencePositiveTotal;
        }
        if (oGoalRows.length > 0) {
          fields['oe goal total'] = oGoalDifferencePositiveTotalSaved;
        }
        await saveWhiteBearFields(month, office, fields);
      } catch {
        // 저장 실패는 화면 표시 없이 무시
      }
    };

    saveTotals();
  }, [
    month,
    office,
    loading,
    error,
    rows.length,
    oGoalRows.length,
    differencePositiveTotal,
    oGoalDifferencePositiveTotalSaved,
  ]);

  const hasParams = month !== '' && office !== '';
  const showSGoalContent = viewMode === 's-goal' && hasParams && !loading && !error && rows.length > 0;
  const showOGoalContent = viewMode === 'alt-goal' && hasParams && !loading && !error && oGoalRows.length > 0;
  const showEditControls = showSGoalContent || showOGoalContent;

  const valueColor = (hasValue: boolean) => (hasValue ? '#111827' : '#cbd5e1');

  const editButtonStyle: React.CSSProperties = {
    padding: '8px 16px',
    border: 'none',
    borderRadius: 8,
    fontWeight: 600,
    fontSize: 14,
    cursor: 'pointer',
  };

  const tdStyle: React.CSSProperties = {
    borderBottom: '1px solid #e5e7eb',
    borderRight: '1px solid #e5e7eb',
    padding: '8px 12px',
    color: '#111827',
  };

  const inputStyle: React.CSSProperties = {
    width: '100%',
    boxSizing: 'border-box',
    padding: 0,
    border: 'none',
    outline: 'none',
    fontSize: 14,
    color: '#111827',
    background: 'transparent',
  };

  const textareaStyle: React.CSSProperties = {
    ...inputStyle,
    resize: 'vertical',
  };

  const thStyle: React.CSSProperties = {
    borderBottom: '1px solid #e5e7eb',
    borderRight: '1px solid #e5e7eb',
    padding: '10px 12px',
    background: '#f8fafc',
    textAlign: 'left',
    fontWeight: 700,
    fontSize: 14,
  };

  const noteThStyle: React.CSSProperties = {
    ...thStyle,
    borderRight: 'none',
  };

  const goalTableSectionStyle: React.CSSProperties = {
    overflowX: 'auto',
    border: '1px solid #e5e7eb',
    borderRadius: 8,
  };

  const noteSectionStyle: React.CSSProperties = {
    ...goalTableSectionStyle,
    marginBottom: 16,
  };

  const noteTableStyle: React.CSSProperties = {
    width: '100%',
    borderCollapse: 'collapse',
  };

  const noteCellStyle: React.CSSProperties = {
    padding: '8px 12px',
  };

  const noteTextStyle: React.CSSProperties = {
    margin: 0,
    whiteSpace: 'pre-wrap',
    fontSize: 14,
    color: '#111827',
  };

  const saveMessageStyle: React.CSSProperties = {
    margin: '0 0 16px',
    padding: '10px 14px',
    borderRadius: 8,
    fontSize: 14,
    background: '#f0fdf4',
    color: '#166534',
    border: '1px solid #bbf7d0',
  };

  const radioLabelStyle: React.CSSProperties = {
    display: 'flex',
    alignItems: 'center',
    gap: 6,
    cursor: 'pointer',
    fontSize: 14,
  };

  const goalTableStyle: React.CSSProperties = {
    width: '100%',
    borderCollapse: 'collapse',
    tableLayout: 'fixed',
  };

  const { thStyle: sGoalThStyle, cellStyle: sGoalCellStyle } = buildGoalColumnStyles(
    thStyle,
    tdStyle,
    S_GOAL_COLUMN_HEADERS.length
  );

  const { thStyle: oeGoalThStyle, cellStyle: oeGoalCellStyle } = buildGoalColumnStyles(
    thStyle,
    tdStyle,
    ALT_GOAL_COLUMN_HEADERS.length
  );

  const oeGoalFooterStyles = buildGoalFooterStyles(oeGoalCellStyle);
  const sGoalFooterStyles = buildGoalFooterStyles(sGoalCellStyle);

  return (
    <main style={{ minHeight: '100vh', background: '#fff', padding: 24 }}>
      <section style={{ maxWidth: 960, margin: '0 auto' }}>
        <div
          style={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            marginBottom: 20,
          }}
        >
          <h1 style={{ margin: 0, fontSize: 22, fontWeight: 700 }}>Sealant & OE Goal</h1>
          {showEditControls && (
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              {isEditing ? (
                <>
                  <button
                    type="button"
                    onClick={handleSave}
                    disabled={isSaving}
                    style={{
                      ...editButtonStyle,
                      background: isSaving ? '#94a3b8' : '#16a34a',
                      color: '#fff',
                      cursor: isSaving ? 'not-allowed' : 'pointer',
                    }}
                  >
                    {isSaving ? 'Saving...' : 'Save'}
                  </button>
                  <button
                    type="button"
                    onClick={handleCancelEdit}
                    disabled={isSaving}
                    style={{
                      ...editButtonStyle,
                      background: '#fff',
                      color: '#64748b',
                      border: '1px solid #e2e8f0',
                      cursor: isSaving ? 'not-allowed' : 'pointer',
                    }}
                  >
                    Cancel
                  </button>
                </>
              ) : (
                <button
                  type="button"
                  onClick={handleStartEdit}
                  style={{
                    ...editButtonStyle,
                    background: '#2563eb',
                    color: '#fff',
                  }}
                >
                  Edit
                </button>
              )}
            </div>
          )}
        </div>

        {saveMessage && (
          <p style={saveMessageStyle}>
            {saveMessage}
          </p>
        )}

        <div
          role="radiogroup"
          aria-label="Goal view"
          style={{ display: 'flex', gap: 16, marginBottom: 20 }}
        >
          <label style={radioLabelStyle}>
            <input
              type="radio"
              name="goal-view"
              value="s-goal"
              checked={viewMode === 's-goal'}
              onChange={() => handleViewModeChange('s-goal')}
            />
            Sealant Goal
          </label>
          <label style={radioLabelStyle}>
            <input
              type="radio"
              name="goal-view"
              value="alt-goal"
              checked={viewMode === 'alt-goal'}
              onChange={() => handleViewModeChange('alt-goal')}
            />
            OE Goal
          </label>
        </div>

        {hasParams && (
          <p style={{ margin: '0 0 20px', color: '#64748b' }}>
            {month} · {office}
          </p>
        )}

        {hasParams && loading && <p style={{ margin: 0, color: '#6b7280' }}>Loading...</p>}
        {hasParams && error && <p style={{ margin: 0, color: '#b91c1c' }}>{error}</p>}

        {showOGoalContent && (
          <>
            <div style={noteSectionStyle}>
              <table style={noteTableStyle}>
                <thead>
                  <tr>
                    <th style={noteThStyle}>
                      <label htmlFor="o-goal-note" style={{ margin: 0 }}>
                        Note
                      </label>
                    </th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td style={noteCellStyle}>
                      {isEditing ? (
                        <textarea
                          id="o-goal-note"
                          value={activeOGoalNote}
                          onChange={(e) => setOGoalNoteDraft(e.target.value)}
                          rows={3}
                          style={textareaStyle}
                        />
                      ) : (
                        activeOGoalNote ? (
                          <p style={noteTextStyle}>
                            {activeOGoalNote}
                          </p>
                        ) : null
                      )}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
            <div style={goalTableSectionStyle}>
            <table style={goalTableStyle}>
              <thead>
                <tr>
                  {ALT_GOAL_COLUMN_HEADERS.map((header) => (
                    <th key={header} style={oeGoalThStyle}>
                      {header}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {oGoalRows.map((row) => {
                  const difference = computeOGGoalDifference(row.done, activeOGoal);
                  return (
                    <tr key={row.date}>
                      <td style={{ ...oeGoalCellStyle, fontWeight: 600 }}>{row.day}</td>
                      <td style={oeGoalCellStyle}>{row.date}</td>
                      <td style={oeGoalCellStyle}>
                        {isEditing ? (
                          <input
                            type="text"
                            value={activeOGoal}
                            onChange={(e) => setOGoalDraft(e.target.value)}
                            style={inputStyle}
                          />
                        ) : (
                          activeOGoal || '—'
                        )}
                      </td>
                      <td style={{ ...oeGoalCellStyle, color: valueColor(!!row.done) }}>
                        {row.done || '—'}
                      </td>
                      <td
                        style={{
                          ...oeGoalCellStyle,
                          borderRight: 'none',
                          color: valueColor(difference !== null),
                        }}
                      >
                        {difference !== null ? difference : '—'}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr>
                  <td colSpan={4} style={oeGoalFooterStyles.labelStyle}>
                    Total
                  </td>
                  <td style={oeGoalFooterStyles.valueStyle}>
                    {oGoalDifferencePositiveTotal}
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>
          </>
        )}

        {showSGoalContent && (
          <>
            <div style={noteSectionStyle}>
              <table style={noteTableStyle}>
                <thead>
                  <tr>
                    <th style={noteThStyle}>
                      <label htmlFor="goal-note" style={{ margin: 0 }}>
                        Note
                      </label>
                    </th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td style={noteCellStyle}>
                      {isEditing ? (
                        <textarea
                          id="goal-note"
                          value={activeNote}
                          onChange={(e) => setNoteDraft(e.target.value)}
                          rows={3}
                          style={textareaStyle}
                        />
                      ) : (
                        activeNote ? (
                          <p style={noteTextStyle}>
                            {activeNote}
                          </p>
                        ) : null
                      )}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
            <div style={goalTableSectionStyle}>
            <table style={goalTableStyle}>
              <thead>
                <tr>
                  {S_GOAL_COLUMN_HEADERS.map((header) => (
                    <th key={header} style={sGoalThStyle}>
                      {header}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map((row) => (
                  <tr key={row.date}>
                    <td
                      style={{
                        ...sGoalCellStyle,
                        fontWeight: 600,
                        color: '#111827',
                      }}
                    >
                      {row.day}
                    </td>
                    <td style={{ ...sGoalCellStyle, color: '#111827' }}>
                      {row.date}
                    </td>
                    <td style={{ ...sGoalCellStyle, color: valueColor(!!row.orangeJuiceTotal) }}>
                      {row.orangeJuiceTotal || '—'}
                    </td>
                    <td style={{ ...sGoalCellStyle, color: valueColor(row.column4 !== null) }}>
                      {row.column4 !== null ? row.column4 : '—'}
                    </td>
                    <td style={{ ...sGoalCellStyle, color: valueColor(row.sugarGood !== null) }}>
                      {row.sugarGood !== null ? row.sugarGood : '—'}
                    </td>
                    <td
                      style={{
                        ...sGoalCellStyle,
                        borderRight: 'none',
                        color: valueColor(row.column6 !== null),
                      }}
                    >
                      {row.column6 !== null ? row.column6 : '—'}
                    </td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr>
                  <td colSpan={5} style={sGoalFooterStyles.labelStyle}>
                    Total
                  </td>
                  <td style={sGoalFooterStyles.valueStyle}>
                    {differencePositiveTotal}
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>
          </>
        )}
      </section>
    </main>
  );
}