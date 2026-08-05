'use client';

import React, { useEffect, useMemo, useState } from 'react';
import { auth, db } from '@/lib/firebase.config';
import { onAuthStateChanged } from 'firebase/auth';
import { collection, doc, getDoc, getDocs } from 'firebase/firestore';

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

type ExtraInputRow = {
  position?: string;
  name?: string;
  doctorPreventative?: string;
  doctorRestorative?: string;
  doctorCraProduction?: string;
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
  mailedProduction?: string;
};

type ProductionSideMetrics = {
  add?: string;
  noShow?: string;
  scheduled?: string;
  seen?: string;
  seenPercent?: string;
  referral?: string;
  postcard?: string;
};

type TableTotals = {
  sales?: number | string;
  coffeeNew?: number | string;
  coffeeReturn?: number | string;
  coffeeTotal?: number | string;
  coffeeNo?: number | string;
  renderedCoffee?: number | string;
  coffeeYes?: number | string;
  orangeJuiceNew?: number | string;
  orangeJuiceReturn?: number | string;
  orangeJuiceTotal?: number | string;
};

type SugarTotals = {
  sugar?: number | string;
  sugarGood?: number | string;
  sugarBad?: number | string;
  paper?: number | string;
};

type FormDoc = {
  id: string;
  date?: string;
  location?: string;
  reasonIfLate?: string;
  checkIn?: string;
  checkOut?: string;
  hoursOpen?: string;
  closer?: string;
  submittedDateTime?: string;
  grandTotal?: string;
  coffeeSales?: string;
  salesWithoutCoffee?: string;
  paperAtOrangeJuice?: string;
  paperAtTea?: string;
  justPaper?: string;
  prophyTotal?: string;
  notes?: string;
  notDue?: string;
  notesLines?: string[];
  notDueLines?: string[];
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
  extraInputRows?: ExtraInputRow[];
  locationSummary?: LocationSummary;
  productionSideMetrics?: ProductionSideMetrics;
};

const TABLE_HEADERS = [
  'Position',
  'Name',
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

const SUGAR_HEADERS = ['Position', 'Name', 'Sealant', 'Sealant (Billable)', 'Sealant (Redo)', 'Prophy Documented'];

const DOCTOR_HEADERS = [
  'Position',
  'Name',
  'Production',
  'Preventative',
  'Restorative',
  'CRA Production',
  'Patient Seen',
  'Insurance',
  'Cash',
  'Dentical',
  'Treatment',
  'Primary Teeth',
  'Permanent Teeth',
];

const PRODUCTION_SIDE_METRIC_ROWS: { key: keyof ProductionSideMetrics; label: string }[] = [
  { key: 'add', label: 'Add On' },
  { key: 'noShow', label: 'No Shows' },
  { key: 'scheduled', label: 'Scheduled' },
  { key: 'seen', label: 'Seen' },
  { key: 'seenPercent', label: 'Seen %' },
  { key: 'referral', label: 'Referral' },
  { key: 'postcard', label: 'Postcard Count' },
];

const NOTES_MAX_LENGTH = 300;

const D_PAGE_DETAIL_CARD_PADDING_PX = 12;
const dPageDetailTableBleedScroll: React.CSSProperties = {
  marginLeft: -D_PAGE_DETAIL_CARD_PADDING_PX,
  marginRight: -D_PAGE_DETAIL_CARD_PADDING_PX,
  width: `calc(100% + ${D_PAGE_DETAIL_CARD_PADDING_PX * 2}px)`,
  overflowX: 'auto',
};

const FILTER_SELECT_STYLE: React.CSSProperties = {
  width: '100%',
  height: 34,
  border: '1px solid #d1d5db',
  borderRadius: 6,
  padding: '0 10px',
  background: '#fff',
  fontWeight: 600,
};

const thStyle: React.CSSProperties = { border: '1px solid #e5e7eb', padding: 8, textAlign: 'left' };
const tdStyle: React.CSSProperties = { border: '1px solid #e5e7eb', padding: 8 };
const labelCellStyle: React.CSSProperties = {
  border: '1px solid #e5e7eb',
  padding: 8,
  background: '#f9fafb',
  fontWeight: 700,
};

function loadMultilineField(text: unknown, lines: unknown): string {
  if (Array.isArray(lines) && lines.length > 0) {
    return lines.map((line) => String(line)).join('\n');
  }
  return String(text ?? '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
}

function parseMoney(value: unknown): number {
  const s = String(value ?? '')
    .trim()
    .replace(/^\$/, '')
    .replace(/,/g, '')
    .trim();
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function formatRoundedNumber(n: number): string {
  if (!Number.isFinite(n)) return '0';
  const cents = Math.round(n * 100);
  if (cents % 100 === 0) return String(cents / 100);
  return (cents / 100).toFixed(2);
}

function formatNumberWithThousandSeparator(n: number): string {
  const rounded = formatRoundedNumber(n);
  const negative = rounded.startsWith('-');
  const raw = negative ? rounded.slice(1) : rounded;
  const [intPart, decPart] = raw.split('.');
  const withCommas = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const body = decPart != null ? `${withCommas}.${decPart}` : withCommas;
  return negative ? `-${body}` : body;
}

function formatCurrencyLabel(value: unknown): string {
  const s = String(value ?? '').trim();
  if (!s || s === '-') return '';
  const n = parseMoney(s);
  if (!Number.isFinite(n)) {
    const core = s.startsWith('$') ? s.slice(1).trim() : s;
    return core ? (s.startsWith('$') ? s : `$${core}`) : '';
  }
  return `$${formatNumberWithThousandSeparator(n)}`;
}

function formatHoursOpenLabel(hoursOpen: unknown): string {
  const raw = String(hoursOpen ?? '').trim();
  if (!raw) return '';
  if (/hour|minute/i.test(raw)) return raw;
  return raw;
}

function formatSeenPercentDisplay(value: unknown): string {
  const t = String(value ?? '').trim();
  if (!t || t === '-') return '';
  if (t.endsWith('%')) return t;
  return `${t}%`;
}

function displayTotal(value: unknown): string {
  const s = String(value ?? '').trim();
  return s === '' ? '' : s;
}

function getMonthKey(dateValue: string | undefined): string {
  const raw = String(dateValue ?? '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw.slice(0, 7);
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
}

function formatYearMonthOptionLabel(yearMonth: string): string {
  if (!/^\d{4}-\d{2}$/.test(yearMonth)) return yearMonth;
  const [year, month] = yearMonth.split('-').map(Number);
  if (!Number.isFinite(year) || !Number.isFinite(month) || month < 1 || month > 12) return yearMonth;
  return new Date(year, month - 1, 1).toLocaleDateString('en-US', { year: 'numeric', month: 'long' });
}

function parseShiftOffices(shift: unknown): string[] {
  if (Array.isArray(shift)) {
    return shift.map((item) => String(item ?? '').trim()).filter(Boolean);
  }
  if (typeof shift === 'string') {
    const raw = shift.trim();
    if (!raw) return [];
    if (raw.startsWith('[') && raw.endsWith(']')) {
      try {
        return parseShiftOffices(JSON.parse(raw));
      } catch {
        // fall through to delimiter split
      }
    }
    return raw
      .split(/[,;\n]/)
      .map((item) => item.trim())
      .filter(Boolean);
  }
  return [];
}

export default function ViewPage() {
  const [pageReady, setPageReady] = useState(false);
  const [allowedOffices, setAllowedOffices] = useState<string[]>([]);
  const [docs, setDocs] = useState<FormDoc[]>([]);
  const [selectedId, setSelectedId] = useState<string>('');
  const [selectedYearMonth, setSelectedYearMonth] = useState('');
  const [selectedOffice, setSelectedOffice] = useState('');
  const [dateFilter, setDateFilter] = useState<string[]>([]);
  const [loading, setLoading] = useState(true);
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
        if (userData?.role !== 'Manager' && userData?.role !== 'HR' && userData?.role !== 'Director') {
          goHome();
          return;
        }

        if (!cancelled) {
          setAllowedOffices(parseShiftOffices(userData?.offices));
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

    const load = async () => {
      try {
        setLoading(true);
        const snap = await getDocs(collection(db, 'simple-forms'));
        const loaded = snap.docs
          .map((d) => {
            const data = d.data() as Omit<FormDoc, 'id'>;
            const { notesLines: _nl, notDueLines: _ndl, ...rest } = data;
            return {
              id: d.id,
              ...rest,
              notes: loadMultilineField(data.notes, data.notesLines).slice(0, NOTES_MAX_LENGTH),
              notDue: loadMultilineField(data.notDue, data.notDueLines).slice(0, NOTES_MAX_LENGTH),
            };
          })
          .filter((docItem) => String(docItem.submittedDateTime ?? '').trim() !== '')
          .sort((a, b) => `${b.date ?? ''}_${b.location ?? ''}`.localeCompare(`${a.date ?? ''}_${a.location ?? ''}`));
        setDocs(loaded);
      } catch {
        setError('Unable to access this page.');
      } finally {
        setLoading(false);
      }
    };

    load();
  }, [pageReady]);

  const filtersReady = selectedYearMonth !== '' && selectedOffice !== '';

  const yearMonthOptions = useMemo(() => {
    const keys = new Set<string>();
    for (const docItem of docs) {
      const key = getMonthKey(docItem.date);
      if (/^\d{4}-\d{2}$/.test(key)) keys.add(key);
    }
    return Array.from(keys).sort((a, b) => b.localeCompare(a));
  }, [docs]);

  const officeOptions = useMemo(() => {
    const allowed = new Set(allowedOffices.map((office) => office.trim()).filter(Boolean));
    if (allowed.size === 0) return [];

    const offices = new Set<string>();
    for (const docItem of docs) {
      const office = String(docItem.location ?? '').trim();
      if (office && allowed.has(office)) offices.add(office);
    }
    return Array.from(offices).sort((a, b) => a.localeCompare(b));
  }, [docs, allowedOffices]);

  useEffect(() => {
    if (selectedOffice && !officeOptions.includes(selectedOffice)) {
      setSelectedOffice('');
      setDateFilter([]);
    }
  }, [officeOptions, selectedOffice]);

  const dateOptions = useMemo(() => {
    if (!filtersReady) return [];
    const dates = new Set<string>();
    for (const docItem of docs) {
      const dateValue = String(docItem.date ?? '').trim();
      if (!dateValue) continue;
      if (getMonthKey(dateValue) !== selectedYearMonth) continue;
      if (String(docItem.location ?? '').trim() !== selectedOffice) continue;
      dates.add(dateValue);
    }
    return Array.from(dates).sort((a, b) => b.localeCompare(a));
  }, [docs, filtersReady, selectedYearMonth, selectedOffice]);

  const filteredDocs = useMemo(() => {
    if (!filtersReady) return [];
    return docs.filter((docItem) => {
      if (String(docItem.submittedDateTime ?? '').trim() === '') return false;
      const dateValue = String(docItem.date ?? '').trim();
      const locationValue = String(docItem.location ?? '').trim();
      if (getMonthKey(dateValue) !== selectedYearMonth) return false;
      if (locationValue !== selectedOffice) return false;
      if (dateFilter.length > 0 && !dateFilter.includes(dateValue)) return false;
      return true;
    });
  }, [docs, filtersReady, selectedYearMonth, selectedOffice, dateFilter]);

  const dateFilterLabel = dateFilter.length === 0 ? 'Date (All in month)' : `Date (${dateFilter.length})`;

  const toggleDateFilterValue = (value: string) => {
    setDateFilter((prev) => (prev.includes(value) ? prev.filter((item) => item !== value) : [...prev, value]));
  };

  const handleYearMonthChange = (value: string) => {
    setSelectedYearMonth(value);
    setDateFilter([]);
  };

  const handleOfficeChange = (value: string) => {
    setSelectedOffice(value);
    setDateFilter([]);
  };

  useEffect(() => {
    if (filteredDocs.length === 0) {
      setSelectedId('');
      return;
    }
    const existsInFiltered = filteredDocs.some((docItem) => docItem.id === selectedId);
    if (!existsInFiltered) {
      setSelectedId(filteredDocs[0].id);
    }
  }, [filteredDocs, selectedId]);

  const selectedDoc = useMemo(() => filteredDocs.find((d) => d.id === selectedId), [filteredDocs, selectedId]);

  const tableRows = selectedDoc?.tableRows || [];
  const sugarRows = selectedDoc?.sugarRows || [];
  const locationSummary = selectedDoc?.locationSummary || {};
  const sideMetrics = selectedDoc?.productionSideMetrics || {};
  const coffeeActualTotals = selectedDoc?.coffeeActualTotals || {};
  const tableTotals = selectedDoc?.tableTotals || {};
  const sugarTotals = selectedDoc?.sugarTotals || {};

  const doctorRows = tableRows.map((tableRow, idx) => {
    const extra = selectedDoc?.extraInputRows?.[idx] || {};
    return {
      position: extra.position || tableRow.position || '',
      name: extra.name || tableRow.name || '',
      production: tableRow.sales || '',
      doctorPreventative: extra.doctorPreventative || '',
      doctorRestorative: extra.doctorRestorative || '',
      doctorCraProduction: extra.doctorCraProduction || '',
      customer: extra.customer || '',
      icecream: extra.icecream || '',
      cake: extra.cake || '',
      donut: extra.donut || '',
      tart: extra.tart || '',
      peach: extra.peach || '',
      peppermint: extra.peppermint || '',
    };
  });

  return (
    <main className="d-page-main" style={{ minHeight: '100vh', background: '#fff', padding: 24 }}>
      <section
        style={{
          maxWidth: 2600,
          margin: '0 auto',
          padding: 20,
        }}
      >
        <div style={{ marginBottom: 14, display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8 }}>
          <h2 style={{ margin: '0 0 8px', fontSize: 28, fontWeight: 800, color: '#111827' }}>Daily Production View</h2>
        </div>

        {loading && <p style={{ margin: 0, color: '#6b7280' }}>Loading...</p>}
        {error && <p style={{ margin: 0, color: '#b91c1c' }}>{error}</p>}

        {!loading && !error && (
          <div style={{ display: 'grid', gridTemplateColumns: '380px 1fr', gap: 16 }}>
            <div style={{ border: '1px solid #e5e7eb', borderRadius: 10, overflow: 'hidden' }}>
              <div style={{ padding: 10, background: '#f9fafb', borderBottom: '1px solid #e5e7eb', fontWeight: 700 }}>
                Date / Office
              </div>
              <div style={{ padding: 10, borderBottom: '1px solid #e5e7eb', display: 'grid', gap: 8 }}>
                <label style={{ display: 'grid', gap: 4 }}>
                  <span style={{ fontSize: 12, fontWeight: 600, color: '#475569' }}>Year / Month</span>
                  <select
                    value={selectedYearMonth}
                    onChange={(e) => handleYearMonthChange(e.target.value)}
                    style={FILTER_SELECT_STYLE}
                  >
                    <option value="">Select year / month</option>
                    {yearMonthOptions.map((ym) => (
                      <option key={ym} value={ym}>
                        {formatYearMonthOptionLabel(ym)}
                      </option>
                    ))}
                  </select>
                </label>
                <label style={{ display: 'grid', gap: 4 }}>
                  <span style={{ fontSize: 12, fontWeight: 600, color: '#475569' }}>Office</span>
                  <select value={selectedOffice} onChange={(e) => handleOfficeChange(e.target.value)} style={FILTER_SELECT_STYLE}>
                    <option value="">Select office</option>
                    {officeOptions.map((office) => (
                      <option key={office} value={office}>
                        {office}
                      </option>
                    ))}
                  </select>
                </label>
                {filtersReady && (
                  <details style={{ width: '100%', border: '1px solid #d1d5db', borderRadius: 6, background: '#fff' }}>
                    <summary style={{ cursor: 'pointer', listStyle: 'none', padding: '8px 10px', fontWeight: 600 }}>{dateFilterLabel}</summary>
                    <div style={{ borderTop: '1px solid #e5e7eb', padding: '8px 10px', display: 'grid', gap: 6, maxHeight: 180, overflowY: 'auto' }}>
                      <label style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer' }}>
                        <input type="checkbox" checked={dateFilter.length === 0} onChange={() => setDateFilter([])} />
                        All dates in month
                      </label>
                      {dateOptions.map((date) => (
                        <label key={date} style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer' }}>
                          <input type="checkbox" checked={dateFilter.includes(date)} onChange={() => toggleDateFilterValue(date)} />
                          {date}
                        </label>
                      ))}
                    </div>
                  </details>
                )}
              </div>
              {filtersReady && filteredDocs.length === 0 && (
                <div style={{ padding: 12, color: '#6b7280' }}>There is no data yet.</div>
              )}
              {filtersReady &&
                filteredDocs.map((d, docIdx) => {
                  const active = d.id === selectedId;
                  const isLastDocRow = docIdx === filteredDocs.length - 1;
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
                        borderBottom: isLastDocRow ? 'none' : '1px solid #e5e7eb',
                        background: active ? '#eef2ff' : '#fff',
                        cursor: 'pointer',
                      }}
                    >
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
                        <div style={{ fontWeight: 700 }}>{d.date || ''}</div>
                      </div>
                      <div style={{ color: '#475569', fontSize: 13 }}>{d.location || ''}</div>
                    </button>
                  );
                })}
            </div>

            <div
              style={{
                border: '1px solid #e5e7eb',
                borderRadius: 10,
                padding: D_PAGE_DETAIL_CARD_PADDING_PX,
                boxSizing: 'border-box',
              }}
            >
              {filtersReady && selectedDoc && (
                <>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, minmax(220px, 1fr))', gap: 10, marginBottom: 14 }}>
                    <div>
                      <strong>Date:</strong> {selectedDoc.date || ''}
                    </div>
                    <div>
                      <strong>Office:</strong> {selectedDoc.location || ''}
                    </div>
                    <div>
                      <strong>Reason if Late:</strong> {selectedDoc.reasonIfLate || ''}
                    </div>
                    <div>
                      <strong>Submitted at:</strong> {selectedDoc.submittedDateTime || ''}
                    </div>
                  </div>

                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, minmax(200px, 1fr))', gap: 10, marginBottom: 14 }}>
                    <div>
                      <strong>Check In:</strong> {selectedDoc.checkIn || ''}
                    </div>
                    <div>
                      <strong>Check Out:</strong> {selectedDoc.checkOut || ''}
                    </div>
                    <div>
                      <strong>Hours Open:</strong>{' '}
                      <span style={{ marginLeft: 6, color: '#374151' }}>{formatHoursOpenLabel(selectedDoc.hoursOpen)}</span>
                    </div>
                    <div>
                      <strong>Closer:</strong> {selectedDoc.closer || ''}
                    </div>
                  </div>

                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, minmax(220px, 1fr))', gap: 10, marginTop: 14 }}>
                    <div>
                      <strong>Submitted Production:</strong> {formatCurrencyLabel(selectedDoc.grandTotal || '')}
                    </div>
                    <div>
                      <strong>CRA Production:</strong> {formatCurrencyLabel(selectedDoc.coffeeSales || '')}
                    </div>
                    <div>
                      <strong>Production W/Out CRA:</strong> {formatCurrencyLabel(selectedDoc.salesWithoutCoffee || '')}
                    </div>
                  </div>

                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, minmax(200px, 1fr))', gap: 10, marginTop: 10 }}>
                    <div>
                      <strong>Prophy @ OE:</strong> {selectedDoc.paperAtOrangeJuice || ''}
                    </div>
                    <div>
                      <strong>Prophy @ TX:</strong> {selectedDoc.paperAtTea || ''}
                    </div>
                    <div>
                      <strong>Just Prophy:</strong> {selectedDoc.justPaper || ''}
                    </div>
                    <div>
                      <strong>Actual Prophy:</strong>{' '}
                      <span style={{ marginLeft: 6, color: '#374151' }}>{selectedDoc.prophyTotal || ''}</span>
                    </div>
                  </div>

                  <h3 style={{ margin: '16px 0 10px' }}>Production</h3>
                  <div
                    style={{
                      ...dPageDetailTableBleedScroll,
                      display: 'flex',
                      flexWrap: 'wrap',
                      gap: 16,
                      alignItems: 'flex-start',
                    }}
                  >
                    <div>
                      <table style={{ width: '100%', maxWidth: 480, borderCollapse: 'collapse', fontSize: 14 }}>
                        <tbody>
                          <tr>
                            <td style={labelCellStyle}>Preventative</td>
                            <td style={tdStyle}>{formatCurrencyLabel(locationSummary.pineapple || '')}</td>
                          </tr>
                          <tr>
                            <td style={labelCellStyle}>Restorative</td>
                            <td style={tdStyle}>{formatCurrencyLabel(locationSummary.rose || '')}</td>
                          </tr>
                          <tr>
                            <td style={labelCellStyle}>CRA Production</td>
                            <td style={tdStyle}>{formatCurrencyLabel(selectedDoc.coffeeSales || '')}</td>
                          </tr>
                          <tr>
                            <td style={labelCellStyle}>1st Review Production</td>
                            <td style={tdStyle}>{formatCurrencyLabel(locationSummary.total || '')}</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                    <div>
                      <table style={{ width: '100%', minWidth: 220, maxWidth: 320, borderCollapse: 'collapse', fontSize: 14 }}>
                        <tbody>
                          {PRODUCTION_SIDE_METRIC_ROWS.map(({ key, label }) => (
                            <tr key={key}>
                              <td style={{ ...labelCellStyle, whiteSpace: 'nowrap' }}>{label}</td>
                              <td style={tdStyle}>
                                {key === 'seenPercent'
                                  ? formatSeenPercentDisplay(sideMetrics.seenPercent)
                                  : String(sideMetrics[key] ?? '').trim() || ''}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  <h3 style={{ margin: '16px 0 10px' }}>Doctors</h3>
                  <div style={dPageDetailTableBleedScroll}>
                    <table style={{ width: '100%', minWidth: 1400, borderCollapse: 'collapse', fontSize: 14 }}>
                      <thead>
                        <tr style={{ background: '#f3f4f6' }}>
                          {DOCTOR_HEADERS.map((h) => (
                            <th key={h} style={thStyle}>
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {doctorRows.map((row, idx) => (
                          <tr key={`doctor-${idx}`}>
                            <td style={tdStyle}>{row.position}</td>
                            <td style={tdStyle}>{row.name}</td>
                            <td style={tdStyle}>{formatCurrencyLabel(row.production)}</td>
                            <td style={tdStyle}>{formatCurrencyLabel(row.doctorPreventative)}</td>
                            <td style={tdStyle}>{formatCurrencyLabel(row.doctorRestorative)}</td>
                            <td style={tdStyle}>{formatCurrencyLabel(row.doctorCraProduction)}</td>
                            <td style={tdStyle}>{row.customer}</td>
                            <td style={tdStyle}>{row.icecream}</td>
                            <td style={tdStyle}>{row.cake}</td>
                            <td style={tdStyle}>{row.donut}</td>
                            <td style={tdStyle}>{row.tart}</td>
                            <td style={tdStyle}>{row.peach}</td>
                            <td style={tdStyle}>{row.peppermint}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  <h3 style={{ margin: '16px 0 10px' }}>CRA / OE</h3>
                  <div style={dPageDetailTableBleedScroll}>
                    <table style={{ width: '100%', minWidth: 1500, borderCollapse: 'collapse', fontSize: 14 }}>
                      <thead>
                        <tr style={{ background: '#f3f4f6' }}>
                          {TABLE_HEADERS.map((h) => (
                            <th key={h} style={thStyle}>
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {tableRows.map((r, idx) => (
                          <tr key={`table-${idx}`}>
                            <td style={tdStyle}>{r.position || ''}</td>
                            <td style={tdStyle}>{r.name || ''}</td>
                            <td style={tdStyle}>{r.coffeeNew || ''}</td>
                            <td style={tdStyle}>{r.coffeeReturn || ''}</td>
                            <td style={tdStyle}>{r.coffeeTotal || ''}</td>
                            <td style={tdStyle}>{r.coffeeNo || ''}</td>
                            <td style={tdStyle}>{r.renderedCoffee || ''}</td>
                            <td style={tdStyle}>{r.coffeeYes || ''}</td>
                            <td style={tdStyle}>{r.orangeJuiceNew || ''}</td>
                            <td style={tdStyle}>{r.orangeJuiceReturn || ''}</td>
                            <td style={tdStyle}>{r.orangeJuiceTotal || ''}</td>
                            <td style={tdStyle} />
                            <td style={tdStyle} />
                          </tr>
                        ))}
                      </tbody>
                      <tfoot>
                        <tr style={{ background: '#f9fafb', fontWeight: 700 }}>
                          <td style={tdStyle}>Total</td>
                          <td style={tdStyle} />
                          <td style={tdStyle}>{displayTotal(tableTotals.coffeeNew)}</td>
                          <td style={tdStyle}>{displayTotal(tableTotals.coffeeReturn)}</td>
                          <td style={tdStyle}>{displayTotal(tableTotals.coffeeTotal)}</td>
                          <td style={tdStyle}>{displayTotal(tableTotals.coffeeNo)}</td>
                          <td style={tdStyle}>{displayTotal(tableTotals.renderedCoffee)}</td>
                          <td style={tdStyle}>{displayTotal(tableTotals.coffeeYes)}</td>
                          <td style={tdStyle}>{displayTotal(tableTotals.orangeJuiceNew)}</td>
                          <td style={tdStyle}>{displayTotal(tableTotals.orangeJuiceReturn)}</td>
                          <td style={tdStyle}>{displayTotal(tableTotals.orangeJuiceTotal)}</td>
                          <td style={tdStyle}>{coffeeActualTotals.orangeJuiceNew || ''}</td>
                          <td style={tdStyle}>{coffeeActualTotals.orangeJuiceReturn || ''}</td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>

                  <h3 style={{ margin: '16px 0 10px' }}>Sealant / Prophy</h3>
                  <div style={dPageDetailTableBleedScroll}>
                    <table style={{ width: '100%', minWidth: 760, borderCollapse: 'collapse', fontSize: 14 }}>
                      <thead>
                        <tr style={{ background: '#f3f4f6' }}>
                          {SUGAR_HEADERS.map((h) => (
                            <th key={h} style={thStyle}>
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {sugarRows.map((r, idx) => (
                          <tr key={`sugar-${idx}`}>
                            <td style={tdStyle}>{r.position || ''}</td>
                            <td style={tdStyle}>{r.name || ''}</td>
                            <td style={tdStyle}>{r.sugar || ''}</td>
                            <td style={tdStyle}>{r.sugarGood || ''}</td>
                            <td style={tdStyle}>{r.sugarBad || ''}</td>
                            <td style={tdStyle}>{r.paper || ''}</td>
                          </tr>
                        ))}
                      </tbody>
                      <tfoot>
                        <tr style={{ background: '#f9fafb', fontWeight: 700 }}>
                          <td style={tdStyle}>Total</td>
                          <td style={tdStyle} />
                          <td style={tdStyle}>{displayTotal(sugarTotals.sugar)}</td>
                          <td style={tdStyle}>{displayTotal(sugarTotals.sugarGood)}</td>
                          <td style={tdStyle}>{displayTotal(sugarTotals.sugarBad)}</td>
                          <td style={tdStyle}>{displayTotal(sugarTotals.paper)}</td>
                        </tr>
                      </tfoot>
                    </table>
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
