'use client';

import React, { Suspense, useEffect, useMemo, useState } from 'react';
import { useSearchParams } from 'next/navigation';
import { db } from '@/lib/firebase.config';
import { collection, deleteField, doc, getDoc, getDocs, serverTimestamp, setDoc } from 'firebase/firestore';

const PINK_BEAR_COLLECTION = 'monthly report';

type ReviewSummarySnapshot = {
  dailyGoal: number;
  monthlyGoal: number | null;
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

type IssueSummarySnapshot = {
  monthly: string;
  yearly: string;
  freeDays: string;
};

type TableRow = {
  position?: string;
  name?: string;
  coffeeNew?: string;
  coffeeReturn?: string;
  coffeeYes?: string;
};

type CoffeeActualTotals = {
  orangeJuiceNew?: string;
  orangeJuiceReturn?: string;
  orangeJuiceTotal?: string;
};

type SugarRow = {
  position?: string;
  name?: string;
  sugarGood?: string;
};

type ExtraInputRow = {
  customer?: unknown;
};

type FormDoc = {
  id: string;
  date?: string;
  location?: string;
  submittedDateTime?: string;
  coffeeSales?: string;
  prophyTotal?: string;
  tableRows?: TableRow[];
  coffeeActualTotals?: CoffeeActualTotals;
  sugarRows?: SugarRow[];
  extraInputRows?: ExtraInputRow[];
  checkIn?: string;
  checkOut?: string;
  closer?: string;
  locationSummary?: {
    mailedProduction?: string;
    pineapple?: string;
    rose?: string;
    total?: unknown;
  };
  productionSideMetrics?: {
    add?: unknown;
    noShow?: unknown;
    scheduled?: unknown;
    seen?: unknown;
    postcard?: unknown;
    referral?: unknown;
  };
};

type Aggregate = {
  docCount: number;
  coffeeSales: number;
  prophyTotal: number;
  mailedProduction: number;
  pineapple: number;
  rose: number;
  visitsAdd: number;
  visitsNoShow: number;
  visitsScheduled: number;
  visitsSeen: number;
  visitsPostcard: number;
  visitsReferral: number;
  doctorCustomerTotal: number;
  orangeJuiceNew: number;
  orangeJuiceReturn: number;
  orangeJuiceTotal: number;
  coffeeRows: Array<{ position: string; name: string; coffeeNew: number; coffeeReturn: number; coffeeYes: number }>;
  sugarRows: Array<{ position: string; name: string; sugarGood: number }>;
};

const OFFICE_SUMMARY_GOAL_KEYS = ['preventative', 'restorative', 'craProduction', 'firstReviewProduction', 'mailedProduction'] as const;
const OFFICE_SUMMARY_TOTAL_COMPONENT_KEYS = ['preventative', 'restorative', 'craProduction'] as const;
const OFFICE_SUMMARY_SAVABLE_GOAL_KEYS = ['preventative', 'restorative', 'craProduction', 'firstReviewProduction'] as const;
const OFFICE_SUMMARY_NO_GOAL_KEYS = ['mailedProduction'] as const;
const OE_CORE_GOAL_KEYS = ['oeNp', 'oeRc', 'oeTotal'] as const;
const OE_PROPHY_GOAL_KEYS = ['actualProphy'] as const;
const SEALANT_GOAL_KEYS = ['sealantRda', 'sealantDds'] as const;
const OE_TABLE_GOAL_KEYS = [...OE_CORE_GOAL_KEYS, ...OE_PROPHY_GOAL_KEYS, ...SEALANT_GOAL_KEYS] as const;

type OfficeSummaryGoalKey = (typeof OFFICE_SUMMARY_GOAL_KEYS)[number];
type OeTableGoalKey = (typeof OE_TABLE_GOAL_KEYS)[number];
type OeCoreGoalKey = (typeof OE_CORE_GOAL_KEYS)[number];
type OeProphyGoalKey = (typeof OE_PROPHY_GOAL_KEYS)[number];
type SealantGoalKey = (typeof SEALANT_GOAL_KEYS)[number];
type OfficeSummaryGoals = Partial<Record<OfficeSummaryGoalKey, number>>;
type OeTableGoals = Partial<Record<OeTableGoalKey, number>>;
type OfficeSummaryGoalOverrides = { monthlyGoal: OfficeSummaryGoals; dailyGoal: OfficeSummaryGoals };
type OeTableGoalOverrides = { monthlyGoal: OeTableGoals; dailyGoal: OeTableGoals };

const VISITS_HEADERS = ['Add On', 'No Shows', 'Scheduled', 'Seen %', 'Postcard Count'] as const;
const SEEN_HEADERS = ['Seen (Office)', 'Seen (Doctor)'] as const;
const DAILY_OFFICE_LOG_HEADERS = ['Date', 'Daily Production', 'Check In', 'Check Out', 'Closer', 'Hours Open'] as const;
const CRA_SUMMARY_HEADERS = ['CRA (New)', 'CRA (Return)', 'CRA (Billable)'] as const;
const CRA_NON_CRA_HEADERS = ['>6YRS (Non CRA)', '<6YRS (CRA)'] as const;
const REFERRAL_STARTED_HEADERS = ['Referred', 'Started'] as const;
const CLEAN_HEADERS = ['Debris', 'Odor', 'Unclean Lobby'] as const;

const BONUS_DEDUCT_COLUMN_HEADERS = ['#', 'Bonus', 'Deduct'] as const;
const BONUS_DEDUCT_ROW_LABELS = [
  '5* Reviews',
  '1* Reviews',
  'Exams^Daily Goal',
  'Sealants',
  'Cleanliness',
  'Injury',
  'Ortho Starts',
  'RDA Sealants',
] as const;

const BONUS_DEDUCT_BONUS_MULTIPLIERS: Partial<Record<(typeof BONUS_DEDUCT_ROW_LABELS)[number], number>> = {
  '5* Reviews': 50,
  'Exams^Daily Goal': 5,
  Sealants: 5,
  'Ortho Starts': 50,
  'RDA Sealants': 5,
};

const BONUS_DEDUCT_DEDUCT_MULTIPLIERS: Partial<Record<(typeof BONUS_DEDUCT_ROW_LABELS)[number], number>> = {
  '1* Reviews': -50,
  Cleanliness: -50,
  Injury: -500,
};

const PROMOTION_COLUMN_HEADERS = ['Nintendo', 'Refer a Friend'] as const;
const PROMOTION_ROW_HEADERS = ['Daily Goal', 'Monthly Goal', 'Actual', 'Daily Average'] as const;

const PROMOTION_DAILY_GOALS: Record<(typeof PROMOTION_COLUMN_HEADERS)[number], number> = {
  Nintendo: 2,
  'Refer a Friend': 1,
};

const PROMOTION_ACTUAL_KEYS = {
  Nintendo: 'nintendo',
  'Refer a Friend': 'referAFriend',
} as const;

type PromotionActualKey = (typeof PROMOTION_ACTUAL_KEYS)[keyof typeof PROMOTION_ACTUAL_KEYS];
type PromotionActuals = Partial<Record<PromotionActualKey, number>>;

const WAIT_TIME_ROWS = [
  { key: '0-15', label: '0-15 min' },
  { key: '16-30', label: '16-30 min' },
  { key: '31-45', label: '31-45 min' },
  { key: '46-60', label: '46-60 min' },
  { key: '1plus', label: '1+ hr' },
  { key: 'notCompleted', label: 'Not Completed' },
] as const;

type WaitTimeKey = (typeof WAIT_TIME_ROWS)[number]['key'];
type WaitTimes = Partial<Record<WaitTimeKey, number>>;

const EXPENSE_ROWS = [
  { key: 'labor', label: 'Labor' },
  { key: 'officeSupply', label: 'Office Supply' },
  { key: 'dentalSupply', label: 'Dental Supply' },
  { key: 'water', label: 'Water' },
  { key: 'electric', label: 'Electric' },
  { key: 'gas', label: 'Gas' },
  { key: 'trash', label: 'Trash' },
  { key: 'blank1', label: '' },
  { key: 'blank2', label: '' },
  { key: 'blank3', label: '' },
  { key: 'total', label: 'Total' },
] as const;

const EXPENSE_CUSTOM_ROW_KEYS = ['blank1', 'blank2', 'blank3'] as const;

type ExpenseKey = (typeof EXPENSE_ROWS)[number]['key'];
type ExpenseCustomRowKey = (typeof EXPENSE_CUSTOM_ROW_KEYS)[number];
type ExpenseLabelKey = `${ExpenseCustomRowKey}Label`;
type Expenses = Partial<Record<ExpenseKey, number>> & Partial<Record<ExpenseLabelKey, string>>;

function isExpenseCustomRowKey(key: ExpenseKey): key is ExpenseCustomRowKey {
  return (EXPENSE_CUSTOM_ROW_KEYS as readonly string[]).includes(key);
}

function getExpenseLabelKey(key: ExpenseCustomRowKey): ExpenseLabelKey {
  return `${key}Label`;
}

const cellStyle: React.CSSProperties = {
  borderBottom: '1px solid #e5e7eb',
  borderRight: '1px solid #e5e7eb',
  padding: '8px 10px',
  fontSize: 13,
  color: '#111827',
  overflowWrap: 'anywhere',
  wordBreak: 'break-word',
  background: '#fff',
};
const headerCellStyle: React.CSSProperties = {
  ...cellStyle,
  textAlign: 'left',
  background: '#f8fafc',
  fontWeight: 700,
  color: '#334155',
};
const labelCellStyle: React.CSSProperties = {
  ...cellStyle,
  fontWeight: 600,
  background: '#fafafa',
  color: '#475569',
};
const inputStyle: React.CSSProperties = {
  width: '100%',
  height: 30,
  border: '1px solid #d1d5db',
  borderRadius: 6,
  padding: '0 8px',
  fontSize: 13,
  boxSizing: 'border-box',
  background: '#fff',
};
const rowStyle = { fontWeight: 600 as const };
const centeredCellStyle: React.CSSProperties = { ...cellStyle, textAlign: 'center', fontWeight: 600 };
const editButtonStyle: React.CSSProperties = {
  height: 36,
  padding: '0 16px',
  borderRadius: 8,
  fontWeight: 600,
  fontSize: 14,
  cursor: 'pointer',
  border: 'none',
};
const pageSectionGap = 24;
const tableGridStyle: React.CSSProperties = {
  display: 'grid',
  gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))',
  gap: 12,
  alignItems: 'start',
};
const tableCardStyle: React.CSSProperties = {
  minWidth: 0,
  maxWidth: '100%',
};
const summaryTablesRowStyle: React.CSSProperties = {
  display: 'grid',
  gridTemplateColumns: 'minmax(0, 1.35fr) minmax(0, 1fr) minmax(0, 1fr) minmax(0, 1fr)',
  gap: 12,
  alignItems: 'start',
};

function TableCard({ title, children }: { title?: string; children: React.ReactNode }) {
  return (
    <div
      style={{
        background: '#fff',
        border: '1px solid #e5e7eb',
        borderRadius: 10,
        overflow: 'hidden',
        boxShadow: '0 1px 3px rgba(15, 23, 42, 0.06)',
      }}
    >
      {title && (
        <div
          style={{
            padding: '10px 14px',
            borderBottom: '1px solid #e5e7eb',
            background: '#f8fafc',
            fontWeight: 700,
            fontSize: 13,
            color: '#334155',
          }}
        >
          {title}
        </div>
      )}
      <div style={{ overflowX: 'auto' }}>{children}</div>
    </div>
  );
}

function SectionBlock({ children }: { children: React.ReactNode }) {
  return (
    <section style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
      {children}
    </section>
  );
}

function StatusMessage({ type, children }: { type: 'info' | 'error' | 'success'; children: React.ReactNode }) {
  const styles: Record<'info' | 'error' | 'success', React.CSSProperties> = {
    info: { background: '#f8fafc', color: '#64748b', border: '1px solid #e2e8f0' },
    error: { background: '#fef2f2', color: '#b91c1c', border: '1px solid #fecaca' },
    success: { background: '#f0fdf4', color: '#166534', border: '1px solid #bbf7d0' },
  };
  return (
    <p
      style={{
        margin: '0 0 16px',
        padding: '10px 14px',
        borderRadius: 8,
        fontSize: 14,
        ...styles[type],
      }}
    >
      {children}
    </p>
  );
}

function getTableStyle(): React.CSSProperties {
  return {
    width: '100%',
    maxWidth: '100%',
    borderCollapse: 'collapse',
    fontSize: 13,
    tableLayout: 'fixed',
  };
}

function parseNumber(value: unknown): number {
  const n = Number(value);
  if (!Number.isFinite(n)) return 0;
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

function addAmount(base: number, value: unknown): number {
  return Math.round((base + parseNumber(value) + Number.EPSILON) * 100) / 100;
}

function formatAmount(value: number): string {
  const rounded = Math.round((value + Number.EPSILON) * 100) / 100;
  return Number.isInteger(rounded) ? String(rounded) : rounded.toFixed(2);
}

function formatDollarAmount(value: number): string {
  const amount = formatAmount(value);
  return amount.startsWith('-') ? `-$${amount.slice(1)}` : `$${amount}`;
}

function formatDollarDailyAverage(total: number, days: number): string {
  if (days <= 0) return '$0';
  return formatDollarAmount(Math.round(((total / days) + Number.EPSILON) * 100) / 100);
}

function formatDailyAverage(total: number, days: number): string {
  if (days <= 0) return '0';
  const avg = Math.round(((total / days) + Number.EPSILON) * 100) / 100;
  return Number.isInteger(avg) ? String(avg) : avg.toFixed(2);
}

function formatInteger(value: number): string {
  return String(Math.round(value));
}

function formatIntegerDailyAverage(total: number, days: number): string {
  if (days <= 0) return '0';
  return String(Math.round(total / days));
}

function formatIntegerGoal(value: number | undefined): string {
  return value === undefined ? '' : formatInteger(value);
}

function formatPercentage(value: number, total: number): string {
  if (total <= 0) return '0%';
  return `${Math.round((value / total) * 100)}%`;
}

function getDocFieldString(value: unknown): string {
  if (value === undefined || value === null) return '';
  return String(value).trim();
}

function parseTimeToMinutes(value: unknown): number | null {
  if (value === undefined || value === null || value === '') return null;

  if (typeof value === 'object' && value !== null) {
    if ('toDate' in value && typeof (value as { toDate: () => Date }).toDate === 'function') {
      const d = (value as { toDate: () => Date }).toDate();
      return d.getHours() * 60 + d.getMinutes();
    }
    if ('seconds' in value && typeof (value as { seconds: number }).seconds === 'number') {
      const d = new Date((value as { seconds: number }).seconds * 1000);
      return d.getHours() * 60 + d.getMinutes();
    }
  }

  const raw = String(value).trim();
  const ampmMatch = raw.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)$/i);
  if (ampmMatch) {
    let hours = Number(ampmMatch[1]);
    const minutes = Number(ampmMatch[2]);
    const isPm = ampmMatch[4].toUpperCase() === 'PM';
    if (isPm && hours !== 12) hours += 12;
    if (!isPm && hours === 12) hours = 0;
    return hours * 60 + minutes;
  }

  const hhmmMatch = raw.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (hhmmMatch) {
    return Number(hhmmMatch[1]) * 60 + Number(hhmmMatch[2]);
  }

  return null;
}

function formatTimeDisplay(value: unknown): string {
  if (value === undefined || value === null || value === '') return '';

  if (typeof value === 'object' && value !== null) {
    if ('toDate' in value && typeof (value as { toDate: () => Date }).toDate === 'function') {
      const d = (value as { toDate: () => Date }).toDate();
      return d.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
    }
    if ('seconds' in value && typeof (value as { seconds: number }).seconds === 'number') {
      const d = new Date((value as { seconds: number }).seconds * 1000);
      return d.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
    }
  }

  return String(value).trim();
}

function formatHoursOpen(checkIn: unknown, checkOut: unknown): string {
  const start = parseTimeToMinutes(checkIn);
  const end = parseTimeToMinutes(checkOut);
  if (start === null || end === null) return '';

  let diffMinutes = end - start;
  if (diffMinutes <= 0) diffMinutes += 24 * 60;

  const hours = Math.floor(diffMinutes / 60);
  const minutes = diffMinutes % 60;
  return `${hours} hrs ${minutes} min`;
}

type DailyOfficeLogRow = {
  id: string;
  date: string;
  dailyProduction?: number;
  checkIn?: string;
  checkOut?: string;
  closer: string;
};

function buildDailyOfficeLogRows(monthlyDocs: FormDoc[]): DailyOfficeLogRow[] {
  return monthlyDocs
    .slice()
    .sort((a, b) => String(a.date ?? '').localeCompare(String(b.date ?? '')))
    .map((doc) => {
      const totalRaw = doc.locationSummary?.total;
      const dailyProduction = totalRaw === undefined || totalRaw === null || totalRaw === ''
        ? undefined
        : parseNumber(totalRaw);

      return {
        id: doc.id,
        date: getDocFieldString(doc.date),
        dailyProduction,
        checkIn: doc.checkIn,
        checkOut: doc.checkOut,
        closer: getDocFieldString(doc.closer),
      };
    });
}

function sumRowsAmount<T>(rows: readonly T[], pick: (row: T) => number): number {
  return rows.reduce((sum, row) => Math.round((sum + pick(row) + Number.EPSILON) * 100) / 100, 0);
}

function sumCoffeeRows(rows: Aggregate['coffeeRows'], pick: (row: Aggregate['coffeeRows'][number]) => number) {
  return sumRowsAmount(rows, pick);
}

function sumSugarRowsByPosition(rows: Aggregate['sugarRows'], positions: readonly string[]) {
  const targets = new Set(positions.map((p) => p.trim().toLowerCase()));
  return sumRowsAmount(
    rows.filter((row) => targets.has(String(row.position ?? '').trim().toLowerCase())),
    (row) => row.sugarGood
  );
}

function personKey(position: unknown, name: unknown): string {
  return `${String(position ?? '').trim() || '-'}__${String(name ?? '').trim() || '-'}`;
}

function normalizeDocMonth(dateValue: unknown): string {
  const match = String(dateValue ?? '').trim().match(/^(\d{4})-(\d{1,2})/);
  if (!match) return '';
  return `${match[1]}-${String(Number(match[2])).padStart(2, '0')}`;
}

function getMonthlyLocationFieldKey(location: string): string {
  const raw = String(location ?? '').trim();
  return /[.#$/\[\]]/.test(raw) ? raw.replace(/[.#$/\[\]]/g, '_') : raw;
}

function getMonthlyProductionDocId(month: string, office: string): string {
  const monthId = normalizeDocMonth(month) || String(month ?? '').trim();
  return `${monthId}_${getMonthlyLocationFieldKey(office)}`;
}

function normalizeReviewSummary(raw: unknown): ReviewSummarySnapshot | null {
  if (!raw || typeof raw !== 'object') return null;
  const data = raw as ReviewSummarySnapshot;
  if (!data.actual || !data.ratingCounts || !data.dailyAverage || !data.facebook) return null;
  return data;
}

function normalizeIssueSummary(raw: unknown): IssueSummarySnapshot | null {
  if (!raw || typeof raw !== 'object') return null;
  const data = raw as Record<string, unknown>;
  const monthly = getDocFieldString(data.monthly);
  const yearly = getDocFieldString(data.yearly);
  const freeDays = getDocFieldString(data.freeDays);
  if (!monthly && !yearly && !freeDays) return null;
  return { monthly, yearly, freeDays };
}

function getMonthlyDocIdCandidates(month: string, office: string): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const docId of [
    getMonthlyProductionDocId(month, office),
    `${normalizeDocMonth(month) || month}_${office.trim()}`,
  ]) {
    if (!docId || seen.has(docId)) continue;
    seen.add(docId);
    result.push(docId);
  }
  return result;
}

async function fetchMonthlyDoc(collectionName: string, month: string, office: string) {
  for (const docId of getMonthlyDocIdCandidates(month, office)) {
    const snap = await getDoc(doc(db, collectionName, docId));
    if (snap.exists()) return snap.data() as Record<string, unknown>;
  }
  return null;
}

async function loadReportSummariesFromDb(month: string, office: string) {
  const data = await fetchMonthlyDoc(PINK_BEAR_COLLECTION, month, office);
  if (!data) return { reviewSummary: null, issueSummary: null };
  const reviews = data.reviews as { summary?: unknown } | undefined;
  return {
    reviewSummary: normalizeReviewSummary(reviews?.summary),
    issueSummary: normalizeIssueSummary(data.issues),
  };
}

function formatSummaryNumber(value: number | null | undefined, fallback = '—'): string {
  if (value === null || value === undefined) return fallback;
  return String(value);
}

function formatSummaryText(value: string | undefined): string {
  return value?.trim() ? value : '—';
}

function emptyOfficeSummaryGoalOverrides(): OfficeSummaryGoalOverrides {
  return { monthlyGoal: {}, dailyGoal: {} };
}

function emptyOeTableGoalOverrides(): OeTableGoalOverrides {
  return { monthlyGoal: {}, dailyGoal: {} };
}

function normalizeClean(raw: unknown): string | undefined {
  if (raw === undefined || raw === null || raw === '') return undefined;
  if (typeof raw === 'string') {
    const trimmed = raw.trim();
    return trimmed === '' ? undefined : trimmed;
  }
  return undefined;
}

function hasClean(value: string | undefined): boolean {
  return value !== undefined && value.trim() !== '';
}

function normalizeNote(raw: unknown): string | undefined {
  if (raw === undefined || raw === null || raw === '') return undefined;
  if (typeof raw === 'string') {
    const trimmed = raw.trim();
    return trimmed === '' ? undefined : raw;
  }
  return undefined;
}

function hasNote(value: string | undefined): boolean {
  return value !== undefined && value.trim() !== '';
}

function emptyPromotionActuals(): PromotionActuals {
  return {};
}

function normalizePromotionActuals(raw: unknown): PromotionActuals {
  if (!raw || typeof raw !== 'object') return emptyPromotionActuals();
  const data = raw as Record<string, unknown>;
  const out: PromotionActuals = {};
  (Object.values(PROMOTION_ACTUAL_KEYS) as PromotionActualKey[]).forEach((key) => {
    if (data[key] === undefined || data[key] === null || data[key] === '') return;
    const n = parseNumber(data[key]);
    if (Number.isFinite(n) && n >= 0) out[key] = n;
  });
  return out;
}

function hasPromotionActuals(values: PromotionActuals): boolean {
  return (Object.values(PROMOTION_ACTUAL_KEYS) as PromotionActualKey[]).some((key) => values[key] !== undefined);
}

function updatePromotionActual(prev: PromotionActuals, key: PromotionActualKey, rawValue: string): PromotionActuals {
  const next = { ...prev };
  const trimmed = rawValue.trim();
  if (trimmed === '') {
    delete next[key];
  } else {
    const parsed = parseNumber(trimmed);
    if (Number.isFinite(parsed) && parsed >= 0) {
      next[key] = Math.round((parsed + Number.EPSILON) * 100) / 100;
    }
  }
  return next;
}

function emptyWaitTimes(): WaitTimes {
  return {};
}

function normalizeWaitTimes(raw: unknown): WaitTimes {
  if (!raw || typeof raw !== 'object') return emptyWaitTimes();
  const data = raw as Record<string, unknown>;
  const out: WaitTimes = {};
  WAIT_TIME_ROWS.forEach(({ key }) => {
    if (data[key] === undefined || data[key] === null || data[key] === '') return;
    const n = parseNumber(data[key]);
    if (Number.isFinite(n) && n >= 0) out[key] = n;
  });
  return out;
}

function hasWaitTimes(values: WaitTimes): boolean {
  return WAIT_TIME_ROWS.some(({ key }) => values[key] !== undefined);
}

function updateWaitTime(prev: WaitTimes, key: WaitTimeKey, rawValue: string): WaitTimes {
  const next = { ...prev };
  const trimmed = rawValue.trim();
  if (trimmed === '') {
    delete next[key];
  } else {
    const parsed = parseNumber(trimmed);
    if (Number.isFinite(parsed) && parsed >= 0) {
      next[key] = Math.round((parsed + Number.EPSILON) * 100) / 100;
    }
  }
  return next;
}

function emptyExpenses(): Expenses {
  return {};
}

function normalizeExpenses(raw: unknown): Expenses {
  if (!raw || typeof raw !== 'object') return emptyExpenses();
  const data = raw as Record<string, unknown>;
  const out: Expenses = {};
  EXPENSE_ROWS.forEach(({ key }) => {
    if (data[key] === undefined || data[key] === null || data[key] === '') return;
    const n = parseNumber(data[key]);
    if (Number.isFinite(n) && n >= 0) out[key] = n;
  });
  EXPENSE_CUSTOM_ROW_KEYS.forEach((key) => {
    const labelKey = getExpenseLabelKey(key);
    const rawLabel = data[labelKey];
    if (typeof rawLabel !== 'string') return;
    const trimmed = rawLabel.trim();
    if (trimmed !== '') out[labelKey] = trimmed;
  });
  return out;
}

function hasExpenses(values: Expenses): boolean {
  return EXPENSE_ROWS.some(({ key }) => key !== 'total' && values[key] !== undefined)
    || EXPENSE_CUSTOM_ROW_KEYS.some((key) => values[getExpenseLabelKey(key)] !== undefined);
}

function sumExpenseAmounts(values: Expenses): number {
  return EXPENSE_ROWS.reduce((sum, { key }) => {
    if (key === 'total') return sum;
    const amount = values[key];
    if (amount === undefined) return sum;
    return Math.round((sum + amount + Number.EPSILON) * 100) / 100;
  }, 0);
}

function withExpenseTotal(values: Expenses): Expenses {
  if (!hasExpenses(values)) return values;
  return { ...values, total: sumExpenseAmounts(values) };
}

function getExpenseFieldKeys(): string[] {
  return [
    ...EXPENSE_ROWS.map(({ key }) => key),
    ...EXPENSE_CUSTOM_ROW_KEYS.flatMap((key) => [key, getExpenseLabelKey(key)]),
  ];
}

function applyRemovedNestedFields<T extends Record<string, unknown>>(
  draft: T,
  saved: T,
  keys: readonly string[],
): Record<string, unknown> {
  const payload: Record<string, unknown> = { ...draft };
  keys.forEach((key) => {
    if (saved[key] !== undefined && draft[key] === undefined) {
      payload[key] = deleteField();
    }
  });
  return payload;
}

function buildExpensesForSave(draft: Expenses, saved: Expenses) {
  const withTotal = withExpenseTotal(draft);
  if (!hasExpenses(withTotal)) {
    return deleteField();
  }
  return applyRemovedNestedFields(withTotal, saved, getExpenseFieldKeys());
}

function buildWaitTimesForSave(draft: WaitTimes, saved: WaitTimes) {
  if (!hasWaitTimes(draft)) {
    return deleteField();
  }
  return applyRemovedNestedFields(
    draft,
    saved,
    WAIT_TIME_ROWS.map(({ key }) => key),
  );
}

function buildPromotionActualsForSave(draft: PromotionActuals, saved: PromotionActuals) {
  if (!hasPromotionActuals(draft)) {
    return deleteField();
  }
  return applyRemovedNestedFields(
    draft,
    saved,
    Object.values(PROMOTION_ACTUAL_KEYS) as PromotionActualKey[],
  );
}

function updateExpense(prev: Expenses, key: ExpenseKey, rawValue: string): Expenses {
  if (key === 'total') return prev;
  const next = { ...prev };
  const trimmed = rawValue.trim();
  if (trimmed === '') {
    delete next[key];
  } else {
    const parsed = parseNumber(trimmed);
    if (Number.isFinite(parsed) && parsed >= 0) {
      next[key] = Math.round((parsed + Number.EPSILON) * 100) / 100;
    }
  }
  return next;
}

function updateExpenseLabel(prev: Expenses, key: ExpenseCustomRowKey, rawValue: string): Expenses {
  const next = { ...prev };
  const labelKey = getExpenseLabelKey(key);
  const trimmed = rawValue.trim();
  if (trimmed === '') {
    delete next[labelKey];
  } else {
    next[labelKey] = trimmed;
  }
  return next;
}

function normalizeGoalSection<K extends string>(raw: unknown, keys: readonly K[]): Partial<Record<K, number>> {
  if (!raw || typeof raw !== 'object') return {};
  const data = raw as Record<string, unknown>;
  const out: Partial<Record<K, number>> = {};
  keys.forEach((key) => {
    if (data[key] === undefined || data[key] === null || data[key] === '') return;
    const n = parseNumber(data[key]);
    if (Number.isFinite(n) && n >= 0) out[key] = n;
  });
  return out;
}

function normalizeOfficeSummaryGoalOverrides(raw: unknown): OfficeSummaryGoalOverrides {
  if (!raw || typeof raw !== 'object') return emptyOfficeSummaryGoalOverrides();
  const data = raw as Record<string, unknown>;
  return stripOfficeSummaryNoGoalFields({
    monthlyGoal: normalizeGoalSection(data.monthlyGoal, OFFICE_SUMMARY_GOAL_KEYS),
    dailyGoal: normalizeGoalSection(data.dailyGoal, OFFICE_SUMMARY_GOAL_KEYS),
  });
}

function stripOfficeSummaryNoGoalFields(goals: OfficeSummaryGoalOverrides): OfficeSummaryGoalOverrides {
  const nextDaily = { ...goals.dailyGoal };
  const nextMonthly = { ...goals.monthlyGoal };
  OFFICE_SUMMARY_NO_GOAL_KEYS.forEach((key) => {
    delete nextDaily[key];
    delete nextMonthly[key];
  });
  return { dailyGoal: nextDaily, monthlyGoal: nextMonthly };
}

function normalizeOeTableGoalOverrides(raw: unknown): OeTableGoalOverrides {
  if (!raw || typeof raw !== 'object') return emptyOeTableGoalOverrides();
  const data = raw as Record<string, unknown>;
  return {
    monthlyGoal: normalizeGoalSection(data.monthlyGoal, OE_TABLE_GOAL_KEYS),
    dailyGoal: normalizeGoalSection(data.dailyGoal, OE_TABLE_GOAL_KEYS),
  };
}

function parseStoredGoalTotal(value: unknown): number | undefined {
  if (value === undefined || value === null || value === '') return undefined;
  const parsed = parseNumber(value);
  if (!Number.isFinite(parsed)) return undefined;
  return parsed;
}

async function loadGoalSettings(month: string, office: string) {
  for (const docId of getMonthlyDocIdCandidates(month, office)) {
    const snap = await getDoc(doc(db, 'monthly production', docId));
    if (!snap.exists()) continue;
    const data = snap.data() as Record<string, unknown>;
    const oGoalTotal = parseStoredGoalTotal(data['oe goal total']);
    const sealantGoalTotal = parseStoredGoalTotal(data['sealant goal total']);
    if (
      data.officeSummaryGoals === undefined &&
      data.oeTableGoals === undefined &&
      data.started === undefined &&
      !hasClean(normalizeClean(data.clean ?? data.lobbyMetrics)) &&
      !hasWaitTimes(normalizeWaitTimes(data.waitTimes)) &&
      !hasExpenses(normalizeExpenses(data.expenses)) &&
      !hasPromotionActuals(normalizePromotionActuals(data.promotionActuals)) &&
      !hasNote(normalizeNote(data['summary note'] ?? data.note)) &&
      oGoalTotal === undefined &&
      sealantGoalTotal === undefined
    ) continue;

    const startedRaw = data.started;
    const started = startedRaw === undefined || startedRaw === null || startedRaw === ''
      ? undefined
      : parseNumber(startedRaw);

    return {
      officeGoals: normalizeOfficeSummaryGoalOverrides(data.officeSummaryGoals),
      oeGoals: normalizeOeTableGoalOverrides(data.oeTableGoals),
      started: started !== undefined && Number.isFinite(started) && started >= 0 ? started : undefined,
      clean: normalizeClean(data.clean ?? data.lobbyMetrics),
      waitTimes: normalizeWaitTimes(data.waitTimes),
      expenses: normalizeExpenses(data.expenses),
      promotionActuals: normalizePromotionActuals(data.promotionActuals),
      note: normalizeNote(data['summary note'] ?? data.note),
      oGoalTotal,
      sealantGoalTotal,
    };
  }

  return null;
}

function buildAggregate(monthlyDocs: FormDoc[]): Aggregate | null {
  if (monthlyDocs.length === 0) return null;

  const agg: Aggregate = {
    docCount: 0,
    coffeeSales: 0,
    prophyTotal: 0,
    pineapple: 0,
    rose: 0,
    mailedProduction: 0,
    visitsAdd: 0,
    visitsNoShow: 0,
    visitsScheduled: 0,
    visitsSeen: 0,
    visitsPostcard: 0,
    visitsReferral: 0,
    doctorCustomerTotal: 0,
    orangeJuiceNew: 0,
    orangeJuiceReturn: 0,
    orangeJuiceTotal: 0,
    coffeeRows: [],
    sugarRows: [],
  };

  monthlyDocs.forEach((doc) => {
    agg.docCount += 1;
    agg.coffeeSales = addAmount(agg.coffeeSales, doc.coffeeSales);
    agg.prophyTotal = addAmount(agg.prophyTotal, doc.prophyTotal);
    agg.pineapple = addAmount(agg.pineapple, doc.locationSummary?.pineapple);
    agg.rose = addAmount(agg.rose, doc.locationSummary?.rose);
    agg.mailedProduction = addAmount(agg.mailedProduction, doc.locationSummary?.mailedProduction);

    const sideMetrics = doc.productionSideMetrics;
    if (sideMetrics) {
      agg.visitsAdd = addAmount(agg.visitsAdd, sideMetrics.add);
      agg.visitsNoShow = addAmount(agg.visitsNoShow, sideMetrics.noShow);
      agg.visitsScheduled = addAmount(agg.visitsScheduled, sideMetrics.scheduled);
      agg.visitsSeen = addAmount(agg.visitsSeen, sideMetrics.seen);
      agg.visitsPostcard = addAmount(agg.visitsPostcard, sideMetrics.postcard);
      agg.visitsReferral = addAmount(agg.visitsReferral, sideMetrics.referral);
    }

    (doc.extraInputRows || []).forEach((row) => {
      agg.doctorCustomerTotal = addAmount(agg.doctorCustomerTotal, row.customer);
    });

    const coffeeActualTotals = doc.coffeeActualTotals;
    if (coffeeActualTotals) {
      agg.orangeJuiceNew = addAmount(agg.orangeJuiceNew, coffeeActualTotals.orangeJuiceNew);
      agg.orangeJuiceReturn = addAmount(agg.orangeJuiceReturn, coffeeActualTotals.orangeJuiceReturn);
      agg.orangeJuiceTotal = addAmount(agg.orangeJuiceTotal, coffeeActualTotals.orangeJuiceTotal);
    }

    const coffeeMap = new Map(agg.coffeeRows.map((r) => [personKey(r.position, r.name), r]));
    (doc.tableRows || []).forEach((row) => {
      const pKey = personKey(row.position, row.name);
      const prev = coffeeMap.get(pKey) || {
        position: pKey.split('__')[0],
        name: pKey.split('__')[1],
        coffeeNew: 0,
        coffeeReturn: 0,
        coffeeYes: 0,
      };
      prev.coffeeNew = addAmount(prev.coffeeNew, row.coffeeNew);
      prev.coffeeReturn = addAmount(prev.coffeeReturn, row.coffeeReturn);
      prev.coffeeYes = addAmount(prev.coffeeYes, row.coffeeYes);
      coffeeMap.set(pKey, prev);
    });
    agg.coffeeRows = Array.from(coffeeMap.values()).sort((a, b) => `${a.position}_${a.name}`.localeCompare(`${b.position}_${b.name}`));

    const sugarMap = new Map(agg.sugarRows.map((r) => [personKey(r.position, r.name), r]));
    (doc.sugarRows || []).forEach((row) => {
      const pKey = personKey(row.position, row.name);
      const prev = sugarMap.get(pKey) || { position: pKey.split('__')[0], name: pKey.split('__')[1], sugarGood: 0 };
      prev.sugarGood = addAmount(prev.sugarGood, row.sugarGood);
      sugarMap.set(pKey, prev);
    });
    agg.sugarRows = Array.from(sugarMap.values()).sort((a, b) => `${a.position}_${a.name}`.localeCompare(`${b.position}_${b.name}`));
  });

  return agg;
}

function getGoalsActualMetrics(aggregate: Aggregate, oeTotalAmount: number) {
  return {
    preventative: aggregate.pineapple,
    restorative: aggregate.rose,
    craProduction: aggregate.coffeeSales,
    firstReviewProduction: aggregate.pineapple + aggregate.rose + aggregate.coffeeSales,
    mailedProduction: aggregate.mailedProduction,
    oeNp: aggregate.orangeJuiceNew,
    oeRc: aggregate.orangeJuiceReturn,
    oeTotal: oeTotalAmount,
    actualProphy: aggregate.prophyTotal,
    sealantRda: sumSugarRowsByPosition(aggregate.sugarRows, ['rda']),
    sealantDds: sumSugarRowsByPosition(aggregate.sugarRows, ['doctor']),
  };
}

type GoalsActualMetrics = ReturnType<typeof getGoalsActualMetrics>;

type GoalColumnDef<K extends string> = {
  key: K;
  header: string;
  getValue: (metrics: GoalsActualMetrics) => number;
  formatTotal: (value: number) => string;
  formatDaily: (value: number, days: number) => string;
  formatGoal: (value: number | undefined) => string;
  goalStep: string;
};

const OFFICE_SUMMARY_COLUMNS: GoalColumnDef<OfficeSummaryGoalKey>[] = [
  { key: 'preventative', header: 'Preventative', getValue: (m) => m.preventative, formatTotal: formatDollarAmount, formatDaily: formatDollarDailyAverage, formatGoal: (v) => (v === undefined ? '' : formatDollarAmount(v)), goalStep: '0.01' },
  { key: 'restorative', header: 'Restorative', getValue: (m) => m.restorative, formatTotal: formatDollarAmount, formatDaily: formatDollarDailyAverage, formatGoal: (v) => (v === undefined ? '' : formatDollarAmount(v)), goalStep: '0.01' },
  { key: 'craProduction', header: 'CRA Production', getValue: (m) => m.craProduction, formatTotal: formatDollarAmount, formatDaily: formatDollarDailyAverage, formatGoal: (v) => (v === undefined ? '' : formatDollarAmount(v)), goalStep: '0.01' },
  { key: 'firstReviewProduction', header: 'Total', getValue: (m) => m.firstReviewProduction, formatTotal: formatDollarAmount, formatDaily: formatDollarDailyAverage, formatGoal: (v) => (v === undefined ? '' : formatDollarAmount(v)), goalStep: '0.01' },
  { key: 'mailedProduction', header: 'Mailed Production', getValue: (m) => m.mailedProduction, formatTotal: formatDollarAmount, formatDaily: formatDollarDailyAverage, formatGoal: (v) => (v === undefined ? '' : formatDollarAmount(v)), goalStep: '0.01' },
];

const OE_CORE_COLUMNS: GoalColumnDef<OeCoreGoalKey>[] = [
  { key: 'oeNp', header: 'OE (NP)', getValue: (m) => m.oeNp, formatTotal: formatInteger, formatDaily: formatIntegerDailyAverage, formatGoal: formatIntegerGoal, goalStep: '1' },
  { key: 'oeRc', header: 'OE (RC)', getValue: (m) => m.oeRc, formatTotal: formatInteger, formatDaily: formatIntegerDailyAverage, formatGoal: formatIntegerGoal, goalStep: '1' },
  { key: 'oeTotal', header: 'OE Total', getValue: (m) => m.oeTotal, formatTotal: formatInteger, formatDaily: formatIntegerDailyAverage, formatGoal: formatIntegerGoal, goalStep: '1' },
];

const ACTUAL_PROPHY_COLUMNS: GoalColumnDef<OeProphyGoalKey>[] = [
  { key: 'actualProphy', header: 'Actual Prophy', getValue: (m) => m.actualProphy, formatTotal: formatInteger, formatDaily: formatIntegerDailyAverage, formatGoal: formatIntegerGoal, goalStep: '1' },
];

const SEALANT_COLUMNS: GoalColumnDef<SealantGoalKey>[] = [
  { key: 'sealantRda', header: 'Sealant (RDA)', getValue: (m) => m.sealantRda, formatTotal: formatInteger, formatDaily: formatIntegerDailyAverage, formatGoal: formatIntegerGoal, goalStep: '1' },
  { key: 'sealantDds', header: 'Sealant (DDS)', getValue: (m) => m.sealantDds, formatTotal: formatInteger, formatDaily: formatIntegerDailyAverage, formatGoal: formatIntegerGoal, goalStep: '1' },
];

function GoalInputCell({
  value,
  isEditing,
  onChange,
  cellStyle: style,
  formatValue,
  step,
}: {
  value: number | undefined;
  isEditing: boolean;
  onChange: (rawValue: string) => void;
  cellStyle: React.CSSProperties;
  formatValue: (value: number | undefined) => string;
  step: string;
}) {
  if (!isEditing) return <td style={style}>{formatValue(value)}</td>;
  return (
    <td style={style}>
      <input
        type="number"
        min={0}
        step={step}
        value={value === undefined ? '' : value}
        onChange={(e) => onChange(e.target.value)}
        style={inputStyle}
      />
    </td>
  );
}

function getComputedMonthlyGoal(
  daily: number | undefined,
  storedMonthly: number | undefined,
  docCount: number
): number | undefined {
  if (daily !== undefined && docCount > 0) {
    return Math.round((daily * docCount + Number.EPSILON) * 100) / 100;
  }
  return storedMonthly;
}

function sumDefinedGoalValues<K extends string>(
  goals: Partial<Record<K, number>>,
  keys: readonly K[]
): number | undefined {
  let sum = 0;
  let hasAny = false;
  for (const key of keys) {
    const value = goals[key];
    if (value !== undefined) {
      sum = Math.round((sum + value + Number.EPSILON) * 100) / 100;
      hasAny = true;
    }
  }
  return hasAny ? sum : undefined;
}

function getOfficeSummaryTotalDailyGoal(dailyGoals: OfficeSummaryGoals): number | undefined {
  return sumDefinedGoalValues(dailyGoals, OFFICE_SUMMARY_TOTAL_COMPONENT_KEYS);
}

function getOfficeSummaryTotalMonthlyGoal(
  dailyGoals: OfficeSummaryGoals,
  monthlyGoals: OfficeSummaryGoals,
  docCount: number
): number | undefined {
  let sum = 0;
  let hasAny = false;
  for (const key of OFFICE_SUMMARY_TOTAL_COMPONENT_KEYS) {
    const monthly = getComputedMonthlyGoal(dailyGoals[key], monthlyGoals[key], docCount);
    if (monthly !== undefined) {
      sum = Math.round((sum + monthly + Number.EPSILON) * 100) / 100;
      hasAny = true;
    }
  }
  return hasAny ? sum : undefined;
}

function withOfficeSummaryTotalGoals(
  goals: OfficeSummaryGoalOverrides,
  docCount: number
): OfficeSummaryGoalOverrides {
  const totalDaily = getOfficeSummaryTotalDailyGoal(goals.dailyGoal);
  const nextDaily = { ...goals.dailyGoal };
  const nextMonthly = { ...goals.monthlyGoal };

  if (totalDaily !== undefined) {
    nextDaily.firstReviewProduction = totalDaily;
    const totalMonthly = getOfficeSummaryTotalMonthlyGoal(goals.dailyGoal, goals.monthlyGoal, docCount)
      ?? (docCount > 0 ? Math.round((totalDaily * docCount + Number.EPSILON) * 100) / 100 : undefined);
    if (totalMonthly !== undefined) {
      nextMonthly.firstReviewProduction = totalMonthly;
    } else {
      delete nextMonthly.firstReviewProduction;
    }
  } else {
    delete nextDaily.firstReviewProduction;
    delete nextMonthly.firstReviewProduction;
  }

  return stripOfficeSummaryNoGoalFields({ dailyGoal: nextDaily, monthlyGoal: nextMonthly });
}

function MetricsGoalsTable<K extends string>({
  columns,
  metrics,
  docCount,
  monthlyGoals,
  dailyGoals,
  isEditing,
  onGoalChange,
  readOnlyGoalKeys,
  getDerivedDailyGoal,
  getDerivedMonthlyGoal,
}: {
  columns: GoalColumnDef<K>[];
  metrics: GoalsActualMetrics;
  docCount: number;
  monthlyGoals: Partial<Record<K, number>>;
  dailyGoals: Partial<Record<K, number>>;
  isEditing: boolean;
  onGoalChange: (key: K, rawValue: string) => void;
  readOnlyGoalKeys?: readonly K[];
  getDerivedDailyGoal?: (key: K, dailyGoals: Partial<Record<K, number>>) => number | undefined;
  getDerivedMonthlyGoal?: (
    key: K,
    dailyGoals: Partial<Record<K, number>>,
    monthlyGoals: Partial<Record<K, number>>,
    docCount: number
  ) => number | undefined;
}) {
  const readOnlyKeys = new Set(readOnlyGoalKeys ?? []);

  return (
    <table style={getTableStyle()}>
      <thead>
        <tr>
          <th style={headerCellStyle} />
          {columns.map((column) => (
            <th key={column.key} style={headerCellStyle}>{column.header}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        <tr style={rowStyle}>
          <td style={cellStyle}>Daily Goal</td>
          {columns.map((column) => {
            const dailyValue = getDerivedDailyGoal?.(column.key, dailyGoals) ?? dailyGoals[column.key];
            if (readOnlyKeys.has(column.key)) {
              return (
                <td key={`daily_goal_${column.key}`} style={cellStyle}>
                  {column.formatGoal(dailyValue)}
                </td>
              );
            }
            return (
              <GoalInputCell
                key={`daily_goal_${column.key}`}
                value={dailyValue}
                isEditing={isEditing}
                onChange={(rawValue) => onGoalChange(column.key, rawValue)}
                cellStyle={cellStyle}
                formatValue={column.formatGoal}
                step={column.goalStep}
              />
            );
          })}
        </tr>
        <tr style={rowStyle}>
          <td style={cellStyle}>Monthly Goal</td>
          {columns.map((column) => {
            const monthlyValue = getDerivedMonthlyGoal?.(column.key, dailyGoals, monthlyGoals, docCount)
              ?? getComputedMonthlyGoal(
                dailyGoals[column.key],
                monthlyGoals[column.key],
                docCount
              );
            return (
              <td key={`monthly_goal_${column.key}`} style={cellStyle}>
                {column.formatGoal(monthlyValue)}
              </td>
            );
          })}
        </tr>
        <tr style={rowStyle}>
          <td style={cellStyle}>Total</td>
          {columns.map((column) => (
            <td key={`total_${column.key}`} style={cellStyle}>
              {column.formatTotal(column.getValue(metrics))}
            </td>
          ))}
        </tr>
        <tr style={rowStyle}>
          <td style={cellStyle}>Daily Average</td>
          {columns.map((column) => (
            <td key={`daily_${column.key}`} style={cellStyle}>
              {column.formatDaily(column.getValue(metrics), docCount)}
            </td>
          ))}
        </tr>
      </tbody>
    </table>
  );
}

function getSealantMonthlyGoal(oeTotal: number): number {
  return Math.round(oeTotal * 20 * 0.03);
}

function getSealantDailyGoal(oeTotal: number, days: number): number {
  if (days <= 0) return 0;
  return Math.round(getSealantMonthlyGoal(oeTotal) / days);
}

function SealantGoalsTable({
  metrics,
  docCount,
}: {
  metrics: GoalsActualMetrics;
  docCount: number;
}) {
  const sealantMonthlyGoal = getSealantMonthlyGoal(metrics.oeTotal);
  const sealantDailyGoal = getSealantDailyGoal(metrics.oeTotal, docCount);
  const sealantRdaTotal = metrics.sealantRda;
  const sealantDdsTotal = metrics.sealantDds;
  const sealantCombinedTotal = sealantRdaTotal + sealantDdsTotal;
  const sealantRdaPct = sealantCombinedTotal > 0
    ? Math.round((sealantRdaTotal / sealantCombinedTotal) * 100)
    : 0;
  const sealantDdsPct = sealantCombinedTotal > 0 ? 100 - sealantRdaPct : 0;

  return (
    <table style={getTableStyle()}>
      <thead>
        <tr>
          <th style={headerCellStyle} />
          {SEALANT_COLUMNS.map((column) => (
            <th key={column.key} style={headerCellStyle}>{column.header}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        <tr style={rowStyle}>
          <td style={cellStyle}>%</td>
          <td style={cellStyle}>{sealantRdaPct}%</td>
          <td style={cellStyle}>{sealantDdsPct}%</td>
        </tr>
        <tr style={rowStyle}>
          <td style={cellStyle}>Daily Goal</td>
          <td colSpan={SEALANT_COLUMNS.length} style={{ ...cellStyle, textAlign: 'center' }}>
            {sealantDailyGoal}
          </td>
        </tr>
        <tr style={rowStyle}>
          <td style={cellStyle}>Monthly Goal</td>
          <td colSpan={SEALANT_COLUMNS.length} style={{ ...cellStyle, textAlign: 'center' }}>
            {sealantMonthlyGoal}
          </td>
        </tr>
        <tr style={rowStyle}>
          <td style={cellStyle}>Total</td>
          {SEALANT_COLUMNS.map((column) => (
            <td key={`total_${column.key}`} style={cellStyle}>
              {column.formatTotal(column.getValue(metrics))}
            </td>
          ))}
        </tr>
        <tr style={rowStyle}>
          <td style={cellStyle}>Daily Average</td>
          {SEALANT_COLUMNS.map((column) => (
            <td key={`daily_${column.key}`} style={cellStyle}>
              {column.formatDaily(column.getValue(metrics), docCount)}
            </td>
          ))}
        </tr>
      </tbody>
    </table>
  );
}

function GoalsTablesSection({
  aggregate,
  oeTotalAmount,
  officeMonthlyGoals,
  officeDailyGoals,
  oeMonthlyGoals,
  oeDailyGoals,
  isEditing,
  onOfficeGoalChange,
  onOeGoalChange,
}: {
  aggregate: Aggregate;
  oeTotalAmount: number;
  officeMonthlyGoals: OfficeSummaryGoals;
  officeDailyGoals: OfficeSummaryGoals;
  oeMonthlyGoals: OeTableGoals;
  oeDailyGoals: OeTableGoals;
  isEditing: boolean;
  onOfficeGoalChange: (key: OfficeSummaryGoalKey, rawValue: string) => void;
  onOeGoalChange: (key: OeTableGoalKey, rawValue: string) => void;
}) {
  const metrics = getGoalsActualMetrics(aggregate, oeTotalAmount);
  const { docCount } = aggregate;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12, width: '100%', minWidth: 0 }}>
      <TableCard>
        <MetricsGoalsTable
          columns={OFFICE_SUMMARY_COLUMNS}
          metrics={metrics}
          docCount={docCount}
          monthlyGoals={officeMonthlyGoals}
          dailyGoals={officeDailyGoals}
          isEditing={isEditing}
          onGoalChange={onOfficeGoalChange}
          readOnlyGoalKeys={['firstReviewProduction', 'mailedProduction']}
          getDerivedDailyGoal={(key, dailyGoals) => {
            if (key === 'mailedProduction') return undefined;
            if (key === 'firstReviewProduction') return getOfficeSummaryTotalDailyGoal(dailyGoals);
            return undefined;
          }}
          getDerivedMonthlyGoal={(key, dailyGoals, monthlyGoals, days) => {
            if (key === 'mailedProduction') return undefined;
            if (key === 'firstReviewProduction') {
              return getOfficeSummaryTotalMonthlyGoal(dailyGoals, monthlyGoals, days);
            }
            return undefined;
          }}
        />
      </TableCard>
      <div style={tableGridStyle}>
        <TableCard>
          <MetricsGoalsTable
            columns={OE_CORE_COLUMNS}
            metrics={metrics}
            docCount={docCount}
            monthlyGoals={oeMonthlyGoals}
            dailyGoals={oeDailyGoals}
            isEditing={isEditing}
            onGoalChange={onOeGoalChange}
          />
        </TableCard>
        <TableCard>
          <MetricsGoalsTable
            columns={ACTUAL_PROPHY_COLUMNS}
            metrics={metrics}
            docCount={docCount}
            monthlyGoals={oeMonthlyGoals}
            dailyGoals={oeDailyGoals}
            isEditing={isEditing}
            onGoalChange={onOeGoalChange}
          />
        </TableCard>
        <TableCard>
          <SealantGoalsTable
            metrics={metrics}
            docCount={docCount}
          />
        </TableCard>
      </div>
    </div>
  );
}

function getCraSummaryTotals(aggregate: Aggregate) {
  return {
    craNew: sumCoffeeRows(aggregate.coffeeRows, (r) => r.coffeeNew),
    craReturn: sumCoffeeRows(aggregate.coffeeRows, (r) => r.coffeeReturn),
    craBillable: sumCoffeeRows(aggregate.coffeeRows, (r) => r.coffeeYes),
  };
}

function CraSummaryTable({ aggregate }: { aggregate: Aggregate }) {
  const totals = getCraSummaryTotals(aggregate);
  const values = [totals.craNew, totals.craReturn, totals.craBillable];

  return (
    <table style={getTableStyle()}>
      <thead>
        <tr>
          <th style={headerCellStyle} />
          {CRA_SUMMARY_HEADERS.map((h) => (
            <th key={h} style={headerCellStyle}>{h}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        <tr style={rowStyle}>
          <td style={cellStyle}>Total</td>
          {values.map((value, idx) => (
            <td key={`cra_total_${idx}`} style={cellStyle}>{formatAmount(value)}</td>
          ))}
        </tr>
      </tbody>
    </table>
  );
}

function CraNonCraTable({ aggregate, oeTotalAmount }: { aggregate: Aggregate; oeTotalAmount: number }) {
  const craBillable = getCraSummaryTotals(aggregate).craBillable;
  const craPct = oeTotalAmount > 0 ? Math.round((craBillable / oeTotalAmount) * 100) : 0;
  const nonCraPct = oeTotalAmount > 0 ? 100 - craPct : 0;
  const values: Record<(typeof CRA_NON_CRA_HEADERS)[number], string> = {
    '>6YRS (Non CRA)': `${nonCraPct}%`,
    '<6YRS (CRA)': `${craPct}%`,
  };

  return (
    <table style={getTableStyle()}>
      <thead>
        <tr>
          {CRA_NON_CRA_HEADERS.map((h) => (
            <th key={h} style={headerCellStyle}>{h}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        <tr style={rowStyle}>
          {CRA_NON_CRA_HEADERS.map((h) => (
            <td key={`cra_non_cra_${h}`} style={cellStyle}>{values[h]}</td>
          ))}
        </tr>
      </tbody>
    </table>
  );
}

function VisitsTable({ aggregate }: { aggregate: Aggregate }) {
  return (
    <table style={getTableStyle()}>
      <thead>
        <tr>
          {VISITS_HEADERS.map((h) => (
            <th key={h} style={headerCellStyle}>{h}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        <tr style={rowStyle}>
          <td style={cellStyle}>{formatInteger(aggregate.visitsAdd)}</td>
          <td style={cellStyle}>{formatInteger(aggregate.visitsNoShow)}</td>
          <td style={cellStyle}>{formatInteger(aggregate.visitsScheduled)}</td>
          <td style={cellStyle}>
            {formatPercentage(aggregate.visitsSeen, aggregate.visitsScheduled)}
          </td>
          <td style={cellStyle}>{formatInteger(aggregate.visitsPostcard)}</td>
        </tr>
      </tbody>
    </table>
  );
}

function SeenTable({ officeSeen, doctorSeen }: { officeSeen: number; doctorSeen: number }) {
  return (
    <table style={getTableStyle()}>
      <thead>
        <tr>
          {SEEN_HEADERS.map((h) => (
            <th key={h} style={headerCellStyle}>{h}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        <tr style={rowStyle}>
          <td style={cellStyle}>{formatAmount(officeSeen)}</td>
          <td style={cellStyle}>{formatAmount(doctorSeen)}</td>
        </tr>
      </tbody>
    </table>
  );
}

type BonusDeductContext = {
  reviewSummary: ReviewSummarySnapshot | null;
  issueSummary: IssueSummarySnapshot | null;
  oGoalTotal: number | undefined;
  sealantGoalTotal: number | undefined;
  sealantRda: number;
  clean: string | undefined;
  started: number | undefined;
};

function getBonusDeductHashNumber(
  label: (typeof BONUS_DEDUCT_ROW_LABELS)[number],
  ctx: BonusDeductContext
): number | undefined {
  const { reviewSummary, issueSummary, oGoalTotal, sealantGoalTotal, sealantRda, clean, started } = ctx;

  switch (label) {
    case '5* Reviews':
      if (!reviewSummary) return undefined;
      return reviewSummary.ratingCounts.google['5'] + reviewSummary.ratingCounts.yelp['5'];
    case '1* Reviews':
      if (!reviewSummary) return undefined;
      return reviewSummary.ratingCounts.google['1'] + reviewSummary.ratingCounts.yelp['1'];
    case 'Exams^Daily Goal':
      return oGoalTotal;
    case 'Sealants':
      return sealantGoalTotal;
    case 'RDA Sealants':
      return sealantRda;
    case 'Ortho Starts':
      return started;
    case 'Injury':
      if (!issueSummary) return undefined;
      return parseNumber(issueSummary.monthly || '0');
    case 'Cleanliness':
      if (!hasClean(clean)) return undefined;
      return parseNumber(clean);
    default:
      return undefined;
  }
}

function getBonusDeductBonusNumber(
  label: (typeof BONUS_DEDUCT_ROW_LABELS)[number],
  ctx: BonusDeductContext
): number {
  const multiplier = BONUS_DEDUCT_BONUS_MULTIPLIERS[label];
  if (multiplier === undefined) return 0;
  const hashNumber = getBonusDeductHashNumber(label, ctx);
  if (hashNumber === undefined) return 0;
  return Math.round((hashNumber * multiplier + Number.EPSILON) * 100) / 100;
}

function getBonusDeductDeductNumber(
  label: (typeof BONUS_DEDUCT_ROW_LABELS)[number],
  ctx: BonusDeductContext
): number {
  const multiplier = BONUS_DEDUCT_DEDUCT_MULTIPLIERS[label];
  if (multiplier === undefined) return 0;
  const hashNumber = getBonusDeductHashNumber(label, ctx);
  if (hashNumber === undefined) return 0;
  return Math.round((hashNumber * multiplier + Number.EPSILON) * 100) / 100;
}

function getBonusDeductRowNet(label: (typeof BONUS_DEDUCT_ROW_LABELS)[number], ctx: BonusDeductContext): number {
  return Math.round((getBonusDeductBonusNumber(label, ctx) + getBonusDeductDeductNumber(label, ctx) + Number.EPSILON) * 100) / 100;
}

function getBonusDeductTotalExcludingRSealants(ctx: BonusDeductContext): number {
  return BONUS_DEDUCT_ROW_LABELS.reduce((sum, label) => {
    if (label === 'RDA Sealants') return sum;
    return Math.round((sum + getBonusDeductRowNet(label, ctx) + Number.EPSILON) * 100) / 100;
  }, 0);
}

function getBonusDeductTotalExcludingSealantsBonus(ctx: BonusDeductContext): number {
  return BONUS_DEDUCT_ROW_LABELS.reduce((sum, label) => {
    const rowNet = label === 'Sealants'
      ? getBonusDeductDeductNumber(label, ctx)
      : getBonusDeductRowNet(label, ctx);
    return Math.round((sum + rowNet + Number.EPSILON) * 100) / 100;
  }, 0);
}

function getBonusDeductHashValue(label: (typeof BONUS_DEDUCT_ROW_LABELS)[number], ctx: BonusDeductContext): string {
  if (label === 'Cleanliness') return ctx.clean ?? '';
  if (label === 'Injury') return ctx.issueSummary ? ctx.issueSummary.monthly || '0' : '';

  const hashNumber = getBonusDeductHashNumber(label, ctx);
  if (hashNumber === undefined) return '';
  if (label === 'Ortho Starts') return formatAmount(hashNumber);
  return formatInteger(hashNumber);
}

function getBonusDeductBonusValue(label: (typeof BONUS_DEDUCT_ROW_LABELS)[number], ctx: BonusDeductContext): string {
  const multiplier = BONUS_DEDUCT_BONUS_MULTIPLIERS[label];
  if (multiplier === undefined) return '';
  const hashNumber = getBonusDeductHashNumber(label, ctx);
  if (hashNumber === undefined) return '';
  return formatAmount(getBonusDeductBonusNumber(label, ctx));
}

function getBonusDeductDeductValue(label: (typeof BONUS_DEDUCT_ROW_LABELS)[number], ctx: BonusDeductContext): string {
  const multiplier = BONUS_DEDUCT_DEDUCT_MULTIPLIERS[label];
  if (multiplier === undefined) return '';
  const hashNumber = getBonusDeductHashNumber(label, ctx);
  if (hashNumber === undefined) return '';
  return formatAmount(getBonusDeductDeductNumber(label, ctx));
}

const bonusDeductTotalRowStyle: React.CSSProperties = {
  borderTop: '2px solid #cbd5e1',
  background: '#f8fafc',
};

function getBonusDeductTotalValueStyle(value: number): React.CSSProperties {
  return {
    ...centeredCellStyle,
    ...bonusDeductTotalRowStyle,
    fontWeight: 700,
    fontSize: 14,
    color: value > 0 ? '#166534' : value < 0 ? '#b91c1c' : '#334155',
  };
}

function BonusDeductTable({
  reviewSummary,
  issueSummary,
  oGoalTotal,
  sealantGoalTotal,
  sealantRda,
  clean,
  started,
}: BonusDeductContext) {
  const context: BonusDeductContext = {
    reviewSummary,
    issueSummary,
    oGoalTotal,
    sealantGoalTotal,
    sealantRda,
    clean,
    started,
  };
  const totalExcludingRSealants = getBonusDeductTotalExcludingRSealants(context);
  const totalExcludingSealantsBonus = getBonusDeductTotalExcludingSealantsBonus(context);

  return (
    <table style={getTableStyle()}>
      <thead>
        <tr>
          <th style={headerCellStyle} />
          {BONUS_DEDUCT_COLUMN_HEADERS.map((header) => (
            <th key={header} style={headerCellStyle}>
              {header}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {BONUS_DEDUCT_ROW_LABELS.map((label) => (
          <tr key={label} style={rowStyle}>
            <td style={labelCellStyle}>{label}</td>
            <td style={cellStyle}>
              {getBonusDeductHashValue(label, context)}
            </td>
            <td style={cellStyle}>
              {getBonusDeductBonusValue(label, context)}
            </td>
            <td style={cellStyle}>
              {getBonusDeductDeductValue(label, context)}
            </td>
          </tr>
        ))}
        <tr style={rowStyle}>
          <td
            style={{
              ...labelCellStyle,
              ...bonusDeductTotalRowStyle,
              fontWeight: 700,
              color: '#1e3a8a',
            }}
          >
            Total
          </td>
          <td style={{ ...cellStyle, ...bonusDeductTotalRowStyle }} />
          <td style={getBonusDeductTotalValueStyle(totalExcludingRSealants)}>
            {formatAmount(totalExcludingRSealants)}
          </td>
          <td style={getBonusDeductTotalValueStyle(totalExcludingSealantsBonus)}>
            {formatAmount(totalExcludingSealantsBonus)}
          </td>
        </tr>
      </tbody>
    </table>
  );
}

function MonthlyNoteSection({
  value,
  isEditing,
  onChange,
}: {
  value: string | undefined;
  isEditing: boolean;
  onChange: (rawValue: string) => void;
}) {
  if (isEditing) {
    return (
      <textarea
        value={value ?? ''}
        onChange={(e) => onChange(e.target.value)}
        rows={4}
        style={{
          width: '100%',
          minHeight: 108,
          border: '1px solid #d1d5db',
          borderRadius: 8,
          padding: '10px 12px',
          fontSize: 13,
          lineHeight: 1.6,
          color: '#111827',
          resize: 'vertical',
          boxSizing: 'border-box',
          fontFamily: 'inherit',
          background: '#fff',
        }}
      />
    );
  }

  if (!hasNote(value)) {
    return null;
  }

  return (
    <p
      style={{
        margin: 0,
        color: '#334155',
        fontSize: 13,
        lineHeight: 1.6,
        whiteSpace: 'pre-wrap',
        overflowWrap: 'anywhere',
      }}
    >
      {value}
    </p>
  );
}

function DailyOfficeLogTable({ rows }: { rows: DailyOfficeLogRow[] }) {
  return (
    <table style={getTableStyle()}>
      <thead>
        <tr>
          {DAILY_OFFICE_LOG_HEADERS.map((h) => (
            <th key={h} style={headerCellStyle}>{h}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        {rows.length === 0 ? (
          <tr>
            <td colSpan={DAILY_OFFICE_LOG_HEADERS.length} style={cellStyle}>
              There's no data yet.
            </td>
          </tr>
        ) : (
          rows.map((row) => (
            <tr key={row.id} style={rowStyle}>
              <td style={cellStyle}>{row.date}</td>
              <td style={cellStyle}>
                {row.dailyProduction === undefined ? '' : formatDollarAmount(row.dailyProduction)}
              </td>
              <td style={cellStyle}>{formatTimeDisplay(row.checkIn)}</td>
              <td style={cellStyle}>{formatTimeDisplay(row.checkOut)}</td>
              <td style={cellStyle}>{row.closer}</td>
              <td style={cellStyle}>{formatHoursOpen(row.checkIn, row.checkOut)}</td>
            </tr>
          ))
        )}
      </tbody>
    </table>
  );
}

function ReferralStartedTable({
  referred,
  started,
  isEditing,
  onStartedChange,
}: {
  referred: number;
  started: number | undefined;
  isEditing: boolean;
  onStartedChange: (rawValue: string) => void;
}) {
  return (
    <table style={getTableStyle()}>
      <thead>
        <tr>
          {REFERRAL_STARTED_HEADERS.map((h) => (
            <th key={h} style={headerCellStyle}>{h}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        <tr style={rowStyle}>
          <td style={cellStyle}>{formatAmount(referred)}</td>
          <GoalInputCell
            value={started}
            isEditing={isEditing}
            onChange={onStartedChange}
            cellStyle={cellStyle}
            formatValue={(value) => (value === undefined ? '' : formatAmount(value))}
            step="1"
          />
        </tr>
      </tbody>
    </table>
  );
}

function getPromotionCellValue(rowHeader: string, columnHeader: (typeof PROMOTION_COLUMN_HEADERS)[number], days: number): string {
  const dailyGoal = PROMOTION_DAILY_GOALS[columnHeader];
  if (rowHeader === 'Daily Goal') return formatAmount(dailyGoal);
  if (rowHeader === 'Monthly Goal') return formatAmount(dailyGoal * days);
  return '';
}

function ReviewRatingRowsFromSummary({ ratingCounts }: { ratingCounts: ReviewSummarySnapshot['ratingCounts'] }) {
  const repugenLow = ratingCounts.repugen['1-2'];
  const repugenMid = ratingCounts.repugen['3'];
  const repugenHigh = ratingCounts.repugen['4-5'];

  return (
    <>
      <tr>
        <td style={{ ...labelCellStyle, textAlign: 'center' }}>1</td>
        <td style={centeredCellStyle}>{ratingCounts.google['1']}</td>
        <td style={centeredCellStyle}>{ratingCounts.yelp['1']}</td>
        <td rowSpan={2} style={centeredCellStyle}>{repugenLow}</td>
      </tr>
      <tr>
        <td style={{ ...labelCellStyle, textAlign: 'center' }}>2</td>
        <td style={centeredCellStyle}>{ratingCounts.google['2']}</td>
        <td style={centeredCellStyle}>{ratingCounts.yelp['2']}</td>
      </tr>
      <tr>
        <td style={{ ...labelCellStyle, textAlign: 'center' }}>3</td>
        <td style={centeredCellStyle}>{ratingCounts.google['3']}</td>
        <td style={centeredCellStyle}>{ratingCounts.yelp['3']}</td>
        <td style={centeredCellStyle}>{repugenMid}</td>
      </tr>
      <tr>
        <td style={{ ...labelCellStyle, textAlign: 'center' }}>4</td>
        <td style={centeredCellStyle}>{ratingCounts.google['4']}</td>
        <td style={centeredCellStyle}>{ratingCounts.yelp['4']}</td>
        <td rowSpan={2} style={centeredCellStyle}>{repugenHigh}</td>
      </tr>
      <tr>
        <td style={{ ...labelCellStyle, textAlign: 'center' }}>5</td>
        <td style={centeredCellStyle}>{ratingCounts.google['5']}</td>
        <td style={centeredCellStyle}>{ratingCounts.yelp['5']}</td>
      </tr>
    </>
  );
}

function ReviewSummaryTable({ summary }: { summary: ReviewSummarySnapshot | null }) {
  if (!summary) {
    return (
      <table style={getTableStyle()}>
        <tbody>
          <tr>
            <td style={{ ...cellStyle, color: '#94a3b8', textAlign: 'center' }} colSpan={4}>
            </td>
          </tr>
        </tbody>
      </table>
    );
  }

  return (
    <table style={getTableStyle()}>
      <thead>
        <tr>
          <th style={headerCellStyle} />
          <th style={{ ...headerCellStyle, textAlign: 'center' }}>Google</th>
          <th style={{ ...headerCellStyle, textAlign: 'center' }}>Yelp</th>
          <th style={{ ...headerCellStyle, textAlign: 'center' }}>RepuGen</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td style={labelCellStyle}>Daily Goal</td>
          <td colSpan={3} style={centeredCellStyle}>{summary.dailyGoal}</td>
        </tr>
        <tr>
          <td style={labelCellStyle}>Monthly Goal</td>
          <td colSpan={3} style={centeredCellStyle}>{formatSummaryNumber(summary.monthlyGoal)}</td>
        </tr>
        <tr>
          <td style={labelCellStyle}>Actual</td>
          <td style={centeredCellStyle}>{summary.actual.google}</td>
          <td style={centeredCellStyle}>{summary.actual.yelp}</td>
          <td style={centeredCellStyle}>{summary.actual.repugen}</td>
        </tr>
        <tr>
          <td style={labelCellStyle}>Daily Average</td>
          <td style={centeredCellStyle}>{formatSummaryNumber(summary.dailyAverage.google)}</td>
          <td style={centeredCellStyle}>{formatSummaryNumber(summary.dailyAverage.yelp)}</td>
          <td style={centeredCellStyle}>{formatSummaryNumber(summary.dailyAverage.repugen)}</td>
        </tr>
        <ReviewRatingRowsFromSummary ratingCounts={summary.ratingCounts} />
        <tr>
          <td colSpan={4} style={{ ...cellStyle, textAlign: 'center', fontWeight: 700, background: '#f8fafc', color: '#334155' }}>
            Facebook
          </td>
        </tr>
        <tr>
          <td style={labelCellStyle}>Recommended</td>
          <td style={cellStyle}>{formatSummaryText(summary.facebook.recommended)}</td>
          <td style={{ ...labelCellStyle, textAlign: 'center' }}>Not Recommended</td>
          <td style={cellStyle}>{formatSummaryText(summary.facebook.notRecommended)}</td>
        </tr>
      </tbody>
    </table>
  );
}

function IssueSummaryTable({ summary }: { summary: IssueSummarySnapshot | null }) {
  if (!summary) {
    return (
      <table style={getTableStyle()}>
        <tbody>
          <tr>
            <td style={{ ...cellStyle, color: '#94a3b8', textAlign: 'center' }} colSpan={2}>
            </td>
          </tr>
        </tbody>
      </table>
    );
  }

  return (
    <table style={getTableStyle()}>
      <tbody>
        <tr>
          <td style={labelCellStyle}>Monthly</td>
          <td style={centeredCellStyle}>{summary.monthly || '0'}</td>
        </tr>
        <tr>
          <td style={labelCellStyle}>Yearly</td>
          <td style={centeredCellStyle}>{formatSummaryText(summary.yearly)}</td>
        </tr>
        <tr>
          <td style={labelCellStyle}>Free Days</td>
          <td style={centeredCellStyle}>{formatSummaryText(summary.freeDays)}</td>
        </tr>
      </tbody>
    </table>
  );
}

function PromotionTable({
  docCount,
  actuals,
  isEditing,
  onActualChange,
}: {
  docCount: number;
  actuals: PromotionActuals;
  isEditing: boolean;
  onActualChange: (key: PromotionActualKey, rawValue: string) => void;
}) {
  return (
    <table style={getTableStyle()}>
      <thead>
        <tr>
          <th style={headerCellStyle} />
          {PROMOTION_COLUMN_HEADERS.map((h) => (
            <th key={h} style={headerCellStyle}>{h}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        {PROMOTION_ROW_HEADERS.map((rowHeader) => (
          <tr key={rowHeader} style={rowStyle}>
            <td style={cellStyle}>{rowHeader}</td>
            {PROMOTION_COLUMN_HEADERS.map((columnHeader) => {
              const actualKey = PROMOTION_ACTUAL_KEYS[columnHeader];

              if (rowHeader === 'Actual') {
                return (
                  <td key={`${rowHeader}_${columnHeader}`} style={cellStyle}>
                    {isEditing ? (
                      <input
                        type="number"
                        min={0}
                        step="1"
                        value={actuals[actualKey] === undefined ? '' : actuals[actualKey]}
                        onChange={(e) => onActualChange(actualKey, e.target.value)}
                        style={inputStyle}
                      />
                    ) : (
                      actuals[actualKey] === undefined ? '' : formatAmount(actuals[actualKey]!)
                    )}
                  </td>
                );
              }

              if (rowHeader === 'Daily Average') {
                const actual = actuals[actualKey];
                return (
                  <td key={`${rowHeader}_${columnHeader}`} style={cellStyle}>
                    {actual === undefined ? '' : formatDailyAverage(actual, docCount)}
                  </td>
                );
              }

              return (
                <td key={`${rowHeader}_${columnHeader}`} style={cellStyle}>
                  {getPromotionCellValue(rowHeader, columnHeader, docCount)}
                </td>
              );
            })}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function CleanTable({
  value,
  isEditing,
  onChange,
}: {
  value: string | undefined;
  isEditing: boolean;
  onChange: (rawValue: string) => void;
}) {
  return (
    <table style={getTableStyle()}>
      <thead>
        <tr>
          {CLEAN_HEADERS.map((h) => (
            <th key={h} style={headerCellStyle}>{h}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        <tr style={rowStyle}>
          <td colSpan={CLEAN_HEADERS.length} style={cellStyle}>
            {isEditing ? (
              <input
                type="text"
                value={value ?? ''}
                onChange={(e) => onChange(e.target.value)}
                style={inputStyle}
              />
            ) : (
              value ?? ''
            )}
          </td>
        </tr>
      </tbody>
    </table>
  );
}

function WaitTimesTable({
  values,
  isEditing,
  onChange,
}: {
  values: WaitTimes;
  isEditing: boolean;
  onChange: (key: WaitTimeKey, rawValue: string) => void;
}) {
  return (
    <table style={getTableStyle()}>
      <tbody>
        {WAIT_TIME_ROWS.map(({ key, label }) => (
          <tr key={key} style={rowStyle}>
            <td style={cellStyle}>{label}</td>
            <td style={cellStyle}>
              {isEditing ? (
                <input
                  type="number"
                  min={0}
                  step="1"
                  value={values[key] === undefined ? '' : values[key]}
                  onChange={(e) => onChange(key, e.target.value)}
                  style={inputStyle}
                />
              ) : (
                values[key] === undefined ? '' : formatAmount(values[key]!)
              )}
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function ExpensesTable({
  values,
  isEditing,
  onChange,
  onLabelChange,
}: {
  values: Expenses;
  isEditing: boolean;
  onChange: (key: ExpenseKey, rawValue: string) => void;
  onLabelChange: (key: ExpenseCustomRowKey, rawValue: string) => void;
}) {
  const expenseTotal = sumExpenseAmounts(values);

  return (
    <table style={getTableStyle()}>
      <tbody>
        {EXPENSE_ROWS.map(({ key, label }) => {
          const isCustomRow = isExpenseCustomRowKey(key);
          const isTotalRow = key === 'total';
          const labelKey = isCustomRow ? getExpenseLabelKey(key) : null;
          const displayLabel = isCustomRow ? (values[labelKey!] ?? '') : label;

          return (
            <tr key={key} style={rowStyle}>
              <td style={cellStyle}>
                {isEditing && isCustomRow ? (
                  <input
                    type="text"
                    value={values[labelKey!] ?? ''}
                    onChange={(e) => onLabelChange(key, e.target.value)}
                    style={inputStyle}
                  />
                ) : (
                  displayLabel
                )}
              </td>
              <td style={cellStyle}>
                {isTotalRow ? (
                  formatDollarAmount(expenseTotal)
                ) : isEditing ? (
                  <input
                    type="number"
                    min={0}
                    step="0.01"
                    value={values[key] === undefined ? '' : values[key]}
                    onChange={(e) => onChange(key, e.target.value)}
                    style={inputStyle}
                  />
                ) : (
                  values[key] === undefined ? '' : formatDollarAmount(values[key]!)
                )}
              </td>
            </tr>
          );
        })}
      </tbody>
    </table>
  );
}

function updateDailyGoalSection<T extends string>(
  prev: { monthlyGoal: Partial<Record<T, number>>; dailyGoal: Partial<Record<T, number>> },
  key: T,
  rawValue: string,
  docCount = 0
) {
  const nextDaily = { ...prev.dailyGoal };
  const nextMonthly = { ...prev.monthlyGoal };
  const trimmed = rawValue.trim();

  if (trimmed === '') {
    delete nextDaily[key];
    delete nextMonthly[key];
  } else {
    const parsed = parseNumber(trimmed);
    if (Number.isFinite(parsed) && parsed >= 0) {
      const daily = Math.round((parsed + Number.EPSILON) * 100) / 100;
      nextDaily[key] = daily;
      if (docCount > 0) {
        nextMonthly[key] = Math.round((daily * docCount + Number.EPSILON) * 100) / 100;
      }
    }
  }
  return { dailyGoal: nextDaily, monthlyGoal: nextMonthly };
}

function syncMonthlyGoalsFromDaily<T extends string>(
  goals: { monthlyGoal: Partial<Record<T, number>>; dailyGoal: Partial<Record<T, number>> },
  keys: readonly T[],
  docCount: number
) {
  if (docCount <= 0) return goals;
  const nextMonthly = { ...goals.monthlyGoal };
  keys.forEach((key) => {
    const daily = goals.dailyGoal[key];
    if (daily !== undefined) {
      nextMonthly[key] = Math.round((daily * docCount + Number.EPSILON) * 100) / 100;
    }
  });
  return { ...goals, monthlyGoal: nextMonthly };
}

function MonthlySummaryPageContent() {
  const searchParams = useSearchParams();
  const selectedMonth = searchParams.get('month') ?? '';
  const selectedOffice = searchParams.get('office') ?? '';
  const [docs, setDocs] = useState<FormDoc[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [saveMessage, setSaveMessage] = useState('');
  const [isEditing, setIsEditing] = useState(false);
  const [officeSummaryGoalsSaved, setOfficeSummaryGoalsSaved] = useState(emptyOfficeSummaryGoalOverrides);
  const [officeSummaryGoalsDraft, setOfficeSummaryGoalsDraft] = useState(emptyOfficeSummaryGoalOverrides);
  const [oeTableGoalsSaved, setOeTableGoalsSaved] = useState(emptyOeTableGoalOverrides);
  const [oeTableGoalsDraft, setOeTableGoalsDraft] = useState(emptyOeTableGoalOverrides);
  const [startedSaved, setStartedSaved] = useState<number | undefined>(undefined);
  const [startedDraft, setStartedDraft] = useState<number | undefined>(undefined);
  const [cleanSaved, setCleanSaved] = useState<string | undefined>(undefined);
  const [cleanDraft, setCleanDraft] = useState<string | undefined>(undefined);
  const [waitTimesSaved, setWaitTimesSaved] = useState<WaitTimes>(emptyWaitTimes);
  const [waitTimesDraft, setWaitTimesDraft] = useState<WaitTimes>(emptyWaitTimes);
  const [expensesSaved, setExpensesSaved] = useState<Expenses>(emptyExpenses);
  const [expensesDraft, setExpensesDraft] = useState<Expenses>(emptyExpenses);
  const [promotionActualsSaved, setPromotionActualsSaved] = useState<PromotionActuals>(emptyPromotionActuals);
  const [promotionActualsDraft, setPromotionActualsDraft] = useState<PromotionActuals>(emptyPromotionActuals);
  const [oGoalTotalSaved, setOGoalTotalSaved] = useState<number | undefined>(undefined);
  const [sealantGoalTotalSaved, setSealantGoalTotalSaved] = useState<number | undefined>(undefined);
  const [noteSaved, setNoteSaved] = useState<string | undefined>(undefined);
  const [noteDraft, setNoteDraft] = useState<string | undefined>(undefined);
  const [reviewSummary, setReviewSummary] = useState<ReviewSummarySnapshot | null>(null);
  const [issueSummary, setIssueSummary] = useState<IssueSummarySnapshot | null>(null);
  const hasSelection = selectedMonth !== '' && selectedOffice !== '';
  const activeOfficeSummaryGoals = isEditing ? officeSummaryGoalsDraft : officeSummaryGoalsSaved;
  const activeOeTableGoals = isEditing ? oeTableGoalsDraft : oeTableGoalsSaved;
  const activeStarted = isEditing ? startedDraft : startedSaved;
  const activeClean = isEditing ? cleanDraft : cleanSaved;
  const activeWaitTimes = isEditing ? waitTimesDraft : waitTimesSaved;
  const activeExpenses = isEditing ? expensesDraft : expensesSaved;
  const activePromotionActuals = isEditing ? promotionActualsDraft : promotionActualsSaved;
  const activeNote = isEditing ? noteDraft : noteSaved;

  const syncDraftsFromSaved = () => {
    setOfficeSummaryGoalsDraft(officeSummaryGoalsSaved);
    setOeTableGoalsDraft(oeTableGoalsSaved);
    setStartedDraft(startedSaved);
    setCleanDraft(cleanSaved);
    setWaitTimesDraft(waitTimesSaved);
    setExpensesDraft(expensesSaved);
    setPromotionActualsDraft(promotionActualsSaved);
    setNoteDraft(noteSaved);
  };

  useEffect(() => {
    const load = async () => {
      try {
        const snap = await getDocs(collection(db, 'simple-forms'));
        setDocs(snap.docs.map((d) => ({ id: d.id, ...(d.data() as Omit<FormDoc, 'id'>) })));
      } catch (e: any) {
        setError(e?.message || 'Error. Please try again.');
      } finally {
        setLoading(false);
      }
    };
    load();
  }, []);

  const monthlyDocs = useMemo(() => {
    if (!hasSelection) return [];
    return docs.filter((doc) => (
      !!doc.submittedDateTime &&
      normalizeDocMonth(doc.date) === selectedMonth &&
      String(doc.location ?? '').trim() === selectedOffice
    ));
  }, [docs, hasSelection, selectedMonth, selectedOffice]);

  const aggregate = useMemo(() => buildAggregate(monthlyDocs), [monthlyDocs]);
  const dailyOfficeLogRows = useMemo(() => buildDailyOfficeLogRows(monthlyDocs), [monthlyDocs]);
  const oeTotalAmount = useMemo(
    () => (aggregate ? aggregate.orangeJuiceTotal : 0),
    [aggregate]
  );

  useEffect(() => {
    let cancelled = false;
    const resetGoals = () => {
      if (cancelled) return;
      setOfficeSummaryGoalsSaved(emptyOfficeSummaryGoalOverrides());
      setOfficeSummaryGoalsDraft(emptyOfficeSummaryGoalOverrides());
      setOeTableGoalsSaved(emptyOeTableGoalOverrides());
      setOeTableGoalsDraft(emptyOeTableGoalOverrides());
      setStartedSaved(undefined);
      setStartedDraft(undefined);
      setCleanSaved(undefined);
      setCleanDraft(undefined);
      setWaitTimesSaved(emptyWaitTimes());
      setWaitTimesDraft(emptyWaitTimes());
      setExpensesSaved(emptyExpenses());
      setExpensesDraft(emptyExpenses());
      setPromotionActualsSaved(emptyPromotionActuals());
      setPromotionActualsDraft(emptyPromotionActuals());
      setOGoalTotalSaved(undefined);
      setSealantGoalTotalSaved(undefined);
      setNoteSaved(undefined);
      setNoteDraft(undefined);
      setIsEditing(false);
    };

    if (!hasSelection) {
      resetGoals();
      return () => { cancelled = true; };
    }

    loadGoalSettings(selectedMonth, selectedOffice)
      .then((merged) => {
        if (cancelled) return;
        if (!merged) {
          resetGoals();
          return;
        }
        setOfficeSummaryGoalsSaved(merged.officeGoals);
        setOfficeSummaryGoalsDraft(merged.officeGoals);
        setOeTableGoalsSaved(merged.oeGoals);
        setOeTableGoalsDraft(merged.oeGoals);
        setStartedSaved(merged.started);
        setStartedDraft(merged.started);
        setCleanSaved(merged.clean);
        setCleanDraft(merged.clean);
        setWaitTimesSaved(merged.waitTimes);
        setWaitTimesDraft(merged.waitTimes);
        setExpensesSaved(merged.expenses);
        setExpensesDraft(merged.expenses);
        setPromotionActualsSaved(merged.promotionActuals);
        setPromotionActualsDraft(merged.promotionActuals);
        setOGoalTotalSaved(merged.oGoalTotal);
        setSealantGoalTotalSaved(merged.sealantGoalTotal);
        setNoteSaved(merged.note);
        setNoteDraft(merged.note);
        setIsEditing(false);
      })
      .catch(() => resetGoals());

    return () => { cancelled = true; };
  }, [hasSelection, selectedMonth, selectedOffice]);

  useEffect(() => {
    let cancelled = false;

    if (!hasSelection) {
      setReviewSummary(null);
      setIssueSummary(null);
      return () => { cancelled = true; };
    }

    loadReportSummariesFromDb(selectedMonth, selectedOffice)
      .then((result) => {
        if (cancelled) return;
        setReviewSummary(result.reviewSummary);
        setIssueSummary(result.issueSummary);
      })
      .catch(() => {
        if (cancelled) return;
        setReviewSummary(null);
        setIssueSummary(null);
      });

    return () => { cancelled = true; };
  }, [hasSelection, selectedMonth, selectedOffice]);

  const handleSaveEdit = async () => {
    if (!aggregate) return;
    try {
      setSaveMessage('');
      const officeGoalsToSave = withOfficeSummaryTotalGoals(
        syncMonthlyGoalsFromDaily(
          officeSummaryGoalsDraft,
          OFFICE_SUMMARY_SAVABLE_GOAL_KEYS,
          aggregate.docCount
        ),
        aggregate.docCount
      );
      const oeGoalsToSave = syncMonthlyGoalsFromDaily(
        oeTableGoalsDraft,
        OE_TABLE_GOAL_KEYS,
        aggregate.docCount
      );
      const expensesToSave = withExpenseTotal(expensesDraft);
      const expensesPayload = buildExpensesForSave(expensesDraft, expensesSaved);
      const waitTimesPayload = buildWaitTimesForSave(waitTimesDraft, waitTimesSaved);
      const promotionActualsPayload = buildPromotionActualsForSave(promotionActualsDraft, promotionActualsSaved);
      const docRef = doc(db, 'monthly production', getMonthlyProductionDocId(selectedMonth, selectedOffice));
      const existingSnap = await getDoc(docRef);
      const existing = existingSnap.exists() ? (existingSnap.data() as Record<string, unknown>) : {};

      await setDoc(
        docRef,
        {
          month: normalizeDocMonth(selectedMonth) || selectedMonth,
          location: getMonthlyLocationFieldKey(selectedOffice),
          locationName: selectedOffice,
          updatedAt: serverTimestamp(),
          ...(existing.daysInOfficeOverrides !== undefined ? { daysInOfficeOverrides: existing.daysInOfficeOverrides } : {}),
          officeSummaryGoals: officeGoalsToSave,
          oeTableGoals: oeGoalsToSave,
          started: startedDraft === undefined ? deleteField() : startedDraft,
          clean: hasClean(cleanDraft) ? cleanDraft : deleteField(),
          waitTimes: waitTimesPayload,
          expenses: expensesPayload,
          promotionActuals: promotionActualsPayload,
          'summary note': hasNote(noteDraft) ? noteDraft!.trim() : deleteField(),
          note: deleteField(),
        },
        { merge: true }
      );
      setOfficeSummaryGoalsSaved(officeGoalsToSave);
      setOfficeSummaryGoalsDraft(officeGoalsToSave);
      setOeTableGoalsSaved(oeGoalsToSave);
      setOeTableGoalsDraft(oeGoalsToSave);
      setStartedSaved(startedDraft);
      setCleanSaved(cleanDraft);
      setWaitTimesSaved(waitTimesDraft);
      setExpensesSaved(expensesToSave);
      setExpensesDraft(expensesToSave);
      setPromotionActualsSaved(promotionActualsDraft);
      const savedNote = hasNote(noteDraft) ? noteDraft!.trim() : undefined;
      setNoteSaved(savedNote);
      setNoteDraft(savedNote);
      setIsEditing(false);
      setSaveMessage('Saved!');
    } catch (e: any) {
      setSaveMessage(`Failed to save: ${e?.message || 'Error'}`);
    }
  };

  return (
    <main style={{ minHeight: '100vh', background: '#f1f5f9', padding: '20px 24px', overflowX: 'hidden', boxSizing: 'border-box' }}>
      <section style={{ width: '100%', maxWidth: 1400, margin: '0 auto', boxSizing: 'border-box' }}>
        <div
          style={{
            display: 'flex',
            flexWrap: 'wrap',
            justifyContent: 'space-between',
            alignItems: 'center',
            gap: 16,
            marginBottom: 20,
            padding: '18px 22px',
            background: '#fff',
            border: '1px solid #e5e7eb',
            borderRadius: 12,
            boxShadow: '0 1px 3px rgba(15, 23, 42, 0.06)',
          }}
        >
          <div>
            <div style={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 10 }}>
              <h2 style={{ margin: 0, fontSize: 24, fontWeight: 700, color: '#0f172a' }}>
                Monthly Summary
              </h2>
              {aggregate && (
                <span
                  style={{
                    padding: '4px 12px',
                    borderRadius: 999,
                    background: '#eff6ff',
                    color: '#2563eb',
                    fontSize: 13,
                    fontWeight: 600,
                  }}
                >
                  {aggregate.docCount} days
                </span>
              )}
            </div>
            {hasSelection && (
              <p style={{ margin: '8px 0 0', color: '#64748b', fontSize: 14 }}>
                {selectedMonth} · {selectedOffice}
              </p>
            )}
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            {isEditing ? (
              <>
                <button
                  type="button"
                  onClick={handleSaveEdit}
                  style={{ ...editButtonStyle, background: '#16a34a', color: '#fff' }}
                >
                  Save
                </button>
                <button
                  type="button"
                  onClick={() => {
                    syncDraftsFromSaved();
                    setIsEditing(false);
                  }}
                  style={{ ...editButtonStyle, background: '#fff', color: '#64748b', border: '1px solid #e2e8f0' }}
                >
                  Cancel
                </button>
              </>
            ) : (
              <button
                type="button"
                onClick={() => {
                  syncDraftsFromSaved();
                  setIsEditing(true);
                }}
                disabled={!aggregate}
                style={{
                  ...editButtonStyle,
                  background: !aggregate ? '#e2e8f0' : '#2563eb',
                  color: !aggregate ? '#94a3b8' : '#fff',
                  cursor: !aggregate ? 'not-allowed' : 'pointer',
                }}
              >
                Edit
              </button>
            )}
          </div>
        </div>

        {saveMessage && (
          <StatusMessage type={saveMessage.includes('실패') ? 'error' : 'success'}>
            {saveMessage}
          </StatusMessage>
        )}
        {!hasSelection && <StatusMessage type="info">Need Month and Office.</StatusMessage>}
        {hasSelection && loading && <StatusMessage type="info">Loading...</StatusMessage>}
        {hasSelection && error && <StatusMessage type="error">{error}</StatusMessage>}
        {hasSelection && !loading && !error && !aggregate && (
          <StatusMessage type="info">There is no data for the month/office.</StatusMessage>
        )}
        {hasSelection && !loading && !error && aggregate && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: pageSectionGap, width: '100%' }}>
            <SectionBlock>
              <TableCard title="Note">
                <div style={{ padding: '12px 14px' }}>
                  <MonthlyNoteSection
                    value={activeNote}
                    isEditing={isEditing}
                    onChange={setNoteDraft}
                  />
                </div>
              </TableCard>
            </SectionBlock>

            <SectionBlock>
              <GoalsTablesSection
              aggregate={aggregate}
              oeTotalAmount={oeTotalAmount}
              officeMonthlyGoals={activeOfficeSummaryGoals.monthlyGoal}
              officeDailyGoals={activeOfficeSummaryGoals.dailyGoal}
              oeMonthlyGoals={activeOeTableGoals.monthlyGoal}
              oeDailyGoals={activeOeTableGoals.dailyGoal}
              isEditing={isEditing}
              onOfficeGoalChange={(key, rawValue) => {
                if (key === 'mailedProduction') return;
                setOfficeSummaryGoalsDraft((prev) =>
                  withOfficeSummaryTotalGoals(
                    updateDailyGoalSection(prev, key, rawValue, aggregate.docCount),
                    aggregate.docCount
                  )
                );
              }}
              onOeGoalChange={(key, rawValue) => {
                setOeTableGoalsDraft((prev) =>
                  updateDailyGoalSection(prev, key, rawValue, aggregate.docCount)
                );
              }}
              />
            </SectionBlock>

            <SectionBlock>
              <div style={tableGridStyle}>
                <div style={tableCardStyle}>
                  <TableCard>
                    <CraSummaryTable aggregate={aggregate} />
                  </TableCard>
                </div>
                <div style={tableCardStyle}>
                  <TableCard>
                    <CraNonCraTable aggregate={aggregate} oeTotalAmount={oeTotalAmount} />
                  </TableCard>
                </div>
                <div style={tableCardStyle}>
                  <TableCard>
                    <ReferralStartedTable
                      referred={aggregate.visitsReferral}
                      started={activeStarted}
                      isEditing={isEditing}
                      onStartedChange={(rawValue) => {
                        const trimmed = rawValue.trim();
                        if (trimmed === '') {
                          setStartedDraft(undefined);
                          return;
                        }
                        const parsed = parseNumber(trimmed);
                        if (Number.isFinite(parsed) && parsed >= 0) {
                          setStartedDraft(Math.round((parsed + Number.EPSILON) * 100) / 100);
                        }
                      }}
                    />
                  </TableCard>
                </div>
                <div style={tableCardStyle}>
                  <TableCard>
                    <SeenTable
                      officeSeen={aggregate.visitsSeen}
                      doctorSeen={aggregate.doctorCustomerTotal}
                    />
                  </TableCard>
                </div>
              </div>
            </SectionBlock>

            <SectionBlock>
              <div style={tableGridStyle}>
                <div style={tableCardStyle}>
                  <TableCard>
                    <VisitsTable aggregate={aggregate} />
                  </TableCard>
                </div>
                <div style={tableCardStyle}>
                  <TableCard>
                    <PromotionTable
                      docCount={aggregate.docCount}
                      actuals={activePromotionActuals}
                      isEditing={isEditing}
                      onActualChange={(key, rawValue) => {
                        setPromotionActualsDraft((prev) => updatePromotionActual(prev, key, rawValue));
                      }}
                    />
                  </TableCard>
                </div>
                <div style={tableCardStyle}>
                  <TableCard>
                    <CleanTable
                      value={activeClean}
                      isEditing={isEditing}
                      onChange={(rawValue) => {
                        const trimmed = rawValue.trim();
                        setCleanDraft(trimmed === '' ? undefined : rawValue);
                      }}
                    />
                  </TableCard>
                </div>
              </div>
            </SectionBlock>

            <SectionBlock>
              <div style={summaryTablesRowStyle}>
                <div style={tableCardStyle}>
                  <TableCard title="Review">
                    <ReviewSummaryTable summary={reviewSummary} />
                  </TableCard>
                </div>
                <div style={tableCardStyle}>
                  <TableCard title="Injury Issue">
                    <IssueSummaryTable summary={issueSummary} />
                  </TableCard>
                </div>
                <div style={tableCardStyle}>
                  <TableCard title="Wait Times">
                    <WaitTimesTable
                      values={activeWaitTimes}
                      isEditing={isEditing}
                      onChange={(key, rawValue) => {
                        setWaitTimesDraft((prev) => updateWaitTime(prev, key, rawValue));
                      }}
                    />
                  </TableCard>
                </div>
                <div style={tableCardStyle}>
                  <TableCard title="Expenses">
                    <ExpensesTable
                      values={activeExpenses}
                      isEditing={isEditing}
                      onChange={(key, rawValue) => {
                        setExpensesDraft((prev) => updateExpense(prev, key, rawValue));
                      }}
                      onLabelChange={(key, rawValue) => {
                        setExpensesDraft((prev) => updateExpenseLabel(prev, key, rawValue));
                      }}
                    />
                  </TableCard>
                </div>
              </div>
            </SectionBlock>

            <SectionBlock>
              <TableCard>
                <BonusDeductTable
                  reviewSummary={reviewSummary}
                  issueSummary={issueSummary}
                  oGoalTotal={oGoalTotalSaved}
                  sealantGoalTotal={sealantGoalTotalSaved}
                  sealantRda={getGoalsActualMetrics(aggregate, oeTotalAmount).sealantRda}
                  clean={cleanSaved}
                  started={startedSaved}
                />
              </TableCard>
            </SectionBlock>

            <SectionBlock>
              <TableCard>
                <DailyOfficeLogTable rows={dailyOfficeLogRows} />
              </TableCard>
            </SectionBlock>
          </div>
        )}
      </section>
    </main>
  );
}

export default function MonthlySummaryPage() {
  return (
    <Suspense
      fallback={
        <main style={{ minHeight: '100vh', background: '#f1f5f9', padding: 24, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
          <p style={{ margin: 0, color: '#64748b', fontSize: 15 }}>Loading...</p>
        </main>
      }
    >
      <MonthlySummaryPageContent />
    </Suspense>
  );
}