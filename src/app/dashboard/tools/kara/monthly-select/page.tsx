'use client';

import React, { useEffect, useMemo, useRef, useState } from 'react';
import { auth, db } from '@/lib/firebase.config';
import { onAuthStateChanged } from 'firebase/auth';
import { collection, doc, getDoc, getDocs } from 'firebase/firestore';

const BLACK_BEAR_COLLECTION = 'simple-forms';

const colors = {
  pageBg: '#f8fafc',
  cardBg: '#ffffff',
  cardBorder: '#e5e7eb',
  title: '#111827',
  text: '#111827',
  placeholder: '#9ca3af',
  label: '#374151',
  accent: '#6b7280',
  accentSoft: '#f3f4f6',
  accentText: '#111827',
  inputBorder: '#d1d5db',
  reportBg: '#ffffff',
  reportBorder: '#e5e7eb',
  reportDisabledBg: '#f9fafb',
  reportDisabledText: '#9ca3af',
  hint: '#6b7280',
};

type SelectDestination = {
  label: string;
  path: string;
};

const SELECT_DESTINATIONS: SelectDestination[] = [
  { label: 'Monthly Summary', path: '/dashboard/tools/kara/monthly-summary' },
  { label: 'Monthly Production', path: '/dashboard/tools/kara/monthly' },
  { label: 'Sealant & OE Goal', path: '/dashboard/tools/kara/sealant-goal' },
  { label: 'Monthly Reviews', path: '/dashboard/tools/kara/review' },
];

type FormDoc = {
  id: string;
  date?: string;
  location?: string;
  submittedDateTime?: string;
};

function normalizeDocMonth(dateValue: unknown): string {
  const raw = String(dateValue ?? '').trim();
  const match = raw.match(/^(\d{4})-(\d{1,2})/);
  if (!match) return '';
  return `${match[1]}-${String(Number(match[2])).padStart(2, '0')}`;
}

function buildDestinationUrl(path: string, month: string, office: string): string {
  const params = new URLSearchParams({ month, office });
  return `${path}?${params.toString()}`;
}

function getUniqueSorted(values: string[], direction: 'asc' | 'desc'): string[] {
  const sorted = Array.from(new Set(values));
  sorted.sort((a, b) => (direction === 'asc' ? a.localeCompare(b) : b.localeCompare(a)));
  return sorted;
}

function getMonthOptions(docs: FormDoc[]): string[] {
  const months = docs
    .filter((doc) => !!doc.submittedDateTime)
    .map((doc) => normalizeDocMonth(doc.date))
    .filter(Boolean);
  return getUniqueSorted(months, 'desc');
}

function normalizeAllowedOffices(officeValue: unknown): string[] {
  if (Array.isArray(officeValue)) {
    return officeValue.map((value) => String(value ?? '').trim()).filter(Boolean);
  }
  if (typeof officeValue === 'string') {
    return officeValue
      .split(',')
      .map((value) => value.trim())
      .filter(Boolean);
  }
  if (officeValue && typeof officeValue === 'object') {
    return Object.values(officeValue as Record<string, unknown>)
      .map((value) => String(value ?? '').trim())
      .filter(Boolean);
  }
  return [];
}

function getOfficeOptions(docs: FormDoc[], month: string, allowedOffices: string[]): string[] {
  if (!month || allowedOffices.length === 0) return [];
  const allowedSet = new Set(allowedOffices);
  const locations = docs
    .filter((doc) => !!doc.submittedDateTime && normalizeDocMonth(doc.date) === month)
    .map((doc) => String(doc.location ?? '').trim())
    .filter((location) => location && allowedSet.has(location));
  return getUniqueSorted(locations, 'asc');
}

const dropdownOptionStyle = (active: boolean): React.CSSProperties => ({
  width: '100%',
  border: 0,
  borderRadius: 8,
  padding: '10px 12px',
  background: active ? colors.accentSoft : 'transparent',
  color: active ? colors.accentText : colors.text,
  textAlign: 'left',
  cursor: 'pointer',
  fontWeight: active ? 600 : 400,
});

const reportButtonStyle = (enabled: boolean): React.CSSProperties => ({
  width: '100%',
  border: `1px solid ${enabled ? colors.reportBorder : colors.cardBorder}`,
  borderRadius: 10,
  padding: '12px 14px',
  background: enabled ? colors.reportBg : colors.reportDisabledBg,
  textAlign: 'left',
  cursor: enabled ? 'pointer' : 'not-allowed',
});

type DropdownFieldProps = {
  label: string;
  value: string;
  isOpen: boolean;
  onToggle: () => void;
  canOpen: boolean;
  dropdownRef: React.RefObject<HTMLDivElement | null>;
  children: React.ReactNode;
};

function DropdownField({
  label,
  value,
  isOpen,
  onToggle,
  canOpen,
  dropdownRef,
  children,
}: DropdownFieldProps) {
  return (
    <div ref={dropdownRef} style={{ position: 'relative' }}>
      <label style={{ display: 'block', marginBottom: 6, fontWeight: 600, fontSize: 13, color: colors.label }}>
        {label}
      </label>
      <button
        type="button"
        onClick={() => canOpen && onToggle()}
        style={{
          width: '100%',
          height: 40,
          border: `1px solid ${isOpen ? colors.accent : colors.inputBorder}`,
          borderRadius: 10,
          padding: '0 12px',
          background: colors.cardBg,
          color: value ? colors.text : colors.placeholder,
          textAlign: 'left',
          cursor: canOpen ? 'pointer' : 'default',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          fontWeight: value ? 600 : 400,
          fontSize: 14,
        }}
      >
        <span>{value || `Select ${label.toLowerCase()}`}</span>
        {canOpen && (
          <span style={{ color: colors.accent, fontSize: 11 }}>{isOpen ? '▲' : '▼'}</span>
        )}
      </button>
      {isOpen && children}
    </div>
  );
}

function DropdownList({ children }: { children: React.ReactNode }) {
  return (
    <ul
      style={{
        position: 'absolute',
        top: 'calc(100% + 4px)',
        left: 0,
        right: 0,
        margin: 0,
        padding: 4,
        listStyle: 'none',
        border: `1px solid ${colors.cardBorder}`,
        borderRadius: 10,
        background: colors.cardBg,
        boxShadow: '0 4px 16px rgba(15, 23, 42, 0.06)',
        maxHeight: 220,
        overflowY: 'auto',
        zIndex: 10,
      }}
    >
      {children}
    </ul>
  );
}

function DropdownOption({
  active,
  onClick,
  children,
}: {
  active: boolean;
  onClick: () => void;
  children: React.ReactNode;
}) {
  return (
    <li>
      <button type="button" onClick={onClick} style={dropdownOptionStyle(active)}>
        {children}
      </button>
    </li>
  );
}

export default function MonthlySelectPage() {
  const [pageReady, setPageReady] = useState(false);
  const [docs, setDocs] = useState<FormDoc[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [selectedMonth, setSelectedMonth] = useState('');
  const [selectedOffice, setSelectedOffice] = useState('');
  const [allowedOffices, setAllowedOffices] = useState<string[]>([]);
  const [monthDropdownOpen, setMonthDropdownOpen] = useState(false);
  const [officeDropdownOpen, setOfficeDropdownOpen] = useState(false);
  const monthDropdownRef = useRef<HTMLDivElement>(null);
  const officeDropdownRef = useRef<HTMLDivElement>(null);

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
          setAllowedOffices(normalizeAllowedOffices(userData?.offices));
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
    const load = async () => {
      try {
        const snap = await getDocs(collection(db, BLACK_BEAR_COLLECTION));
        const loaded = snap.docs.map((d) => ({ id: d.id, ...(d.data() as Omit<FormDoc, 'id'>) }));
        setDocs(loaded);
      } catch (e: unknown) {
        const message = e instanceof Error ? e.message : 'Error';
        setError(message);
      } finally {
        setLoading(false);
      }
    };
    load();
  }, []);

  const monthOptions = useMemo(() => getMonthOptions(docs), [docs]);
  const officeOptions = useMemo(
    () => getOfficeOptions(docs, selectedMonth, allowedOffices),
    [docs, selectedMonth, allowedOffices]
  );

  useEffect(() => {
    if (selectedMonth && !monthOptions.includes(selectedMonth)) {
      setSelectedMonth('');
    }
    if (selectedOffice && !officeOptions.includes(selectedOffice)) {
      setSelectedOffice('');
    }
  }, [monthOptions, officeOptions, selectedMonth, selectedOffice]);

  useEffect(() => {
    if (!monthDropdownOpen && !officeDropdownOpen) return;

    const handleClickOutside = (event: MouseEvent) => {
      const target = event.target as Node;
      if (monthDropdownRef.current && !monthDropdownRef.current.contains(target)) {
        setMonthDropdownOpen(false);
      }
      if (officeDropdownRef.current && !officeDropdownRef.current.contains(target)) {
        setOfficeDropdownOpen(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [monthDropdownOpen, officeDropdownOpen]);

  const handleMonthSelect = (month: string) => {
    setSelectedMonth(month);
    setSelectedOffice('');
    setMonthDropdownOpen(false);
    setOfficeDropdownOpen(false);
  };

  const handleOfficeSelect = (office: string) => {
    setSelectedOffice(office);
    setOfficeDropdownOpen(false);
  };

  const canNavigate = selectedMonth !== '' && selectedOffice !== '';

  const handleDestinationClick = (path: string) => {
    if (!canNavigate) return;
    window.open(buildDestinationUrl(path, selectedMonth, selectedOffice), '_blank', 'noopener,noreferrer');
  };

  if (!pageReady) {
    return <main style={{ minHeight: '100vh', background: colors.pageBg }} />;
  }

  return (
    <main style={{ minHeight: '100vh', background: colors.pageBg, padding: 32 }}>
      <section
        style={{
          maxWidth: 520,
          margin: '32px auto',
          border: `1px solid ${colors.cardBorder}`,
          borderRadius: 14,
          padding: '28px 26px',
          background: colors.cardBg,
          boxShadow: '0 2px 12px rgba(15, 23, 42, 0.05)',
        }}
      >
        <h1 style={{ margin: '0 0 20px', fontSize: 24, fontWeight: 700, color: colors.title }}>
          Monthly Reports
        </h1>

        {loading && <p style={{ margin: 0, color: colors.hint }}>Loading...</p>}
        {error && <p style={{ margin: 0, color: '#ef4444' }}>{error}</p>}

        {!loading && !error && (
          <div style={{ display: 'grid', gap: 14 }}>
            <DropdownField
              label="Month"
              value={selectedMonth}
              isOpen={monthDropdownOpen}
              onToggle={() => setMonthDropdownOpen((prev) => !prev)}
              canOpen={monthOptions.length > 0}
              dropdownRef={monthDropdownRef}
            >
              <DropdownList>
                {monthOptions.map((month) => (
                  <DropdownOption
                    key={month}
                    active={month === selectedMonth}
                    onClick={() => handleMonthSelect(month)}
                  >
                    {month}
                  </DropdownOption>
                ))}
              </DropdownList>
            </DropdownField>

            <DropdownField
              label="Office"
              value={selectedOffice}
              isOpen={officeDropdownOpen}
              onToggle={() => setOfficeDropdownOpen((prev) => !prev)}
              canOpen={!!selectedMonth && officeOptions.length > 0}
              dropdownRef={officeDropdownRef}
            >
              <DropdownList>
                {officeOptions.map((office) => (
                  <DropdownOption
                    key={office}
                    active={office === selectedOffice}
                    onClick={() => handleOfficeSelect(office)}
                  >
                    {office}
                  </DropdownOption>
                ))}
              </DropdownList>
            </DropdownField>

            <div
              style={{
                marginTop: 4,
                paddingTop: 18,
                borderTop: `1px solid ${colors.cardBorder}`,
              }}
            >
              <p style={{ margin: '0 0 10px', fontWeight: 600, fontSize: 14, color: colors.text }}>
                Reports
              </p>
              {!canNavigate && (
                <p style={{ margin: '0 0 10px', fontSize: 13, color: colors.hint }}>
                  Please select Month and Office first.
                </p>
              )}
              <div style={{ display: 'grid', gap: 8 }}>
                {SELECT_DESTINATIONS.map((destination) => (
                  <button
                    key={destination.path}
                    type="button"
                    onClick={() => handleDestinationClick(destination.path)}
                    disabled={!canNavigate}
                    style={reportButtonStyle(canNavigate)}
                  >
                    <div
                      style={{
                        fontWeight: 600,
                        color: canNavigate ? colors.text : colors.reportDisabledText,
                      }}
                    >
                      {destination.label}
                    </div>
                  </button>
                ))}
              </div>
            </div>
          </div>
        )}
      </section>
    </main>
  );
}
