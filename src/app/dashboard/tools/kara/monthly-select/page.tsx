'use client'

import React, { useEffect, useMemo, useState } from 'react';
import { collection, doc, getDoc, getDocs } from 'firebase/firestore';
import { auth, db } from '@/lib/firebase.config';
import { onAuthStateChanged } from 'firebase/auth';

function yearMonthFromDate(date: string): string {
  const matched = /^(\d{4})-(\d{2})/.exec(date || '');
  return matched ? `${matched[1]}-${matched[2]}` : '';
}

type PatientItem = {
  office: string;
  source: string;
  reason: string;
};

type ShowDoc = {
  id: string;
  yearMonth: string;
  patients: PatientItem[];
};

function normalizePatient(patients: unknown): PatientItem[] {
  if (!Array.isArray(patients)) return [];

  const items: PatientItem[] = [];
  for (const item of patients) {
    if (!item || typeof item !== 'object') continue;
    const office = typeof item.office === 'string' ? item.office : '';
    const source = typeof item.source === 'string' ? item.source : '';
    const reason = typeof item.reason === 'string' ? item.reason : '';
    if (!office && !source && !reason) continue;
    items.push({ office, source, reason });
  }
  return items;
}

export default function Page() {
  const [pageReady, setPageReady] = useState(false);
  const [month, setMonth] = useState('');
  const [office, setOffice] = useState('');
  const [showDocs, setShowDocs] = useState<ShowDoc[]>([]);

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
    let cancelled = false;

    const loadShow = async () => {
      try {
        const snap = await getDocs(collection(db, 'show-noshow'));
        const docs: ShowDoc[] = [];

        for (const item of snap.docs) {
          const yearMonth = yearMonthFromDate(item.id);
          if (!yearMonth) continue;
          docs.push({
            id: item.id,
            yearMonth,
            patients: normalizePatient(item.data()?.patients),
          });
        }

        if (!cancelled) setShowDocs(docs);
      } catch {
        if (!cancelled) setShowDocs([]);
      }
    };

    loadShow();
    return () => {
      cancelled = true;
    };
  }, []);

  const availableMonths = useMemo(() => {
    const months = new Set<string>();
    for (const doc of showDocs) months.add(doc.yearMonth);
    return Array.from(months).sort((a, b) => b.localeCompare(a));
  }, [showDocs]);

  const officeOptions = useMemo(() => {
    const offices = new Set<string>();
    for (const doc of showDocs) {
      if (doc.yearMonth !== month) continue;
      for (const item of doc.patients) {
        if (item.office) offices.add(item.office);
      }
    }
    return Array.from(offices).sort((a, b) => a.localeCompare(b));
  }, [showDocs, month]);

  const sourceRows = useMemo(() => {
    if (!month || !office) return [];

    const counts = new Map<string, number>();
    for (const doc of showDocs) {
      if (doc.yearMonth !== month) continue;
      for (const item of doc.patients) {
        if (item.office !== office || !item.source) continue;
        counts.set(item.source, (counts.get(item.source) || 0) + 1);
      }
    }

    const total = Array.from(counts.values()).reduce((sum, n) => sum + n, 0);
    return Array.from(counts.entries())
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
      .map(([source, count]) => ({
        source,
        count,
        percentage: total > 0 ? (count / total) * 100 : 0,
      }));
  }, [showDocs, month, office]);

  const reasonRows = useMemo(() => {
    if (!month || !office) return [];

    const counts = new Map<string, number>();
    for (const doc of showDocs) {
      if (doc.yearMonth !== month) continue;
      for (const item of doc.patients) {
        if (item.office !== office || !item.reason) continue;
        counts.set(item.reason, (counts.get(item.reason) || 0) + 1);
      }
    }

    const total = Array.from(counts.values()).reduce((sum, n) => sum + n, 0);
    return Array.from(counts.entries())
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
      .map(([reason, count]) => ({
        reason,
        count,
        percentage: total > 0 ? (count / total) * 100 : 0,
      }));
  }, [showDocs, month, office]);

  useEffect(() => {
    if (availableMonths.length === 0) {
      if (month) setMonth('');
      return;
    }
    if (!availableMonths.includes(month)) {
      setMonth(availableMonths[0]);
    }
  }, [availableMonths, month]);

  useEffect(() => {
    if (office && !officeOptions.includes(office)) {
      setOffice('');
    }
  }, [office, officeOptions]);

  const pageStyle: React.CSSProperties = {
    minHeight: '100vh',
    margin: 0,
    padding: '48px 24px',
    background: '#f1f5f9',
    fontFamily:
      'ui-sans-serif, system-ui, -apple-system, "Segoe UI", sans-serif',
    color: '#0f172a',
  };

  const cardStyle: React.CSSProperties = {
    maxWidth: 1300,
    margin: '0 auto',
    background: '#fff',
    borderRadius: 12,
    boxShadow: '0 10px 30px rgba(15, 23, 42, 0.08)',
    overflow: 'hidden',
  };

  const filtersStyle: React.CSSProperties = {
    display: 'flex',
    gap: 20,
    flexWrap: 'wrap',
    padding: '20px 24px',
    borderBottom: '1px solid #e2e8f0',
  };

  const fieldStyle: React.CSSProperties = {
    flex: 1,
    minWidth: 200,
  };

  const labelStyle: React.CSSProperties = {
    display: 'block',
    marginBottom: 5,
    fontWeight: 'bold',
  };

  const inputStyle: React.CSSProperties = {
    padding: '8px 10px',
    border: '1px solid #e6e8eb',
    borderRadius: 4,
    fontSize: '1em',
    backgroundColor: '#ffffff',
    color: '#3b4252',
    width: '100%',
    boxSizing: 'border-box',
  };

  const tableStyle: React.CSSProperties = {
    width: '100%',
    borderCollapse: 'collapse',
  };

  const thStyle: React.CSSProperties = {
    padding: '12px 24px',
    fontSize: 13,
    fontWeight: 600,
    letterSpacing: '0.02em',
    color: '#475569',
    background: '#f8fafc',
    borderBottom: '1px solid #e2e8f0',
  };

  const tdStyle: React.CSSProperties = {
    padding: '12px 24px',
    fontSize: 14,
    height: 44,
    borderBottom: '1px solid #e2e8f0',
  };

  if (!pageReady) {
    return (
      <main style={{ minHeight: '100vh', background: '#f1f5f9' }} />
    );
  }

  return (
    <main style={pageStyle}>
      <h1
          style={{
            color: '#4b5563',
            textAlign: 'center',
            marginBottom: '24px',
            fontSize: '2.2rem',
            fontWeight: 'bold',
          }}
        >
          Appointment Details
      </h1>
      <div style={cardStyle}>
        <div style={filtersStyle}>
          <div style={fieldStyle}>
            <label style={labelStyle}>Month:</label>
            <select
              value={month}
              onChange={(e) => setMonth(e.target.value)}
              style={inputStyle}
            >
              {availableMonths.length === 0 ? (
                <option value="">No months</option>
              ) : (
                availableMonths.map((value) => (
                  <option key={value} value={value}>
                    {value}
                  </option>
                ))
              )}
            </select>
          </div>
          <div style={fieldStyle}>
            <label style={labelStyle}>Office:</label>
            <select
              value={office}
              onChange={(e) => setOffice(e.target.value)}
              style={inputStyle}
            >
              <option value="">Select Office</option>
              {officeOptions.map((option) => (
                <option key={option} value={option}>
                  {option}
                </option>
              ))}
            </select>
          </div>
        </div>

        <table style={tableStyle}>
          <thead>
            <tr>
              <th style={{ ...thStyle, textAlign: 'center' }}>Discovery Source</th>
              <th style={{ ...thStyle, textAlign: 'center' }}>Count</th>
              <th style={{ ...thStyle, textAlign: 'center' }}>Percentage</th>
            </tr>
          </thead>
          <tbody>
            {sourceRows.length === 0
              ? Array.from({ length: 10 }, (_, index) => (
                  <tr
                    key={index}
                    style={{
                      background: index % 2 === 0 ? '#fff' : '#f8fafc',
                    }}
                  >
                    <td style={{ ...tdStyle, textAlign: 'center' }} />
                    <td style={{ ...tdStyle, textAlign: 'center' }} />
                    <td style={{ ...tdStyle, textAlign: 'center' }} />
                  </tr>
                ))
              : sourceRows.map((row, index) => (
                  <tr
                    key={row.source}
                    style={{
                      background: index % 2 === 0 ? '#fff' : '#f8fafc',
                    }}
                  >
                    <td style={{ ...tdStyle, textAlign: 'center', fontWeight: 500 }}>
                      {row.source}
                    </td>
                    <td style={{ ...tdStyle, textAlign: 'center' }}>
                      {row.count.toLocaleString()}
                    </td>
                    <td style={{ ...tdStyle, textAlign: 'center' }}>
                      {row.percentage.toFixed(1)}%
                    </td>
                  </tr>
                ))}
          </tbody>
        </table>

        <div style={{ height: 125 }} />

        <table style={tableStyle}>
          <thead>
            <tr>
              <th style={{ ...thStyle, textAlign: 'center' }}>Cancellation / Reschedule Reason</th>
              <th style={{ ...thStyle, textAlign: 'center' }}>Count</th>
              <th style={{ ...thStyle, textAlign: 'center' }}>Percentage</th>
            </tr>
          </thead>
          <tbody>
            {reasonRows.length === 0
              ? Array.from({ length: 10 }, (_, index) => (
                  <tr
                    key={index}
                    style={{
                      background: index % 2 === 0 ? '#fff' : '#f8fafc',
                    }}
                  >
                    <td style={{ ...tdStyle, textAlign: 'center' }} />
                    <td style={{ ...tdStyle, textAlign: 'center' }} />
                    <td style={{ ...tdStyle, textAlign: 'center' }} />
                  </tr>
                ))
              : reasonRows.map((row, index) => (
                  <tr
                    key={row.reason}
                    style={{
                      background: index % 2 === 0 ? '#fff' : '#f8fafc',
                    }}
                  >
                    <td style={{ ...tdStyle, textAlign: 'center', fontWeight: 500 }}>
                      {row.reason}
                    </td>
                    <td style={{ ...tdStyle, textAlign: 'center' }}>
                      {row.count.toLocaleString()}
                    </td>
                    <td style={{ ...tdStyle, textAlign: 'center' }}>
                      {row.percentage.toFixed(1)}%
                    </td>
                  </tr>
                ))}
          </tbody>
        </table>
      </div>
    </main>
  );
}

