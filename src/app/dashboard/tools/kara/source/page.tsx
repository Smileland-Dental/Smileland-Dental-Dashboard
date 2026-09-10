'use client'

import React, { useEffect, useMemo, useState } from 'react';
import { collection, doc, getDoc, getDocs } from 'firebase/firestore';
import { auth, db } from '@/lib/firebase.config';
import { onAuthStateChanged } from 'firebase/auth';

function yearMonthFromDate(date: string): string {
  const matched = /^(\d{4})-(\d{2})/.exec(date || '');
  return matched ? `${matched[1]}-${matched[2]}` : '';
}

type SourceItem = {
  office: string;
  source: string;
};

type UnscheduledItem = {
  office: string;
  reason: string;
  type_of_visit: string;
  unscheduled: string;
};

type ShowDoc = {
  id: string;
  yearMonth: string;
  source: SourceItem[];
  unscheduled: UnscheduledItem[];
};

function normalizeSource(source: unknown): SourceItem[] {
  if (!Array.isArray(source)) return [];

  const items: SourceItem[] = [];
  for (const item of source) {
    if (!item || typeof item !== 'object') continue;
    const office = typeof item.office === 'string' ? item.office : '';
    const source = typeof item.source === 'string' ? item.source : '';
    if (!office && !source) continue;
    items.push({ office, source });
  }
  return items;
}

function normalizeUnscheduled(unscheduled: unknown): UnscheduledItem[] {
  if (!Array.isArray(unscheduled)) return [];

  const items: UnscheduledItem[] = [];
  for (const item of unscheduled) {
    if (!item || typeof item !== 'object') continue;
    const office = typeof item.office === 'string' ? item.office : '';
    const reason = typeof item.reason === 'string' ? item.reason : '';
    const type_of_visit = typeof item.type_of_visit === 'string' ? item.type_of_visit : '';
    const unscheduled = typeof item.unscheduled === 'string' ? item.unscheduled : '';
    if (!office && !reason && !type_of_visit && !unscheduled) continue;
    items.push({ office, reason, type_of_visit, unscheduled });
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
        const snap = await getDocs(collection(db, 'patientlog-details'));
        const docs: ShowDoc[] = [];

        for (const item of snap.docs) {
          const yearMonth = yearMonthFromDate(item.id);
          if (!yearMonth) continue;
          docs.push({
            id: item.id,
            yearMonth,
            source: normalizeSource(item.data()?.source),
            unscheduled: normalizeUnscheduled (item.data()?.unscheduled),
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
      for (const item of doc.source) {
        if (item.office) offices.add(item.office);
      }
      for (const item of doc.unscheduled) {
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
      for (const item of doc.source) {
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

  const counts = new Map<
    string,
    { unscheduled: string; visit_type: string; count: number }
  >();

  for (const doc of showDocs) {
    if (doc.yearMonth !== month) continue;

    for (const item of doc.unscheduled) {
      if (item.office !== office || !item.reason) continue;

      const existing = counts.get(item.reason);

      if (existing) {
        existing.count += 1;
      } else {
        counts.set(item.reason, {
          unscheduled: item.unscheduled,
          visit_type: item.type_of_visit,
          count: 1,
        });
      }
    }
  }

  const total = Array.from(counts.values()).reduce(
    (sum, item) => sum + item.count,
    0
  );
  const unscheduledOrder = [
  'No Show',
  'Cancelled',
  'Rescheduled',
  'Other',
];

  return Array.from(counts.entries())
  .sort((a, b) => {
    const aOrder = unscheduledOrder.indexOf(a[1].unscheduled);
    const bOrder = unscheduledOrder.indexOf(b[1].unscheduled);

    const safeAOrder = aOrder === -1 ? 999 : aOrder;
    const safeBOrder = bOrder === -1 ? 999 : bOrder;

    return (
      safeAOrder - safeBOrder ||
      a[1].unscheduled.localeCompare(b[1].unscheduled)
    );
  })
  .map(([reason, data]) => ({
    unscheduled: data.unscheduled,
    visit_type: data.visit_type,
    reason,
    count: data.count,
    percentage: total > 0 ? (data.count / total) * 100 : 0,
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
              <th style={{ ...thStyle, textAlign: 'center' }}>Unscheduled</th>
              <th style={{ ...thStyle, textAlign: 'center' }}>Type of Visit</th>
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
                    <td style={{ ...tdStyle, textAlign: 'center' }} />
                    <td style={{ ...tdStyle, textAlign: 'center' }} />
                  </tr>
                ))
              : reasonRows.map((row, index) => (
                  <tr
                   key={`${row.unscheduled}-${row.visit_type}-${row.reason}`}
                    style={{
                      background: index % 2 === 0 ? '#fff' : '#f8fafc',
                    }}
                  >
                    <td style={{ ...tdStyle, textAlign: 'center', fontWeight: 500 }}>
                      {row.unscheduled}
                    </td>
                    <td style={{ ...tdStyle, textAlign: 'center' }}>
                      {row.visit_type}
                    </td>
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