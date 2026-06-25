'use client';

import React, { useState, useEffect, useCallback, useRef } from 'react';
import { db, auth } from '@/lib/firebase.config';
import { doc, setDoc, getDoc, Timestamp, deleteDoc } from 'firebase/firestore';
import { getStorage, ref, uploadBytes } from 'firebase/storage';
import { onAuthStateChanged } from 'firebase/auth';
import { pdf, Document, Page, View, Text, StyleSheet } from '@react-pdf/renderer';

interface StaffMember {
  no: number;
  name: string;
  active?: boolean;
}

interface StaffList {
  [position: string]: StaffMember[];
}

interface DoctorMember {
  no: number;
  name: string;
}

interface AttendanceRow {
  date: string;
  filledBy: string;
  checkedBy: string;
  position: string;
  count: number;
  no: number;
  name: string;
  present: boolean;
  startTardy?: string;
  lateLunch?: string;
  needsAdj?: boolean;
  overtime?: string;
  otCorp?: string;
  subAnother?: boolean;
  incident?: string;
  notes?: string;
}

interface DoctorRow {
  date: string;
  filledBy: string;
  checkedBy: string;
  position: string;
  count: number;
  no: number;
  name: string;
  present: boolean;
  checkIn?: string;
  lunchOut?: string;
  lunchIn?: string;
  checkOut?: string;
}

export default function AttendanceTrack() {
  const [staffList, setStaffList] = useState<{ staff: StaffList; doctors: DoctorMember[] } | null>(null);
  const [userOfficesOptions, setUserOfficesOptions] = useState<string[]>([]); // 사용자의 offices 옵션들

  // 마지막 저장된 데이터 추적 (자동 저장 최적화용)
  const lastSavedDataRef = useRef<string>('');
  // 초기 로드 완료 플래그 (초기 로드 시 자동 저장 방지)
  const isInitialLoadRef = useRef<boolean>(true);
  // 🔒 보안: Rate limiting을 위한 ref
  const lastAutoSaveTimeRef = useRef<number>(0);
  const lastApiCallTimeRef = useRef<number>(0);
  const autoSaveCountRef = useRef<number>(0);
  const autoSaveResetTimeRef = useRef<number>(Date.now());
  const AUTO_SAVE_DEBOUNCE_MS = 1000;
  const AUTO_SAVE_MIN_INTERVAL_MS = 1000;

  const theme = {
    bg: '#FFFFFF',
    card: '#FFFFFF',
    cardAlt: '#F8F9FA',
    gray: '#5F6B73',
    accent: '#4A8F65',
    text: '#212529',
    textMuted: '#6C757D',
    border: '#DEE2E6',
    borderLight: '#E9ECEF',
    positionBg: '#F1F3F5',
    shadow: '0 1px 4px rgba(0, 0, 0, 0.06)',
    radiusSm: '8px',
    radiusLg: '12px',
    font: "'Segoe UI', system-ui, -apple-system, BlinkMacSystemFont, sans-serif",
  };

  const ui = {
    page: {
      minHeight: '100vh',
      padding: '28px 24px',
      fontFamily: theme.font,
      background: theme.bg,
      color: theme.text,
    },
    title: {
      fontSize: '2.25rem',
      fontWeight: 600,
      color: '#000000',
      marginBottom: '8px',
      textAlign: 'center' as const,
      letterSpacing: '-0.02em',
    },
    subtitle: {
      marginTop: '4px',
      marginBottom: '28px',
      fontSize: '15px',
      textAlign: 'center' as const,
      color: theme.textMuted,
      lineHeight: 1.5,
    },
    formCard: {
      marginBottom: '24px',
      padding: '20px 24px',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      gap: '20px',
      flexWrap: 'wrap' as const,
      background: theme.cardAlt,
      borderRadius: theme.radiusLg,
      boxShadow: 'none',
      border: `1px solid ${theme.borderLight}`,
    },
    label: { display: 'flex', alignItems: 'center', gap: '8px', color: theme.text, fontWeight: 500 },
    input: {
      padding: '8px 12px',
      borderRadius: theme.radiusSm,
      border: `1px solid ${theme.border}`,
      background: theme.card,
      color: theme.text,
      fontSize: '14px',
    },
    tableCard: {
      background: theme.card,
      borderRadius: theme.radiusLg,
      overflow: 'hidden',
      boxShadow: theme.shadow,
      border: `1px solid ${theme.border}`,
    },
    table: {
      width: '100%',
      tableLayout: 'fixed' as const,
      borderCollapse: 'separate' as const,
      borderSpacing: 0,
      background: theme.card,
    },
    th: {
      padding: '12px 8px',
      borderBottom: `1px solid ${theme.border}`,
      borderRight: '1px solid rgba(255, 255, 255, 0.25)',
      fontSize: '0.95rem',
      fontWeight: 600,
      background: theme.gray,
      color: '#F5FBFB',
    },
    td: {
      padding: '10px 8px',
      borderBottom: `1px solid ${theme.borderLight}`,
      borderRight: `1px solid ${theme.borderLight}`,
      color: theme.text,
    },
    tdCenter: {
      padding: '10px 8px',
      borderBottom: `1px solid ${theme.borderLight}`,
      borderRight: `1px solid ${theme.borderLight}`,
      textAlign: 'center' as const,
      color: theme.text,
    },
    addBtn: {
      padding: '6px 14px',
      background: theme.accent,
      color: 'white',
      border: 'none',
      borderRadius: '20px',
      cursor: 'pointer',
      fontSize: '12px',
      fontWeight: 500,
      position: 'absolute' as const,
      right: '12px',
    },
    removeBtn: {
      padding: '0 6px',
      background: 'transparent',
      color: theme.textMuted,
      border: `1px solid ${theme.border}`,
      borderRadius: '4px',
      cursor: 'pointer',
      fontSize: '14px',
      lineHeight: 1.2,
    },
    submitBtn: {
      padding: '12px 32px',
      background: theme.gray,
      color: '#FFFFFF',
      border: 'none',
      borderRadius: '24px',
      cursor: 'pointer',
      fontSize: '16px',
      fontWeight: 600,
      boxShadow: '0 2px 6px rgba(0, 0, 0, 0.12)',
    },
    officeBadge: {
      padding: '8px 14px',
      borderRadius: theme.radiusSm,
      border: `1px solid ${theme.border}`,
      backgroundColor: theme.cardAlt,
      fontWeight: 500,
      color: theme.text,
    },
    checkbox: { width: '18px', height: '18px', accentColor: theme.gray },
    cellInput: { width: '100%', border: 'none', background: 'transparent', padding: '2px 0', color: theme.text },
  };

  // 🔒 보안: 입력 검증 및 XSS 방지 함수
  const validateInput = (value: string | undefined | null, maxLength: number = 500): string => {
    if (!value || typeof value !== 'string') return '';
    // XSS 방지: 위험한 문자 제거
    let sanitized = value
      .replace(/[<>\"']/g, '') // HTML 태그 및 따옴표 제거
      .replace(/javascript:/gi, '') // javascript: 프로토콜 제거
      .replace(/on\w+=/gi, ''); // 이벤트 핸들러 제거 (onclick, onerror 등)
    // 길이 제한
    if (sanitized.length > maxLength) {
      sanitized = sanitized.substring(0, maxLength);
    }
    return sanitized;
  };

  // 🔒 보안: 날짜 형식 검증
  const validateDate = (date: string): boolean => {
    const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
    if (!dateRegex.test(date)) return false;
    const parsedDate = new Date(date);
    return !isNaN(parsedDate.getTime());
  };

  const OFFICE_OPTIONS = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'] as const;
  const INCIDENT_OPTIONS = ['', 'Late In', 'Early Out', 'Long Lunch', 'Leave and Come Back', 'Voluntary Early Out'] as const;
  const POSITION_ORDER = ['Front Office', 'Biller', 'Dental Assistant', 'RDA', 'Sub', 'Extern'] as const;
  const CUSTOM_ROW_NO_THRESHOLD = 1000;

  // 🔒 보안: 오피스 값 검증
  const validateOffice = (office: string): boolean =>
    (OFFICE_OPTIONS as readonly string[]).includes(office);

  // 🔒 보안: Position 값 검증
  const validatePosition = (position: string): boolean =>
    (POSITION_ORDER as readonly string[]).includes(position);

  // 🔒 보안: Incident 값 검증
  const validateIncident = (incident: string): boolean =>
    (INCIDENT_OPTIONS as readonly string[]).includes(incident);

  const getAttendanceDocId = (date: string, office: string) =>
    `${date}_${office}`.replace(/[^a-zA-Z0-9_-]/g, '');

  // --- PDF 생성 관련 상수/스타일 ---
  const pdfStyles = StyleSheet.create({
    page: { padding: 20, fontFamily: 'Helvetica', fontSize: 7 },
    header: { marginBottom: 10, borderBottomWidth: 2, borderColor: '#333', paddingBottom: 6, alignItems: 'center' },
    headerTitle: { fontSize: 14, fontWeight: 'bold', marginBottom: 4 },
    headerSub: { fontSize: 8 },
    row: { flexDirection: 'row', borderBottomWidth: 0.5, borderColor: '#333' },
    cell: { padding: 2, fontSize: 6, flex: 1, borderRightWidth: 0.5, borderColor: '#333', justifyContent: 'center', alignItems: 'center' },
    cellBold: { fontWeight: 'bold', backgroundColor: '#e3f2fd' },
    cellGray: { backgroundColor: '#fce4ec', fontWeight: 'bold' },
    footer: { marginTop: 15, paddingTop: 8, borderTopWidth: 1, borderColor: '#ddd', alignItems: 'center', fontSize: 7, color: '#666' },
    dailyRecap: { textAlign: 'right', fontSize: 10, fontWeight: 'bold', marginTop: 4, marginBottom: 8 },
  });

  // PDF 생성 유틸 함수
  function safeStr(v: unknown, max: number): string {
    if (v == null) return '';
    return String(v).trim().slice(0, max).replace(/[<>]/g, '');
  }

  function isChecked(v: unknown): boolean {
    return v === true || v === 1 || (typeof v === 'string' && (v === 'true' || v === '1'));
  }

  // 시간 형식 변환 (HH:MM -> H:MM AM/PM)
  function formatTimeToAMPM(timeStr: string): string {
    if (!timeStr || typeof timeStr !== 'string' || timeStr.trim() === '') return '';
    if (timeStr.length > 10) return '';
    const timeMatch = timeStr.match(/^(\d{1,2}):(\d{2})$/);
    if (!timeMatch) return safeStr(timeStr, 10);
    let hour = parseInt(timeMatch[1], 10);
    const minute = timeMatch[2];
    if (isNaN(hour) || hour < 0 || hour > 23) return safeStr(timeStr, 10);
    if (!/^\d{2}$/.test(minute) || parseInt(minute, 10) < 0 || parseInt(minute, 10) > 59) return safeStr(timeStr, 10);
    const period = hour >= 12 ? 'PM' : 'AM';
    if (hour === 0) hour = 12;
    else if (hour > 12) hour = hour - 12;
    return `${hour}:${minute} ${period}`;
  }

  function createAttendancePDFDocument(props: {
    safeDate: string;
    safeOffice: string;
    safeFilledBy: string;
    safeCheckedBy: string;
    staffData: AttendanceRow[];
    doctorData: DoctorRow[];
    generatedDate: string;
  }) {
    const { safeDate, safeOffice, safeFilledBy, safeCheckedBy, staffData, doctorData, generatedDate } = props;
    const s = pdfStyles;

    // Position별 그룹핑
    const staffByPosition: { [key: string]: AttendanceRow[] } = {};
    staffData.forEach(row => {
      let pos = row.position || '';
      if (!validatePosition(pos) && pos !== 'Dental Assistant') {
        pos = '';
      }
      if (pos === 'Dental Assistant') pos = 'DA';
      if (!staffByPosition[pos]) staffByPosition[pos] = [];
      staffByPosition[pos].push(row);
    });

    // Staff 테이블 헤더
    const staffHeaderRow = React.createElement(View, { key: 'staff-header', style: [s.row, s.cellBold] },
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, 'Name')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, 'Present')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, 'Start Shift Tardy (Min)')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, 'Late from Lunch (Min)')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, 'Needs Clock Adj.')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, 'Overtime')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, 'OT Corp Authorized By')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, 'Sub. at Another Office')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, 'Incident Description')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, 'Notes')),
    );

    // Staff 테이블 데이터 행
    const staffRows: React.ReactElement[] = [];
    Object.keys(staffByPosition).forEach(pos => {
      const rows = staffByPosition[pos];
      const presentCount = rows.filter(r => r.present === true).length;
      
      // Position 구분 행
      staffRows.push(
        React.createElement(View, { key: `pos-${pos}`, style: [s.row, s.cellGray] },
          React.createElement(View, { style: s.cell }, React.createElement(Text, null, '')),
          React.createElement(View, { style: s.cell }, React.createElement(Text, null, '')),
          React.createElement(View, { style: s.cell }, React.createElement(Text, null, '')),
          React.createElement(View, { style: s.cell }, React.createElement(Text, null, '')),
          React.createElement(View, { style: s.cell }, React.createElement(Text, null, pos)),
          React.createElement(View, { style: s.cell }, React.createElement(Text, null, String(presentCount))),
          React.createElement(View, { style: s.cell }, React.createElement(Text, null, '')),
          React.createElement(View, { style: s.cell }, React.createElement(Text, null, '')),
          React.createElement(View, { style: s.cell }, React.createElement(Text, null, '')),
          React.createElement(View, { style: s.cell }, React.createElement(Text, null, '')),
        )
      );
      
      // 직원 행
      rows.forEach((row, idx) => {
        const safeName = safeStr(row.name, 100);
        if (safeName && safeName.trim()) {
          staffRows.push(
            React.createElement(View, { key: `staff-${pos}-${idx}`, style: s.row },
              React.createElement(View, { style: s.cell }, React.createElement(Text, null, safeName)),
              React.createElement(View, { style: s.cell }, React.createElement(Text, null, isChecked(row.present) ? 'O' : '')),
              React.createElement(View, { style: s.cell }, React.createElement(Text, null, safeStr(row.startTardy, 50))),
              React.createElement(View, { style: s.cell }, React.createElement(Text, null, safeStr(row.lateLunch, 50))),
              React.createElement(View, { style: s.cell }, React.createElement(Text, null, isChecked(row.needsAdj) ? 'O' : '')),
              React.createElement(View, { style: s.cell }, React.createElement(Text, null, safeStr(row.overtime, 50))),
              React.createElement(View, { style: s.cell }, React.createElement(Text, null, safeStr(row.otCorp, 50))),
              React.createElement(View, { style: s.cell }, React.createElement(Text, null, isChecked(row.subAnother) ? 'O' : '')),
              React.createElement(View, { style: s.cell }, React.createElement(Text, null, safeStr(row.incident, 50))),
              React.createElement(View, { style: s.cell }, React.createElement(Text, null, safeStr(row.notes, 500))),
            )
          );
        }
      });
    });

    // Doctor 테이블 헤더
    const doctorPresentCount = doctorData.filter(d => d.present === true).length;
    const doctorHeaderRow1 = React.createElement(View, { key: 'doctor-header-1', style: [s.row, s.cellBold] },
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, 'Name')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, 'Present')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, 'Check In')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, 'Lunch Out')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, 'Lunch In')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, 'Check Out')),
    );
    const doctorHeaderRow2 = React.createElement(View, { key: 'doctor-header-2', style: [s.row, s.cellGray] },
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, '')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, '')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, 'Doctor')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, String(doctorPresentCount))),
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, '')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, '')),
    );

    // Doctor 테이블 데이터 행
    const doctorRows = doctorData.map((row, idx) => {
      const safeName = safeStr(row.name, 100);
      if (!safeName || !safeName.trim()) return null;
      return React.createElement(View, { key: `doctor-${idx}`, style: s.row },
        React.createElement(View, { style: s.cell }, React.createElement(Text, null, safeName)),
        React.createElement(View, { style: s.cell }, React.createElement(Text, null, isChecked(row.present) ? 'O' : '')),
        React.createElement(View, { style: s.cell }, React.createElement(Text, null, formatTimeToAMPM(safeStr(row.checkIn, 10)))),
        React.createElement(View, { style: s.cell }, React.createElement(Text, null, formatTimeToAMPM(safeStr(row.lunchOut, 10)))),
        React.createElement(View, { style: s.cell }, React.createElement(Text, null, formatTimeToAMPM(safeStr(row.lunchIn, 10)))),
        React.createElement(View, { style: s.cell }, React.createElement(Text, null, formatTimeToAMPM(safeStr(row.checkOut, 10)))),
      );
    }).filter(Boolean) as React.ReactElement[];

    // Daily Recap 계산
    let totalStaffPresent = 0;
    Object.keys(staffByPosition).forEach(pos => {
      totalStaffPresent += staffByPosition[pos].filter(r => r.present === true).length;
    });
    const totalPresent = totalStaffPresent + doctorPresentCount;

    const header = React.createElement(View, { style: s.header },
      React.createElement(Text, { style: s.headerTitle }, `${safeOffice} Attendance Tract`),
      React.createElement(Text, { style: s.headerSub }, `Date: ${safeDate} | Filled Out By: ${safeFilledBy} | Management that Checked Times Today on Time Clock: ${safeCheckedBy}`),
    );

    const staffTable = React.createElement(View, { key: 'staff-table' }, staffHeaderRow, ...staffRows);
    const doctorTable = React.createElement(View, { key: 'doctor-table', style: { marginTop: 10 } }, doctorHeaderRow1, doctorHeaderRow2, ...doctorRows);
    const dailyRecap = React.createElement(View, { style: s.dailyRecap }, React.createElement(Text, null, `Daily Recap: ${totalPresent}`));
    const footer = React.createElement(View, { style: s.footer }, React.createElement(Text, null, `Generated: ${generatedDate}`));

    return React.createElement(Document, null,
      React.createElement(Page, { size: 'LETTER', orientation: 'landscape', style: s.page }, header, staffTable, doctorTable, dailyRecap, footer),
    );
  }

  // 날짜 상태 (캘리포니아 시간대)
  const [trackDate, setTrackDate] = useState(() => {
    const now = new Date();
    const formatter = new Intl.DateTimeFormat('en-US', {
      timeZone: 'America/Los_Angeles',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit'
    });
    const parts = formatter.formatToParts(now);
    const year = parts.find(p => p.type === 'year')?.value || '';
    const month = parts.find(p => p.type === 'month')?.value || '';
    const day = parts.find(p => p.type === 'day')?.value || '';
    return `${year}-${month}-${day}`;
  });

  const [selectedOffice, setSelectedOffice] = useState(''); // 오피스 선택
  const [filledBy, setFilledBy] = useState('');
  const [checkedBy, setCheckedBy] = useState('');

  const [tableRows, setTableRows] = useState<Array<{
    id: string;
    position: string;
    no: number;
    name: string;
    present: boolean;
    startTardy: string;
    lateLunch: string;
    needsAdj: boolean;
    overtime: string;
    otCorp: string;
    subAnother: boolean;
    incident: string;
    notes: string;
  }>>([]);

  // Doctor 테이블 데이터 상태
  const [doctorRows, setDoctorRows] = useState<Array<{
    id: string;
    no: number;
    name: string;
    present: boolean;
    checkIn: string;
    lunchOut: string;
    lunchIn: string;
    checkOut: string;
  }>>([]);

  const compareByName = (a: { name: string }, b: { name: string }) =>
    (a.name || '').localeCompare(b.name || '', 'en', { sensitivity: 'base' });

  const isCustomStaffRow = (row: { no: number }) => row.no >= CUSTOM_ROW_NO_THRESHOLD;

  const isCustomDoctorRow = (row: { no: number }) => row.no >= CUSTOM_ROW_NO_THRESHOLD;

  const isBlankCustomStaffRow = (row: typeof tableRows[number]) =>
    isCustomStaffRow(row) &&
    !(row.name || '').trim() &&
    !row.present &&
    !row.subAnother &&
    !(row.startTardy || '').trim() &&
    !(row.lateLunch || '').trim() &&
    !row.needsAdj &&
    !(row.overtime || '').trim() &&
    !(row.otCorp || '').trim() &&
    !(row.incident || '').trim() &&
    !(row.notes || '').trim();

  const isBlankCustomDoctorRow = (row: typeof doctorRows[number]) =>
    isCustomDoctorRow(row) &&
    !(row.name || '').trim() &&
    !row.present &&
    !(row.checkIn || '').trim() &&
    !(row.lunchOut || '').trim() &&
    !(row.lunchIn || '').trim() &&
    !(row.checkOut || '').trim();

  const getStaffRowsMissingIncident = (rows: typeof tableRows) =>
    rows.filter(
      row =>
        !isBlankCustomStaffRow(row) &&
        !row.present &&
        !row.subAnother &&
        !(row.incident || '').trim()
    );

  function buildAttendancePayload(
    rowsToSave: typeof tableRows,
    doctorsToSave: typeof doctorRows,
    date: string,
    rawFilledBy: string,
    rawCheckedBy: string
  ) {
    const safeFilledBy = validateInput(rawFilledBy, 100);
    const safeCheckedBy = validateInput(rawCheckedBy, 100);

    const staffData: AttendanceRow[] = rowsToSave
      .filter(row => !isBlankCustomStaffRow(row))
      .map(row => {
      const validatedPosition = validatePosition(row.position) ? row.position : '';
      const validatedIncident = validateIncident(row.incident || '') ? (row.incident || '') : '';
      return {
        date,
        filledBy: safeFilledBy,
        checkedBy: safeCheckedBy,
        position: validatedPosition,
        count: 0,
        no: typeof row.no === 'number' && row.no >= 0 && row.no <= 9999 ? row.no : 0,
        name: validateInput(row.name, 100),
        present: typeof row.present === 'boolean' ? row.present : false,
        startTardy: validateInput(row.startTardy, 50),
        lateLunch: validateInput(row.lateLunch, 50),
        needsAdj: typeof row.needsAdj === 'boolean' ? row.needsAdj : false,
        overtime: validateInput(row.overtime, 50),
        otCorp: validateInput(row.otCorp, 50),
        subAnother: typeof row.subAnother === 'boolean' ? row.subAnother : false,
        incident: validatedIncident,
        notes: validateInput(row.notes, 500)
      };
    });

    const positionCounts: { [key: string]: number } = {};
    staffData.forEach(row => {
      if (!positionCounts[row.position]) positionCounts[row.position] = 0;
      if (row.present) positionCounts[row.position]++;
    });
    staffData.forEach(row => {
      row.count = positionCounts[row.position] || 0;
    });

    const doctorPresentCount = doctorsToSave.filter(r => r.present && !isBlankCustomDoctorRow(r)).length;
    const doctorData: DoctorRow[] = doctorsToSave
      .filter(row => !isBlankCustomDoctorRow(row))
      .map(row => ({
      date,
      filledBy: safeFilledBy,
      checkedBy: safeCheckedBy,
      position: 'Doctor',
      count: doctorPresentCount,
      no: typeof row.no === 'number' ? row.no : 0,
      name: validateInput(row.name, 100),
      present: typeof row.present === 'boolean' ? row.present : false,
      checkIn: validateInput(row.checkIn, 10),
      lunchOut: validateInput(row.lunchOut, 10),
      lunchIn: validateInput(row.lunchIn, 10),
      checkOut: validateInput(row.checkOut, 10)
    }));

    return { staffData, doctorData, safeFilledBy, safeCheckedBy };
  }

  type StaffTableRow = {
    id: string;
    position: string;
    no: number;
    name: string;
    present: boolean;
    startTardy: string;
    lateLunch: string;
    needsAdj: boolean;
    overtime: string;
    otCorp: string;
    subAnother: boolean;
    incident: string;
    notes: string;
  };

  type DoctorTableRow = {
    id: string;
    no: number;
    name: string;
    present: boolean;
    checkIn: string;
    lunchOut: string;
    lunchIn: string;
    checkOut: string;
  };

  const buildStaffTableRows = (staff: StaffList, savedStaffData: AttendanceRow[]): StaffTableRow[] => {
    const savedRowsMap = new Map<string, StaffTableRow>();
    savedStaffData.forEach((row: AttendanceRow) => {
      const id = `${row.position}-${row.no}`;
      savedRowsMap.set(id, {
        id,
        position: row.position,
        no: row.no,
        name: row.name,
        present: row.present || false,
        startTardy: row.startTardy || '',
        lateLunch: row.lateLunch || '',
        needsAdj: row.needsAdj || false,
        overtime: row.overtime || '',
        otCorp: row.otCorp || '',
        subAnother: row.subAnother || false,
        incident: row.incident || '',
        notes: row.notes || ''
      });
    });

    const newRows: StaffTableRow[] = [];

    POSITION_ORDER.forEach(position => {
      const members = staff[position] || [];
      const activeMembers = members
        .filter(m => {
          if (position === 'Sub' || position === 'Extern') return true;
          return m.active === true;
        })
        .sort(compareByName);

      activeMembers.forEach(member => {
        const rowId = `${position}-${member.no}`;
        const savedRow = savedRowsMap.get(rowId);

        if (savedRow) {
          newRows.push({
            ...savedRow,
            name: member.name || savedRow.name || '',
            position,
            no: member.no
          });
        } else {
          newRows.push({
            id: rowId,
            position,
            no: member.no,
            name: member.name || '',
            present: false,
            startTardy: '',
            lateLunch: '',
            needsAdj: false,
            overtime: '',
            otCorp: '',
            subAnother: false,
            incident: '',
            notes: ''
          });
        }
      });
    });

    savedRowsMap.forEach((savedRow, id) => {
      if (
        savedRow.no >= CUSTOM_ROW_NO_THRESHOLD &&
        !newRows.some(r => r.position === savedRow.position && r.no === savedRow.no)
      ) {
        newRows.push({ ...savedRow, id });
      }
    });

    return POSITION_ORDER.flatMap(position => {
      const positionRows = newRows.filter(r => r.position === position);
      const regularRows = positionRows.filter(r => !isCustomStaffRow(r)).sort(compareByName);
      const customRows = positionRows.filter(r => isCustomStaffRow(r));
      return [...regularRows, ...customRows];
    });
  };

  const buildDoctorTableRows = (doctors: DoctorMember[], savedDoctorData: DoctorRow[]): DoctorTableRow[] => {
    const savedDoctorsMap = new Map<string, DoctorTableRow>();
    savedDoctorData.forEach((row: DoctorRow) => {
      const id = `doctor-${row.no}`;
      savedDoctorsMap.set(id, {
        id,
        no: row.no,
        name: row.name,
        present: row.present || false,
        checkIn: row.checkIn || '',
        lunchOut: row.lunchOut || '',
        lunchIn: row.lunchIn || '',
        checkOut: row.checkOut || ''
      });
    });

    const newRows = [...doctors].sort(compareByName).map(doctor => {
      const rowId = `doctor-${doctor.no}`;
      const savedRow = savedDoctorsMap.get(rowId);

      if (savedRow) {
        return { ...savedRow, name: doctor.name || savedRow.name || '' };
      }

      return {
        id: rowId,
        no: doctor.no,
        name: doctor.name || '',
        present: false,
        checkIn: '',
        lunchOut: '',
        lunchIn: '',
        checkOut: ''
      };
    });

    savedDoctorsMap.forEach((savedRow, id) => {
      if (
        savedRow.no >= CUSTOM_ROW_NO_THRESHOLD &&
        !newRows.some(r => r.no === savedRow.no)
      ) {
        newRows.push({ ...savedRow, id });
      }
    });

    const regularRows = newRows.filter(r => !isCustomDoctorRow(r)).sort(compareByName);
    const customRows = newRows.filter(r => isCustomDoctorRow(r));
    return [...regularRows, ...customRows];
  };

  const resetFormAfterSubmit = useCallback(() => {
    setFilledBy('');
    setCheckedBy('');
    setSelectedOffice('');
    setDoctorRows([]);
    setTableRows([]);
    setStaffList(null);
    isInitialLoadRef.current = true;
    lastSavedDataRef.current = '';
  }, []);

  // 출석 데이터 저장
  const saveAttendanceData = useCallback(async (
    silent: boolean = false,
    override?: {
      tableRows?: typeof tableRows;
      doctorRows?: typeof doctorRows;
      immediate?: boolean;
    }
  ) => {
    const rowsToSave = override?.tableRows ?? tableRows;
    const doctorsToSave = override?.doctorRows ?? doctorRows;

    if (isInitialLoadRef.current && silent && !override?.immediate) {
      return;
    }

    // 🔒 보안: Rate limiting - 자동 저장은 최소 1초 간격, 분당 최대 30회
    const now = Date.now();
    if (silent && !override?.immediate) {
      if (now - lastAutoSaveTimeRef.current < AUTO_SAVE_MIN_INTERVAL_MS) {
        return;
      }
      // 분당 30회 제한
      if (now - autoSaveResetTimeRef.current > 60000) {
        autoSaveCountRef.current = 0;
        autoSaveResetTimeRef.current = now;
      }
      if (autoSaveCountRef.current >= 30) {
        return; // 분당 30회 초과
      }
      autoSaveCountRef.current++;
      lastAutoSaveTimeRef.current = now;
    }

    if (!trackDate) {
      if (!silent) alert('Please select a date');
      return;
    }

    // 🔒 보안: 오피스 선택 확인
    if (!selectedOffice) {
      if (!silent) alert('Please select an office');
      return;
    }

    // 🔒 보안: 날짜 형식 검증
    if (!validateDate(trackDate)) {
      if (!silent) alert('Invalid date format');
      return;
    }

    // 🔒 보안: 오피스 값 검증
    if (!validateOffice(selectedOffice)) {
      if (!silent) alert('Invalid office value');
      return;
    }

    try {
      const { staffData, doctorData, safeFilledBy, safeCheckedBy } = buildAttendancePayload(
        rowsToSave,
        doctorsToSave,
        trackDate,
        filledBy,
        checkedBy
      );

      const data = {
        date: trackDate,
        office: selectedOffice,
        filledBy: safeFilledBy,
        checkedBy: safeCheckedBy,
        staffData,
        doctorData,
        createdAt: Timestamp.now()
      };

      const safeDocId = getAttendanceDocId(trackDate, selectedOffice);
      await setDoc(doc(db, 'attendance-data', safeDocId), data);
      
      // 마지막 저장된 데이터 업데이트 (자동 저장 최적화용)
      lastSavedDataRef.current = JSON.stringify({ tableRows: rowsToSave, doctorRows: doctorsToSave, filledBy, checkedBy });
      
    } catch (error) {
      if (!silent) {
        // 🔒 보안: 상세한 에러 메시지 노출 최소화
        alert('Error saving attendance data. Please try again.');
      }
    }
  }, [trackDate, filledBy, checkedBy, tableRows, doctorRows, selectedOffice]);

  // 임시 row 추가 (특정 position에)
  const addTempRow = useCallback((position: string) => {
    if (!validatePosition(position)) {
      return;
    }
    const positionRows = tableRows.filter(r => r.position === position);
    const customRows = positionRows.filter(r => isCustomStaffRow(r));
    const tempNo = CUSTOM_ROW_NO_THRESHOLD + customRows.length;

    const newTempRow = {
      id: `${position}-${tempNo}`,
      position,
      no: tempNo,
      name: '',
      present: false,
      startTardy: '',
      lateLunch: '',
      needsAdj: false,
      overtime: '',
      otCorp: '',
      subAnother: false,
      incident: '',
      notes: ''
    };

    const positionIndex = tableRows.findIndex(r => r.position === position);
    let newRows: typeof tableRows;
    if (positionIndex === -1) {
      newRows = [...tableRows, newTempRow];
    } else {
      let insertIndex = positionIndex;
      for (let i = positionIndex; i < tableRows.length; i++) {
        if (tableRows[i].position === position) {
          insertIndex = i + 1;
        } else {
          break;
        }
      }
      newRows = [...tableRows];
      newRows.splice(insertIndex, 0, newTempRow);
    }

    setTableRows(newRows);
    if (trackDate && selectedOffice && !isInitialLoadRef.current) {
      saveAttendanceData(true, { tableRows: newRows, immediate: true });
    }
  }, [tableRows, trackDate, selectedOffice, saveAttendanceData]);

  const addDoctor = useCallback(() => {
    const customDoctors = doctorRows.filter(r => isCustomDoctorRow(r));
    const tempNo = CUSTOM_ROW_NO_THRESHOLD + customDoctors.length;

    const newDoctor = {
      id: `doctor-${tempNo}`,
      no: tempNo,
      name: '',
      present: false,
      checkIn: '',
      lunchOut: '',
      lunchIn: '',
      checkOut: ''
    };

    const newDoctorRows = [...doctorRows, newDoctor];
    setDoctorRows(newDoctorRows);
    if (trackDate && selectedOffice && !isInitialLoadRef.current) {
      saveAttendanceData(true, { doctorRows: newDoctorRows, immediate: true });
    }
  }, [doctorRows, trackDate, selectedOffice, saveAttendanceData]);

  const removeCustomStaffRow = useCallback((rowId: string) => {
    const newRows = tableRows.filter(r => r.id !== rowId);
    setTableRows(newRows);
    if (trackDate && selectedOffice && !isInitialLoadRef.current) {
      saveAttendanceData(true, { tableRows: newRows, immediate: true });
    }
  }, [tableRows, trackDate, selectedOffice, saveAttendanceData]);

  const removeCustomDoctorRow = useCallback((rowId: string) => {
    const newRows = doctorRows.filter(r => r.id !== rowId);
    setDoctorRows(newRows);
    if (trackDate && selectedOffice && !isInitialLoadRef.current) {
      saveAttendanceData(true, { doctorRows: newRows, immediate: true });
    }
  }, [doctorRows, trackDate, selectedOffice, saveAttendanceData]);

  // Submit (데이터 저장 + PDF 생성)
  const handleSubmit = useCallback(async () => {
    if (!trackDate) {
      alert('Please select a date');
      return;
    }

    // 🔒 보안: 오피스 선택 확인
    if (!selectedOffice) {
      alert('Please select an office');
      return;
    }

    // 🔒 보안: 날짜 형식 검증
    if (!validateDate(trackDate)) {
      alert('Invalid date format');
      return;
    }

    // 🔒 보안: 오피스 값 검증
    if (!validateOffice(selectedOffice)) {
      alert('Invalid office value');
      return;
    }

    const rowsMissingIncident = getStaffRowsMissingIncident(tableRows);
    if (rowsMissingIncident.length > 0) {
      const labels = rowsMissingIncident
        .map(r => r.name?.trim() || `${r.position} #${r.no}`)
        .join(', ');
      alert(`Please select an Incident Description for each staff member not marked Present.\n\nMissing: ${labels}`);
      return;
    }

    try {
      // 🔒 보안: Rate limiting - API 호출은 최소 5초 간격
      const currentTime = Date.now();
      if (currentTime - lastApiCallTimeRef.current < 5000) {
        alert('Please wait a moment before submitting again.');
        return;
      }
      lastApiCallTimeRef.current = currentTime;

      const { staffData, doctorData, safeFilledBy, safeCheckedBy } = buildAttendancePayload(
        tableRows,
        doctorRows,
        trackDate,
        filledBy,
        checkedBy
      );

      const safeDocId = getAttendanceDocId(trackDate, selectedOffice);
      await setDoc(doc(db, 'attendance-data', safeDocId), {
        date: trackDate,
        office: selectedOffice,
        filledBy: safeFilledBy,
        checkedBy: safeCheckedBy,
        staffData,
        doctorData,
        createdAt: Timestamp.now()
      });

      // 2. PDF 생성 (client-side)
      const currentDate = new Date();
      const formatter = new Intl.DateTimeFormat('en-US', {
        timeZone: 'America/Los_Angeles',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: true
      });
      const generatedDate = formatter.format(currentDate);

      // 🔒 보안: 파일명 및 경로 검증 (특수문자 제거)
      const safeDate = trackDate.replace(/[^a-zA-Z0-9_-]/g, '');
      const safeOffice = selectedOffice.replace(/[^a-zA-Z0-9_-]/g, '');

      try {
        const pdfBuffer = await pdf(createAttendancePDFDocument({
          safeDate: trackDate,
          safeOffice: selectedOffice,
          safeFilledBy,
          safeCheckedBy,
          staffData,
          doctorData,
          generatedDate,
        })).toBlob();

        if (!pdfBuffer || pdfBuffer.size === 0) {
          throw new Error('PDF is empty');
        }

        try {
          const storage = getStorage();
          // 캘리포니아 시간으로 짧은 타임스탬프 생성 (예: 230pm)
          const now = new Date();
          const laTime = new Date(now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
          let hours = laTime.getHours();
          const minutes = laTime.getMinutes();
          const ampm = hours >= 12 ? 'pm' : 'am';
          hours = hours % 12;
          hours = hours ? hours : 12;
          const timeStamp = `${hours}${minutes.toString().padStart(2, '0')}${ampm}`;
          
          const filename = `3) ${safeDate}_${safeOffice}_Attendance Tract_${timeStamp}.pdf`;
          const storageRef = ref(storage, `endofday-pdfs/${safeOffice}/${safeDate}/${filename}`);
          
          // PDF 업로드
          await uploadBytes(storageRef, pdfBuffer);
          
        } catch (storageError: any) {
          const errorMsg = storageError?.message || '알 수 없는 오류';
          throw new Error(`An error occurred while submitting. Please try again.: ${errorMsg}`);
        }
      } catch (pdfError: any) {
        throw new Error(`An error occurred while submitting. Please try again.: ${pdfError?.message || 'Unknown error'}`);
      }

      alert('Submitted Successfully!');
      
      await deleteDoc(doc(db, 'attendance-data', safeDocId));
      
      resetFormAfterSubmit();
      
    } catch (error) {
      // 🔒 보안: 상세한 에러 메시지 노출 최소화
      alert('An error occurred while submitting. Please try again.');
    }
  }, [trackDate, filledBy, checkedBy, tableRows, doctorRows, selectedOffice, resetFormAfterSubmit]);

  // Staff List + 출석 임시 데이터 로드 (오피스/날짜 변경 시 1회, 원자적 적용)
  useEffect(() => {
    if (!selectedOffice || !validateOffice(selectedOffice)) {
      setStaffList({ staff: {}, doctors: [] });
      setTableRows([]);
      setDoctorRows([]);
      setFilledBy('');
      setCheckedBy('');
      return;
    }

    let cancelled = false;
    const loadOffice = selectedOffice;
    const loadDate = trackDate;
    const canLoadAttendance = Boolean(loadDate && validateDate(loadDate));

    setStaffList(null);
    setTableRows([]);
    setDoctorRows([]);
    setFilledBy('');
    setCheckedBy('');
    isInitialLoadRef.current = true;
    lastSavedDataRef.current = '';

    (async () => {
      try {
        const staffPromise = getDoc(doc(db, 'staff-list', loadOffice));
        const attendancePromise = canLoadAttendance
          ? getDoc(doc(db, 'attendance-data', getAttendanceDocId(loadDate, loadOffice)))
          : Promise.resolve(null);

        const [staffSnap, attendanceSnap] = await Promise.all([staffPromise, attendancePromise]);
        if (cancelled) return;

        const staff = staffSnap.exists() ? (staffSnap.data().staff || {}) : {};
        const doctors: DoctorMember[] = staffSnap.exists() ? (staffSnap.data().doctors || []) : [];
        const staffData: AttendanceRow[] = [];
        const doctorData: DoctorRow[] = [];
        let loadedFilledBy = '';
        let loadedCheckedBy = '';

        if (canLoadAttendance && attendanceSnap?.exists()) {
          const docData = attendanceSnap.data();
          staffData.push(...(docData.staffData || []));
          doctorData.push(...(docData.doctorData || []));
          loadedFilledBy = docData.filledBy || '';
          loadedCheckedBy = docData.checkedBy || '';
        }

        setStaffList({ staff, doctors });
        setFilledBy(loadedFilledBy);
        setCheckedBy(loadedCheckedBy);
        setTableRows(buildStaffTableRows(staff, staffData));
        setDoctorRows(buildDoctorTableRows(doctors, doctorData));
      } catch {
        if (!cancelled) {
          alert('Error loading data. Please try again.');
        }
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [selectedOffice, trackDate]);

  // 자동 저장 (debounce 적용, 깜빡임 방지)
  useEffect(() => {
    // 오피스가 선택되지 않았거나 초기 로드 시에는 저장하지 않음
    if (!selectedOffice || (tableRows.length === 0 && doctorRows.length === 0)) return;
    
    // 초기 로드 시에는 자동 저장하지 않음 (데이터 로드 후 한 번만 설정)
    if (isInitialLoadRef.current) {
      // 초기 로드 완료 후 현재 데이터를 마지막 저장된 데이터로 설정
      setTimeout(() => {
        lastSavedDataRef.current = JSON.stringify({ tableRows, doctorRows, filledBy, checkedBy });
        isInitialLoadRef.current = false;
      }, 1000);
      return;
    }
    
    // 데이터가 실제로 변경되었는지 확인
    const currentData = JSON.stringify({ tableRows, doctorRows, filledBy, checkedBy });
    if (currentData === lastSavedDataRef.current) {
      return; // 변경사항이 없으면 저장하지 않음
    }
    
    // debounce: 1초 후에 저장
    const timer = setTimeout(() => {
      if (trackDate && selectedOffice) {
        saveAttendanceData(true);
      }
    }, AUTO_SAVE_DEBOUNCE_MS);

    return () => clearTimeout(timer);
  }, [tableRows, doctorRows, filledBy, checkedBy, trackDate, selectedOffice, saveAttendanceData]);

  // 사용자 offices 정보 로드
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      if (!currentUser) return;

      try {
        const userDoc = await getDoc(doc(db, 'users', currentUser.uid));
        if (!userDoc.exists()) return;

        const userData = userDoc.data();

        if (userData?.offices) {
          const officesArray = Array.isArray(userData.offices)
            ? userData.offices
            : [userData.offices];

          const validOptions = officesArray.filter((g: string) =>
            (OFFICE_OPTIONS as readonly string[]).includes(g)
          );

          if (validOptions.length > 0) {
            setUserOfficesOptions(validOptions);
            if (validOptions.length === 1) {
              setSelectedOffice(validOptions[0]);
            }
          }
        }
      } catch (error) {
        console.error('Failed to load user offices:', error);
      }
    });

    if (process.env.NODE_ENV === 'production' &&
        typeof window !== 'undefined' &&
        window.location.protocol !== 'https:') {
      window.location.href = window.location.href.replace('http:', 'https:');
    }

    return () => {
      unsubscribe();
    };
  }, []);

  return (
    <div style={ui.page}>
      <h1 style={ui.title}>
        Attendance Tract
      </h1>
      
      <div style={ui.subtitle}>
        Management is required to review all staff members' times on Time Clock to fill out the Attendance Tract Sheet accurately and to request any necessary clock adjustments.
      </div>

      <div style={ui.formCard}>
        {/* offices 옵션이 있는 경우에만 Office 표시 */}
        {userOfficesOptions.length > 0 && (
          <label style={ui.label}>
            <span>Office:</span>
            {userOfficesOptions.length === 1 ? (
              <span style={ui.officeBadge}>
                {selectedOffice}
              </span>
            ) : (
              <select
                value={selectedOffice}
                onChange={(e) => setSelectedOffice(e.target.value)}
                style={{ ...ui.input, minWidth: '120px' }}
              >
                <option value="">Select Office</option>
                {userOfficesOptions.map(office => (
                  <option key={office} value={office}>{office}</option>
                ))}
              </select>
            )}
          </label>
        )}
        
        <label style={ui.label}>
          <span>Date:</span>
          <input
            type="date"
            value={trackDate}
            onChange={(e) => setTrackDate(e.target.value)}
            style={ui.input}
          />
        </label>
        
        <label style={ui.label}>
          <span>Filled Out By:</span>
          <input
            type="text"
            value={filledBy}
            onChange={(e) => {
              const validatedValue = validateInput(e.target.value, 100);
              setFilledBy(validatedValue);
            }}
            style={{ ...ui.input, minWidth: '150px' }}
          />
        </label>
        
        <label style={ui.label}>
          <span>Management that Checked Times Today on Time Clock:</span>
          <input
            type="text"
            value={checkedBy}
            onChange={(e) => {
              const validatedValue = validateInput(e.target.value, 100);
              setCheckedBy(validatedValue);
            }}
            style={{ ...ui.input, minWidth: '150px' }}
          />
        </label>
      </div>

      {/* Staff 테이블 */}
      {staffList && (
        <div style={{ marginTop: '20px', overflowX: 'auto' }}>
          <div style={ui.tableCard}>
            {/* 고정 헤더 테이블 */}
            <div style={{ position: 'sticky', top: 0, zIndex: 100, background: theme.card }}>
              <table style={ui.table}>
                <colgroup>
                  <col style={{ width: '5%' }} />
                  <col style={{ width: '12%' }} />
                  <col style={{ width: '6%' }} />
                  <col style={{ width: '10%' }} />
                  <col style={{ width: '10%' }} />
                  <col style={{ width: '8%' }} />
                  <col style={{ width: '8%' }} />
                  <col style={{ width: '10%' }} />
                  <col style={{ width: '8%' }} />
                  <col style={{ width: '13%' }} />
                  <col style={{ width: '10%' }} />
                </colgroup>
                <thead>
                  <tr>
                    <th style={ui.th}>No.</th>
                    <th style={ui.th}>Name</th>
                    <th style={ui.th}>Present</th>
                    <th style={ui.th}>Start Shift<br />Tardy (Min)</th>
                    <th style={ui.th}>Late from<br />Lunch (Min)</th>
                    <th style={ui.th}>Needs Clock Adj.</th>
                    <th style={ui.th}>Overtime</th>
                    <th style={ui.th}>OT Corp Authorized<br />By (Initials)</th>
                    <th style={ui.th}>Sub. at Another<br />Office</th>
                    <th style={ui.th}>Incident Description</th>
                    <th style={ui.th}>Notes</th>
                  </tr>
                </thead>
              </table>
            </div>
            {/* 스크롤 가능한 본문 */}
            <div style={{ maxHeight: '70vh', overflowY: 'auto' }}>
              <table style={ui.table}>
                <colgroup>
                  <col style={{ width: '5%' }} />
                  <col style={{ width: '12%' }} />
                  <col style={{ width: '6%' }} />
                  <col style={{ width: '10%' }} />
                  <col style={{ width: '10%' }} />
                  <col style={{ width: '8%' }} />
                  <col style={{ width: '8%' }} />
                  <col style={{ width: '10%' }} />
                  <col style={{ width: '8%' }} />
                  <col style={{ width: '13%' }} />
                  <col style={{ width: '10%' }} />
                </colgroup>
                <tbody>
              {(() => {
                const rows: React.ReactElement[] = [];
                
                POSITION_ORDER.forEach(position => {
                  const positionRows = tableRows.filter(r => r.position === position);
                  const presentCount = positionRows.filter(r => r.present).length;
                  
                  rows.push(
                    <tr key={`header-${position}`} style={{ background: theme.positionBg, fontWeight: 600, color: theme.text }}>
                      <td colSpan={11} style={{ ...ui.tdCenter, padding: '12px', position: 'relative' }}>
                        <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', position: 'relative' }}>
                          <span style={{ textAlign: 'center' }}>{position} <span style={{ marginLeft: '10px', color: theme.gray }}>{presentCount}</span></span>
                          <button
                            onClick={() => addTempRow(position)}
                            style={ui.addBtn}
                          >
                            Add Staff
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                  
                  // Position별 행들 추가
                  positionRows.forEach((row, rowIdx) => {
                  
                  // 데이터 행
                  rows.push(
                    <tr key={row.id} style={{ background: rowIdx % 2 === 0 ? theme.card : theme.cardAlt }}>
                      <td style={ui.tdCenter}>
                        {isCustomStaffRow(row) ? (
                          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '4px' }}>
                            <span>{row.no}</span>
                            <button
                              type="button"
                              onClick={() => removeCustomStaffRow(row.id)}
                              style={ui.removeBtn}
                              title="Remove row"
                            >
                              ×
                            </button>
                          </div>
                        ) : (
                          row.no
                        )}
                      </td>
                      <td style={{ ...ui.td, fontWeight: 600 }}>
                        {(row.position === 'Sub' || row.position === 'Extern' || isCustomStaffRow(row)) ? (
                          <input
                            type="text"
                            value={row.name}
                            onChange={(e) => {
                              // 🔒 보안: 입력 검증
                              const validatedValue = validateInput(e.target.value, 100);
                              const newRows = [...tableRows];
                              const rowIndex = newRows.findIndex(r => r.id === row.id);
                              if (rowIndex !== -1) {
                                newRows[rowIndex].name = validatedValue;
                                setTableRows(newRows);
                              }
                            }}
                            style={{ ...ui.cellInput, fontWeight: 600 }}
                            placeholder="Enter name"
                          />
                        ) : (
                          row.name
                        )}
                      </td>
                      <td style={ui.tdCenter}>
                        <input
                          type="checkbox"
                          checked={row.present}
                          onChange={(e) => {
                            const newRows = [...tableRows];
                            const rowIndex = newRows.findIndex(r => r.id === row.id);
                            if (rowIndex !== -1) {
                              newRows[rowIndex].present = e.target.checked;
                              setTableRows(newRows);
                            }
                          }}
                          style={ui.checkbox}
                        />
                      </td>
                      <td style={ui.td}>
                        <input
                          type="text"
                          value={row.startTardy}
                            onChange={(e) => {
                              // 🔒 보안: 입력 검증
                              const validatedValue = validateInput(e.target.value, 50);
                              const newRows = [...tableRows];
                              const rowIndex = newRows.findIndex(r => r.id === row.id);
                              if (rowIndex !== -1) {
                                newRows[rowIndex].startTardy = validatedValue;
                                setTableRows(newRows);
                              }
                            }}
                          style={ui.cellInput}
                        />
                      </td>
                      <td style={ui.td}>
                        <input
                          type="text"
                          value={row.lateLunch}
                            onChange={(e) => {
                              // 🔒 보안: 입력 검증
                              const validatedValue = validateInput(e.target.value, 50);
                              const newRows = [...tableRows];
                              const rowIndex = newRows.findIndex(r => r.id === row.id);
                              if (rowIndex !== -1) {
                                newRows[rowIndex].lateLunch = validatedValue;
                                setTableRows(newRows);
                              }
                            }}
                          style={ui.cellInput}
                        />
                      </td>
                      <td style={ui.tdCenter}>
                        <input
                          type="checkbox"
                          checked={row.needsAdj}
                          onChange={(e) => {
                            const newRows = [...tableRows];
                            const rowIndex = newRows.findIndex(r => r.id === row.id);
                            if (rowIndex !== -1) {
                              newRows[rowIndex].needsAdj = e.target.checked;
                              setTableRows(newRows);
                            }
                          }}
                          style={ui.checkbox}
                        />
                      </td>
                      <td style={ui.td}>
                        <input
                          type="text"
                          value={row.overtime}
                            onChange={(e) => {
                              // 🔒 보안: 입력 검증
                              const validatedValue = validateInput(e.target.value, 50);
                              const newRows = [...tableRows];
                              const rowIndex = newRows.findIndex(r => r.id === row.id);
                              if (rowIndex !== -1) {
                                newRows[rowIndex].overtime = validatedValue;
                                setTableRows(newRows);
                              }
                            }}
                          style={ui.cellInput}
                        />
                      </td>
                      <td style={ui.td}>
                        <input
                          type="text"
                          value={row.otCorp}
                            onChange={(e) => {
                              // 🔒 보안: 입력 검증
                              const validatedValue = validateInput(e.target.value, 50);
                              const newRows = [...tableRows];
                              const rowIndex = newRows.findIndex(r => r.id === row.id);
                              if (rowIndex !== -1) {
                                newRows[rowIndex].otCorp = validatedValue;
                                setTableRows(newRows);
                              }
                            }}
                          style={ui.cellInput}
                        />
                      </td>
                      <td style={ui.tdCenter}>
                        <input
                          type="checkbox"
                          checked={row.subAnother}
                          onChange={(e) => {
                            const newRows = [...tableRows];
                            const rowIndex = newRows.findIndex(r => r.id === row.id);
                            if (rowIndex !== -1) {
                              newRows[rowIndex].subAnother = e.target.checked;
                              setTableRows(newRows);
                            }
                          }}
                          style={ui.checkbox}
                        />
                      </td>
                      <td style={ui.td}>
                        <select
                          value={row.incident || ''}
                          onChange={(e) => {
                            // 🔒 보안: Incident 값 검증
                            const value = e.target.value;
                            if (!validateIncident(value)) {
                              return; // 유효하지 않은 값은 무시
                            }
                            const newRows = [...tableRows];
                            const rowIndex = newRows.findIndex(r => r.id === row.id);
                            if (rowIndex !== -1) {
                              newRows[rowIndex].incident = value;
                              setTableRows(newRows);
                            }
                          }}
                          style={{ ...ui.cellInput, fontSize: '14px' }}
                        >
                          {INCIDENT_OPTIONS.map(option => (
                            <option key={option} value={option}>
                              {option || 'Select...'}
                            </option>
                          ))}
                        </select>
                      </td>
                      <td style={ui.td}>
                        <input
                          type="text"
                          value={row.notes}
                            onChange={(e) => {
                              // 🔒 보안: 입력 검증
                              const validatedValue = validateInput(e.target.value, 500);
                              const newRows = [...tableRows];
                              const rowIndex = newRows.findIndex(r => r.id === row.id);
                              if (rowIndex !== -1) {
                                newRows[rowIndex].notes = validatedValue;
                                setTableRows(newRows);
                              }
                            }}
                          style={ui.cellInput}
                        />
                      </td>
                    </tr>
                  );
                  });
                });
                
                return rows;
              })()}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* Doctor 테이블 */}
      {staffList && selectedOffice && staffList.doctors && staffList.doctors.length > 0 && (
        <div key={`doctor-${selectedOffice}-${trackDate}`} style={{ marginTop: '40px', overflowX: 'auto' }}>
          <table style={{ ...ui.table, ...ui.tableCard }}>
            <thead>
              <tr>
                <th style={ui.th}>No.</th>
                <th style={ui.th}>Name</th>
                <th style={ui.th}>Present</th>
                <th style={ui.th}>Check In</th>
                <th style={ui.th}>Lunch Out</th>
                <th style={ui.th}>Lunch In</th>
                <th style={ui.th}>Check Out</th>
              </tr>
              <tr style={{ background: theme.positionBg, fontWeight: 600, color: theme.text }}>
                <td colSpan={7} style={{ ...ui.tdCenter, padding: '12px', position: 'relative' }}>
                  <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', position: 'relative' }}>
                    <span style={{ textAlign: 'center' }}>Doctor <span style={{ marginLeft: '10px', color: theme.gray }}>{doctorRows.filter(r => r.present).length}</span></span>
                    <button
                      onClick={addDoctor}
                      style={ui.addBtn}
                    >
                      Add Doctor
                    </button>
                  </div>
                </td>
              </tr>
            </thead>
            <tbody>
              {doctorRows.map((row, rowIdx) => (
                <tr key={row.id} style={{ background: rowIdx % 2 === 0 ? theme.card : theme.cardAlt }}>
                  <td style={ui.tdCenter}>
                    {isCustomDoctorRow(row) ? (
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '4px' }}>
                        <span>{row.no}</span>
                        <button
                          type="button"
                          onClick={() => removeCustomDoctorRow(row.id)}
                          style={ui.removeBtn}
                          title="Remove row"
                        >
                          ×
                        </button>
                      </div>
                    ) : (
                      row.no
                    )}
                  </td>
                  <td style={{ ...ui.td, fontWeight: 600 }}>
                    {isCustomDoctorRow(row) ? (
                      <input
                        type="text"
                        value={row.name}
                        onChange={(e) => {
                          // 🔒 보안: 입력 검증 (길이 제한)
                          const validatedValue = validateInput(e.target.value, 100);
                          const newRows = [...doctorRows];
                          const rowIndex = newRows.findIndex(r => r.id === row.id);
                          if (rowIndex !== -1) {
                            newRows[rowIndex].name = validatedValue;
                            setDoctorRows(newRows);
                          }
                        }}
                        style={{ ...ui.cellInput, fontWeight: 600 }}
                        placeholder="Enter name"
                        maxLength={100}
                      />
                    ) : (
                      row.name
                    )}
                  </td>
                  <td style={ui.tdCenter}>
                    <input
                      type="checkbox"
                      checked={row.present}
                      onChange={(e) => {
                        const newRows = [...doctorRows];
                        const rowIndex = newRows.findIndex(r => r.id === row.id);
                        if (rowIndex !== -1) {
                          newRows[rowIndex].present = e.target.checked;
                          setDoctorRows(newRows);
                        }
                      }}
                      style={ui.checkbox}
                    />
                  </td>
                  <td style={ui.td}>
                    <input
                      type="time"
                      value={row.checkIn}
                      onChange={(e) => {
                        const newRows = [...doctorRows];
                        const rowIndex = newRows.findIndex(r => r.id === row.id);
                        if (rowIndex !== -1) {
                          newRows[rowIndex].checkIn = e.target.value;
                          setDoctorRows(newRows);
                        }
                      }}
                      style={ui.cellInput}
                    />
                  </td>
                  <td style={ui.td}>
                    <input
                      type="time"
                      value={row.lunchOut}
                      onChange={(e) => {
                        const newRows = [...doctorRows];
                        const rowIndex = newRows.findIndex(r => r.id === row.id);
                        if (rowIndex !== -1) {
                          newRows[rowIndex].lunchOut = e.target.value;
                          setDoctorRows(newRows);
                        }
                      }}
                      style={ui.cellInput}
                    />
                  </td>
                  <td style={ui.td}>
                    <input
                      type="time"
                      value={row.lunchIn}
                      onChange={(e) => {
                        const newRows = [...doctorRows];
                        const rowIndex = newRows.findIndex(r => r.id === row.id);
                        if (rowIndex !== -1) {
                          newRows[rowIndex].lunchIn = e.target.value;
                          setDoctorRows(newRows);
                        }
                      }}
                      style={ui.cellInput}
                    />
                  </td>
                  <td style={ui.td}>
                    <input
                      type="time"
                      value={row.checkOut}
                      onChange={(e) => {
                        const newRows = [...doctorRows];
                        const rowIndex = newRows.findIndex(r => r.id === row.id);
                        if (rowIndex !== -1) {
                          newRows[rowIndex].checkOut = e.target.value;
                          setDoctorRows(newRows);
                        }
                      }}
                      style={ui.cellInput}
                    />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <div style={{ marginTop: '32px', textAlign: 'center', paddingBottom: '24px' }}>
        <button
          onClick={handleSubmit}
          style={ui.submitBtn}
        >
          Submit
        </button>
      </div>
    </div>
  );
}
