'use client';

import React, { useState, useEffect, useCallback } from "react";
import { doc, setDoc, getDoc, deleteDoc, onSnapshot } from "firebase/firestore";
import { db, auth } from "@/lib/firebase.config";
import { getStorage, ref, uploadBytes } from 'firebase/storage';
import { onAuthStateChanged } from 'firebase/auth';
import { pdf, Document, Page, View, Text, StyleSheet } from '@react-pdf/renderer';

// 🔒 보안: 입력 검증 함수
const validateInput = (value: string, maxLength: number = 500): string => {
  if (typeof value !== 'string') return '';
  // 길이 제한
  if (value.length > maxLength) {
    return value.substring(0, maxLength);
  }
  return value;
};

// 🔒 보안: 날짜 형식 검증
const validateDate = (date: string): boolean => {
  if (!date || typeof date !== 'string') return false;
  // MM/DD/YYYY 또는 YYYY-MM-DD 형식 확인
  const dateRegex1 = /^\d{2}\/\d{2}\/\d{4}$/;
  const dateRegex2 = /^\d{4}-\d{2}-\d{2}$/;
  return dateRegex1.test(date) || dateRegex2.test(date);
};

// 🔒 보안: 오피스 값 검증
const validateOffice = (office: string): boolean => {
  if (!office || typeof office !== 'string') return false;
  const allowedOffices = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];
  return allowedOffices.includes(office);
};

// 고정된 폼 이름 배열 (재사용)
const FIXED_FORM_NAMES = [
            'Daily Front Office Duties',
            'Attendance Tract Sheet',
            'Clock Adjustment',
            'Excuse Note',
            'Time Off Request',
            'Request for Sched. Change',
            'Written Warning',
            'Record of Conversation',
            'Incident Notice',
            'Restroom Log',
            'Add on Treatment Log',
            'Scheduled Appts Log',
            'Mileage Log',
            'Lobby Inspection Log',
            'RDA Sheets',
            'X-Ray/IOPs Before Treatment',
            'Covid-19 Daily Screening Log',
            'Other:',
            'Other:'
          ];

export default function FaxCoverPage() {
  // 시간 형식 변환 헬퍼 함수 (24시간제 -> 12시간제)
  const convertTo12Hour = (timeStr: string): string => {
    if (!timeStr) return '12:00 AM';
    // 이미 AM/PM이 포함되어 있으면 그대로 반환
    if (timeStr.includes('AM') || timeStr.includes('PM')) return timeStr;
    // 24시간제 형식 (HH:MM)을 12시간제로 변환
    const parts = timeStr.split(':');
    if (parts.length === 2) {
      let hour = parseInt(parts[0], 10);
      const minute = parts[1].trim();
      const ampm = hour >= 12 ? 'PM' : 'AM';
      hour = hour % 12;
      if (hour === 0) hour = 12;
      return `${hour.toString().padStart(2, '0')}:${minute} ${ampm}`;
    }
    return '12:00 AM';
  };

  // 시간 파싱 헬퍼 함수
  const parseTime = (timeStr: string): { hour: string; minute: string; ampm: string } => {
    // 빈 값이면 빈 값 반환
    if (!timeStr || timeStr.trim() === '') {
      return { hour: '', minute: '', ampm: '' };
    }
    const converted = convertTo12Hour(timeStr);
    const parts = converted.split(' ');
    const timePart = parts[0] || '12:00';
    const ampm = parts[1] || 'AM';
    const timeParts = timePart.split(':');
    const hour = timeParts[0] || '12';
    const minute = timeParts[1] || '00';
    return { hour, minute, ampm };
  };

  // 날짜 형식 변환: MM/DD/YYYY -> YYYY-MM-DD
  const convertDateToISO = (dateStr: string): string => {
    if (!dateStr) return '';
    // 이미 YYYY-MM-DD 형식인 경우
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return dateStr;
    // MM/DD/YYYY 형식을 YYYY-MM-DD로 변환
    const parts = dateStr.split('/');
    if (parts.length === 3) {
      const month = parts[0].padStart(2, '0');
      const day = parts[1].padStart(2, '0');
      const year = parts[2];
      return `${year}-${month}-${day}`;
    }
    return '';
  };

  // --- PDF 생성 관련 상수/스타일 ---
  const pdfStyles = StyleSheet.create({
    page: { padding: 15, fontFamily: 'Helvetica', fontSize: 8 },
    header: { marginBottom: 6, borderBottomWidth: 1, borderColor: '#333', paddingBottom: 3, alignItems: 'center' },
    headerTitle: { fontSize: 16, fontWeight: 'bold', marginBottom: 2 },
    headerSubtitle: { fontSize: 9, color: '#666' },
    infoSection: { flexDirection: 'row', justifyContent: 'space-between', marginBottom: 5, fontSize: 8 },
    infoItem: { flexDirection: 'row', gap: 3 },
    infoLabel: { fontWeight: 'bold' },
    table: { marginTop: 5, marginBottom: 5 },
    tableRow: { flexDirection: 'row', borderBottomWidth: 0.5, borderColor: '#333' },
    tableCell: { padding: 2.5, fontSize: 7, flex: 1, borderRightWidth: 0.5, borderColor: '#333', justifyContent: 'center', alignItems: 'center' },
    tableCellNo: { flex: 0.3 },
    tableCellName: { flex: 2 },
    tableCellQty: { flex: 0.8 },
    tableHeader: { backgroundColor: '#f0f0f0', fontWeight: 'bold' },
    section: { marginTop: 6, marginBottom: 4 },
    sectionTitle: { fontSize: 10, fontWeight: 'bold', marginBottom: 2, borderBottomWidth: 1, borderColor: '#ddd', paddingBottom: 1 },
    sectionContent: { fontSize: 8, marginBottom: 2 },
    sectionRow: { flexDirection: 'row', marginBottom: 2 },
    sectionLabel: { fontWeight: 'bold', marginRight: 4 },
    sideBySideRow: { flexDirection: 'row', gap: 10, marginTop: 6, marginBottom: 4 },
    sideBySideColumn: { flex: 1 },
    footer: { marginTop: 8, paddingTop: 5, borderTopWidth: 1, borderColor: '#ddd', alignItems: 'center', fontSize: 7, color: '#666' },
  });

  // PDF 생성 유틸 함수
  function safeStr(v: unknown, max: number): string {
    if (v == null) return '';
    return String(v).trim().slice(0, max).replace(/[<>]/g, '');
  }

  function createFaxCoverPDFDocument(props: {
    safeFaxDate: string;
    safeSelectedOffice: string;
    formData: {
      date: string;
      officeTimeCheckIn: string;
      officeName: string;
      timeCheckOut: string;
      name: string;
    };
    tableData: Array<{
      nameOfForm: string;
      otherText: string;
      qty: string;
    }>;
    productionData: Array<{
      date: string;
      note: string;
      status: string;
    }>;
    todayData: {
      addOns: string;
      noShows: string;
      seen: string;
    };
    nextDayData: {
      opener: string;
      closer: string;
    };
    callLogData: {
      whoCalled: string;
      appointmentsMade: string;
    };
    supervisorData: {
      officeSupervisorManager: string;
      checkOutBy: string;
    };
    generatedDate: string;
  }) {
    const { safeFaxDate, safeSelectedOffice, formData, tableData, productionData, todayData, nextDayData, callLogData, supervisorData, generatedDate } = props;
    const s = pdfStyles;

    // Header
    const header = React.createElement(View, { style: s.header },
      React.createElement(Text, { style: s.headerTitle }, 'End of Day Fax Cover'),
      React.createElement(Text, { style: s.headerSubtitle }, '(Check out only when leaving the office)'),
    );

    // Info Section 1: Date, Office
    const infoSection1 = React.createElement(View, { style: s.infoSection },
      React.createElement(View, { style: s.infoItem },
        React.createElement(Text, { style: s.infoLabel }, 'Date: '),
        React.createElement(Text, null, safeStr(formData.date, 20)),
      ),
      React.createElement(View, { style: s.infoItem },
        React.createElement(Text, { style: s.infoLabel }, 'Office: '),
        React.createElement(Text, null, safeSelectedOffice),
      ),
    );

    // Info Section 2: Time Check In, Name, Time Check Out, Name
    const infoSection2 = React.createElement(View, { style: s.infoSection },
      React.createElement(View, { style: s.infoItem },
        React.createElement(Text, { style: s.infoLabel }, 'Time Check In: '),
        React.createElement(Text, null, safeStr(formData.officeTimeCheckIn, 20)),
      ),
      React.createElement(View, { style: s.infoItem },
        React.createElement(Text, { style: s.infoLabel }, 'Name: '),
        React.createElement(Text, null, safeStr(formData.officeName, 100)),
      ),
      React.createElement(View, { style: s.infoItem },
        React.createElement(Text, { style: s.infoLabel }, 'Time Check Out: '),
        React.createElement(Text, null, safeStr(formData.timeCheckOut, 20)),
      ),
      React.createElement(View, { style: s.infoItem },
        React.createElement(Text, { style: s.infoLabel }, 'Name: '),
        React.createElement(Text, null, safeStr(formData.name, 100)),
      ),
    );

    // Table Header
    const tableHeader = React.createElement(View, { style: [s.tableRow, s.tableHeader] },
      React.createElement(View, { style: [s.tableCell, s.tableCellNo] }, React.createElement(Text, null, 'No.')),
      React.createElement(View, { style: [s.tableCell, s.tableCellName] }, React.createElement(Text, null, 'Name of Form')),
      React.createElement(View, { style: [s.tableCell, s.tableCellQty] }, React.createElement(Text, null, 'Qty')),
    );

    // Table Rows
    const tableRows = tableData.map((row, index) => {
      const formName = row.nameOfForm === 'Other:' 
        ? `Other: ${safeStr(row.otherText, 200)}`
        : safeStr(row.nameOfForm, 200);
      return React.createElement(View, { key: index, style: s.tableRow },
        React.createElement(View, { style: [s.tableCell, s.tableCellNo] }, React.createElement(Text, null, String(index + 1))),
        React.createElement(View, { style: [s.tableCell, s.tableCellName] }, React.createElement(Text, null, formName)),
        React.createElement(View, { style: [s.tableCell, s.tableCellQty] }, React.createElement(Text, null, safeStr(row.qty, 20))),
      );
    });

    // Total Pages Row
    const totalPages = tableData.reduce((sum, row) => {
      const qty = parseFloat(safeStr(row.qty, 20)) || 0;
      return Math.max(0, Math.min(sum + qty, 999999));
    }, 0);
    const totalRow = React.createElement(View, { style: [s.tableRow, { backgroundColor: '#f8f9fa' }] },
      React.createElement(View, { style: [s.tableCell, s.tableCellNo, { flex: 2.3 }] }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Total Pages')),
      React.createElement(View, { style: [s.tableCell, s.tableCellQty] }, React.createElement(Text, { style: { fontWeight: 'bold' } }, String(totalPages))),
    );

    const table = React.createElement(View, { style: s.table }, tableHeader, ...tableRows, totalRow);

    // Production Section
    const productionHeader = React.createElement(View, { style: [s.tableRow, s.tableHeader] },
      React.createElement(View, { style: [s.tableCell, { flex: 1 }] }, React.createElement(Text, null, 'Date')),
      React.createElement(View, { style: [s.tableCell, { flex: 2 }] }, React.createElement(Text, null, 'Note')),
      React.createElement(View, { style: [s.tableCell, { flex: 0.8 }] }, React.createElement(Text, null, 'Status')),
    );
    const productionRows = productionData.map((row, index) => 
      React.createElement(View, { key: index, style: s.tableRow },
        React.createElement(View, { style: [s.tableCell, { flex: 1 }] }, React.createElement(Text, null, safeStr(row.date, 20))),
        React.createElement(View, { style: [s.tableCell, { flex: 2 }] }, React.createElement(Text, null, safeStr(row.note, 500))),
        React.createElement(View, { style: [s.tableCell, { flex: 0.8 }] }, React.createElement(Text, null, safeStr(row.status, 50))),
      )
    );
    const productionTable = React.createElement(View, { style: s.table }, productionHeader, ...productionRows);

        // Today + Next Day Section (side by side)
    const todayNextDaySection = React.createElement(View, { style: s.sideBySideRow },
      // Today (left)
      React.createElement(View, { style: s.sideBySideColumn },
        React.createElement(Text, { style: s.sectionTitle }, 'Today'),
        React.createElement(View, { style: s.sectionContent },
          React.createElement(View, { style: s.sectionRow },
            React.createElement(Text, { style: s.sectionLabel }, 'Add On\'s: '),
            React.createElement(Text, null, safeStr(todayData.addOns, 500)),
          ),
          React.createElement(View, { style: s.sectionRow },
            React.createElement(Text, { style: s.sectionLabel }, 'No Shows: '),
            React.createElement(Text, null, safeStr(todayData.noShows, 500)),
          ),
          React.createElement(View, { style: s.sectionRow },
            React.createElement(Text, { style: s.sectionLabel }, 'Seen: '),
            React.createElement(Text, null, safeStr(todayData.seen, 500)),
          ),
        ),
      ),
      // Next Day (right)
      React.createElement(View, { style: s.sideBySideColumn },
        React.createElement(Text, { style: s.sectionTitle }, 'Next Day'),
        React.createElement(View, { style: s.sectionContent },
          React.createElement(View, { style: s.sectionRow },
            React.createElement(Text, { style: s.sectionLabel }, 'Opener: '),
            React.createElement(Text, null, safeStr(nextDayData.opener, 200)),
          ),
          React.createElement(View, { style: s.sectionRow },
            React.createElement(Text, { style: s.sectionLabel }, 'Closer: '),
            React.createElement(Text, null, safeStr(nextDayData.closer, 200)),
          ),
        ),
      ),
    );

    // Call Log Section
    const callLogSection = React.createElement(View, { style: s.section },
      React.createElement(Text, { style: s.sectionTitle }, 'Call Log'),
      React.createElement(View, { style: s.sectionContent },
        React.createElement(View, { style: s.sectionRow },
          React.createElement(Text, { style: s.sectionLabel }, 'Who called: '),
          React.createElement(Text, null, safeStr(callLogData.whoCalled, 500)),
        ),
        React.createElement(View, { style: s.sectionRow },
          React.createElement(Text, { style: s.sectionLabel }, 'Appointments made: '),
          React.createElement(Text, null, safeStr(callLogData.appointmentsMade, 500)),
        ),
      ),
    );

    // Supervisor Section
    const supervisorSection = React.createElement(View, { style: s.section },
      React.createElement(Text, { style: s.sectionTitle }, 'Office Supervisor/Manager'),
      React.createElement(View, { style: s.sectionContent },
        React.createElement(View, { style: s.sectionRow },
          React.createElement(Text, { style: s.sectionLabel }, 'Supervisor/Manager: '),
          React.createElement(Text, null, safeStr(supervisorData.officeSupervisorManager, 200)),
        ),
        React.createElement(View, { style: s.sectionRow },
          React.createElement(Text, { style: s.sectionLabel }, 'Check out by: '),
          React.createElement(Text, null, safeStr(supervisorData.checkOutBy, 200)),
        ),
      ),
    );

    const footer = React.createElement(View, { style: s.footer }, React.createElement(Text, null, `Generated: ${generatedDate}`));

    return React.createElement(Document, null,
      React.createElement(Page, { size: 'A4', orientation: 'portrait', style: s.page }, 
        header, 
        infoSection1, 
        infoSection2, 
        table, 
        productionTable, 
        todayNextDaySection, 
        callLogSection, 
        supervisorSection, 
        footer
      ),
    );
  }

  // 초기 데이터 생성 함수 (중복 제거)
  const createInitialData = () => {
    const now = new Date();
    const californiaTime = new Date(now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
    const formattedDate = californiaTime.toLocaleDateString('en-US', {
      month: '2-digit',
      day: '2-digit',
      year: 'numeric'
    });

    return {
      formData: {
        date: formattedDate,
        officeTimeCheckIn: '',
        officeName: '',
        timeCheckOut: '',
        name: '',
      },
      tableData: FIXED_FORM_NAMES.map(nameOfForm => ({
        nameOfForm,
        otherText: '',
        qty: ''
      })),
      productionData: [
        { date: '', note: '', status: '' },
        { date: '', note: '', status: '' },
        { date: '', note: '', status: '' }
      ],
      todayData: {
        addOns: '',
        noShows: '',
        seen: ''
      },
      nextDayData: {
        opener: '',
        closer: ''
      },
      callLogData: {
        whoCalled: '',
        appointmentsMade: ''
      },
      supervisorData: {
        officeSupervisorManager: '',
        checkOutBy: ''
      }
    };
  };

  const [loading, setLoading] = useState(false);
  const [autoSaveStatus, setAutoSaveStatus] = useState('');
  const [submitStatus, setSubmitStatus] = useState('');
  const [progress, setProgress] = useState(0);
  const [isUpdatingFromFirebase, setIsUpdatingFromFirebase] = useState(false);
  const [isAuthorized, setIsAuthorized] = useState<boolean | null>(null); // null: 확인 중, true: 인증됨, false: 인증 실패
  const [userOfficesOptions, setuserOfficesOptions] = useState<string[]>([]); // 사용자의 offices 옵션들
 
  const [userSessionId] = useState(() => Math.random().toString(36).substr(2, 9));

  // 마지막 저장된 데이터 추적
  const [lastSavedData, setLastSavedData] = useState({});
  
  // 날짜 상태
  const [faxDate, setFaxDate] = useState(() => {
    const now = new Date();
    const californiaTime = new Date(now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
    return californiaTime.toISOString().split('T')[0];
  });

  // 오피스 선택 상태
  const [selectedOffice, setSelectedOffice] = useState('');

  // 오피스 옵션
  const officeOptions = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

  // 폼 데이터 상태
  const [formData, setFormData] = useState(() => {
    // 캘리포니아 시간으로 현재 날짜 설정
    const now = new Date();
    const californiaTime = new Date(now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
    const formattedDate = californiaTime.toLocaleDateString('en-US', {
      month: '2-digit',
      day: '2-digit',
      year: 'numeric'
    });
    
    return {
      date: formattedDate,
      officeTimeCheckIn: '',
      officeName: '',
      timeCheckOut: '',
      name: '',
    };
  });

  // Production 데이터 상태 (3줄)
  const [productionData, setProductionData] = useState([
    { date: '', note: '', status: '' },
    { date: '', note: '', status: '' },
    { date: '', note: '', status: '' }
  ]);

  // Today 데이터 상태
  const [todayData, setTodayData] = useState({
    addOns: '',
    noShows: '',
    seen: ''
  });

  // Next Day 데이터 상태
  const [nextDayData, setNextDayData] = useState({
    opener: '',
    closer: ''
  });

  // Call Log 데이터 상태
  const [callLogData, setCallLogData] = useState({
    whoCalled: '',
    appointmentsMade: ''
  });

  // Office Supervisor/Manager 데이터 상태
  const [supervisorData, setSupervisorData] = useState({
    officeSupervisorManager: '',
    checkOutBy: ''
  });

  // 테이블 데이터 상태 (19줄) - 고정된 폼 이름으로 초기화
  const [tableData, setTableData] = useState(() => {
    return FIXED_FORM_NAMES.map(nameOfForm => ({
      nameOfForm,
      otherText: '',
      qty: ''
    }));
  });

  // 마지막 저장된 테이블 데이터 추적
  const [lastSavedTableData, setLastSavedTableData] = useState<Array<{ nameOfForm: string; otherText: string; qty: string }>>([]);
  // 마지막 저장된 Production 데이터 추적
  const [lastSavedProductionData, setLastSavedProductionData] = useState<Array<{ date: string; note: string; status: string }>>([]);
  // 마지막 저장된 Today 데이터 추적
  const [lastSavedTodayData, setLastSavedTodayData] = useState<{ addOns: string; noShows: string; seen: string }>({ addOns: '', noShows: '', seen: '' });
  // 마지막 저장된 Next Day 데이터 추적
  const [lastSavedNextDayData, setLastSavedNextDayData] = useState<{ opener: string; closer: string }>({ opener: '', closer: '' });
  // 마지막 저장된 Call Log 데이터 추적
  const [lastSavedCallLogData, setLastSavedCallLogData] = useState<{ whoCalled: string; appointmentsMade: string }>({ whoCalled: '', appointmentsMade: '' });
  // 마지막 저장된 Supervisor 데이터 추적
  const [lastSavedSupervisorData, setLastSavedSupervisorData] = useState<{ officeSupervisorManager: string; checkOutBy: string }>({ officeSupervisorManager: '', checkOutBy: '' });

  // 자동 저장 함수
  const autoSave = useCallback(async () => {
    if (!faxDate || isUpdatingFromFirebase) return;

    // 🔒 보안: 오피스 선택 확인
    if (!selectedOffice) return;

    // 🔒 보안: 오피스 값 검증
    if (!validateOffice(selectedOffice)) {
      return;
    }

    // 🔒 보안: 날짜 형식 검증
    if (formData.date && !validateDate(formData.date)) {
      return;
    }

    // 데이터가 실제로 변경되었는지 확인
    const hasFormChanges = JSON.stringify(formData) !== JSON.stringify(lastSavedData);
    const hasTableChanges = JSON.stringify(tableData) !== JSON.stringify(lastSavedTableData);
    const hasProductionChanges = JSON.stringify(productionData) !== JSON.stringify(lastSavedProductionData);
    const hasTodayChanges = JSON.stringify(todayData) !== JSON.stringify(lastSavedTodayData);
    const hasNextDayChanges = JSON.stringify(nextDayData) !== JSON.stringify(lastSavedNextDayData);
    const hasCallLogChanges = JSON.stringify(callLogData) !== JSON.stringify(lastSavedCallLogData);
    const hasSupervisorChanges = JSON.stringify(supervisorData) !== JSON.stringify(lastSavedSupervisorData);
    if (!hasFormChanges && !hasTableChanges && !hasProductionChanges && !hasTodayChanges && !hasNextDayChanges && !hasCallLogChanges && !hasSupervisorChanges) return;

    try {
      // 🔒 보안: 입력 검증 및 정리
      const validatedFormData = {
        date: validateInput(formData.date, 20),
        officeTimeCheckIn: validateInput(formData.officeTimeCheckIn, 20),
        officeName: validateInput(formData.officeName, 100),
        timeCheckOut: validateInput(formData.timeCheckOut, 20),
        name: validateInput(formData.name, 100),
      };

      const validatedTableData = tableData.map(row => ({
        nameOfForm: validateInput(row.nameOfForm, 200),
        otherText: validateInput(row.otherText, 200),
        qty: validateInput(row.qty, 20),
      }));

      const validatedProductionData = productionData.map(row => ({
        date: validateInput(row.date, 20),
        note: validateInput(row.note, 500),
        status: validateInput(row.status, 50),
      }));

      const validatedTodayData = {
        addOns: validateInput(todayData.addOns, 500),
        noShows: validateInput(todayData.noShows, 500),
        seen: validateInput(todayData.seen, 500),
      };

      const validatedNextDayData = {
        opener: validateInput(nextDayData.opener, 200),
        closer: validateInput(nextDayData.closer, 200),
      };

      const validatedCallLogData = {
        whoCalled: validateInput(callLogData.whoCalled, 500),
        appointmentsMade: validateInput(callLogData.appointmentsMade, 500),
      };

      const validatedSupervisorData = {
        officeSupervisorManager: validateInput(supervisorData.officeSupervisorManager, 200),
        checkOutBy: validateInput(supervisorData.checkOutBy, 200),
      };

      const dataToSave = {
        faxDate,
        selectedOffice,
        ...validatedFormData,
        tableData: validatedTableData,
        productionData: validatedProductionData,
        todayData: validatedTodayData,
        nextDayData: validatedNextDayData,
        callLogData: validatedCallLogData,
        supervisorData: validatedSupervisorData,
        timestamp: new Date().toISOString(),
        autoSaved: true,
        lastUpdatedBy: userSessionId
      };

      const isoDate = convertDateToISO(validatedFormData.date) || faxDate;
      // 🔒 보안: 문서 ID 검증 및 추가 sanitization
      const safeOffice = selectedOffice.replace(/[^a-zA-Z0-9_-]/g, '');
      const safeIsoDate = isoDate.replace(/[^a-zA-Z0-9_-]/g, '');
      // 경로 탐색 공격 방지
      if (safeOffice.includes('..') || safeIsoDate.includes('..')) {
        return;
      }
      const docId = `${safeIsoDate}_${safeOffice}_fax_cover`;
      // 문서 ID 길이 제한 (Firebase 제한: 1500 bytes)
      if (docId.length > 1500) {
        return;
      }
      await setDoc(doc(db, "fax-cover", docId), dataToSave);
      
      // 저장 성공 시 마지막 저장된 데이터 업데이트
      setLastSavedData({ ...formData });
      setLastSavedTableData([...tableData]);
      setLastSavedProductionData([...productionData]);
      setLastSavedTodayData({ ...todayData });
      setLastSavedNextDayData({ ...nextDayData });
      setLastSavedCallLogData({ ...callLogData });
      setLastSavedSupervisorData({ ...supervisorData });
      
    } catch (error) {
      // Auto-save error silently handled
    }
  }, [faxDate, selectedOffice, formData, tableData, productionData, todayData, nextDayData, callLogData, supervisorData, lastSavedData, lastSavedTableData, lastSavedProductionData, lastSavedTodayData, lastSavedNextDayData, lastSavedCallLogData, lastSavedSupervisorData, isUpdatingFromFirebase, userSessionId]);

  // 데이터 변경 시에만 자동 저장
  useEffect(() => {
    // 초기 로드 시에는 저장하지 않음
    if (Object.values(formData).some(value => value !== '') || 
        tableData.some(row => row.nameOfForm || row.qty) ||
        productionData.some(row => row.date || row.note || row.status) ||
        Object.values(todayData).some(value => value !== '') ||
        Object.values(nextDayData).some(value => value !== '') ||
        Object.values(callLogData).some(value => value !== '') ||
        Object.values(supervisorData).some(value => value !== '')) {
      autoSave();
    }
  }, [formData, tableData, productionData, todayData, nextDayData, callLogData, supervisorData]);

  // 데이터 로드
  const loadData = async () => {
    // formData.date와 selectedOffice가 모두 있어야 로드
    if (!formData.date || !selectedOffice) return;
    
    const isoDate = convertDateToISO(formData.date);
    if (!isoDate) return;

    try {
      setSubmitStatus('Loading data...');
      
      // 🔒 보안: 문서 ID 검증
      const safeOffice = selectedOffice.replace(/[^a-zA-Z0-9_-]/g, '');
      const safeIsoDate = isoDate.replace(/[^a-zA-Z0-9_-]/g, '');
      if (safeOffice.includes('..') || safeIsoDate.includes('..')) {
        return;
      }
      const docId = `${safeIsoDate}_${safeOffice}_fax_cover`;
      if (docId.length > 1500) {
        return;
      }
      const docRef = doc(db, "fax-cover", docId);
      const docSnap = await getDoc(docRef);
      
      if (docSnap.exists()) {
        const data = docSnap.data();
        
        // Firebase에서 업데이트되는 동안 자동 저장 방지
        setIsUpdatingFromFirebase(true);
        
        // 날짜가 없으면 캘리포니아 시간으로 설정
        let dateValue = data.date || '';
        if (!dateValue) {
          const now = new Date();
          const californiaTime = new Date(now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
          dateValue = californiaTime.toLocaleDateString('en-US', {
            month: '2-digit',
            day: '2-digit',
            year: 'numeric'
          });
        }
        setFormData(prevData => ({
          ...prevData,
          date: dateValue,
          officeTimeCheckIn: data.officeTimeCheckIn ? convertTo12Hour(data.officeTimeCheckIn) : '',
          officeName: data.officeName || '',
          timeCheckOut: data.timeCheckOut ? convertTo12Hour(data.timeCheckOut) : '',
          name: data.name || '',
        }));

        // 테이블 데이터 로드
        if (data.tableData && Array.isArray(data.tableData)) {
          // 고정된 폼 이름 유지하면서 데이터 복원
          const normalizedTableData = FIXED_FORM_NAMES.map((fixedName, index) => {
            const savedRow = data.tableData[index];
            return {
              nameOfForm: fixedName,
              otherText: savedRow?.otherText || '',
              qty: savedRow?.qty || ''
            };
          });
          // 19개만 유지 (20번째 행 제거)
          const trimmedData = normalizedTableData.slice(0, 19);
          setTableData(trimmedData);
          setLastSavedTableData([...trimmedData]);
        } else {
          // 고정된 폼 이름으로 초기화
          const initial = createInitialData();
          setTableData(initial.tableData);
          setLastSavedTableData(initial.tableData);
        }

        // Production 데이터 로드
        if (data.productionData && Array.isArray(data.productionData)) {
          setProductionData(data.productionData);
          setLastSavedProductionData([...data.productionData]);
        } else {
          const initialProductionData = [
            { date: '', note: '', status: '' },
            { date: '', note: '', status: '' },
            { date: '', note: '', status: '' }
          ];
          setProductionData(initialProductionData);
          setLastSavedProductionData(initialProductionData);
        }

        // Today 데이터 로드
        if (data.todayData) {
          setTodayData({
            addOns: data.todayData.addOns || '',
            noShows: data.todayData.noShows || '',
            seen: data.todayData.seen || ''
          });
          setLastSavedTodayData({
            addOns: data.todayData.addOns || '',
            noShows: data.todayData.noShows || '',
            seen: data.todayData.seen || ''
          });
        } else {
          const initialTodayData = {
            addOns: '',
            noShows: '',
            seen: ''
          };
          setTodayData(initialTodayData);
          setLastSavedTodayData(initialTodayData);
        }

        // Next Day 데이터 로드
        if (data.nextDayData) {
          setNextDayData({
            opener: data.nextDayData.opener || '',
            closer: data.nextDayData.closer || ''
          });
          setLastSavedNextDayData({
            opener: data.nextDayData.opener || '',
            closer: data.nextDayData.closer || ''
          });
        } else {
          const initialNextDayData = {
            opener: '',
            closer: ''
          };
          setNextDayData(initialNextDayData);
          setLastSavedNextDayData(initialNextDayData);
        }

        // Call Log 데이터 로드
        if (data.callLogData) {
          setCallLogData({
            whoCalled: data.callLogData.whoCalled || '',
            appointmentsMade: data.callLogData.appointmentsMade || ''
          });
          setLastSavedCallLogData({
            whoCalled: data.callLogData.whoCalled || '',
            appointmentsMade: data.callLogData.appointmentsMade || ''
          });
        } else {
          const initialCallLogData = {
            whoCalled: '',
            appointmentsMade: ''
          };
          setCallLogData(initialCallLogData);
          setLastSavedCallLogData(initialCallLogData);
        }

        // Supervisor 데이터 로드
        if (data.supervisorData) {
          setSupervisorData({
            officeSupervisorManager: data.supervisorData.officeSupervisorManager || '',
            checkOutBy: data.supervisorData.checkOutBy || ''
          });
          setLastSavedSupervisorData({
            officeSupervisorManager: data.supervisorData.officeSupervisorManager || '',
            checkOutBy: data.supervisorData.checkOutBy || ''
          });
        } else {
          const initialSupervisorData = {
            officeSupervisorManager: '',
            checkOutBy: ''
          };
          setSupervisorData(initialSupervisorData);
          setLastSavedSupervisorData(initialSupervisorData);
        }
        
        // 로드된 데이터를 마지막 저장된 데이터로 설정
        setLastSavedData({
          date: data.date || '',
          officeTimeCheckIn: data.officeTimeCheckIn || '',
          officeName: data.officeName || '',
          timeCheckOut: data.timeCheckOut || '',
          name: data.name || '',
        });
        
        // 짧은 지연 후 Firebase 업데이트 플래그 해제
        setTimeout(() => {
          setIsUpdatingFromFirebase(false);
        }, 100);
        
        setSubmitStatus('Data loaded successfully');
        setTimeout(() => setSubmitStatus(''), 2000);
      } else {
        // 데이터가 없으면 초기화
        const initial = createInitialData();
        setFormData(initial.formData);
        setLastSavedData(initial.formData);
        setTableData(initial.tableData);
        setLastSavedTableData(initial.tableData);
        setProductionData(initial.productionData);
        setLastSavedProductionData(initial.productionData);
        setTodayData(initial.todayData);
        setLastSavedTodayData(initial.todayData);
        setNextDayData(initial.nextDayData);
        setLastSavedNextDayData(initial.nextDayData);
        setCallLogData(initial.callLogData);
        setLastSavedCallLogData(initial.callLogData);
        setSupervisorData(initial.supervisorData);
        setLastSavedSupervisorData(initial.supervisorData);
        setSubmitStatus('');
      }
      
    } catch (error) {
      // 🔒 보안: 에러 메시지에 민감한 정보 노출 방지
      setSubmitStatus('');
    }
  };

  // formData.date 또는 selectedOffice 변경 시 데이터 로드
  useEffect(() => {
    if (formData.date && selectedOffice) {
      const isoDate = convertDateToISO(formData.date);
      if (isoDate) {
        setFaxDate(isoDate);
        // 약간의 지연 후 로드 (faxDate 업데이트 반영)
        setTimeout(() => {
          loadData();
        }, 100);
      }
    }
  }, [formData.date, selectedOffice]);

  // 실시간 데이터 동기화
  useEffect(() => {
    if (!formData.date || !selectedOffice) return;

    const isoDate = convertDateToISO(formData.date) || faxDate;
    if (!isoDate) return;

    // 🔒 보안: 문서 ID 검증
    const safeOffice = selectedOffice.replace(/[^a-zA-Z0-9_-]/g, '');
    const safeIsoDate = isoDate.replace(/[^a-zA-Z0-9_-]/g, '');
    if (safeOffice.includes('..') || safeIsoDate.includes('..')) {
      return;
    }
    const docId = `${safeIsoDate}_${safeOffice}_fax_cover`;
    if (docId.length > 1500) {
      return;
    }
    const docRef = doc(db, "fax-cover", docId);
    
    const unsubscribe = onSnapshot(docRef, (docSnap) => {
      if (docSnap.exists()) {
        const data = docSnap.data();
        
        // Firebase에서 업데이트되는 동안 자동 저장 방지
        setIsUpdatingFromFirebase(true);
        
        setFormData(prevData => {
          // 날짜가 없으면 캘리포니아 시간으로 설정
          let dateValue = data.date || '';
          if (!dateValue) {
            const now = new Date();
            const californiaTime = new Date(now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
            dateValue = californiaTime.toLocaleDateString('en-US', {
              month: '2-digit',
              day: '2-digit',
              year: 'numeric'
            });
          }
          return {
            ...prevData,
            date: dateValue,
            officeTimeCheckIn: data.officeTimeCheckIn ? convertTo12Hour(data.officeTimeCheckIn) : '',
            officeName: data.officeName || '',
            timeCheckOut: data.timeCheckOut ? convertTo12Hour(data.timeCheckOut) : '',
            name: data.name || '',
          };
        });

        // Production 데이터 업데이트
        if (data.productionData && Array.isArray(data.productionData)) {
          setProductionData(data.productionData);
          setLastSavedProductionData([...data.productionData]);
        }

        // Today 데이터 업데이트
        if (data.todayData) {
          setTodayData({
            addOns: data.todayData.addOns || '',
            noShows: data.todayData.noShows || '',
            seen: data.todayData.seen || ''
          });
          setLastSavedTodayData({
            addOns: data.todayData.addOns || '',
            noShows: data.todayData.noShows || '',
            seen: data.todayData.seen || ''
          });
        }

        // Next Day 데이터 업데이트
        if (data.nextDayData) {
          setNextDayData({
            opener: data.nextDayData.opener || '',
            closer: data.nextDayData.closer || ''
          });
          setLastSavedNextDayData({
            opener: data.nextDayData.opener || '',
            closer: data.nextDayData.closer || ''
          });
        }

        // Call Log 데이터 업데이트
        if (data.callLogData) {
          setCallLogData({
            whoCalled: data.callLogData.whoCalled || '',
            appointmentsMade: data.callLogData.appointmentsMade || ''
          });
          setLastSavedCallLogData({
            whoCalled: data.callLogData.whoCalled || '',
            appointmentsMade: data.callLogData.appointmentsMade || ''
          });
        }

        // Supervisor 데이터 업데이트
        if (data.supervisorData) {
          setSupervisorData({
            officeSupervisorManager: data.supervisorData.officeSupervisorManager || '',
            checkOutBy: data.supervisorData.checkOutBy || ''
          });
          setLastSavedSupervisorData({
            officeSupervisorManager: data.supervisorData.officeSupervisorManager || '',
            checkOutBy: data.supervisorData.checkOutBy || ''
          });
        }

        // 테이블 데이터 업데이트
        if (data.tableData && Array.isArray(data.tableData)) {
          // 고정된 폼 이름 유지하면서 데이터 복원
          const normalizedTableData = FIXED_FORM_NAMES.map((fixedName, index) => {
            const savedRow = data.tableData[index];
            return {
              nameOfForm: fixedName,
              otherText: savedRow?.otherText || '',
              qty: savedRow?.qty || ''
            };
          });
          // 19개만 유지 (20번째 행 제거)
          const trimmedData = normalizedTableData.slice(0, 19);
          setTableData(trimmedData);
          setLastSavedTableData([...trimmedData]);
        }
        
        // 실시간 업데이트된 데이터를 마지막 저장된 데이터로 설정
        // 날짜가 없으면 캘리포니아 시간으로 설정
        let savedDateValue = data.date || '';
        if (!savedDateValue) {
          const now = new Date();
          const californiaTime = new Date(now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
          savedDateValue = californiaTime.toLocaleDateString('en-US', {
            month: '2-digit',
            day: '2-digit',
            year: 'numeric'
          });
        }
        setLastSavedData({
          date: savedDateValue,
          officeTimeCheckIn: data.officeTimeCheckIn || '',
          officeName: data.officeName || '',
          timeCheckOut: data.timeCheckOut || '',
          name: data.name || '',
        });
        
        // 짧은 지연 후 Firebase 업데이트 플래그 해제
        setTimeout(() => {
          setIsUpdatingFromFirebase(false);
        }, 100);
        
        // 다른 사용자의 업데이트만 표시 (자신의 업데이트는 제외)
        if (data.timestamp && 
            new Date(data.timestamp).getTime() > Date.now() - 5000 && 
            data.lastUpdatedBy && 
            data.lastUpdatedBy !== userSessionId) {
          setAutoSaveStatus('🔄 Updated from another user');
          setTimeout(() => setAutoSaveStatus(''), 2000);
        }
      }
    }, (error) => {
      setAutoSaveStatus('❌ Error');
      setTimeout(() => setAutoSaveStatus(''), 3000);
    });

    return () => {
      unsubscribe();
    };
  }, [faxDate, selectedOffice, userSessionId]);

  // 컴포넌트 마운트 시 사용자 인증 및 role 확인
  useEffect(() => {
    // Firebase Auth 상태 변경 감지
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      try {
        if (!currentUser) {
          alert('Please log in.');
          setIsAuthorized(false);
          return;
        }

        // Firestore에서 사용자 role 확인
        const userDoc = await getDoc(doc(db, 'users', currentUser.uid));
        if (!userDoc.exists()) {
          alert('User information could not be found.');
          setIsAuthorized(false);
          return;
        }

        const userData = userDoc.data();

        if (userData?.role !== 'MANAGER' && userData?.role !== 'USER') {
          alert('You do not have access to this page.');
          setIsAuthorized(false);
          // 다른 페이지로 리다이렉트하거나 홈으로 이동
          if (typeof window !== 'undefined') {
            window.location.href = '/';
          }
          return;
        }

        setIsAuthorized(true);

        // offices 처리: 배열이거나 단일 값일 수 있음
        if (userData?.offices) {
          const officesArray = Array.isArray(userData.offics) 
            ? userData.offices 
            : [userData.offices];
          
          // officeOptions에 포함된 값들만 필터링
          const validOptions = officesArray.filter((g: string) => officeOptions.includes(g));
          
          if (validOptions.length > 0) {
            setuserOfficesOptions(validOptions);
            // 단일 값이면 자동 선택
            if (validOptions.length === 1) {
              setSelectedOffice(validOptions[0]);
            }
          }
        }
      } catch (error: any) {
        alert('An error occurred while verifying authentication.');
        setIsAuthorized(false);
      }
    });

    // 프로덕션 환경에서 HTTPS 강제 (클라이언트 사이드)
    if (process.env.NODE_ENV === 'production' && 
        typeof window !== 'undefined' && 
        window.location.protocol !== 'https:') {
      // HTTP로 접속한 경우 HTTPS로 리다이렉트
      window.location.href = window.location.href.replace('http:', 'https:');
    }

    // cleanup 함수
    return () => {
      unsubscribe();
    };
  }, []);

  // 데이터 업데이트 함수
  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value } = e.target;
    setFormData((prev) => ({
      ...prev,
      [name]: value,
    }));
  };

  // 테이블 데이터 업데이트 함수
  const handleTableChange = (index: number, field: 'otherText' | 'qty', value: string) => {
    setTableData(prev => {
      const newData = [...prev];
      newData[index] = {
        ...newData[index],
        [field]: value
      };
      return newData;
    });
  };

  // 제출 처리
  const handleSubmit = async () => {
    // 🔒 보안: 오피스 선택 확인
    if (!selectedOffice) {
      alert('Please select an office');
      return;
    }

    // 🔒 보안: 오피스 값 검증
    if (!validateOffice(selectedOffice)) {
      alert('Invalid office value');
      return;
    }

    // 🔒 보안: 날짜 형식 검증
    if (formData.date && !validateDate(formData.date)) {
      alert('Invalid date format');
      return;
    }

    // 확인 다이얼로그
    const confirmed = confirm('Would you like to submit?');
    if (!confirmed) {
      return;
    }

    try {
      setLoading(true);
      setSubmitStatus('Saving...');
      setProgress(10);

      // 🔒 보안: 입력 검증 및 정리
      const validatedFormData = {
        date: validateInput(formData.date, 20),
        officeTimeCheckIn: validateInput(formData.officeTimeCheckIn, 20),
        officeName: validateInput(formData.officeName, 100),
        timeCheckOut: validateInput(formData.timeCheckOut, 20),
        name: validateInput(formData.name, 100),
      };

      const validatedTableData = tableData.map(row => ({
        nameOfForm: validateInput(row.nameOfForm, 200),
        otherText: validateInput(row.otherText, 200),
        qty: validateInput(row.qty, 20),
      }));

      const validatedProductionData = productionData.map(row => ({
        date: validateInput(row.date, 20),
        note: validateInput(row.note, 500),
        status: validateInput(row.status, 50),
      }));

      const validatedTodayData = {
        addOns: validateInput(todayData.addOns, 500),
        noShows: validateInput(todayData.noShows, 500),
        seen: validateInput(todayData.seen, 500),
      };

      const validatedNextDayData = {
        opener: validateInput(nextDayData.opener, 200),
        closer: validateInput(nextDayData.closer, 200),
      };

      const validatedCallLogData = {
        whoCalled: validateInput(callLogData.whoCalled, 500),
        appointmentsMade: validateInput(callLogData.appointmentsMade, 500),
      };

      const validatedSupervisorData = {
        officeSupervisorManager: validateInput(supervisorData.officeSupervisorManager, 200),
        checkOutBy: validateInput(supervisorData.checkOutBy, 200),
      };

      // ISO 날짜 변환 (한 번만 수행)
      const isoDate = convertDateToISO(validatedFormData.date) || faxDate;
      
      // 🔒 보안: 파일명 및 경로 검증
      const safeOffice = selectedOffice.replace(/[^a-zA-Z0-9_-]/g, '');
      const safeIsoDate = isoDate.replace(/[^a-zA-Z0-9_-]/g, '');
      const safeFaxDate = safeStr(validatedFormData.date, 20);

      // 현재 시간 가져오기 (생성 날짜용)
      const now = new Date();
      const californiaTime = new Date(now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
      const generatedDate = californiaTime.toLocaleString('en-US', {
        month: '2-digit',
        day: '2-digit',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        hour12: true
      });

      // 1. PDF 생성
      setSubmitStatus('Generating PDF...');
      setProgress(30);

      try {
        // 클라이언트에서 직접 PDF 생성
        const blob = await pdf(createFaxCoverPDFDocument({
          safeFaxDate,
          safeSelectedOffice: safeOffice,
          formData: validatedFormData,
          tableData: validatedTableData,
          productionData: validatedProductionData,
          todayData: validatedTodayData,
          nextDayData: validatedNextDayData,
          callLogData: validatedCallLogData,
          supervisorData: validatedSupervisorData,
          generatedDate,
        })).toBlob();

        setSubmitStatus('Processing...');
        setProgress(60);

        // PDF를 Firebase Storage에 저장
        setSubmitStatus('Saving...');
        setProgress(70);
        try {
          const storage = getStorage();
          const tsNow = new Date();
          const tsLaTime = new Date(tsNow.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
          let tsHours = tsLaTime.getHours();
          const tsMinutes = tsLaTime.getMinutes();
          const tsAmpm = tsHours >= 12 ? 'pm' : 'am';
          tsHours = tsHours % 12;
          tsHours = tsHours ? tsHours : 12;
          const timeStamp = `${tsHours}${tsMinutes.toString().padStart(2, '0')}${tsAmpm}`;
          
          const filename = `1) ${safeIsoDate}_${safeOffice || 'Unknown'}_End of Day Fax Cover_${timeStamp}.pdf`;
          const storageRef = ref(storage, `endofday-pdfs/${safeOffice}/${safeIsoDate}/${filename}`);
          
          // PDF 업로드
          await uploadBytes(storageRef, blob);
          
          setSubmitStatus('✅ Complete!');
        } catch (storageError: any) {
          // 🔒 보안: 에러 메시지에 민감한 정보 노출 방지
          alert('An error occurred while submitting. Please try again.');
          setSubmitStatus('❌ Submission failed. Please try again.');
        }
        
        // 2. 데이터 삭제
        setSubmitStatus('Cleaning up...');
        setProgress(80);
        // 🔒 보안: 문서 ID 검증 (위에서 이미 safeOffice, safeIsoDate 생성됨)
        const docId = `${safeIsoDate}_${safeOffice}_fax_cover`;
        
        // 자동 저장 방지 (삭제 후 재생성 방지)
        setIsUpdatingFromFirebase(true);
        
        try {
          await deleteDoc(doc(db, "fax-cover", docId));
        } catch (deleteError) {
          // 삭제 실패해도 계속 진행
        }
        
        // 3. 폼 초기화
        const initial = createInitialData();
        setFormData(initial.formData);
        setTableData(initial.tableData);
        setProductionData(initial.productionData);
        setTodayData(initial.todayData);
        setNextDayData(initial.nextDayData);
        setCallLogData(initial.callLogData);
        setSupervisorData(initial.supervisorData);

        // 마지막 저장된 데이터도 초기화하여 자동 저장 방지
        setLastSavedData(initial.formData);
        setLastSavedTableData(initial.tableData);
        setLastSavedProductionData(initial.productionData);
        setLastSavedTodayData(initial.todayData);
        setLastSavedNextDayData(initial.nextDayData);
        setLastSavedCallLogData(initial.callLogData);
        setLastSavedSupervisorData(initial.supervisorData);

        setSubmitStatus('Complete!');
        setProgress(100);
        
        // 자동 저장 재활성화 (약간의 지연 후)
        setTimeout(() => {
          setIsUpdatingFromFirebase(false);
        }, 2000);
        
        // 2초 후 모달 닫기
        setTimeout(() => {
          setLoading(false);
          setSubmitStatus('');
          setProgress(0);
        }, 2000);

      } catch (pdfError: any) {
        // 🔒 보안: 에러 메시지에 민감한 정보 노출 방지
        throw new Error('An error occurred while submitting. Please try again.');
      }
    } catch (error: any) {
      // 🔒 보안: 에러 메시지에 민감한 정보 노출 방지
      setSubmitStatus('❌ Submission failed. Please try again.');
      setProgress(0);
      setTimeout(() => {
        setLoading(false);
        setSubmitStatus('');
      }, 3000);
    }
  };

  // 스타일 정의
  const styles: {
    body: React.CSSProperties;
    container: React.CSSProperties;
    header: React.CSSProperties;
    subtitle: React.CSSProperties;
    formGroup: React.CSSProperties;
    label: React.CSSProperties;
    input: React.CSSProperties;
    submitButton: React.CSSProperties;
    statusMessage: React.CSSProperties;
    autoSaveStatus: React.CSSProperties;
    table: React.CSSProperties;
    th: React.CSSProperties;
    td: React.CSSProperties;
  } = {
    body: {
      fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
      padding: '20px',
      background: 'linear-gradient(135deg, #f5f7fa 0%, #e8ecf1 100%)',
      color: '#2c3e50',
      lineHeight: '1.6',
      minHeight: '100vh'
    },
    container: {
      maxWidth: '67%',
      width: '67%',
      margin: '20px auto',
      padding: '30px',
      backgroundColor: 'white',
      borderRadius: '8px',
      boxShadow: '0 4px 8px rgba(0, 0, 0, 0.1)',
      position: 'relative'
    },
    header: {
      color: '#2c3e50',
      textAlign: 'center',
      marginBottom: '10px',
      paddingBottom: '10px',
      borderBottom: '2px solid #e9ecef',
      fontSize: '2em',
      fontWeight: 'bold'
    },
    subtitle: {
      fontSize: '0.875rem',
      color: '#666',
      textAlign: 'center',
      marginBottom: '30px',
      fontStyle: 'italic'
    },
    formGroup: {
      marginBottom: '25px'
    },
    label: {
      display: 'block',
      fontWeight: '500',
      color: '#2c3e50',
      marginBottom: '8px'
    },
    input: {
      width: '100%',
      padding: '10px 12px',
      fontSize: '14px',
      border: '1px solid #e9ecef',
      borderRadius: '4px',
      backgroundColor: 'white',
      boxSizing: 'border-box',
      transition: 'border-color 0.2s'
    },
    submitButton: {
      display: 'block',
      width: '150px',
      margin: '30px auto 0 auto',
      padding: '12px 20px',
      backgroundColor: '#3498db',
      color: 'white',
      border: 'none',
      borderRadius: '4px',
      cursor: 'pointer',
      fontSize: '16px',
      transition: 'background-color 0.2s'
    },
    statusMessage: {
      marginTop: '15px',
      fontWeight: 'bold',
      textAlign: 'center',
      padding: '10px',
      borderRadius: '4px'
    },
    autoSaveStatus: {
      position: 'absolute',
      top: '10px',
      right: '10px',
      padding: '8px 16px',
      backgroundColor: '#51cf66',
      color: 'white',
      borderRadius: '20px',
      fontSize: '14px',
      fontWeight: 'bold',
      boxShadow: '0 2px 8px rgba(0,0,0,0.2)',
      zIndex: 1000
    },
    table: {
      borderCollapse: 'collapse',
      width: '100%',
      marginTop: '20px',
      backgroundColor: 'white',
      boxShadow: '0 1px 3px rgba(0, 0, 0, 0.1)'
    },
    th: {
      border: '1px solid #e9ecef',
      padding: '10px',
      textAlign: 'left',
      verticalAlign: 'top',
      backgroundColor: '#2c3e50',
      color: 'white',
      fontWeight: '500'
    },
    td: {
      border: '1px solid #e9ecef',
      padding: '10px',
      textAlign: 'left',
      verticalAlign: 'top'
    }
  };

  // 제출 중 브라우저 네비게이션 방지
  useEffect(() => {
    const handleBeforeUnload = (e: BeforeUnloadEvent) => {
      if (loading) {
        e.preventDefault();
        e.returnValue = '';
      }
    };

    const handlePopState = (e: PopStateEvent) => {
      if (loading) {
        e.preventDefault();
        window.history.pushState(null, '', window.location.href);
      }
    };

    if (loading) {
      window.addEventListener('beforeunload', handleBeforeUnload);
      window.addEventListener('popstate', handlePopState);
      window.history.pushState(null, '', window.location.href);
    }

    return () => {
      window.removeEventListener('beforeunload', handleBeforeUnload);
      window.removeEventListener('popstate', handlePopState);
    };
  }, [loading]);

  // 인증 확인 중이거나 인증 실패 시 로딩 화면 표시
  if (isAuthorized === null) {
    return (
      <div style={{
        display: 'flex',
        justifyContent: 'center',
        alignItems: 'center',
        height: '100vh',
        background: `
          radial-gradient(circle at 10% 20%, rgba(120, 200, 255, 0.1) 0%, transparent 50%),
          radial-gradient(circle at 90% 80%, rgba(255, 182, 193, 0.1) 0%, transparent 50%),
          radial-gradient(circle at 50% 50%, rgba(144, 238, 144, 0.05) 0%, transparent 50%),
          linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%)
        `,
        fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
      } as React.CSSProperties}>
        <div style={{ textAlign: 'center' } as React.CSSProperties}>
          <div style={{ fontSize: '24px', marginBottom: '20px' } as React.CSSProperties}>🔐</div>
          <div style={{ fontSize: '18px', color: '#2c3e50' } as React.CSSProperties}>Verifying authentication...</div>
        </div>
      </div>
    );
  }

  if (isAuthorized === false) {
    return (
      <div style={{
        display: 'flex',
        justifyContent: 'center',
        alignItems: 'center',
        height: '100vh',
        background: `
          radial-gradient(circle at 10% 20%, rgba(120, 200, 255, 0.1) 0%, transparent 50%),
          radial-gradient(circle at 90% 80%, rgba(255, 182, 193, 0.1) 0%, transparent 50%),
          radial-gradient(circle at 50% 50%, rgba(144, 238, 144, 0.05) 0%, transparent 50%),
          linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%)
        `,
        fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
      } as React.CSSProperties}>
        <div style={{ textAlign: 'center' } as React.CSSProperties}>
          <div style={{ fontSize: '24px', marginBottom: '20px' } as React.CSSProperties}>🚫</div>
          <div style={{ fontSize: '18px', color: '#d32f2f', marginBottom: '10px' } as React.CSSProperties}>You do not have access to this page.</div>
          <div style={{ fontSize: '14px', color: '#666' } as React.CSSProperties}>You do not have access to this page.</div>
        </div>
      </div>
    );
  }

  return (
    <>
      <style>{`
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
      `}</style>
      <div style={styles.body}>
      {/* 로딩 모달 */}
      {loading && (
        <div style={{
          position: "fixed",
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          backgroundColor: "rgba(0, 0, 0, 0.7)",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          zIndex: 9999
        }}>
          <div style={{
            backgroundColor: "white",
            padding: "40px",
            borderRadius: "12px",
            textAlign: "center",
            boxShadow: "0 10px 30px rgba(0, 0, 0, 0.3)",
            maxWidth: "400px",
            width: "90%"
          }}>
            <div style={{
              border: "4px solid #f3f3f3",
              borderTop: "4px solid #4a90e2",
              borderRadius: "50%",
              width: "50px",
              height: "50px",
              animation: "spin 1s linear infinite",
              margin: "0 auto 20px"
            }}></div>
            <h3 style={{
              color: "#333",
              fontSize: "1.2rem",
              fontWeight: "600",
              margin: "0 0 10px 0"
            }}>
              {submitStatus || "Processing..."}
            </h3>
            <p style={{
              color: "#666",
              fontSize: "0.9rem",
              margin: "0 0 20px 0",
              lineHeight: "1.4"
            }}>
              {!submitStatus && 'Processing... Please wait'}
            </p>
            {/* 진행률 바 */}
            <div style={{
              width: "100%",
              backgroundColor: "#e9ecef",
              borderRadius: "10px",
              overflow: "hidden",
              marginBottom: "20px"
            }}>
              <div style={{
                width: `${progress}%`,
                height: "8px",
                backgroundColor: "#4a90e2",
                borderRadius: "10px",
                transition: "width 0.3s ease",
                background: "linear-gradient(90deg, #4a90e2, #357abd)"
              }}></div>
            </div>
            <p style={{
              color: "#495057",
              fontSize: "0.8rem",
              margin: "0 0 20px 0",
              fontWeight: "500"
            }}>
              {progress}% Complete
            </p>
            <div style={{
              backgroundColor: "#f8f9fa",
              padding: "15px",
              borderRadius: "8px",
              border: "1px solid #e9ecef"
            }}>
              <p style={{
                color: "#495057",
                fontSize: "0.8rem",
                margin: 0,
                fontWeight: "500"
              }}>
                ⚠️ Please do not close.
              </p>
            </div>
          </div>
        </div>
      )}

      <div style={styles.container}>
        {/* 자동 저장 상태 표시 */}
        {autoSaveStatus && (
          <div style={styles.autoSaveStatus}>
            {autoSaveStatus}
          </div>
        )}

        <h2 style={styles.header}>End of Day Fax Cover</h2>
        <p style={styles.subtitle}>(Check out only when leaving the office)</p>

        {/* 날짜 및 오피스 선택 */}
        <div style={{ display: 'flex', gap: '20px', marginBottom: '20px', flexWrap: 'wrap' }}>
          <div style={{ ...styles.formGroup, flex: '1', minWidth: '200px' }}>
            <label style={styles.label} htmlFor="faxDate">Date:</label>
            <input
              type="text"
              id="faxDate"
              value={formData.date}
              onChange={(e) => {
                const validated = validateInput(e.target.value, 20);
                setFormData(prev => ({ ...prev, date: validated }));
              }}
              onBlur={(e) => {
                // 포커스를 잃을 때 데이터 로드
                if (formData.date && selectedOffice) {
                  const isoDate = convertDateToISO(formData.date);
                  if (isoDate) {
                    setFaxDate(isoDate);
                  }
                }
              }}
              maxLength={20}
              style={styles.input}
            />
          </div>
          {/* offices 옵션이 있는 경우에만 Office 표시 */}
          {userOfficesOptions.length > 0 && (
            <div style={{ ...styles.formGroup, flex: '1', minWidth: '200px' }}>
              <label style={styles.label} htmlFor="selectedOffice">Office:</label>
              {userOfficesOptions.length === 1 ? (
                <div style={{
                  ...styles.input,
                  display: 'flex',
                  alignItems: 'center',
                  backgroundColor: '#e9ecef',
                  fontWeight: '600',
                  color: '#2c3e50'
                }}>
                  {selectedOffice}
                </div>
              ) : (
                <select
                  id="selectedOffice"
                  value={selectedOffice}
                  onChange={(e) => setSelectedOffice(e.target.value)}
                  style={styles.input}
                >
                  <option value="">-- Select Office --</option>
                  {userOfficesOptions.map(office => (
                    <option key={office} value={office}>{office}</option>
                  ))}
                </select>
              )}
            </div>
          )}
        </div>

        {/* Time Check In, Name, Time Check Out, Name - 한 줄에 */}
        {selectedOffice && (
        <form style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap' }}>
            <div style={{ ...styles.formGroup, flex: '1', minWidth: '150px' }}>
              <label htmlFor="officeTimeCheckIn" style={styles.label}>
                Time Check In
              </label>
              <input
                type="time"
                id="officeTimeCheckIn"
                name="officeTimeCheckIn"
                value={(() => {
                  // 12시간제를 24시간제로 변환 (time input용)
                  if (!formData.officeTimeCheckIn) return '';
                  const parsed = parseTime(formData.officeTimeCheckIn);
                  if (!parsed.hour || !parsed.minute) return '';
                  let hour24 = parseInt(parsed.hour, 10);
                  if (parsed.ampm === 'PM' && hour24 !== 12) hour24 += 12;
                  if (parsed.ampm === 'AM' && hour24 === 12) hour24 = 0;
                  return `${hour24.toString().padStart(2, '0')}:${parsed.minute}`;
                })()}
                onChange={(e) => {
                  // 24시간제를 12시간제로 변환하여 저장
                  const time24 = e.target.value;
                  if (!time24) {
                    setFormData(prev => ({ ...prev, officeTimeCheckIn: '' }));
                    return;
                  }
                  const [hour24, minute] = time24.split(':');
                  let hour12 = parseInt(hour24, 10);
                  const ampm = hour12 >= 12 ? 'PM' : 'AM';
                  if (hour12 === 0) hour12 = 12;
                  else if (hour12 > 12) hour12 -= 12;
                  const validated = validateInput(`${hour12.toString().padStart(2, '0')}:${minute} ${ampm}`, 20);
                  setFormData(prev => ({ ...prev, officeTimeCheckIn: validated }));
                }}
                maxLength={20}
                style={styles.input}
              />
            </div>

            <div style={{ ...styles.formGroup, flex: '1', minWidth: '150px' }}>
              <label htmlFor="officeName" style={styles.label}>
                Name
              </label>
              <input
                type="text"
                id="officeName"
                name="officeName"
                value={formData.officeName}
                onChange={handleChange}
                style={styles.input}
              />
            </div>

            <div style={{ ...styles.formGroup, flex: '1', minWidth: '150px' }}>
              <label htmlFor="timeCheckOut" style={styles.label}>
                Time Check Out
              </label>
              <input
                type="time"
                id="timeCheckOut"
                name="timeCheckOut"
                value={(() => {
                  // 12시간제를 24시간제로 변환 (time input용)
                  if (!formData.timeCheckOut) return '';
                  const parsed = parseTime(formData.timeCheckOut);
                  if (!parsed.hour || !parsed.minute) return '';
                  let hour24 = parseInt(parsed.hour, 10);
                  if (parsed.ampm === 'PM' && hour24 !== 12) hour24 += 12;
                  if (parsed.ampm === 'AM' && hour24 === 12) hour24 = 0;
                  return `${hour24.toString().padStart(2, '0')}:${parsed.minute}`;
                })()}
                onChange={(e) => {
                  // 24시간제를 12시간제로 변환하여 저장
                  const time24 = e.target.value;
                  if (!time24) {
                    setFormData(prev => ({ ...prev, timeCheckOut: '' }));
                    return;
                  }
                  const [hour24, minute] = time24.split(':');
                  let hour12 = parseInt(hour24, 10);
                  const ampm = hour12 >= 12 ? 'PM' : 'AM';
                  if (hour12 === 0) hour12 = 12;
                  else if (hour12 > 12) hour12 -= 12;
                  const validated = validateInput(`${hour12.toString().padStart(2, '0')}:${minute} ${ampm}`, 20);
                  setFormData(prev => ({ ...prev, timeCheckOut: validated }));
                }}
                maxLength={20}
                style={styles.input}
              />
            </div>

            <div style={{ ...styles.formGroup, flex: '1', minWidth: '150px' }}>
              <label htmlFor="name" style={styles.label}>
                Name
              </label>
              <input
                type="text"
                id="name"
                name="name"
                value={formData.name}
                onChange={handleChange}
                style={styles.input}
              />
            </div>
          </div>
        </form>
        )}

        {/* 테이블 */}
        {selectedOffice && (
        <table style={styles.table}>
          <thead>
            <tr>
              <th style={{ ...styles.th, width: '60px' }}>No.</th>
              <th style={styles.th}>Name of Form</th>
              <th style={styles.th}>Qty</th>
            </tr>
          </thead>
          <tbody>
            {tableData.map((row, index) => (
              <tr key={index}>
                <td style={{ ...styles.td, textAlign: 'center', fontWeight: 'bold' }}>
                  {index + 1}
                </td>
                <td style={styles.td}>
                  {row.nameOfForm === 'Other:' ? (
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                      {/* "Other:" 텍스트 */}
                      <span style={{
                        fontSize: '13px',
                        color: '#2c3e50',
                        fontWeight: '500',
                        whiteSpace: 'nowrap'
                      }}>
                        Other:
                      </span>
                      {/* 입력 필드 */}
                    <input
                      type="text"
                      value={row.otherText}
                      onChange={(e) => {
                        const validated = validateInput(e.target.value, 200);
                        handleTableChange(index, 'otherText', validated);
                      }}
                      maxLength={200}
                      style={{
                        ...styles.input,
                        margin: 0,
                        fontSize: '13px',
                        padding: '6px 8px',
                        flex: 1,
                        boxSizing: 'border-box'
                      }}
                    />
                    </div>
                  ) : (
                    <div style={{
                      padding: '6px 8px',
                      fontSize: '13px',
                      color: '#2c3e50',
                      fontWeight: row.nameOfForm ? '500' : 'normal'
                    }}>
                      {row.nameOfForm || '\u00A0'}
                    </div>
                  )}
                </td>
                <td style={styles.td}>
                  <input
                    type="text"
                    value={row.qty}
                    onChange={(e) => {
                      const validated = validateInput(e.target.value, 20);
                      handleTableChange(index, 'qty', validated);
                    }}
                    maxLength={20}
                    style={{
                      ...styles.input,
                      margin: 0,
                      fontSize: '13px',
                      padding: '6px 8px',
                      width: '100%',
                      boxSizing: 'border-box'
                    }}
                  />
                </td>
              </tr>
            ))}
            {/* Total Pages 행 */}
            <tr style={{ backgroundColor: '#f8f9fa', fontWeight: 'bold' }}>
              <td style={{ ...styles.td, textAlign: 'center' }} colSpan={2}>
                Total Pages
              </td>
              <td style={{ ...styles.td, textAlign: 'center', fontWeight: 'bold', fontSize: '16px' }}>
                {tableData.reduce((sum, row) => {
                  const qty = parseFloat(row.qty) || 0;
                  return sum + qty;
                }, 0)}
              </td>
            </tr>
          </tbody>
        </table>
        )}

        {/* Production 섹션 */}
        {selectedOffice && (
        <div style={{ marginTop: '40px', paddingTop: '20px', borderTop: '2px solid #e9ecef' }}>
          <h3 style={{
            color: '#2c3e50',
            fontSize: '1.5em',
            fontWeight: 'bold',
            marginBottom: '20px'
          }}>
            Production
          </h3>

          {/* 3줄: 각 줄마다 Date, Note, Status */}
          {productionData.map((row, index) => (
            <div key={index} style={{ ...styles.formGroup, marginBottom: '20px' }}>
              <div style={{ display: 'flex', gap: '20px', alignItems: 'flex-start', flexWrap: 'wrap' }}>
                {/* Date 입력 */}
                <div style={{ flex: '1', minWidth: '150px' }}>
                  <label style={styles.label}>Date</label>
                  <input
                    type="date"
                    value={row.date}
                    onChange={(e) => {
                      const newData = [...productionData];
                      newData[index] = { ...newData[index], date: e.target.value };
                      setProductionData(newData);
                    }}
                    style={styles.input}
                  />
                </div>

                {/* Note 입력 */}
                <div style={{ flex: '2', minWidth: '200px' }}>
                  <label style={styles.label}>Note</label>
                  <input
                    type="text"
                    value={row.note}
                    onChange={(e) => {
                      const newData = [...productionData];
                      newData[index] = { ...newData[index], note: e.target.value };
                      setProductionData(newData);
                    }}
                    style={styles.input}
                  />
                </div>

                {/* Status 선택 */}
                <div style={{ flex: '1', minWidth: '150px' }}>
                  <label style={styles.label}>Status</label>
                  <div style={{ display: 'flex', gap: '15px', alignItems: 'center', marginTop: '8px' }}>
                    <label style={{ display: 'flex', alignItems: 'center', gap: '6px', cursor: 'pointer' }}>
                      <input
                        type="radio"
                        name={`productionStatus_${index}`}
                        value="final"
                        checked={row.status === 'final'}
                        onChange={(e) => {
                          const newData = [...productionData];
                          newData[index] = { ...newData[index], status: e.target.value };
                          setProductionData(newData);
                        }}
                        style={{ cursor: 'pointer' }}
                      />
                      <span style={{ fontSize: '14px', color: '#2c3e50' }}>Final</span>
                    </label>
                    <label style={{ display: 'flex', alignItems: 'center', gap: '6px', cursor: 'pointer' }}>
                      <input
                        type="radio"
                        name={`productionStatus_${index}`}
                        value="not final"
                        checked={row.status === 'not final'}
                        onChange={(e) => {
                          const newData = [...productionData];
                          newData[index] = { ...newData[index], status: e.target.value };
                          setProductionData(newData);
                        }}
                        style={{ cursor: 'pointer' }}
                      />
                      <span style={{ fontSize: '14px', color: '#2c3e50' }}>Not Final</span>
                    </label>
                  </div>
                </div>
              </div>
            </div>
          ))}
        </div>
        )}

        {/* Today 섹션 */}
        {selectedOffice && (
        <div style={{ marginTop: '40px', paddingTop: '20px', borderTop: '2px solid #e9ecef' }}>
          <h3 style={{
            color: '#2c3e50',
            fontSize: '1.5em',
            fontWeight: 'bold',
            marginBottom: '20px'
          }}>
            Today
          </h3>

          {/* 첫 번째 줄: Add On's */}
          <div style={styles.formGroup}>
            <label htmlFor="addOns" style={styles.label}>
              Add On's
            </label>
            <input
              type="text"
              id="addOns"
              name="addOns"
              value={todayData.addOns}
              onChange={(e) => {
                const validated = validateInput(e.target.value, 500);
                setTodayData(prev => ({ ...prev, addOns: validated }));
              }}
              maxLength={500}
              style={styles.input}
            />
          </div>

          {/* 두 번째 줄: No Shows */}
          <div style={styles.formGroup}>
            <label htmlFor="noShows" style={styles.label}>
              No Shows
            </label>
            <input
              type="text"
              id="noShows"
              name="noShows"
              value={todayData.noShows}
              onChange={(e) => {
                const validated = validateInput(e.target.value, 500);
                setTodayData(prev => ({ ...prev, noShows: validated }));
              }}
              style={styles.input}
            />
          </div>

          {/* 세 번째 줄: Seen */}
          <div style={styles.formGroup}>
            <label htmlFor="seen" style={styles.label}>
              Seen
            </label>
            <input
              type="text"
              id="seen"
              name="seen"
              value={todayData.seen}
              onChange={(e) => {
                const validated = validateInput(e.target.value, 500);
                setTodayData(prev => ({ ...prev, seen: validated }));
              }}
              maxLength={500}
              style={styles.input}
            />
          </div>
        </div>
        )}

        {/* Next Day 섹션 */}
        {selectedOffice && (
        <div style={{ marginTop: '40px', paddingTop: '20px', borderTop: '2px solid #e9ecef' }}>
          <h3 style={{
            color: '#2c3e50',
            fontSize: '1.5em',
            fontWeight: 'bold',
            marginBottom: '20px'
          }}>
            Next Day
          </h3>

          {/* Opener 입력 */}
          <div style={styles.formGroup}>
            <label htmlFor="opener" style={styles.label}>
              Opener
            </label>
            <input
              type="text"
              id="opener"
              name="opener"
              value={nextDayData.opener}
              onChange={(e) => {
                const validated = validateInput(e.target.value, 200);
                setNextDayData(prev => ({ ...prev, opener: validated }));
              }}
              style={styles.input}
            />
          </div>

          {/* Closer 입력 */}
          <div style={styles.formGroup}>
            <label htmlFor="closer" style={styles.label}>
              Closer
            </label>
            <input
              type="text"
              id="closer"
              name="closer"
              value={nextDayData.closer}
              onChange={(e) => {
                const validated = validateInput(e.target.value, 200);
                setNextDayData(prev => ({ ...prev, closer: validated }));
              }}
              maxLength={200}
              style={styles.input}
            />
          </div>
        </div>
        )}

        {/* Call Log 섹션 */}
        {selectedOffice && (
        <div style={{ marginTop: '40px', paddingTop: '20px', borderTop: '2px solid #e9ecef' }}>
          <h3 style={{
            color: '#2c3e50',
            fontSize: '1.5em',
            fontWeight: 'bold',
            marginBottom: '20px'
          }}>
            Call Log
          </h3>

          {/* Who called 입력 */}
          <div style={styles.formGroup}>
            <label htmlFor="whoCalled" style={styles.label}>
              Who called
            </label>
            <input
              type="text"
              id="whoCalled"
              name="whoCalled"
              value={callLogData.whoCalled}
              onChange={(e) => {
                const validated = validateInput(e.target.value, 500);
                setCallLogData(prev => ({ ...prev, whoCalled: validated }));
              }}
              style={styles.input}
            />
          </div>

          {/* How many appointments made 입력 */}
          <div style={styles.formGroup}>
            <label htmlFor="appointmentsMade" style={styles.label}>
              How many appointments made
            </label>
            <input
              type="text"
              id="appointmentsMade"
              name="appointmentsMade"
              value={callLogData.appointmentsMade}
              onChange={(e) => {
                const validated = validateInput(e.target.value, 500);
                setCallLogData(prev => ({ ...prev, appointmentsMade: validated }));
              }}
              maxLength={500}
              style={styles.input}
            />
          </div>
        </div>
        )}

        {/* Office Supervisor/Manager 섹션 */}
        {selectedOffice && (
        <div style={{ marginTop: '40px', paddingTop: '20px', borderTop: '2px solid #e9ecef' }}>
          <h3 style={{
            color: '#2c3e50',
            fontSize: '1.5em',
            fontWeight: 'bold',
            marginBottom: '20px'
          }}>
            (For Corporate Use Only)
          </h3>

          {/* Office Supervisor/Manager 이름 입력 */}
          <div style={styles.formGroup}>
            <label htmlFor="officeSupervisorManager" style={styles.label}>
              Office Supervisor/Manager
            </label>
            <input
              type="text"
              id="officeSupervisorManager"
              name="officeSupervisorManager"
              value={supervisorData.officeSupervisorManager}
              onChange={(e) => {
                const validated = validateInput(e.target.value, 200);
                setSupervisorData(prev => ({ ...prev, officeSupervisorManager: validated }));
              }}
              maxLength={200}
              style={styles.input}
            />
          </div>

          {/* Check out by 입력 */}
          <div style={styles.formGroup}>
            <label htmlFor="checkOutBy" style={styles.label}>
              Check out by
            </label>
            <input
              type="text"
              id="checkOutBy"
              name="checkOutBy"
              value={supervisorData.checkOutBy}
              onChange={(e) => {
                const validated = validateInput(e.target.value, 200);
                setSupervisorData(prev => ({ ...prev, checkOutBy: validated }));
              }}
              maxLength={200}
              style={styles.input}
            />
          </div>
        </div>
        )}

        {/* 제출 버튼 */}
        {selectedOffice && (
        <div style={{ textAlign: 'center', marginTop: '20px' }}>
          <button
            onClick={handleSubmit}
            disabled={loading}
            style={{
              ...styles.submitButton,
              backgroundColor: loading ? '#bdc3c7' : '#3498db',
              cursor: loading ? 'not-allowed' : 'pointer'
            }}
          >
            {loading ? 'Submitting...' : 'Button for Corporate'}
          </button>
        </div>
        )}

        {/* 상태 메시지 */}
        {submitStatus && (
          <div style={{
            ...styles.statusMessage,
            backgroundColor: submitStatus.includes('failed') || submitStatus.includes('Error') ? '#f8d7da' : 
                           submitStatus.includes('successfully') || submitStatus.includes('Complete') ? '#d4edda' : '#d1ecf1',
            color: submitStatus.includes('failed') || submitStatus.includes('Error') ? '#721c24' : 
                   submitStatus.includes('successfully') || submitStatus.includes('Complete') ? '#155724' : '#0c5460',
            border: submitStatus.includes('failed') || submitStatus.includes('Error') ? '1px solid #f5c6cb' : 
                    submitStatus.includes('successfully') || submitStatus.includes('Complete') ? '1px solid #c3e6cb' : '1px solid #bee5eb'
          }}>
            {submitStatus}
          </div>
        )}

      </div>
    </div>
    </>
  );
}


