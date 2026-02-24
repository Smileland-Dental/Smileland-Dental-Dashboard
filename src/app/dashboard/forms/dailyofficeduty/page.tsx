'use client'

import React, { useState, useEffect, useCallback, useRef } from "react";
import { doc, setDoc, getDoc, deleteDoc, onSnapshot } from "firebase/firestore";
import { getStorage, ref, uploadBytes } from 'firebase/storage';
import { db, auth } from "@/lib/firebase.config";
// Firebase 인증 직접 사용
import { onAuthStateChanged } from 'firebase/auth';
import { pdf, Document, Page, View, Text, StyleSheet } from '@react-pdf/renderer';

// 🔒 보안: 입력 검증 함수
function safeStr(v: unknown, max: number): string {
  if (v == null) return '';
  return String(v).trim().slice(0, max).replace(/[<>]/g, '');
}


// Daily Office Duty PDF 생성 함수
function createDailyOfficeDutyPDFDocument(props: {
  safeDutyDate: string;
  safeSelectedOffice: string;
  safeDutyData: { [key: string]: string };
  generatedDate: string;
}) {
  const { safeDutyDate, safeSelectedOffice, safeDutyData, generatedDate } = props;

  const pdfStyles = StyleSheet.create({
    page: { padding: 20, fontFamily: 'Helvetica', fontSize: 9 },
    header: { marginBottom: 15, borderBottomWidth: 2, borderColor: '#333', paddingBottom: 8, alignItems: 'center' },
    headerTitle: { fontSize: 16, fontWeight: 'bold', marginBottom: 4 },
    infoSection: { marginBottom: 15, padding: 8, borderWidth: 1, borderColor: '#ccc', flexDirection: 'row', justifyContent: 'space-between', flexWrap: 'wrap' },
    infoItem: { fontSize: 9, marginBottom: 4 },
    table: { marginTop: 10 },
    tableRow: { flexDirection: 'row', borderBottomWidth: 0.5, borderColor: '#333' },
    tableHeader: { flexDirection: 'row', borderBottomWidth: 1, borderColor: '#333', backgroundColor: '#f0f0f0', fontWeight: 'bold' },
    tableCellNo: { padding: 4, fontSize: 8, flex: 0.3, borderRightWidth: 0.5, borderColor: '#333', justifyContent: 'center', alignItems: 'center' },
    tableCellDesc: { padding: 4, fontSize: 7, flex: 2.5, borderRightWidth: 0.5, borderColor: '#333', justifyContent: 'flex-start', alignItems: 'flex-start' },
    tableCellDetails: { padding: 4, fontSize: 7, flex: 1.5, borderRightWidth: 0.5, borderColor: '#333', justifyContent: 'flex-start', alignItems: 'flex-start' },
    tableCellName: { padding: 4, fontSize: 8, flex: 1, borderRightWidth: 0.5, borderColor: '#333', justifyContent: 'center', alignItems: 'center' },
    tableCellTime: { padding: 4, fontSize: 8, flex: 0.7, borderRightWidth: 0.5, borderColor: '#333', justifyContent: 'center', alignItems: 'center' },
    dutyDetails: { fontSize: 6, color: '#666', marginTop: 2 },
  });

  const s = pdfStyles;

  // 헤더
  const header = React.createElement(View, { style: s.header },
    React.createElement(Text, { style: s.headerTitle }, 'Daily Office Duty'),
  );

  // 정보 섹션
  const infoSection = React.createElement(View, { style: s.infoSection },
    React.createElement(Text, { style: s.infoItem }, `Date: ${safeDutyDate || '-'}`),
    React.createElement(Text, { style: s.infoItem }, `Location: ${safeSelectedOffice || '-'}`),
    React.createElement(Text, { style: s.infoItem }, `Generated: ${generatedDate || '-'}`),
  );

  // 테이블 헤더
  const tableHeader = React.createElement(View, { style: s.tableHeader },
    React.createElement(View, { style: s.tableCellNo }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'No.')),
    React.createElement(View, { style: s.tableCellDesc }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Duty Description')),
    React.createElement(View, { style: s.tableCellDetails }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Details')),
    React.createElement(View, { style: s.tableCellName }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Done By')),
    React.createElement(View, { style: s.tableCellTime }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Time')),
    React.createElement(View, { style: s.tableCellName }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Checked By')),
    React.createElement(View, { style: s.tableCellTime }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Time')),
  );

  // Duty Description 텍스트를 줄바꿈 처리하는 함수
  const renderDutyDescription = (text: string, details?: string | null) => {
    const lines = text.split('<br>');
    const elements: any[] = [];
    
    lines.forEach((line, index) => {
      if (index > 0) {
        elements.push(React.createElement(Text, { key: `br-${index}` }, '\n'));
      }
      elements.push(React.createElement(Text, { key: `line-${index}` }, line));
    });
    
    if (details && details !== null) {
      elements.push(React.createElement(Text, { key: 'details', style: s.dutyDetails }, `\n${details}`));
    }
    
    return React.createElement(Text, null, ...elements);
  };

  // 테이블 데이터 행 (32개 행)
  const dutyDescriptions = [
    { main: 'Turn Off Answering Service', details: null },
    { main: 'All charts filed back?', details: null },
    { main: 'Charts pulled for next day', details: null },
    { main: 'Check eligibility', details: '1st of every month come in early to check eligibility by 8:30 am' },
    { main: 'If pt is not eligible call and inform', details: null },
    { main: 'Insurance breakdown for next day\'s patients', details: 'Call and get ins. info if necessary' },
    { main: 'Check ledger for any balance on the account', details: 'Fill out "Account with Balances Form" and fax\nCalled to inform patient of balance?' },
    { main: 'Morning confirmations', details: 'At least by noon' },
    { main: 'No shows entered on ledger', details: null },
    { main: 'No shows stamped in patient charts', details: null },
    { main: 'Reconfirming completed?', details: 'Start at 4:00pm' },
    { main: 'One week reminders completed?', details: null },
    { main: 'Call all treatment patients from today for post op', details: null },
    { main: 'Total lab case deposits/deliveries', details: 'Name/DOB' },
    { main: 'Check all undelivered lab cases and make appointments', details: 'Any Lab case that is more than 3 weeks old must be sent to corporate along with $20 deposit' },
    { main: 'Check all lab cases for next day', details: 'Call lab for next day pick up\'s' },
    { main: 'N₂O/ Compressor Off', details: null },
    { main: 'Did you read the meter on the Oxygen/N₂O/Helium tank?', details: null },
    { main: 'How many tanks are empty & need to be replaced?', details: null },
    { main: 'Check restrooms initial logs hourly', details: null },
    { main: 'Swept/Mopped', details: null },
    { main: 'Cleaned Breakroom', details: null },
    { main: 'Sterilizers: Cycle Complete', details: '(Do Not Push Stop)' },
    { main: 'Drained Ultrasonic', details: null },
    { main: 'Spore Test', details: 'Every Monday' },
    { main: 'Turn Off All TV\'s and Computers at the End of the Day', details: null },
    { main: 'Postcards Ready for Pick-up', details: null },
    { main: 'Clean traps everyday', details: '(chair)' },
    { main: 'Clean main trap 1st/15th', details: '(by vacuum)' },
    { main: 'Did you flush the lines with hot water?', details: null },
    { main: 'Check all doors are locked', details: null },
    { main: 'Turn On Answering Service', details: null },
  ];

  const tableRows = dutyDescriptions.map((duty, index) => {
    const rowNum = index + 1;
    const rowKey = `Row${rowNum}`;
    
    // Details 필드 처리
    let detailsText = '';
    if (rowNum === 2) {
      detailsText = safeStr(safeDutyData[`${rowKey}_YesNo`], 100);
    } else if (rowNum === 4) {
      detailsText = '';
    } else if (rowNum === 5) {
      detailsText = `How many pt's did you call: ${safeStr(safeDutyData[`${rowKey}_CallNum`], 50)}`;
    } else if (rowNum === 7) {
      detailsText = safeStr(safeDutyData[`${rowKey}_YesNo`], 100);
    } else if (rowNum === 14) {
      detailsText = safeStr(safeDutyData[`${rowKey}_Name/DOB`], 200);
    } else if (rowNum === 15) {
      detailsText = safeStr(safeDutyData[`${rowKey}_LabCases`], 200);
    } else if (rowNum === 18) {
      detailsText = safeStr(safeDutyData[`${rowKey}_YesNo`], 100);
    } else if (rowNum === 19) {
      const o2 = safeStr(safeDutyData[`${rowKey}_O2`], 10) || '0';
      const n2o = safeStr(safeDutyData[`${rowKey}_N2O`], 10) || '0';
      const he = safeStr(safeDutyData[`${rowKey}_He`], 10) || '0';
      detailsText = `O2: ${o2}\nN2O: ${n2o}\nHe: ${he}`;
    } else if (rowNum === 21) {
      detailsText = safeStr(safeDutyData[`${rowKey}_YesNo`], 100);
    } else if (rowNum === 22) {
      detailsText = safeStr(safeDutyData[`${rowKey}_YesNo`], 100);
    } else if (rowNum === 23) {
      detailsText = safeStr(safeDutyData[`${rowKey}_YesNo`], 100);
    } else if (rowNum === 24) {
      detailsText = safeStr(safeDutyData[`${rowKey}_YesNo`], 100);
    } else if (rowNum === 30) {
      detailsText = safeStr(safeDutyData[`${rowKey}_YesNo`], 100);
    }

    const doneBy = safeStr(safeDutyData[`${rowKey}_Done`], 100);
    const time = safeStr(safeDutyData[`${rowKey}_Time`], 20);
    const checkedBy = safeStr(safeDutyData[`${rowKey}_Checked`], 100);
    const checkedTime = safeStr(safeDutyData[`${rowKey}_Checked_Time`], 20);

    return React.createElement(View, { key: rowNum, style: s.tableRow },
      React.createElement(View, { style: s.tableCellNo }, React.createElement(Text, null, String(rowNum))),
      React.createElement(View, { style: s.tableCellDesc }, renderDutyDescription(duty.main, duty.details)),
      React.createElement(View, { style: s.tableCellDetails }, React.createElement(Text, null, detailsText || '-')),
      React.createElement(View, { style: s.tableCellName }, React.createElement(Text, null, doneBy || '-')),
      React.createElement(View, { style: s.tableCellTime }, React.createElement(Text, null, time || '-')),
      React.createElement(View, { style: s.tableCellName }, React.createElement(Text, null, checkedBy || '-')),
      React.createElement(View, { style: s.tableCellTime }, React.createElement(Text, null, checkedTime || '-')),
    );
  });

  // 테이블
  const table = React.createElement(View, { style: s.table }, tableHeader, ...tableRows);

  return React.createElement(Document, null,
    React.createElement(Page, { size: 'A4', orientation: 'portrait', style: s.page },
      header,
      infoSection,
      table
    ),
  );
}

type RowConfig = {
  num: number;
  title: string;
  suffix?: string;
  instructions?: string[];
  details?: string[];
  yesNo?: string;
  noLabel?: string;
  deadlineDisplay?: string;
  textarea?: { field: string; placeholder: string };
  textInput?: { field: string; placeholder: string };
  numberInputs?: { field: string; label: string }[];
};

const ROW_CONFIGS: RowConfig[] = [
  {
    num: 1,
    title: 'Turn Off Answering Service',
    instructions: [
      '1) Go to phone system website',
      '2) Log in',
      '4) Ensure your office is selected',
      '5) Select override office hours',
      '6) Select "Reset Default Office Hours"',
    ],
    deadlineDisplay: 'Deadline: 9:00 AM',
  },
  {
    num: 2,
    title: 'All charts filed back?',
    yesNo: 'Row2_YesNo',
    deadlineDisplay: 'Deadline: 4 PM',
  },
  {
    num: 3,
    title: 'Charts pulled for next day',
    deadlineDisplay: 'Deadline: 12:00 PM',
  },
  {
    num: 4,
    title: 'Check eligibility',
    details: ['1st of every month come in early to check eligibility by 8:30 am'],
    deadlineDisplay: 'Deadline: 4:30 PM',
  },
  {
    num: 5,
    title: 'If pt is not eligible call and inform',
    textInput: { field: 'Row5_CallNum', placeholder: "How many pt's did you call?" },
  },
  {
    num: 6,
    title: "Insurance breakdown for next day's patients",
    details: ['Call and get ins. info if necessary'],
    deadlineDisplay: 'Deadline: 4:30 PM',
  },
  {
    num: 7,
    title: 'Check ledger for any balance on the account',
    details: [
      'Fill out "Account with Balances Form" and fax to the AR Department',
      'Called to inform patient of balance?',
    ],
    yesNo: 'Row7_YesNo',
    deadlineDisplay: 'Deadline: 12:00 PM',
  },
  {
    num: 8,
    title: 'Morning confirmations',
    details: ['At least by noon'],
    deadlineDisplay: 'Deadline: 12:00 PM',
  },
  { num: 9, title: 'No shows entered on ledger' },
  { num: 10, title: 'No shows stamped in patient charts' },
  {
    num: 11,
    title: 'Reconfirming completed?',
    details: ['Start at 4:00pm'],
    deadlineDisplay: 'Deadline: 4:30 PM',
  },
  {
    num: 12,
    title: 'One week reminders completed?',
    deadlineDisplay: 'Deadline: 3:00 PM',
  },
  { num: 13, title: 'Call all treatment patients from today for post op' },
  {
    num: 14,
    title: 'Total lab case deposits/deliveries',
    textarea: { field: 'Row14_Name/DOB', placeholder: 'Name/DOB(mm/dd/yyyy) 1\n\nName/DOB(mm/dd/yyyy) 2\n\nName/DOB(mm/dd/yyyy) 3' },
  },
  {
    num: 15,
    title: 'Check all undelivered lab cases and make appointments',
    details: ['Any Lab case that is more than 3 weeks old must be sent to corporate along with $20 deposit'],
    textarea: { field: 'Row15_LabCases', placeholder: '1)\n\n2)\n\n3)' },
  },
  {
    num: 16,
    title: 'Check all lab cases for next day',
    details: ["Call lab for next day pick up's"],
  },
  { num: 17, title: 'N₂O/ Compressor Off' },
  {
    num: 18,
    title: 'Did you read the meter on the Oxygen/N₂O/Helium tank?',
    yesNo: 'Row18_YesNo',
  },
  {
    num: 19,
    title: 'How many tanks are empty & need to be replaced?',
    numberInputs: [
      { field: 'Row19_O2', label: 'O₂:' },
      { field: 'Row19_N2O', label: 'N₂O:' },
      { field: 'Row19_He', label: 'He:' },
    ],
  },
  { num: 20, title: 'Check restrooms initial logs hourly' },
  { num: 21, title: 'Swept/Mopped', yesNo: 'Row21_YesNo' },
  { num: 22, title: 'Cleaned Breakroom', yesNo: 'Row22_YesNo' },
  {
    num: 23,
    title: 'Sterilizers: Cycle Complete',
    yesNo: 'Row23_YesNo',
    noLabel: 'No (Do Not Push Stop)',
  },
  { num: 24, title: 'Drained Ultrasonic', yesNo: 'Row24_YesNo' },
  {
    num: 25,
    title: 'Spore Test',
    details: ['Every Monday'],
    deadlineDisplay: 'Deadline: 11:00 AM (Mondays only)',
  },
  { num: 26, title: "Turn Off All TV's and Computers at the End of the Day" },
  { num: 27, title: 'Postcards Ready for Pick-up' },
  { num: 28, title: 'Clean traps everyday', suffix: '(chair)' },
  { num: 29, title: 'Clean main trap 1st/15th', suffix: '(by vacuum)' },
  { num: 30, title: 'Did you flush the lines with hot water?', yesNo: 'Row30_YesNo' },
  { num: 31, title: 'Check all doors are locked' },
  {
    num: 32,
    title: 'Turn On Answering Service',
    instructions: [
      '1) Go to phone system website',
      '2) Log in',
      '3) Click on user icon',
      '4) Select override office hours',
      '5) Select "Office is Closed"',
      '6) For "1day"',
      '7) Call the office to verify that calls were transferred correctly. (For Holiday Weekends, Set for 1 week)',
    ],
  },
];

export default function DailyOfficeDuties() {
  const [loading, setLoading] = useState(false);
  const [autoSaveStatus, setAutoSaveStatus] = useState('');
  const [submitStatus, setSubmitStatus] = useState('');
  const [progress, setProgress] = useState(0);
  const [isUpdatingFromFirebase, setIsUpdatingFromFirebase] = useState(false);
  const [isAuthorized, setIsAuthorized] = useState<boolean | null>(null); // null: 확인 중, true: 인증됨, false: 인증 실패
  const [OfficesOptions, setOfficesOptions] = useState<string[]>([]);
  
  // Rate limiting을 위한 ref
  const lastUpdateDutyDataCall = useRef<number>(0);
  const lastSubmitCall = useRef<number>(0);
  const fieldRateLimit = useRef<Record<string, number>>({});
  
  // 사용자 세션 ID 생성 (페이지 로드 시 한 번만)
  const [userSessionId] = useState(() => Math.random().toString(36).substr(2, 9));
  
  // 마지막 저장된 데이터 추적
  const [lastSavedData, setLastSavedData] = useState({});
  
  // 날짜 상태
  const [dutyDate, setDutyDate] = useState(() => {
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

  // 오피스 선택 상태
  const [selectedOffice, setSelectedOffice] = useState('');
  
  // 오피스 옵션
  const officeOptions = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

  // 모든 업무 항목 상태
  const [dutyData, setDutyData] = useState({
    Row1_Done: '', Row1_Time: '', Row1_Checked: '', Row1_Checked_Time: '',
    Row2_YesNo: '', Row2_Done: '', Row2_Time: '', Row2_Checked: '', Row2_Checked_Time: '',
    Row3_Done: '', Row3_Time: '', Row3_Checked: '', Row3_Checked_Time: '',
    Row4_Done: '', Row4_Time: '', Row4_Checked: '', Row4_Checked_Time: '',
    Row5_CallNum: '', Row5_Done: '', Row5_Time: '', Row5_Checked: '', Row5_Checked_Time: '',
    Row6_Done: '', Row6_Time: '', Row6_Checked: '', Row6_Checked_Time: '',
    Row7_YesNo: '', Row7_Done: '', Row7_Time: '', Row7_Checked: '', Row7_Checked_Time: '',
    Row8_Done: '', Row8_Time: '', Row8_Checked: '', Row8_Checked_Time: '',
    Row9_Done: '', Row9_Time: '', Row9_Checked: '', Row9_Checked_Time: '',
    Row10_Done: '', Row10_Time: '', Row10_Checked: '', Row10_Checked_Time: '',
    Row11_Done: '', Row11_Time: '', Row11_Checked: '', Row11_Checked_Time: '',
    Row12_Done: '', Row12_Time: '', Row12_Checked: '', Row12_Checked_Time: '',
    Row13_Done: '', Row13_Time: '', Row13_Checked: '', Row13_Checked_Time: '',
    'Row14_Name/DOB': '', Row14_Done: '', Row14_Time: '', Row14_Checked: '', Row14_Checked_Time: '',
    Row15_LabCases: '', Row15_Done: '', Row15_Time: '', Row15_Checked: '', Row15_Checked_Time: '',
    Row16_Done: '', Row16_Time: '', Row16_Checked: '', Row16_Checked_Time: '',
    Row17_Done: '', Row17_Time: '', Row17_Checked: '', Row17_Checked_Time: '',
    Row18_YesNo: '', Row18_Done: '', Row18_Time: '', Row18_Checked: '', Row18_Checked_Time: '',
    Row19_O2: '', Row19_N2O: '', Row19_He: '', Row19_Done: '', Row19_Time: '', Row19_Checked: '', Row19_Checked_Time: '',
    Row20_Done: '', Row20_Time: '', Row20_Checked: '', Row20_Checked_Time: '',
    Row21_YesNo: '', Row21_Done: '', Row21_Time: '', Row21_Checked: '', Row21_Checked_Time: '',
    Row22_YesNo: '', Row22_Done: '', Row22_Time: '', Row22_Checked: '', Row22_Checked_Time: '',
    Row23_YesNo: '', Row23_Done: '', Row23_Time: '', Row23_Checked: '', Row23_Checked_Time: '',
    Row24_YesNo: '', Row24_Done: '', Row24_Time: '', Row24_Checked: '', Row24_Checked_Time: '',
    Row25_Done: '', Row25_Time: '', Row25_Checked: '', Row25_Checked_Time: '',
    Row26_Done: '', Row26_Time: '', Row26_Checked: '', Row26_Checked_Time: '',
    Row27_Done: '', Row27_Time: '', Row27_Checked: '', Row27_Checked_Time: '',
    Row28_Done: '', Row28_Time: '', Row28_Checked: '', Row28_Checked_Time: '',
    Row29_Done: '', Row29_Time: '', Row29_Checked: '', Row29_Checked_Time: '',
    Row30_YesNo: '', Row30_Done: '', Row30_Time: '', Row30_Checked: '', Row30_Checked_Time: '',
    Row31_Done: '', Row31_Time: '', Row31_Checked: '', Row31_Checked_Time: '',
    Row32_Done: '', Row32_Time: '', Row32_Checked: '', Row32_Checked_Time: '',
  });

  // 🔒 보안: Firebase에서 로드한 데이터에서 허용된 필드만 추출
  const VALID_DUTY_KEYS = useRef(new Set(Object.keys(dutyData))).current;
  const filterFirebaseData = (data: Record<string, any>): Record<string, string> => {
    const filtered: Record<string, string> = {};
    for (const key of VALID_DUTY_KEYS) {
      if (key in data && typeof data[key] === 'string') {
        filtered[key] = data[key].slice(0, 500).replace(/[<>]/g, '');
      }
    }
    return filtered;
  };

  // 마감 시간 정의
  const DEADLINES = {
    'Row1_Done': { time: '09:00', message: 'Turn Off Answering Service should be done by 9:00 AM' },
    'Row2_Done': { time: '16:00', message: 'All charts filed back should be done by 4:00 PM' },
    'Row3_Done': { time: '12:00', message: 'Charts pulled for next day should be done by 12:00 PM' },
    'Row4_Done': { time: '16:30', message: 'Check eligibility should be done by 4:30 PM' },
    'Row6_Done': { time: '16:30', message: 'Insurance breakdown should be done by 4:30 PM' },
    'Row7_Done': { time: '12:00', message: 'Check ledger for balance should be done by 12:00 PM' },
    'Row8_Done': { time: '12:00', message: 'Morning confirmations should be done by 12:00 PM' },
    'Row11_Done': { time: '16:30', message: 'Reconfirming should be completed by 4:30 PM' },
    'Row12_Done': { time: '15:00', message: 'One week reminders should be done by 3:00 PM' },
    'Row25_Done': { time: '11:00', message: 'Spore Test should be done by 11:00 AM' }
  };

  // 현재 캘리포니아 시간 가져오기
  const getCurrentCaliforniaTime = () => {
    const now = new Date();
    return new Date(now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
  };

  // 12시간 형식 시간 가져오기
  const getCurrentTime12Hour = () => {
    const californiaTime = getCurrentCaliforniaTime();
    let hours = californiaTime.getHours();
    const minutes = californiaTime.getMinutes();
    const ampm = hours >= 12 ? 'PM' : 'AM';
    
    hours = hours % 12;
    hours = hours ? hours : 12;
    
    return `${hours}:${minutes.toString().padStart(2, '0')} ${ampm}`;
  };

  // 자동 저장 함수
  const autoSave = useCallback(async () => {
    if (!dutyDate || !selectedOffice || isUpdatingFromFirebase) return;

    // 데이터가 실제로 변경되었는지 확인
    const hasChanges = JSON.stringify(dutyData) !== JSON.stringify(lastSavedData);
    if (!hasChanges) return;

    try {
      const dataToSave = {
        dutyDate,
        selectedOffice,
        ...dutyData,
        timestamp: new Date().toISOString(),
        autoSaved: true,
        lastUpdatedBy: userSessionId
      };

      // 🔒 보안: 저장 전 데이터 검증
      const validatedDutyData: { [key: string]: string } = {};
      for (const [key, value] of Object.entries(dutyData)) {
        validatedDutyData[key] = validateInput(value as string, 500);
      }
      
      const validatedDataToSave = {
        ...dataToSave,
        ...validatedDutyData
      };
      
      const docId = `${dutyDate}_${selectedOffice}`;
      await setDoc(doc(db, "daily-office-duties", docId), validatedDataToSave);
      
      // 저장 성공 시 마지막 저장된 데이터 업데이트
      setLastSavedData({ ...dutyData });
      
    } catch {
      // auto-save failure is silently ignored; real-time listener will resync
    }
  }, [dutyDate, selectedOffice, dutyData, lastSavedData, isUpdatingFromFirebase, userSessionId]);

  // 데이터 변경 시에만 자동 저장
  useEffect(() => {
    // 초기 로드 시에는 저장하지 않음
    if (Object.values(dutyData).some(value => value !== '')) {
      autoSave();
    }
  }, [dutyData]);

  // 데이터 로드
  const loadData = async () => {
    if (!dutyDate || !selectedOffice) return;

    try {
      setSubmitStatus('Loading data...');
      
      const docId = `${dutyDate}_${selectedOffice}`;
      const docSnap = await getDoc(doc(db, "daily-office-duties", docId));
      
      if (docSnap.exists()) {
        const safeData = filterFirebaseData(docSnap.data());
        
        // Firebase에서 업데이트되는 동안 자동 저장 방지
        setIsUpdatingFromFirebase(true);
        
        setDutyData(prevData => ({
          ...prevData,
          ...safeData
        }));
        
        // 로드된 데이터를 마지막 저장된 데이터로 설정
        setLastSavedData({ ...safeData });
        
        // 짧은 지연 후 Firebase 업데이트 플래그 해제
        setTimeout(() => {
          setIsUpdatingFromFirebase(false);
        }, 100);
        
        setSubmitStatus('Data loaded successfully');
        setTimeout(() => setSubmitStatus(''), 2000);
      } else {
        // 데이터가 없으면 초기화
        const initialData: { [key: string]: string } = {};
        Object.keys(dutyData).forEach(key => {
          initialData[key] = '';
        });
        setDutyData(initialData as typeof dutyData);
        setLastSavedData(initialData);
        setSubmitStatus('No data found - initialized empty form');
        setTimeout(() => setSubmitStatus(''), 2000);
      }
      
    } catch (error) {
      setSubmitStatus('Error loading data. Please try again.');
      setTimeout(() => setSubmitStatus(''), 3000);
    }
  };

  // 날짜 또는 오피스 변경 시 데이터 로드
  useEffect(() => {
    if (dutyDate && selectedOffice) {
      loadData();
    }
  }, [dutyDate, selectedOffice]);

  // 실시간 데이터 동기화
  useEffect(() => {
    if (!dutyDate || !selectedOffice) return;

    const docId = `${dutyDate}_${selectedOffice}`;
    const docRef = doc(db, "daily-office-duties", docId);
    
    const unsubscribe = onSnapshot(docRef, (docSnap) => {
      if (docSnap.exists()) {
        const rawData = docSnap.data();
        const safeData = filterFirebaseData(rawData);
        
        // Firebase에서 업데이트되는 동안 자동 저장 방지
        setIsUpdatingFromFirebase(true);
        
        setDutyData(prevData => {
          return {
            ...prevData,
            ...safeData
          };
        });
        
        // 실시간 업데이트된 데이터를 마지막 저장된 데이터로 설정
        setLastSavedData({ ...safeData });
        
        // 짧은 지연 후 Firebase 업데이트 플래그 해제
        setTimeout(() => {
          setIsUpdatingFromFirebase(false);
        }, 100);
        
        // 다른 사용자의 업데이트만 표시 (자신의 업데이트는 제외)
        if (rawData.timestamp && 
            new Date(rawData.timestamp).getTime() > Date.now() - 5000 && 
            rawData.lastUpdatedBy && 
            rawData.lastUpdatedBy !== userSessionId) {
          setAutoSaveStatus('🔄 Updated from another user');
          setTimeout(() => setAutoSaveStatus(''), 2000);
        }
      }
    }, (error) => {
      setAutoSaveStatus('❌ Connection error');
      setTimeout(() => setAutoSaveStatus(''), 3000);
    });

    return () => {
      unsubscribe();
    };
  }, [dutyDate, selectedOffice]);

  // 컴포넌트 마운트 시 사용자 인증 및 role 확인
  useEffect(() => {
    // Firebase Auth 상태 변경 감지
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      try {
        if (!currentUser) {
          setIsAuthorized(false);
          return;
        }

        // Firestore에서 사용자 role 확인
        const userDoc = await getDoc(doc(db, 'users', currentUser.uid));
        if (!userDoc.exists()) {
          setIsAuthorized(false);
          return;
        }

        const userData = userDoc.data();

        if (userData?.role !== 'Manager' && userData?.role !== 'User') {
          setIsAuthorized(false);
          return;
        }

        setIsAuthorized(true);

        // offices 처리: 배열이거나 단일 값일 수 있음
        if (userData?.offices) {
          const officesArray = Array.isArray(userData.offices) 
            ? userData.offices 
            : [userData.offices];
          const validOptions = officesArray.filter((g: string) => officeOptions.includes(g));
          if (validOptions.length > 0) {
            setOfficesOptions(validOptions);
            if (validOptions.length === 1) {
              setSelectedOffice(validOptions[0]);
            }
          }
        }
      } catch (error: any) {
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

  // 컴포넌트 마운트 시 초기 로드는 dutyDate 변경 시 로드로 대체됨


  // 🔒 보안: 입력 검증 함수
  const validateInput = (value: string, maxLength: number = 500): string => {
    if (typeof value !== 'string') return '';
    return value.slice(0, maxLength).replace(/[<>]/g, '');
  };

  // 데이터 업데이트 함수 (Rate limiting 적용)
  const updateDutyData = (field: string, value: string) => {
    // Rate limiting: 입력 반응성을 위해 완화된 제한 적용
    // (자동 저장은 별도 debounce로 처리되므로 입력 자체는 빠르게 반응)
    const now = Date.now();
    const lastCall = fieldRateLimit.current[field] || 0;

    // 전역 rate limiting: 모든 업데이트에 대해 50ms 제한 (입력 반응성 향상)
    if (now - lastUpdateDutyDataCall.current < 50) {
      return;
    }
    lastUpdateDutyDataCall.current = now;

    // 개별 필드 rate limiting: 동일 필드에 대해 100ms 제한 (입력 반응성 향상)
    if (now - lastCall < 100) {
      return;
    }
    fieldRateLimit.current[field] = now;

    // 🔒 보안: 입력 검증 및 길이 제한
    const validatedValue = validateInput(value, 500);
    
    setDutyData(prev => {
      const newData: { [key: string]: string } = { ...prev, [field]: validatedValue };
      
      // Done by 필드가 입력되면 시간 자동 기록
      if (field.endsWith('_Done') && validatedValue.trim() !== '' && (prev[field as keyof typeof prev] as string)?.trim() === '') {
        const rowNumber = field.match(/Row(\d+)_Done/)?.[1];
        if (rowNumber) {
          const timeField = `Row${rowNumber}_Time`;
          if (!newData[timeField]) {
            newData[timeField] = getCurrentTime12Hour();
          }
        }
      }
      
      // Done by 필드가 비워지면 시간도 비움
      if (field.endsWith('_Done') && validatedValue.trim() === '') {
        const rowNumber = field.match(/Row(\d+)_Done/)?.[1];
        if (rowNumber) {
          const timeField = `Row${rowNumber}_Time`;
          newData[timeField] = '';
        }
      }

      // Checked by 필드가 입력되면 시간 자동 기록
      if (field.endsWith('_Checked') && validatedValue.trim() !== '' && (prev[field as keyof typeof prev] as string)?.trim() === '') {
        const rowNumber = field.match(/Row(\d+)_Checked/)?.[1];
        if (rowNumber) {
          const checkedTimeField = `Row${rowNumber}_Checked_Time`;
          if (!newData[checkedTimeField]) {
            newData[checkedTimeField] = getCurrentTime12Hour();
          }
        }
      }

      // Checked by 필드가 비워지면 시간도 비움
      if (field.endsWith('_Checked') && validatedValue.trim() === '') {
        const rowNumber = field.match(/Row(\d+)_Checked/)?.[1];
        if (rowNumber) {
          const checkedTimeField = `Row${rowNumber}_Checked_Time`;
          newData[checkedTimeField] = '';
        }
      }
      
      return newData as typeof prev;
    });
  };

  // 마감 시간 체크 - 각 행별로 개별 체크
  const isRowOverdue = (rowItemName: keyof typeof DEADLINES) => {
    const californiaTime = getCurrentCaliforniaTime();
    const currentMinutes = californiaTime.getHours() * 60 + californiaTime.getMinutes();
    const currentDay = californiaTime.getDay();
    
    if (!DEADLINES[rowItemName]) return false;
    
    const deadline = DEADLINES[rowItemName];
    const [hours, minutes] = deadline.time.split(':').map(Number);
    const deadlineMinutes = hours * 60 + minutes;
    
    const isCompleted = dutyData[rowItemName as keyof typeof dutyData]?.trim() !== '';
    const isOverdue = currentMinutes > deadlineMinutes;
    
    // Spore Test는 월요일만 체크
    if (rowItemName === 'Row25_Done' && currentDay !== 1) {
      return false;
    }
    
    return isOverdue && !isCompleted;
  };


  // 제출 처리 (Rate limiting 적용)
  const handleSubmit = async () => {
    // Rate limiting: 최근 3초 내 호출 방지 (PDF 생성은 무거운 작업)
    const now = Date.now();
    if (now - lastSubmitCall.current < 3000) {
      alert('⚠️ Please try again.');
      return;
    }
    lastSubmitCall.current = now;

    // 확인 다이얼로그
    const confirmed = confirm('Would you like to submit?');
    if (!confirmed) {
      return;
    }

    try {
      setLoading(true);
      setSubmitStatus('Saving...');
      setProgress(10);

      // 🔒 보안: 경로 조작 방지를 위한 값 검증
      if (!dutyDate || !/^\d{4}-\d{2}-\d{2}$/.test(dutyDate)) {
        alert('Invalid date format.');
        setLoading(false);
        return;
      }
      if (!selectedOffice || !officeOptions.includes(selectedOffice)) {
        alert('Invalid office selection.');
        setLoading(false);
        return;
      }

      // 1. PDF 생성 (클라이언트 사이드)
      setSubmitStatus('Submitting...');
      setProgress(30);
      
      // 데이터 sanitize
      const safeDutyDate = dutyDate.trim().slice(0, 10);
      const safeSelectedOffice = selectedOffice.trim().slice(0, 100);
      const safeDutyData: { [key: string]: string } = {};
      for (const [key, value] of Object.entries(dutyData)) {
        if (typeof value === 'string' && value.length <= 500) {
          safeDutyData[key] = value.replace(/[<>]/g, '');
        } else {
          safeDutyData[key] = '';
        }
      }

      // 생성 날짜 포맷팅
      const generatedDate = new Date().toLocaleString('en-US', {
        timeZone: 'America/Los_Angeles',
        year: 'numeric',
        month: 'long',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      });

      // PDF 문서 생성
      setSubmitStatus('Creating PDF document...');
      setProgress(40);
      
      const pdfDoc = createDailyOfficeDutyPDFDocument({
        safeDutyDate,
        safeSelectedOffice,
        safeDutyData,
        generatedDate
      });

      // PDF blob 생성
      setSubmitStatus('Processing PDF...');
      setProgress(50);
      
      const pdfBlob = await pdf(pdfDoc).toBlob();

      if (pdfBlob && pdfBlob.size > 0) {
        // PDF를 Firebase Storage에 저장
        setSubmitStatus('Saving...');
        setProgress(60);
        
        try {
          const storage = getStorage();
          const now = new Date();
          const laTime = new Date(now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
          let hours = laTime.getHours();
          const minutes = laTime.getMinutes();
          const ampm = hours >= 12 ? 'pm' : 'am';
          hours = hours % 12;
          hours = hours ? hours : 12;
          const timeStamp = `${hours}${minutes.toString().padStart(2, '0')}${ampm}`;
          const filename = `2) ${dutyDate}_${selectedOffice}_Daily Office Duty_${timeStamp}.pdf`;
          const storageRef = ref(storage, `endofday-pdfs/${selectedOffice}/${dutyDate}/${filename}`);
          
          // PDF 업로드
          await uploadBytes(storageRef, pdfBlob);
          
          setSubmitStatus('✅ PDF saved to archive successfully!');
        } catch (storageError) {
          alert('An error occurred while submitting. Please try again.');
          setSubmitStatus('❌ Submission failed. Please try again.');
          setProgress(0);
          setTimeout(() => { setLoading(false); setSubmitStatus(''); }, 3000);
          return;
        }
        
        // 2. 데이터 삭제
        setSubmitStatus('Cleaning up...');
        setProgress(80);
        const docId = `${dutyDate}_${selectedOffice}`;
        await deleteDoc(doc(db, "daily-office-duties", docId));
        
        // 3. 폼 초기화
        setDutyData(prevData => {
          const initialData: { [key: string]: string } = {};
          Object.keys(prevData).forEach(key => {
            initialData[key] = '';
          });
          return initialData as typeof prevData;
        });

        setSubmitStatus('Complete!');
        setProgress(100);
        
        // 2초 후 모달 닫기
        setTimeout(() => {
          setLoading(false);
          setSubmitStatus('');
          setProgress(0);
        }, 2000);
      } else {
        throw new Error('PDF is empty');
      }

    } catch (error) {
      setSubmitStatus('❌ Submission failed. Please try again.');
      setProgress(0);
      setTimeout(() => {
        setLoading(false);
        setSubmitStatus('');
      }, 3000);
    }
  };

  // 스타일 정의
  const styles: { [key: string]: React.CSSProperties } = {
    body: {
      fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
      padding: '20px',
      background: 'linear-gradient(135deg, #fce4ec 0%, #f8bbd0 100%)',
      color: '#2c3e50',
      lineHeight: '1.6',
      minHeight: '100vh'
    },
    container: {
      maxWidth: '85%',
      width: '85%',
      margin: '20px auto',
      padding: '30px',
      backgroundColor: 'white',
      borderRadius: '8px',
      boxShadow: '0 4px 8px rgba(0, 0, 0, 0.1)',
      position: 'relative'
    },
    header: {
      color: '#2c3e50',
      textAlign: 'center' as const,
      marginBottom: '30px',
      paddingBottom: '10px',
      borderBottom: '2px solid #e9ecef',
      fontSize: '2em',
      fontWeight: 'bold'
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
    table: {
      borderCollapse: 'collapse',
      width: '100%',
      marginTop: '20px',
      backgroundColor: 'white',
      boxShadow: '0 1px 3px rgba(11, 4, 4, 0.1)'
    },
    th: {
      border: '1px solid #e9ecef',
      padding: '10px',
      textAlign: 'left' as const,
      verticalAlign: 'top' as const,
      backgroundColor: '#2c3e50',
      color: 'white',
      fontWeight: '500'
    },
    td: {
      border: '1px solid #e9ecef',
      padding: '10px',
      textAlign: 'left' as const,
      verticalAlign: 'top' as const
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
      textAlign: 'center' as const,
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
    overdueWarning: {
      color: '#d32f2f',
      fontWeight: 'bold',
      fontSize: '0.9em',
      marginTop: '4px',
      backgroundColor: '#ffcdd2',
      padding: '4px 8px',
      borderRadius: '4px',
      display: 'inline-block'
    },
    dutyDetails: {
      fontSize: '0.9em',
      color: '#555',
      marginTop: '4px'
    },
    deadlineInfo: {
      fontSize: '0.8em',
      color: '#666',
      fontStyle: 'italic',
      marginTop: '2px'
    },
    inlineOption: {
      display: 'inline-block',
      marginRight: '15px',
      verticalAlign: 'middle' as const
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
      <div style={styles.body}>
        <div style={{
          display: 'flex',
          justifyContent: 'center',
          alignItems: 'center',
          height: '100vh',
          background: 'linear-gradient(135deg, #fce4ec 0%, #f8bbd0 100%)',
          fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
        }}>
          <div style={{ textAlign: 'center' }}>
            <div style={{ fontSize: '24px', marginBottom: '20px' }}>🔐</div>
            <div style={{ fontSize: '18px', color: '#2c3e50' }}>Verifying authentication...</div>
          </div>
        </div>
      </div>
    );
  }

  if (isAuthorized === false) {
    return (
      <div style={styles.body}>
        <div style={{
          display: 'flex',
          justifyContent: 'center',
          alignItems: 'center',
          height: '100vh',
          background: 'linear-gradient(135deg, #fce4ec 0%, #f8bbd0 100%)',
          fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
        }}>
          <div style={{ textAlign: 'center' }}>
            <div style={{ fontSize: '24px', marginBottom: '20px' }}>🚫</div>
            <div style={{ fontSize: '18px', color: '#d32f2f', marginBottom: '10px' }}>You do not have access to this page.</div>
            <div style={{ fontSize: '14px', color: '#666' }}>You do not have access to this page.</div>
          </div>
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
        .overdue-row td {
          background-color: #ffcdd2 !important;
        }
        .overdue-row input,
        .overdue-row textarea {
          background-color: #ffcdd2 !important;
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
            textAlign: "center" as const,
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

        <h2 style={styles.header}>Daily Office Duty</h2>

        {/* 날짜 및 오피스 선택 */}
        <div style={{ display: 'flex', gap: '20px', marginBottom: '20px', flexWrap: 'wrap' }}>
          <div style={styles.formGroup}>
            <label style={styles.label} htmlFor="dutyDate">Date:</label>
            <input
              type="date"
              id="dutyDate"
              value={dutyDate}
              onChange={(e) => setDutyDate(e.target.value)}
              style={styles.input}
              required
            />
          </div>
          {OfficesOptions.length > 0 && (
            <div style={styles.formGroup}>
              <label style={styles.label} htmlFor="selectedOffice">Office:</label>
              {OfficesOptions.length === 1 ? (
                <span style={{
                  ...styles.input,
                  display: 'inline-block',
                  backgroundColor: '#e9ecef',
                  fontWeight: '600',
                  color: '#2c3e50'
                }}>
                  {selectedOffice}
                </span>
              ) : (
                <select
                  id="selectedOffice"
                  value={selectedOffice}
                  onChange={(e) => setSelectedOffice(e.target.value)}
                  style={styles.input}
                  required
                >
                  <option value="">-- Select Office --</option>
                  {OfficesOptions.map(office => (
                    <option key={office} value={office}>{office}</option>
                  ))}
                </select>
              )}
            </div>
          )}
        </div>

        {/* 업무 테이블 */}
        <table style={styles.table}>
          <thead>
            <tr>
              <th style={{ ...styles.th, width: '60px' }}>No.</th>
              <th style={styles.th}>Duty</th>
              <th style={styles.th}>Done by</th>
              <th style={{ ...styles.th, width: '100px' }}>Time</th>
              <th style={styles.th}>Checked by</th>
              <th style={{ ...styles.th, width: '100px' }}>Time</th>
            </tr>
          </thead>
          <tbody>
            {ROW_CONFIGS.map((row) => {
              const doneKey = `Row${row.num}_Done`;
              const checkedKey = `Row${row.num}_Checked`;
              const timeKey = `Row${row.num}_Time`;
              const deadlineEntry = DEADLINES[doneKey as keyof typeof DEADLINES];
              const overdue = deadlineEntry ? isRowOverdue(doneKey as keyof typeof DEADLINES) : false;
              const cellInput: React.CSSProperties = { ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' };
              const cellInputRO: React.CSSProperties = { ...cellInput, backgroundColor: '#f8f9fa', color: '#6c757d' };
              const textareaCSS: React.CSSProperties = { ...cellInput, minHeight: '80px', resize: 'vertical', whiteSpace: 'pre-wrap' };

              return (
                <tr key={row.num} className={overdue ? 'overdue-row' : ''}>
                  <td style={styles.td}><strong>{row.num}</strong></td>
                  <td style={styles.td}>
                    <strong>{row.title}</strong>{row.suffix && ` ${row.suffix}`}
                    {row.instructions && (
                      <div style={styles.dutyDetails}>
                        {row.instructions.map((line, i) => (
                          <React.Fragment key={i}>{i > 0 && <br/>}{line}</React.Fragment>
                        ))}
                      </div>
                    )}
                    {row.details?.map((d, i) => (
                      <div key={i} style={styles.dutyDetails}>{d}</div>
                    ))}
                    {row.yesNo && (
                      <>
                        <br/>
                        <div style={styles.inlineOption}>
                          <input type="radio" id={`${row.yesNo}_Yes`} name={row.yesNo} value="Yes"
                            checked={(dutyData as any)[row.yesNo] === 'Yes'}
                            onChange={(e) => updateDutyData(row.yesNo!, e.target.value)}
                            style={{ marginRight: '5px', cursor: 'pointer' }} />
                          <label htmlFor={`${row.yesNo}_Yes`}>Yes</label>
                        </div>
                        <div style={styles.inlineOption}>
                          <input type="radio" id={`${row.yesNo}_No`} name={row.yesNo} value="No"
                            checked={(dutyData as any)[row.yesNo] === 'No'}
                            onChange={(e) => updateDutyData(row.yesNo!, e.target.value)}
                            style={{ marginRight: '5px', cursor: 'pointer' }} />
                          <label htmlFor={`${row.yesNo}_No`}>{row.noLabel || 'No'}</label>
                        </div>
                      </>
                    )}
                    {row.textInput && (
                      <>
                        <br/>
                        <input type="text"
                          value={(dutyData as any)[row.textInput.field] || ''}
                          onChange={(e) => updateDutyData(row.textInput!.field, e.target.value)}
                          placeholder={row.textInput.placeholder}
                          style={cellInput} />
                      </>
                    )}
                    {row.textarea && (
                      <textarea
                        value={(dutyData as any)[row.textarea.field] || ''}
                        onChange={(e) => updateDutyData(row.textarea!.field, e.target.value)}
                        placeholder={row.textarea.placeholder}
                        style={textareaCSS} />
                    )}
                    {row.numberInputs?.map((ni) => (
                      <div key={ni.field} style={styles.inlineOption}>
                        <label htmlFor={ni.field}>{ni.label}</label>
                        <input type="text" id={ni.field}
                          value={(dutyData as any)[ni.field] || ''}
                          onChange={(e) => updateDutyData(ni.field, e.target.value)}
                          style={{ width: '50px', marginLeft: '5px' }} />
                      </div>
                    ))}
                    {row.deadlineDisplay && (
                      <div style={styles.deadlineInfo}>{row.deadlineDisplay}</div>
                    )}
                    {overdue && deadlineEntry && (
                      <div style={styles.overdueWarning}>⚠️ {deadlineEntry.message}</div>
                    )}
                  </td>
                  <td style={styles.td}>
                    <input type="text" value={(dutyData as any)[doneKey] || ''}
                      onChange={(e) => updateDutyData(doneKey, e.target.value)} style={cellInput} />
                  </td>
                  <td style={styles.td}>
                    <input type="text" value={(dutyData as any)[timeKey] || ''} readOnly style={cellInputRO} />
                  </td>
                  <td style={styles.td}>
                    <input type="text" value={(dutyData as any)[checkedKey] || ''}
                      onChange={(e) => updateDutyData(checkedKey, e.target.value)} style={cellInput} />
                  </td>
                  <td style={styles.td}>
                    <input type="text" value={(dutyData as any)[`Row${row.num}_Checked_Time`] || ''} readOnly style={cellInputRO} />
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>

        {/* 제출 버튼 */}
        <div style={{ textAlign: 'center' as const, marginTop: '20px' }}>
          <button
            onClick={handleSubmit}
            disabled={loading}
            style={{
              ...styles.submitButton,
              backgroundColor: loading ? '#bdc3c7' : '#3498db',
              cursor: loading ? 'not-allowed' : 'pointer'
            }}
          >
            {loading ? 'Submitting...' : 'Submit'}
          </button>
        </div>

        {/* 상태 메시지 */}
        {submitStatus && (
          <div style={{
            ...styles.statusMessage,
            backgroundColor: submitStatus.includes('failed') || submitStatus.includes('Error') ? '#f8d7da' : 
                           submitStatus.includes('successfully') ? '#d4edda' : '#d1ecf1',
            color: submitStatus.includes('failed') || submitStatus.includes('Error') ? '#721c24' : 
                   submitStatus.includes('successfully') ? '#155724' : '#0c5460',
            border: submitStatus.includes('failed') || submitStatus.includes('Error') ? '1px solid #f5c6cb' : 
                    submitStatus.includes('successfully') ? '1px solid #c3e6cb' : '1px solid #bee5eb'
          }}>
            {submitStatus}
          </div>
        )}

      </div>
    </div>
    </>
  );
}



