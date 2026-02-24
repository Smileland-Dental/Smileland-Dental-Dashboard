'use client'

import React, { useState, useEffect, useCallback, useRef } from "react";
import { doc, setDoc, collection, getDocs, deleteDoc, getDoc, query, where } from "firebase/firestore";
import { db, auth } from "@/lib/firebase.config";
import { getAuth, onAuthStateChanged } from 'firebase/auth';
import { getStorage, ref, uploadBytes } from 'firebase/storage';
import { pdf, Document, Page, View, Text, StyleSheet } from '@react-pdf/renderer';

// 타입 정의
interface PatientRowProps {
  row: any;
  updatePatientRow: (id: number, field: string, value: any) => void;
  removePatientRow: (id: number) => void;
  patientOfficeOptions: string[];
  getVisitTypeOptions: (office: string) => string[];
  remarkOptions: string[];
  otherDutyOptions: string[];
  inputStyle: React.CSSProperties;
  buttonStyle: React.CSSProperties;
}

// 개별 환자 행 컴포넌트 (메모이제이션 최적화)
const PatientRow = React.memo(({ 
  row, 
  updatePatientRow, 
  removePatientRow, 
  patientOfficeOptions, 
  getVisitTypeOptions, 
  remarkOptions, 
  otherDutyOptions,
  inputStyle,
  buttonStyle
}: PatientRowProps) => {
  const visitTypeOptions = getVisitTypeOptions(row.office);
  
  return (
    <tr style={{ backgroundColor: row.id % 2 === 0 ? '#f9f9f9' : 'white' }}>
      <td style={{ padding: '8px', textAlign: 'center', fontWeight: 'bold' }}>
        {row.id}
      </td>
      <td style={{ padding: '8px' }}>
        <input
          type="text"
          value={row.name}
          onChange={(e) => updatePatientRow(row.id, 'name', e.target.value)}
          style={{ ...inputStyle, margin: 0, fontSize: '14px' }}
        />
      </td>
      <td style={{ padding: '8px' }}>
        <select
          value={row.office}
          onChange={(e) => updatePatientRow(row.id, 'office', e.target.value)}
          style={{ ...inputStyle, margin: 0, fontSize: '14px' }}
        >
          <option value=""></option>
          {patientOfficeOptions.map(office => (
            <option key={office} value={office}>{office}</option>
          ))}
        </select>
      </td>
      <td style={{ padding: '8px' }}>
        <input
          type="date"
          value={row.appt_date}
          onChange={(e) => updatePatientRow(row.id, 'appt_date', e.target.value)}
          style={{ ...inputStyle, margin: 0, fontSize: '14px' }}
        />
      </td>
      <td style={{ padding: '8px' }}>
        <select
          value={row.visit_type}
          onChange={(e) => updatePatientRow(row.id, 'visit_type', e.target.value)}
          style={{ ...inputStyle, margin: 0, fontSize: '14px' }}
        >
          <option value=""></option>
          {visitTypeOptions.map(type => (
            <option key={type} value={type}>{type}</option>
          ))}
        </select>
      </td>
      <td style={{ padding: '8px', textAlign: 'center' }}>
        <input
          type="checkbox"
          checked={row.call_in}
          onChange={(e) => updatePatientRow(row.id, 'call_in', e.target.checked)}
          style={{ transform: 'scale(1.2)' }}
        />
      </td>
      <td style={{ padding: '8px', textAlign: 'center' }}>
        <input
          type="checkbox"
          checked={row.call_out}
          onChange={(e) => updatePatientRow(row.id, 'call_out', e.target.checked)}
          style={{ transform: 'scale(1.2)' }}
        />
      </td>
      <td style={{ padding: '8px' }}>
        <input
          type="time"
          value={row.time}
          onChange={(e) => updatePatientRow(row.id, 'time', e.target.value)}
          disabled={row.call_in || row.call_out}
          style={{ 
            ...inputStyle, 
            margin: 0, 
            fontSize: '14px',
            backgroundColor: (row.call_in || row.call_out) ? '#f0f0f0' : 'white',
            cursor: (row.call_in || row.call_out) ? 'not-allowed' : 'text'
          }}
        />
      </td>
      <td style={{ padding: '8px' }}>
        <select
          value={row.remark}
          onChange={(e) => updatePatientRow(row.id, 'remark', e.target.value)}
          style={{ ...inputStyle, margin: 0, fontSize: '14px' }}
        >
          <option value=""></option>
          {remarkOptions.map(remark => (
            <option key={remark} value={remark}>{remark}</option>
          ))}
        </select>
      </td>
      <td style={{ padding: '8px' }}>
        <select
          value={row.other_duty}
          onChange={(e) => updatePatientRow(row.id, 'other_duty', e.target.value)}
          style={{ ...inputStyle, margin: 0, fontSize: '14px' }}
        >
          <option value=""></option>
          {otherDutyOptions.map(duty => (
            <option key={duty} value={duty}>{duty}</option>
          ))}
        </select>
      </td>
      <td style={{ padding: '8px', textAlign: 'center' }}>
        <button
          onClick={() => removePatientRow(row.id)}
          style={{
            ...buttonStyle,
            backgroundColor: '#dc3545',
            padding: '6px 12px',
            fontSize: '12px',
            margin: 0
          }}
        >
          Delete
        </button>
      </td>
    </tr>
  );
}, (prevProps, nextProps) => {
  // 커스텀 비교 함수로 불필요한 리렌더링 방지
  return (
    prevProps.row.id === nextProps.row.id &&
    prevProps.row.name === nextProps.row.name &&
    prevProps.row.office === nextProps.row.office &&
    prevProps.row.appt_date === nextProps.row.appt_date &&
    prevProps.row.visit_type === nextProps.row.visit_type &&
    prevProps.row.call_in === nextProps.row.call_in &&
    prevProps.row.call_out === nextProps.row.call_out &&
    prevProps.row.time === nextProps.row.time &&
    prevProps.row.remark === nextProps.row.remark &&
    prevProps.row.other_duty === nextProps.row.other_duty
  );
});

// Firebase 데이터 sanitize 함수 (강화된 버전)
function sanitizeFirebaseDataClient(data: any, depth: number = 0): any {
  // 깊이 제한 (순환 참조 및 깊은 중첩 방지)
  if (depth > 20) return null;
  
  // 기본적인 데이터 검증
  if (data === null || data === undefined) return null;
  
  // 원시 타입 처리
  if (typeof data !== 'object') {
    // 문자열인 경우 길이 제한 및 특수 문자 제거
    if (typeof data === 'string') {
      // Firebase 문자열 필드 최대 크기: 1MB (안전하게 900KB로 제한)
      if (data.length > 900 * 1024) {
        return data.slice(0, 900 * 1024);
      }
      // 위험한 문자 제거 (XSS 방지)
      return data.replace(/[\x00-\x1F\x7F-\x9F]/g, '').trim();
    }
    // 숫자, 불린 등은 그대로 반환
    if (typeof data === 'number' && (isNaN(data) || !isFinite(data))) {
      return 0;
    }
    return data;
  }
  
  // 배열 처리
  if (Array.isArray(data)) {
    // 배열 크기 제한 (Firebase 제한 고려)
    if (data.length > 10000) {
      return data.slice(0, 10000).map(item => sanitizeFirebaseDataClient(item, depth + 1));
    }
    return data.map(item => sanitizeFirebaseDataClient(item, depth + 1));
  }
  
  // 객체 처리
  const sanitized: any = {};
  let keyCount = 0;
  const maxKeys = 1000; // Firebase 문서 필드 수 제한 고려
  
  for (const key in data) {
    if (Object.prototype.hasOwnProperty.call(data, key)) {
      // 키 개수 제한
      if (keyCount >= maxKeys) break;
      
      // 키 길이 제한 (Firebase 제한)
      if (key.length > 1500 || key.length === 0) continue;
      
      // 키에 허용되지 않은 문자 제거
      const safeKey = key.replace(/[.$[\]#\/]/g, '_').slice(0, 1500);
      
      // 값 sanitize
      sanitized[safeKey] = sanitizeFirebaseDataClient(data[key], depth + 1);
      keyCount++;
    }
  }
  
  return sanitized;
}

// Firebase Document ID sanitize 함수
function sanitizeFirebaseDocIdClient(docId: string): string {
  // Firebase Document ID 제한: 1-1500자, 특수문자 제한
  return docId
    .replace(/[\/\s]/g, '_') // 슬래시와 공백을 언더스코어로
    .replace(/[^a-zA-Z0-9_-]/g, '') // 허용된 문자만 유지
    .slice(0, 1500); // 길이 제한
}

// PDF 생성 유틸 함수
function safeStr(v: unknown, max: number): string {
  if (v == null) return '';
  return String(v).trim().slice(0, max).replace(/[<>]/g, '');
}

// 시간을 12시간 형식으로 변환하는 함수
function convertTo12Hour(timeStr: string): string {
  if (!timeStr || timeStr === '-') return '-';
  try {
    const [hours, minutes] = timeStr.split(':');
    const hour = parseInt(hours);
    const min = minutes || '00';
    if (hour === 0) return `12:${min} AM`;
    if (hour < 12) return `${hour}:${min} AM`;
    if (hour === 12) return `12:${min} PM`;
    return `${hour - 12}:${min} PM`;
  } catch {
    return timeStr;
  }
}

// Patient Log PDF 생성 함수
function createPatientLogPDFDocument(props: {
  safeDutyDate: string;
  safeUserName: string;
  safeWorkOffice: string;
  safeWorkHoursFrom: string;
  safeWorkHoursTo: string;
  safeDailyWorkReport: string;
  patientList: any[];
  totalAppointments: number;
  incomingCalls: number;
  outgoingCalls: number;
  generatedDate: string;
}) {
  const { 
    safeDutyDate, 
    safeUserName, 
    safeWorkOffice, 
    safeWorkHoursFrom, 
    safeWorkHoursTo,
    safeDailyWorkReport,
    patientList,
    totalAppointments,
    incomingCalls,
    outgoingCalls,
    generatedDate 
  } = props;

  const pdfStyles = StyleSheet.create({
    page: { padding: 20, fontFamily: 'Helvetica', fontSize: 9 },
    header: { marginBottom: 15, borderBottomWidth: 2, borderColor: '#333', paddingBottom: 8, alignItems: 'center' },
    headerTitle: { fontSize: 18, fontWeight: 'bold', marginBottom: 4 },
    infoSection: { marginBottom: 15, padding: 8, borderWidth: 1, borderColor: '#ccc', flexDirection: 'row', justifyContent: 'space-between', flexWrap: 'wrap' },
    infoItem: { fontSize: 8, marginBottom: 4 },
    stats: { marginBottom: 15, padding: 8, borderWidth: 1, borderColor: '#ccc', backgroundColor: '#f9f9f9', flexDirection: 'row', justifyContent: 'center', gap: 30 },
    statItem: { alignItems: 'center' },
    statValue: { fontSize: 14, fontWeight: 'bold', marginBottom: 2 },
    statLabel: { fontSize: 9, fontWeight: 'bold' },
    table: { marginTop: 10 },
    tableRow: { flexDirection: 'row', borderBottomWidth: 0.5, borderColor: '#333' },
    tableHeader: { flexDirection: 'row', borderBottomWidth: 1, borderColor: '#333', backgroundColor: '#f0f0f0', fontWeight: 'bold' },
    tableCell: { padding: 4, fontSize: 7, flex: 1, borderRightWidth: 0.5, borderColor: '#333', justifyContent: 'center', alignItems: 'center' },
    dailyReport: { marginTop: 15, padding: 8, borderWidth: 1, borderColor: '#ccc' },
    dailyReportTitle: { fontSize: 12, fontWeight: 'bold', marginBottom: 4 },
    dailyReportContent: { fontSize: 8, lineHeight: 1.4 },
    footer: { marginTop: 15, paddingTop: 8, borderTopWidth: 1, borderColor: '#ddd', alignItems: 'center', fontSize: 8, color: '#666' },
  });

  const s = pdfStyles;

  // 헤더
  const header = React.createElement(View, { style: s.header },
    React.createElement(Text, { style: s.headerTitle }, 'Patient Log'),
  );

  // 정보 섹션
  const infoSection = React.createElement(View, { style: s.infoSection },
    React.createElement(Text, { style: s.infoItem }, `Duty Date: ${safeDutyDate || '-'}`),
    React.createElement(Text, { style: s.infoItem }, `Name: ${safeUserName || '-'}`),
    React.createElement(Text, { style: s.infoItem }, `Work Office: ${safeWorkOffice || '-'}`),
    React.createElement(Text, { style: s.infoItem }, `Work Hours: ${convertTo12Hour(safeWorkHoursFrom)} - ${convertTo12Hour(safeWorkHoursTo)}`),
  );

  // 통계 섹션
  const stats = React.createElement(View, { style: s.stats },
    React.createElement(View, { style: s.statItem },
      React.createElement(Text, { style: s.statValue }, String(totalAppointments)),
      React.createElement(Text, { style: s.statLabel }, 'Total Appointments'),
    ),
    React.createElement(View, { style: s.statItem },
      React.createElement(Text, { style: s.statValue }, String(incomingCalls)),
      React.createElement(Text, { style: s.statLabel }, 'Incoming Calls'),
    ),
    React.createElement(View, { style: s.statItem },
      React.createElement(Text, { style: s.statValue }, String(outgoingCalls)),
      React.createElement(Text, { style: s.statLabel }, 'Outgoing Calls'),
    ),
  );

  // 테이블 헤더
  const tableHeader = React.createElement(View, { style: s.tableHeader },
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'No.')),
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Name')),
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Office')),
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Appt. Date')),
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Visit Type')),
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Call In')),
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Call Out')),
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Time')),
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Remark')),
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Other Duty')),
  );

  // 테이블 데이터 행
  const tableRows = patientList.map((row: any, index: number) => {
    const safeName = safeStr(row?.name, 50);
    const safeOffice = safeStr(row?.office, 50);
    const safeApptDate = safeStr(row?.appt_date || row?.apptDate, 20);
    const safeVisitType = safeStr(row?.visit_type || row?.visitType, 50);
    const safeTime = safeStr(row?.time, 20);
    const safeRemark = safeStr(row?.remark, 100);
    const safeOtherDuty = safeStr(row?.other_duty || row?.otherDuty, 100);
    // 체크박스 값 확인: call_in 또는 callIn이 true이면 'O', 아니면 빈 문자열
    const callIn = row?.call_in === true || row?.callIn === true;
    const callOut = row?.call_out === true || row?.callOut === true;

    return React.createElement(View, { key: index, style: s.tableRow },
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, String(index + 1))),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeName || '-')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeOffice || '-')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeApptDate || '-')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeVisitType || '-')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, callIn ? 'O' : '')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, callOut ? 'O' : '')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, convertTo12Hour(safeTime))),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeRemark || '-')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeOtherDuty || '-')),
    );
  });

  // 테이블
  const table = patientList.length > 0
    ? React.createElement(View, { style: s.table }, tableHeader, ...tableRows)
    : React.createElement(View, { style: { padding: 40, alignItems: 'center' } },
        React.createElement(Text, { style: { fontSize: 10, color: '#666' } }, 'No patient data recorded.'),
      );

  // Daily Work Report
  const dailyReport = safeDailyWorkReport
    ? React.createElement(View, { style: s.dailyReport },
        React.createElement(Text, { style: s.dailyReportTitle }, 'Daily Work Report'),
        React.createElement(Text, { style: s.dailyReportContent }, safeDailyWorkReport),
      )
    : null;

  // 푸터
  const footer = React.createElement(View, { style: s.footer },
    React.createElement(Text, null, `Generated: ${generatedDate}`),
  );

  return React.createElement(Document, null,
    React.createElement(Page, { size: 'A4', orientation: 'portrait', style: s.page }, 
      header, 
      infoSection, 
      stats, 
      table, 
      dailyReport,
      footer
    ),
  );
}

function PatientLogSystem(): React.ReactElement {
  
  // 기본 상태
  const [loading, setLoading] = useState(false);
  const [autoSaveStatus, setAutoSaveStatus] = useState(''); // 자동 저장 상태 표시
  
  // 마지막 저장된 데이터 추적 (dailyofficeduty 방식)
  const [lastSavedData, setLastSavedData] = useState({});
  const [submitStatus, setSubmitStatus] = useState('');
  const [progress, setProgress] = useState(0);
  const [previousDocId, setPreviousDocId] = useState<string | null>(null); // 이전 document ID 추적
  const [previousFormData, setPreviousFormData] = useState<{
    dutyDate: string;
    userName: string;
    workOffice: string;
    workHoursFrom: string;
    workHoursTo: string;
  } | null>(null); // 이전 basic information 추적
  const previousFormDataRef = useRef<{
    dutyDate: string;
    userName: string;
    workOffice: string;
    workHoursFrom: string;
    workHoursTo: string;
  } | null>(null); // 이전 basic information ref (최신 값 추적)
  const [isUnlocked, setIsUnlocked] = useState(false); // 아래 섹션 lock 상태
  const [isAuthorized, setIsAuthorized] = useState<boolean | null>(null); // null: 확인 중, true: 인증됨, false: 인증 실패
  const [userOfficesOptions, setuserOfficesOptions] = useState<string[]>([]); // 사용자의 offices 옵션들

  // Rate limiting을 위한 ref
  const lastUpdatePatientRowCall = useRef<number>(0);
  const lastSubmitCall = useRef<number>(0);

  // 폼 데이터 상태 (원본과 동일한 구조)
  const [formData, setFormData] = useState({
    dutyDate: new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' }),
    userName: '',
    workOffice: '',
    workHoursFrom: '',
    workHoursTo: '',
    dailyWorkReport: ''
  });

  // 환자 로그 상태 (원본과 동일한 필드명) - 기본 30행
  const [patientRows, setPatientRows] = useState(() => {
    return Array.from({ length: 30 }, (_, index) => ({
      id: index + 1,
      name: '',
      office: '',
      appt_date: '',
      visit_type: '',
      call_in: false,
      call_out: false,
      time: '',
      remark: '',
      other_duty: ''
    }));
  });

  // patientRows를 useMemo로 최적화

  // 실시간 카운트 계산 (단순화)
  const appointments = patientRows.filter(row => row.appt_date && row.name).length;
  const incomingCalls = patientRows.filter(row => row.call_in).length;
  const outgoingCalls = patientRows.filter(row => row.call_out).length;

  // Office 옵션 (단순화)
  const workOfficeOptions = ['Bernard', 'Call Center', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];
  const patientOfficeOptions = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];
  
  // Visit Type 옵션을 Office에 따라 동적으로 생성 (useCallback 최적화)
  const getVisitTypeOptions = useCallback((office: string): string[] => {
    if (office === 'Ortho') {
      return ['Adjustment', 'Bonding', 'Consult', 'Full Deband', 'Partial Deband', 'Records', 'Retainer Check', 'RPE Check'];
    } else {
      return ['Emergency', 'New Patient', 'RCRA', 'Recall', 'Tx'];
    }
  }, []);

  // Remark 옵션 (단순화)
  const remarkOptions = ['Disc', 'Elsewhere', 'LMA', 'LMW', 'NA', 'Not Interested', 'Wrong'];

  // Other Duty 옵션 (단순화)
  const otherDutyOptions = [
    'Accounts with Balances', 'Booking ASL Interpreters', 'Break', 'Confirming', 
    'Incoming Call Report', 'Insurance Verifications', 'Lunch', 'Marketing Data', 
    'Medi-cal Eligibility', 'Monthly Report', 'Nintendo Switch Raffle', 'One Week\'s Reconfirming', 
    'Other', 'Postcards', 'Refer a friend', 'Reviews', 'Routing Slips', 
    'Sending Replacement Staff', 'Training'
  ];

  // Document ID 생성 함수 (모든 Basic Information 포함)
  const generateDocId = (dutyDate: string, userName: string, workOffice: string, workHoursFrom: string, workHoursTo: string): string => {
    return `${dutyDate}_${userName}_${workOffice}_${workHoursFrom}_${workHoursTo}`.replace(/[^a-zA-Z0-9_-]/g, '_');
  };

  // Basic Information 완료 체크 함수
  const isBasicInfoComplete = () => {
    return formData.dutyDate && 
           formData.userName && 
           formData.workOffice && 
           formData.workHoursFrom && 
           formData.workHoursTo;
  };

  // 자동 저장 함수 (매우 빠른 저장)
  const autoSave = useCallback(async () => {
    // unlock되지 않았으면 저장하지 않음
    if (!isUnlocked) return;
    
    // Basic Information이 완료되지 않으면 저장하지 않음
    if (!isBasicInfoComplete()) return;

    // 이름이 공백으로 끝나면 아직 입력 중이므로 저장하지 않음
    if (formData.userName.trim().endsWith(' ')) {
      return;
    }

    try {
      const currentDocId = sanitizeFirebaseDocIdClient(generateDocId(formData.dutyDate, formData.userName, formData.workOffice, formData.workHoursFrom, formData.workHoursTo));
      
      // 이전 formData가 있고 basic information이 변경되었으면 기존 document를 찾아서 patientRows를 가져옴
      let rowsToSave = patientRows;
      let previousDocIdToDelete: string | null = null;
      
      const prevFormData = previousFormDataRef.current;
      if (prevFormData && 
          (prevFormData.dutyDate !== formData.dutyDate ||
           prevFormData.userName !== formData.userName ||
           prevFormData.workOffice !== formData.workOffice ||
           prevFormData.workHoursFrom !== formData.workHoursFrom ||
           prevFormData.workHoursTo !== formData.workHoursTo)) {
        try {
          const prevDocId = sanitizeFirebaseDocIdClient(generateDocId(
            prevFormData.dutyDate,
            prevFormData.userName,
            prevFormData.workOffice,
            prevFormData.workHoursFrom,
            prevFormData.workHoursTo
          ));
          
          const previousDocRef = doc(db, "patient-logs", prevDocId);
          const previousDocSnap = await getDoc(previousDocRef);
          
          if (previousDocSnap.exists()) {
            const previousData = previousDocSnap.data();
            // 기존 patientRows를 유지
            if (previousData.patientRows && Array.isArray(previousData.patientRows)) {
              rowsToSave = previousData.patientRows;
              // state에도 반영
              setPatientRows(previousData.patientRows);
            }
            
            // 나중에 삭제할 document ID 저장
            previousDocIdToDelete = prevDocId;
          }
        } catch (error) {
          // 기존 document를 찾지 못해도 계속 진행
        }
      }

      // 모든 row 저장 (빈 row도 포함하여 나중에 로드할 때 정확히 복원)
      const dataToSave = {
        ...formData,
        patientRows: rowsToSave, // 기존 patientRows 또는 현재 patientRows
        timestamp: new Date().toISOString(),
        autoSaved: true
      };

      // Firebase 데이터 보안 검증
      const safeDataToSave = sanitizeFirebaseDataClient(dataToSave);
      
      // 자동 저장 시작 (메시지 표시 안함)
      
      // Firebase 저장 (비동기 처리로 UI 블로킹 방지)
      setDoc(doc(db, "patient-logs", currentDocId), safeDataToSave)
        .then(async () => {
          // 기존 document 삭제 (새 document 저장 성공 후)
          if (previousDocIdToDelete && previousDocIdToDelete !== currentDocId) {
            try {
              await deleteDoc(doc(db, "patient-logs", previousDocIdToDelete));
            } catch (deleteError) {
              // 삭제 실패해도 계속 진행
            }
          }
          
          // 현재 basic information을 이전 값으로 저장 (state와 ref 모두 업데이트)
          const newPreviousFormData = {
            dutyDate: formData.dutyDate,
            userName: formData.userName,
            workOffice: formData.workOffice,
            workHoursFrom: formData.workHoursFrom,
            workHoursTo: formData.workHoursTo
          };
          setPreviousFormData(newPreviousFormData);
          previousFormDataRef.current = newPreviousFormData;
          setPreviousDocId(currentDocId);
        })
        .catch((error) => {
          setAutoSaveStatus('❌ Save failed');
          
          // 2초 후 상태 메시지 제거
          setTimeout(() => {
            setAutoSaveStatus('');
          }, 2000);
        });
      
    } catch (error) {
      // 에러 발생 시 조용히 처리
    }
  }, [formData, patientRows, isUnlocked]);

  // 데이터 변경 시 자동 저장 (입력 완료 후 저장)
  useEffect(() => {
    // 초기 로드 시에는 저장하지 않음
    // unlock되지 않았으면 저장하지 않음
    if (!isUnlocked) return;
    
    // Basic Information이 완료되지 않았으면 저장하지 않음
    if (!isBasicInfoComplete()) return;
    
    // 이름이 공백으로 끝나면 아직 입력 중이므로 저장하지 않음
    if (formData.userName.trim().endsWith(' ')) {
      return;
    }

    const timeoutId = setTimeout(() => {
      autoSave();
    }, 2000); // 2초 debounce - 입력이 완료된 후 저장

    return () => clearTimeout(timeoutId);
  }, [formData, patientRows, isUnlocked, autoSave]);
  
  // 입력 필드에서 포커스를 잃을 때 저장하는 함수
  const handleFieldBlur = useCallback(() => {
    // unlock되지 않았으면 저장하지 않음
    if (!isUnlocked) return;
    
    // 포커스를 잃을 때 즉시 저장 (debounce 없이)
    // Basic Information이 완료되었는지 확인
    const isComplete = formData.dutyDate && 
                       formData.userName && 
                       formData.workOffice && 
                       formData.workHoursFrom && 
                       formData.workHoursTo;
    if (isComplete) {
      autoSave();
    }
  }, [formData, autoSave, isUnlocked]);

  // 저장된 데이터를 현재 환자 로그에 로드하는 함수 (최적화 + 보안 강화)
  const loadExistingData = async () => {
    if (!isBasicInfoComplete()) {
      return;
    }

    try {
      // 인증 상태 확인
      const auth = getAuth();
      const currentUser = auth.currentUser;
      
      const currentDocId = sanitizeFirebaseDocIdClient(generateDocId(formData.dutyDate, formData.userName, formData.workOffice, formData.workHoursFrom, formData.workHoursTo));
      
      // 이전 formData가 있고 basic information이 변경되었으면 이전 document에서 patientRows를 가져옴
      if (previousFormData && 
          (previousFormData.dutyDate !== formData.dutyDate ||
           previousFormData.userName !== formData.userName ||
           previousFormData.workOffice !== formData.workOffice ||
           previousFormData.workHoursFrom !== formData.workHoursFrom ||
           previousFormData.workHoursTo !== formData.workHoursTo)) {
        try {
          const prevDocId = sanitizeFirebaseDocIdClient(generateDocId(
            previousFormData.dutyDate,
            previousFormData.userName,
            previousFormData.workOffice,
            previousFormData.workHoursFrom,
            previousFormData.workHoursTo
          ));
          
          const previousDocRef = doc(db, "patient-logs", prevDocId);
          const previousDocSnap = await getDoc(previousDocRef);
          
          if (previousDocSnap.exists()) {
            const previousData = previousDocSnap.data();
            
            // 보안 강화: 데이터 소유권 확인
            if (currentUser && previousData.userId && previousData.userId !== currentUser.uid) {
              // 다른 사용자의 데이터 접근 시도 차단
            } else if (previousData.patientRows && Array.isArray(previousData.patientRows)) {
              // 기존 patientRows를 state에 반영
              const loadedRows = previousData.patientRows.map((row: any, index: number) => ({
                ...row,
                id: index + 1
              }));
              
              // 저장된 row 개수와 30 중 더 큰 값으로 배열 생성 (30개 이상도 유지, 최소 30개 보장)
              const minRows = Math.max(loadedRows.length, 30);
              const newRows = Array.from({ length: minRows }, (_, index) => {
                if (index < loadedRows.length) {
                  return loadedRows[index];
                }
                return {
                  id: index + 1,
                  name: '',
                  office: '',
                  appt_date: '',
                  visit_type: '',
                  call_in: false,
                  call_out: false,
                  time: '',
                  remark: '',
                  other_duty: ''
                };
              });
              
              setPatientRows(newRows);
            }
          }
        } catch (error) {
          // 이전 document를 찾지 못해도 계속 진행
        }
      }
      
      const docRef = doc(db, "patient-logs", currentDocId);
      
      // 직접 document 참조로 조회 (전체 collection 스캔 방지)
      const docSnap = await getDoc(docRef);

      if (docSnap.exists()) {
        const matchingLog = docSnap.data();
        
        // 보안 강화: 데이터 소유권 확인 (Firebase Security Rules와 함께)
        if (currentUser && matchingLog.userId && matchingLog.userId !== currentUser.uid) {
          // 다른 사용자의 데이터는 로드하지 않음 (Security Rules에서도 차단됨)
          return;
        }
        
        // 현재 docId를 previousDocId로 설정
        setPreviousDocId(currentDocId);
        // 현재 basic information을 previousFormData로 설정 (state와 ref 모두 업데이트)
        const newPreviousFormData = {
          dutyDate: formData.dutyDate,
          userName: formData.userName,
          workOffice: formData.workOffice,
          workHoursFrom: formData.workHoursFrom,
          workHoursTo: formData.workHoursTo
        };
        setPreviousFormData(newPreviousFormData);
        previousFormDataRef.current = newPreviousFormData;
        
        if (matchingLog && matchingLog.patientRows) {
          // 기존 저장된 환자 로그를 현재 patientRows에 로드
          const loadedRows = matchingLog.patientRows.map((row: any, index: number) => ({
            ...row,
            id: index + 1
          }));
          
          // 저장된 row 개수와 30 중 더 큰 값으로 배열 생성 (30개 이상도 유지, 최소 30개 보장)
          const minRows = Math.max(loadedRows.length, 30);
          const newRows = Array.from({ length: minRows }, (_, index) => {
            if (index < loadedRows.length) {
              return loadedRows[index];
            }
            return {
              id: index + 1,
              name: '',
              office: '',
              appt_date: '',
              visit_type: '',
              call_in: false,
              call_out: false,
              time: '',
              remark: '',
              other_duty: ''
            };
          });
          
          setPatientRows(newRows);
          
          // Daily Work Report도 로드
          if (matchingLog.dailyWorkReport) {
            setFormData(prev => ({
              ...prev,
              dailyWorkReport: matchingLog.dailyWorkReport
            }));
          }
        }
      } else {
        // 문서가 없으면 previousDocId를 현재 docId로 설정 (새 문서 생성 준비)
        setPreviousDocId(currentDocId);
        // 현재 basic information을 previousFormData로 설정 (state와 ref 모두 업데이트)
        const newPreviousFormData = {
          dutyDate: formData.dutyDate,
          userName: formData.userName,
          workOffice: formData.workOffice,
          workHoursFrom: formData.workHoursFrom,
          workHoursTo: formData.workHoursTo
        };
        setPreviousFormData(newPreviousFormData);
        previousFormDataRef.current = newPreviousFormData;
      }
    } catch (error) {
      // 에러 발생 시 조용히 처리
    }
  };

  // patient-logs에서 같은 userName, workOffice인데 다른 날짜(dutyDate)에 데이터가 있는지만 확인 (show-noshow 확인 안 함)
  const checkUnsubmittedData = useCallback(async () => {
    if (!formData.userName || !formData.workOffice) return;

    try {
      const logsQuery = query(
        collection(db, "patient-logs"),
        where("userName", "==", formData.userName),
        where("workOffice", "==", formData.workOffice)
      );
      const logsSnapshot = await getDocs(logsQuery);
      const otherDates: string[] = [];

      for (const logDoc of logsSnapshot.docs) {
        const logDate = logDoc.data().dutyDate;
        if (logDate && logDate !== formData.dutyDate) otherDates.push(logDate);
      }

      if (otherDates.length > 0) {
        const datesList = [...new Set(otherDates)].sort().join(', ');
        alert(`You have data for these other dates: ${datesList}`);
      }
    } catch (error) {
      // 에러 발생 시 조용히 처리
    }
  }, [formData.userName, formData.workOffice, formData.dutyDate]);

  // 기본 정보가 입력되면 기존 데이터 로드 (unlock 후에만)
  useEffect(() => {
    // unlock되지 않았으면 로드하지 않음
    if (!isUnlocked) return;
    
    const timeoutId = setTimeout(() => {
      loadExistingData();
    }, 50); // 0.05초 debounce로 매우 빠르게

    return () => clearTimeout(timeoutId);
  }, [formData.dutyDate, formData.userName, formData.workOffice, formData.workHoursFrom, formData.workHoursTo, isUnlocked]);

  // unlock된 상태에서 basic information이 변경되면 다시 lock
  useEffect(() => {
    // unlock되지 않았으면 실행하지 않음
    if (!isUnlocked) return;
    
    // previousFormData가 없으면 (처음 unlock한 경우) 실행하지 않음
    const prevFormData = previousFormDataRef.current;
    if (!prevFormData) return;
    
    // basic information이 변경되었는지 확인
    if (prevFormData.dutyDate !== formData.dutyDate ||
        prevFormData.userName !== formData.userName ||
        prevFormData.workOffice !== formData.workOffice ||
        prevFormData.workHoursFrom !== formData.workHoursFrom ||
        prevFormData.workHoursTo !== formData.workHoursTo) {
      // basic information이 변경되었으면 다시 lock
      setIsUnlocked(false);
      
      // patientRows 초기화
      setPatientRows(() => {
        return Array.from({ length: 30 }, (_, index) => ({
          id: index + 1,
          name: '',
          office: '',
          appt_date: '',
          visit_type: '',
          call_in: false,
          call_out: false,
          time: '',
          remark: '',
          other_duty: ''
        }));
      });
      
      // dailyWorkReport 초기화
      setFormData(prev => ({
        ...prev,
        dailyWorkReport: ''
      }));
      
      // previousFormData 초기화
      setPreviousFormData(null);
      previousFormDataRef.current = null;
      setPreviousDocId(null);
    }
  }, [formData.dutyDate, formData.userName, formData.workOffice, formData.workHoursFrom, formData.workHoursTo, isUnlocked]);

  // 이름이나 오피스가 변경되면 미제출 데이터 확인
  useEffect(() => {
    const timeoutId = setTimeout(() => {
      checkUnsubmittedData();
    }, 500); // 0.5초 debounce

    return () => clearTimeout(timeoutId);
  }, [formData.userName, formData.workOffice, checkUnsubmittedData]);


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
          const officesArray = Array.isArray(userData.offices) 
            ? userData.offices 
            : [userData.offices];
          
          // workOfficeOptions에 포함된 값들만 필터링
          const validOptions = officesArray.filter((g: string) => workOfficeOptions.includes(g));
          
          if (validOptions.length > 0) {
            setuserOfficesOptions(validOptions);
            // 단일 값이면 자동 선택
            if (validOptions.length === 1) {
              setFormData(prev => ({ ...prev, workOffice: validOptions[0] }));
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

  // 저장된 로그 불러오기

  // 입력 값 검증 함수
  const validateInput = useCallback((field: string, value: any): any => {
    // 문자열 필드 길이 제한
    if (typeof value === 'string') {
      const maxLengths: { [key: string]: number } = {
        dutyDate: 50,
        userName: 100,
        workOffice: 100,
        workHoursFrom: 20,
        workHoursTo: 20,
        dailyWorkReport: 2000,
        name: 100,
        office: 50,
        appt_date: 20,
        visit_type: 50,
        time: 20,
        remark: 200,
        other_duty: 200
      };
      
      const maxLength = maxLengths[field] || 500;
      if (value.length > maxLength) {
        return value.slice(0, maxLength);
      }
      
      // 제어 문자 제거 (XSS 방지)
      return value.replace(/[\x00-\x1F\x7F-\x9F]/g, '');
    }
    
    return value;
  }, []);

  // 폼 데이터 업데이트 (입력 검증 포함)
  const updateFormData = useCallback((field: string, value: any) => {
    const validatedValue = validateInput(field, value);
    setFormData(prev => {
      // 값이 같으면 업데이트하지 않음
      if ((prev as any)[field] === validatedValue) return prev;
      return { ...prev, [field]: validatedValue };
    });
  }, [validateInput]);


  // 환자 행 추가 (useCallback 최적화)
  const addPatientRow = useCallback(() => {
    setPatientRows(prevRows => {
      const newId = Math.max(...prevRows.map(row => row.id)) + 1;
      return [...prevRows, {
        id: newId,
        name: '',
        office: '',
        appt_date: '',
        visit_type: '',
        call_in: false,
        call_out: false,
        time: '',
        remark: '',
        other_duty: ''
      }];
    });
  }, []);

  // 환자 행 삭제 (useCallback 최적화)
  const removePatientRow = useCallback((id: number) => {
    setPatientRows(prevRows => {
      if (prevRows.length > 1) {
        return prevRows.filter(row => row.id !== id);
      }
      return prevRows;
    });
  }, []);

  // 환자 행 업데이트 (입력 검증 포함, Rate limiting 적용)
  const updatePatientRow = useCallback((id: number, field: string, value: any) => {
    // Rate limiting: 입력 반응성을 위해 완화된 제한 적용
    // (자동 저장은 별도 debounce로 처리되므로 입력 자체는 빠르게 반응)
    const now = Date.now();
    const rowKey = `lastUpdate_${id}_${field}`;
    const lastCall = (window as any)[rowKey] || 0;
    
    // 전역 rate limiting: 모든 업데이트에 대해 50ms 제한 (입력 반응성 향상)
    if (now - lastUpdatePatientRowCall.current < 50) {
      return;
    }
    lastUpdatePatientRowCall.current = now;

    // 개별 행 rate limiting: 동일 필드에 대해 100ms 제한 (입력 반응성 향상)
    if (now - lastCall < 100) {
      return;
    }
    (window as any)[rowKey] = now;

    // 입력 값 검증
    const validatedValue = validateInput(field, value);
    
    setPatientRows(prevRows => {
      const rowIndex = prevRows.findIndex(row => row.id === id);
      
      if (rowIndex === -1) {
        return prevRows;
      }
      
      const row = prevRows[rowIndex];
      
      // 값이 같으면 업데이트하지 않음
      if ((row as any)[field] === validatedValue) {
        return prevRows;
      }
      
      const updatedRow = { ...row, [field]: validatedValue };
      
      // Office가 변경되면 visit_type을 초기화
      if (field === 'office' && row.office !== validatedValue) {
        updatedRow.visit_type = '';
      }
      // Call In 또는 Call Out이 체크되면 현재 시간을 Time에 자동 입력
      if ((field === 'call_in' || field === 'call_out') && validatedValue === true) {
        const now = new Date();
        const timeString = now.toTimeString().slice(0, 5);
        updatedRow.time = timeString;
      }
      // Call In과 Call Out이 모두 체크 해제되면 Time을 비움
      if ((field === 'call_in' && validatedValue === false && !row.call_out) || 
          (field === 'call_out' && validatedValue === false && !row.call_in)) {
        updatedRow.time = '';
      }
      
      const newRows = [...prevRows];
      newRows[rowIndex] = updatedRow;
      return newRows;
    });
  }, [validateInput]);

  // 폼 초기화 함수
  const resetForm = () => {
    setFormData({
      dutyDate: new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' }),
      userName: '',
      workOffice: '',
      workHoursFrom: '',
      workHoursTo: '',
      dailyWorkReport: ''
    });
    setPatientRows([{
      id: 1,
      name: '',
      office: '',
      appt_date: '',
      visit_type: '',
      call_in: false,
      call_out: false,
      time: '',
      remark: '',
      other_duty: ''
    }]);
    setLastSavedData({});
    setPreviousDocId(null);
    setPreviousFormData(null);
    previousFormDataRef.current = null;
    setIsUnlocked(false);
  };

  // PDF 생성 및 제출 (Rate limiting 적용)
  const handleSubmit = async () => {
    // Rate limiting: 최근 3초 내 호출 방지 (PDF 생성은 무거운 작업)
    // (실수로 두 번 클릭하는 것을 방지하되, 사용자 경험을 해치지 않도록)
    const now = Date.now();
    if (now - lastSubmitCall.current < 3000) {
      alert('⚠️ Please try again.');
      return;
    }
    lastSubmitCall.current = now;

    // 이미 제출 중이면 중복 호출 방지
    if (loading) {
      return;
    }

    if (!isBasicInfoComplete()) {
      alert('⚠️ Please fill in all Basic Information fields.');
      return;
    }

    // 인증 상태 확인
    const auth = getAuth();
    const currentUser = auth.currentUser;
    if (!currentUser) {
      alert('⚠️ Please log in.');
      return;
    }
    
    // 로그 제거 (보안 강화)

    try {
      setLoading(true);
      setSubmitStatus('Saving...');
      setProgress(10);

      const currentDocId = sanitizeFirebaseDocIdClient(generateDocId(formData.dutyDate, formData.userName, formData.workOffice, formData.workHoursFrom, formData.workHoursTo));

      // 1. 클라이언트 사이드에서 PDF 생성
      setSubmitStatus('Submitting...');
      setProgress(30);
      
      // PDF용 데이터 준비
      const patientListForPdf = patientRows.filter(row => 
        row.name || row.office || row.appt_date || row.visit_type || 
        row.call_in || row.call_out || row.time || row.remark || row.other_duty
      );
      
      // 통계 계산
      const totalAppointments = patientListForPdf.filter(row => row.appt_date && row.name).length;
      const incomingCalls = patientListForPdf.filter(row => row.call_in).length;
      const outgoingCalls = patientListForPdf.filter(row => row.call_out).length;
      
      // 날짜 포맷팅
      const generatedDate = new Date().toLocaleDateString('en-US', { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        hour12: true
      });
      
      // 데이터 sanitize
      const safeDutyDate = (formData.dutyDate || '').trim().slice(0, 50).replace(/[<>]/g, '');
      const safeUserName = (formData.userName || '').trim().slice(0, 100).replace(/[<>]/g, '');
      const safeWorkOffice = (formData.workOffice || '').trim().slice(0, 100).replace(/[<>]/g, '');
      const safeWorkHoursFrom = (formData.workHoursFrom || '').trim().slice(0, 20).replace(/[<>]/g, '');
      const safeWorkHoursTo = (formData.workHoursTo || '').trim().slice(0, 20).replace(/[<>]/g, '');
      const safeDailyWorkReport = (formData.dailyWorkReport || '').trim().slice(0, 2000).replace(/[<>]/g, '');
      
      // PDF 문서 생성
      setSubmitStatus('Processing...');
      setProgress(50);
      
      const pdfDoc = createPatientLogPDFDocument({
        safeDutyDate,
        safeUserName,
        safeWorkOffice,
        safeWorkHoursFrom,
        safeWorkHoursTo,
        safeDailyWorkReport,
        patientList: patientListForPdf,
        totalAppointments,
        incomingCalls,
        outgoingCalls,
        generatedDate,
      });
      
      // PDF blob 생성
      setSubmitStatus('Processing...');
      setProgress(60);
      
      const pdfBlob = await pdf(pdfDoc).toBlob();
      
      // 파일명 생성 (강화된 검증)
      const safeDate = (formData.dutyDate || new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' }))
        .replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 50);
      const safeName = (formData.userName || 'Unknown')
        .replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 50);
      const safeOffice = (formData.workOffice || 'Unknown')
        .replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 50);
      const tsNow = new Date();
      const tsLaTime = new Date(tsNow.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
      let tsHours = tsLaTime.getHours();
      const tsMinutes = tsLaTime.getMinutes();
      const tsAmpm = tsHours >= 12 ? 'pm' : 'am';
      tsHours = tsHours % 12;
      tsHours = tsHours ? tsHours : 12;
      const timeStamp = `${tsHours}${tsMinutes.toString().padStart(2, '0')}${tsAmpm}`;
      const filename = `7) ${safeDate}_${safeOffice}_${safeName}_Patient Log_${timeStamp}.pdf`.slice(0, 255);
      const date = safeDate;
      const name = safeName;
      const office = safeOffice;
      
      // Firebase에 자동 저장
      setSubmitStatus('Saving...');
      setProgress(70);
      
      try {
        // PDF를 Firebase Storage에 저장 (endofday-pdfs 경로 유지)
        const storage = getStorage();
        const storageRef = ref(storage, `endofday-pdfs/${office}/${date}/${filename}`);
        
        await uploadBytes(storageRef, pdfBlob);
        
        // patient-logs 컬렉션에서 제출한 문서 삭제
        try {
          await deleteDoc(doc(db, 'patient-logs', currentDocId));
        } catch (deleteError) {
          // 문서가 없어도 무시
        }
        
        // show-noshow 컬렉션: document는 appt_date만, 같은 appt_date면 같은 문서 아래로 merge
        const rowsWithDate = patientListForPdf.filter(row => row.appt_date && row.appt_date.trim() !== '');
        const byApptDate = new Map<string, { name: string; office: string }[]>();
        for (const row of rowsWithDate) {
          const d = (row.appt_date || '').trim();
          if (!byApptDate.has(d)) byApptDate.set(d, []);
          byApptDate.get(d)!.push({
            name: safeStr(row.name, 100),
            office: safeStr(row.office, 100),
          });
        }
        // show-noshow: appt_date + patients만 (submissions, duty_date, name, office, createdAt 없음)
        for (const [apptDate, patients] of byApptDate) {
          const showDocId = apptDate.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 1500);
          const docRef = doc(db, 'show-noshow', showDocId);
          const existing = await getDoc(docRef);
          const existingPatients = existing.exists() && Array.isArray(existing.data()?.patients)
            ? existing.data()!.patients
            : [];
          const mergedPatients = [...existingPatients, ...patients];
          await setDoc(docRef, sanitizeFirebaseDataClient({
            appt_date: apptDate,
            patients: mergedPatients,
          }));
        }
        if (byApptDate.size === 0) {
          const showDocId = formData.dutyDate.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 1500);
          const docRef = doc(db, 'show-noshow', showDocId);
          const existing = await getDoc(docRef);
          const existingPatients = existing.exists() && Array.isArray(existing.data()?.patients)
            ? existing.data()!.patients
            : [];
          await setDoc(docRef, sanitizeFirebaseDataClient({
            appt_date: formData.dutyDate,
            patients: existingPatients,
          }));
        }
        
        setSubmitStatus('Submitted Successfully!');
        setProgress(100);
        
        // 폼 초기화
        resetForm();
        
        setTimeout(() => {
          setLoading(false);
          setSubmitStatus('');
          setProgress(0);
        }, 2000);
      } catch (storageError: any) {
        // 에러 메시지 추출 및 보안 강화
        const errorMessage = storageError?.message || 'Error';
        
        // 민감한 정보 필터링
        const sensitiveKeywords = ['password', 'token', 'secret', 'key', 'credential', 'auth', 'login', 'session', 'cookie', 'bearer', 'jwt', 'api', 'apikey'];
        const hasSensitiveInfo = sensitiveKeywords.some(keyword => 
          errorMessage.toLowerCase().includes(keyword.toLowerCase())
        );
        
        const safeErrorMessage = hasSensitiveInfo 
          ? 'Error.' 
          : (errorMessage.length > 100 ? errorMessage.substring(0, 100) + '...' : errorMessage).replace(/[<>\"'&]/g, '');
        
        alert(`Error: ${safeErrorMessage}`);
        setLoading(false);
        setSubmitStatus('');
        setProgress(0);
      }

    } catch (error) {
      // 에러 메시지 추출 및 보안 강화
      const errorMessage = (error as any).message || 'Error';
      
      // 민감한 정보 필터링
      const sensitiveKeywords = ['password', 'token', 'secret', 'key', 'credential', 'auth', 'login', 'session', 'cookie', 'bearer', 'jwt', 'api', 'apikey'];
      const hasSensitiveInfo = sensitiveKeywords.some(keyword => 
        errorMessage.toLowerCase().includes(keyword.toLowerCase())
      );
      
      const safeErrorMessage = hasSensitiveInfo 
        ? 'An error occurred while submitting.' 
        : (errorMessage.length > 100 ? errorMessage.substring(0, 100) + '...' : errorMessage).replace(/[<>\"'&]/g, '');
      
      // 화면에 에러 메시지 표시
      setSubmitStatus('❌ Submission failed: ' + safeErrorMessage);
      setProgress(0);
      
      // 사용자에게 alert로도 표시 (콘솔이 막혀있으므로)
      alert('❌ Submission failed. Please try again.');
      
      setTimeout(() => {
        setLoading(false);
        setSubmitStatus('');
      }, 5000); // 5초로 연장하여 사용자가 메시지를 읽을 수 있도록
    }
  };

  // 제출 중 브라우저 네비게이션 방지
  useEffect(() => {
    const handleBeforeUnload = (e: any) => {
      if (loading) {
        e.preventDefault();
        e.returnValue = '';
      }
    };

    const handlePopState = (e: any) => {
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

  // 원본 HTML 스타일을 React 스타일로 변환
  const containerStyle = {
    maxWidth: '1500px',
    margin: '40px auto',
    padding: '30px',
    backgroundColor: 'rgba(255, 255, 255, 0.95)',
    backdropFilter: 'blur(5px)',
    borderRadius: '12px',
    boxShadow: '0 8px 16px rgba(0, 0, 0, 0.15)',
    fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
    color: '#023047',
    lineHeight: '1.6'
  };

  const bodyStyle = {
    padding: '20px',
    background: 'linear-gradient(to bottom, #a2d2ff, #f0f8ff)',
    minHeight: '100vh'
  };

  const headerStyle = {
    color: '#0077B6',
    textAlign: 'center',
    marginBottom: '30px',
    paddingBottom: '10px',
    borderBottom: '2px solid #BDE0FE',
    fontSize: '2.5em',
    fontWeight: 'bold'
  };

  const sectionStyle = {
    marginBottom: '30px',
    padding: '20px',
    backgroundColor: '#f0f8ff',
    borderRadius: '8px',
    border: '1px solid #BDE0FE'
  };

  const inputStyle = {
    padding: '8px 12px',
    border: '1px solid #BDE0FE',
    borderRadius: '4px',
    fontSize: '1em',
    backgroundColor: 'white',
    color: '#023047',
    width: '100%'
  };

  const buttonStyle = {
    backgroundColor: '#0077B6',
    color: 'white',
    border: 'none',
    padding: '12px 24px',
    borderRadius: '6px',
    fontSize: '1em',
    fontWeight: 'bold',
    cursor: 'pointer',
    margin: '5px',
    transition: 'all 0.3s ease'
  };

  const tableStyle = {
    width: '100%',
    borderCollapse: 'collapse',
    marginTop: '20px',
    backgroundColor: 'white',
    borderRadius: '8px',
    overflow: 'hidden',
    boxShadow: '0 4px 8px rgba(0, 0, 0, 0.1)'
  };

  // 인증 확인 중이거나 인증 실패 시 로딩 화면 표시
  if (isAuthorized === null) {
    return (
      <div style={bodyStyle}>
        <div style={{
          display: 'flex',
          justifyContent: 'center',
          alignItems: 'center',
          height: '100vh',
          background: 'linear-gradient(to bottom, #a2d2ff, #f0f8ff)',
          fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
        }}>
          <div style={{ textAlign: 'center' }}>
            <div style={{ fontSize: '24px', marginBottom: '20px' }}>🔐</div>
            <div style={{ fontSize: '18px', color: '#023047' }}>Verifying authentication...</div>
          </div>
        </div>
      </div>
    );
  }

  if (isAuthorized === false) {
    return (
      <div style={bodyStyle}>
        <div style={{
          display: 'flex',
          justifyContent: 'center',
          alignItems: 'center',
          height: '100vh',
          background: 'linear-gradient(to bottom, #a2d2ff, #f0f8ff)',
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
      `}</style>
      <div style={bodyStyle}>
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
              borderRadius: "20px",
              textAlign: "center",
              boxShadow: "0 20px 40px rgba(0, 0, 0, 0.3)",
              maxWidth: "400px",
              width: "90%"
            }}>
              <div style={{
                fontSize: "18px",
                fontWeight: "600",
                color: "#4a6fa1",
                marginBottom: "20px"
              }}>
                {submitStatus}
              </div>
              {progress > 0 && (
                <div style={{
                  width: "100%",
                  height: "8px",
                  backgroundColor: "#f0f0f0",
                  borderRadius: "4px",
                  overflow: "hidden",
                  marginBottom: "10px"
                }}>
                  <div style={{
                    width: `${progress}%`,
                    height: "100%",
                    background: "linear-gradient(90deg, #4a90e2, #51cf66)",
                    transition: "width 0.3s ease",
                    borderRadius: "4px"
                  }} />
                </div>
              )}
              <div style={{
                fontSize: "14px",
                color: "#666",
                marginTop: "10px"
              }}>
                {progress}%
              </div>
            </div>
          </div>
        )}

        <div style={containerStyle}>
        {/* 헤더 */}
        <div style={{ position: 'relative' }}>
        <h1 style={{ 
          color: '#0077B6', 
          textAlign: 'center', 
          marginBottom: '20px', 
          fontSize: '2.5rem', 
          fontWeight: 'bold',
          textShadow: '2px 2px 4px rgba(0,0,0,0.1)'
        }}>Patient Log</h1>
          {autoSaveStatus && (
            <div style={{
              position: 'absolute',
              top: '10px',
              right: '10px',
              padding: '8px 16px',
              backgroundColor: autoSaveStatus.includes('실패') ? '#ff6b6b' : '#51cf66',
              color: 'white',
              borderRadius: '20px',
              fontSize: '14px',
              fontWeight: 'bold',
              boxShadow: '0 2px 8px rgba(0,0,0,0.2)',
              zIndex: 1000
            }}>
              {autoSaveStatus}
            </div>
          )}
        </div>

        {/* 기본 정보 섹션 */}
        <div style={sectionStyle}>
          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap', marginBottom: '20px' }}>
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Duty Date:
              </label>
              <input
                type="date"
                value={formData.dutyDate}
                onChange={(e) => updateFormData('dutyDate', e.target.value)}
                onBlur={handleFieldBlur}
                style={inputStyle}
                required
              />
            </div>
            
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Name:
              </label>
              <input
                type="text"
                value={formData.userName}
                onChange={(e) => updateFormData('userName', e.target.value)}
                onBlur={handleFieldBlur}
                placeholder="Enter your full name"
                style={inputStyle}
                required
              />
            </div>

            {userOfficesOptions.length > 0 && (
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Work Office:
              </label>
              {userOfficesOptions.length === 1 ? (
                <div style={{
                  ...inputStyle,
                  display: 'flex',
                  alignItems: 'center',
                  backgroundColor: '#e9ecef',
                  fontWeight: '600',
                  color: '#023047'
                }}>
                  {formData.workOffice}
                </div>
              ) : (
                <select
                  value={formData.workOffice}
                  onChange={(e) => updateFormData('workOffice', e.target.value)}
                  onBlur={handleFieldBlur}
                  style={inputStyle}
                  required
                >
                  <option value="">Select Office</option>
                  {userOfficesOptions.map(office => (
                    <option key={office} value={office}>{office}</option>
                  ))}
                </select>
              )}
            </div>
            )}
          </div>

          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap', marginBottom: '20px' }}>
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Work Hours From:
              </label>
              <input
                type="time"
                value={formData.workHoursFrom}
                onChange={(e) => updateFormData('workHoursFrom', e.target.value)}
                onBlur={handleFieldBlur}
                style={inputStyle}
                required
              />
            </div>
            
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Work Hours To:
              </label>
              <input
                type="time"
                value={formData.workHoursTo}
                onChange={(e) => updateFormData('workHoursTo', e.target.value)}
                onBlur={handleFieldBlur}
                style={inputStyle}
                required
              />
            </div>
          </div>

          {/* Unlock 버튼 */}
          {!isUnlocked && (
            <div style={{ display: 'flex', justifyContent: 'center', marginTop: '20px' }}>
              <button
                onClick={async () => {
                  if (isBasicInfoComplete()) {
                    // unlock 시 기존 데이터 로드
                    await loadExistingData();
                    // unlock 상태로 변경
                    setIsUnlocked(true);
                    // unlock 후 저장 시작 (기존 데이터가 없으면 새 document로 저장)
                    setTimeout(() => {
                      autoSave();
                    }, 500); // 데이터 로드 후 저장
                  } else {
                    alert('⚠️ Please fill in all Basic Information fields.');
                  }
                }}
                style={{
                  ...buttonStyle,
                  backgroundColor: '#0077B6',
                  padding: '12px 30px',
                  fontSize: '1.1em'
                }}
              >
                🔓 Unlock Patient Log
              </button>
            </div>
          )}
        </div>

        {/* 환자 로그 테이블 */}
        <div style={sectionStyle}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '15px' }}>
              {autoSaveStatus && (
                <span style={{ 
                  fontSize: '12px', 
                  color: autoSaveStatus.includes('❌') ? '#dc3545' : autoSaveStatus.includes('💾') ? '#007bff' : '#28a745',
                  fontWeight: 'bold',
                  padding: '4px 8px',
                  backgroundColor: autoSaveStatus.includes('❌') ? '#f8d7da' : autoSaveStatus.includes('💾') ? '#d1ecf1' : '#e8f5e8',
                  borderRadius: '12px',
                  border: `1px solid ${autoSaveStatus.includes('❌') ? '#dc3545' : autoSaveStatus.includes('💾') ? '#007bff' : '#28a745'}`,
                  marginLeft: '10px'
                }}>
                  {autoSaveStatus}
                </span>
              )}
              {formData.userName && formData.userName.trim().endsWith(' ') && (
                <span style={{ 
                  fontSize: '12px', 
                  color: '#ffc107',
                  fontWeight: 'bold',
                  padding: '4px 8px',
                  backgroundColor: '#fff3cd',
                  borderRadius: '12px',
                  border: '1px solid #ffc107',
                  marginLeft: '10px'
                }}>
                  ⏳ Typing...
                </span>
              )}
            </div>
          </div>

          {/* 실시간 카운트 표시 */}
          <div style={{ 
            display: 'flex', 
            justifyContent: 'center',
            gap: '30px', 
            marginBottom: '20px',
            padding: '15px',
            backgroundColor: '#e3f2fd',
            borderRadius: '8px',
            border: '1px solid #bbdefb'
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span style={{ fontWeight: 'bold', color: '#1565c0' }}>📅 Appointments:</span>
              <span style={{ 
                fontSize: '18px', 
                fontWeight: 'bold', 
                color: '#0077B6',
                minWidth: '20px'
              }}>
                {appointments}
              </span>
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span style={{ fontWeight: 'bold', color: '#1565c0' }}>📞 Incoming Calls:</span>
              <span style={{ 
                fontSize: '18px', 
                fontWeight: 'bold', 
                color: '#28a745',
                minWidth: '20px'
              }}>
                {incomingCalls}
              </span>
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span style={{ fontWeight: 'bold', color: '#1565c0' }}>📱 Outgoing Calls:</span>
              <span style={{ 
                fontSize: '18px', 
                fontWeight: 'bold', 
                color: '#ff6b35',
                minWidth: '20px'
              }}>
                {outgoingCalls}
              </span>
            </div>
          </div>

          {/* Unlock 체크 후 테이블 표시 */}
          {!isUnlocked ? (
            <div style={{
              padding: '40px',
              textAlign: 'center',
              backgroundColor: '#f8f9fa',
              border: '2px dashed #dee2e6',
              borderRadius: '8px',
              margin: '20px 0'
            }}>
              <div style={{ fontSize: '48px', marginBottom: '20px' }}>🔒</div>
            </div>
          ) : (
            <>
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', margin: '10px 0' }}>
              <thead style={{ backgroundColor: '#0077B6', color: 'white' }}>
                <tr>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '60px' }}>#</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '150px' }}>Patient's Name</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '100px' }}>Office</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '120px' }}>Appt. Date</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '120px' }}>Type of Visit</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '80px' }}>Call In</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '80px' }}>Call Out</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '80px' }}>Time</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '150px' }}>Remark</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '150px' }}>Other Duty</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '80px' }}>Action</th>
                </tr>
              </thead>
              <tbody>
                {patientRows.map((row) => (
                  <PatientRow
                    key={row.id}
                    row={row}
                    updatePatientRow={updatePatientRow}
                    removePatientRow={removePatientRow}
                    patientOfficeOptions={patientOfficeOptions}
                    getVisitTypeOptions={getVisitTypeOptions}
                    remarkOptions={remarkOptions}
                    otherDutyOptions={otherDutyOptions}
                    inputStyle={inputStyle}
                    buttonStyle={buttonStyle}
                  />
                ))}
              </tbody>
            </table>
            </div>

            <div style={{ display: 'flex', justifyContent: 'center', marginTop: '20px' }}>
              <button onClick={addPatientRow} style={buttonStyle}>
                + Add
              </button>
            </div>
            </>
          )}
        </div>

        {/* 일일 업무 보고서 */}
        <div style={sectionStyle}>
          <h2 style={{ color: '#0077B6', marginBottom: '15px' }}>Daily Work Report</h2>
          {!isUnlocked ? (
            <div style={{
              padding: '20px',
              textAlign: 'center',
              backgroundColor: '#f8f9fa',
              border: '2px dashed #dee2e6',
              borderRadius: '8px',
              margin: '10px 0'
            }}>
              <div style={{ fontSize: '24px', marginBottom: '10px' }}>🔒</div>
            </div>
          ) : (
            <textarea
              value={formData.dailyWorkReport}
              onChange={(e) => updateFormData('dailyWorkReport', e.target.value)}
              rows={4}
              placeholder="Enter your daily work report here..."
              style={{
                ...inputStyle,
                minHeight: '100px',
                resize: 'vertical'
              }}
            />
          )}
        </div>

        {/* PDF 생성 버튼 */}
        <div style={{ display: 'flex', justifyContent: 'center', gap: '20px', marginBottom: '30px' }}>
          {!isUnlocked ? (
            <div style={{
              padding: '20px',
              textAlign: 'center',
              backgroundColor: '#f8f9fa',
              border: '2px dashed #dee2e6',
              borderRadius: '8px',
              margin: '10px 0'
            }}>
              <div style={{ fontSize: '24px', marginBottom: '10px' }}>🔒</div>
            </div>
          ) : (
            <button 
              onClick={handleSubmit} 
              disabled={loading} 
              style={{ ...buttonStyle, backgroundColor: '#28a745' }}
            >
              {loading ? 'Submitting...' : 'Submit'}
            </button>
          )}
        </div>

      </div>
    </div>
    </>
  );
}

export default function PatientLogPage() {
  return <PatientLogSystem />;
}


