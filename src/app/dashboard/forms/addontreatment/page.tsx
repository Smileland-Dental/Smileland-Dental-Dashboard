'use client'

import React, { useState, useEffect, useCallback } from "react";
import { doc, setDoc, getDoc, deleteDoc, onSnapshot, collection, getDocs } from "firebase/firestore";
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

// 🔒 보안: 문서 ID sanitization
const sanitizeDocId = (docId: string): string => {
  if (!docId || typeof docId !== 'string') return '';
  // 경로 탐색 공격 방지
  let sanitized = docId.replace(/\.\./g, '_');
  // 허용된 문자만 유지
  sanitized = sanitized.replace(/[^a-zA-Z0-9_-]/g, '_');
  // 길이 제한 (Firebase 제한: 1500 bytes)
  if (sanitized.length > 1500) {
    sanitized = sanitized.substring(0, 1500);
  }
  return sanitized;
};

// 🔒 보안: Firebase 데이터 sanitization
const sanitizeFirebaseData = (data: any): any => {
  if (data === null || data === undefined) return data;
  
  if (typeof data === 'string') {
    return validateInput(data, 1000);
  }
  
  if (Array.isArray(data)) {
    return data.map(item => sanitizeFirebaseData(item));
  }
  
  if (typeof data === 'object') {
    const sanitized: any = {};
    for (const key in data) {
      if (data.hasOwnProperty(key)) {
        // 키도 sanitize
        const safeKey = validateInput(key, 200);
        sanitized[safeKey] = sanitizeFirebaseData(data[key]);
      }
    }
    return sanitized;
  }
  
  return data;
};

export default function AddOnTreatment() {

  const [loading, setLoading] = useState(false);
  const [autoSaveStatus, setAutoSaveStatus] = useState('');
  const [submitStatus, setSubmitStatus] = useState('');
  const [progress, setProgress] = useState(0);
  const [isUpdatingFromFirebase, setIsUpdatingFromFirebase] = useState(false);
  const [isAuthorized, setIsAuthorized] = useState<boolean | null>(null); // null: 확인 중, true: 인증됨, false: 인증 실패
  const [userOfficeBasedOptions, setUserOfficeBasedOptions] = useState<string[]>([]); // 사용자의 office_basedes 옵션들
  
  // 🔒 보안: 사용자 세션 ID 생성 (페이지 로드 시 한 번만)
  // 더 안전한 UUID 생성 방식을 사용하는 것을 권장합니다 (예: crypto.randomUUID)
  const [userSessionId] = useState(() => {
    // crypto.randomUUID가 사용 가능한 경우 사용, 아니면 fallback
    if (typeof window !== 'undefined' && typeof crypto !== 'undefined' && crypto.randomUUID) {
      return crypto.randomUUID();
    }
    // Fallback: 더 안전한 랜덤 문자열 생성
    const array = new Uint8Array(16);
    if (typeof window !== 'undefined' && typeof crypto !== 'undefined' && crypto.getRandomValues) {
      crypto.getRandomValues(array);
    } else {
      // 최후의 수단: Math.random 사용 (보안상 권장하지 않음)
      for (let i = 0; i < array.length; i++) {
        array[i] = Math.floor(Math.random() * 256);
      }
    }
    return Array.from(array, byte => byte.toString(16).padStart(2, '0')).join('');
  });
  
  // 마지막 저장된 데이터 추적
  const [lastSavedData, setLastSavedData] = useState({});
  
  // 현재 캘리포니아 시간 가져오기
  const getCurrentCaliforniaTime = () => {
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
  };
  
  // 날짜 상태
  const [dutyDate, setDutyDate] = useState('');

  // 오피스 선택 상태
  const [selectedOffice, setSelectedOffice] = useState('');
  
  // 오피스 옵션
  const officeOptions = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

  // 환자 데이터 상태 - 동적으로 관리
  const [patientData, setPatientData] = useState<Record<string, string>>({});
  const [rowCount, setRowCount] = useState(20);

  // 12시간 형식 시간 가져오기
  const getCurrentTime12Hour = () => {
    const now = new Date();
    const californiaTime = new Date(now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
    let hours = californiaTime.getHours();
    const minutes = californiaTime.getMinutes();
    const ampm = hours >= 12 ? 'PM' : 'AM';
    
    hours = hours % 12;
    hours = hours ? hours : 12;
    
    return `${hours}:${minutes.toString().padStart(2, '0')} ${ampm}`;
  };

  // 시간을 12시간 형식으로 변환하는 함수
  const convertTo12Hour = (timeStr: string): string => {
    if (!timeStr || timeStr === '-') return '-';
    
    // 이미 AM/PM이 포함된 경우 그대로 반환 (중복 변환 방지)
    if (timeStr.includes('AM') || timeStr.includes('PM')) {
      return timeStr;
    }
    
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
  };

  // --- PDF 생성 관련 상수/스타일 ---
  const pdfStyles = StyleSheet.create({
    page: { padding: 22, fontFamily: 'Helvetica', fontSize: 8 },
    header: { marginBottom: 10, borderBottomWidth: 2, borderColor: '#333', paddingBottom: 6, alignItems: 'center' },
    headerTitle: { fontSize: 14, fontWeight: 'bold', marginBottom: 4 },
    infoSection: { flexDirection: 'row', justifyContent: 'space-between', marginBottom: 15, fontSize: 10 },
    row: { flexDirection: 'row', borderBottomWidth: 0.5, borderColor: '#333' },
    cell: { padding: 4, fontSize: 8, flex: 1, borderRightWidth: 0.5, borderColor: '#333', justifyContent: 'center', alignItems: 'center' },
    cellNo: { flex: 0.3 },
    cellName: { flex: 2 },
    cellDob: { flex: 1.5 },
    cellTime: { flex: 1 },
    cellBold: { fontWeight: 'bold' },
    cellGray: { backgroundColor: '#f0f0f0' },
    footer: { marginTop: 15, paddingTop: 8, borderTopWidth: 1, borderColor: '#ddd', alignItems: 'center', fontSize: 7, color: '#666' },
  });

  // PDF 생성 유틸 함수
  function safeStr(v: unknown, max: number): string {
    if (v == null) return '';
    return String(v).trim().slice(0, max).replace(/[<>]/g, '');
  }

  function createAddOnTreatmentPDFDocument(props: {
    safeDutyDate: string;
    safeSelectedOffice: string;
    patientRows: Array<{
      rowNumber: number;
      patientName: string;
      dob: string;
      time: string;
    }>;
    generatedDate: string;
  }) {
    const { safeDutyDate, safeSelectedOffice, patientRows, generatedDate } = props;
    const s = pdfStyles;

    // Header
    const header = React.createElement(View, { style: s.header },
      React.createElement(Text, { style: s.headerTitle }, 'Add-On Treatment'),
    );

    // Info Section
    const infoSection = React.createElement(View, { style: s.infoSection },
      React.createElement(Text, null, `Date: ${safeDutyDate}`),
      React.createElement(Text, null, `Location: ${safeSelectedOffice}`),
    );

    // Table Header
    const tableHeader = React.createElement(View, { key: 'header', style: [s.row, s.cellGray] },
      React.createElement(View, { style: [s.cell, s.cellNo, s.cellBold] }, React.createElement(Text, null, 'No.')),
      React.createElement(View, { style: [s.cell, s.cellName, s.cellBold] }, React.createElement(Text, null, 'Name of Patient')),
      React.createElement(View, { style: [s.cell, s.cellDob, s.cellBold] }, React.createElement(Text, null, 'DOB')),
      React.createElement(View, { style: [s.cell, s.cellTime, s.cellBold] }, React.createElement(Text, null, 'Time')),
    );

    // Data Rows
    const dataRows = patientRows.map((row, index) => {
      const safePatientName = safeStr(row.patientName, 100);
      const safeDob = safeStr(row.dob, 20);
      const safeTime = safeStr(convertTo12Hour(row.time), 20);
      
      return React.createElement(View, { key: index, style: s.row },
        React.createElement(View, { style: [s.cell, s.cellNo] }, React.createElement(Text, null, String(index + 1))),
        React.createElement(View, { style: [s.cell, s.cellName] }, React.createElement(Text, null, safePatientName || '-')),
        React.createElement(View, { style: [s.cell, s.cellDob] }, React.createElement(Text, null, safeDob || '-')),
        React.createElement(View, { style: [s.cell, s.cellTime] }, React.createElement(Text, null, safeTime || '-')),
      );
    });

    const table = React.createElement(View, null, tableHeader, ...dataRows);

    const footer = React.createElement(View, { style: s.footer }, React.createElement(Text, null, `Generated: ${generatedDate}`));

    return React.createElement(Document, null,
      React.createElement(Page, { size: 'A4', orientation: 'portrait', style: s.page }, header, infoSection, table, footer),
    );
  }

  function sanitizeFilename(filename: string): string {
    return filename.replace(/[^a-zA-Z0-9._-]/g, '_').replace(/\.\./g, '_').slice(0, 255);
  }

  // 행 추가 함수
  const addRow = () => {
    setRowCount(prev => prev + 1);
  };

  // 행 클리어 함수
  const clearRow = (rowNumber: number) => {
    setPatientData(prev => {
      const newData = { ...prev };
      newData[`Row${rowNumber}_PatientName`] = '';
      newData[`Row${rowNumber}_DOB`] = '';
      newData[`Row${rowNumber}_Time`] = '';
      return newData;
    });
  };

  // 자동 저장 함수
  const autoSave = useCallback(async () => {
    if (!dutyDate || !selectedOffice || isUpdatingFromFirebase) return;

    // 데이터가 실제로 변경되었는지 확인
    const currentData = { ...patientData, rowCount, selectedOffice };
    const hasChanges = JSON.stringify(currentData) !== JSON.stringify(lastSavedData);
    if (!hasChanges) return;

    try {
      const dataToSave = {
        dutyDate,
        selectedOffice,
        ...patientData,
        rowCount,
        timestamp: new Date().toISOString(),
        autoSaved: true,
        lastUpdatedBy: userSessionId
      };

      // 🔒 보안: 문서 ID 검증 및 추가 sanitization
      const safeOffice = selectedOffice.replace(/[^a-zA-Z0-9_-]/g, '');
      const safeDutyDate = dutyDate.replace(/[^a-zA-Z0-9_-]/g, '');
      // 경로 탐색 공격 방지
      if (safeOffice.includes('..') || safeDutyDate.includes('..')) {
        return;
      }
      const docId = sanitizeDocId(`${safeDutyDate}_${safeOffice}_addon_treatment`);
      // 문서 ID 길이 제한 (Firebase 제한: 1500 bytes)
      if (docId.length > 1500) {
        return;
      }
      const safeDataToSave = sanitizeFirebaseData(dataToSave);
      await setDoc(doc(db, "addon-treatment", docId), safeDataToSave);
      
      // 저장 성공 시 마지막 저장된 데이터 업데이트
      setLastSavedData(currentData);
      
    } catch (error) {
      // Auto-save error silently handled
    }
  }, [dutyDate, selectedOffice, patientData, rowCount, lastSavedData, isUpdatingFromFirebase, userSessionId]);

  // 데이터 변경 시에만 자동 저장
  useEffect(() => {
    // 초기 로드 시에는 저장하지 않음
    if (Object.values(patientData).some(value => value !== '')) {
      autoSave();
    }
  }, [patientData]);

  // 데이터 로드
  const loadData = async () => {
    if (!dutyDate || !selectedOffice) return;

    try {
      setSubmitStatus('Loading data...');
      
      // 🔒 보안: 문서 ID 검증
      const safeOffice = selectedOffice.replace(/[^a-zA-Z0-9_-]/g, '');
      const safeDutyDate = dutyDate.replace(/[^a-zA-Z0-9_-]/g, '');
      if (safeOffice.includes('..') || safeDutyDate.includes('..')) {
        return;
      }
      const docId = sanitizeDocId(`${safeDutyDate}_${safeOffice}_addon_treatment`);
      if (docId.length > 1500) {
        return;
      }
      const docSnap = await getDocs(collection(db, "addon-treatment")).then(snapshot => {
        const foundDoc = snapshot.docs.find(d => d.id === docId);
        return foundDoc ? { 
          exists: (): boolean => true, 
          data: (): any => foundDoc.data() 
        } : { 
          exists: (): boolean => false,
          data: (): any => undefined
        };
      });
      
      if (docSnap.exists()) {
        const data = docSnap.data();
        if (!data) {
          setSubmitStatus('No data found - initialized empty form');
          setTimeout(() => setSubmitStatus(''), 2000);
          return;
        }
        
        // Firebase에서 업데이트되는 동안 자동 저장 방지
        setIsUpdatingFromFirebase(true);
        
        setPatientData(prevData => ({
          ...prevData,
          ...data
        }));
        
        // rowCount 복원
        if (data.rowCount) {
          setRowCount(data.rowCount);
        }
        
        // 로드된 데이터를 마지막 저장된 데이터로 설정
        setLastSavedData({ ...data, rowCount: data.rowCount || 20 });
        
        // 짧은 지연 후 Firebase 업데이트 플래그 해제
        setTimeout(() => {
          setIsUpdatingFromFirebase(false);
        }, 100);
        
        setSubmitStatus('Data loaded successfully');
        setTimeout(() => setSubmitStatus(''), 2000);
      } else {
        // 데이터가 없으면 초기화
        setPatientData({});
        setRowCount(20);
        setLastSavedData({ rowCount: 20 });
        setSubmitStatus('No data found - initialized empty form');
        setTimeout(() => setSubmitStatus(''), 2000);
      }
      
    } catch (error) {
      // 🔒 보안: 에러 메시지에 민감한 정보 노출 방지
      setSubmitStatus('Error loading data. Please try again.');
      setTimeout(() => setSubmitStatus(''), 3000);
    }
  };

  // 날짜 또는 오피스 변경 시 데이터 로드
  useEffect(() => {
    loadData();
  }, [dutyDate, selectedOffice]);

  // 실시간 데이터 동기화
  useEffect(() => {
    if (!dutyDate || !selectedOffice) return;

    // 🔒 보안: 문서 ID 검증
    const safeOffice = selectedOffice.replace(/[^a-zA-Z0-9_-]/g, '');
    const safeDutyDate = dutyDate.replace(/[^a-zA-Z0-9_-]/g, '');
    if (safeOffice.includes('..') || safeDutyDate.includes('..')) {
      return;
    }
    const docId = sanitizeDocId(`${safeDutyDate}_${safeOffice}_addon_treatment`);
    if (docId.length > 1500) {
      return;
    }
    const docRef = doc(db, "addon-treatment", docId);
    
    const unsubscribe = onSnapshot(docRef, (docSnap) => {
      if (docSnap.exists()) {
        const data = docSnap.data();
        
        // Firebase에서 업데이트되는 동안 자동 저장 방지
        setIsUpdatingFromFirebase(true);
        
        setPatientData(prevData => {
          return {
            ...prevData,
            ...data
          };
        });
        
        // rowCount 복원
        if (data.rowCount) {
          setRowCount(data.rowCount);
        }
        
        // 실시간 업데이트된 데이터를 마지막 저장된 데이터로 설정
        setLastSavedData({ ...data, rowCount: data.rowCount || 20 });
        
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
      setAutoSaveStatus('❌ Connection error');
      setTimeout(() => setAutoSaveStatus(''), 3000);
    });

    return () => {
      unsubscribe();
    };
  }, [dutyDate, selectedOffice]);

  // 데이터 업데이트 함수
  const updatePatientData = (field: string, value: string) => {
    setPatientData(prev => {
      const newData = { ...prev, [field]: value };
      
      // Patient Name 필드가 입력되면 시간 자동 기록
      if (field.endsWith('_PatientName') && value.trim() !== '' && (prev[field] || '').trim() === '') {
        const rowNumber = field.match(/Row(\d+)_PatientName/)?.[1];
        if (rowNumber) {
          const timeField = `Row${rowNumber}_Time`;
          if (!newData[timeField]) {
            newData[timeField] = getCurrentTime12Hour();
          }
        }
      }
      
      // Patient Name 필드가 비워지면 시간도 비움
      if (field.endsWith('_PatientName') && value.trim() === '') {
        const rowNumber = field.match(/Row(\d+)_PatientName/)?.[1];
        if (rowNumber) {
          const timeField = `Row${rowNumber}_Time`;
          newData[timeField] = '';
        }
      }
      
      return newData;
    });
  };

  // Office 변경 처리
  const handleOfficeChange = (newOffice: string) => {
    // 빈 값으로 선택하면 비밀번호 없이 변경 허용 (초기화)
    if (newOffice === '') {
      setSelectedOffice('');
      return;
    }
    
    // 선택된 office의 첫 알파벳 대문자를 비밀번호로 사용
    const officePassword = newOffice.charAt(0).toUpperCase();
    const password = prompt(`Enter password to change office: `);
    if (password === null) return;
    if (password !== officePassword) {
      alert("Incorrect password. Office change cancelled.");
      return;
    }
    setSelectedOffice(newOffice);
  };

  // 제출 처리
  const handleSubmit = async () => {
    if (!selectedOffice) {
      alert("Please select an office before submitting.");
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

      // 입력 검증
      if (!dutyDate || !selectedOffice) {
        throw new Error('Please fill out all required fields.');
      }

      if (typeof dutyDate !== 'string' || typeof selectedOffice !== 'string') {
        throw new Error('Invalid input format');
      }

      // 날짜 형식 검증 (YYYY-MM-DD)
      if (!/^\d{4}-\d{2}-\d{2}$/.test(dutyDate)) {
        throw new Error('Invalid date format');
      }

      // 날짜 유효성 검증
      const dateObj = new Date(dutyDate + 'T00:00:00');
      if (isNaN(dateObj.getTime())) {
        throw new Error('Invalid date format');
      }

      const [dateYear, dateMonth, dateDay] = dutyDate.split('-').map(Number);
      if (dateYear < 1900 || dateYear > 2100 || dateMonth < 1 || dateMonth > 12 || dateDay < 1 || dateDay > 31) {
        throw new Error('Invalid date format');
      }

      const reconstructedDate = `${dateYear}-${String(dateMonth).padStart(2, '0')}-${String(dateDay).padStart(2, '0')}`;
      if (reconstructedDate !== dutyDate) {
        throw new Error('Invalid date format');
      }

      // 오피스 선택 검증
      if (!['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'].includes(selectedOffice)) {
        throw new Error('Invalid office selection');
      }

      // 1. PDF 생성
      setSubmitStatus('Submitting...');
      setProgress(30);
      
      // 환자 데이터를 배열 형태로 변환
      const patientRows: Array<{
        rowNumber: number;
        patientName: string;
        dob: string;
        time: string;
      }> = [];
      for (let i = 1; i <= rowCount; i++) {
        const patientName = patientData[`Row${i}_PatientName`];
        const dob = patientData[`Row${i}_DOB`];
        const time = patientData[`Row${i}_Time`];
        
        if (patientName || dob || time) {
          patientRows.push({
            rowNumber: i,
            patientName: patientName || '',
            dob: dob || '',
            time: time || ''
          });
        }
      }

      // 환자 행 데이터 검증
      if (patientRows.length > 1000) {
        throw new Error('Too many patient rows');
      }

      for (const row of patientRows) {
        if (typeof row.rowNumber !== 'number' || row.rowNumber < 1 || row.rowNumber > 10000) {
          throw new Error('Invalid patient data format');
        }
        if (typeof row.patientName !== 'string' || row.patientName.length > 100) {
          throw new Error('Invalid patient data format');
        }
        if (typeof row.dob !== 'string' || row.dob.length > 20) {
          throw new Error('Invalid patient data format');
        }
        if (typeof row.time !== 'string' || row.time.length > 20) {
          throw new Error('Invalid patient data format');
        }
      }

      const safeDutyDate = dutyDate.trim().slice(0, 50).replace(/[<>]/g, '');
      const safeSelectedOffice = selectedOffice.trim().slice(0, 50).replace(/[<>]/g, '');

      const generatedDate = new Date().toLocaleDateString('en-US', {
        year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit', hour12: true,
      });

      // PDF 생성 (클라이언트 사이드)
      setSubmitStatus('Processing PDF...');
      setProgress(60);
      
      const pdfDoc = createAddOnTreatmentPDFDocument({
        safeDutyDate,
        safeSelectedOffice,
        patientRows,
        generatedDate,
      });

      const blob = await pdf(pdfDoc).toBlob();
        
      // PDF를 Firebase Storage에 저장 (endofday-pdfs에만 저장)
      setSubmitStatus('Saving...');
      setProgress(70);
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
        
        const filename = `4) ${safeDutyDate}_${safeSelectedOffice}_Add On Treatment_${timeStamp}.pdf`;
        const storageRef = ref(storage, `endofday-pdfs/${safeSelectedOffice}/${safeDutyDate}/${filename}`);
        
        // PDF 업로드
        await uploadBytes(storageRef, blob);
        
      } catch (storageError: any) {
        alert('An error occurred while submitting. Please try again.');
        throw storageError;
      }
      
      // 2. 데이터 삭제
      setSubmitStatus('Cleaning up...');
      setProgress(80);
      const docId = sanitizeDocId(`${safeDutyDate}_${safeSelectedOffice}_addon_treatment`);
      await deleteDoc(doc(db, "addon-treatment", docId));
      
      // 3. 폼 초기화
      setPatientData({});
      setRowCount(20);

      setSubmitStatus('Complete!');
      setProgress(100);
      
      // 2초 후 모달 닫기
      setTimeout(() => {
        setLoading(false);
        setSubmitStatus('');
        setProgress(0);
      }, 2000);

    } catch (error: any) {
      setSubmitStatus('❌ Submission failed. Please try again.');
      setProgress(0);
      setTimeout(() => {
        setLoading(false);
        setSubmitStatus('');
      }, 3000);
    }
  };

  // 스타일 정의
  const styles = {
    body: {
      fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
      padding: '20px',
      background: `
        radial-gradient(circle at 10% 20%, rgba(120, 200, 255, 0.1) 0%, transparent 50%),
        radial-gradient(circle at 90% 80%, rgba(255, 182, 193, 0.1) 0%, transparent 50%),
        radial-gradient(circle at 50% 50%, rgba(144, 238, 144, 0.05) 0%, transparent 50%),
        linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%)
      `,
      color: '#2c3e50',
      lineHeight: '1.6',
      minHeight: '100vh',
      position: 'relative' as const
    },
    container: {
      maxWidth: '1200px',
      margin: '40px auto',
      padding: '30px',
      backgroundColor: 'white',
      borderRadius: '12px',
      boxShadow: '0 4px 20px rgba(0, 0, 0, 0.1)',
      border: '1px solid #e9ecef',
      position: 'relative' as const,
      overflow: 'hidden'
    },
    header: {
      color: '#2c3e50',
      textAlign: 'center' as const,
      marginBottom: '30px',
      paddingBottom: '20px',
      borderBottom: '3px solid transparent',
      borderImage: 'linear-gradient(90deg, #4CAF50, #2196F3, #FF9800) 1',
      fontSize: '2.2em',
      fontWeight: '600',
      letterSpacing: '-0.5px',
      position: 'relative' as const
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
      fontSize: '16px',
      border: '1px solid #e9ecef',
      borderRadius: '4px',
      backgroundColor: 'white',
      boxSizing: 'border-box' as const,
      transition: 'border-color 0.2s'
    },
    table: {
      borderCollapse: 'collapse' as const,
      width: '100%',
      marginTop: '20px',
      backgroundColor: 'white',
      borderRadius: '6px',
      boxShadow: '0 1px 3px rgba(0, 0, 0, 0.1)',
      border: '1px solid #dee2e6',
      overflow: 'hidden'
    },
    th: {
      border: '1px solid #dee2e6',
      padding: '12px 10px',
      textAlign: 'center' as const,
      verticalAlign: 'middle' as const,
      background: 'linear-gradient(135deg, #4CAF50 0%, #45a049 100%)',
      color: 'white',
      fontWeight: '600',
      fontSize: '14px',
      textShadow: '0 1px 2px rgba(0, 0, 0, 0.1)'
    },
    td: {
      border: '1px solid #dee2e6',
      padding: '10px',
      textAlign: 'center' as const,
      verticalAlign: 'middle' as const,
      backgroundColor: 'white',
      transition: 'background-color 0.2s ease'
    },
    submitButton: {
      display: 'block',
      width: '150px',
      margin: '30px auto 0 auto',
      padding: '12px 20px',
      background: 'linear-gradient(135deg, #2196F3 0%, #1976D2 100%)',
      color: 'white',
      border: 'none',
      borderRadius: '8px',
      cursor: 'pointer',
      fontSize: '16px',
      fontWeight: '600',
      transition: 'all 0.3s ease',
      boxShadow: '0 2px 8px rgba(33, 150, 243, 0.3)',
      textShadow: '0 1px 2px rgba(0, 0, 0, 0.1)'
    },
    autoSaveStatus: {
      position: 'absolute' as const,
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
    footer: {
      marginTop: '40px',
      paddingTop: '20px',
      borderTop: '1px solid #e9ecef',
      textAlign: 'center' as const,
      fontSize: '0.9em',
      color: '#6c757d'
    }
  };

  // 컴포넌트 마운트 시 사용자 인증 및 role 확인
  useEffect(() => {
    // Firebase Auth 상태 변경 감지
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      try {
        if (!currentUser) {
          alert('Please log in');
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

        if (userData?.role !== 'manager') {
          alert('You do not have access to this page.');
          setIsAuthorized(false);
          // 다른 페이지로 리다이렉트하거나 홈으로 이동
          if (typeof window !== 'undefined') {
            window.location.href = '/';
          }
          return;
        }

        setIsAuthorized(true);
        setDutyDate(getCurrentCaliforniaTime());

        // office_based 처리: 배열이거나 단일 값일 수 있음
        if (userData?.office_based) {
          const officeBasedArray = Array.isArray(userData.office_based) 
            ? userData.office_based 
            : [userData.office_based];
          
          // officeOptions에 포함된 값들만 필터링
          const validOptions = officeBasedArray.filter((g: string) => officeOptions.includes(g));
          
          if (validOptions.length > 0) {
            setUserOfficeBasedOptions(validOptions);
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
            position: "fixed" as const,
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
                {submitStatus === 'Saving...'}
                {submitStatus === 'Generating PDF...'}
                {submitStatus === 'Processing PDF...'}
                {submitStatus === 'Cleaning up...'}
                {submitStatus === 'Complete!'}
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

          <h2 style={styles.header}>
            Add-On Treatment
          </h2>

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
            {/* office_basedes 옵션이 있는 경우에만 Office 표시 */}
            {userOfficeBasedOptions.length > 0 && (
              <div style={styles.formGroup}>
                <label style={styles.label} htmlFor="selectedOffice">Office:</label>
                {userOfficeBasedOptions.length === 1 ? (
                  <span style={{
                    ...styles.input,
                    display: 'inline-flex',
                    alignItems: 'center',
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
                    <option value="">--Select Office--</option>
                    {userOfficeBasedOptions.map(office => (
                      <option key={office} value={office}>{office}</option>
                    ))}
                  </select>
                )}
              </div>
            )}
          </div>

          {/* 환자 로그 테이블 */}
          <table style={styles.table}>
            <thead>
              <tr>
                <th style={{ ...styles.th, width: '60px' }}>No.</th>
                <th style={styles.th}>Name of Patient</th>
                <th style={styles.th}>DOB</th>
                <th style={styles.th}>Time</th>
                <th style={{ ...styles.th, width: '80px' }}>Action</th>
              </tr>
            </thead>
            <tbody>
              {Array.from({ length: rowCount }, (_, index) => {
                const rowNumber = index + 1;
                return (
                  <tr 
                    key={rowNumber}
                    style={{
                      transition: 'background-color 0.2s ease'
                    }}
                    onMouseEnter={(e) => {
                      e.currentTarget.style.backgroundColor = '#f8f9fa';
                    }}
                    onMouseLeave={(e) => {
                      e.currentTarget.style.backgroundColor = 'white';
                    }}
                  >
                    <td style={styles.td}>{rowNumber}</td>
                    <td style={styles.td}>
                      <input
                        type="text"
                        value={patientData[`Row${rowNumber}_PatientName`] || ''}
                        onChange={(e) => updatePatientData(`Row${rowNumber}_PatientName`, e.target.value)}
                        style={{ ...styles.input, margin: 0, fontSize: '14px', padding: '8px' }}
                      />
                    </td>
                    <td style={styles.td}>
                      <input
                        type="text"
                        value={patientData[`Row${rowNumber}_DOB`] || ''}
                        onChange={(e) => updatePatientData(`Row${rowNumber}_DOB`, e.target.value)}
                        placeholder="mm/dd/yyyy"
                        style={{ ...styles.input, margin: 0, fontSize: '14px', padding: '8px' }}
                      />
                    </td>
                    <td style={styles.td}>
                      <input
                        type="text"
                        value={convertTo12Hour(patientData[`Row${rowNumber}_Time`] || '')}
                        onChange={(e) => updatePatientData(`Row${rowNumber}_Time`, e.target.value)}
                        style={{ ...styles.input, margin: 0, fontSize: '14px', padding: '8px' }}
                      />
                    </td>
                    <td style={styles.td}>
                      <button
                        type="button"
                        onClick={() => clearRow(rowNumber)}
                        style={{
                          background: 'linear-gradient(135deg, #FF9800 0%, #F57C00 100%)',
                          color: 'white',
                          border: 'none',
                          padding: '8px 16px',
                          borderRadius: '6px',
                          cursor: 'pointer',
                          fontSize: '12px',
                          fontWeight: '600',
                          transition: 'all 0.3s ease',
                          boxShadow: '0 2px 6px rgba(255, 152, 0, 0.3)',
                          textShadow: '0 1px 2px rgba(0, 0, 0, 0.1)'
                        }}
                      >
                        Clear
                      </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>

          {/* 행 추가 버튼 */}
          <div style={{ textAlign: 'center', marginTop: '20px', marginBottom: '10px' }}>
            <button
              type="button"
              onClick={addRow}
              style={{
                background: 'linear-gradient(135deg, #4CAF50 0%, #45a049 100%)',
                color: 'white',
                border: 'none',
                padding: '12px 24px',
                borderRadius: '8px',
                cursor: 'pointer',
                fontSize: '14px',
                fontWeight: '600',
                transition: 'all 0.3s ease',
                boxShadow: '0 2px 8px rgba(76, 175, 80, 0.3)',
                textShadow: '0 1px 2px rgba(0, 0, 0, 0.1)'
              }}
            >
              + Add Row
            </button>
          </div>

          {/* 제출 버튼 */}
          <div style={{ textAlign: 'center', marginTop: '10px' }}>
            <button
              type="button"
              onClick={handleSubmit}
              disabled={loading || !selectedOffice}
              style={{
                ...styles.submitButton,
                background: loading || !selectedOffice ? 'linear-gradient(135deg, #6c757d 0%, #5a6268 100%)' : 'linear-gradient(135deg, #2196F3 0%, #1976D2 100%)',
                cursor: loading || !selectedOffice ? 'not-allowed' : 'pointer'
              }}
              onMouseEnter={(e) => {
                if (!loading) {
                  e.currentTarget.style.transform = 'translateY(-1px)';
                  e.currentTarget.style.boxShadow = '0 4px 12px rgba(33, 150, 243, 0.4)';
                }
              }}
              onMouseLeave={(e) => {
                if (!loading) {
                  e.currentTarget.style.transform = 'translateY(0)';
                  e.currentTarget.style.boxShadow = '0 2px 8px rgba(33, 150, 243, 0.3)';
                }
              }}
            >
              {loading ? 'Submitting...' : 'Submit'}
            </button>
          </div>

          {/* 상태 메시지 */}
          {submitStatus && (
            <div style={{
              marginTop: '15px',
              fontWeight: 'bold',
              textAlign: 'center',
              padding: '10px',
              borderRadius: '4px',
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
