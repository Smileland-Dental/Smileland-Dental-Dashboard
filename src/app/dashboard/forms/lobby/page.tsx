'use client';

import React, { useState, useEffect, useCallback, useRef } from 'react';
import { db, auth } from '@/lib/firebase.config';
import { doc, setDoc, getDoc, onSnapshot, deleteDoc } from 'firebase/firestore';
import { getStorage, ref, uploadBytes } from 'firebase/storage';
import { onAuthStateChanged } from 'firebase/auth';
import { pdf, Document, Page, View, Text, StyleSheet } from '@react-pdf/renderer';

export default function LobbyInspectionPage() {
  // 상태 관리
  const [inspectionDate, setInspectionDate] = useState('');
  const [selectedOffice, setSelectedOffice] = useState('');
  const [lobbyData, setLobbyData] = useState<any>({});
  const [loading, setLoading] = useState(false);
  const [submitStatus, setSubmitStatus] = useState('');
  const [progress, setProgress] = useState(0);
  const [isUpdatingFromFirebase, setIsUpdatingFromFirebase] = useState(false);
  const [userSessionId] = useState(() => Math.random().toString(36).substr(2, 9));
  const [lastSavedData, setLastSavedData] = useState<any>({});
  const [isAuthorized, setIsAuthorized] = useState<boolean | null>(null); // null: 확인 중, true: 인증됨, false: 인증 실패
  const [userOfficeBasedOptions, setUserOfficeBasedOptions] = useState<string[]>([]); // 사용자의 office_based 옵션들

  // Rate limiting을 위한 ref
  const lastUpdateLobbyDataCall = useRef<number>(0);
  const lastLoadDataCall = useRef<number>(0);
  const lastSubmitCall = useRef<number>(0);

  // 오피스 옵션 (알파벳 순)
  const officeOptions = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

  // Office별 Check 옵션 정의 (알파벳 순)
  const OFFICE_CHECK_OPTIONS = {
    "Bernard": ["Carmen", "Cynthia", "Elisa", "Ranjit"],
    "California": ["Kindal"],
    "Delano": ["Helen", "Stephanie"],
    "Fresno": ["Cynthia"],
    "Ming": ["Kindal", "Hopie"],
    "Ortho": ["Kindal"],
    "Tulare": ["Dianne"],
    "Visalia": ["Dianne"]
  };

  // --- PDF 생성 관련 상수/스타일 ---
  const PDF_COLUMN_NAMES = [
    'Time', 'Check', 'Ipads/Games Working', 'Wipe Ipads/Games', 'Pick Up Litter/Sweep',
    'Entrance Area', 'Pass Out Water', 'Sweep/Vacuum', 'Wipe Ipads/Games',
    'Take Out Trash', 'Wipe Desk Tops', 'Wipe Seats', 'Wipe Windows/Door Handles', 'Checked Time',
  ];

  const PDF_ROW_HEADERS = [
    'Manager Inspection',
    '8 am', '9 am', '10 am', '11 am',
    'Manager Inspection',
    '12 pm', '1 pm', 'Sweep/Mop', '2 pm', '3 pm',
    'Manager Inspection',
    '4 pm', '5 pm', '6 pm', 'Sweep/Mop', '7 pm',
    'Deep Clean Manager Inspection',
  ];

  const pdfStyles = StyleSheet.create({
    page: { padding: 22, fontFamily: 'Helvetica', fontSize: 8 },
    header: { marginBottom: 10, borderBottomWidth: 2, borderColor: '#333', paddingBottom: 6, alignItems: 'center' },
    headerTitle: { fontSize: 14, fontWeight: 'bold', marginBottom: 4 },
    headerSub: { fontSize: 10 },
    row: { flexDirection: 'row', borderBottomWidth: 0.5, borderColor: '#333' },
    cell: { padding: 3, fontSize: 5, flex: 1, borderRightWidth: 0.5, borderColor: '#333', justifyContent: 'center', alignItems: 'center' },
    cellFlex5: { flex: 5 },
    cellFlex6: { flex: 6 },
    cellBold: { fontWeight: 'bold' },
    cellGray: { backgroundColor: '#f0f0f0' },
    footer: { marginTop: 15, paddingTop: 8, borderTopWidth: 1, borderColor: '#ddd', alignItems: 'center', fontSize: 7, color: '#666' },
  });

  // PDF 생성 유틸 함수
  function safeStr(v: unknown, max: number): string {
    if (v == null) return '';
    return String(v).trim().slice(0, max).replace(/[<>]/g, '');
  }

  function isChecked(v: unknown): boolean {
    return v === true || v === 1 || (typeof v === 'string' && (v === 'true' || v === '1'));
  }

  function createLobbyPDFDocument(props: {
    safeSelectedOffice: string;
    formattedDate: string;
    lobbyData: Record<string, unknown>;
    generatedDate: string;
  }) {
    const { safeSelectedOffice, formattedDate, lobbyData, generatedDate } = props;
    const s = pdfStyles;

    // Row 1: Hourly Cleaning / End of Day Cleaning
    const row1 = React.createElement(View, { key: 'r1', style: [s.row, s.cellGray] },
      React.createElement(View, { style: [s.cell] }, React.createElement(Text, null, '')),
      React.createElement(View, { style: [s.cell] }, React.createElement(Text, null, '')),
      React.createElement(View, { style: [s.cell, s.cellFlex5, s.cellBold] }, React.createElement(Text, null, 'Hourly Cleaning')),
      React.createElement(View, { style: [s.cell, s.cellFlex6, s.cellBold] }, React.createElement(Text, null, 'End of Day Cleaning')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, null, '')),
    );

    // Row 2: Column names (14 columns)
    const row2 = React.createElement(View, { key: 'r2', style: [s.row, s.cellGray] },
      ...PDF_COLUMN_NAMES.map((name, i) =>
        React.createElement(View, { key: i, style: s.cell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, name))
      ),
    );

    // Data rows
    const dataRows = PDF_ROW_HEADERS.map((header, rowIndex) => {
      const isSweepMop = header.includes('Sweep/Mop');
      const checkValue = safeStr(lobbyData[`Row${rowIndex + 1}_Check`], 50);
      const checkedTime = safeStr(lobbyData[`Row${rowIndex + 1}_CheckedTime`], 50);
      const taskCells = PDF_COLUMN_NAMES.slice(2, -1).map((_, colIndex) => {
        if (isSweepMop) return React.createElement(View, { key: colIndex, style: s.cell }, React.createElement(Text, null, ''));
        const v = lobbyData[`Row${rowIndex + 1}_Col${colIndex + 3}`];
        return React.createElement(View, { key: colIndex, style: s.cell }, React.createElement(Text, null, isChecked(v) ? 'O' : ''));
      });
      return React.createElement(View, { key: rowIndex, style: s.row },
        React.createElement(View, { style: s.cell }, React.createElement(Text, null, header)),
        React.createElement(View, { style: s.cell }, React.createElement(Text, null, checkValue)),
        ...taskCells,
        React.createElement(View, { style: s.cell }, React.createElement(Text, null, checkedTime)),
      );
    });

    const table = React.createElement(View, null, row1, row2, ...dataRows);

    const header = React.createElement(View, { style: s.header },
      React.createElement(Text, { style: s.headerTitle }, 'Lobby Inspection Log'),
      React.createElement(Text, { style: s.headerSub }, `${safeSelectedOffice} (${formattedDate})`),
    );

    const footer = React.createElement(View, { style: s.footer }, React.createElement(Text, null, `Generated: ${generatedDate}`));

    return React.createElement(Document, null,
      React.createElement(Page, { size: 'A4', orientation: 'landscape', style: s.page }, header, table, footer),
    );
  }

  function sanitizeFilename(filename: string): string {
    return filename.replace(/[^a-zA-Z0-9._-]/g, '_').replace(/\.\./g, '_').slice(0, 255);
  }

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

  // 12시간제 시간 포맷
  const getCurrentTime12Hour = () => {
    const now = new Date();
    const laTime = new Date(now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
    let hours = laTime.getHours();
    const minutes = laTime.getMinutes();
    const ampm = hours >= 12 ? 'pm' : 'am';
    hours = hours % 12;
    hours = hours ? hours : 12;
    
    return `${hours}:${minutes.toString().padStart(2, '0')}${ampm}`;
  };

  // 자동 저장 함수
  const autoSave = useCallback(async () => {
    if (!inspectionDate || !selectedOffice || isUpdatingFromFirebase) return;

    // 데이터가 실제로 변경되었는지 확인
    const hasChanges = JSON.stringify(lobbyData) !== JSON.stringify(lastSavedData);
    if (!hasChanges) return;

    try {
      const dataToSave = {
        inspectionDate,
        selectedOffice,
        ...lobbyData,
        timestamp: new Date().toISOString(),
        autoSaved: true,
        lastUpdatedBy: userSessionId
      };

      const docId = `${inspectionDate}_${selectedOffice}_lobby`;
      await setDoc(doc(db, "lobby-inspections", docId), dataToSave);
      
      // 저장 성공 시 마지막 저장된 데이터 업데이트
      setLastSavedData({ ...lobbyData });
      
    } catch (error) {
      // Auto-save error silently handled
    }
  }, [inspectionDate, selectedOffice, lobbyData, lastSavedData, isUpdatingFromFirebase]);

  // 데이터 변경 시에만 자동 저장
  useEffect(() => {
    // 초기 로드 시에는 저장하지 않음
    if (Object.values(lobbyData).some(value => value !== '')) {
      autoSave();
    }
  }, [lobbyData]);

  // 데이터 로드 함수 (Rate limiting 적용)
  const loadData = async () => {
    if (!inspectionDate || !selectedOffice) return;

    try {
      // Rate limiting: 최근 1.5초 내 호출 방지
      // (필터 변경 후 빠르게 새로고침할 수 있도록 허용하되, 과도한 호출 방지)
      const now = Date.now();
      if (now - lastLoadDataCall.current < 1500) {
        return;
      }
      lastLoadDataCall.current = now;

      setSubmitStatus('Loading data...');
      
      const docId = `${inspectionDate}_${selectedOffice}_lobby`;
      const docSnap = await getDoc(doc(db, "lobby-inspections", docId));
      
      if (docSnap.exists()) {
        const data = docSnap.data();
        
        // Firebase에서 업데이트되는 동안 자동 저장 방지
        setIsUpdatingFromFirebase(true);
        
        setLobbyData((prevData: any) => ({
          ...prevData,
          ...data
        }));
        
        // 로드된 데이터를 마지막 저장된 데이터로 설정
        setLastSavedData({ ...data });
        
        // 짧은 지연 후 Firebase 업데이트 플래그 해제
        setTimeout(() => {
          setIsUpdatingFromFirebase(false);
        }, 100);
        
        setSubmitStatus('Data loaded successfully');
        setTimeout(() => setSubmitStatus(''), 2000);
      } else {
        // 초기 데이터 설정
        const initialData: any = {};
        setLobbyData(initialData);
        setLastSavedData(initialData);
        setSubmitStatus('No data found - initialized empty form');
        setTimeout(() => setSubmitStatus(''), 2000);
      }
      
    } catch (error: any) {
      setSubmitStatus('Error loading data: ' + error.message);
      setTimeout(() => setSubmitStatus(''), 3000);
    }
  };

  // 날짜, 오피스 변경 시 데이터 로드
  useEffect(() => {
    if (inspectionDate && selectedOffice) {
      loadData();
    }
  }, [inspectionDate, selectedOffice]);

  // 실시간 데이터 동기화
  useEffect(() => {
    if (!inspectionDate || !selectedOffice) return;

    const docId = `${inspectionDate}_${selectedOffice}_lobby`;
    const docRef = doc(db, "lobby-inspections", docId);
    
    const unsubscribe = onSnapshot(docRef, (docSnap: any) => {
      if (docSnap.exists()) {
        const data = docSnap.data();
        
        // 데이터가 실제로 변경되었는지 확인
        const hasChanges = JSON.stringify(data) !== JSON.stringify(lobbyData);
        if (!hasChanges) {
          return;
        }
        
        // Firebase에서 업데이트되는 동안 자동 저장 방지
        setIsUpdatingFromFirebase(true);
        
        setLobbyData((prevData: any) => {
          return {
            ...prevData,
            ...data
          };
        });
        
        // 실시간 업데이트된 데이터를 마지막 저장된 데이터로 설정
        setLastSavedData({ ...data });
        
        // 짧은 지연 후 Firebase 업데이트 플래그 해제
        setTimeout(() => {
          setIsUpdatingFromFirebase(false);
        }, 100);
        
        // 다른 사용자의 업데이트는 조용히 처리 (알림 없음)
      }
    }, (error: any) => {
      // Real-time listener error silently handled
    });

    return () => {
      unsubscribe();
    };
  }, [inspectionDate, selectedOffice]);

  // 데이터 업데이트 함수 (Rate limiting 적용)
  const updateLobbyData = (field: string, value: any) => {
    // Rate limiting: 최근 300ms 내 동일한 필드에 대한 호출 방지
    // (여러 필드를 빠르게 업데이트할 수 있도록 허용하되, 과도한 호출 방지)
    const now = Date.now();
    const fieldKey = `lastUpdate_${field}`;
    const lastCall = (window as any)[fieldKey] || 0;
    
    // 전역 rate limiting: 모든 업데이트에 대해 300ms 제한
    if (now - lastUpdateLobbyDataCall.current < 300) {
      return;
    }
    lastUpdateLobbyDataCall.current = now;

    // 개별 필드 rate limiting: 동일 필드에 대해 800ms 제한
    if (now - lastCall < 800) {
      return;
    }
    (window as any)[fieldKey] = now;

    setLobbyData((prev: any) => {
      const newData = { ...prev, [field]: value };
      
      // Check 필드가 입력되면 시간 자동 기록
      if (field.includes('_Check') && value && value.toString().trim() !== '' && (!prev[field] || prev[field].toString().trim() === '')) {
        const rowNumber = field.match(/Row(\d+)_Check/)?.[1];
        if (rowNumber) {
          const timeField = `Row${rowNumber}_CheckedTime`;
          if (!newData[timeField]) {
            newData[timeField] = getCurrentTime12Hour();
          }
        }
      }
      
      // Check 필드가 비워지면 시간도 비움
      if (field.includes('_Check') && (!value || value.toString().trim() === '')) {
        const rowNumber = field.match(/Row(\d+)_Check/)?.[1];
        if (rowNumber) {
          const timeField = `Row${rowNumber}_CheckedTime`;
          newData[timeField] = '';
        }
      }
      
      return newData;
    });
  };

  // 제출 함수 (Rate limiting 적용)
  const handleSubmit = async () => {
    // Rate limiting: 최근 3초 내 호출 방지 (PDF 생성은 무거운 작업)
    // (실수로 두 번 클릭하는 것을 방지하되, 사용자 경험을 해치지 않도록)
    const now = Date.now();
    if (now - lastSubmitCall.current < 3000) {
      alert('⚠️ An error occurred while submitting. Please try again.');
      return;
    }
    lastSubmitCall.current = now;

    // 이미 제출 중이면 중복 호출 방지
    if (loading) {
      return;
    }

    if (!inspectionDate || !selectedOffice) {
      alert('Please select a date and office first.');
      return;
    }

    // 확인 다이얼로그
    const confirmed = confirm('Would you like to submit?');
    if (!confirmed) {
      return;
    }

    setLoading(true);
    setSubmitStatus('Saving...');
    setProgress(10);

    try {
      // 1. PDF 생성
      setSubmitStatus('Submitting...');
      setProgress(30);
      
      // Firebase Auth에서 현재 사용자 확인
      const currentUser = auth.currentUser;
      if (!currentUser) {
        alert('Please log in.');
        setLoading(false);
        setSubmitStatus('');
        return;
      }

      // Firestore에서 사용자 role 확인 (Firebase Rules로 보호됨)
      const userDoc = await getDoc(doc(db, 'users', currentUser.uid));
      if (!userDoc.exists()) {
        alert('User information could not be found.');
        setLoading(false);
        setSubmitStatus('');
        return;
      }

      const userData = userDoc.data();
      if (userData?.role !== 'manager') {
        alert('You do not have access to this page.');
        setLoading(false);
        setSubmitStatus('');
        return;
      }

      // 입력 검증
      if (!inspectionDate || !selectedOffice || !lobbyData) {
        throw new Error('Please fill out all required fields.');
      }

      if (
        typeof inspectionDate !== 'string' ||
        typeof selectedOffice !== 'string' ||
        typeof lobbyData !== 'object' ||
        lobbyData === null ||
        Array.isArray(lobbyData)
      ) {
        throw new Error('Invalid input format');
      }

      // 날짜 형식 검증 (YYYY-MM-DD)
      if (!/^\d{4}-\d{2}-\d{2}$/.test(inspectionDate)) {
        throw new Error('Invalid date format');
      }

      // 날짜 유효성 검증
      const dateObj = new Date(inspectionDate + 'T00:00:00');
      if (isNaN(dateObj.getTime())) {
        throw new Error('Invalid date format');
      }

      const [dateYear, dateMonth, dateDay] = inspectionDate.split('-').map(Number);
      if (dateYear < 1900 || dateYear > 2100 || dateMonth < 1 || dateMonth > 12 || dateDay < 1 || dateDay > 31) {
        throw new Error('Invalid date format');
      }

      const reconstructedDate = `${dateYear}-${String(dateMonth).padStart(2, '0')}-${String(dateDay).padStart(2, '0')}`;
      if (reconstructedDate !== inspectionDate) {
        throw new Error('Invalid date format');
      }

      // 오피스 선택 검증
      if (!['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'].includes(selectedOffice)) {
        throw new Error('Invalid office selection');
      }

      // 날짜 포맷팅 (MM/DD/YYYY)
      const laDate = new Date(dateObj.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
      const month = String(laDate.getMonth() + 1).padStart(2, '0');
      const day = String(laDate.getDate()).padStart(2, '0');
      const year = laDate.getFullYear();
      const formattedDate = `${month}/${day}/${year}`;

      const generatedDate = new Date().toLocaleDateString('en-US', {
        year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit', hour12: true,
      });

      // PDF 생성 (클라이언트 사이드)
      setSubmitStatus('Processing...');
      setProgress(60);
      
      const safeSelectedOffice = selectedOffice.trim().slice(0, 100).replace(/[<>]/g, '');
      const pdfDoc = createLobbyPDFDocument({
        safeSelectedOffice,
        formattedDate,
        lobbyData: lobbyData as Record<string, unknown>,
        generatedDate,
      });

      const blob = await pdf(pdfDoc).toBlob();
        
      // PDF를 Firebase Storage에 저장 (endofday-pdfs 저장)
      setSubmitStatus('Saving...');
      setProgress(70);
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
        
        const filename = `5) ${inspectionDate}_${selectedOffice}_Lobby Inspection Log_${timeStamp}.pdf`;
        const storageRef = ref(storage, `endofday-pdfs/${selectedOffice}/${inspectionDate}/${filename}`);
        
        // PDF 업로드
        await uploadBytes(storageRef, blob);
        
      } catch (storageError: any) {
        // 저장 실패 시 사용자에게 알림
        const errorMsg = storageError?.message || 'Error';
        alert(`An error occurred while submitting. Please try again.: ${errorMsg}`);
        throw storageError;
      }
      
      // 2. 데이터 삭제
      setSubmitStatus('Cleaning up...');
      setProgress(80);
      const docId = `${inspectionDate}_${selectedOffice}_lobby`;
      await deleteDoc(doc(db, "lobby-inspections", docId));
      
      // 3. 폼 초기화
      setLobbyData({});
      setLastSavedData({});

      setSubmitStatus('Complete!');
      setProgress(100);
      
      setTimeout(() => {
        setLoading(false);
        setSubmitStatus('');
        setProgress(0);
        alert('Submitted successfully!');
      }, 2000);

    } catch (error: any) {
      setSubmitStatus('❌ Submission failed: ' + error.message);
      setProgress(0);
      setTimeout(() => {
        setLoading(false);
        setSubmitStatus('');
        alert('❌ Submission failed: ' + error.message);
      }, 3000);
    }
  };

  // 스타일 정의
  const styles = {
    body: {
      fontFamily: "Arial, sans-serif",
      background: 'linear-gradient(135deg, #ffecd2 0%, #fcb69f 100%)',
      color: '#333',
      lineHeight: '1.6',
      overflowX: 'hidden' as const,
      minHeight: '100vh',
      margin: 0,
      padding: 0
    },
    container: {
      maxWidth: '95vw',
      minWidth: '320px',
      width: '100%',
      margin: '20px auto',
      background: 'linear-gradient(145deg, #ffffff 0%, #f8fafc 100%)',
      borderRadius: '20px',
      boxShadow: '0 20px 40px rgba(0, 0, 0, 0.1), 0 0 0 1px rgba(255, 255, 255, 0.1)',
      padding: '25px 2vw 30px 2vw',
      overflowX: 'auto' as const,
      backdropFilter: 'blur(10px)',
      border: '1px solid rgba(255, 255, 255, 0.2)'
    },
    header: {
      fontSize: '1.8rem',
      fontWeight: '700',
      marginBottom: '25px',
      color: '#4a6fa1',
      textAlign: 'center' as const,
      textShadow: '0 2px 4px rgba(0, 0, 0, 0.1)',
      letterSpacing: '-0.5px'
    },
    infoRow: {
      display: 'flex',
      gap: '18px',
      alignItems: 'center',
      justifyContent: 'center',
      marginBottom: '30px',
      flexWrap: 'wrap' as const,
      background: 'linear-gradient(135deg, rgba(255, 154, 158, 0.05) 0%, rgba(254, 207, 239, 0.05) 100%)',
      padding: '20px',
      borderRadius: '15px',
      border: '1px solid rgba(255, 154, 158, 0.1)'
    },
    label: {
      fontWeight: '600',
      color: '#ff9a9e',
      fontSize: '16px'
    },
    input: {
      width: '160px',
      padding: '10px 15px',
      border: '2px solid rgba(255, 154, 158, 0.2)',
      borderRadius: '10px',
      background: 'rgba(255, 255, 255, 0.9)',
      fontSize: '14px',
      fontWeight: '500',
      transition: 'all 0.3s ease'
    },
    table: {
      width: '100%',
      maxWidth: '100%',
      minWidth: '1400px',
      tableLayout: 'auto' as const,
      margin: '0 auto',
      overflowX: 'auto' as const,
      borderRadius: '12px',
      overflow: 'hidden',
      boxShadow: '0 8px 32px rgba(255, 154, 158, 0.1)'
    },
    tableContainer: {
      overflowX: 'auto' as const,
      width: '100%',
      margin: '0 auto',
      borderRadius: '12px',
      background: '#fff',
      boxShadow: '0 8px 32px rgba(255, 154, 158, 0.1)'
    },
    th: {
      background: '#2e3a4e',
      color: '#fff',
      fontWeight: '500',
      padding: '12px 15px'
    },
    td: {
      padding: '12px 15px',
      borderBottom: '1px solid #e1e5ea'
    },
    checkbox: {
      display: 'block',
      margin: '0 auto',
      position: 'relative' as const,
      top: '50%',
      transform: 'translateY(-50%)'
    },
    textInput: {
      width: '100%',
      padding: '6px 10px',
      border: '1px solid #c4cdd5',
      borderRadius: '4px'
    },
    select: {
      width: '160px',
      padding: '10px 15px',
      border: '2px solid rgba(255, 154, 158, 0.2)',
      borderRadius: '10px',
      background: 'rgba(255, 255, 255, 0.9)',
      fontSize: '14px',
      fontWeight: '500',
      transition: 'all 0.3s ease'
    },
    submitButton: {
      display: 'inline-block',
      padding: '12px 32px',
      background: 'linear-gradient(135deg, #ff9a9e 0%, #fecfef 100%)',
      color: '#fff',
      fontSize: '1.1rem',
      fontWeight: '600',
      border: 'none',
      borderRadius: '50px',
      cursor: 'pointer',
      transition: 'all 0.3s ease',
      boxShadow: '0 8px 25px rgba(255, 154, 158, 0.3)',
      position: 'relative' as const,
      overflow: 'hidden',
      letterSpacing: '0.5px'
    },
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
        setInspectionDate(getCurrentCaliforniaTime());

        // office_based 처리: 배열이거나 단일 값일 수 있음
        if (userData?.office_based) {
          const officebasedArray = Array.isArray(userData.office_based) 
            ? userData.office_based 
            : [userData.office_based];
          
          // officeOptions에 포함된 값들만 필터링
          const validOptions = officebasedArray.filter((g: string) => officeOptions.includes(g));
          
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

  // 컬럼 정의
  const COLUMN_NAMES = [
    'Time', 'Check', 'Ipads/Games Working', 'Wipe Ipads/Games', 'Pick Up Litter/Sweep',
    'Entrance Area', 'Pass Out Water', 'Sweep/Vacuum', 'Wipe Ipads/Games',
    'Take Out Trash', 'Wipe Desk Tops', 'Wipe Seats', 'Wipe Windows/Door Handles', 'Checked Time'
  ];

  // 행 헤더 정의
  const ROW_HEADERS = [
    'Manager Inspection',
    '8 am', '9 am', '10 am', '11 am',
    'Manager Inspection',
    '12 pm', '1 pm', 'Sweep/Mop', '2 pm', '3 pm',
    'Manager Inspection',
    '4 pm', '5 pm', '6 pm', 'Sweep/Mop', '7 pm',
    'Deep Clean Manager Inspection'
  ];


  // 인증 확인 중이거나 인증 실패 시 로딩 화면 표시
  if (isAuthorized === null) {
    return (
      <div style={{
        display: 'flex',
        justifyContent: 'center',
        alignItems: 'center',
        height: '100vh',
        background: 'linear-gradient(135deg, #ffecd2 0%, #fcb69f 100%)',
        fontFamily: 'Arial, sans-serif'
      }}>
        <div style={{ textAlign: 'center' }}>
          <div style={{ fontSize: '24px', marginBottom: '20px' }}>🔐</div>
          <div style={{ fontSize: '18px', color: '#333' }}>Verifying authentication...</div>
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
        background: 'linear-gradient(135deg, #ffecd2 0%, #fcb69f 100%)',
        fontFamily: 'Arial, sans-serif'
      }}>
        <div style={{ textAlign: 'center' }}>
          <div style={{ fontSize: '24px', marginBottom: '20px' }}>🚫</div>
          <div style={{ fontSize: '18px', color: '#d32f2f', marginBottom: '10px' }}>You do not have access to this page.</div>
          <div style={{ fontSize: '14px', color: '#666' }}>You do not have access to this page.</div>
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

      <div style={styles.container}>
        <h2 style={styles.header}>
          <span style={{fontSize:'2.8em',verticalAlign:'middle',textShadow:'0 2px 4px rgba(0,0,0,0.1)',margin:'0 12px'}}>🛋️</span> 
          Lobby Inspection Log 
          <span style={{fontSize:'2.8em',verticalAlign:'middle',textShadow:'0 2px 4px rgba(0,0,0,0.1)',margin:'0 12px'}}>🛋️</span>
        </h2>
        
        <div style={{textAlign:'center', fontSize:'14px', color:'#4a6fa1', marginBottom:'20px', fontStyle:'italic',fontWeight:'500',textShadow:'0 1px 2px rgba(0,0,0,0.05)'}}>
          Only perform tasks as needed.<br/>
          <span style={{color:'#4a6fa1'}}>Do not use cavi wipes, use front office wipes.</span>
        </div>

        {/* 정보 입력 섹션 */}
        <div style={styles.infoRow}>
          <label style={styles.label} htmlFor="date">📅 Date:</label>
          <input
            type="date"
            id="date"
            value={inspectionDate}
            onChange={(e: any) => setInspectionDate(e.target.value)}
            style={styles.input}
          />
          {/* officee_based 옵션이 있는 경우에만 Office 표시 */}
          {userOfficeBasedOptions.length > 0 && (
            <>
              <label style={styles.label} htmlFor="office">🏢 Office:</label>
              {userOfficeBasedOptions.length === 1 ? (
                <span style={{
                  ...styles.select,
                  display: 'inline-flex',
                  alignItems: 'center',
                  backgroundColor: '#e9ecef',
                  fontWeight: '600',
                  color: '#4a6fa1'
                }}>
                  {selectedOffice}
                </span>
              ) : (
                <select
                  id="office"
                  value={selectedOffice}
                  onChange={(e: any) => setSelectedOffice(e.target.value)}
                  style={styles.select}
                >
                  <option value="">--Select Office--</option>
                  {userOfficeBasedOptions.map(office => (
                    <option key={office} value={office}>{office}</option>
                  ))}
                </select>
              )}
            </>
          )}
        </div>

        {/* 테이블 - Date와 Office가 모두 선택되었을 때만 표시 */}
        {inspectionDate && selectedOffice && (
          <div style={styles.tableContainer}>
            <table style={styles.table}>
            <thead>
              <tr>
                <th colSpan={2}></th>
                <th colSpan={5} style={{...styles.th, background:'#e3e8f0', fontWeight:'700'}}>
                  <span style={{color:'#4a6fa1'}}>Hourly Cleaning</span>
                </th>
                <th colSpan={6} style={{...styles.th, background:'#f3f7fa', fontWeight:'700'}}>
                  <span style={{color:'#4a6fa1'}}>End of Day Cleaning</span>
                </th>
                <th></th>
              </tr>
              <tr>
                <th style={{...styles.th, background:'#2e3a4e', color:'white'}}>Time</th>
                <th style={{...styles.th, background:'#2e3a4e', color:'white'}}>Check</th>
                <th style={{...styles.th, background:'#2e3a4e', color:'white'}}>Ipads/Games Working</th>
                <th style={{...styles.th, background:'#2e3a4e', color:'white'}}>Wipe Ipads/Games</th>
                <th style={{...styles.th, background:'#2e3a4e', color:'white'}}>Pick Up Litter/Sweep</th>
                <th style={{...styles.th, background:'#2e3a4e', color:'white'}}>Entrance Area</th>
                <th style={{...styles.th, background:'#2e3a4e', color:'white'}}>Pass Out Water</th>
                <th style={{...styles.th, background:'#2e3a4e', color:'white'}}>Sweep/Vacuum</th>
                <th style={{...styles.th, background:'#2e3a4e', color:'white'}}>Wipe Ipads/Games</th>
                <th style={{...styles.th, background:'#2e3a4e', color:'white'}}>Take Out Trash</th>
                <th style={{...styles.th, background:'#2e3a4e', color:'white'}}>Wipe Desk Tops</th>
                <th style={{...styles.th, background:'#2e3a4e', color:'white'}}>Wipe Seats</th>
                <th style={{...styles.th, background:'#2e3a4e', color:'white'}}>Wipe Windows/Door Handles</th>
                <th style={{...styles.th, background:'#2e3a4e', color:'white'}}>Checked Time</th>
              </tr>
              <tr>
                <th colSpan={2}></th>
                <th colSpan={5} style={{...styles.th, background:'#f7fafd', fontSize:'13px', fontWeight:'400', color:'#4a6fa1', textAlign:'center'}}>
                  Check/Perform each hour
                </th>
                <th colSpan={6} style={{...styles.th, background:'#fafdff', fontSize:'13px', fontWeight:'400', color:'#4a6fa1', textAlign:'center'}}>
                  or as needed
                </th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {ROW_HEADERS.map((header, rowIndex) => {
                const isManagerInspection = header.includes('Manager Inspection');
                const isSweepMop = header.includes('Sweep/Mop');
                const checkOptions = (OFFICE_CHECK_OPTIONS as any)[selectedOffice] || [];
                
                return (
                  <tr key={rowIndex} style={{background: rowIndex % 2 === 0 ? '#f9fbfc' : 'white'}}>
                    <td style={styles.td}>{header}</td>
                    <td style={styles.td}>
                      {isManagerInspection ? (
                        checkOptions.length > 0 ? (
                          <select
                            value={lobbyData[`Row${rowIndex + 1}_Check`] || ''}
                            onChange={(e: any) => updateLobbyData(`Row${rowIndex + 1}_Check`, e.target.value)}
                            style={styles.select}
                          >
                            <option value=""></option>
                            {checkOptions.map((opt: any) => (
                              <option key={opt} value={opt}>{opt}</option>
                            ))}
                          </select>
                        ) : (
                          <input
                            type="checkbox"
                            checked={lobbyData[`Row${rowIndex + 1}_Check`] === true}
                            onChange={(e: any) => updateLobbyData(`Row${rowIndex + 1}_Check`, e.target.checked)}
                            style={styles.checkbox}
                          />
                        )
                      ) : (
                        <input
                          type="text"
                          value={lobbyData[`Row${rowIndex + 1}_Check`] || ''}
                          onChange={(e: any) => updateLobbyData(`Row${rowIndex + 1}_Check`, e.target.value)}
                          style={styles.textInput}
                        />
                      )}
                    </td>
                    {COLUMN_NAMES.slice(2, -1).map((columnName, colIndex) => {
                      if (isSweepMop) {
                        return <td key={colIndex} style={styles.td}></td>;
                      }
                      
                      return (
                        <td key={colIndex} style={styles.td}>
                          <input
                            type="checkbox"
                            checked={lobbyData[`Row${rowIndex + 1}_Col${colIndex + 3}`] === true}
                            onChange={(e: any) => updateLobbyData(`Row${rowIndex + 1}_Col${colIndex + 3}`, e.target.checked)}
                            style={styles.checkbox}
                          />
                        </td>
                      );
                    })}
                    <td style={styles.td}>
                      <span style={{color: '#666', fontSize: '12px'}}>
                        {lobbyData[`Row${rowIndex + 1}_CheckedTime`] || ''}
                      </span>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        )}

        {/* 제출 버튼 - Date와 Office가 모두 선택되었을 때만 표시 */}
        {inspectionDate && selectedOffice && (
          <div style={{textAlign:'center', marginTop:'18px'}}>
            <button 
              type="button" 
              onClick={handleSubmit}
              style={{
                ...styles.submitButton,
                background: loading ? 'linear-gradient(135deg, #6c757d 0%, #5a6268 100%)' : styles.submitButton.background,
                cursor: loading ? 'not-allowed' : 'pointer'
              }}
              disabled={loading || !selectedOffice}
            >
              🚀 Submit
            </button>
          </div>
        )}
      </div>
      </div>
    </>
  );
}


