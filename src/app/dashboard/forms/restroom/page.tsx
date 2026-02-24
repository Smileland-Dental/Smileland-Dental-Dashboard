'use client'

import React, { useState, useEffect, useCallback } from "react";
import { doc, setDoc, getDoc, deleteDoc, onSnapshot, collection, getDocs } from "firebase/firestore";
import { db, auth } from "@/lib/firebase.config";
import { getStorage, ref, uploadBytes } from 'firebase/storage';
import { onAuthStateChanged } from 'firebase/auth';
import { pdf, Document, Page, View, Text, StyleSheet } from '@react-pdf/renderer';

export default function RestroomInspection() {
  const [loading, setLoading] = useState(false);
  const [submitStatus, setSubmitStatus] = useState('');
  const [progress, setProgress] = useState(0);
  const [isUpdatingFromFirebase, setIsUpdatingFromFirebase] = useState(false);
  const [isAuthorized, setIsAuthorized] = useState<boolean | null>(null); // null: 확인 중, true: 인증됨, false: 인증 실패
  const [userOfficesOptions, setUserOfficesOptions] = useState<string[]>([]); // 사용자의 offices 옵션들
  
  // 사용자 세션 ID 생성 (페이지 로드 시 한 번만)
  const [userSessionId] = useState(() => Math.random().toString(36).substr(2, 9));
  
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
  const [inspectionDate, setInspectionDate] = useState('');

  // 오피스 및 화장실 선택 상태
  const [selectedOffice, setSelectedOffice] = useState('');
  const [selectedRestroom, setSelectedRestroom] = useState('');
  
  // 오피스 옵션 (알파벳 순)
  const officeOptions = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];
  const restroomOptions = ['1', '2', '3'];

  // Office별 Check 옵션 정의 (알파벳 순)
  const OFFICE_CHECK_OPTIONS = {
    "Bernard": ["Carmen", "Cynthia", "Elisa", "Ranjit"],
    "California": ["Kindal"],
    "Delano": ["Helen", "Stephanie"],
    "Fresno": ["Cynthia"],
    "Ming": ["Kindal", "Hopie", "Marbella"],
    "Ortho": ["Kindal"],
    "Tulare": ["Dianne"],
    "Visalia": ["Dianne"]
  };


  // 컬럼 순서: Time, Check, Pick up Paper, Wipe Sinks and Mirrors, Wipe Toilets, Wipe Baby Table, Empty Trash, Toilet Paper, Soap, Toilet Seat Covers, Refresh Spray, Checked Time
  const COLUMN_NAMES = [
    'Time', 'Check', 'Pick up Paper', 'Wipe Sinks and Mirrors', 'Wipe Toilets',
    'Wipe Baby Table', 'Empty Trash', 'Toilet Paper', 'Soap',
    'Toilet Seat Covers', 'Refresh Spray', 'Checked Time'
  ];

  // 행 헤더 정의 (구글시트 행 순서와 일치)
  const ROW_HEADERS = [
    'Manager Inspection',
    '8 am', '9 am', '10 am', '11 am',
    'Manager Inspection',
    '12 pm', '1 pm', 'Sweep/Mop', '2 pm', '3 pm',
    'Manager Inspection',
    '4 pm', '5 pm', '6 pm', 'Sweep/Mop', '7 pm',
    'Deep Clean Manager Inspection'
  ];

  // --- PDF 생성 관련 상수/스타일 ---
  const PDF_COLUMN_NAMES = [
    'Time', 'Check', 'Pick up Paper', 'Wipe Sinks and Mirrors', 'Wipe Toilets',
    'Wipe Baby Table', 'Empty Trash', 'Toilet Paper', 'Soap',
    'Toilet Seat Covers', 'Refresh Spray', 'Checked Time',
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
    cell: { padding: 3, fontSize: 6, flex: 1, borderRightWidth: 0.5, borderColor: '#333', justifyContent: 'center', alignItems: 'center' },
    cellFlex2: { flex: 2 },
    cellFlex4: { flex: 4 },
    cellFlex5: { flex: 5 },
    cellFlex1: { flex: 1 },
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

  function createRestroomPDFDocument(props: {
    safeInspectionDate: string;
    safeSelectedOffice: string;
    safeSelectedRestroom: string;
    restroomData: Record<string, unknown>;
    generatedDate: string;
  }) {
    const { safeInspectionDate, safeSelectedOffice, safeSelectedRestroom, restroomData, generatedDate } = props;
    const s = pdfStyles;

    // Row 1: SPOTLESS / STOCKED
    const row1 = React.createElement(View, { key: 'r1', style: [s.row, s.cellGray] },
      React.createElement(View, { style: [s.cell, s.cellFlex2] }, React.createElement(Text, null, '')),
      React.createElement(View, { style: [s.cell, s.cellFlex4, s.cellBold] }, React.createElement(Text, null, 'SPOTLESS')),
      React.createElement(View, { style: [s.cell, s.cellFlex5, s.cellBold] }, React.createElement(Text, null, 'STOCKED')),
      React.createElement(View, { style: [s.cell, s.cellFlex1] }, React.createElement(Text, null, '')),
    );

    // Row 2: Column names
    const row2 = React.createElement(View, { key: 'r2', style: [s.row, s.cellGray] },
      ...PDF_COLUMN_NAMES.map((name, i) =>
        React.createElement(View, { key: i, style: s.cell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, name))
      ),
    );

    // Row 3: Perform each hour / Replenish as needed
    const row3 = React.createElement(View, { key: 'r3', style: [s.row, s.cellGray] },
      React.createElement(View, { style: [s.cell, s.cellFlex2] }, React.createElement(Text, null, '')),
      React.createElement(View, { style: [s.cell, s.cellFlex4] }, React.createElement(Text, null, 'Perform each hour')),
      React.createElement(View, { style: [s.cell, s.cellFlex5] }, React.createElement(Text, null, 'Replenish as needed')),
      React.createElement(View, { style: [s.cell, s.cellFlex1] }, React.createElement(Text, null, '')),
    );

    // Data rows
    const dataRows = PDF_ROW_HEADERS.map((header, rowIndex) => {
      const isMopRow = header.toLowerCase().includes('sweep/mop');
      const checkValue = safeStr(restroomData[`Row${rowIndex + 1}_Check`], 50);
      const checkedTime = safeStr(restroomData[`Row${rowIndex + 1}_CheckedTime`], 50);
      const taskCells = PDF_COLUMN_NAMES.slice(2, -1).map((_, colIndex) => {
        if (isMopRow) return React.createElement(View, { key: colIndex, style: s.cell }, React.createElement(Text, null, ''));
        const v = restroomData[`Row${rowIndex + 1}_Col${colIndex + 3}`];
        return React.createElement(View, { key: colIndex, style: s.cell }, React.createElement(Text, null, isChecked(v) ? 'O' : ''));
      });
      return React.createElement(View, { key: rowIndex, style: s.row },
        React.createElement(View, { style: s.cell }, React.createElement(Text, null, header)),
        React.createElement(View, { style: s.cell }, React.createElement(Text, null, checkValue)),
        ...taskCells,
        React.createElement(View, { style: s.cell }, React.createElement(Text, null, checkedTime)),
      );
    });

    const table = React.createElement(View, null, row1, row2, row3, ...dataRows);

    const header = React.createElement(View, { style: s.header },
      React.createElement(Text, { style: s.headerTitle }, 'Restroom Inspection Log'),
      React.createElement(Text, { style: s.headerSub }, `${safeSelectedOffice} - Restroom ${safeSelectedRestroom} (${safeInspectionDate})`),
    );

    const footer = React.createElement(View, { style: s.footer }, React.createElement(Text, null, `Generated: ${generatedDate}`));

    return React.createElement(Document, null,
      React.createElement(Page, { size: 'A4', orientation: 'landscape', style: s.page }, header, table, footer),
    );
  }

  function sanitizeFilename(filename: string): string {
    return filename.replace(/[^a-zA-Z0-9._-]/g, '_').replace(/\.\./g, '_').slice(0, 255);
  }

  // 모든 화장실 검사 데이터 상태
  const [restroomData, setRestroomData] = useState<{ [key: string]: any }>({});

  // 12시간 형식 시간 가져오기 (캘리포니아 시간대)
  const getCurrentTime12Hour = () => {
    const now = new Date();
    const timeFormatter = new Intl.DateTimeFormat('en-US', {
      timeZone: 'America/Los_Angeles',
      hour: 'numeric',
      minute: '2-digit',
      hour12: true
    });
    const timeString = timeFormatter.format(now);
    
    // 시간 문자열 파싱 (예: "1:45 PM" 또는 "12:30 AM")
    const match = timeString.match(/(\d+):(\d+)\s*(AM|PM)/i);
    if (match) {
      const hours = match[1];
      const minutes = match[2];
      const ampm = match[3].toUpperCase();
      return `${hours}:${minutes} ${ampm}`;
    }
    
    // 폴백: 기존 방식 사용
    const laTime = new Date(now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
    let hours = laTime.getHours();
    const minutes = laTime.getMinutes();
    const ampm = hours >= 12 ? 'PM' : 'AM';
    
    hours = hours % 12;
    hours = hours ? hours : 12;
    
    return `${hours}:${minutes.toString().padStart(2, '0')} ${ampm}`;
  };

  // 자동 저장 함수
  const autoSave = useCallback(async () => {
    if (!inspectionDate || !selectedRestroom || isUpdatingFromFirebase) return;

    // 데이터가 실제로 변경되었는지 확인
    const hasChanges = JSON.stringify(restroomData) !== JSON.stringify(lastSavedData);
    if (!hasChanges) return;

    try {
      const dataToSave = {
        inspectionDate,
        selectedOffice,
        selectedRestroom,
        ...restroomData,
        timestamp: new Date().toISOString(),
        autoSaved: true,
        lastUpdatedBy: userSessionId
      };

      const docId = `${inspectionDate}_${selectedOffice}_${selectedRestroom}`;
      await setDoc(doc(db, "restroom-inspections", docId), dataToSave);
      
      // 저장 성공 시 마지막 저장된 데이터 업데이트
      setLastSavedData({ ...restroomData });
      
    } catch (error) {
      // Auto-save error silently handled
    }
  }, [inspectionDate, selectedOffice, selectedRestroom, restroomData, lastSavedData, isUpdatingFromFirebase]);

  // 데이터 변경 시에만 자동 저장
  useEffect(() => {
    // 초기 로드 시에는 저장하지 않음
    if (Object.values(restroomData).some(value => value !== '')) {
      autoSave();
    }
  }, [restroomData]);

  // 데이터 로드
  const loadData = async () => {
    if (!inspectionDate || !selectedOffice || !selectedRestroom) return;

    try {
      const docId = `${inspectionDate}_${selectedOffice}_${selectedRestroom}`;
      const docSnap = await getDocs(collection(db, "restroom-inspections")).then((snapshot: any) => {
        const foundDoc = snapshot.docs.find((d: any) => d.id === docId);
        return foundDoc ? { exists: () => true, data: () => foundDoc.data() } : { exists: () => false, data: undefined };
      });
      
      if (docSnap.exists() && docSnap.data) {
        const data = docSnap.data();
        
        // Firebase에서 업데이트되는 동안 자동 저장 방지
        setIsUpdatingFromFirebase(true);
        
        setRestroomData((prevData: any) => ({
          ...prevData,
          ...data
        }));
        
        // 로드된 데이터를 마지막 저장된 데이터로 설정
        setLastSavedData({ ...data });
        
        // 짧은 지연 후 Firebase 업데이트 플래그 해제
        setTimeout(() => {
          setIsUpdatingFromFirebase(false);
        }, 100);
      } else {
        // 데이터가 없으면 초기화
        const initialData = {};
        setRestroomData(initialData);
        setLastSavedData(initialData);

                
        // 플래그 해제
        setTimeout(() => {
          setIsUpdatingFromFirebase(false);
        }, 100);
      }
      
    } catch (error: any) {
      // 에러 발생 시에도 플래그 해제
      setTimeout(() => {
        setIsUpdatingFromFirebase(false);
      }, 100);
      
      setSubmitStatus('Error loading data: ' + error.message);
      setTimeout(() => setSubmitStatus(''), 3000);
    }
  };

  // Restroom 변경 처리 - 데이터 초기화 후 변경
  const handleRestroomChange = (newRestroom: string) => {
    // 초기화 중 자동 저장 방지
    setIsUpdatingFromFirebase(true);
    // 이전 데이터 초기화
    setRestroomData({});
    setLastSavedData({});
    // 새 restroom 설정
    setSelectedRestroom(newRestroom);
  };

  // 날짜, 오피스, 화장실 변경 시 데이터 로드
  useEffect(() => {
    if (selectedRestroom) {
      loadData();
    }
  }, [inspectionDate, selectedOffice, selectedRestroom]);

  // 실시간 데이터 동기화
  useEffect(() => {
    if (!inspectionDate || !selectedOffice || !selectedRestroom) return;

    const docId = `${inspectionDate}_${selectedOffice}_${selectedRestroom}`;
    const docRef = doc(db, "restroom-inspections", docId);
    
    const unsubscribe = onSnapshot(docRef, (docSnap: any) => {
      if (docSnap.exists() && docSnap.data) {
        const data = docSnap.data();
        
        // Firebase에서 업데이트되는 동안 자동 저장 방지
        setIsUpdatingFromFirebase(true);
        
        setRestroomData((prevData: any) => {
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
  }, [inspectionDate, selectedOffice, selectedRestroom]);

  // 데이터 업데이트 함수
  const updateRestroomData = (field: string, value: any) => {
    setRestroomData((prev: any) => {
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
    if (!selectedOffice || !selectedRestroom) {
      alert('Please select an office and restroom first.');
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
      if (!inspectionDate || !selectedOffice || !selectedRestroom || !restroomData) {
        throw new Error('Please fill out all required fields.');
      }

      if (
        typeof inspectionDate !== 'string' ||
        typeof selectedOffice !== 'string' ||
        typeof selectedRestroom !== 'string' ||
        typeof restroomData !== 'object' ||
        restroomData === null ||
        Array.isArray(restroomData)
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

      // 화장실 선택 검증
      if (!['1', '2', '3'].includes(selectedRestroom)) {
        throw new Error('Invalid restroom selection');
      }

      // 1. PDF 생성
      setSubmitStatus('Processing...');
      setProgress(30);

      const safeInspectionDate = inspectionDate.trim().slice(0, 20).replace(/[<>]/g, '');
      const safeSelectedOffice = selectedOffice.trim().slice(0, 100).replace(/[<>]/g, '');
      const safeSelectedRestroom = selectedRestroom.trim().slice(0, 50).replace(/[<>]/g, '');

      const generatedDate = new Date().toLocaleDateString('en-US', {
        year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit', hour12: true,
      });

      // PDF 생성 (클라이언트 사이드)
      setSubmitStatus('Processing...');
      setProgress(60);
      
      const pdfDoc = createRestroomPDFDocument({
        safeInspectionDate,
        safeSelectedOffice,
        safeSelectedRestroom,
        restroomData: restroomData as Record<string, unknown>,
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
        
        const filename = `6) ${inspectionDate}_${selectedOffice}_Restroom_${selectedRestroom}_Inspection_Log_${timeStamp}.pdf`;
        const storageRef = ref(storage, `endofday-pdfs/${selectedOffice}/${inspectionDate}/${filename}`);
        
        // PDF 업로드
        await uploadBytes(storageRef, blob);
        
      } catch (storageError: any) {
        const errorMsg = storageError?.message || 'Error';
        alert(`An error occurred while submitting. Please try again.: ${errorMsg}`);
        throw storageError;
      }
      
      // 2. 데이터 삭제
      setSubmitStatus('Cleaning up...');
      setProgress(80);
      const docId = `${inspectionDate}_${selectedOffice}_${selectedRestroom}`;
      await deleteDoc(doc(db, "restroom-inspections", docId));
      
      // 3. 폼 초기화
      setRestroomData({});
      setLastSavedData({});

      setSubmitStatus('Complete!');
      setProgress(100);
      
      // 2초 후 모달 닫기
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
  const styles: { [key: string]: React.CSSProperties } = {
    body: {
      fontFamily: "Arial, sans-serif",
      backgroundColor: '#C2E6E6',
      color: '#333',
      lineHeight: '1.6',
      overflowX: 'hidden',
      minHeight: '100vh',
      margin: 0,
      padding: 0
    },
    container: {
      maxWidth: '95vw',
      minWidth: '320px',
      width: '100%',
      margin: '20px auto',
      backgroundColor: '#ffffff',
      borderRadius: '20px',
      boxShadow: '0 20px 40px rgba(0, 0, 0, 0.1), 0 0 0 1px rgba(255, 255, 255, 0.1)',
      padding: '25px 2vw 30px 2vw',
      overflowX: 'auto',
      backdropFilter: 'blur(10px)',
      border: '1px solid rgba(255, 255, 255, 0.2)'
    },
    header: {
      fontSize: '1.8rem',
      fontWeight: '700',
      marginBottom: '10px',
      color: '#4a6fa1',
      textAlign: 'center',
      textShadow: '0 2px 4px rgba(0, 0, 0, 0.1)',
      letterSpacing: '-0.5px'
    },
    infoRow: {
      display: 'flex',
      gap: '18px',
      alignItems: 'center',
      justifyContent: 'center',
      marginBottom: '30px',
      flexWrap: 'wrap',
      backgroundColor: 'rgba(194, 230, 230, 0.5)',
      padding: '20px',
      borderRadius: '15px',
      border: '1px solid rgba(74, 111, 161, 0.1)'
    },
    label: {
      fontWeight: '600',
      color: '#4a6fa1',
      fontSize: '16px'
    },
    input: {
      width: '160px',
      padding: '10px 15px',
      border: '2px solid rgba(74, 111, 161, 0.2)',
      borderRadius: '10px',
      backgroundColor: 'rgba(255, 255, 255, 0.9)',
      fontSize: '14px',
      fontWeight: '500',
      transition: 'all 0.3s ease'
    },
    select: {
      width: '100px',
      padding: '10px 15px',
      border: '2px solid rgba(74, 111, 161, 0.2)',
      borderRadius: '10px',
      backgroundColor: 'rgba(255, 255, 255, 0.9)',
      fontSize: '14px',
      fontWeight: '500',
      transition: 'all 0.3s ease'
    },
    tableContainer: {
      overflowX: 'auto',
      width: '100%',
      margin: '0 auto',
      borderRadius: '12px',
      backgroundColor: '#ffffff',
      boxShadow: '0 8px 32px rgba(74, 111, 161, 0.1)'
    },
    table: {
      width: '100%',
      maxWidth: '100%',
      minWidth: '1400px',
      tableLayout: 'auto',
      margin: '0 auto',
      overflowX: 'auto',
      borderRadius: '12px',
      overflow: 'hidden',
      boxShadow: '0 8px 32px rgba(74, 111, 161, 0.1)',
      borderCollapse: 'collapse'
    },
    th: {
      padding: '12px 15px',
      backgroundColor: '#2e3a4e',
      color: '#fff',
      fontWeight: '500',
      textAlign: 'center'
    },
    td: {
      padding: '12px 15px',
      borderBottom: '1px solid #e1e5ea',
      textAlign: 'center',
      verticalAlign: 'middle'
    },
    checkboxCell: {
      display: 'flex',
      justifyContent: 'center',
      alignItems: 'center',
      height: '100%',
      padding: 0
    },
    checkbox: {
      margin: 0,
      position: 'relative',
      left: 0
    },
    textInput: {
      width: '100%',
      padding: '6px 10px',
      border: '1px solid #c4cdd5',
      borderRadius: '4px',
      fontSize: '12px'
    },
    submitButton: {
      display: 'inline-block',
      padding: '12px 32px',
      backgroundColor: '#C2E6E6',
      color: '#333',
      fontSize: '1.1rem',
      fontWeight: '600',
      border: 'none',
      borderRadius: '50px',
      cursor: 'pointer',
      transition: 'all 0.3s ease',
      boxShadow: '0 8px 25px rgba(74, 111, 161, 0.3)',
      position: 'relative',
      overflow: 'hidden',
      letterSpacing: '0.5px'
    },
  };

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

        if (userData?.role !== 'Manager' && userData?.role !== 'Employee') {
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

        // offices 처리: 배열이거나 단일 값일 수 있음
        if (userData?.offices) {
          const officesArray = Array.isArray(userData.offices) 
            ? userData.offices 
            : [userData.offices];
          
          // officeOptions에 포함된 값들만 필터링
          const validOptions = officesArray.filter((g: string) => officeOptions.includes(g));
          
          if (validOptions.length > 0) {
            setUserOfficesOptions(validOptions);
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
        backgroundColor: '#C2E6E6',
        fontFamily: 'Arial, sans-serif'
      } as React.CSSProperties}>
        <div style={{ textAlign: 'center' } as React.CSSProperties}>
          <div style={{ fontSize: '24px', marginBottom: '20px' } as React.CSSProperties}>🔐</div>
          <div style={{ fontSize: '18px', color: '#333' } as React.CSSProperties}>Verifying authentication...</div>
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
        backgroundColor: '#C2E6E6',
        fontFamily: 'Arial, sans-serif'
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
          } as React.CSSProperties}>
            <div style={{
              backgroundColor: "white",
              padding: "40px",
              borderRadius: "20px",
              textAlign: "center",
              boxShadow: "0 20px 40px rgba(0, 0, 0, 0.3)",
              maxWidth: "400px",
              width: "90%"
            } as React.CSSProperties}>
              <div style={{
                fontSize: "18px",
                fontWeight: "600",
                color: "#4a6fa1",
                marginBottom: "20px"
              } as React.CSSProperties}>
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
                } as React.CSSProperties}>
                  <div style={{
                    width: `${progress}%`,
                    height: "100%",
                    background: "linear-gradient(90deg, #4a90e2, #51cf66)",
                    transition: "width 0.3s ease",
                    borderRadius: "4px"
                  } as React.CSSProperties} />
                </div>
              )}
              <div style={{
                fontSize: "14px",
                color: "#666",
                marginTop: "10px"
              } as React.CSSProperties}>
                {progress}%
              </div>
            </div>
          </div>
        )}

        <div style={styles.container}>
          <h2 style={styles.header}>
            <span style={{fontSize:'2.8em',verticalAlign:'middle',textShadow:'0 2px 4px rgba(0,0,0,0.1)',margin:'0 12px'} as React.CSSProperties}>🚻</span> 
            Restroom Inspection Log 
            <span style={{fontSize:'2.8em',verticalAlign:'middle',textShadow:'0 2px 4px rgba(0,0,0,0.1)',margin:'0 12px'} as React.CSSProperties}>🚻</span>
          </h2>
          
          <div style={{textAlign:'center', fontSize:'14px', color:'#4a6fa1', marginBottom:'20px', fontStyle:'italic',fontWeight:'500',textShadow:'0 1px 2px rgba(0,0,0,0.05)'} as React.CSSProperties}>
            Only perform tasks as needed.
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
            {/* offices 옵션이 있는 경우에만 Office 표시 */}
            {userOfficesOptions.length > 0 && (
              <>
                <label style={styles.label} htmlFor="office">🏢 Office:</label>
                {userOfficesOptions.length === 1 ? (
                  <span style={{
                    ...styles.select,
                    display: 'inline-flex',
                    alignItems: 'center',
                    backgroundColor: '#e9ecef',
                    fontWeight: '600',
                    color: '#4a6fa1'
                  } as React.CSSProperties}>
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
                    {userOfficesOptions.map(office => (
                      <option key={office} value={office}>{office}</option>
                    ))}
                  </select>
                )}
              </>
            )}
            <label style={styles.label} htmlFor="restroom">🚻 Restroom:</label>
            <select
              id="restroom"
              value={selectedRestroom}
              onChange={(e: any) => handleRestroomChange(e.target.value)}
              style={styles.select}
            >
              <option value="" disabled hidden>Select Restroom</option>
              {restroomOptions.map(restroom => (
                <option key={restroom} value={restroom}>{restroom}</option>
              ))}
            </select>
          </div>

          {/* 검사 테이블 */}
          {selectedOffice && selectedRestroom && (
            <div style={styles.tableContainer}>
              <table style={styles.table}>
                <thead>
                  <tr>
                    <th colSpan={2}></th>
                    <th colSpan={4} style={{backgroundColor:'#e3e8f0',fontWeight:'700'} as React.CSSProperties}>
                      <span style={{color:'#4a6fa1'} as React.CSSProperties}>SPOTLESS</span>
                    </th>
                    <th colSpan={5} style={{backgroundColor:'#f3f7fa',fontWeight:'700'} as React.CSSProperties}>
                      <span style={{color:'#4a6fa1'} as React.CSSProperties}>STOCKED</span>
                    </th>
                    <th></th>
                  </tr>
                  <tr>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Time</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Check</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Pick up Paper</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Wipe Sinks and Mirrors</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Wipe Toilets</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Wipe Baby Table</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Empty Trash</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Toilet Paper</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Soap</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Toilet Seat Covers</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Refresh Spray</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Checked Time</th>
                  </tr>
                  <tr>
                    <th colSpan={2}></th>
                    <th colSpan={4} style={{backgroundColor:'#f7fafd',fontSize:'13px',fontWeight:'400',color:'#4a6fa1',textAlign:'center'} as React.CSSProperties}>
                      Perform each hour
                    </th>
                    <th colSpan={5} style={{backgroundColor:'#fafdff',fontSize:'13px',fontWeight:'400',color:'#4a6fa1',textAlign:'center'} as React.CSSProperties}>
                      Replenish as needed
                    </th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {ROW_HEADERS.map((header: string, rowIndex: number) => {
                    const isMopRow = header.toLowerCase().includes('sweep/mop');
                    const isManagerInspection = header.includes('Manager Inspection');
                    const checkOptions = (OFFICE_CHECK_OPTIONS as any)[selectedOffice] || [];
                    
                    return (
                      <tr key={rowIndex}>
                        <td style={styles.td}>{header}</td>
                        <td style={styles.td}>
                          {isManagerInspection ? (
                            checkOptions.length > 0 ? (
                              <select
                                value={restroomData[`Row${rowIndex + 1}_Check`] || ''}
                                onChange={(e: any) => updateRestroomData(`Row${rowIndex + 1}_Check`, e.target.value)}
                                style={styles.textInput}
                              >
                                <option value=""></option>
                                {checkOptions.map((opt: any) => (
                                  <option key={opt} value={opt}>{opt}</option>
                                ))}
                              </select>
                            ) : (
                              <input
                                type="checkbox"
                                checked={restroomData[`Row${rowIndex + 1}_Check`] === true}
                                onChange={(e: any) => updateRestroomData(`Row${rowIndex + 1}_Check`, e.target.checked)}
                                style={styles.checkbox}
                              />
                            )
                          ) : (
                            <input
                              type="text"
                              value={restroomData[`Row${rowIndex + 1}_Check`] || ''}
                              onChange={(e: any) => updateRestroomData(`Row${rowIndex + 1}_Check`, e.target.value)}
                              style={styles.textInput}
                            />
                          )}
                        </td>
                        {COLUMN_NAMES.slice(2, -1).map((_, colIndex) => {
                          if (isMopRow) {
                            return <td key={colIndex} style={styles.td}></td>;
                          }
                          return (
                            <td key={colIndex} style={styles.td}>
                              <input
                                type="checkbox"
                                checked={restroomData[`Row${rowIndex + 1}_Col${colIndex + 3}`] === true}
                                onChange={(e: any) => updateRestroomData(`Row${rowIndex + 1}_Col${colIndex + 3}`, e.target.checked)}
                                style={styles.checkbox}
                              />
                            </td>
                          );
                        })}
                        <td style={styles.td}>
                          <span>{restroomData[`Row${rowIndex + 1}_CheckedTime`] || ''}</span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}

          {/* 제출 버튼 */}
          {selectedOffice && selectedRestroom && (
            <div style={{textAlign:'center',marginTop:'30px'} as React.CSSProperties}>
              <button
                type="button"
                onClick={handleSubmit}
                disabled={loading || !selectedOffice}
                style={{
                  ...styles.submitButton,
                  backgroundColor: loading || !selectedOffice ? '#bdc3c7' : '#C2E6E6',
                  cursor: loading || !selectedOffice ? 'not-allowed' : 'pointer'
                } as React.CSSProperties}
              >
                🚀 Submit
              </button>
            </div>
          )}

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
            } as React.CSSProperties}>
              {submitStatus}
            </div>
          )}
        </div>
      </div>
    </>
  );
}






