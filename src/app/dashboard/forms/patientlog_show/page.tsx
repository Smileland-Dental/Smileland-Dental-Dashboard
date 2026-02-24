'use client'

import React, { useState, useEffect, useRef } from "react";
import { doc, collection, getDocs, getDoc, updateDoc, deleteDoc, query, where } from "firebase/firestore";
import { db, auth } from "@/lib/firebase.config";
import { onAuthStateChanged } from 'firebase/auth';
import { getStorage, ref, uploadBytes } from 'firebase/storage';
import { pdf, Document, Page, View, Text, StyleSheet } from '@react-pdf/renderer';

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

export default function ShowCheckSystem() {
  // 입력 값 검증 함수
  const validateInput = React.useCallback((field: string, value: any): any => {
    // 문자열 필드 길이 제한
    if (typeof value === 'string') {
      const maxLengths: { [key: string]: number } = {
        name: 100,
        startDate: 20,
        endDate: 20,
        selectedOffice: 50
      };
      
      const maxLength = maxLengths[field] || 500;
      if (value.length > maxLength) {
        return value.slice(0, maxLength);
      }
      
      // 날짜 필드 형식 검증 (YYYY-MM-DD)
      if ((field === 'startDate' || field === 'endDate') && value) {
        const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
        if (!dateRegex.test(value)) {
          // 잘못된 형식이면 빈 문자열 반환
          return '';
        }
      }
      
      // Office 값 whitelist 검증
      if (field === 'selectedOffice' && value && value !== 'All') {
        const validOffices = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];
        if (!validOffices.includes(value)) {
          return '';
        }
      }
      
      // 제어 문자 제거 (XSS 방지)
      return value.replace(/[\x00-\x1F\x7F-\x9F]/g, '');
    }
    
    return value;
  }, []);

  // 상태 관리
  const [loading, setLoading] = useState(false);
  const [appointments, setAppointments] = useState<any[]>([]);
  const appointmentsRef = useRef<any[]>([]); // 최신 appointments 추적
  const [filteredAppointments, setFilteredAppointments] = useState<any[]>([]);
  const [startDate, setStartDate] = useState(new Date().toISOString().split('T')[0]);
  const [endDate, setEndDate] = useState(new Date().toISOString().split('T')[0]);
  const [selectedOffice, setSelectedOffice] = useState('');
  const [name, setName] = useState('');
  const [pdfLoading, setPdfLoading] = useState(false);
  const [submitStatus, setSubmitStatus] = useState('');
  const [progress, setProgress] = useState(0);
  const [isAuthorized, setIsAuthorized] = useState<boolean | null>(null); // null: 확인 중, true: 인증됨, false: 인증 실패

  // Rate limiting을 위한 ref
  const lastUpdateStatusCall = useRef<number>(0);
  const lastLoadAppointmentsCall = useRef<number>(0);
  const lastGeneratePDFCall = useRef<number>(0);
  const loadAppointmentsTimeoutRef = useRef<NodeJS.Timeout | null>(null);

  // Office 옵션

  const officeOptions = ['All', 'Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

  // --- PDF 생성 관련 상수/스타일 ---
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
    footer: { marginTop: 15, paddingTop: 8, borderTopWidth: 1, borderColor: '#ddd', alignItems: 'center', fontSize: 8, color: '#666' },
  });

  // PDF 생성 유틸 함수
  function safeStr(v: unknown, max: number): string {
    if (v == null) return '';
    return String(v).trim().slice(0, max).replace(/[<>]/g, '');
  }

  function createShowCheckPDFDocument(props: {
    safeStartDate: string;
    safeEndDate: string;
    safeSelectedOffice: string;
    safeGeneratedBy: string;
    safeAppointments: any[];
    showCount: number;
    noShowCount: number;
    showRate: string;
    generatedDate: string;
  }) {
    const { safeStartDate, safeEndDate, safeSelectedOffice, safeGeneratedBy, safeAppointments, showCount, noShowCount, showRate, generatedDate } = props;
    const s = pdfStyles;

    const dateRange = safeStartDate === safeEndDate ? safeStartDate : `${safeStartDate} to ${safeEndDate}`;

    // 헤더
    const header = React.createElement(View, { style: s.header },
      React.createElement(Text, { style: s.headerTitle }, 'Attendance Show Check'),
    );

    // 정보 섹션
    const infoSection = React.createElement(View, { style: s.infoSection },
      React.createElement(Text, { style: s.infoItem }, `Appointment Date Range: ${dateRange}`),
      React.createElement(Text, { style: s.infoItem }, `Appointment Office: ${safeSelectedOffice}`),
      React.createElement(Text, { style: s.infoItem }, `Checked by: ${safeGeneratedBy}`),
      React.createElement(Text, { style: s.infoItem }, `Total Appointments: ${safeAppointments.length}`),
    );

    // 통계 섹션
    const stats = React.createElement(View, { style: s.stats },
      React.createElement(View, { style: s.statItem },
        React.createElement(Text, { style: s.statValue }, String(showCount)),
        React.createElement(Text, { style: s.statLabel }, 'Show'),
      ),
      React.createElement(View, { style: s.statItem },
        React.createElement(Text, { style: s.statValue }, String(noShowCount)),
        React.createElement(Text, { style: s.statLabel }, 'No Show'),
      ),
      React.createElement(View, { style: s.statItem },
        React.createElement(Text, { style: s.statValue }, `${showRate}%`),
        React.createElement(Text, { style: s.statLabel }, 'Show Rate'),
      ),
    );

    // 테이블 헤더
    const tableHeader = React.createElement(View, { style: s.tableHeader },
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'No.')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Name')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Office')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Appt. Date')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Status')),
    );

    // 테이블 데이터 행
    const tableRows = safeAppointments.map((apt: any, index: number) => {
      const safeName = safeStr(apt.name, 100);
      const safeOffice = safeStr(apt.office, 50);
      const safeApptDate = safeStr(apt.appt_date, 20);
      const safeShowStatus = apt.showStatus === 'show' ? 'Show' : 'No Show';

      return React.createElement(View, { key: index, style: s.tableRow },
        React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, String(index + 1))),
        React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeName || '-')),
        React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeOffice || '-')),
        React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeApptDate || '-')),
        React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeShowStatus)),
      );
    });

    // 테이블
    const table = safeAppointments.length > 0
      ? React.createElement(View, { style: s.table }, tableHeader, ...tableRows)
      : React.createElement(View, { style: { padding: 40, alignItems: 'center' } },
          React.createElement(Text, { style: { fontSize: 10, color: '#666' } }, 'No appointments found for the selected criteria.'),
        );

    // 푸터
    const footer = React.createElement(View, { style: s.footer },
      React.createElement(Text, null, `Generated: ${generatedDate}`),
    );

    return React.createElement(Document, null,
      React.createElement(Page, { size: 'A4', orientation: 'portrait', style: s.page }, header, infoSection, stats, table, footer),
    );
  }

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

        if (userData?.role !== 'Manager' && userData?.role !== 'User') {
          alert('You do not have access to this page.');
          setIsAuthorized(false);
          // 다른 페이지로 리다이렉트하거나 홈으로 이동
          if (typeof window !== 'undefined') {
            window.location.href = '/';
          }
          return;
        }

        setIsAuthorized(true);
        
        // 인증 성공 후 데이터 로드
        await loadAppointments();
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
      // Rate limiting timeout 정리
      if (loadAppointmentsTimeoutRef.current) {
        clearTimeout(loadAppointmentsTimeoutRef.current);
      }
    };
  }, []);

  // 필터 변경 시 데이터 필터링
  useEffect(() => {
    filterAppointments();
  }, [appointments, startDate, endDate, selectedOffice]);

  // 제출 중 브라우저 네비게이션 방지
  useEffect(() => {
    const handleBeforeUnload = (e: any) => {
      if (pdfLoading) {
        e.preventDefault();
        e.returnValue = '';
      }
    };

    const handlePopState = (e: any) => {
      if (pdfLoading) {
        e.preventDefault();
        window.history.pushState(null, '', window.location.href);
      }
    };

    if (pdfLoading) {
      window.addEventListener('beforeunload', handleBeforeUnload);
      window.addEventListener('popstate', handlePopState);
      window.history.pushState(null, '', window.location.href);
    }

    return () => {
      window.removeEventListener('beforeunload', handleBeforeUnload);
      window.removeEventListener('popstate', handlePopState);
    };
  }, [pdfLoading]);

  // show-noshow 컬렉션에서 제출된 Patient Log 불러오기 (날짜별 문서)
  const loadAppointments = async () => {
    try {
      const now = Date.now();
      if (now - lastLoadAppointmentsCall.current < 1500) return;
      lastLoadAppointmentsCall.current = now;

      const currentUser = auth.currentUser;
      if (!currentUser) return;
      if (loading) return;

      setLoading(true);
      const showSnapshot = await getDocs(collection(db, "show-noshow"));

      const allAppointments: any[] = [];
      const apptDatesSet = new Set<string>();

      showSnapshot.forEach((docSnap) => {
        const data = docSnap.data();
        const appt_date = typeof data.appt_date === 'string' ? data.appt_date.trim() : '';
        if (!appt_date) return;

        apptDatesSet.add(appt_date);

        // 새 구조: document에 patients만 있음
        if (Array.isArray(data.patients)) {
          data.patients.forEach((p: any, index: number) => {
            const name = typeof p.name === 'string' ? p.name.replace(/[\x00-\x1F\x7F-\x9F]/g, '').slice(0, 100) : '';
            const office = typeof p.office === 'string' ? p.office.replace(/[\x00-\x1F\x7F-\x9F]/g, '').slice(0, 50) : '';
            const showStatus = (p.showStatus === 'show' || p.showStatus === 'no-show' || p.showStatus === 'pending') ? p.showStatus : 'pending';
            allAppointments.push({
              name,
              office,
              appt_date,
              showStatus,
              docId: sanitizeFirebaseDocIdClient(docSnap.id),
              rowIndex: index,
              dutyDate: '',
              userName: '',
              workOffice: '',
            });
          });
          return;
        }
        // 구 구조: submissions 아래 patients
        const submissions = Array.isArray(data.submissions) ? data.submissions : [];
        submissions.forEach((sub: any, subIdx: number) => {
          const patients = Array.isArray(sub.patients) ? sub.patients : [];
          patients.forEach((p: any, index: number) => {
            const name = typeof p.name === 'string' ? p.name.replace(/[\x00-\x1F\x7F-\x9F]/g, '').slice(0, 100) : '';
            const office = typeof p.office === 'string' ? p.office.replace(/[\x00-\x1F\x7F-\x9F]/g, '').slice(0, 50) : '';
            const showStatus = (p.showStatus === 'show' || p.showStatus === 'no-show' || p.showStatus === 'pending') ? p.showStatus : 'pending';
            allAppointments.push({
              name,
              office,
              appt_date,
              showStatus,
              docId: sanitizeFirebaseDocIdClient(docSnap.id),
              submissionIndex: subIdx,
              rowIndex: index,
              dutyDate: typeof sub.duty_date === 'string' ? sub.duty_date.slice(0, 50) : '',
              userName: typeof sub.name === 'string' ? sub.name.replace(/[\x00-\x1F\x7F-\x9F]/g, '').slice(0, 100) : '',
              workOffice: typeof sub.office === 'string' ? sub.office.slice(0, 50) : '',
            });
          });
        });
      });

      setAppointments(allAppointments);
      appointmentsRef.current = allAppointments;

      // 필터 날짜: document에 있는 appt_date 기준으로 기본값 설정
      if (apptDatesSet.size > 0) {
        const sorted = Array.from(apptDatesSet).sort();
        setStartDate(sorted[0]);
        setEndDate(sorted[sorted.length - 1]);
      }
    } catch (error) {
      // 에러 시 조용히 처리
    } finally {
      setLoading(false);
    }
  };

  // 필터링 함수
  const filterAppointments = () => {
    let filtered = appointments;
    
    // 날짜 범위 필터
    if (startDate && endDate) {
      filtered = filtered.filter(apt => {
        const aptDate = new Date(apt.appt_date);
        const start = new Date(startDate);
        const end = new Date(endDate);
        return aptDate >= start && aptDate <= end;
      });
    }
    
    // 오피스 필터
    if (selectedOffice && selectedOffice !== 'All') {
      filtered = filtered.filter(apt => apt.office === selectedOffice);
    }
    
    setFilteredAppointments(filtered);
  };

  // Show/No Show 상태 업데이트 (Rate limiting 적용)
  const updateShowStatus = async (appointment: any, newStatus: string) => {
    try {
      // Rate limiting: 최근 800ms 내 동일한 appointment에 대한 호출 방지
      // (같은 appointment를 실수로 빠르게 여러 번 클릭하는 것 방지)
      const now = Date.now();
      const appointmentKey = `${appointment.docId}-${appointment.submissionIndex ?? 0}-${appointment.rowIndex}`;
      const lastCallKey = `lastUpdate_${appointmentKey}`;
      
      // 로컬 스토리지에 마지막 호출 시간 저장 (브라우저 재시작 시 초기화됨)
      const lastCall = (window as any)[lastCallKey] || 0;
      if (now - lastCall < 800) {
        return; // 800ms 내 중복 호출 방지
      }
      (window as any)[lastCallKey] = now;

      // 전역 rate limiting: 모든 업데이트에 대해 300ms 제한
      // (여러 appointment를 빠르게 업데이트할 수 있도록 허용하되, 과도한 호출 방지)
      if (now - lastUpdateStatusCall.current < 300) {
        return;
      }
      lastUpdateStatusCall.current = now;

      // 상태 값 whitelist 검증 (show, no-show, pending만 허용)
      const validStatuses = ['show', 'no-show', 'pending'];
      if (!validStatuses.includes(newStatus)) {
        alert('⚠️ Invalid value.');
        return;
      }
      
      // 인증 상태 확인
      const currentUser = auth.currentUser;
      
      if (!currentUser) {
        alert('⚠️ Please log in.');
        return;
      }

      // show-noshow 문서: 새 구조는 patients[rowIndex], 구 구조는 submissions[submissionIndex].patients[rowIndex]
      const safeDocId = sanitizeFirebaseDocIdClient(appointment.docId);
      const docRef = doc(db, "show-noshow", safeDocId);
      const docSnap = await getDoc(docRef);
      const currentData = docSnap.exists() ? docSnap.data() : null;
      const subIdx = appointment.submissionIndex ?? 0;

      if (currentData && Array.isArray(currentData.patients) && currentData.submissions === undefined) {
        // 새 구조: document에 patients만 있음
        const patients = [...currentData.patients];
        if (patients[appointment.rowIndex]) {
          patients[appointment.rowIndex] = { ...patients[appointment.rowIndex], showStatus: newStatus };
          await updateDoc(docRef, sanitizeFirebaseDataClient({
            patients,
            lastUpdated: new Date().toISOString(),
            updatedAt: new Date().toISOString()
          }));
        }
      } else if (currentData && Array.isArray(currentData.submissions) && currentData.submissions[subIdx]) {
        // 구 구조: submissions
        const submissions = [...currentData.submissions];
        const patients = [...(submissions[subIdx].patients || [])];
        if (patients[appointment.rowIndex]) {
          patients[appointment.rowIndex] = { ...patients[appointment.rowIndex], showStatus: newStatus };
          submissions[subIdx] = { ...submissions[subIdx], patients };
          await updateDoc(docRef, sanitizeFirebaseDataClient({
            submissions,
            lastUpdated: new Date().toISOString(),
            updatedAt: new Date().toISOString()
          }));
        }
      }

      if (currentData && (
        (Array.isArray(currentData.patients) && currentData.submissions === undefined) ||
        (Array.isArray(currentData.submissions) && currentData.submissions[subIdx])
      )) {
        // 로컬 상태 업데이트
        setAppointments(prev => {
          const updated = prev.map((apt: any) =>
            apt.docId === appointment.docId && apt.rowIndex === appointment.rowIndex &&
            (apt.submissionIndex == null ? appointment.submissionIndex == null : apt.submissionIndex === subIdx)
              ? { ...apt, showStatus: newStatus }
              : apt
          );
          
          // appointmentsRef도 업데이트 (최신 값 보장)
          appointmentsRef.current = updated;
          
          // filteredAppointments도 즉시 업데이트 (useEffect 대기 없이)
          // filterAppointments 함수 로직을 재사용
          let filtered = updated;
          
          // 날짜 범위 필터
          if (startDate && endDate) {
            filtered = filtered.filter(apt => {
              const aptDate = new Date(apt.appt_date);
              const start = new Date(startDate);
              const end = new Date(endDate);
              return aptDate >= start && aptDate <= end;
            });
          }
          
          // 오피스 필터
          if (selectedOffice && selectedOffice !== 'All') {
            filtered = filtered.filter(apt => apt.office === selectedOffice);
          }
          
          // 즉시 filteredAppointments 업데이트
          setFilteredAppointments(filtered);
          
          return updated;
        });
        
        // 로그 제거 (보안 강화)
      }
    } catch (error) {
      // 로그 제거 (보안 강화)
    }
  };

  // PDF 생성 및 제출 (보안 강화 + Rate limiting)
  const handleGeneratePDF = async () => {
    // Rate limiting: 최근 3초 내 호출 방지 (PDF 생성은 무거운 작업)
    // (실수로 두 번 클릭하는 것을 방지하되, 사용자 경험을 해치지 않도록)
    const now = Date.now();
    if (now - lastGeneratePDFCall.current < 3000) {
      alert('⚠️ Please try again.');
      return;
    }
    lastGeneratePDFCall.current = now;

    // 이미 PDF 생성 중이면 중복 호출 방지
    if (pdfLoading) {
      return;
    }

    if (filteredAppointments.length === 0) {
      alert('⚠️ No appointments to generate PDF.');
      return;
    }

    // pending 상태가 있는지 체크 (actions 또는 showStatus)
    const hasPendingStatus = filteredAppointments.some(apt => 
      apt.actions === 'pending' || 
      apt.showStatus === 'pending' || 
      (!apt.actions && !apt.showStatus) ||
      (apt.actions !== 'show' && apt.actions !== 'no-show' && apt.showStatus !== 'show' && apt.showStatus !== 'no-show')
    );
    
    if (hasPendingStatus) {
      alert('⚠️ Please select Show or No Show for all appointments before submitting.');
      return;
    }

    if (!name || !name.trim()) {
      alert('⚠️ Please enter your Name.');
      return;
    }

    // 인증 상태 확인
    const currentUser = auth.currentUser;
    
    if (!currentUser) {
      alert('⚠️ Please log in again.');
      return;
    }
    
    // 토큰 미리 갱신 (서버 요청 전에 토큰이 최신인지 확인)
    try {
      await currentUser.getIdToken(true);
    } catch (tokenError) {
      // 토큰 갱신 실패는 무시 (로그 제거)
    }

    try {
      setPdfLoading(true);
      setSubmitStatus('Saving...');
      setProgress(10);

      // PDF용 데이터 준비
      setSubmitStatus('Processing...');
      setProgress(30);
      
      // 최신 appointments state를 기반으로 필터링 (클로저 문제 방지)
      // appointmentsRef.current를 사용하여 최신 상태 보장
      const latestAppointments = appointmentsRef.current.length > 0 ? appointmentsRef.current : appointments;
      
      // 날짜 범위 필터
      let pdfFilteredAppointments = latestAppointments;
      if (startDate && endDate) {
        pdfFilteredAppointments = pdfFilteredAppointments.filter(apt => {
          const aptDate = new Date(apt.appt_date);
          const start = new Date(startDate);
          const end = new Date(endDate);
          return aptDate >= start && aptDate <= end;
        });
      }
      
      // 오피스 필터
      if (selectedOffice && selectedOffice !== 'All') {
        pdfFilteredAppointments = pdfFilteredAppointments.filter(apt => apt.office === selectedOffice);
      }
      
      // pending 상태의 appointment는 PDF에서 제외
      pdfFilteredAppointments = pdfFilteredAppointments.filter(apt => 
        apt.showStatus === 'show' || apt.showStatus === 'no-show'
      );
      
      // PDF 생성에 사용할 데이터 저장 (삭제 시 사용)
      const appointmentsForPdf = pdfFilteredAppointments;
      
      // PDF 생성 전에 데이터가 있는지 확인
      if (appointmentsForPdf.length === 0) {
        alert('⚠️ Please select show or no show.');
        setPdfLoading(false);
        setSubmitStatus('');
        setProgress(0);
        return;
      }
      
      // 통계 계산
      const showCount = appointmentsForPdf.filter(apt => apt.showStatus === 'show').length;
      const noShowCount = appointmentsForPdf.filter(apt => apt.showStatus === 'no-show').length;
      const showRate = appointmentsForPdf.length > 0 ? ((showCount / (showCount + noShowCount)) * 100).toFixed(1) : '0';

      // 생성 날짜 포맷팅
      const currentDate = new Date();
      const formatter = new Intl.DateTimeFormat('en-US', {
        timeZone: 'America/Los_Angeles',
        year: 'numeric',
        month: 'long',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        hour12: true
      });
      const generatedDate = formatter.format(currentDate);

      // 데이터 검증 및 sanitization
      const safeStartDate = startDate.replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 20);
      const safeEndDate = endDate.replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 20);
      const safeSelectedOffice = (selectedOffice || 'All').replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 100);
      const safeGeneratedBy = (name || 'Supervisor').replace(/[\x00-\x1F\x7F-\x9F]/g, '').slice(0, 100);

      // PDF 생성 (client-side)
      setSubmitStatus('Processing...');
      setProgress(60);
      
      try {
        const pdfBlob = await pdf(createShowCheckPDFDocument({
          safeStartDate: startDate,
          safeEndDate: endDate,
          safeSelectedOffice: selectedOffice || 'All',
          safeGeneratedBy,
          safeAppointments: appointmentsForPdf,
          showCount,
          noShowCount,
          showRate,
          generatedDate,
        })).toBlob();

        if (!pdfBlob || pdfBlob.size === 0) {
          throw new Error('error');
        }
        
        // 파일명 생성 (강화된 검증)
        const safeDateRange = startDate === endDate 
          ? startDate.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 50)
          : `${startDate.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 50)}_to_${endDate.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 50)}`;
        const safeOfficeName = (selectedOffice || 'All_Offices').replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 50);
        const tsNow = new Date();
        const tsLaTime = new Date(tsNow.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
        let tsHours = tsLaTime.getHours();
        const tsMinutes = tsLaTime.getMinutes();
        const tsAmpm = tsHours >= 12 ? 'pm' : 'am';
        tsHours = tsHours % 12;
        tsHours = tsHours ? tsHours : 12;
        const timeStamp = `${tsHours}${tsMinutes.toString().padStart(2, '0')}${tsAmpm}`;
        const filename = `${safeDateRange}_${safeOfficeName}_Show Check_${timeStamp}.pdf`.slice(0, 255);
        // 제출 날짜(현재 날짜)를 폴더 이름으로 사용
        const submissionDate = new Date().toISOString().split('T')[0]; // YYYY-MM-DD 형식
        const date = submissionDate.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 50); // 저장 경로에 사용할 날짜 (제출 날짜 사용)
        const office = 'Appointment Show'; // 저장 경로의 office는 "Appointment Show"로 고정
        
        // Firebase에 자동 저장
        setSubmitStatus('Saving...');
        setProgress(70);
        
        try {
          // PDF를 Firebase Storage에 저장
          const storage = getStorage();
          const storageRef = ref(storage, `endofday-pdfs/${office}/${date}/${filename}`);
          
          // PDF 업로드 (endofday-pdfs Storage에만 저장)
          await uploadBytes(storageRef, pdfBlob);
          
          setSubmitStatus('Saved Successfully!');
          setProgress(80);
          
          // 데이터 삭제
          setSubmitStatus('Cleaning up...');
          setProgress(90);
          
          try {
            await deleteProcessedAppointments(appointmentsForPdf);
          } catch (deleteError) {
            // 로그 제거 (보안 강화)
          }
          
          setSubmitStatus('Complete!');
          setProgress(100);
          
          setTimeout(() => {
            setPdfLoading(false);
            setSubmitStatus('');
            setProgress(0);
          }, 2000);
        } catch (storageError: any) {
          setPdfLoading(false);
          setSubmitStatus('');
          setProgress(0);
        }
      } catch (pdfError: any) {
        throw new Error('error');
      }

    } catch (error) {
      // 로그 제거 (보안 강화)
      setSubmitStatus('error');
      setProgress(0);
      setTimeout(() => {
        setPdfLoading(false);
        setSubmitStatus('');
      }, 3000);
    }
  };

  // 처리된 약속 데이터 삭제: 새 구조는 patients에서 제거, 구 구조는 submissions[].patients에서 제거
  const deleteProcessedAppointments = async (appointmentsToDelete?: any[]) => {
    try {
      const appointmentsToProcess = appointmentsToDelete || filteredAppointments;
      const byDocFlat = new Map<string, Set<number>>(); // 새 구조: docId -> rowIndices
      const byDocSubs = new Map<string, Map<number, Set<number>>>(); // 구 구조: docId -> subIdx -> rowIndices
      appointmentsToProcess.forEach(apt => {
        const docId = apt.docId;
        if (apt.submissionIndex == null) {
          if (!byDocFlat.has(docId)) byDocFlat.set(docId, new Set());
          byDocFlat.get(docId)!.add(apt.rowIndex);
        } else {
          if (!byDocSubs.has(docId)) byDocSubs.set(docId, new Map());
          const subMap = byDocSubs.get(docId)!;
          const subIdx = apt.submissionIndex;
          if (!subMap.has(subIdx)) subMap.set(subIdx, new Set());
          subMap.get(subIdx)!.add(apt.rowIndex);
        }
      });

      for (const [docId, rowIndices] of byDocFlat) {
        const safeDocId = sanitizeFirebaseDocIdClient(docId);
        const docRef = doc(db, "show-noshow", safeDocId);
        const docSnap = await getDoc(docRef);
        const data = docSnap.exists() ? docSnap.data() : null;
        if (!data || !Array.isArray(data.patients) || data.submissions !== undefined) continue;
        const patients = data.patients.filter((_: any, i: number) => !rowIndices.has(i));
        if (patients.length === 0) await deleteDoc(docRef);
        else await updateDoc(docRef, sanitizeFirebaseDataClient({ patients, lastUpdated: new Date().toISOString() }));
      }
      for (const [docId, subMap] of byDocSubs) {
        const safeDocId = sanitizeFirebaseDocIdClient(docId);
        const docRef = doc(db, "show-noshow", safeDocId);
        const docSnap = await getDoc(docRef);
        const currentData = docSnap.exists() ? docSnap.data() : null;
        if (!currentData || !Array.isArray(currentData.submissions)) continue;
        const submissions = currentData.submissions as any[];
        const updatedSubmissions = submissions
          .map((sub: any, subIdx: number) => {
            const toRemove = subMap.get(subIdx);
            if (!toRemove || !Array.isArray(sub.patients)) return sub;
            const patients = sub.patients.filter((_: any, i: number) => !toRemove.has(i));
            return { ...sub, patients };
          })
          .filter((sub: any) => Array.isArray(sub.patients) && sub.patients.length > 0);
        if (updatedSubmissions.length === 0) await deleteDoc(docRef);
        else await updateDoc(docRef, sanitizeFirebaseDataClient({ submissions: updatedSubmissions, lastUpdated: new Date().toISOString() }));
      }

      const toRemove = new Set(appointmentsToProcess.map(p =>
        p.submissionIndex == null ? `${p.docId}-${p.rowIndex}` : `${p.docId}-${p.submissionIndex}-${p.rowIndex}`
      ));
      const updated = appointmentsRef.current.filter(apt =>
        !toRemove.has(apt.submissionIndex == null ? `${apt.docId}-${apt.rowIndex}` : `${apt.docId}-${apt.submissionIndex}-${apt.rowIndex}`)
      );
      setAppointments(updated);
      appointmentsRef.current = updated;
      setFilteredAppointments(prev =>
        prev.filter(apt => !toRemove.has(apt.submissionIndex == null ? `${apt.docId}-${apt.rowIndex}` : `${apt.docId}-${apt.submissionIndex}-${apt.rowIndex}`))
      );
    } catch (error) {
      throw error;
    }
  };

  // 스타일 정의
  const containerStyle = {
    maxWidth: '1400px',
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
    padding: '8px 16px',
    borderRadius: '6px',
    fontSize: '14px',
    fontWeight: 'bold',
    cursor: 'pointer',
    margin: '2px',
    border: 'none',
    transition: 'all 0.3s ease'
  };

  const showButtonStyle = {
    ...buttonStyle,
    backgroundColor: '#28a745',
    color: 'white'
  };

  const noShowButtonStyle = {
    ...buttonStyle,
    backgroundColor: '#dc3545',
    color: 'white'
  };

  const pendingButtonStyle = {
    ...buttonStyle,
    backgroundColor: '#ffc107',
    color: '#212529'
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

  const getStatusColor = (status: string, index: number): string => {
    switch (status) {
      case 'show': return '#d4edda'; // 초록색 배경
      case 'no-show': return '#f8d7da'; // 빨간색 배경
      case 'pending': return '#fff3cd'; // 노란색 배경
      default: return index % 2 === 0 ? '#f9f9f9' : 'white'; // 기본 교대로 배경색
    }
  };

  const getStatusText = (status: string): string => {
    switch (status) {
      case 'show': return 'Show ✅';
      case 'no-show': return 'No Show ❌';
      default: return 'Pending ⏳';
    }
  };

  // 인증 확인 중이거나 인증 실패 시 로딩 화면 표시
  if (isAuthorized === null) {
    return (
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
    );
  }

  if (isAuthorized === false) {
    return (
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
        {pdfLoading && (
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
                {submitStatus || "Processing..."}
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
        <h1 style={{ 
          color: '#0077B6', 
          textAlign: 'center', 
          marginBottom: '20px', 
          fontSize: '2.5rem', 
          fontWeight: 'bold',
          textShadow: '2px 2px 4px rgba(0,0,0,0.1)'
        }}>Appointment Show/No Show Check</h1>

        {/* Name 입력 섹션 */}
        <div style={sectionStyle}>
          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap', marginBottom: '20px' }}>
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Name:
              </label>
              <input
                type="text"
                value={name}
                onChange={(e) => {
                  const validatedValue = validateInput('name', e.target.value);
                  setName(validatedValue);
                }}
                placeholder="Enter your name"
                style={inputStyle}
              />
            </div>
          </div>
        </div>

        {/* 필터 섹션 */}
        <div style={sectionStyle}>
          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap', marginBottom: '20px' }}>
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Start Date:
              </label>
              <input
                type="date"
                value={startDate}
                onChange={(e) => {
                  const validatedValue = validateInput('startDate', e.target.value);
                  setStartDate(validatedValue);
                }}
                style={inputStyle}
              />
            </div>
            
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                End Date:
              </label>
              <input
                type="date"
                value={endDate}
                onChange={(e) => {
                  const validatedValue = validateInput('endDate', e.target.value);
                  setEndDate(validatedValue);
                }}
                style={inputStyle}
              />
            </div>

            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Office:
              </label>
              <select
                value={selectedOffice}
                onChange={(e) => {
                  const validatedValue = validateInput('selectedOffice', e.target.value);
                  setSelectedOffice(validatedValue);
                }}
                style={inputStyle}
              >
                {officeOptions.map((office: string) => (
                  <option key={office} value={office}>{office}</option>
                ))}
              </select>
            </div>
            

            <div style={{ flex: '1', minWidth: '200px', display: 'flex', alignItems: 'end' }}>
              <button 
                onClick={loadAppointments}
                disabled={loading}
                style={{
                  ...buttonStyle,
                  backgroundColor: '#0077B6',
                  color: 'white',
                  width: '100%'
                }}
              >
                {loading ? 'Loading...' : '🔄 Refresh Data'}
              </button>
            </div>
          </div>

          <div style={{ 
            padding: '10px', 
            backgroundColor: '#e3f2fd',
            borderRadius: '5px', 
            fontSize: '14px',
            color: '#1565c0'
          }}>
            📊 <strong>Total Appointments:</strong> {filteredAppointments.length}
            {startDate && endDate && ` | Date Range: ${startDate} to ${endDate}`}
            {selectedOffice && selectedOffice !== 'All' && ` | Office: ${selectedOffice}`}
          </div>
            </div>

        {/* 약속 목록 테이블 */}
        <div style={sectionStyle}>
          {loading ? (
            <div style={{ textAlign: 'center', padding: '40px', color: '#666' }}>
              Loading appointments...
            </div>
          ) : filteredAppointments.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '40px', color: '#666' }}>
              No appointments found for the selected criteria.
            </div>
          ) : (
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', margin: '10px 0' }}>
              <thead style={{ backgroundColor: '#0077B6', color: 'white' }}>
                <tr>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Patient Name</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Office</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Appt. Date</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Status</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                  {filteredAppointments.map((appointment: any, index: number) => (
                    <tr 
                      key={`${appointment.docId}-${appointment.submissionIndex ?? 0}-${appointment.rowIndex}`} 
                      style={{ 
                        backgroundColor: getStatusColor(appointment.showStatus, index),
                        opacity: appointment.showStatus === 'pending' ? 1 : 0.8
                      }}
                    >
                      <td style={{ padding: '12px 8px', fontWeight: 'bold' }}>
                        {appointment.name}
                    </td>
                      <td style={{ padding: '12px 8px', textAlign: 'center' }}>
                        {appointment.office}
                    </td>
                      <td style={{ padding: '12px 8px', textAlign: 'center' }}>
                        {appointment.appt_date}
                    </td>
                      <td style={{ padding: '12px 8px', textAlign: 'center', fontWeight: 'bold' }}>
                        {getStatusText(appointment.showStatus)}
                    </td>
                      <td style={{ padding: '12px 8px', textAlign: 'center' }}>
                        <div style={{ display: 'flex', gap: '4px', justifyContent: 'center' }}>
                          <button
                            onClick={() => updateShowStatus(appointment, 'show')}
                            style={showButtonStyle}
                            disabled={appointment.showStatus === 'show'}
                          >
                            Show
                          </button>
                          <button
                            onClick={() => updateShowStatus(appointment, 'no-show')}
                            style={noShowButtonStyle}
                            disabled={appointment.showStatus === 'no-show'}
                          >
                            No Show
                          </button>
                        <button
                            onClick={() => updateShowStatus(appointment, 'pending')}
                            style={pendingButtonStyle}
                            disabled={appointment.showStatus === 'pending'}
                          >
                            Pending
                        </button>
                        </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          )}
        </div>

        {/* 통계 섹션 */}
        {filteredAppointments.length > 0 && (
        <div style={sectionStyle}>
            <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap' }}>
              {(() => {
                const showCount = filteredAppointments.filter(apt => apt.showStatus === 'show').length;
                const noShowCount = filteredAppointments.filter(apt => apt.showStatus === 'no-show').length;
                const pendingCount = filteredAppointments.filter(apt => apt.showStatus === 'pending').length;
                const showRate = filteredAppointments.length > 0 ? ((showCount / (showCount + noShowCount)) * 100).toFixed(1) : 0;
                
                return (
                  <>
                    <div style={{ 
                      flex: '1', 
                      minWidth: '150px', 
                      padding: '15px', 
                      backgroundColor: '#d4edda', 
                      borderRadius: '8px',
                      textAlign: 'center'
                    }}>
                      <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#155724' }}>{showCount}</div>
                      <div style={{ color: '#155724' }}>Show</div>
                    </div>
                    <div style={{ 
                      flex: '1', 
                      minWidth: '150px', 
                      padding: '15px', 
                      backgroundColor: '#f8d7da', 
                      borderRadius: '8px',
                      textAlign: 'center'
                    }}>
                      <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#721c24' }}>{noShowCount}</div>
                      <div style={{ color: '#721c24' }}>No Show</div>
                    </div>
                    <div style={{ 
                      flex: '1', 
                      minWidth: '150px', 
                      padding: '15px', 
                      backgroundColor: '#fff3cd', 
                      borderRadius: '8px',
                      textAlign: 'center'
                    }}>
                      <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#856404' }}>{pendingCount}</div>
                      <div style={{ color: '#856404' }}>Pending</div>
                    </div>
                    <div style={{ 
                      flex: '1', 
                      minWidth: '150px', 
                      padding: '15px', 
                      backgroundColor: '#e3f2fd', 
                      borderRadius: '8px',
                      textAlign: 'center'
                    }}>
                      <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#1565c0' }}>{showRate}%</div>
                      <div style={{ color: '#1565c0' }}>Show Rate</div>
                    </div>
                  </>
                );
              })()}
            </div>
        </div>
        )}

        {/* Submit 버튼 */}
        {filteredAppointments.length > 0 && (
          <div style={{textAlign: 'center', margin: '30px 0'}}>
          {(() => {
            const hasPendingStatus = filteredAppointments.some(apt => 
              apt.actions === 'pending' || 
              apt.showStatus === 'pending' || 
              (!apt.actions && !apt.showStatus) ||
              (apt.actions !== 'show' && apt.actions !== 'no-show' && apt.showStatus !== 'show' && apt.showStatus !== 'no-show')
            );
            const nameEmpty = !name || !name.trim();
            const canSubmit = !pdfLoading && !hasPendingStatus && !nameEmpty;
            const titleMsg = nameEmpty
              ? '⚠️ Please enter your Name.'
              : hasPendingStatus
                ? '⚠️ Please select Show or No Show for all appointments before submitting.'
                : '';
            
            return (
              <button 
                onClick={handleGeneratePDF}
                disabled={!canSubmit}
                style={{
                  backgroundColor: canSubmit ? '#3b82f6' : '#9ca3af',
                  color: 'white',
                  border: 'none',
                  padding: '15px 30px',
                  borderRadius: '8px',
                  fontSize: '16px',
                  fontWeight: 'bold',
                  cursor: canSubmit ? 'pointer' : 'not-allowed',
                  boxShadow: '0 4px 6px rgba(0, 0, 0, 0.1)',
                  transition: 'all 0.2s ease'
                }}
                title={titleMsg}
              >
                {pdfLoading ? '📄 Submitting...' : 'Submit'}
              </button>
            );
          })()}
        </div>
        )}
      </div>
    </div>
    </>
  );
}


