'use client'

import React, { useState, useEffect, useRef } from "react";
import { doc, setDoc, collection, getDocs, getDoc, updateDoc, deleteDoc, query, where } from "firebase/firestore";
import { db, auth } from "@/lib/firebase.config";
import { onAuthStateChanged } from 'firebase/auth';
// Firebase Storage
import { getStorage, ref, uploadBytes, getDownloadURL } from 'firebase/storage';
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
        const validOffices = ['A', 'B', 'C', 'D'];
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
  const officeOptions = ['All', 'A', 'B', 'C', 'D'];

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
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Time')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Visit Type')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Status')),
    );

    // 테이블 데이터 행
    const tableRows = safeAppointments.map((apt: any, index: number) => {
      const safeName = safeStr(apt.name, 100);
      const safeOffice = safeStr(apt.office, 50);
      const safeApptDate = safeStr(apt.appt_date, 20);
      const safeTime = safeStr(apt.time, 20);
      const safeVisitType = safeStr(apt.visit_type, 50);
      const safeShowStatus = apt.showStatus === 'show' ? 'Show' : 'No Show';

      return React.createElement(View, { key: index, style: s.tableRow },
        React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, String(index + 1))),
        React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeName || '-')),
        React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeOffice || '-')),
        React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeApptDate || '-')),
        React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, convertTo12Hour(safeTime))),
        React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeVisitType || '-')),
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

        if (userData?.role !== 'admin') {
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
        alert('error');
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

  // Firebase에서 모든 환자 로그 불러오기 (Rate limiting 적용)
  const loadAppointments = async () => {
    try {
      // Rate limiting: 최근 1.5초 내 호출 방지
      // (필터 변경 후 빠르게 새로고침할 수 있도록 허용하되, 과도한 호출 방지)
      const now = Date.now();
      if (now - lastLoadAppointmentsCall.current < 1500) {
        return;
      }
      lastLoadAppointmentsCall.current = now;

      // 인증 상태 확인
      const currentUser = auth.currentUser;
      
      if (!currentUser) {
        // Firebase Security Rules가 이미 접근을 제어하므로 조용히 실패
        return;
      }

      // 이미 로딩 중이면 중복 호출 방지
      if (loading) {
        return;
      }

      setLoading(true);
      const querySnapshot = await getDocs(collection(db, "patient-logs"));
      
      // 제출된 PDF 목록 가져오기 (제출된 것만 필터링하기 위해)
      const pdfQuerySnapshot = await getDocs(collection(db, "pdf-documents"));
      const submittedDocs = new Set<string>();
      
      pdfQuerySnapshot.forEach((pdfDoc) => {
        const pdfData = pdfDoc.data();
        if (pdfData.type === 'Patient Log' && pdfData.date && pdfData.name && pdfData.office) {
          // 제출된 문서의 키 생성: date_name_office
          const key = `${pdfData.date}_${pdfData.name}_${pdfData.office}`;
          submittedDocs.add(key);
        }
      });
      
      const allAppointments: any[] = [];
      
      querySnapshot.forEach((doc) => {
        const data = doc.data();
        
        // 보안 강화: 데이터 소유권 확인 (Firebase Security Rules와 함께)
        if (currentUser && data.userId && data.userId !== currentUser.uid) {
          // 다른 사용자의 데이터는 로드하지 않음 (Security Rules에서도 차단됨)
          return;
        }
        
        // 제출된 데이터만 필터링
        if (data.dutyDate && data.userName && data.workOffice) {
          const docKey = `${data.dutyDate}_${data.userName}_${data.workOffice}`;
          if (!submittedDocs.has(docKey)) {
            // 제출되지 않은 데이터는 건너뛰기
            return;
          }
        }
        
        if (data.patientRows) {
          data.patientRows.forEach((row: any, index: number) => {
            if (row.appt_date && row.name) {
              // 기본적인 데이터 sanitization (XSS 방지)
              const sanitizedRow = {
                name: typeof row.name === 'string' ? row.name.replace(/[\x00-\x1F\x7F-\x9F]/g, '').slice(0, 100) : '',
                office: typeof row.office === 'string' ? row.office.replace(/[\x00-\x1F\x7F-\x9F]/g, '').slice(0, 50) : '',
                appt_date: typeof row.appt_date === 'string' ? row.appt_date.slice(0, 20) : '',
                visit_type: typeof row.visit_type === 'string' ? row.visit_type.replace(/[\x00-\x1F\x7F-\x9F]/g, '').slice(0, 50) : '',
                time: typeof row.time === 'string' ? row.time.slice(0, 20) : '',
                remark: typeof row.remark === 'string' ? row.remark.replace(/[\x00-\x1F\x7F-\x9F]/g, '').slice(0, 200) : '',
                other_duty: typeof row.other_duty === 'string' ? row.other_duty.replace(/[\x00-\x1F\x7F-\x9F]/g, '').slice(0, 200) : '',
                call_in: Boolean(row.call_in),
                call_out: Boolean(row.call_out),
                showStatus: (row.showStatus === 'show' || row.showStatus === 'no-show' || row.showStatus === 'pending') ? row.showStatus : 'pending'
              };
              
              allAppointments.push({
                ...sanitizedRow,
                docId: sanitizeFirebaseDocIdClient(doc.id),
                rowIndex: index,
                dutyDate: typeof data.dutyDate === 'string' ? data.dutyDate.slice(0, 50) : '',
                userName: typeof data.userName === 'string' ? data.userName.replace(/[\x00-\x1F\x7F-\x9F]/g, '').slice(0, 100) : '',
                workOffice: typeof data.workOffice === 'string' ? data.workOffice.slice(0, 50) : '',
                workHoursFrom: typeof data.workHoursFrom === 'string' ? data.workHoursFrom.slice(0, 20) : '',
                workHoursTo: typeof data.workHoursTo === 'string' ? data.workHoursTo.slice(0, 20) : ''
              });
            }
          });
        }
      });
      
      setAppointments(allAppointments);
      appointmentsRef.current = allAppointments;
    } catch (error) {
      // 로그 제거 (보안 강화)
      alert('error');
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
      const appointmentKey = `${appointment.docId}-${appointment.rowIndex}`;
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

      // 해당 document를 다시 가져와서 patientRows 업데이트
      const safeDocId = sanitizeFirebaseDocIdClient(appointment.docId);
      const docRef = doc(db, "patient-logs", safeDocId);
      const querySnapshot = await getDocs(collection(db, "patient-logs"));
      let currentData: any = null;
      
      querySnapshot.forEach((document) => {
        if (document.id === appointment.docId) {
          currentData = document.data();
        }
      });
      
      // 보안 강화: 데이터 소유권 확인
      if (currentData && currentData.userId && currentData.userId !== currentUser.uid) {
        alert('⚠️ You do not have access to this page.');
        return;
      }
      
      if (currentData && currentData.patientRows) {
        // 해당 row의 showStatus 업데이트
        const updatedPatientRows = currentData.patientRows.map((row: any, index: number) => {
          if (index === appointment.rowIndex) {
            return { ...row, showStatus: newStatus };
          }
          return row;
        });
        
        // Firebase 업데이트 (보안 검증 적용 + 사용자 정보 강제 추가)
        const safeUpdateData = sanitizeFirebaseDataClient({
          patientRows: updatedPatientRows,
          lastUpdated: new Date().toISOString(),
          // 보안 강화: 사용자 정보 강제 추가
          userId: currentUser.uid, // 항상 현재 사용자 ID로 설정
          ...(currentUser.email && { userEmail: currentUser.email }),
          updatedAt: new Date().toISOString()
        });
        await updateDoc(docRef, safeUpdateData);
        
        // 로컬 상태 업데이트
        setAppointments(prev => {
          const updated = prev.map((apt: any) => 
            apt.docId === appointment.docId && apt.rowIndex === appointment.rowIndex
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
      alert('error');
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
        const filename = `${safeDateRange}_${safeOfficeName}_Show Check.pdf`.slice(0, 255);
        // 제출 날짜(현재 날짜)를 폴더 이름으로 사용
        const submissionDate = new Date().toISOString().split('T')[0]; // YYYY-MM-DD 형식
        const date = submissionDate.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 50); // 저장 경로에 사용할 날짜 (제출 날짜 사용)
        const office = 'Appointment Show'; // 저장 경로의 office는 "Appointment Show"로 고정
        
        // Firebase에 자동 저장
        setSubmitStatus('Saving to archive...');
        setProgress(70);
        
        try {
          // PDF를 Firebase Storage에 저장
          const storage = getStorage();
          const storageRef = ref(storage, `endofday-pdfs/${office}/${date}/${filename}`);
          
          // PDF 업로드
          await uploadBytes(storageRef, pdfBlob);
          
          // 다운로드 URL 가져오기
          const downloadUrl = await getDownloadURL(storageRef);
          
          // Firestore에 메타데이터 저장
          const safeMetadata = sanitizeFirebaseDataClient({
            filename: filename,
            office: office,
            date: date,
            name: name || 'Supervisor',
            type: 'Show Check',
            source: 's_route',
          });
          
          await setDoc(doc(db, 'pdf-documents', `${date}_${safeOfficeName}_showcheck_${Date.now()}`), {
            ...safeMetadata,
            url: downloadUrl,
            storagePath: `endofday-pdfs/${office}/${date}/${filename}`,
            createdAt: new Date(),
          });
          
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
          alert('error');
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

  // 처리된 약속 데이터 삭제 함수
  // appointmentsToDelete: PDF 생성에 사용된 appointment 배열 (선택적)
  const deleteProcessedAppointments = async (appointmentsToDelete?: any[]) => {
    try {
      // PDF 생성에 사용된 appointment가 있으면 그것을 사용, 없으면 현재 filteredAppointments 사용
      const appointmentsToProcess = appointmentsToDelete || filteredAppointments;
      
      // 현재 필터링된 약속들의 document ID별로 그룹화
      const documentsToCheck = new Map();
      
      appointmentsToProcess.forEach(appointment => {
        const docId = appointment.docId;
        if (!documentsToCheck.has(docId)) {
          documentsToCheck.set(docId, []);
        }
        documentsToCheck.get(docId).push(appointment.rowIndex);
      });

      const deletedDocuments: string[] = [];
      const updatedDocuments: string[] = [];

      // 각 document 확인 및 처리
      for (const [docId, processedRowIndices] of documentsToCheck) {
        const safeDocId = sanitizeFirebaseDocIdClient(docId);
        const docRef = doc(db, "patient-logs", safeDocId);
        
        // 현재 document 데이터 가져오기
        const querySnapshot = await getDocs(collection(db, "patient-logs"));
        let currentData: any = null;
        
        querySnapshot.forEach((document) => {
          if (document.id === docId) {
            currentData = document.data();
          }
        });

        if (currentData && currentData.patientRows) {
          // 모든 약속(appt_date가 있는 것들) 찾기
          const allAppointmentRows = currentData.patientRows
            .map((row: any, index: number) => ({ ...row, originalIndex: index }))
            .filter((row: any) => row.appt_date && row.name);

          // 처리되지 않은 약속들 찾기 (현재 필터링된 것들 제외)
          const unprocessedAppointments = allAppointmentRows.filter((row: any) => 
            !processedRowIndices.includes(row.originalIndex)
          );

          if (unprocessedAppointments.length === 0 && allAppointmentRows.length > 0) {
            // 약속이 있었고 모든 약속이 처리됨 → 전체 document 삭제
            await deleteDoc(docRef);
            deletedDocuments.push(docId);
            // 로그 제거 (보안 강화)
          } else if (allAppointmentRows.length === 0) {
            // 애초에 약속이 없는 document → 삭제하지 않음 (로그 제거)
          } else {
            // 아직 처리 안 된 약속이 있음 → 처리된 것들만 빈 상태로 초기화
            // showStatus를 pending으로 설정하지 않고 필드를 제거하거나 undefined로 설정
            const updatedPatientRows = currentData.patientRows.map((row: any, index: number) => {
              if (processedRowIndices.includes(index)) {
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
                  // showStatus 필드를 제거 (pending으로 설정하지 않음)
                };
              }
              return row;
            });

            const safeUpdateData = sanitizeFirebaseDataClient({
              patientRows: updatedPatientRows,
              lastUpdated: new Date().toISOString()
            });
            await updateDoc(docRef, safeUpdateData);
            updatedDocuments.push(docId);
            // 로그 제거 (보안 강화)
          }
        }
      }

      // 로컬 상태 업데이트 - 처리된 약속들 제거
      setAppointments(prevAppointments => {
        const updated = prevAppointments.filter(apt => !filteredAppointments.some(filtered => 
          filtered.docId === apt.docId && filtered.rowIndex === apt.rowIndex
        ));
        appointmentsRef.current = updated; // appointmentsRef도 업데이트
        return updated;
      });
      
      setFilteredAppointments([]);
      
      // 로그 제거 (보안 강화)

    } catch (error) {
      // 로그 제거 (보안 강화)
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
          <div style={{ fontSize: '18px', color: '#023047' }}>인증 확인 중...</div>
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
          <div style={{ fontSize: '18px', color: '#d32f2f', marginBottom: '10px' }}>접근 권한이 없습니다</div>
          <div style={{ fontSize: '14px', color: '#666' }}>관리자 권한이 필요합니다</div>
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
        }}>📋 Appointment Show/No Show Check</h1>

        {/* Name 입력 섹션 */}
        <div style={sectionStyle}>
          <h2 style={{ color: '#0077B6', marginBottom: '20px' }}>👤 Name</h2>
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
          <h2 style={{ color: '#0077B6', marginBottom: '20px' }}>🔍 Filter Appointments</h2>
          
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
          <h2 style={{ color: '#0077B6', marginBottom: '20px' }}>👥 Appointments to Check</h2>
          
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
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Visit Type</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Status</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                  {filteredAppointments.map((appointment: any, index: number) => (
                    <tr 
                      key={`${appointment.docId}-${appointment.rowIndex}`} 
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
                      <td style={{ padding: '12px 8px', textAlign: 'center' }}>
                        {appointment.visit_type}
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
            <h2 style={{ color: '#0077B6', marginBottom: '20px' }}>📊 Statistics</h2>
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
            // pending 상태가 있는지 체크 (actions 또는 showStatus)
            const hasPendingStatus = filteredAppointments.some(apt => 
              apt.actions === 'pending' || 
              apt.showStatus === 'pending' || 
              (!apt.actions && !apt.showStatus) ||
              (apt.actions !== 'show' && apt.actions !== 'no-show' && apt.showStatus !== 'show' && apt.showStatus !== 'no-show')
            );
            
            return (
              <button 
                onClick={handleGeneratePDF}
                disabled={pdfLoading || hasPendingStatus}
                style={{
                  backgroundColor: (pdfLoading || hasPendingStatus) ? '#9ca3af' : '#3b82f6',
                  color: 'white',
                  border: 'none',
                  padding: '15px 30px',
                  borderRadius: '8px',
                  fontSize: '16px',
                  fontWeight: 'bold',
                  cursor: (pdfLoading || hasPendingStatus) ? 'not-allowed' : 'pointer',
                  boxShadow: '0 4px 6px rgba(0, 0, 0, 0.1)',
                  transition: 'all 0.2s ease'
                }}
                title={hasPendingStatus ? '⚠️ Please select Show or No Show for all appointments before submitting.' : ''}
              >
                {pdfLoading ? '📄 Submitting...' : '📄 Submit'}
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
