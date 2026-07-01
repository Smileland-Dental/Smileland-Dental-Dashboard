'use client'

import React, { useState, useEffect, useRef, useMemo } from "react";
import { doc, collection, getDocs, getDoc, updateDoc, deleteDoc, query, where } from "firebase/firestore";
import { db } from "@/lib/firebase.config";
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
  const maxKeys = 1000;
  
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

function getCurrentDate(): string {
  return new Date().toISOString().split('T')[0];
}

function getCurrentMonth(): string {
  return getCurrentDate().slice(0, 7);
}

function formatMonthLabel(monthStr: string): string {
  return monthStr.replace('-', '/');
}

function formatShortDate(dateStr: string): string {
  return String(Number(dateStr.split('-')[2]));
}

function normalizePatientField(value: unknown, maxLen: number): string {
  return typeof value === 'string' ? value.replace(/[\x00-\x1F\x7F-\x9F]/g, '').slice(0, maxLen) : '';
}

function isSamePatient(patient: any, appointment: any): boolean {
  return (
    normalizePatientField(patient?.name, 100) === appointment.name &&
    normalizePatientField(patient?.office, 50) === appointment.office
  );
}

function isSameAppointment(a: any, b: any): boolean {
  return a.docId === b.docId && a.name === b.name && a.office === b.office;
}

function removePatientsFromList(patients: any[], toRemove: any[]): any[] {
  const remaining = [...patients];
  for (const apt of toRemove) {
    const idx = remaining.findIndex(p => isSamePatient(p, apt));
    if (idx >= 0) remaining.splice(idx, 1);
  }
  return remaining;
}

export default function ShowCheckSystem() {
  // 입력 값 검증 함수
  const validateInput = React.useCallback((field: string, value: any): any => {
    // 문자열 필드 길이 제한
    if (typeof value === 'string') {
      const maxLengths: { [key: string]: number } = {
        name: 100,
        selectedDate: 20,
        selectedMonth: 7,
        selectedOffice: 50
      };
      
      const maxLength = maxLengths[field] || 500;
      if (value.length > maxLength) {
        return value.slice(0, maxLength);
      }
      
      if (field === 'selectedMonth' && value) {
        const monthRegex = /^\d{4}-\d{2}$/;
        if (!monthRegex.test(value)) {
          return '';
        }
      }

      if (field === 'selectedDate' && value) {
        const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
        if (!dateRegex.test(value)) {
          return '';
        }
      }
      
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
  const [selectedMonth, setSelectedMonth] = useState(getCurrentMonth);
  const [selectedDate, setSelectedDate] = useState(getCurrentDate);
  const [selectedOffice, setSelectedOffice] = useState('');
  const [name, setName] = useState('');
  const [pdfLoading, setPdfLoading] = useState(false);
  const [submitStatus, setSubmitStatus] = useState('');
  const [progress, setProgress] = useState(0);

  // Office 옵션

  const officeOptions = ['All', 'Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

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

  // 컴포넌트 마운트 시 데이터 로드
  useEffect(() => {
    loadAppointments();

    // 프로덕션 환경에서 HTTPS 강제 (클라이언트 사이드)
    if (process.env.NODE_ENV === 'production' && 
        typeof window !== 'undefined' && 
        window.location.protocol !== 'https:') {
      window.location.href = window.location.href.replace('http:', 'https:');
    }
  }, []);

  // 필터 변경 시 데이터 필터링
  useEffect(() => {
    filterAppointments();
  }, [appointments, selectedDate, selectedOffice]);

  const availableMonths = useMemo(() => {
    const months = new Set<string>();
    appointments.forEach(apt => {
      if (selectedOffice && selectedOffice !== 'All' && apt.office !== selectedOffice) return;
      if (typeof apt.appt_date === 'string' && apt.appt_date.length >= 7) {
        months.add(apt.appt_date.slice(0, 7));
      }
    });
    return Array.from(months).sort();
  }, [appointments, selectedOffice]);

  const datesForSelectedMonth = useMemo(() => {
    const byDate = new Map<string, { total: number; pending: number }>();
    appointments.forEach(apt => {
      if (selectedOffice && selectedOffice !== 'All' && apt.office !== selectedOffice) return;
      if (!apt.appt_date?.startsWith(selectedMonth)) return;
      const entry = byDate.get(apt.appt_date) || { total: 0, pending: 0 };
      entry.total++;
      if (apt.showStatus === 'pending' || !apt.showStatus) entry.pending++;
      byDate.set(apt.appt_date, entry);
    });
    return Array.from(byDate.entries())
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([date, stats]) => ({ date, ...stats }));
  }, [appointments, selectedMonth, selectedOffice]);

  useEffect(() => {
    if (availableMonths.length === 0) return;
    if (!availableMonths.includes(selectedMonth)) {
      setSelectedMonth(availableMonths[availableMonths.length - 1]);
    }
  }, [availableMonths, selectedOffice]);

  // Month 변경 시 해당 달의 날짜로 맞춤 (pending 우선)
  useEffect(() => {
    if (datesForSelectedMonth.length === 0) return;
    const dates = datesForSelectedMonth.map(d => d.date);
    if (!dates.includes(selectedDate)) {
      const withPending = datesForSelectedMonth.find(d => d.pending > 0);
      setSelectedDate(withPending?.date ?? dates[dates.length - 1]);
    }
  }, [datesForSelectedMonth, selectedMonth, selectedOffice]);

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

  const loadAppointments = async () => {
    try {
      if (loading) return;

      setLoading(true);
      const showSnapshot = await getDocs(collection(db, "show-noshow"));

      const allAppointments: any[] = [];

      showSnapshot.forEach((docSnap) => {
        const data = docSnap.data();
        const appt_date = typeof data.appt_date === 'string' ? data.appt_date.trim() : '';
        if (!appt_date) return;

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
              docId: docSnap.id,
              rowIndex: index,
            });
          });
        }
      });

      setAppointments(allAppointments);
      appointmentsRef.current = allAppointments;
    } catch (error) {
      // 에러 시 조용히 처리
    } finally {
      setLoading(false);
    }
  };

  // 필터링 함수
  const filterAppointmentList = (list: any[]) => {
    let filtered = list;

    if (selectedDate) {
      filtered = filtered.filter(apt =>
        typeof apt.appt_date === 'string' && apt.appt_date === selectedDate
      );
    }

    if (selectedOffice && selectedOffice !== 'All') {
      filtered = filtered.filter(apt => apt.office === selectedOffice);
    }

    return filtered;
  };

  const filterAppointments = () => {
    setFilteredAppointments(filterAppointmentList(appointments));
  };

  const updateShowStatus = async (appointment: any, newStatus: string) => {
    try {
      // 상태 값 whitelist 검증 (show, no-show, pending 허용)
      const validStatuses = ['show', 'no-show', 'pending'];
      if (!validStatuses.includes(newStatus)) {
        alert('⚠️ Invalid value.');
        return;
      }
      
      // show-noshow 문서: patients[rowIndex]
      const docRef = doc(db, "show-noshow", appointment.docId);
      const docSnap = await getDoc(docRef);
      const currentData = docSnap.exists() ? docSnap.data() : null;

      if (currentData && Array.isArray(currentData.patients)) {
        const patients = [...currentData.patients];
        const patientIndex = patients.findIndex(p => isSamePatient(p, appointment));
        if (patientIndex >= 0) {
          patients[patientIndex] = { ...patients[patientIndex], showStatus: newStatus };
          await updateDoc(docRef, sanitizeFirebaseDataClient({
            patients,
            lastUpdated: new Date().toISOString(),
            updatedAt: new Date().toISOString()
          }));
        }
      }

      if (currentData && Array.isArray(currentData.patients) && currentData.patients.some(p => isSamePatient(p, appointment))) {
        // 로컬 상태 업데이트
        setAppointments(prev => {
          const updated = prev.map((apt: any) =>
            isSameAppointment(apt, appointment)
              ? { ...apt, showStatus: newStatus }
              : apt
          );
          
          // appointmentsRef도 업데이트 (최신 값 보장)
          appointmentsRef.current = updated;
          
          // filteredAppointments도 즉시 업데이트 (useEffect 대기 없이)
          const filtered = filterAppointmentList(updated);
          
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

  const handleGeneratePDF = async () => {
    // 이미 PDF 생성 중이면 중복 호출 방지
    if (pdfLoading) {
      return;
    }

    if (filteredAppointments.length === 0) {
      alert('⚠️ No appointments to generate PDF.');
      return;
    }

    // 미처리(pending) 상태가 있는지 체크
    const hasPendingStatus = filteredAppointments.some(apt => {
      const s = apt.showStatus ?? apt.actions;
      return s !== 'show' && s !== 'no-show';
    });
    
    if (hasPendingStatus) {
      alert('⚠️ Please select Show or No Show for all appointments before submitting.');
      return;
    }

    if (!name || !name.trim()) {
      alert('⚠️ Please enter your Name.');
      return;
    }

    try {
      setPdfLoading(true);
      setSubmitStatus('Saving...');
      setProgress(10);

      // PDF용 데이터 준비
      setSubmitStatus('Processing...');
      setProgress(30);
      
      const latestAppointments = appointmentsRef.current.length > 0 ? appointmentsRef.current : appointments;
      const pdfFilteredAppointments = filterAppointmentList(latestAppointments);

      // 제출 후 DB에서 제거할 행: show / no-show (pending은 유지)
      const appointmentsToDeleteFromDb = pdfFilteredAppointments.filter(apt =>
        apt.showStatus === 'show' || apt.showStatus === 'no-show'
      );
      
      const appointmentsForPdf = appointmentsToDeleteFromDb;
      
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
      const denom = showCount + noShowCount;
      const showRate = denom > 0 ? ((showCount / denom) * 100).toFixed(1) : '0';

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
      const safeStartDate = selectedDate.replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 20);
      const safeEndDate = selectedDate.replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 20);
      const safeSelectedOffice = (selectedOffice || 'All').replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 100);
      const safeGeneratedBy = (name || 'Supervisor').replace(/[\x00-\x1F\x7F-\x9F]/g, '').slice(0, 100);

      // PDF 생성 (client-side)
      setSubmitStatus('Processing...');
      setProgress(60);
      
      try {
        const pdfBlob = await pdf(createShowCheckPDFDocument({
          safeStartDate: selectedDate,
          safeEndDate: selectedDate,
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
        const safeDateRange = selectedDate.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 50);
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
        const office = 'Appointment Show';
        
        // Firebase에 자동 저장
        setSubmitStatus('Saving...');
        setProgress(70);
        
        try {
          const storage = getStorage();
          const storageRef = ref(storage, `endofday-pdfs/${office}/${date}/${filename}`);
          
          await uploadBytes(storageRef, pdfBlob);
          
          setSubmitStatus('Saved Successfully!');
          setProgress(80);
          
          // 데이터 삭제
          setSubmitStatus('Cleaning up...');
          setProgress(90);
          
          try {
            await deleteProcessedAppointments(appointmentsToDeleteFromDb);
          } catch (deleteError) {
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
      setSubmitStatus('error');
      setProgress(0);
      setTimeout(() => {
        setPdfLoading(false);
        setSubmitStatus('');
      }, 3000);
    }
  };

  const deleteProcessedAppointments = async (appointmentsToDelete?: any[]) => {
    try {
      const appointmentsToProcess = appointmentsToDelete || filteredAppointments;
      const byDoc = new Map<string, any[]>();
      appointmentsToProcess.forEach(apt => {
        if (!byDoc.has(apt.docId)) byDoc.set(apt.docId, []);
        byDoc.get(apt.docId)!.push(apt);
      });

      const successfullyRemoved: any[] = [];

      for (const [docId, apts] of byDoc) {
        const docRef = doc(db, "show-noshow", docId);
        const docSnap = await getDoc(docRef);
        const data = docSnap.exists() ? docSnap.data() : null;
        if (!data || !Array.isArray(data.patients)) continue;

        const patients = removePatientsFromList(data.patients, apts);
        if (patients.length === data.patients.length) continue;

        if (patients.length === 0) await deleteDoc(docRef);
        else await updateDoc(docRef, sanitizeFirebaseDataClient({ patients, lastUpdated: new Date().toISOString() }));

        for (const apt of apts) {
          const wasInDoc = data.patients.some(p => isSamePatient(p, apt));
          const stillInDoc = patients.some(p => isSamePatient(p, apt));
          if (wasInDoc && !stillInDoc) successfullyRemoved.push(apt);
        }
      }

      if (successfullyRemoved.length === 0) {
        throw new Error('No matching patients found in database');
      }

      const updated = appointmentsRef.current.filter(apt =>
        !successfullyRemoved.some(removed => isSameAppointment(apt, removed))
      );
      setAppointments(updated);
      appointmentsRef.current = updated;
      setFilteredAppointments(prev =>
        prev.filter(apt => !successfullyRemoved.some(removed => isSameAppointment(apt, removed)))
      );
    } catch (error) {
      throw error;
    }
  };

  const handleDelete = async (appointment: any) => {
    try {
      await deleteProcessedAppointments([appointment]);
    } catch (error) {
      alert('⚠️ Failed to remove appointment. Please try again.');
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

  const deleteButtonStyle = {
    ...buttonStyle,
    backgroundColor: '#6c757d',
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
      case 'show': return '#d4edda';
      case 'no-show': return '#f8d7da'; 
      case 'pending': return '#fff3cd';
      default: return index % 2 === 0 ? '#f9f9f9' : 'white'; 
    }
  };

  const getStatusText = (status: string): string => {
    switch (status) {
      case 'show': return 'Show';
      case 'no-show': return 'No Show';
      default: return 'Pending';
    }
  };

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

            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Month:
              </label>
              <select
                value={availableMonths.includes(selectedMonth) ? selectedMonth : ''}
                onChange={(e) => {
                  const validatedValue = validateInput('selectedMonth', e.target.value);
                  if (validatedValue) setSelectedMonth(validatedValue);
                }}
                style={inputStyle}
                disabled={availableMonths.length === 0}
              >
                {availableMonths.length === 0 ? (
                  <option value="">No data</option>
                ) : (
                  availableMonths.map(month => (
                    <option key={month} value={month}>{formatMonthLabel(month)}</option>
                  ))
                )}
              </select>
            </div>

            <div style={{ flex: '1', minWidth: '240px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Date:
              </label>
              <select
                value={datesForSelectedMonth.some(d => d.date === selectedDate) ? selectedDate : ''}
                onChange={(e) => {
                  const validatedValue = validateInput('selectedDate', e.target.value);
                  if (validatedValue) setSelectedDate(validatedValue);
                }}
                style={inputStyle}
                disabled={datesForSelectedMonth.length === 0}
              >
                {datesForSelectedMonth.length === 0 ? (
                  <option value="">No appointments this month</option>
                ) : (
                  datesForSelectedMonth.map(({ date }) => (
                    <option key={date} value={date}>
                      {formatShortDate(date)}
                    </option>
                  ))
                )}
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
                      key={`${appointment.docId}-${appointment.rowIndex ?? index}`} 
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
                        <div style={{ display: 'flex', gap: '4px', justifyContent: 'center', flexWrap: 'wrap' }}>
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
                            onClick={() => handleDelete(appointment)}
                            style={deleteButtonStyle}
                          >
                            Delete
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
                const rateDenom = showCount + noShowCount;
                const showRate = rateDenom > 0 ? ((showCount / rateDenom) * 100).toFixed(1) : '0';
                
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
            const hasPendingStatus = filteredAppointments.some(apt => {
              const s = apt.showStatus ?? apt.actions;
              return s !== 'show' && s !== 'no-show';
            });
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