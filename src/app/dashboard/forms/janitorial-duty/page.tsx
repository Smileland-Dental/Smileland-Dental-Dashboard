'use client';

import React, { useState, useEffect, useCallback, useRef } from 'react';
import { db, auth } from '@/lib/firebase.config';
import { doc, setDoc, getDoc, onSnapshot, deleteDoc } from 'firebase/firestore';
import { getStorage, ref, uploadBytes } from 'firebase/storage';
import { onAuthStateChanged } from 'firebase/auth';
import { pdf, Document, Page, View, Text, StyleSheet } from '@react-pdf/renderer';
import Script from 'next/script';

// 기본 duties 배열 (공통 duties 1-16)
const commonDuties = [
  'Stock Breakroom (Cups / Papertowels / Soap)',
  'Outside Walk Through and Inspection (Pick Up Trash)',
  'Water Plants (Facing Parking Lot)',
  'Sweep Breakroom',
  'Vacuum (2nd Floor)',
  'Mop (2nd Floor)',
  'Sweep / Mop Restrooms (1st & 2nd Floor)',
  'Sweep Vault',
  'Wipe Table in Meeting Rooms (1st & 2nd Floor / Conference Room)',
  'Sweep / Mop Breakroom',
  'Wipe / Clean Microwave, Air Fryer, Toaster',
  'Wipe / Clean Keurig, Water Disp. Machine',
  'Wipe Countertops & Tables in Break Room',
  'Vacuum (1st Floor)',
  'Throw Out Boxes',
  'Throw Out Trashes (Do Not Leave Trashed Food Inside the Office Overnight)'
];

// 요일별 특정 duties
const daySpecificDuties: { [key: string]: string[] } = {
  'Monday': ['Sweep / Mop Vault'],
  'Tuesday': ['Submit Order for Cleaning Supplies'],
  'Wednesday': ['Wipe Breakroom Seats'],
  'Thursday': ['Scrub Breakroom Sink (1st & 2nd Floor)', 'Clean / Wipe Microwave (Wash Glass Tray)', 'Wash Air Fryer'],
  'Friday': ['Dust Windows Sills', "Clean / Wipe Refrigerator (Throw out left overs on Friday's)"]
};

// 요일별 전체 duties 배열 생성 함수
function getDutiesForDay(dayName: string): string[] {
  return [...commonDuties, ...(daySpecificDuties[dayName] || [])];
}

const maxDuties = [17, 17, 17, 19, 18]; // Monday~Friday (공통 16개 + 요일별 특정)
const weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];

// Time 라벨 생성 함수
function getTimeLabel(page: number, index: number): string {
  const sheetName = weekdays[page - 1];
  const r = index;
  
  if (sheetName === 'Monday') {
    if (r >= 1 && r <= 8) return 'AM';
    else if (r >= 9 && r <= 16) return 'PM';
    else if (r === 17) return 'Mon';
  } else if (sheetName === 'Tuesday') {
    if (r >= 1 && r <= 8) return 'AM';
    else if (r >= 9 && r <= 16) return 'PM';
    else if (r === 17) return 'Tues';
  } else if (sheetName === 'Wednesday') {
    if (r >= 1 && r <= 8) return 'AM';
    else if (r >= 9 && r <= 16) return 'PM';
    else if (r === 17) return 'Wed';
  } else if (sheetName === 'Thursday') {
    if (r >= 1 && r <= 8) return 'AM';
    else if (r >= 9 && r <= 16) return 'PM';
    else if (r >= 17 && r <= 19) return 'Thurs';
  } else if (sheetName === 'Friday') {
    if (r >= 1 && r <= 8) return 'AM';
    else if (r >= 9 && r <= 16) return 'PM';
    else if (r >= 17 && r <= 18) return 'Fri';
  }
  return '';
}

export default function JanitorialDutyPage() {
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

  // 초기 날짜와 요일 계산
  const getInitialDateAndDay = () => {
    const now = new Date();
    const laTimeString = now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' });
    const laDate = new Date(laTimeString);
    const dayOfWeek = laDate.getDay();
    const pageForToday = dayOfWeek === 0 || dayOfWeek === 6 ? 1 : dayOfWeek;
    return {
      date: getCurrentCaliforniaTime(),
      day: pageForToday
    };
  };

  const initialValues = getInitialDateAndDay();

  // 날짜로부터 요일 계산 함수
  const getDayFromDate = (dateString: string): number => {
    const date = new Date(dateString + 'T00:00:00');
    const dayOfWeek = date.getDay();
    // 주말(0=일요일, 6=토요일)이면 월요일(1)로 설정, 그 외에는 실제 요일
    return dayOfWeek === 0 || dayOfWeek === 6 ? 1 : dayOfWeek;
  };

  const [selectedDate, setSelectedDate] = useState(initialValues.date);
  const [selectedDay, setSelectedDay] = useState(initialValues.day);
  const [dutyData, setDutyData] = useState<any>({});
  const [loading, setLoading] = useState(false);
  const [isInitialLoad, setIsInitialLoad] = useState(true);
  const [autoSaveStatus, setAutoSaveStatus] = useState('');
  const [submitStatus, setSubmitStatus] = useState('');
  const [progress, setProgress] = useState(0);
  const [isUpdatingFromFirebase, setIsUpdatingFromFirebase] = useState(false);
  const [userSessionId] = useState(() => Math.random().toString(36).substr(2, 9));
  const [lastSavedData, setLastSavedData] = useState<any>({});
  const [isAuthorized, setIsAuthorized] = useState<boolean | null>(null);
  const select2Initialized = useRef(false);

  // --- PDF 생성 관련 상수/스타일 ---
  const pdfStyles = StyleSheet.create({
    page: { padding: 22, fontFamily: 'Helvetica', fontSize: 8 },
    header: { marginBottom: 10, borderBottomWidth: 2, borderColor: '#333', paddingBottom: 6, alignItems: 'center' },
    headerTitle: { fontSize: 14, fontWeight: 'bold', marginBottom: 4 },
    headerSub: { fontSize: 10 },
    row: { flexDirection: 'row', borderBottomWidth: 0.5, borderColor: '#333' },
    cell: { padding: 3, fontSize: 6, flex: 1, borderRightWidth: 0.5, borderColor: '#333', justifyContent: 'center', alignItems: 'center' },
    cellDuty: { flex: 5 },
    cellTime: { flex: 1.5 },
    cellBy: { flex: 1.5 },
    cellBold: { fontWeight: 'bold' },
    cellGray: { backgroundColor: '#f0f0f0' },
    footer: { marginTop: 15, paddingTop: 8, borderTopWidth: 1, borderColor: '#ddd', alignItems: 'center', fontSize: 7, color: '#666' },
  });

  function safeStr(v: unknown, max: number): string {
    if (v == null) return '';
    return String(v).trim().slice(0, max).replace(/[<>]/g, '');
  }

  function isCheckedValue(v: unknown): boolean {
    return v === true || v === 1 || (typeof v === 'string' && (v === 'true' || v === '1'));
  }

  function createJanitorialPDFDocument(props: {
    dayName: string;
    formattedDate: string;
    duties: string[];
    dutyData: Record<string, unknown>;
    selectedDay: number;
    generatedDate: string;
  }) {
    const { dayName, formattedDate, duties, dutyData, selectedDay, generatedDate } = props;
    const s = pdfStyles;

    // Header row
    const headerRow = React.createElement(View, { key: 'header', style: [s.row, s.cellGray] },
      React.createElement(View, { style: s.cell }, React.createElement(Text, { style: s.cellBold }, 'Time')),
      React.createElement(View, { style: [s.cell, s.cellDuty] }, React.createElement(Text, { style: s.cellBold }, 'Daily Duties')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, { style: s.cellBold }, 'Check')),
      React.createElement(View, { style: s.cell }, React.createElement(Text, { style: s.cellBold }, 'Skip')),
      React.createElement(View, { style: [s.cell, s.cellTime] }, React.createElement(Text, { style: s.cellBold }, 'Time')),
      React.createElement(View, { style: [s.cell, s.cellBy] }, React.createElement(Text, { style: s.cellBold }, 'By')),
    );

    // Data rows
    const dataRows = duties.map((duty, idx) => {
      const index = idx + 1;
      const timeLabel = getTimeLabel(selectedDay, index);
      const checkValue = isCheckedValue(dutyData[`Duty${index}_Check`]);
      const skipValue = isCheckedValue(dutyData[`Duty${index}_Skip`]);
      const timeValue = safeStr(dutyData[`Duty${index}_Time`], 50);
      const byValue = dutyData[`Duty${index}_By`];
      const byStr = Array.isArray(byValue) ? byValue.join(', ') : safeStr(byValue, 50);

      return React.createElement(View, { key: index, style: s.row },
        React.createElement(View, { style: s.cell }, React.createElement(Text, null, timeLabel)),
        React.createElement(View, { style: [s.cell, s.cellDuty] }, React.createElement(Text, null, duty)),
        React.createElement(View, { style: s.cell }, React.createElement(Text, null, checkValue ? 'O' : '')),
        React.createElement(View, { style: s.cell }, React.createElement(Text, null, skipValue ? 'O' : '')),
        React.createElement(View, { style: [s.cell, s.cellTime] }, React.createElement(Text, null, timeValue)),
        React.createElement(View, { style: [s.cell, s.cellBy] }, React.createElement(Text, null, byStr)),
      );
    });

    const table = React.createElement(View, null, headerRow, ...dataRows);

    const header = React.createElement(View, { style: s.header },
      React.createElement(Text, { style: s.headerTitle }, 'Janitorial Duty'),
      React.createElement(Text, { style: s.headerSub }, `${dayName} (${formattedDate})`),
    );

    const footer = React.createElement(View, { style: s.footer }, React.createElement(Text, null, `Generated: ${generatedDate}`));

    return React.createElement(Document, null,
      React.createElement(Page, { size: 'A4', orientation: 'landscape', style: s.page }, header, table, footer),
    );
  }

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
    if (!selectedDate || !selectedDay || isUpdatingFromFirebase) return;

    const hasChanges = JSON.stringify(dutyData) !== JSON.stringify(lastSavedData);
    if (!hasChanges) return;

    try {
      const dataToSave = {
        selectedDate,
        selectedDay,
        ...dutyData,
        timestamp: new Date().toISOString(),
        autoSaved: true,
        lastUpdatedBy: userSessionId
      };

      const dayName = weekdays[selectedDay - 1];
      const docId = `${selectedDate}_${dayName}_janitorial`;
      await setDoc(doc(db, "janitorial-duties", docId), dataToSave);
      
      setLastSavedData({ ...dutyData });

    } catch (error) {
    }
  }, [selectedDate, selectedDay, dutyData, lastSavedData, isUpdatingFromFirebase, userSessionId]);

  // 데이터 변경 시에만 자동 저장
  useEffect(() => {
    if (Object.values(dutyData).some(value => value !== '' && value !== null && value !== undefined)) {
      autoSave();
    }
  }, [dutyData]);

  // 데이터 로드 함수
  const loadData = async () => {
    if (!selectedDate || !selectedDay) return Promise.resolve();

    try {
      setSubmitStatus('Loading data...');
      
      const dayName = weekdays[selectedDay - 1];
      const docId = `${selectedDate}_${dayName}_janitorial`;
      const docSnap = await getDoc(doc(db, "janitorial-duties", docId));
      
      if (docSnap.exists()) {
        const data = docSnap.data();
        
        setIsUpdatingFromFirebase(true);
        
        setDutyData((prevData: any) => ({
          ...prevData,
          ...data
        }));
        
        setLastSavedData({ ...data });
        
        setTimeout(() => {
          setIsUpdatingFromFirebase(false);
        }, 100);
        
        setSubmitStatus('Data loaded successfully');
        setTimeout(() => setSubmitStatus(''), 2000);
      } else {
        setDutyData({});
        setLastSavedData({});
        setSubmitStatus('No data found - initialized empty form');
        setTimeout(() => setSubmitStatus(''), 2000);
      }
      
      return Promise.resolve();
    } catch (error: any) {
      setSubmitStatus('Error loading data: ' + error.message);
      setTimeout(() => setSubmitStatus(''), 3000);
      return Promise.resolve();
    }
  };

  // 날짜, 요일 변경 시 데이터 로드
  useEffect(() => {
    if (selectedDate && selectedDay) {
      loadData().then(() => {
        // 데이터 로드 완료 후 초기 로드 상태 해제
        setIsInitialLoad(false);
      });
    }
  }, [selectedDate, selectedDay]);

  // 실시간 데이터 동기화
  useEffect(() => {
    if (!selectedDate || !selectedDay) return;

    const dayName = weekdays[selectedDay - 1];
    const docId = `${selectedDate}_${dayName}_janitorial`;
    const docRef = doc(db, "janitorial-duties", docId);
    
    // 날짜나 요일이 변경되면 먼저 데이터 초기화
    setDutyData({});
    setLastSavedData({});
    
    const unsubscribe = onSnapshot(docRef, (docSnap: any) => {
      if (docSnap.exists()) {
        const data = docSnap.data();
        
        // selectedDate와 selectedDay를 제외한 실제 duty 데이터만 추출
        const { selectedDate: _, selectedDay: __, timestamp, autoSaved, lastUpdatedBy, ...dutyFields } = data;
        
        const hasChanges = JSON.stringify(dutyFields) !== JSON.stringify(dutyData);
        if (!hasChanges) return;
        
        setIsUpdatingFromFirebase(true);
        
        // 이전 데이터를 병합하지 않고 새로운 데이터로 완전히 교체
        setDutyData(dutyFields);
        setLastSavedData(dutyFields);
        
        setTimeout(() => {
          setIsUpdatingFromFirebase(false);
        }, 100);
        
        if (data.timestamp && 
            new Date(data.timestamp).getTime() > Date.now() - 5000 && 
            data.lastUpdatedBy && 
            data.lastUpdatedBy !== userSessionId) {
          setAutoSaveStatus('🔄 Updated by another user.');
          setTimeout(() => setAutoSaveStatus(''), 3000);
        }
      } else {
        // 문서가 없으면 데이터 초기화
        setDutyData({});
        setLastSavedData({});
      }
    }, (error: any) => {
      setAutoSaveStatus('❌ 연결 오류 - 실시간 동기화가 중단되었습니다');
      setTimeout(() => setAutoSaveStatus(''), 3000);
    });

    return () => {
      unsubscribe();
    };
  }, [selectedDate, selectedDay, userSessionId]);

  // 데이터 업데이트 함수
  const updateDutyData = (field: string, value: any) => {
    setDutyData((prev: any) => {
      const newData = { ...prev, [field]: value };
      
      // By 필드가 선택되면 시간 자동 기록
      if (field.includes('_By') && value && value.length > 0) {
        const rowNumber = field.match(/Duty(\d+)_By/)?.[1];
        if (rowNumber) {
          const timeField = `Duty${rowNumber}_Time`;
          if (!newData[timeField]) {
            newData[timeField] = getCurrentTime12Hour();
          }
        }
      }
      
      return newData;
    });
  };

  // Select2 초기화 (페이지 변경 시 재초기화)
  useEffect(() => {
    if (select2Initialized.current && typeof window !== 'undefined' && (window as any).$ && (window as any).$.fn && (window as any).$.fn.select2) {
      // Day가 변경되면 Select2 값도 강제로 업데이트
      try {
        (window as any).$('.select2-multi').each(function(this: any) {
          const $this = (window as any).$(this);
          const name = $this.attr('name');
          if (name) {
            // dutyData에서 현재 값을 가져와서 Select2에 설정
            const currentValue = dutyData[name] || [];
            if ($this.hasClass('select2-hidden-accessible')) {
              // 이미 Select2로 변환된 경우
              $this.val(currentValue).trigger('change');
            } else {
              // 아직 Select2로 변환되지 않은 경우
              $this.select2({
                width: '120px',
                placeholder: '',
                allowClear: true
              });
              $this.val(currentValue).trigger('change');
            }
          }
        });
      } catch (error) {
      }
    }
  }, [selectedDay, dutyData]);

  // 제출 함수
  const handleSubmit = async () => {
    if (!selectedDate || !selectedDay) {
      alert('Please select a date and day first.');
      return;
    }

    const confirmed = confirm('Would you like to submit?');
    if (!confirmed) return;

    setLoading(true);
    setSubmitStatus('Saving...');
    setProgress(10);

    try {
      // 1. PDF 생성
      setSubmitStatus('Processing...');
      setProgress(30);
      
      const dayName = weekdays[selectedDay - 1];
      const duties = getDutiesForDay(dayName);

      // 날짜 포맷팅 (MM/DD/YYYY)
      const dateObj = new Date(selectedDate + 'T00:00:00');
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

      const pdfDoc = createJanitorialPDFDocument({
        dayName,
        formattedDate,
        duties,
        dutyData: dutyData as Record<string, unknown>,
        selectedDay,
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
        
        const office = 'Janitor';
        const filename = `${selectedDate}_${dayName}_Janitorial Duty_${timeStamp}.pdf`;
        const storageRef = ref(storage, `endofday-pdfs/${office}/${selectedDate}/${filename}`);
        
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
      const docId = `${selectedDate}_${dayName}_janitorial`;
      await deleteDoc(doc(db, "janitorial-duties", docId));
      
      // 3. 폼 초기화
      setDutyData({});
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

  // 테이블 행 렌더링
  const renderDutyRows = () => {
    const dayName = weekdays[selectedDay - 1];
    const currentDuties = getDutiesForDay(dayName);
    
    return currentDuties.map((duty, idx) => {
      const index = idx + 1;
      const timeLabel = getTimeLabel(selectedDay, index);
      const checkKey = `Duty${index}_Check`;
      const byKey = `Duty${index}_By`;
      const timeKey = `Duty${index}_Time`;
      const skipKey = `Duty${index}_Skip`;
      
      // AM과 PM 사이 구분선 (index가 9일 때, PM의 첫 번째 행)
      const isPmFirstRow = index === 9;
      
      return (
        <tr 
          key={index}
          style={isPmFirstRow ? {
            borderTop: '3px solid #495057'
          } : {}}
        >
          <td style={{padding: '12px', textAlign: 'center'}}>{timeLabel}</td>
          <td style={{padding: '12px'}}>{duty}</td>
          <td style={{padding: '12px', textAlign: 'center'}}>
            <input
              type="checkbox"
              name={checkKey}
              checked={dutyData[checkKey] || false}
              disabled={dutyData[skipKey] || false}
              onChange={(e) => {
                if (e.target.checked) {
                  // Check가 체크되면 Skip을 해제
                  updateDutyData(checkKey, true);
                  updateDutyData(skipKey, false);
                } else {
                  updateDutyData(checkKey, false);
                }
              }}
            />
          </td>
          <td style={{padding: '12px', textAlign: 'center', background: 'linear-gradient(135deg, #fcd34d 0%, #fbbf24 100%)'}}>
            <input
              type="checkbox"
              name={skipKey}
              checked={dutyData[skipKey] || false}
              disabled={dutyData[checkKey] || false}
              onChange={(e) => {
                if (e.target.checked) {
                  // Skip이 체크되면 Check를 해제
                  updateDutyData(skipKey, true);
                  updateDutyData(checkKey, false);
                } else {
                  updateDutyData(skipKey, false);
                }
              }}
            />
          </td>
          <td style={{padding: '12px'}}>
            <input
              type="text"
              name={timeKey}
              value={dutyData[timeKey] || ''}
              readOnly
              style={{width: '100%', padding: '6px', border: '1px solid #c4cdd5', borderRadius: '4px'}}
            />
          </td>
          <td style={{padding: '12px'}}>
            <select
              name={byKey}
              multiple
              size={4}
              className="select2-multi"
              value={dutyData[byKey] || []}
              onChange={(e) => {
                const values = Array.from(e.target.selectedOptions, (opt: HTMLOptionElement) => opt.value);
                updateDutyData(byKey, values);
              }}
            >
              <option value="A">A</option>
              <option value="M">M</option>
              <option value="Other">Other</option>
            </select>
          </td>
        </tr>
      );
    });
  };

  // Select2 스크립트 로드 완료 핸들러
  const handleSelect2Load = () => {
    if (typeof window !== 'undefined' && (window as any).$ && (window as any).$.fn && (window as any).$.fn.select2) {
      // Select2가 로드되었으므로 초기화 시도
      setTimeout(() => {
        if (!select2Initialized.current) {
          try {
            (window as any).$('.select2-multi').select2({
              width: '120px',
              placeholder: '',
              allowClear: true
            });

            (window as any).$('.select2-multi').on('select2:close', function(this: any) {
              const $this = (window as any).$(this);
              if ($this.val() && $this.val().length > 0) {
                $this.prop('disabled', true);
                $this.trigger('change.select2');
                
                const name = $this.attr('name');
                const match = name.match(/Duty(\d+)_By/);
                if (match) {
                  const index = parseInt(match[1]);
                  updateDutyData(`Duty${index}_By`, $this.val());
                  
                  const dutyCheckbox = document.querySelector(`input[name="Duty${index}_Check"]`) as HTMLInputElement;
                  if (dutyCheckbox && !dutyCheckbox.checked) {
                    updateDutyData(`Duty${index}_Check`, true);
                  }
                }
              }
            });

            select2Initialized.current = true;
          } catch (error) {
          }
        }
      }, 100);
    }
  };

  // 페이지 로드 시 즉시 스타일 적용 (FOUC 방지)
  useEffect(() => {
    if (typeof document !== 'undefined') {
      // head에 스타일 태그 추가
      const style = document.createElement('style');
      style.id = 'janitorial-inline-styles';
      style.textContent = `
        html, body {
          margin: 0 !important;
          padding: 0 !important;
          background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%) !important;
          background-attachment: fixed !important;
          min-height: 100vh !important;
        }
        .janitorial-container {
          width: 80% !important;
          max-width: 1600px !important;
          min-width: 900px !important;
          margin: 20px auto !important;
          background: #ffffff !important;
          border-radius: 24px !important;
          box-shadow: 0 20px 60px rgba(0, 0, 0, 0.1) !important;
          padding: 50px !important;
          border: 1px solid rgba(0, 0, 0, 0.05) !important;
          position: relative !important;
        }
        .janitorial-container.initial-loading {
          visibility: hidden !important;
        }
      `;
      
      // 이미 존재하면 제거하고 다시 추가
      const existingStyle = document.getElementById('janitorial-inline-styles');
      if (existingStyle) {
        existingStyle.remove();
      }
      document.head.insertBefore(style, document.head.firstChild);
      
      // body에도 직접 스타일 적용 (즉시 반영)
      if (document.body) {
        document.body.style.margin = '0';
        document.body.style.padding = '0';
        document.body.style.background = 'linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%)';
        document.body.style.backgroundAttachment = 'fixed';
        document.body.style.minHeight = '100vh';
      }
    }
  }, []);

  // 초기 로드 완료 처리
  useEffect(() => {
    // 컴포넌트가 마운트되고 스타일이 적용된 후 초기 로드 상태 해제
    const timer = setTimeout(() => {
      setIsInitialLoad(false);
    }, 200);
    return () => clearTimeout(timer);
  }, []);

  // 컴포넌트 마운트 시 사용자 인증 및 role 확인
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      try {
        if (!currentUser) {
          alert('Please log in.');
          setIsAuthorized(false);
          return;
        }

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
          if (typeof window !== 'undefined') {
            window.location.href = '/';
          }
          return;
        }

        setIsAuthorized(true);
      } catch (error: any) {
        alert('An error occurred while verifying authentication.');
        setIsAuthorized(false);
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

  // 인증 확인 중이거나 인증 실패 시 로딩 화면 표시
  if (isAuthorized === null) {
    return (
      <div style={{
        display: 'flex',
        justifyContent: 'center',
        alignItems: 'center',
        height: '100vh',
        background: 'linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%)',
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
        background: 'linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%)',
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
      <style dangerouslySetInnerHTML={{__html: `
        html, body {
          margin: 0 !important;
          padding: 0 !important;
          background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%) !important;
          background-attachment: fixed !important;
          min-height: 100vh !important;
        }
      `}} />
      <Script 
        src="https://code.jquery.com/jquery-3.6.0.min.js" 
        strategy="beforeInteractive"
        onLoad={() => {
          // jQuery 로드 후 Select2 로드
        }}
      />
      <Script 
        src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js" 
        strategy="lazyOnload"
        onLoad={handleSelect2Load}
      />
      <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
      <style jsx global>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        body {
          font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
          background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
          background-attachment: fixed;
          color: #1a202c;
          line-height: 1.6;
          min-height: 100vh;
          margin: 0;
          padding: 20px;
          overflow-x: hidden;
        }
        html {
          overflow-x: hidden;
          background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
          background-attachment: fixed;
        }
        .container {
          width: 80%;
          max-width: 1600px;
          min-width: 900px;
          margin: 20px auto;
          background: #ffffff;
          border-radius: 24px;
          box-shadow: 0 20px 60px rgba(0, 0, 0, 0.1);
          padding: 50px;
          border: 1px solid rgba(0, 0, 0, 0.05);
          position: relative;
        }
        h2 {
          font-size: 2.5rem;
          font-weight: 700;
          margin-bottom: 40px;
          color: #1a202c;
          text-align: center;
          letter-spacing: -0.5px;
        }
        .date-row {
          display: flex;
          align-items: center;
          justify-content: center;
          margin-bottom: 40px;
          background: linear-gradient(135deg, #90cdf4 0%, #63b3ed 100%);
          padding: 25px;
          border-radius: 16px;
          border: none;
          gap: 25px;
          box-shadow: 0 10px 30px rgba(144, 205, 244, 0.2);
        }
        .date-row label {
          font-weight: 600;
          color: #ffffff;
          font-size: 1.1rem;
        }
        .date-row select, .date-row input {
          padding: 14px 24px;
          border: 2px solid rgba(255, 255, 255, 0.3);
          border-radius: 12px;
          font-size: 1rem;
          font-weight: 500;
          background: rgba(255, 255, 255, 0.95);
          color: #1a202c;
          cursor: pointer;
          min-width: 150px;
          transition: all 0.2s ease;
        }
        .date-row select:focus, .date-row input:focus {
          outline: none;
          border-color: rgba(255, 255, 255, 0.6);
          background: #ffffff;
          box-shadow: 0 0 0 4px rgba(255, 255, 255, 0.2);
        }
        .date-row select:disabled {
          background: rgba(255, 255, 255, 0.7);
          color: #718096;
          cursor: not-allowed;
        }
        table {
          width: 100%;
          border-collapse: separate;
          border-spacing: 0;
          margin-bottom: 30px;
          border-radius: 16px;
          overflow: hidden;
          box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
          background: #ffffff;
        }
        th, td {
          padding: 18px 24px;
          text-align: left;
          border-bottom: 1px solid #e2e8f0;
        }
        thead {
          background: linear-gradient(135deg, #90cdf4 0%, #63b3ed 100%);
        }
        thead th {
          color: #1a202c;
          font-weight: 600;
          font-size: 0.95rem;
          text-transform: uppercase;
          letter-spacing: 0.8px;
        }
        tbody tr:nth-child(even) {
          background: #f7fafc;
        }
        tbody tr:hover {
          background: #edf2f7;
          transition: background 0.2s ease;
        }
        .submit-btn {
          display: block;
          width: 100%;
          max-width: 320px;
          margin: 40px auto;
          padding: 18px 40px;
          font-size: 1.1rem;
          font-weight: 600;
          background: linear-gradient(135deg, #90cdf4 0%, #63b3ed 100%);
          color: #fff;
          border: none;
          border-radius: 14px;
          cursor: pointer;
          box-shadow: 0 10px 30px rgba(144, 205, 244, 0.3);
          transition: all 0.3s ease;
        }
        .submit-btn:hover {
          transform: translateY(-3px);
          box-shadow: 0 15px 40px rgba(144, 205, 244, 0.4);
        }
        .submit-btn:active {
          transform: translateY(-1px);
        }
        .submit-btn:disabled {
          opacity: 0.6;
          cursor: not-allowed;
          transform: none;
        }
        .auto-save-status {
          position: fixed;
          top: 20px;
          right: 20px;
          padding: 12px 20px;
          background-color: #51cf66;
          color: white;
          border-radius: 25px;
          fontSize: 14px;
          fontWeight: bold;
          boxShadow: 0 4px 12px rgba(0,0,0,0.3);
          zIndex: 1000;
          maxWidth: 300px;
          textAlign: center;
        }
      `}</style>

      <div className={`container janitorial-container ${isInitialLoad ? 'initial-loading' : ''}`} style={isInitialLoad ? { visibility: 'hidden' } : {}}>
        {autoSaveStatus && (
          <div style={{
            position: 'fixed',
            top: '20px',
            right: '20px',
            padding: '12px 20px',
            backgroundColor: autoSaveStatus.includes('❌') ? '#ef4444' : 
                            autoSaveStatus.includes('🔄') ? '#4a90e2' : 
                            autoSaveStatus.includes('💾') ? '#51cf66' : '#51cf66',
            color: 'white',
            borderRadius: '25px',
            fontSize: '14px',
            fontWeight: 'bold',
            boxShadow: '0 4px 12px rgba(0,0,0,0.3)',
            zIndex: 1000,
            maxWidth: '300px',
            textAlign: 'center'
          }}>
            {autoSaveStatus}
          </div>
        )}

        <h2>Janitorial Duty</h2>
        
        <div className="date-row">
          <label htmlFor="date">📅 Date:</label>
          <input
            type="date"
            id="date"
            value={selectedDate}
            onChange={(e) => {
              const newDate = e.target.value;
              setSelectedDate(newDate);
              // 날짜가 변경되면 해당 날짜의 실제 요일로 Day도 자동 업데이트
              const newDay = getDayFromDate(newDate);
              setSelectedDay(newDay);
            }}
          />
          <label htmlFor="day">📆 Day:</label>
          <select
            id="day"
            value={selectedDay}
            disabled
            style={{
              padding: '12px 20px',
              border: '2px solid #dee2e6',
              borderRadius: '8px',
              fontSize: '1rem',
              fontWeight: '500',
              background: '#f8f9fa',
              color: '#6c757d',
              cursor: 'not-allowed',
              minWidth: '150px'
            }}
          >
            {weekdays.map((day, idx) => (
              <option key={idx + 1} value={idx + 1}>{day}</option>
            ))}
          </select>
        </div>

        {selectedDate && selectedDay && (
          <>
            <table>
              <thead>
                <tr>
                  <th>Time</th>
                  <th>Daily Duties</th>
                  <th>Check</th>
                  <th>Skip/Not Needed</th>
                  <th>Time</th>
                  <th>By</th>
                </tr>
              </thead>
              <tbody>
                {renderDutyRows()}
              </tbody>
            </table>

            <button
              className="submit-btn"
              onClick={handleSubmit}
              disabled={loading}
            >
              {loading ? 'Submitting...' : 'Submit'}
            </button>
          </>
        )}

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
      </div>
    </>
  );
}



