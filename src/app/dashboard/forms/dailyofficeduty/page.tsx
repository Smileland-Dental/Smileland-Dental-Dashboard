'use client'

import React, { useState, useEffect, useCallback } from "react";
import { doc, setDoc, getDoc, deleteDoc, onSnapshot, collection, getDocs } from "firebase/firestore";
import { getStorage, ref, uploadBytes, getDownloadURL } from 'firebase/storage';
import { db } from "@/lib/firebase.config";

export default function DailyOfficeDuties() {
  const [loading, setLoading] = useState(false);
  const [autoSaveStatus, setAutoSaveStatus] = useState('');
  const [submitStatus, setSubmitStatus] = useState('');
  const [progress, setProgress] = useState(0);
  const [isUpdatingFromFirebase, setIsUpdatingFromFirebase] = useState(false);
  
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
  // 비밀번호 확인 상태
  const [officePasswordVerified, setOfficePasswordVerified] = useState(false);
  const [showPasswordModal, setShowPasswordModal] = useState(false);
  const [pendingOffice, setPendingOffice] = useState('');
  
  // 오피스 옵션
  const officeOptions = ['Ming', 'Bernard', 'Delano', 'Tulare', 'Visalia', 'Fresno', 'California', 'Ortho'];
  
  // 오피스 비밀번호 가져오기 함수
  const getOfficePassword = (office: string): string => {
    if (!office) return '';
    return office.charAt(0).toUpperCase();
  };

  // 모든 업무 항목 상태
  const [dutyData, setDutyData] = useState({
    // Row 1
    Row1_Done: '',
    Row1_Checked: '',
    Row1_Time: '',
    
    // Row 2
    Row2_YesNo: '',
    Row2_Done: '',
    Row2_Checked: '',
    Row2_Time: '',
    
    // Row 3
    Row3_Done: '',
    Row3_Checked: '',
    Row3_Time: '',
    
    // Row 4
    Row4_Done: '',
    Row4_Checked: '',
    Row4_Time: '',
    
    // Row 5
    Row5_CallNum: '',
    Row5_Done: '',
    Row5_Checked: '',
    Row5_Time: '',
    
    // Row 6
    Row6_Done: '',
    Row6_Checked: '',
    Row6_Time: '',
    
    // Row 7
    Row7_YesNo: '',
    Row7_Done: '',
    Row7_Checked: '',
    Row7_Time: '',
    
    // Row 8
    Row8_Done: '',
    Row8_Checked: '',
    Row8_Time: '',
    
    // Row 9
    Row9_Done: '',
    Row9_Checked: '',
    Row9_Time: '',
    
    // Row 10
    Row10_Done: '',
    Row10_Checked: '',
    Row10_Time: '',
    
    // Row 11
    Row11_Done: '',
    Row11_Checked: '',
    Row11_Time: '',
    
    // Row 12
    Row12_Done: '',
    Row12_Checked: '',
    Row12_Time: '',
    
    // Row 13
    Row13_Done: '',
    Row13_Checked: '',
    Row13_Time: '',
    
    // Row 14
    'Row14_Name/DOB': '',
    Row14_Done: '',
    Row14_Checked: '',
    Row14_Time: '',
    
    // Row 15
    Row15_LabCases: '',
    Row15_Done: '',
    Row15_Checked: '',
    Row15_Time: '',
    
    // Row 16
    Row16_Done: '',
    Row16_Checked: '',
    Row16_Time: '',
    
    // Row 17
    Row17_Done: '',
    Row17_Checked: '',
    Row17_Time: '',
    
    // Row 18
    Row18_YesNo: '',
    Row18_Done: '',
    Row18_Checked: '',
    Row18_Time: '',
    
    // Row 19
    Row19_O2: '',
    Row19_N2O: '',
    Row19_He: '',
    Row19_Done: '',
    Row19_Checked: '',
    Row19_Time: '',
    
    // Row 20
    Row20_Done: '',
    Row20_Checked: '',
    Row20_Time: '',
    
    // Row 21
    Row21_YesNo: '',
    Row21_Done: '',
    Row21_Checked: '',
    Row21_Time: '',
    
    // Row 22
    Row22_YesNo: '',
    Row22_Done: '',
    Row22_Checked: '',
    Row22_Time: '',
    
    // Row 23
    Row23_YesNo: '',
    Row23_Done: '',
    Row23_Checked: '',
    Row23_Time: '',
    
    // Row 24
    Row24_YesNo: '',
    Row24_Done: '',
    Row24_Checked: '',
    Row24_Time: '',
    
    // Row 25
    Row25_Done: '',
    Row25_Checked: '',
    Row25_Time: '',
    
    // Row 26
    Row26_Done: '',
    Row26_Checked: '',
    Row26_Time: '',
    
    // Row 27
    Row27_Done: '',
    Row27_Checked: '',
    Row27_Time: '',
    
    // Row 28
    Row28_Done: '',
    Row28_Checked: '',
    Row28_Time: '',
    
    // Row 29
    Row29_Done: '',
    Row29_Checked: '',
    Row29_Time: '',
    
    // Row 30
    Row30_YesNo: '',
    Row30_Done: '',
    Row30_Checked: '',
    Row30_Time: '',
    
    // Row 31
    Row31_Done: '',
    Row31_Checked: '',
    Row31_Time: '',
    
    // Row 32
    Row32_Done: '',
    Row32_Checked: '',
    Row32_Time: ''
  });

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
    if (!dutyDate || isUpdatingFromFirebase || !officePasswordVerified) return;

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
      
    } catch (error) {
      console.error("Auto-save error:", error);
    }
  }, [dutyDate, selectedOffice, dutyData, lastSavedData, isUpdatingFromFirebase, officePasswordVerified, userSessionId]);

  // 데이터 변경 시에만 자동 저장
  useEffect(() => {
    // 초기 로드 시에는 저장하지 않음
    if (Object.values(dutyData).some(value => value !== '')) {
      autoSave();
    }
  }, [dutyData]);

  // 데이터 로드
  const loadData = async () => {
    if (!dutyDate || !selectedOffice || !officePasswordVerified) return;

    try {
      console.log("Loading data for date:", dutyDate, "office:", selectedOffice);
      setSubmitStatus('Loading data...');
      
      const docId = `${dutyDate}_${selectedOffice}`;
      const docSnap = await getDocs(collection(db, "daily-office-duties")).then(snapshot => {
        const foundDoc = snapshot.docs.find(d => d.id === docId);
        return foundDoc ? { 
          exists: (): boolean => true, 
          data: (): any => foundDoc.data() 
        } : { 
          exists: (): boolean => false,
          data: (): undefined => undefined
        };
      });
      
      if (docSnap.exists()) {
        const data = docSnap.data();
        console.log("Data loaded from Firebase:", data);
        
        // Firebase에서 업데이트되는 동안 자동 저장 방지
        setIsUpdatingFromFirebase(true);
        
        setDutyData(prevData => ({
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
        console.log("No data found for date:", dutyDate);
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
      console.error("Error loading data:", error);
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      setSubmitStatus('Error loading data: ' + errorMessage);
      setTimeout(() => setSubmitStatus(''), 3000);
    }
  };

  // 날짜 또는 오피스 변경 시 데이터 로드
  useEffect(() => {
    if (dutyDate && selectedOffice && officePasswordVerified) {
      loadData();
    }
  }, [dutyDate, selectedOffice, officePasswordVerified]);

  // 실시간 데이터 동기화
  useEffect(() => {
    if (!dutyDate || !selectedOffice || !officePasswordVerified) return;

    console.log("Setting up real-time listener for date:", dutyDate, "office:", selectedOffice);
    const docId = `${dutyDate}_${selectedOffice}`;
    const docRef = doc(db, "daily-office-duties", docId);
    
    const unsubscribe = onSnapshot(docRef, (docSnap) => {
      if (docSnap.exists()) {
        const data = docSnap.data();
        console.log("Real-time data received:", data);
        
        // Firebase에서 업데이트되는 동안 자동 저장 방지
        setIsUpdatingFromFirebase(true);
        
        setDutyData(prevData => {
          console.log("Updating dutyData from:", prevData, "to:", { ...prevData, ...data });
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
        
        // 다른 사용자의 업데이트만 표시 (자신의 업데이트는 제외)
        if (data.timestamp && 
            new Date(data.timestamp).getTime() > Date.now() - 5000 && 
            data.lastUpdatedBy && 
            data.lastUpdatedBy !== userSessionId) {
          setAutoSaveStatus('🔄 Updated from another user');
          setTimeout(() => setAutoSaveStatus(''), 2000);
        }
      } else {
        console.log("Real-time listener: No document exists for date:", dutyDate);
      }
    }, (error) => {
      console.error("Real-time listener error:", error);
      setAutoSaveStatus('❌ Connection error');
      setTimeout(() => setAutoSaveStatus(''), 3000);
    });

    return () => {
      console.log("Cleaning up real-time listener for date:", dutyDate, "office:", selectedOffice);
      unsubscribe();
    };
  }, [dutyDate, selectedOffice, officePasswordVerified]);

  // 컴포넌트 마운트 시 초기 로드는 dutyDate 변경 시 로드로 대체됨

  // 수동 저장 함수
  const saveData = async () => {
    if (!dutyDate) return;

    try {
      setAutoSaveStatus('💾 Saving...');
      
      // 🔒 보안: 저장 전 데이터 검증
      const validatedDutyData: { [key: string]: string } = {};
      for (const [key, value] of Object.entries(dutyData)) {
        validatedDutyData[key] = validateInput(value as string, 500);
      }
      
      const dataToSave = {
        dutyDate,
        ...validatedDutyData,
        timestamp: new Date().toISOString()
      };

      await setDoc(doc(db, "daily-office-duties", dutyDate), dataToSave);
      
      setAutoSaveStatus('💾 Saved ✅');
      setTimeout(() => setAutoSaveStatus(''), 2000);
      
    } catch (error) {
      console.error("Save error:", error);
      setAutoSaveStatus('💾 Save failed ❌');
      setTimeout(() => setAutoSaveStatus(''), 3000);
    }
  };

  // 🔒 보안: 입력 검증 함수
  const validateInput = (value: string, maxLength: number = 500): string => {
    if (typeof value !== 'string') return '';
    // 길이 제한
    if (value.length > maxLength) {
      return value.substring(0, maxLength);
    }
    return value;
  };

  // 데이터 업데이트 함수
  const updateDutyData = (field: string, value: string) => {
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

  // 전체 마감 시간 체크 (기존 호환성을 위해 유지)
  const checkDeadlines = () => {
    const overdueItems: string[] = [];
    
    (Object.keys(DEADLINES) as Array<keyof typeof DEADLINES>).forEach(itemName => {
      if (isRowOverdue(itemName)) {
        overdueItems.push(DEADLINES[itemName].message);
      }
    });
    
    return overdueItems;
  };

  // 제출 처리
  const handleSubmit = async () => {
    // Submit 비밀번호 확인 (모든 오피스 동일)
    const submitPassword = 'Halloween';
    const password = prompt(`Are you sure you want to submit? Submitting will reset today's data. Enter password to proceed:`);
    if (password === null) return;
    if (password !== submitPassword) {
      alert("Incorrect password. Submission cancelled.");
      return;
    }

    try {
      setLoading(true);
      setSubmitStatus('Saving...');
      setProgress(10);

      // 1. 프린트 페이지 생성 및 열기
      setSubmitStatus('Opening print page...');
      setProgress(30);
      
      const response = await fetch('/api/generate-daily-office-duty-pdf', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          dutyDate,
          selectedOffice,
          dutyData,
          submitPassword: password // 서버 측 검증을 위해 비밀번호 전송
        }),
      });

      if (response.ok) {
        // PDF blob 받기
        setSubmitStatus('Saving PDF to archive...');
        setProgress(60);
        const blob = await response.blob();
        
        // PDF를 Firebase Storage에 저장
        try {
          const storage = getStorage();
          const filename = `2) ${dutyDate}_${selectedOffice}_Daily Office Duty.pdf`;
          const storageRef = ref(storage, `endofday-pdfs/${selectedOffice}/${dutyDate}/${filename}`);
          
          // PDF 업로드
          await uploadBytes(storageRef, blob);
          
          // 다운로드 URL 가져오기
          const downloadUrl = await getDownloadURL(storageRef);
          
          // Firestore에 메타데이터 저장
          await setDoc(doc(db, 'pdf-documents', `${dutyDate}_${selectedOffice}_daily-office-duty_${Date.now()}`), {
            filename,
            office: selectedOffice,
            date: dutyDate,
            type: 'Daily Office Duty',
            url: downloadUrl,
            storagePath: `endofday-pdfs/${selectedOffice}/${dutyDate}/${filename}`,
            createdAt: new Date(),
          });
          
          console.log('PDF saved successfully to Firebase Storage');
          setSubmitStatus('✅ PDF saved to archive successfully!');
        } catch (storageError: any) {
          console.error('Storage error:', storageError);
          const errorMsg = storageError?.message || '알 수 없는 오류';
          alert(`PDF 저장 중 오류가 발생했습니다: ${errorMsg}`);
          setSubmitStatus('❌ PDF 저장 실패');
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

        setSubmitStatus('Complete! PDF saved to archive.');
        setProgress(100);
        
        // 2초 후 모달 닫기
        setTimeout(() => {
          setLoading(false);
          setSubmitStatus('');
          setProgress(0);
        }, 2000);
      } else {
        const errorData = await response.json();
        throw new Error(errorData.error || 'PDF generation failed');
      }

    } catch (error) {
      console.error('Submit error:', error);
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      setSubmitStatus('❌ Submission failed: ' + errorMessage);
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
      background: 'linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%)',
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
    overdueRow: {
      backgroundColor: '#ffebee !important',
      borderLeft: '4px solid #f44336 !important'
    },
    overdueWarning: {
      color: '#f44336',
      fontWeight: 'bold',
      fontSize: '0.9em',
      marginTop: '4px'
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

  const overdueItems = checkDeadlines();

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
          <div style={styles.formGroup}>
            <label style={styles.label} htmlFor="selectedOffice">Office:</label>
            <select
              id="selectedOffice"
              value={selectedOffice}
              onChange={(e) => {
                const newOffice = e.target.value;
                if (newOffice && newOffice !== selectedOffice) {
                  setPendingOffice(newOffice);
                  setShowPasswordModal(true);
                }
              }}
              style={styles.input}
              required
            >
              <option value="">-- Select Office --</option>
              {officeOptions.map(office => (
                <option key={office} value={office}>{office}</option>
              ))}
            </select>
          </div>
        </div>

        {/* 업무 테이블 */}
        <table style={styles.table}>
          <thead>
            <tr>
              <th style={{ ...styles.th, width: '60px' }}>No.</th>
              <th style={styles.th}>Duty</th>
              <th style={styles.th}>Done by</th>
              <th style={styles.th}>Checked by</th>
              <th style={{ ...styles.th, width: '100px' }}>Time</th>
            </tr>
          </thead>
          <tbody>
            {/* Row 1 */}
            <tr style={isRowOverdue('Row1_Done') ? styles.overdueRow : {}}>
              <td style={styles.td}><strong>1</strong></td>
              <td style={styles.td}>
                <strong>Turn Off Answering Service</strong>
                <div style={styles.dutyDetails}>
                  1) Go to: https://smileland.my3cx.us<br/>
                  2) Log in<br/>
                  4) Ensure Your Office is Selected under "Department"<br/>
                  5) Select Override Office Hours<br/>
                  6) Select "Reset Default Office Hours"
                </div>
                <div style={styles.deadlineInfo}>Deadline: 9:00 AM</div>
                {isRowOverdue('Row1_Done') && (
                  <div style={styles.overdueWarning}>⚠️ Turn Off Answering Service should be done by 9:00 AM</div>
                )}
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row1_Done}
                  onChange={(e) => updateDutyData('Row1_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row1_Checked}
                  onChange={(e) => updateDutyData('Row1_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row1_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 2 */}
            <tr style={isRowOverdue('Row2_Done') ? styles.overdueRow : {}}>
              <td style={styles.td}><strong>2</strong></td>
              <td style={styles.td}>
                <strong>All charts filed back?</strong><br/>
                <div style={styles.inlineOption}>
                  <input
                    type="radio"
                    id="Row2_YesNo_Yes"
                    name="Row2_YesNo"
                    value="Yes"
                    checked={dutyData.Row2_YesNo === 'Yes'}
                    onChange={(e) => updateDutyData('Row2_YesNo', e.target.value)}
                    style={{ marginRight: '5px', cursor: 'pointer' }}
                  />
                  <label htmlFor="Row2_YesNo_Yes">Yes</label>
                </div>
                <div style={styles.inlineOption}>
                  <input
                    type="radio"
                    id="Row2_YesNo_No"
                    name="Row2_YesNo"
                    value="No"
                    checked={dutyData.Row2_YesNo === 'No'}
                    onChange={(e) => updateDutyData('Row2_YesNo', e.target.value)}
                    style={{ marginRight: '5px', cursor: 'pointer' }}
                  />
                  <label htmlFor="Row2_YesNo_No">No</label>
                </div>
                <div style={styles.deadlineInfo}>Deadline: 4 PM</div>
                {isRowOverdue('Row2_Done') && (
                  <div style={styles.overdueWarning}>⚠️ All charts filed back should be done by 4:00 PM</div>
                )}
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row2_Done}
                  onChange={(e) => updateDutyData('Row2_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row2_Checked}
                  onChange={(e) => updateDutyData('Row2_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row2_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 3 */}
            <tr style={isRowOverdue('Row3_Done') ? styles.overdueRow : {}}>
              <td style={styles.td}><strong>3</strong></td>
              <td style={styles.td}>
                <strong>Charts pulled for next day</strong>
                <div style={styles.deadlineInfo}>Deadline: 12:00 PM</div>
                {isRowOverdue('Row3_Done') && (
                  <div style={styles.overdueWarning}>⚠️ Charts pulled for next day should be done by 12:00 PM</div>
                )}
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row3_Done}
                  onChange={(e) => updateDutyData('Row3_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row3_Checked}
                  onChange={(e) => updateDutyData('Row3_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row3_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 4 */}
            <tr style={isRowOverdue('Row4_Done') ? styles.overdueRow : {}}>
              <td style={styles.td}><strong>4</strong></td>
              <td style={styles.td}>
                <strong>Check eligibility</strong>
                <div style={styles.dutyDetails}>
                  1st of every month come in early to check eligibility by 8:30 am
                </div>
                <div style={styles.deadlineInfo}>Deadline: 4:30 PM</div>
                {isRowOverdue('Row4_Done') && (
                  <div style={styles.overdueWarning}>⚠️ Check eligibility should be done by 4:30 PM</div>
                )}
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row4_Done}
                  onChange={(e) => updateDutyData('Row4_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row4_Checked}
                  onChange={(e) => updateDutyData('Row4_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row4_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 5 */}
            <tr>
              <td style={styles.td}><strong>5</strong></td>
              <td style={styles.td}>
                <strong>If pt is not eligible call and inform</strong><br/>
                <input
                  type="text"
                  value={dutyData.Row5_CallNum}
                  onChange={(e) => updateDutyData('Row5_CallNum', e.target.value)}
                  placeholder="How many pt's did you call?"
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row5_Done}
                  onChange={(e) => updateDutyData('Row5_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row5_Checked}
                  onChange={(e) => updateDutyData('Row5_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row5_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 6 */}
            <tr style={isRowOverdue('Row6_Done') ? styles.overdueRow : {}}>
              <td style={styles.td}><strong>6</strong></td>
              <td style={styles.td}>
                <strong>Insurance breakdown for next day's patients</strong>
                <div style={styles.dutyDetails}>Call and get ins. info if necessary</div>
                <div style={styles.deadlineInfo}>Deadline: 4:30 PM</div>
                {isRowOverdue('Row6_Done') && (
                  <div style={styles.overdueWarning}>⚠️ Insurance breakdown should be done by 4:30 PM</div>
                )}
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row6_Done}
                  onChange={(e) => updateDutyData('Row6_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row6_Checked}
                  onChange={(e) => updateDutyData('Row6_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row6_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 7 */}
            <tr style={isRowOverdue('Row7_Done') ? styles.overdueRow : {}}>
              <td style={styles.td}><strong>7</strong></td>
              <td style={styles.td}>
                <strong>Check ledger for any balance on the account</strong><br/>
                <div style={styles.dutyDetails}>
                  Fill out "Account with Balances Form" and fax to the AR Department at (661)328-1905
                </div>
                <div style={styles.dutyDetails}>Called to inform patient of balance?</div>
                <div style={styles.inlineOption}>
                  <input
                    type="radio"
                    id="Row7_YesNo_Yes"
                    name="Row7_YesNo"
                    value="Yes"
                    checked={dutyData.Row7_YesNo === 'Yes'}
                    onChange={(e) => updateDutyData('Row7_YesNo', e.target.value)}
                    style={{ marginRight: '5px', cursor: 'pointer' }}
                  />
                  <label htmlFor="Row7_YesNo_Yes">Yes</label>
                </div>
                <div style={styles.inlineOption}>
                  <input
                    type="radio"
                    id="Row7_YesNo_No"
                    name="Row7_YesNo"
                    value="No"
                    checked={dutyData.Row7_YesNo === 'No'}
                    onChange={(e) => updateDutyData('Row7_YesNo', e.target.value)}
                    style={{ marginRight: '5px', cursor: 'pointer' }}
                  />
                  <label htmlFor="Row7_YesNo_No">No</label>
                </div>
                <div style={styles.deadlineInfo}>Deadline: 12:00 PM</div>
                {isRowOverdue('Row7_Done') && (
                  <div style={styles.overdueWarning}>⚠️ Check ledger for balance should be done by 12:00 PM</div>
                )}
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row7_Done}
                  onChange={(e) => updateDutyData('Row7_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row7_Checked}
                  onChange={(e) => updateDutyData('Row7_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row7_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 8 */}
            <tr style={isRowOverdue('Row8_Done') ? styles.overdueRow : {}}>
              <td style={styles.td}><strong>8</strong></td>
              <td style={styles.td}>
                <strong>Morning confirmations</strong><br/>
                <div style={styles.dutyDetails}>At least by noon</div>
                <div style={styles.deadlineInfo}>Deadline: 12:00 PM</div>
                {isRowOverdue('Row8_Done') && (
                  <div style={styles.overdueWarning}>⚠️ Morning confirmations should be done by 12:00 PM</div>
                )}
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row8_Done}
                  onChange={(e) => updateDutyData('Row8_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row8_Checked}
                  onChange={(e) => updateDutyData('Row8_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row8_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 9 */}
            <tr>
              <td style={styles.td}><strong>9</strong></td>
              <td style={styles.td}>
                <strong>No shows entered on ledger</strong>
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row9_Done}
                  onChange={(e) => updateDutyData('Row9_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row9_Checked}
                  onChange={(e) => updateDutyData('Row9_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row9_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 10 */}
            <tr>
              <td style={styles.td}><strong>10</strong></td>
              <td style={styles.td}>
                <strong>No shows stamped in patient charts</strong>
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row10_Done}
                  onChange={(e) => updateDutyData('Row10_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row10_Checked}
                  onChange={(e) => updateDutyData('Row10_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row10_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 11 */}
            <tr style={isRowOverdue('Row11_Done') ? styles.overdueRow : {}}>
              <td style={styles.td}><strong>11</strong></td>
              <td style={styles.td}>
                <strong>Reconfirming completed?</strong><br/>
                <div style={styles.dutyDetails}>Start at 4:00pm</div>
                <div style={styles.deadlineInfo}>Deadline: 4:30 PM</div>
                {isRowOverdue('Row11_Done') && (
                  <div style={styles.overdueWarning}>⚠️ Reconfirming should be completed by 4:30 PM</div>
                )}
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row11_Done}
                  onChange={(e) => updateDutyData('Row11_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row11_Checked}
                  onChange={(e) => updateDutyData('Row11_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row11_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 12 */}
            <tr style={isRowOverdue('Row12_Done') ? styles.overdueRow : {}}>
              <td style={styles.td}><strong>12</strong></td>
              <td style={styles.td}>
                <strong>One week reminders completed?</strong>
                <div style={styles.deadlineInfo}>Deadline: 3:00 PM</div>
                {isRowOverdue('Row12_Done') && (
                  <div style={styles.overdueWarning}>⚠️ One week reminders should be done by 3:00 PM</div>
                )}
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row12_Done}
                  onChange={(e) => updateDutyData('Row12_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row12_Checked}
                  onChange={(e) => updateDutyData('Row12_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row12_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 13 */}
            <tr>
              <td style={styles.td}><strong>13</strong></td>
              <td style={styles.td}>
                <strong>Call all treatment patients from today for post op</strong>
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row13_Done}
                  onChange={(e) => updateDutyData('Row13_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row13_Checked}
                  onChange={(e) => updateDutyData('Row13_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row13_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 14 */}
            <tr>
              <td style={styles.td}><strong>14</strong></td>
              <td style={styles.td}>
                <strong>Total lab case deposits/deliveries</strong>
                <textarea
                  value={dutyData['Row14_Name/DOB']}
                  onChange={(e) => updateDutyData('Row14_Name/DOB', e.target.value)}
                  placeholder="Name/DOB(mm/dd/yyyy) 1&#10;Name/DOB(mm/dd/yyyy) 2&#10;Name/DOB(mm/dd/yyyy) 3"
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    minHeight: '60px',
                    resize: 'vertical'
                  }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row14_Done}
                  onChange={(e) => updateDutyData('Row14_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row14_Checked}
                  onChange={(e) => updateDutyData('Row14_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row14_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 15 */}
            <tr>
              <td style={styles.td}><strong>15</strong></td>
              <td style={styles.td}>
                <strong>Check all undelivered lab cases and make appointments</strong>
                <div style={styles.dutyDetails}>
                  Any Lab case that is more than 3 weeks old must be sent to corporate along with $20 deposit
                </div>
                <textarea
                  value={dutyData.Row15_LabCases}
                  onChange={(e) => updateDutyData('Row15_LabCases', e.target.value)}
                  placeholder="1)&#10;2)&#10;3)"
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    minHeight: '60px',
                    resize: 'vertical'
                  }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row15_Done}
                  onChange={(e) => updateDutyData('Row15_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row15_Checked}
                  onChange={(e) => updateDutyData('Row15_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row15_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 16 */}
            <tr>
              <td style={styles.td}><strong>16</strong></td>
              <td style={styles.td}>
                <strong>Check all lab cases for next day</strong>
                <div style={styles.dutyDetails}>Call lab for next day pick up's</div>
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row16_Done}
                  onChange={(e) => updateDutyData('Row16_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row16_Checked}
                  onChange={(e) => updateDutyData('Row16_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row16_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 17 */}
            <tr>
              <td style={styles.td}><strong>17</strong></td>
              <td style={styles.td}>
                <strong>N₂O/ Compressor Off</strong>
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row17_Done}
                  onChange={(e) => updateDutyData('Row17_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row17_Checked}
                  onChange={(e) => updateDutyData('Row17_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row17_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 18 */}
            <tr>
              <td style={styles.td}><strong>18</strong></td>
              <td style={styles.td}>
                <strong>Did you read the meter on the Oxygen/N₂O/Helium tank?</strong><br/>
                <div style={styles.inlineOption}>
                  <input
                    type="radio"
                    id="Row18_YesNo_Yes"
                    name="Row18_YesNo"
                    value="Yes"
                    checked={dutyData.Row18_YesNo === 'Yes'}
                    onChange={(e) => updateDutyData('Row18_YesNo', e.target.value)}
                    style={{ marginRight: '5px', cursor: 'pointer' }}
                  />
                  <label htmlFor="Row18_YesNo_Yes">Yes</label>
                </div>
                <div style={styles.inlineOption}>
                  <input
                    type="radio"
                    id="Row18_YesNo_No"
                    name="Row18_YesNo"
                    value="No"
                    checked={dutyData.Row18_YesNo === 'No'}
                    onChange={(e) => updateDutyData('Row18_YesNo', e.target.value)}
                    style={{ marginRight: '5px', cursor: 'pointer' }}
                  />
                  <label htmlFor="Row18_YesNo_No">No</label>
                </div>
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row18_Done}
                  onChange={(e) => updateDutyData('Row18_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row18_Checked}
                  onChange={(e) => updateDutyData('Row18_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row18_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 19 */}
            <tr>
              <td style={styles.td}><strong>19</strong></td>
              <td style={styles.td}>
                <strong>How many tanks are empty & need to be replaced?</strong>
                <div style={styles.inlineOption}>
                  <label htmlFor="Row19_O2">O₂:</label>
                  <input
                    type="number"
                    id="Row19_O2"
                    value={dutyData.Row19_O2}
                    onChange={(e) => updateDutyData('Row19_O2', e.target.value)}
                    style={{ width: '50px', marginLeft: '5px' }}
                  />
                </div>
                <div style={styles.inlineOption}>
                  <label htmlFor="Row19_N2O">N₂O:</label>
                  <input
                    type="number"
                    id="Row19_N2O"
                    value={dutyData.Row19_N2O}
                    onChange={(e) => updateDutyData('Row19_N2O', e.target.value)}
                    style={{ width: '50px', marginLeft: '5px' }}
                  />
                </div>
                <div style={styles.inlineOption}>
                  <label htmlFor="Row19_He">He:</label>
                  <input
                    type="number"
                    id="Row19_He"
                    value={dutyData.Row19_He}
                    onChange={(e) => updateDutyData('Row19_He', e.target.value)}
                    style={{ width: '50px', marginLeft: '5px' }}
                  />
                </div>
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row19_Done}
                  onChange={(e) => updateDutyData('Row19_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row19_Checked}
                  onChange={(e) => updateDutyData('Row19_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row19_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 20 */}
            <tr>
              <td style={styles.td}><strong>20</strong></td>
              <td style={styles.td}>
                <strong>Check restrooms initial logs hourly</strong>
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row20_Done}
                  onChange={(e) => updateDutyData('Row20_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row20_Checked}
                  onChange={(e) => updateDutyData('Row20_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row20_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 21 */}
            <tr>
              <td style={styles.td}><strong>21</strong></td>
              <td style={styles.td}>
                <strong>Swept/Mopped</strong><br/>
                <div style={styles.inlineOption}>
                  <input
                    type="radio"
                    id="Row21_YesNo_Yes"
                    name="Row21_YesNo"
                    value="Yes"
                    checked={dutyData.Row21_YesNo === 'Yes'}
                    onChange={(e) => updateDutyData('Row21_YesNo', e.target.value)}
                    style={{ marginRight: '5px', cursor: 'pointer' }}
                  />
                  <label htmlFor="Row21_YesNo_Yes">Yes</label>
                </div>
                <div style={styles.inlineOption}>
                  <input
                    type="radio"
                    id="Row21_YesNo_No"
                    name="Row21_YesNo"
                    value="No"
                    checked={dutyData.Row21_YesNo === 'No'}
                    onChange={(e) => updateDutyData('Row21_YesNo', e.target.value)}
                    style={{ marginRight: '5px', cursor: 'pointer' }}
                  />
                  <label htmlFor="Row21_YesNo_No">No</label>
                </div>
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row21_Done}
                  onChange={(e) => updateDutyData('Row21_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row21_Checked}
                  onChange={(e) => updateDutyData('Row21_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row21_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 22 */}
            <tr>
              <td style={styles.td}><strong>22</strong></td>
              <td style={styles.td}>
                <strong>Cleaned Breakroom</strong><br/>
                <div style={styles.inlineOption}>
                  <input
                    type="radio"
                    id="Row22_YesNo_Yes"
                    name="Row22_YesNo"
                    value="Yes"
                    checked={dutyData.Row22_YesNo === 'Yes'}
                    onChange={(e) => updateDutyData('Row22_YesNo', e.target.value)}
                    style={{ marginRight: '5px', cursor: 'pointer' }}
                  />
                  <label htmlFor="Row22_YesNo_Yes">Yes</label>
                </div>
                <div style={styles.inlineOption}>
                  <input
                    type="radio"
                    id="Row22_YesNo_No"
                    name="Row22_YesNo"
                    value="No"
                    checked={dutyData.Row22_YesNo === 'No'}
                    onChange={(e) => updateDutyData('Row22_YesNo', e.target.value)}
                    style={{ marginRight: '5px', cursor: 'pointer' }}
                  />
                  <label htmlFor="Row22_YesNo_No">No</label>
                </div>
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row22_Done}
                  onChange={(e) => updateDutyData('Row22_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row22_Checked}
                  onChange={(e) => updateDutyData('Row22_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row22_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 23 */}
            <tr>
              <td style={styles.td}><strong>23</strong></td>
              <td style={styles.td}>
                <strong>Sterilizers: Cycle Complete</strong><br/>
                <div style={styles.inlineOption}>
                  <input
                    type="radio"
                    id="Row23_YesNo_Yes"
                    name="Row23_YesNo"
                    value="Yes"
                    checked={dutyData.Row23_YesNo === 'Yes'}
                    onChange={(e) => updateDutyData('Row23_YesNo', e.target.value)}
                    style={{ marginRight: '5px', cursor: 'pointer' }}
                  />
                  <label htmlFor="Row23_YesNo_Yes">Yes</label>
                </div>
                <div style={styles.inlineOption}>
                  <input
                    type="radio"
                    id="Row23_YesNo_No"
                    name="Row23_YesNo"
                    value="No"
                    checked={dutyData.Row23_YesNo === 'No'}
                    onChange={(e) => updateDutyData('Row23_YesNo', e.target.value)}
                    style={{ marginRight: '5px', cursor: 'pointer' }}
                  />
                  <label htmlFor="Row23_YesNo_No">No (Do Not Push Stop)</label>
                </div>
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row23_Done}
                  onChange={(e) => updateDutyData('Row23_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row23_Checked}
                  onChange={(e) => updateDutyData('Row23_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row23_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 24 */}
            <tr>
              <td style={styles.td}><strong>24</strong></td>
              <td style={styles.td}>
                <strong>Drained Ultrasonic</strong><br/>
                <div style={styles.inlineOption}>
                  <input
                    type="radio"
                    id="Row24_YesNo_Yes"
                    name="Row24_YesNo"
                    value="Yes"
                    checked={dutyData.Row24_YesNo === 'Yes'}
                    onChange={(e) => updateDutyData('Row24_YesNo', e.target.value)}
                    style={{ marginRight: '5px', cursor: 'pointer' }}
                  />
                  <label htmlFor="Row24_YesNo_Yes">Yes</label>
                </div>
                <div style={styles.inlineOption}>
                  <input
                    type="radio"
                    id="Row24_YesNo_No"
                    name="Row24_YesNo"
                    value="No"
                    checked={dutyData.Row24_YesNo === 'No'}
                    onChange={(e) => updateDutyData('Row24_YesNo', e.target.value)}
                    style={{ marginRight: '5px', cursor: 'pointer' }}
                  />
                  <label htmlFor="Row24_YesNo_No">No</label>
                </div>
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row24_Done}
                  onChange={(e) => updateDutyData('Row24_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row24_Checked}
                  onChange={(e) => updateDutyData('Row24_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row24_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 25 */}
            <tr style={isRowOverdue('Row25_Done') ? styles.overdueRow : {}}>
              <td style={styles.td}><strong>25</strong></td>
              <td style={styles.td}>
                <strong>Spore Test</strong>
                <div style={styles.dutyDetails}>Every Monday</div>
                <div style={styles.deadlineInfo}>Deadline: 11:00 AM (Mondays only)</div>
                {isRowOverdue('Row25_Done') && (
                  <div style={styles.overdueWarning}>⚠️ Spore Test should be done by 11:00 AM</div>
                )}
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row25_Done}
                  onChange={(e) => updateDutyData('Row25_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row25_Checked}
                  onChange={(e) => updateDutyData('Row25_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row25_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 26 */}
            <tr>
              <td style={styles.td}><strong>26</strong></td>
              <td style={styles.td}>
                <strong>Turn Off All TV's and Computers at the End of the Day</strong>
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row26_Done}
                  onChange={(e) => updateDutyData('Row26_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row26_Checked}
                  onChange={(e) => updateDutyData('Row26_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row26_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 27 */}
            <tr>
              <td style={styles.td}><strong>27</strong></td>
              <td style={styles.td}>
                <strong>Postcards Ready for Pick-up</strong>
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row27_Done}
                  onChange={(e) => updateDutyData('Row27_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row27_Checked}
                  onChange={(e) => updateDutyData('Row27_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row27_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 28 */}
            <tr>
              <td style={styles.td}><strong>28</strong></td>
              <td style={styles.td}>
                <strong>Clean traps everyday</strong> (chair)
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row28_Done}
                  onChange={(e) => updateDutyData('Row28_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row28_Checked}
                  onChange={(e) => updateDutyData('Row28_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row28_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 29 */}
            <tr>
              <td style={styles.td}><strong>29</strong></td>
              <td style={styles.td}>
                <strong>Clean main trap 1st/15th</strong> (by vacuum)
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row29_Done}
                  onChange={(e) => updateDutyData('Row29_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row29_Checked}
                  onChange={(e) => updateDutyData('Row29_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row29_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 30 */}
            <tr>
              <td style={styles.td}><strong>30</strong></td>
              <td style={styles.td}>
                <strong>Did you flush the lines with hot water?</strong><br/>
                <div style={styles.inlineOption}>
                  <input
                    type="radio"
                    id="Row30_YesNo_Yes"
                    name="Row30_YesNo"
                    value="Yes"
                    checked={dutyData.Row30_YesNo === 'Yes'}
                    onChange={(e) => updateDutyData('Row30_YesNo', e.target.value)}
                    style={{ marginRight: '5px', cursor: 'pointer' }}
                  />
                  <label htmlFor="Row30_YesNo_Yes">Yes</label>
                </div>
                <div style={styles.inlineOption}>
                  <input
                    type="radio"
                    id="Row30_YesNo_No"
                    name="Row30_YesNo"
                    value="No"
                    checked={dutyData.Row30_YesNo === 'No'}
                    onChange={(e) => updateDutyData('Row30_YesNo', e.target.value)}
                    style={{ marginRight: '5px', cursor: 'pointer' }}
                  />
                  <label htmlFor="Row30_YesNo_No">No</label>
                </div>
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row30_Done}
                  onChange={(e) => updateDutyData('Row30_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row30_Checked}
                  onChange={(e) => updateDutyData('Row30_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row30_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 31 */}
            <tr>
              <td style={styles.td}><strong>31</strong></td>
              <td style={styles.td}>
                <strong>Check all doors are locked</strong>
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row31_Done}
                  onChange={(e) => updateDutyData('Row31_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row31_Checked}
                  onChange={(e) => updateDutyData('Row31_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row31_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>

            {/* Row 32 - 마지막 행 */}
            <tr>
              <td style={styles.td}><strong>32</strong></td>
              <td style={styles.td}>
                <strong>Turn On Answering Service</strong><br/>
                <div style={styles.dutyDetails}>
                  1) Go to: https://smileland.my3cx.us<br/>
                  2) Log in<br/>
                  3) Click on User Icon<br/>
                  4) Select Override Office Hours<br/>
                  5) Select "Office is Closed"<br/>
                  6) For "1day"<br/>
                  7) Call the office to verify that calls were transferred correctly. (For Holiday Weekends, Set for 1 week)
                </div>
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row32_Done}
                  onChange={(e) => updateDutyData('Row32_Done', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row32_Checked}
                  onChange={(e) => updateDutyData('Row32_Checked', e.target.value)}
                  style={{ ...styles.input, margin: 0, fontSize: '13px', padding: '6px 8px' }}
                />
              </td>
              <td style={styles.td}>
                <input
                  type="text"
                  value={dutyData.Row32_Time}
                  readOnly
                  style={{ 
                    ...styles.input, 
                    margin: 0, 
                    fontSize: '13px', 
                    padding: '6px 8px',
                    backgroundColor: '#f8f9fa',
                    color: '#6c757d'
                  }}
                />
              </td>
            </tr>
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

      {/* 비밀번호 확인 모달 */}
      {showPasswordModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          backgroundColor: 'rgba(0, 0, 0, 0.5)',
          display: 'flex',
          justifyContent: 'center',
          alignItems: 'center',
          zIndex: 1000
        }}>
          <div style={{
            backgroundColor: 'white',
            padding: '30px',
            borderRadius: '8px',
            maxWidth: '400px',
            width: '90%',
            boxShadow: '0 4px 6px rgba(0, 0, 0, 0.1)'
          }}>
            <h3 style={{
              marginTop: 0,
              marginBottom: '20px',
              color: '#2c3e50',
              fontSize: '18px'
            }}>
              Office Password Required
            </h3>
            <p style={{
              marginBottom: '20px',
              color: '#666',
              fontSize: '14px'
            }}>
              Enter the password for <strong>{pendingOffice}</strong>
            </p>
            <input
              type="password"
              id="officePasswordInput"
              autoFocus
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  const input = e.target as HTMLInputElement;
                  const password = input.value;
                  const expectedPassword = getOfficePassword(pendingOffice);
                  
                  if (password === expectedPassword) {
                    setSelectedOffice(pendingOffice);
                    setOfficePasswordVerified(true);
                    setShowPasswordModal(false);
                    setPendingOffice('');
                  } else {
                    alert("Incorrect password. Access denied.");
                    input.value = '';
                  }
                }
              }}
              style={{
                width: '100%',
                padding: '10px',
                fontSize: '16px',
                border: '1px solid #e9ecef',
                borderRadius: '4px',
                marginBottom: '20px',
                boxSizing: 'border-box'
              }}
              placeholder="Enter password"
            />
            <div style={{
              display: 'flex',
              gap: '10px',
              justifyContent: 'flex-end'
            }}>
              <button
                onClick={() => {
                  setShowPasswordModal(false);
                  setPendingOffice('');
                  // 선택을 취소하고 드롭다운을 빈 값으로 복원
                  const selectElement = document.getElementById('selectedOffice') as HTMLSelectElement;
                  if (selectElement) {
                    selectElement.value = selectedOffice || '';
                  }
                }}
                style={{
                  padding: '10px 20px',
                  backgroundColor: '#95a5a6',
                  color: 'white',
                  border: 'none',
                  borderRadius: '4px',
                  cursor: 'pointer',
                  fontSize: '14px'
                }}
              >
                Cancel
              </button>
              <button
                onClick={() => {
                  const input = document.getElementById('officePasswordInput') as HTMLInputElement;
                  const password = input?.value || '';
                  const expectedPassword = getOfficePassword(pendingOffice);
                  
                  if (password === expectedPassword) {
                    setSelectedOffice(pendingOffice);
                    setOfficePasswordVerified(true);
                    setShowPasswordModal(false);
                    setPendingOffice('');
                  } else {
                    alert("Incorrect password. Access denied.");
                    if (input) input.value = '';
                  }
                }}
                style={{
                  padding: '10px 20px',
                  backgroundColor: '#3498db',
                  color: 'white',
                  border: 'none',
                  borderRadius: '4px',
                  cursor: 'pointer',
                  fontSize: '14px'
                }}
              >
                Verify
              </button>
            </div>
          </div>
        </div>
      )}
      </div>
    </div>
    </>
  );
}
