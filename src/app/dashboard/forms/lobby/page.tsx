'use client';

import React, { useState, useEffect, useCallback } from 'react';
import { db } from '@/lib/firebase.config';
import { doc, setDoc, getDoc, onSnapshot, deleteDoc } from 'firebase/firestore';
import { getStorage, ref, uploadBytes, getDownloadURL } from 'firebase/storage';

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

  // 오피스 옵션 (알파벳 순)
  const officeOptions = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

  // Office별 Check 옵션 정의 (알파벳 순)
  const OFFICE_CHECK_OPTIONS = {
    "Bernard": ["Carmen", "Cynthia", "Elisa", "Ranjit"],
    "California": ["Kindal"],
    "Delano": ["Helen", "Stephanie"],
    "Fresno": ["Cynthia"],
    "Ming": ["Kindal"],
    "Ortho": ["Kindal"],
    "Tulare": ["Dianne"],
    "Visalia": ["Dianne"]
  };

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
      console.error("Auto-save error:", error);
    }
  }, [inspectionDate, selectedOffice, lobbyData, lastSavedData, isUpdatingFromFirebase]);

  // 데이터 변경 시에만 자동 저장
  useEffect(() => {
    // 초기 로드 시에는 저장하지 않음
    if (Object.values(lobbyData).some(value => value !== '')) {
      autoSave();
    }
  }, [lobbyData]);

  // 데이터 로드 함수
  const loadData = async () => {
    if (!inspectionDate || !selectedOffice) return;

    try {
      console.log("Loading data for date:", inspectionDate, "office:", selectedOffice);
      setSubmitStatus('Loading data...');
      
      const docId = `${inspectionDate}_${selectedOffice}_lobby`;
      const docSnap = await getDoc(doc(db, "lobby-inspections", docId));
      
      if (docSnap.exists()) {
        const data = docSnap.data();
        console.log("Data loaded from Firebase:", data);
        
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
        console.log("No data found for date:", inspectionDate);
        // 초기 데이터 설정
        const initialData: any = {};
        setLobbyData(initialData);
        setLastSavedData(initialData);
        setSubmitStatus('No data found - initialized empty form');
        setTimeout(() => setSubmitStatus(''), 2000);
      }
      
    } catch (error: any) {
      console.error("Error loading data:", error);
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

    console.log("Setting up real-time listener for date:", inspectionDate, "office:", selectedOffice);
    const docId = `${inspectionDate}_${selectedOffice}_lobby`;
    const docRef = doc(db, "lobby-inspections", docId);
    
    const unsubscribe = onSnapshot(docRef, (docSnap: any) => {
      if (docSnap.exists()) {
        const data = docSnap.data();
        console.log("Real-time data received:", data);
        
        // 데이터가 실제로 변경되었는지 확인
        const hasChanges = JSON.stringify(data) !== JSON.stringify(lobbyData);
        if (!hasChanges) {
          console.log("No changes detected, skipping update");
          return;
        }
        
        // Firebase에서 업데이트되는 동안 자동 저장 방지
        setIsUpdatingFromFirebase(true);
        
        setLobbyData((prevData: any) => {
          console.log("Updating lobbyData from:", prevData, "to:", { ...prevData, ...data });
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
      } else {
        console.log("Real-time listener: No document exists for date:", inspectionDate);
      }
    }, (error: any) => {
      console.error("Real-time listener error:", error);
    });

    return () => {
      console.log("Cleaning up real-time listener for date:", inspectionDate, "office:", selectedOffice);
      unsubscribe();
    };
  }, [inspectionDate, selectedOffice]);

  // 데이터 업데이트 함수
  const updateLobbyData = (field: string, value: any) => {
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

  // Office 변경 처리
  const handleOfficeChange = (newOffice: string) => {
    // 빈 값으로 선택하면 비밀번호 없이 변경 허용 (초기화)
    if (newOffice === '') {
      setSelectedOffice('');
      return;
    }
    
    // 주의: 약한 비밀번호 검증 (보안 취약점)
    // 선택된 office의 첫 알파벳 대문자를 비밀번호로 사용
    // TODO: 강력한 비밀번호 정책 적용 또는 서버 사이드 인증으로 이동
    const officePassword = newOffice.charAt(0).toUpperCase();
    const password = prompt(`Enter password to change office: `);
    if (password === null) return;
    if (password !== officePassword) {
      alert("Incorrect password. Office change cancelled.");
      return;
    }
    setSelectedOffice(newOffice);
  };

  // 제출 함수
  const handleSubmit = async () => {
    if (!inspectionDate || !selectedOffice) {
      alert('Please select a date and office first.');
      return;
    }

    // 비밀번호는 서버 사이드에서 검증해야 함 (현재는 클라이언트 사이드 검증만)
    // TODO: 서버 사이드 인증으로 이동 필요
    const password = prompt(`Are you sure you want to submit? Submitting will reset today's data. Enter password to proceed:`);
    if (password === null) return;
    
    // 주의: 비밀번호가 클라이언트 코드에 노출됨 (보안 취약점)
    // 프로덕션에서는 서버 사이드 인증 API를 통해 검증해야 함
    if (password !== 'Halloween') {
      alert("Incorrect password. Submission cancelled.");
      return;
    }

    setLoading(true);
    setSubmitStatus('Saving...');
    setProgress(10);

    try {
      // 1. PDF 생성
      setSubmitStatus('Generating PDF...');
      setProgress(30);
      
      // PDF 생성 API 호출
      const response = await fetch('/api/generate-lobby-pdf', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          inspectionDate,
          selectedOffice,
          lobbyData
        }),
      });

      if (response.ok) {
        // PDF blob 받기
        setSubmitStatus('Processing PDF...');
        setProgress(60);
        const blob = await response.blob();
        
        // PDF를 Firebase Storage에 저장
        setSubmitStatus('Saving PDF...');
        setProgress(70);
        try {
          const storage = getStorage();
          const filename = `5) ${inspectionDate}_${selectedOffice}_Lobby Inspection Log.pdf`;
          const storageRef = ref(storage, `endofday-pdfs/${selectedOffice}/${inspectionDate}/${filename}`);
          
          // PDF 업로드
          await uploadBytes(storageRef, blob);
          
          // 다운로드 URL 가져오기
          const downloadUrl = await getDownloadURL(storageRef);
          
          // Firestore에 메타데이터 저장
          await setDoc(doc(db, 'pdf-documents', `${inspectionDate}_${selectedOffice}_lobby_${Date.now()}`), {
            filename,
            office: selectedOffice,
            date: inspectionDate,
            type: 'Lobby Inspection Log',
            url: downloadUrl,
            storagePath: `endofday-pdfs/${selectedOffice}/${inspectionDate}/${filename}`,
            createdAt: new Date(),
          });
          
          console.log('PDF saved successfully');
        } catch (storageError: any) {
          console.error('Storage error:', storageError);
          // 저장 실패 시 사용자에게 알림
          const errorMsg = storageError?.message || '알 수 없는 오류';
          alert(`PDF 저장 중 오류가 발생했습니다: ${errorMsg}`);
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
          alert('✅ Submitted successfully!');
        }, 2000);
      } else {
        let errorMessage = 'PDF generation failed';
        try {
          const errorData = await response.json();
          errorMessage = errorData.error || errorMessage;
        } catch (e) {
          errorMessage = `HTTP ${response.status}: ${response.statusText}`;
        }
        throw new Error(errorMessage);
      }

    } catch (error: any) {
      console.error('Submit error:', error);
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

  // 컴포넌트 마운트 시 오늘 날짜 설정 및 HTTPS 체크
  useEffect(() => {
    setInspectionDate(getCurrentCaliforniaTime());
    
    // 프로덕션 환경에서 HTTPS 강제 (클라이언트 사이드)
    if (process.env.NODE_ENV === 'production' && 
        typeof window !== 'undefined' && 
        window.location.protocol !== 'https:') {
      // HTTP로 접속한 경우 HTTPS로 리다이렉트
      window.location.href = window.location.href.replace('http:', 'https:');
    }
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
          <label style={styles.label} htmlFor="office">🏢 Office:</label>
          <select
            id="office"
            value={selectedOffice}
            onChange={(e: any) => handleOfficeChange(e.target.value)}
            style={styles.select}
          >
            <option value="">--Select Office--</option>
            {officeOptions.map(office => (
              <option key={office} value={office}>{office}</option>
            ))}
          </select>
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
