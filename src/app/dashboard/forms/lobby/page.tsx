'use client';

import React, { useState, useEffect, useCallback } from 'react';
import { db } from '@/lib/firebase.config';
import { doc, setDoc, getDoc, onSnapshot, collection, getDocs, deleteDoc } from 'firebase/firestore';

export default function LobbyInspectionPage() {
  // 상태 관리
  const [inspectionDate, setInspectionDate] = useState('');
  const [selectedOffice, setSelectedOffice] = useState('Bernard');
  const [lobbyData, setLobbyData] = useState<any>({});
  const [loading, setLoading] = useState(false);
  const [autoSaveStatus, setAutoSaveStatus] = useState('');
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
    const laTime = new Date(now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
    const yyyy = laTime.getFullYear();
    const mm = String(laTime.getMonth() + 1).padStart(2, '0');
    const dd = String(laTime.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
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
      setAutoSaveStatus('💾 Saving...');
      
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
      
      setAutoSaveStatus('💾 Saved ✅');
      setTimeout(() => setAutoSaveStatus(''), 2000);
      
    } catch (error) {
      console.error("Auto-save error:", error);
      setAutoSaveStatus('💾 Save failed ❌');
      setTimeout(() => setAutoSaveStatus(''), 3000);
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
        
        // 다른 사용자의 업데이트만 표시 (자신의 업데이트는 제외)
        if (data.timestamp && 
            new Date(data.timestamp).getTime() > Date.now() - 5000 && 
            data.lastUpdatedBy && 
            data.lastUpdatedBy !== userSessionId) {
          setAutoSaveStatus('🔄 Updated by another user.');
          setTimeout(() => setAutoSaveStatus(''), 3000);
        }
      } else {
        console.log("Real-time listener: No document exists for date:", inspectionDate);
      }
    }, (error: any) => {
      console.error("Real-time listener error:", error);
      setAutoSaveStatus('❌ 연결 오류 - 실시간 동기화가 중단되었습니다');
      setTimeout(() => setAutoSaveStatus(''), 3000);
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

  // 제출 함수
  const handleSubmit = async () => {
    if (!inspectionDate || !selectedOffice) {
      alert('Please select a date and office first.');
      return;
    }

    // Office 기반 비밀번호 검증
    const expectedPassword = selectedOffice.toLowerCase();
    const password = prompt(`Are you sure you want to submit? Submitting will reset today's data. Enter password to proceed (Password: ${expectedPassword}):`);
    if (password === null) return;
    if (password !== expectedPassword) {
      alert("Incorrect password. Submission cancelled.");
      return;
    }

    // 날짜를 요일로 변환
    const dateObj = new Date(inspectionDate + 'T00:00:00');
    const laDate = new Date(dateObj.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
    const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const dayName = days[laDate.getDay()];

    // 날짜 포맷팅 (MM/DD/YYYY)
    const month = String(laDate.getMonth() + 1).padStart(2, '0');
    const day = String(laDate.getDate()).padStart(2, '0');
    const year = laDate.getFullYear();
    const formattedDate = `${month}/${day}/${year}`;

    // 확인 메시지
    const confirmMessage = `Are you sure you want to submit the ${formattedDate} ${dayName} Lobby Inspection Log?\n\nSubmitting will reset all data entered today.`;

    if (!confirm(confirmMessage)) return;

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
        
        // 4. PDF 다운로드
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `5) ${inspectionDate}_${selectedOffice}_Lobby Inspection Log.pdf`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        
        setTimeout(() => {
          setLoading(false);
          setSubmitStatus('');
          setProgress(0);
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
      }, 3000);
    }
  };

  // 스타일 정의
  const styles = {
    body: {
      fontFamily: "'Roboto', sans-serif",
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
    autoSaveStatus: {
      position: 'fixed' as const,
      top: '20px',
      right: '20px',
      padding: '12px 20px',
      backgroundColor: '#51cf66',
      color: 'white',
      borderRadius: '25px',
      fontSize: '14px',
      fontWeight: 'bold',
      boxShadow: '0 4px 12px rgba(0,0,0,0.3)',
      zIndex: 1000,
      maxWidth: '300px',
      textAlign: 'center' as const,
      backdropFilter: 'blur(10px)',
      border: '1px solid rgba(255, 255, 255, 0.2)'
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

  // 컴포넌트 마운트 시 오늘 날짜 설정
  useEffect(() => {
    setInspectionDate(getCurrentCaliforniaTime());
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
        @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap');
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
        {/* 자동 저장 상태 표시 */}
        {autoSaveStatus && (
          <div style={{
            ...styles.autoSaveStatus,
            backgroundColor: autoSaveStatus.includes('❌') ? '#ff6b6b' : 
                            autoSaveStatus.includes('🔄') ? '#4a90e2' : 
                            autoSaveStatus.includes('💾') ? '#51cf66' : '#51cf66'
          }}>
            {autoSaveStatus}
          </div>
        )}

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
            onChange={(e: any) => setSelectedOffice(e.target.value)}
            style={styles.select}
          >
            {officeOptions.map(office => (
              <option key={office} value={office}>{office}</option>
            ))}
          </select>
        </div>

        {/* 안내 메시지 - Date나 Office가 선택되지 않았을 때 */}
        {(!inspectionDate || !selectedOffice) && (
          <div style={{
            textAlign: 'center',
            padding: '40px 20px',
            background: 'linear-gradient(135deg, rgba(255, 154, 158, 0.1) 0%, rgba(254, 207, 239, 0.1) 100%)',
            borderRadius: '15px',
            border: '1px solid rgba(255, 154, 158, 0.2)',
            margin: '20px 0'
          }}>
            <div style={{fontSize: '18px', color: '#4a6fa1', fontWeight: '600', marginBottom: '10px'}}>
              📋 Please select Date and Office to view the inspection form
            </div>
            <div style={{fontSize: '14px', color: '#666'}}>
              {!inspectionDate && !selectedOffice && 'Select both Date and Office above to start'}
              {!inspectionDate && selectedOffice && 'Please select a Date to continue'}
              {inspectionDate && !selectedOffice && 'Please select an Office to continue'}
            </div>
          </div>
        )}

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
              style={styles.submitButton}
              disabled={loading}
            >
              🚀 Submit Report
            </button>
          </div>
        )}
      </div>
      </div>
    </>
  );
}
