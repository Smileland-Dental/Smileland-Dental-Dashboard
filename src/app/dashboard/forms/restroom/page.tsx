'use client'

import React, { useState, useEffect, useCallback } from "react";
import { doc, setDoc, getDoc, deleteDoc, onSnapshot, collection, getDocs } from "firebase/firestore";
import { db } from "@/lib/firebase.config";
import { getStorage, ref, uploadBytes, getDownloadURL } from 'firebase/storage';

export default function RestroomInspection() {
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
  const [inspectionDate, setInspectionDate] = useState(() => {
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
    "Ming": ["Kindal"],
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
      setAutoSaveStatus('💾 Saving...');
      
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
      
      setAutoSaveStatus('💾 Saved ✅');
      setTimeout(() => setAutoSaveStatus(''), 2000);
      
    } catch (error) {
      console.error("Auto-save error:", error);
      setAutoSaveStatus('💾 Save failed ❌');
      setTimeout(() => setAutoSaveStatus(''), 3000);
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
      console.log("Loading data for date:", inspectionDate, "office:", selectedOffice, "restroom:", selectedRestroom);
      setSubmitStatus('Loading data...');
      
      const docId = `${inspectionDate}_${selectedOffice}_${selectedRestroom}`;
      const docSnap = await getDocs(collection(db, "restroom-inspections")).then((snapshot: any) => {
        const foundDoc = snapshot.docs.find((d: any) => d.id === docId);
        return foundDoc ? { exists: () => true, data: () => foundDoc.data() } : { exists: () => false, data: undefined };
      });
      
      if (docSnap.exists() && docSnap.data) {
        const data = docSnap.data();
        console.log("Data loaded from Firebase:", data);
        
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
        
        setSubmitStatus('Data loaded successfully');
        setTimeout(() => setSubmitStatus(''), 2000);
      } else {
        console.log("No data found for date:", inspectionDate);
        // 데이터가 없으면 초기화
        const initialData = {};
        setRestroomData(initialData);
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

  // 날짜, 오피스, 화장실 변경 시 데이터 로드
  useEffect(() => {
    if (selectedRestroom) {
      loadData();
    }
  }, [inspectionDate, selectedOffice, selectedRestroom]);

  // 실시간 데이터 동기화
  useEffect(() => {
    if (!inspectionDate || !selectedOffice || !selectedRestroom) return;

    console.log("Setting up real-time listener for date:", inspectionDate, "office:", selectedOffice, "restroom:", selectedRestroom);
    const docId = `${inspectionDate}_${selectedOffice}_${selectedRestroom}`;
    const docRef = doc(db, "restroom-inspections", docId);
    
    const unsubscribe = onSnapshot(docRef, (docSnap: any) => {
      if (docSnap.exists() && docSnap.data) {
        const data = docSnap.data();
        console.log("Real-time data received:", data);
        
        // Firebase에서 업데이트되는 동안 자동 저장 방지
        setIsUpdatingFromFirebase(true);
        
        setRestroomData((prevData: any) => {
          console.log("Updating restroomData from:", prevData, "to:", { ...prevData, ...data });
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
          setAutoSaveStatus('🔄 동시 저장됨 - 다른 사용자가 업데이트했습니다');
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
      console.log("Cleaning up real-time listener for date:", inspectionDate, "office:", selectedOffice, "restroom:", selectedRestroom);
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
    if (!selectedOffice) {
      alert('Please select an office first.');
      return;
    }

    const password = prompt(`Are you sure you want to submit? Submitting will reset today's data. Enter password to proceed:`);
    if (password === null) return;
    if (password !== 'Halloween') {
      alert("Incorrect password. Submission cancelled.");
      return;
    }

    try {
      setLoading(true);
      setSubmitStatus('Saving...');
      setProgress(10);

      // 1. PDF 생성
      setSubmitStatus('Generating PDF...');
      setProgress(30);
      
      const response = await fetch('/api/generate-restroom-pdf', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          inspectionDate,
          selectedOffice,
          selectedRestroom,
          restroomData
        }),
      });

      if (response.ok) {
        // PDF blob 받기
        setSubmitStatus('Processing PDF...');
        setProgress(60);
        const blob = await response.blob();
        
        // PDF를 Firebase Storage에 저장
        setSubmitStatus('Saving PDF to archive...');
        setProgress(70);
        try {
          const storage = getStorage();
          const filename = `6) ${inspectionDate}_${selectedOffice}_Restroom_${selectedRestroom}_Inspection_Log.pdf`;
          const storageRef = ref(storage, `endofday-pdfs/${selectedOffice}/${inspectionDate}/${filename}`);
          
          // PDF 업로드
          await uploadBytes(storageRef, blob);
          
          // 다운로드 URL 가져오기
          const downloadUrl = await getDownloadURL(storageRef);
          
          // Firestore에 메타데이터 저장
          await setDoc(doc(db, 'pdf-documents', `${inspectionDate}_${selectedOffice}_restroom_${selectedRestroom}_${Date.now()}`), {
            filename,
            office: selectedOffice,
            date: inspectionDate,
            type: `Restroom ${selectedRestroom} Inspection Log`,
            url: downloadUrl,
            storagePath: `endofday-pdfs/${selectedOffice}/${inspectionDate}/${filename}`,
            createdAt: new Date(),
          });
          
          console.log('PDF saved successfully to Firebase Storage');
        } catch (storageError: any) {
          console.error('Storage error:', storageError);
          // 저장 실패 시 사용자에게 알림
          const errorMsg = storageError?.message || '알 수 없는 오류';
          alert(`PDF 저장 중 오류가 발생했습니다: ${errorMsg}`);
        }
        
        // 2. 데이터 삭제
        setSubmitStatus('Cleaning up...');
        setProgress(80);
        const docId = `${inspectionDate}_${selectedOffice}_${selectedRestroom}`;
        await deleteDoc(doc(db, "restroom-inspections", docId));
        
        // 3. 폼 초기화
        setRestroomData({});

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
    autoSaveStatus: {
      position: 'fixed',
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
      textAlign: 'center',
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
              borderRadius: "12px",
              textAlign: "center",
              boxShadow: "0 10px 30px rgba(0, 0, 0, 0.3)",
              maxWidth: "400px",
              width: "90%"
            } as React.CSSProperties}>
              <div style={{
                border: "4px solid #f3f3f3",
                borderTop: "4px solid #4a90e2",
                borderRadius: "50%",
                width: "50px",
                height: "50px",
                animation: "spin 1s linear infinite",
                margin: "0 auto 20px"
              } as React.CSSProperties}></div>
              <h3 style={{
                color: "#333",
                fontSize: "1.2rem",
                fontWeight: "600",
                margin: "0 0 10px 0"
              } as React.CSSProperties}>
                {submitStatus || "Processing..."}
              </h3>
              <p style={{
                color: "#666",
                fontSize: "0.9rem",
                margin: "0 0 20px 0",
                lineHeight: "1.4"
              } as React.CSSProperties}>
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
              } as React.CSSProperties}>
                <div style={{
                  width: `${progress}%`,
                  height: "8px",
                  backgroundColor: "#4a90e2",
                  borderRadius: "10px",
                  transition: "width 0.3s ease",
                  background: "linear-gradient(90deg, #4a90e2, #357abd)"
                } as React.CSSProperties}></div>
              </div>
              <p style={{
                color: "#495057",
                fontSize: "0.8rem",
                margin: "0 0 20px 0",
                fontWeight: "500"
              } as React.CSSProperties}>
                {progress}% Complete
              </p>
              <div style={{
                backgroundColor: "#f8f9fa",
                padding: "15px",
                borderRadius: "8px",
                border: "1px solid #e9ecef"
              } as React.CSSProperties}>
                <p style={{
                  color: "#495057",
                  fontSize: "0.8rem",
                  margin: 0,
                  fontWeight: "500"
                } as React.CSSProperties}>
                  ⚠️ Please do not close.
                </p>
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
            } as React.CSSProperties}>
              {autoSaveStatus}
            </div>
          )}

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
            <label style={styles.label} htmlFor="restroom">🚻 Restroom:</label>
            <select
              id="restroom"
              value={selectedRestroom}
              onChange={(e: any) => setSelectedRestroom(e.target.value)}
              style={styles.select}
            >
              <option value="" disabled hidden>Select Restroom</option>
              {restroomOptions.map(restroom => (
                <option key={restroom} value={restroom}>{restroom}</option>
              ))}
            </select>
          </div>

          {/* 검사 테이블 */}
          {selectedRestroom && (
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
                        {COLUMN_NAMES.slice(2, -1).map((columnName, colIndex) => {
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
          {selectedRestroom && (
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
                🚀 Submit Report
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

