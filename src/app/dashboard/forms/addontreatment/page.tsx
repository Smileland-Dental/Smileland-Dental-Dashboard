'use client'

import React, { useState, useEffect, useCallback } from "react";
import { doc, setDoc, getDoc, deleteDoc, onSnapshot, collection, getDocs } from "firebase/firestore";
import { db } from "@/lib/firebase.config";
import { enableAllSecurityMeasures, sanitizeFirebaseDataClient, sanitizeFirebaseDocIdClient } from "@/lib/security-client";

export default function AddOnTreatment() {
  // 보안 조치 활성화
  useEffect(() => {
    enableAllSecurityMeasures();
  }, []);

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
    const californiaTime = new Date(now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
    return californiaTime.toISOString().split('T')[0];
  });

  // 오피스 선택 상태
  const [selectedOffice, setSelectedOffice] = useState('Bernard');
  
  // 오피스 옵션
  const officeOptions = ['Ming', 'Bernard', 'Delano', 'Tulare', 'Visalia', 'Fresno', 'California', 'Ortho'];

  // 환자 데이터 상태 - 동적으로 관리
  const [patientData, setPatientData] = useState<Record<string, string>>({});
  const [rowCount, setRowCount] = useState(20);

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

  // 시간을 12시간 형식으로 변환하는 함수
  const convertTo12Hour = (timeStr: string): string => {
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
  };

  // 행 추가 함수
  const addRow = () => {
    setRowCount(prev => prev + 1);
  };

  // 행 클리어 함수
  const clearRow = (rowNumber: number) => {
    setPatientData(prev => {
      const newData = { ...prev };
      newData[`Row${rowNumber}_PatientName`] = '';
      newData[`Row${rowNumber}_DOB`] = '';
      newData[`Row${rowNumber}_Time`] = '';
      return newData;
    });
  };

  // 자동 저장 함수
  const autoSave = useCallback(async () => {
    if (!dutyDate || !selectedOffice || isUpdatingFromFirebase) return;

    // 데이터가 실제로 변경되었는지 확인
    const currentData = { ...patientData, rowCount, selectedOffice };
    const hasChanges = JSON.stringify(currentData) !== JSON.stringify(lastSavedData);
    if (!hasChanges) return;

    try {
      const dataToSave = {
        dutyDate,
        selectedOffice,
        ...patientData,
        rowCount,
        timestamp: new Date().toISOString(),
        autoSaved: true,
        lastUpdatedBy: userSessionId
      };

      const docId = sanitizeFirebaseDocIdClient(`${dutyDate}_${selectedOffice}_addon_treatment`);
      const safeDataToSave = sanitizeFirebaseDataClient(dataToSave);
      await setDoc(doc(db, "addon-treatment", docId), safeDataToSave);
      
      // 저장 성공 시 마지막 저장된 데이터 업데이트
      setLastSavedData(currentData);
      
    } catch (error) {
      // Production에서는 에러 로깅 비활성화
      if (process.env.NODE_ENV !== 'production') {
        console.error("Auto-save error:", error);
      }
    }
  }, [dutyDate, selectedOffice, patientData, rowCount, lastSavedData, isUpdatingFromFirebase, userSessionId]);

  // 데이터 변경 시에만 자동 저장
  useEffect(() => {
    // 초기 로드 시에는 저장하지 않음
    if (Object.values(patientData).some(value => value !== '')) {
      autoSave();
    }
  }, [patientData]);

  // 데이터 로드
  const loadData = async () => {
    if (!dutyDate || !selectedOffice) return;

    try {
      // Production에서는 로그 출력 안 함
      if (process.env.NODE_ENV !== 'production') {
        console.log("Loading data for date:", dutyDate, "office:", selectedOffice);
      }
      setSubmitStatus('Loading data...');
      
      const docId = sanitizeFirebaseDocIdClient(`${dutyDate}_${selectedOffice}_addon_treatment`);
      const docSnap = await getDocs(collection(db, "addon-treatment")).then(snapshot => {
        const foundDoc = snapshot.docs.find(d => d.id === docId);
        return foundDoc ? { 
          exists: (): boolean => true, 
          data: (): any => foundDoc.data() 
        } : { 
          exists: (): boolean => false,
          data: (): any => undefined
        };
      });
      
      if (docSnap.exists()) {
        const data = docSnap.data();
        if (!data) {
          setSubmitStatus('No data found - initialized empty form');
          setTimeout(() => setSubmitStatus(''), 2000);
          return;
        }
        
        // Production에서는 로그 출력 안 함
        if (process.env.NODE_ENV !== 'production') {
          console.log("Data loaded from Firebase:", data);
        }
        
        // Firebase에서 업데이트되는 동안 자동 저장 방지
        setIsUpdatingFromFirebase(true);
        
        setPatientData(prevData => ({
          ...prevData,
          ...data
        }));
        
        // rowCount 복원
        if (data.rowCount) {
          setRowCount(data.rowCount);
        }
        
        // 로드된 데이터를 마지막 저장된 데이터로 설정
        setLastSavedData({ ...data, rowCount: data.rowCount || 20 });
        
        // 짧은 지연 후 Firebase 업데이트 플래그 해제
        setTimeout(() => {
          setIsUpdatingFromFirebase(false);
        }, 100);
        
        setSubmitStatus('Data loaded successfully');
        setTimeout(() => setSubmitStatus(''), 2000);
      } else {
        console.log("No data found for date:", dutyDate);
        // 데이터가 없으면 초기화
        setPatientData({});
        setRowCount(20);
        setLastSavedData({ rowCount: 20 });
        setSubmitStatus('No data found - initialized empty form');
        setTimeout(() => setSubmitStatus(''), 2000);
      }
      
    } catch (error) {
      // Production에서는 에러 로깅 비활성화
      if (process.env.NODE_ENV !== 'production') {
        console.error("Error loading data:", error);
      }
      setSubmitStatus('Error loading data: ' + ((error as any).message || 'Unknown error'));
      setTimeout(() => setSubmitStatus(''), 3000);
    }
  };

  // 날짜 또는 오피스 변경 시 데이터 로드
  useEffect(() => {
    loadData();
  }, [dutyDate, selectedOffice]);

  // 실시간 데이터 동기화
  useEffect(() => {
    if (!dutyDate || !selectedOffice) return;

    // Production에서는 로그 출력 안 함
    if (process.env.NODE_ENV !== 'production') {
      console.log("Setting up real-time listener for date:", dutyDate, "office:", selectedOffice);
    }
    const docId = sanitizeFirebaseDocIdClient(`${dutyDate}_${selectedOffice}_addon_treatment`);
    const docRef = doc(db, "addon-treatment", docId);
    
    const unsubscribe = onSnapshot(docRef, (docSnap) => {
      if (docSnap.exists()) {
        const data = docSnap.data();
        // Production에서는 로그 출력 안 함
        if (process.env.NODE_ENV !== 'production') {
          console.log("Real-time data received:", data);
        }
        
        // Firebase에서 업데이트되는 동안 자동 저장 방지
        setIsUpdatingFromFirebase(true);
        
        setPatientData(prevData => {
          // Production에서는 로그 출력 안 함
          if (process.env.NODE_ENV !== 'production') {
            console.log("Updating patientData from:", prevData, "to:", { ...prevData, ...data });
          }
          return {
            ...prevData,
            ...data
          };
        });
        
        // rowCount 복원
        if (data.rowCount) {
          setRowCount(data.rowCount);
        }
        
        // 실시간 업데이트된 데이터를 마지막 저장된 데이터로 설정
        setLastSavedData({ ...data, rowCount: data.rowCount || 20 });
        
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
        // Production에서는 로그 출력 안 함
        if (process.env.NODE_ENV !== 'production') {
          console.log("Real-time listener: No document exists for date:", dutyDate);
        }
      }
    }, (error) => {
      // Production에서는 에러 로깅 비활성화
      if (process.env.NODE_ENV !== 'production') {
        console.error("Real-time listener error:", error);
      }
      setAutoSaveStatus('❌ Connection error');
      setTimeout(() => setAutoSaveStatus(''), 3000);
    });

    return () => {
      // Production에서는 로그 출력 안 함
      if (process.env.NODE_ENV !== 'production') {
        console.log("Cleaning up real-time listener for date:", dutyDate, "office:", selectedOffice);
      }
      unsubscribe();
    };
  }, [dutyDate, selectedOffice]);

  // 데이터 업데이트 함수
  const updatePatientData = (field: string, value: string) => {
    setPatientData(prev => {
      const newData = { ...prev, [field]: value };
      
      // Patient Name 필드가 입력되면 시간 자동 기록
      if (field.endsWith('_PatientName') && value.trim() !== '' && (prev[field] || '').trim() === '') {
        const rowNumber = field.match(/Row(\d+)_PatientName/)?.[1];
        if (rowNumber) {
          const timeField = `Row${rowNumber}_Time`;
          if (!newData[timeField]) {
            newData[timeField] = getCurrentTime12Hour();
          }
        }
      }
      
      // Patient Name 필드가 비워지면 시간도 비움
      if (field.endsWith('_PatientName') && value.trim() === '') {
        const rowNumber = field.match(/Row(\d+)_PatientName/)?.[1];
        if (rowNumber) {
          const timeField = `Row${rowNumber}_Time`;
          newData[timeField] = '';
        }
      }
      
      return newData;
    });
  };

  // 제출 처리
  const handleSubmit = async () => {
    const password = prompt(`Are you sure you want to submit? Submitting will reset today's data. Enter password to proceed (Password: ${selectedOffice}):`);
    if (password === null) return;
    if (password !== selectedOffice) {
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
      
      // 환자 데이터를 배열 형태로 변환
      const patientRows: Array<{
        rowNumber: number;
        patientName: string;
        dob: string;
        time: string;
      }> = [];
      for (let i = 1; i <= rowCount; i++) {
        const patientName = patientData[`Row${i}_PatientName`];
        const dob = patientData[`Row${i}_DOB`];
        const time = patientData[`Row${i}_Time`];
        
        if (patientName || dob || time) {
          patientRows.push({
            rowNumber: i,
            patientName: patientName || '',
            dob: dob || '',
            time: time || ''
          });
        }
      }
      
      const response = await fetch('/api/generate-addon-treatment-pdf', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          dutyDate,
          patientRows
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
        const docId = sanitizeFirebaseDocIdClient(`${dutyDate}_${selectedOffice}_addon_treatment`);
        await deleteDoc(doc(db, "addon-treatment", docId));
        
        // 3. 폼 초기화
        setPatientData({});
        setRowCount(20);

        setSubmitStatus('Complete!');
        setProgress(100);
        
        // PDF 다운로드
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `4) ${dutyDate}_${selectedOffice}_Add On Treatment.pdf`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        
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
      // Production에서는 에러 로깅 비활성화
      if (process.env.NODE_ENV !== 'production') {
        console.error('Submit error:', error);
      }
      setSubmitStatus('❌ Submission failed: ' + ((error as any).message || 'Unknown error'));
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
      fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
      padding: '20px',
      background: `
        radial-gradient(circle at 10% 20%, rgba(120, 200, 255, 0.1) 0%, transparent 50%),
        radial-gradient(circle at 90% 80%, rgba(255, 182, 193, 0.1) 0%, transparent 50%),
        radial-gradient(circle at 50% 50%, rgba(144, 238, 144, 0.05) 0%, transparent 50%),
        linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%)
      `,
      color: '#2c3e50',
      lineHeight: '1.6',
      minHeight: '100vh',
      position: 'relative' as const
    },
    container: {
      maxWidth: '1200px',
      margin: '40px auto',
      padding: '30px',
      backgroundColor: 'white',
      borderRadius: '12px',
      boxShadow: '0 4px 20px rgba(0, 0, 0, 0.1)',
      border: '1px solid #e9ecef',
      position: 'relative' as const,
      overflow: 'hidden'
    },
    header: {
      color: '#2c3e50',
      textAlign: 'center' as const,
      marginBottom: '30px',
      paddingBottom: '20px',
      borderBottom: '3px solid transparent',
      borderImage: 'linear-gradient(90deg, #4CAF50, #2196F3, #FF9800) 1',
      fontSize: '2.2em',
      fontWeight: '600',
      letterSpacing: '-0.5px',
      position: 'relative' as const
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
      fontSize: '16px',
      border: '1px solid #e9ecef',
      borderRadius: '4px',
      backgroundColor: 'white',
      boxSizing: 'border-box' as const,
      transition: 'border-color 0.2s'
    },
    table: {
      borderCollapse: 'collapse' as const,
      width: '100%',
      marginTop: '20px',
      backgroundColor: 'white',
      borderRadius: '6px',
      boxShadow: '0 1px 3px rgba(0, 0, 0, 0.1)',
      border: '1px solid #dee2e6',
      overflow: 'hidden'
    },
    th: {
      border: '1px solid #dee2e6',
      padding: '12px 10px',
      textAlign: 'center' as const,
      verticalAlign: 'middle' as const,
      background: 'linear-gradient(135deg, #4CAF50 0%, #45a049 100%)',
      color: 'white',
      fontWeight: '600',
      fontSize: '14px',
      textShadow: '0 1px 2px rgba(0, 0, 0, 0.1)'
    },
    td: {
      border: '1px solid #dee2e6',
      padding: '10px',
      textAlign: 'center' as const,
      verticalAlign: 'middle' as const,
      backgroundColor: 'white',
      transition: 'background-color 0.2s ease'
    },
    submitButton: {
      display: 'block',
      width: '150px',
      margin: '30px auto 0 auto',
      padding: '12px 20px',
      background: 'linear-gradient(135deg, #2196F3 0%, #1976D2 100%)',
      color: 'white',
      border: 'none',
      borderRadius: '8px',
      cursor: 'pointer',
      fontSize: '16px',
      fontWeight: '600',
      transition: 'all 0.3s ease',
      boxShadow: '0 2px 8px rgba(33, 150, 243, 0.3)',
      textShadow: '0 1px 2px rgba(0, 0, 0, 0.1)'
    },
    autoSaveStatus: {
      position: 'absolute' as const,
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
    footer: {
      marginTop: '40px',
      paddingTop: '20px',
      borderTop: '1px solid #e9ecef',
      textAlign: 'center' as const,
      fontSize: '0.9em',
      color: '#6c757d'
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
            position: "fixed" as const,
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
              textAlign: "center",
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

          <h2 style={styles.header}>
            <span style={{ marginRight: '10px' }}>🏥</span>
            Add-On Treatment
            <span style={{ marginLeft: '10px' }}>💊</span>
          </h2>

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
                onChange={(e) => setSelectedOffice(e.target.value)}
                style={styles.input}
                required
              >
                {officeOptions.map(office => (
                  <option key={office} value={office}>{office}</option>
                ))}
              </select>
            </div>
          </div>

          {/* 환자 로그 테이블 */}
          <table style={styles.table}>
            <thead>
              <tr>
                <th style={{ ...styles.th, width: '60px' }}>No.</th>
                <th style={styles.th}>Name of Patient</th>
                <th style={styles.th}>DOB</th>
                <th style={styles.th}>Time</th>
                <th style={{ ...styles.th, width: '80px' }}>Action</th>
              </tr>
            </thead>
            <tbody>
              {Array.from({ length: rowCount }, (_, index) => {
                const rowNumber = index + 1;
                return (
                  <tr 
                    key={rowNumber}
                    style={{
                      transition: 'background-color 0.2s ease'
                    }}
                    onMouseEnter={(e) => {
                      e.currentTarget.style.backgroundColor = '#f8f9fa';
                    }}
                    onMouseLeave={(e) => {
                      e.currentTarget.style.backgroundColor = 'white';
                    }}
                  >
                    <td style={styles.td}>{rowNumber}</td>
                    <td style={styles.td}>
                      <input
                        type="text"
                        value={patientData[`Row${rowNumber}_PatientName`] || ''}
                        onChange={(e) => updatePatientData(`Row${rowNumber}_PatientName`, e.target.value)}
                        style={{ ...styles.input, margin: 0, fontSize: '14px', padding: '8px' }}
                      />
                    </td>
                    <td style={styles.td}>
                      <input
                        type="text"
                        value={patientData[`Row${rowNumber}_DOB`] || ''}
                        onChange={(e) => updatePatientData(`Row${rowNumber}_DOB`, e.target.value)}
                        placeholder="mm/dd/yyyy"
                        style={{ ...styles.input, margin: 0, fontSize: '14px', padding: '8px' }}
                      />
                    </td>
                    <td style={styles.td}>
                      <input
                        type="text"
                        value={convertTo12Hour(patientData[`Row${rowNumber}_Time`] || '')}
                        onChange={(e) => updatePatientData(`Row${rowNumber}_Time`, e.target.value)}
                        style={{ ...styles.input, margin: 0, fontSize: '14px', padding: '8px' }}
                      />
                    </td>
                    <td style={styles.td}>
                      <button
                        type="button"
                        onClick={() => clearRow(rowNumber)}
                        style={{
                          background: 'linear-gradient(135deg, #FF9800 0%, #F57C00 100%)',
                          color: 'white',
                          border: 'none',
                          padding: '8px 16px',
                          borderRadius: '6px',
                          cursor: 'pointer',
                          fontSize: '12px',
                          fontWeight: '600',
                          transition: 'all 0.3s ease',
                          boxShadow: '0 2px 6px rgba(255, 152, 0, 0.3)',
                          textShadow: '0 1px 2px rgba(0, 0, 0, 0.1)'
                        }}
                      >
                        Clear
                      </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>

          {/* 행 추가 버튼 */}
          <div style={{ textAlign: 'center', marginTop: '20px', marginBottom: '10px' }}>
            <button
              type="button"
              onClick={addRow}
              style={{
                background: 'linear-gradient(135deg, #4CAF50 0%, #45a049 100%)',
                color: 'white',
                border: 'none',
                padding: '12px 24px',
                borderRadius: '8px',
                cursor: 'pointer',
                fontSize: '14px',
                fontWeight: '600',
                transition: 'all 0.3s ease',
                boxShadow: '0 2px 8px rgba(76, 175, 80, 0.3)',
                textShadow: '0 1px 2px rgba(0, 0, 0, 0.1)'
              }}
            >
              + Add Row
            </button>
          </div>

          {/* 제출 버튼 */}
          <div style={{ textAlign: 'center', marginTop: '10px' }}>
            <button
              type="button"
              onClick={handleSubmit}
              disabled={loading}
              style={{
                ...styles.submitButton,
                background: loading ? 'linear-gradient(135deg, #6c757d 0%, #5a6268 100%)' : 'linear-gradient(135deg, #2196F3 0%, #1976D2 100%)',
                cursor: loading ? 'not-allowed' : 'pointer'
              }}
              onMouseEnter={(e) => {
                if (!loading) {
                  e.currentTarget.style.transform = 'translateY(-1px)';
                  e.currentTarget.style.boxShadow = '0 4px 12px rgba(33, 150, 243, 0.4)';
                }
              }}
              onMouseLeave={(e) => {
                if (!loading) {
                  e.currentTarget.style.transform = 'translateY(0)';
                  e.currentTarget.style.boxShadow = '0 2px 8px rgba(33, 150, 243, 0.3)';
                }
              }}
            >
              {loading ? 'Submitting...' : 'Submit'}
            </button>
          </div>

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
            }}>
              {submitStatus}
            </div>
          )}

          {/* Footer */}
          <div style={styles.footer}>
            <p style={{ marginBottom: '0' }}>Smileland Dental</p>
          </div>
        </div>
      </div>
    </>
  );
}