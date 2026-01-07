'use client'

import React, { useState, useEffect, useCallback } from "react";
import { doc, setDoc, collection, getDocs, deleteDoc, getDoc } from "firebase/firestore";
import { db } from "@/lib/firebase.config";

// 개별 환자 행 컴포넌트 (메모이제이션 최적화)
const PatientRow = React.memo(({ 
  row, 
  updatePatientRow, 
  removePatientRow, 
  patientOfficeOptions, 
  getVisitTypeOptions, 
  remarkOptions, 
  otherDutyOptions,
  inputStyle,
  buttonStyle
}) => {
  const visitTypeOptions = getVisitTypeOptions(row.office);
  
  return (
    <tr style={{ backgroundColor: row.id % 2 === 0 ? '#f9f9f9' : 'white' }}>
      <td style={{ padding: '8px', textAlign: 'center', fontWeight: 'bold' }}>
        {row.id}
      </td>
      <td style={{ padding: '8px' }}>
        <input
          type="text"
          value={row.name}
          onChange={(e) => updatePatientRow(row.id, 'name', e.target.value)}
          style={{ ...inputStyle, margin: 0, fontSize: '14px' }}
        />
      </td>
      <td style={{ padding: '8px' }}>
        <select
          value={row.office}
          onChange={(e) => updatePatientRow(row.id, 'office', e.target.value)}
          style={{ ...inputStyle, margin: 0, fontSize: '14px' }}
        >
          <option value=""></option>
          {patientOfficeOptions.map(office => (
            <option key={office} value={office}>{office}</option>
          ))}
        </select>
      </td>
      <td style={{ padding: '8px' }}>
        <input
          type="date"
          value={row.appt_date}
          onChange={(e) => updatePatientRow(row.id, 'appt_date', e.target.value)}
          style={{ ...inputStyle, margin: 0, fontSize: '14px' }}
        />
      </td>
      <td style={{ padding: '8px' }}>
        <select
          value={row.visit_type}
          onChange={(e) => updatePatientRow(row.id, 'visit_type', e.target.value)}
          style={{ ...inputStyle, margin: 0, fontSize: '14px' }}
        >
          <option value=""></option>
          {visitTypeOptions.map(type => (
            <option key={type} value={type}>{type}</option>
          ))}
        </select>
      </td>
      <td style={{ padding: '8px', textAlign: 'center' }}>
        <input
          type="checkbox"
          checked={row.call_in}
          onChange={(e) => updatePatientRow(row.id, 'call_in', e.target.checked)}
          style={{ transform: 'scale(1.2)' }}
        />
      </td>
      <td style={{ padding: '8px', textAlign: 'center' }}>
        <input
          type="checkbox"
          checked={row.call_out}
          onChange={(e) => updatePatientRow(row.id, 'call_out', e.target.checked)}
          style={{ transform: 'scale(1.2)' }}
        />
      </td>
      <td style={{ padding: '8px' }}>
        <input
          type="time"
          value={row.time}
          onChange={(e) => updatePatientRow(row.id, 'time', e.target.value)}
          disabled={row.call_in || row.call_out}
          style={{ 
            ...inputStyle, 
            margin: 0, 
            fontSize: '14px',
            backgroundColor: (row.call_in || row.call_out) ? '#f0f0f0' : 'white',
            cursor: (row.call_in || row.call_out) ? 'not-allowed' : 'text'
          }}
        />
      </td>
      <td style={{ padding: '8px' }}>
        <select
          value={row.remark}
          onChange={(e) => updatePatientRow(row.id, 'remark', e.target.value)}
          style={{ ...inputStyle, margin: 0, fontSize: '14px' }}
        >
          <option value=""></option>
          {remarkOptions.map(remark => (
            <option key={remark} value={remark}>{remark}</option>
          ))}
        </select>
      </td>
      <td style={{ padding: '8px' }}>
        <select
          value={row.other_duty}
          onChange={(e) => updatePatientRow(row.id, 'other_duty', e.target.value)}
          style={{ ...inputStyle, margin: 0, fontSize: '14px' }}
        >
          <option value=""></option>
          {otherDutyOptions.map(duty => (
            <option key={duty} value={duty}>{duty}</option>
          ))}
        </select>
      </td>
      <td style={{ padding: '8px', textAlign: 'center' }}>
        <button
          onClick={() => removePatientRow(row.id)}
          style={{
            ...buttonStyle,
            backgroundColor: '#dc3545',
            padding: '6px 12px',
            fontSize: '12px',
            margin: 0
          }}
        >
          Delete
        </button>
      </td>
    </tr>
  );
}, (prevProps, nextProps) => {
  // 커스텀 비교 함수로 불필요한 리렌더링 방지
  return (
    prevProps.row.id === nextProps.row.id &&
    prevProps.row.name === nextProps.row.name &&
    prevProps.row.office === nextProps.row.office &&
    prevProps.row.appt_date === nextProps.row.appt_date &&
    prevProps.row.visit_type === nextProps.row.visit_type &&
    prevProps.row.call_in === nextProps.row.call_in &&
    prevProps.row.call_out === nextProps.row.call_out &&
    prevProps.row.time === nextProps.row.time &&
    prevProps.row.remark === nextProps.row.remark &&
    prevProps.row.other_duty === nextProps.row.other_duty
  );
});

function PatientLogSystem() {
  
  // 기본 상태
  const [loading, setLoading] = useState(false);
  const [savedLogs, setSavedLogs] = useState([]);
  const [autoSaveStatus, setAutoSaveStatus] = useState(''); // 자동 저장 상태 표시
  
  // 마지막 저장된 데이터 추적 (dailyofficeduty 방식)
  const [lastSavedData, setLastSavedData] = useState({});
  const [submitStatus, setSubmitStatus] = useState('');
  const [progress, setProgress] = useState(0);

  // 폼 데이터 상태 (원본과 동일한 구조)
  const [formData, setFormData] = useState({
    dutyDate: new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' }),
    userName: '',
    workOffice: '',
    workHoursFrom: '',
    workHoursTo: '',
    dailyWorkReport: ''
  });

  // 환자 로그 상태 (원본과 동일한 필드명) - 기본 30행
  const [patientRows, setPatientRows] = useState(() => {
    return Array.from({ length: 30 }, (_, index) => ({
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
    }));
  });

  // patientRows를 useMemo로 최적화

  // 실시간 카운트 계산 (단순화)
  const appointments = patientRows.filter(row => row.appt_date && row.name).length;
  const incomingCalls = patientRows.filter(row => row.call_in).length;
  const outgoingCalls = patientRows.filter(row => row.call_out).length;

  // Office 옵션 (단순화)
  const workOfficeOptions = ['Bernard', 'Call Center', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];
  const patientOfficeOptions = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];
  
  // Visit Type 옵션을 Office에 따라 동적으로 생성 (useCallback 최적화)
  const getVisitTypeOptions = useCallback((office) => {
    if (office === 'Ortho') {
      return ['Adjustment', 'Bonding', 'Consult', 'Full Deband', 'Partial Deband', 'Records', 'Retainer Check', 'RPE Check'];
    } else {
      return ['Emergency', 'New Patient', 'RCRA', 'Recall', 'Tx'];
    }
  }, []);

  // Remark 옵션 (단순화)
  const remarkOptions = ['Disc', 'Elsewhere', 'LMA', 'LMW', 'NA', 'Not Interested', 'Wrong'];

  // Other Duty 옵션 (단순화)
  const otherDutyOptions = [
    'Accounts with Balances', 'Booking ASL Interpreters', 'Break', 'Confirming', 
    'Incoming Call Report', 'Insurance Verifications', 'Lunch', 'Marketing Data', 
    'Medi-cal Eligibility', 'Monthly Report', 'Nintendo Switch Raffle', 'One Week\'s Reconfirming', 
    'Other', 'Postcards', 'Refer a friend', 'Reviews', 'Routing Slips', 
    'Sending Replacement Staff', 'Training'
  ];

  // Document ID 생성 함수 (모든 Basic Information 포함)
  const generateDocId = (dutyDate, userName, workOffice, workHoursFrom, workHoursTo) => {
    return `${dutyDate}_${userName}_${workOffice}_${workHoursFrom}_${workHoursTo}`.replace(/[^a-zA-Z0-9_-]/g, '_');
  };

  // Basic Information 완료 체크 함수
  const isBasicInfoComplete = () => {
    return formData.dutyDate && 
           formData.userName && 
           formData.workOffice && 
           formData.workHoursFrom && 
           formData.workHoursTo;
  };

  // 자동 저장 함수 (매우 빠른 저장)
  const autoSave = useCallback(async () => {
    // Basic Information이 완료되지 않으면 저장하지 않음
    if (!isBasicInfoComplete()) return;

    // 이름이 공백으로 끝나면 아직 입력 중이므로 저장하지 않음
    if (formData.userName.trim().endsWith(' ')) {
      return;
    }

    try {
      const dataToSave = {
        ...formData,
        patientRows: patientRows.filter(row => 
          row.name || row.office || row.appt_date || row.visit_type || 
          row.call_in || row.call_out || row.time || row.remark || row.other_duty
        ),
        timestamp: new Date().toISOString(),
        autoSaved: true
      };

      const docId = generateDocId(formData.dutyDate, formData.userName, formData.workOffice, formData.workHoursFrom, formData.workHoursTo);
      
      // 자동 저장 시작 표시
      setAutoSaveStatus('💾 Saving...');
      
      // Firebase 저장 (비동기 처리로 UI 블로킹 방지)
      setDoc(doc(db, "patient-logs", docId), dataToSave)
        .then(() => {
          // 저장 성공 표시
          setAutoSaveStatus('✅ Auto-saved');
          
          // 1초 후 상태 메시지 제거
          setTimeout(() => {
            setAutoSaveStatus('');
          }, 1000);
        })
        .catch((error) => {
          console.error("Auto-save error:", error);
          setAutoSaveStatus('❌ Save failed');
          
          // 2초 후 상태 메시지 제거
          setTimeout(() => {
            setAutoSaveStatus('');
          }, 2000);
        });
      
    } catch (error) {
      console.error("Auto-save error:", error);
    }
  }, [formData, patientRows]);

  // 데이터 변경 시 자동 저장 (매우 빠른 debounce)
  useEffect(() => {
    // 초기 로드 시에는 저장하지 않음
    if (Object.values(formData).some(value => value !== '') || patientRows.length > 0) {
      const timeoutId = setTimeout(() => {
        autoSave();
      }, 300); // 0.3초 debounce로 매우 빠르게

      return () => clearTimeout(timeoutId);
    }
  }, [formData, patientRows, autoSave]);

  // 저장된 데이터를 현재 환자 로그에 로드하는 함수 (최적화)
  const loadExistingData = async () => {
    if (!isBasicInfoComplete()) {
      return;
    }

    try {
      const docId = generateDocId(formData.dutyDate, formData.userName, formData.workOffice, formData.workHoursFrom, formData.workHoursTo);
      const docRef = doc(db, "patient-logs", docId);
      
      // 직접 document 참조로 조회 (전체 collection 스캔 방지)
      const docSnap = await getDoc(docRef);

      if (docSnap.exists()) {
        const matchingLog = docSnap.data();
        
        if (matchingLog && matchingLog.patientRows) {
          // 기존 저장된 환자 로그를 현재 patientRows에 로드
          const loadedRows = matchingLog.patientRows.map((row, index) => ({
            ...row,
            id: index + 1
          }));
          
          // 30개 행을 유지하되, 저장된 데이터로 채움
          const newRows = Array.from({ length: 30 }, (_, index) => {
            if (index < loadedRows.length) {
              return loadedRows[index];
            }
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
            };
          });
          
          setPatientRows(newRows);
          
          // Daily Work Report도 로드
          if (matchingLog.dailyWorkReport) {
            setFormData(prev => ({
              ...prev,
              dailyWorkReport: matchingLog.dailyWorkReport
            }));
          }
        }
      }
    } catch (error) {
      console.error("Error loading existing data:", error);
    }
  };

  // 기본 정보가 입력되면 기존 데이터 로드 (매우 빠른 로드)
  useEffect(() => {
    const timeoutId = setTimeout(() => {
      loadExistingData();
    }, 50); // 0.05초 debounce로 매우 빠르게

    return () => clearTimeout(timeoutId);
  }, [formData.dutyDate, formData.userName, formData.workOffice, formData.workHoursFrom, formData.workHoursTo]);


  // 컴포넌트 마운트 시 저장된 로그 불러오기
  useEffect(() => {
    loadSavedLogs();
  }, []);

  // 저장된 로그 불러오기
  const loadSavedLogs = async () => {
    try {
      const querySnapshot = await getDocs(collection(db, "patient-logs"));
      const logs: any[] = [];
      querySnapshot.forEach((doc) => {
        logs.push({ id: doc.id, ...doc.data() });
      });
      setSavedLogs(logs.sort((a: any, b: any) => new Date(b.dutyDate).getTime() - new Date(a.dutyDate).getTime()));
    } catch (error) {
      console.error("Error loading logs:", error);
    }
  };

  // 폼 데이터 업데이트 (최고 성능 업데이트)
  const updateFormData = useCallback((field, value) => {
    setFormData(prev => {
      // 값이 같으면 업데이트하지 않음
      if (prev[field] === value) return prev;
      return { ...prev, [field]: value };
    });
  }, []);


  // 환자 행 추가 (useCallback 최적화)
  const addPatientRow = useCallback(() => {
    setPatientRows(prevRows => {
      const newId = Math.max(...prevRows.map(row => row.id)) + 1;
      return [...prevRows, {
        id: newId,
        name: '',
        office: '',
        appt_date: '',
        visit_type: '',
        call_in: false,
        call_out: false,
        time: '',
        remark: '',
        other_duty: ''
      }];
    });
  }, []);

  // 환자 행 삭제 (useCallback 최적화)
  const removePatientRow = useCallback((id) => {
    setPatientRows(prevRows => {
      if (prevRows.length > 1) {
        return prevRows.filter(row => row.id !== id);
      }
      return prevRows;
    });
  }, []);

  // 환자 행 업데이트 (최고 성능 업데이트)
  const updatePatientRow = useCallback((id, field, value) => {
    setPatientRows(prevRows => {
      const rowIndex = prevRows.findIndex(row => row.id === id);
      
      if (rowIndex === -1) {
        return prevRows;
      }
      
      const row = prevRows[rowIndex];
      
      // 값이 같으면 업데이트하지 않음
      if (row[field] === value) {
        return prevRows;
      }
      
      const updatedRow = { ...row, [field]: value };
      
      // Office가 변경되면 visit_type을 초기화
      if (field === 'office' && row.office !== value) {
        updatedRow.visit_type = '';
      }
      // Call In 또는 Call Out이 체크되면 현재 시간을 Time에 자동 입력
      if ((field === 'call_in' || field === 'call_out') && value === true) {
        const now = new Date();
        const timeString = now.toTimeString().slice(0, 5);
        updatedRow.time = timeString;
      }
      // Call In과 Call Out이 모두 체크 해제되면 Time을 비움
      if ((field === 'call_in' && value === false && !row.call_out) || 
          (field === 'call_out' && value === false && !row.call_in)) {
        updatedRow.time = '';
      }
      
      const newRows = [...prevRows];
      newRows[rowIndex] = updatedRow;
      return newRows;
    });
  }, []);

  // 폼 초기화 함수
  const resetForm = () => {
    setFormData({
      dutyDate: new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' }),
      userName: '',
      workOffice: '',
      workHoursFrom: '',
      workHoursTo: '',
      dailyWorkReport: ''
    });
    setPatientRows([{
      id: 1,
      name: '',
      office: '',
      appt_date: '',
      visit_type: '',
      call_in: false,
      call_out: false,
      time: '',
      remark: '',
      other_duty: ''
    }]);
    setLastSavedData({});
  };

  // PDF 생성 및 제출
  const handleSubmit = async () => {
    if (!isBasicInfoComplete()) {
      alert('⚠️ Please fill in all Basic Information fields.');
      return;
    }

    try {
      setLoading(true);
      setSubmitStatus('Saving...');
      setProgress(10);

      // 1. Firebase에 데이터 저장
      const dataToSave = {
        ...formData,
        patientRows: patientRows.filter(row => 
          row.name || row.office || row.appt_date || row.visit_type || 
          row.call_in || row.call_out || row.time || row.remark || row.other_duty
        ),
        timestamp: new Date().toISOString()
      };

      const docId = generateDocId(formData.dutyDate, formData.userName, formData.workOffice, formData.workHoursFrom, formData.workHoursTo);
      await setDoc(doc(db, "patient-logs", docId), dataToSave);

      // 2. API로 PDF 생성
      setSubmitStatus('Generating PDF...');
      setProgress(30);
      
      const response = await fetch('/api/generate-pdf', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          patientData: {
            ...formData,
            patientLogs: patientRows.filter(row => 
              row.name || row.office || row.appt_date || row.visit_type || 
              row.call_in || row.call_out || row.time || row.remark || row.other_duty
            )
          }
        }),
      });

      if (response.ok) {
        // PDF blob 받기
        setSubmitStatus('Processing PDF...');
        setProgress(60);
        const blob = await response.blob();
        
        // 파일명 생성
        const date = formData.dutyDate || new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' });
        const name = formData.userName || 'Unknown';
        const office = formData.workOffice || 'Unknown';
        const filename = `${date}_${name}_${office}_Patient_Log.pdf`;
        
        setSubmitStatus('✅ Submitted Successfully!');
        setProgress(100);
        
        // PDF 다운로드
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        
        // 폼 초기화 (제출 완료 느낌)
        resetForm();
        
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
      console.error('Submit 오류:', error);
      setSubmitStatus('❌ Submission failed: ' + error.message);
      setProgress(0);
      setTimeout(() => {
        setLoading(false);
        setSubmitStatus('');
      }, 3000);
    }
  };

  // 제출 중 브라우저 네비게이션 방지
  useEffect(() => {
    const handleBeforeUnload = (e) => {
      if (loading) {
        e.preventDefault();
        e.returnValue = '';
      }
    };

    const handlePopState = (e) => {
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

  // 원본 HTML 스타일을 React 스타일로 변환
  const containerStyle = {
    maxWidth: '1500px',
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
    backgroundColor: '#0077B6',
    color: 'white',
    border: 'none',
    padding: '12px 24px',
    borderRadius: '6px',
    fontSize: '1em',
    fontWeight: 'bold',
    cursor: 'pointer',
    margin: '5px',
    transition: 'all 0.3s ease'
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
                  ⚠️ Please do not close
                </p>
              </div>
            </div>
          </div>
        )}

        <div style={containerStyle}>
        {/* 헤더 */}
        <div style={{ position: 'relative' }}>
        <h1 style={headerStyle}>🌴 Patient Log</h1>
          {autoSaveStatus && (
            <div style={{
              position: 'absolute',
              top: '10px',
              right: '10px',
              padding: '8px 16px',
              backgroundColor: autoSaveStatus.includes('실패') ? '#ff6b6b' : '#51cf66',
              color: 'white',
              borderRadius: '20px',
              fontSize: '14px',
              fontWeight: 'bold',
              boxShadow: '0 2px 8px rgba(0,0,0,0.2)',
              zIndex: 1000
            }}>
              {autoSaveStatus}
            </div>
          )}
        </div>

        {/* 기본 정보 섹션 */}
        <div style={sectionStyle}>
          <h2 style={{ color: '#0077B6', marginBottom: '15px' }}>📋 Basic Information</h2>
          
          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap', marginBottom: '20px' }}>
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Duty Date:
              </label>
              <input
                type="date"
                value={formData.dutyDate}
                onChange={(e) => updateFormData('dutyDate', e.target.value)}
                style={inputStyle}
                required
              />
            </div>
            
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Name:
              </label>
              <input
                type="text"
                value={formData.userName}
                onChange={(e) => updateFormData('userName', e.target.value)}
                placeholder="Enter your full name"
                style={inputStyle}
                required
              />
            </div>

            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Work Office:
              </label>
              <select
                value={formData.workOffice}
                onChange={(e) => updateFormData('workOffice', e.target.value)}
                style={inputStyle}
                required
              >
                <option value="">Select Office</option>
                {workOfficeOptions.map(office => (
                  <option key={office} value={office}>{office}</option>
                ))}
              </select>
            </div>
          </div>

          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap' }}>
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Work Hours From:
              </label>
              <input
                type="time"
                value={formData.workHoursFrom}
                onChange={(e) => updateFormData('workHoursFrom', e.target.value)}
                style={inputStyle}
                required
              />
            </div>
            
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Work Hours To:
              </label>
              <input
                type="time"
                value={formData.workHoursTo}
                onChange={(e) => updateFormData('workHoursTo', e.target.value)}
                style={inputStyle}
                required
              />
            </div>
          </div>
        </div>

        {/* 환자 로그 테이블 */}
        <div style={sectionStyle}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '15px' }}>
            <h2 style={{ color: '#0077B6', margin: 0 }}>👥 Patient Log</h2>
              {!isBasicInfoComplete() && (
                <span style={{ 
                  fontSize: '12px', 
                  color: '#dc3545', 
                  fontWeight: 'bold',
                  padding: '4px 8px',
                  backgroundColor: '#f8d7da',
                  borderRadius: '12px',
                  border: '1px solid #dc3545'
                }}>
                  ⚠️ Complete Basic Information First
                </span>
              )}
              {isBasicInfoComplete() && (
                <span style={{ 
                  fontSize: '12px', 
                  color: '#28a745', 
                  fontWeight: 'bold',
                  padding: '4px 8px',
                  backgroundColor: '#e8f5e8',
                  borderRadius: '12px',
                  border: '1px solid #28a745'
                }}>
                  ✅ Data Loaded
                </span>
              )}
              {autoSaveStatus && (
                <span style={{ 
                  fontSize: '12px', 
                  color: autoSaveStatus.includes('❌') ? '#dc3545' : autoSaveStatus.includes('💾') ? '#007bff' : '#28a745',
                  fontWeight: 'bold',
                  padding: '4px 8px',
                  backgroundColor: autoSaveStatus.includes('❌') ? '#f8d7da' : autoSaveStatus.includes('💾') ? '#d1ecf1' : '#e8f5e8',
                  borderRadius: '12px',
                  border: `1px solid ${autoSaveStatus.includes('❌') ? '#dc3545' : autoSaveStatus.includes('💾') ? '#007bff' : '#28a745'}`,
                  marginLeft: '10px'
                }}>
                  {autoSaveStatus}
                </span>
              )}
              {formData.userName && formData.userName.trim().endsWith(' ') && (
                <span style={{ 
                  fontSize: '12px', 
                  color: '#ffc107',
                  fontWeight: 'bold',
                  padding: '4px 8px',
                  backgroundColor: '#fff3cd',
                  borderRadius: '12px',
                  border: '1px solid #ffc107',
                  marginLeft: '10px'
                }}>
                  ⏳ Typing...
                </span>
              )}
            </div>
          </div>

          {/* 실시간 카운트 표시 */}
          <div style={{ 
            display: 'flex', 
            justifyContent: 'center',
            gap: '30px', 
            marginBottom: '20px',
            padding: '15px',
            backgroundColor: '#e3f2fd',
            borderRadius: '8px',
            border: '1px solid #bbdefb'
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span style={{ fontWeight: 'bold', color: '#1565c0' }}>📅 Appointments:</span>
              <span style={{ 
                fontSize: '18px', 
                fontWeight: 'bold', 
                color: '#0077B6',
                minWidth: '20px'
              }}>
                {appointments}
              </span>
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span style={{ fontWeight: 'bold', color: '#1565c0' }}>📞 Incoming Calls:</span>
              <span style={{ 
                fontSize: '18px', 
                fontWeight: 'bold', 
                color: '#28a745',
                minWidth: '20px'
              }}>
                {incomingCalls}
              </span>
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span style={{ fontWeight: 'bold', color: '#1565c0' }}>📱 Outgoing Calls:</span>
              <span style={{ 
                fontSize: '18px', 
                fontWeight: 'bold', 
                color: '#ff6b35',
                minWidth: '20px'
              }}>
                {outgoingCalls}
              </span>
            </div>
          </div>

          {/* Basic Information 완료 체크 후 테이블 표시 */}
          {!isBasicInfoComplete() ? (
            <div style={{
              padding: '40px',
              textAlign: 'center',
              backgroundColor: '#f8f9fa',
              border: '2px dashed #dee2e6',
              borderRadius: '8px',
              margin: '20px 0'
            }}>
              <div style={{ fontSize: '48px', marginBottom: '20px' }}>🔒</div>
              <h3 style={{ color: '#6c757d', marginBottom: '10px' }}>
                Complete Basic Information First
              </h3>
              <p style={{ color: '#6c757d', margin: 0 }}>
                Please fill in all Basic Information fields above to access Patient Log
              </p>
            </div>
          ) : (
            <>
            <div style={{ overflowX: 'auto' }}>
              <table style={tableStyle}>
              <thead style={{ backgroundColor: '#0077B6', color: 'white' }}>
                <tr>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '60px' }}>#</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '150px' }}>Patient's Name</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '100px' }}>Office</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '120px' }}>Appt. Date</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '120px' }}>Type of Visit</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '80px' }}>Call In</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '80px' }}>Call Out</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '80px' }}>Time</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '150px' }}>Remark</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '150px' }}>Other Duty</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '80px' }}>Action</th>
                </tr>
              </thead>
              <tbody>
                {patientRows.map((row) => (
                  <PatientRow
                    key={row.id}
                    row={row}
                    updatePatientRow={updatePatientRow}
                    removePatientRow={removePatientRow}
                    patientOfficeOptions={patientOfficeOptions}
                    getVisitTypeOptions={getVisitTypeOptions}
                    remarkOptions={remarkOptions}
                    otherDutyOptions={otherDutyOptions}
                    inputStyle={inputStyle}
                    buttonStyle={buttonStyle}
                  />
                ))}
              </tbody>
            </table>
            </div>

            <div style={{ display: 'flex', justifyContent: 'center', marginTop: '20px' }}>
              <button onClick={addPatientRow} style={buttonStyle}>
                + Add Row
              </button>
            </div>
            </>
          )}
        </div>

        {/* 일일 업무 보고서 */}
        <div style={sectionStyle}>
          <h2 style={{ color: '#0077B6', marginBottom: '15px' }}>📝 Daily Work Report</h2>
          {!isBasicInfoComplete() ? (
            <div style={{
              padding: '20px',
              textAlign: 'center',
              backgroundColor: '#f8f9fa',
              border: '2px dashed #dee2e6',
              borderRadius: '8px',
              margin: '10px 0'
            }}>
              <div style={{ fontSize: '24px', marginBottom: '10px' }}>🔒</div>
              <p style={{ color: '#6c757d', margin: 0 }}>
                Complete Basic Information first to access Daily Work Report
              </p>
            </div>
          ) : (
            <textarea
              value={formData.dailyWorkReport}
              onChange={(e) => updateFormData('dailyWorkReport', e.target.value)}
              rows={4}
              placeholder="Enter your daily work report here..."
              style={{
                ...inputStyle,
                minHeight: '100px',
                resize: 'vertical'
              }}
            />
          )}
        </div>

        {/* PDF 생성 버튼 */}
        <div style={{ display: 'flex', justifyContent: 'center', gap: '20px', marginBottom: '30px' }}>
          {!isBasicInfoComplete() ? (
            <div style={{
              padding: '20px',
              textAlign: 'center',
              backgroundColor: '#f8f9fa',
              border: '2px dashed #dee2e6',
              borderRadius: '8px',
              margin: '10px 0'
            }}>
              <div style={{ fontSize: '24px', marginBottom: '10px' }}>🔒</div>
              <p style={{ color: '#6c757d', margin: 0 }}>
                Complete Basic Information first to generate PDF
              </p>
            </div>
          ) : (
            <button 
              onClick={handleSubmit} 
              disabled={loading} 
              style={{ ...buttonStyle, backgroundColor: '#28a745' }}
            >
              {loading ? 'Submitting...' : '📄 Submit + Generate PDF'}
            </button>
          )}
        </div>

      </div>
    </div>
    </>
  );
}

export default function PatientLogPage() {
  return <PatientLogSystem />;
}
