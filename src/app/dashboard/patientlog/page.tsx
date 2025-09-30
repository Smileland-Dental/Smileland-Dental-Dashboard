'use client'

import React, { useState, useEffect, useCallback } from "react";
import { doc, setDoc, collection, getDocs, deleteDoc } from "firebase/firestore";
import { db } from "@/lib/firebase.config";

export default function PatientLogSystem() {
  // 기본 상태
  const [loading, setLoading] = useState(false);
  const [savedLogs, setSavedLogs] = useState([]);
  const [autoSaveStatus, setAutoSaveStatus] = useState(''); // 자동 저장 상태 표시

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

  // Office 옵션
  const workOfficeOptions = ['Ming', 'Bernard', 'Delano', 'Tulare', 'Visalia', 'Fresno', 'California', 'Ortho', 'Corporate'];
  const patientOfficeOptions = ['Ming', 'Bernard', 'Delano', 'Tulare', 'Visalia', 'Fresno', 'California', 'Ortho'];
  
  // Visit Type 옵션을 Office에 따라 동적으로 생성
  const getVisitTypeOptions = (office) => {
    if (office === 'Ortho') {
      return ['Adjustment', 'Consult', 'Partial Deband', 'Full Deband', 'Records', 'Bonding', 'Retainer Check', 'RPE Check'];
    } else {
      return ['New Patient', 'Recall', 'Tx', 'Emergency', 'RCRA'];
    }
  };

  // Remark 옵션
  const remarkOptions = ['NA', 'LMA', 'LMW', 'Disc', 'Wrong', 'Not Interested', 'Elsewhere'];

  // Other Duty 옵션
  const otherDutyOptions = [
    'Refer a friend', 'Postcards', 'Routing Slips', 'Nintendo Switch Raffle', 
    'Marketing Data', 'Reviews', 'Confirming', 'One Week\'s Reconfirming', 
    'Incoming Call Report', 'Monthly Report', 'Insurance Verifications', 
    'Medi-cal Eligibility', 'Accounts with Balances', 'Sending Replacement Staff', 
    'Booking ASL Interpreters', 'Lunch', 'Break', 'Other'
  ];

  // Document ID 생성 함수
  const generateDocId = (dutyDate, userName, workOffice, workHoursFrom, workHoursTo) => {
    return `${dutyDate}_${userName}_${workOffice}_${workHoursFrom}_${workHoursTo}`.replace(/[^a-zA-Z0-9_-]/g, '_');
  };

  // 자동 저장 함수 (debounced)
  const autoSave = useCallback(async () => {
    // 기본 정보가 모두 입력된 경우에만 자동 저장
    if (!formData.dutyDate || !formData.userName || !formData.workOffice || !formData.workHoursFrom || !formData.workHoursTo) {
      return;
    }

    try {
      setAutoSaveStatus('Saving...');
      
      const dataToSave = {
        ...formData,
        patientRows: patientRows.filter(row => 
          row.name || row.office || row.appt_date || row.visit_type || 
          row.call_in || row.call_out || row.time || row.remark || row.other_duty
        ),
        timestamp: new Date().toISOString(),
        autoSaved: true // 자동 저장 표시
      };

      const docId = generateDocId(formData.dutyDate, formData.userName, formData.workOffice, formData.workHoursFrom, formData.workHoursTo);
      await setDoc(doc(db, "patient-logs", docId), dataToSave);
      
      setAutoSaveStatus('Auto-saved ✅');
      
      // 2초 후 상태 메시지 제거
      setTimeout(() => setAutoSaveStatus(''), 2000);
      
    } catch (error) {
      console.error("Auto-save error:", error);
      setAutoSaveStatus('Auto-save failed ❌');
      setTimeout(() => setAutoSaveStatus(''), 3000);
    }
  }, [formData, patientRows]);

  // debounce를 위한 useEffect
  useEffect(() => {
    const timeoutId = setTimeout(() => {
      autoSave();
    }, 2000); // 2초 후 자동 저장

    return () => clearTimeout(timeoutId);
  }, [autoSave]);

  // 저장된 데이터를 현재 환자 로그에 로드하는 함수
  const loadExistingData = async () => {
    if (!formData.dutyDate || !formData.userName || !formData.workOffice || !formData.workHoursFrom || !formData.workHoursTo) {
      return;
    }

    try {
      const docId = generateDocId(formData.dutyDate, formData.userName, formData.workOffice, formData.workHoursFrom, formData.workHoursTo);
      const docRef = doc(db, "patient-logs", docId);
      const docSnap = await getDocs(collection(db, "patient-logs")).then(snapshot => {
        const foundDoc = snapshot.docs.find(d => d.id === docId);
        return foundDoc ? { exists: () => true, data: () => foundDoc.data() } : { exists: () => false };
      });

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

  // 기본 정보가 모두 입력되면 기존 데이터 로드
  useEffect(() => {
    loadExistingData();
  }, [formData.dutyDate, formData.userName, formData.workOffice, formData.workHoursFrom, formData.workHoursTo]);

  // 컴포넌트 마운트 시 저장된 로그 불러오기
  useEffect(() => {
    loadSavedLogs();
  }, []);

  // 저장된 로그 불러오기
  const loadSavedLogs = async () => {
    try {
      const querySnapshot = await getDocs(collection(db, "patient-logs"));
      const logs = [];
      querySnapshot.forEach((doc) => {
        logs.push({ id: doc.id, ...doc.data() });
      });
      setSavedLogs(logs.sort((a, b) => new Date(b.dutyDate) - new Date(a.dutyDate)));
    } catch (error) {
      console.error("Error loading logs:", error);
    }
  };

  // 폼 데이터 업데이트
  const updateFormData = (field, value) => {
    setFormData(prev => ({
      ...prev,
      [field]: value
    }));
  };

  // 환자 행 추가
  const addPatientRow = () => {
    const newId = Math.max(...patientRows.map(row => row.id)) + 1;
    setPatientRows([...patientRows, {
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
    }]);
  };

  // 환자 행 삭제
  const removePatientRow = (id) => {
    if (patientRows.length > 1) {
      setPatientRows(patientRows.filter(row => row.id !== id));
    }
  };

  // 환자 행 업데이트
  const updatePatientRow = (id, field, value) => {
    setPatientRows(patientRows.map(row => {
      if (row.id === id) {
        const updatedRow = { ...row, [field]: value };
        // Office가 변경되면 visit_type을 초기화
        if (field === 'office') {
          updatedRow.visit_type = '';
        }
        // Call In 또는 Call Out이 체크되면 현재 시간을 Time에 자동 입력
        if ((field === 'call_in' || field === 'call_out') && value === true)s {
          const now = new Date();
          const timeString = now.toTimeString().slice(0, 5); // HH:MM 형태
          updatedRow.time = timeString;
        }
        // Call In과 Call Out이 모두 체크 해제되면 Time을 비움
        if ((field === 'call_in' && value === false && !row.call_out) || 
            (field === 'call_out' && value === false && !row.call_in)) {
          updatedRow.time = '';
        }
        return updatedRow;
      }
      return row;
    }));
  };

  // Firebase에 저장
  const handleSave = async () => {
    if (!formData.dutyDate || !formData.userName || !formData.workOffice || !formData.workHoursFrom || !formData.workHoursTo) {
      alert('⚠️ Please fill in all basic information.');
      return;
    }

    try {
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
      
      alert('✅ Successfully saved!');
      loadSavedLogs();
    } catch (error) {
      console.error("Error saving document:", error);
      alert('❌ An error occurred while saving.');
    }
  };

  // PDF 생성 및 제출
  const handleSubmit = async () => {
    if (!formData.dutyDate || !formData.userName || !formData.workOffice || !formData.workHoursFrom || !formData.workHoursTo) {
      alert('⚠️ Please fill in all basic information.');
      return;
    }

    try {
      setLoading(true);

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

      const result = await response.json();

      if (result.success) {
        // 파일명 생성
        const date = formData.dutyDate || new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' });
        const name = formData.userName || 'Unknown';
        const office = formData.workOffice || 'Unknown';
        const filename = `${date}_${name}_${office}_Patient_Log`;
        
        // 인쇄 창 열기
        const printWindow = window.open('', '_blank');
        if (printWindow) {
          printWindow.document.write(result.html);
          printWindow.document.close();
          
          printWindow.onload = function() {
            setTimeout(() => {
              printWindow.print();
              alert(`📄 ${filename}.pdf\n\nPlease select "Save as PDF" in the print dialog to save!`);
            }, 1000);
          };
        }
      } else {
        throw new Error(result.error || 'PDF generation failed');
      }

    } catch (error) {
      console.error('Submit 오류:', error);
      alert('Submission failed: ' + error.message);
    } finally {
      setLoading(false);
    }
  };


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
    <div style={bodyStyle}>
      <div style={containerStyle}>
        {/* 헤더 */}
        <div style={{ position: 'relative' }}>
        <h1 style={headerStyle}>🌴 Patient Log System</h1>
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
              {formData.dutyDate && formData.userName && formData.workOffice && formData.workHoursFrom && formData.workHoursTo && (
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
                {patientRows.filter(row => row.appt_date && row.name).length}
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
                {patientRows.filter(row => row.call_in).length}
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
                {patientRows.filter(row => row.call_out).length}
              </span>
            </div>
          </div>

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
                {patientRows.map((row, index) => (
                  <tr key={row.id} style={{ backgroundColor: index % 2 === 0 ? '#f9f9f9' : 'white' }}>
                    <td style={{ padding: '8px', textAlign: 'center', fontWeight: 'bold' }}>
                      {index + 1}
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
                        {getVisitTypeOptions(row.office).map(type => (
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
                      {patientRows.length > 1 && (
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
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Add Row 버튼을 테이블 하단으로 이동 */}
          <div style={{ display: 'flex', justifyContent: 'center', marginTop: '20px' }}>
            <button onClick={addPatientRow} style={buttonStyle}>
              + Add Row
            </button>
          </div>
        </div>

        {/* 일일 업무 보고서 */}
        <div style={sectionStyle}>
          <h2 style={{ color: '#0077B6', marginBottom: '15px' }}>📝 Daily Work Report</h2>
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
        </div>

        {/* 액션 버튼들 */}
        <div style={{ display: 'flex', justifyContent: 'center', gap: '20px', marginBottom: '30px' }}>
          <button 
            onClick={handleSubmit} 
            disabled={loading} 
            style={{ ...buttonStyle, backgroundColor: '#28a745' }}
          >
            {loading ? 'Submitting...' : '📄 Submit + Generate PDF'}
          </button>
        </div>

      </div>
    </div>
  );
}
