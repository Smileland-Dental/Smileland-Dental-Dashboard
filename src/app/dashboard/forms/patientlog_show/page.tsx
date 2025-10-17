'use client'

import React, { useState, useEffect } from "react";
import { doc, setDoc, collection, getDocs, updateDoc, deleteDoc } from "firebase/firestore";
import { db } from "@/lib/firebase.config";
import { enableAllSecurityMeasures, sanitizeFirebaseDataClient, sanitizeFirebaseDocIdClient } from "@/lib/security-client";

export default function ShowCheckSystem() {
  // 보안 조치 활성화
  useEffect(() => {
    enableAllSecurityMeasures({
      disableConsole: true,
      disableRightClick: true,
      disableShortcuts: true,
      disableCopy: false,
      disableSelection: false,
      monitorDevTools: false
    });
  }, []);

  // 상태 관리
  const [loading, setLoading] = useState(false);
  const [appointments, setAppointments] = useState<any[]>([]);
  const [filteredAppointments, setFilteredAppointments] = useState<any[]>([]);
  const [startDate, setStartDate] = useState(new Date().toISOString().split('T')[0]);
  const [endDate, setEndDate] = useState(new Date().toISOString().split('T')[0]);
  const [selectedOffice, setSelectedOffice] = useState('');
  const [name, setName] = useState('');
  const [pdfLoading, setPdfLoading] = useState(false);
  const [submitStatus, setSubmitStatus] = useState('');
  const [progress, setProgress] = useState(0);

  // Office 옵션
  const officeOptions = ['All', 'Bernard', 'Call Center', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

  // 컴포넌트 마운트 시 데이터 로드
  useEffect(() => {
    loadAppointments();
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

  // Firebase에서 모든 환자 로그 불러오기
  const loadAppointments = async () => {
    try {
      setLoading(true);
      const querySnapshot = await getDocs(collection(db, "patient-logs"));
      const allAppointments: any[] = [];
      
      querySnapshot.forEach((doc) => {
        const data = doc.data();
        if (data.patientRows) {
          data.patientRows.forEach((row: any, index: number) => {
            if (row.appt_date && row.name) {
              allAppointments.push({
            ...row,
                docId: doc.id,
                rowIndex: index,
                dutyDate: data.dutyDate,
                userName: data.userName,
                workOffice: data.workOffice,
                workHoursFrom: data.workHoursFrom,
                workHoursTo: data.workHoursTo,
                showStatus: row.showStatus || 'pending' // pending, show, no-show
              });
            }
          });
        }
      });
      
      setAppointments(allAppointments);
    } catch (error) {
      // Production에서는 에러 로깅 비활성화
      if (process.env.NODE_ENV !== 'production') {
        console.error("Error loading appointments:", error);
      }
      alert('❌ 데이터 로드 중 오류가 발생했습니다.');
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

  // Show/No Show 상태 업데이트
  const updateShowStatus = async (appointment: any, newStatus: string) => {
    try {
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
      
      if (currentData && currentData.patientRows) {
        // 해당 row의 showStatus 업데이트
        const updatedPatientRows = currentData.patientRows.map((row: any, index: number) => {
          if (index === appointment.rowIndex) {
            return { ...row, showStatus: newStatus };
      }
      return row;
        });
        
        // Firebase 업데이트 (보안 검증 적용)
        const safeUpdateData = sanitizeFirebaseDataClient({
          patientRows: updatedPatientRows,
          lastUpdated: new Date().toISOString()
        });
        await updateDoc(docRef, safeUpdateData);
        
        // 로컬 상태 업데이트
        setAppointments(prev => 
          prev.map((apt: any) => 
            apt.docId === appointment.docId && apt.rowIndex === appointment.rowIndex
              ? { ...apt, showStatus: newStatus }
              : apt
          )
        );
        
        // Production에서는 로깅 비활성화
        if (process.env.NODE_ENV !== 'production') {
          console.log(`✅ ${appointment.name}의 상태가 ${newStatus}로 업데이트되었습니다.`);
        }
      }
    } catch (error) {
      // Production에서는 에러 로깅 비활성화
      if (process.env.NODE_ENV !== 'production') {
        console.error("Error updating show status:", error);
      }
      alert('❌ 상태 업데이트 중 오류가 발생했습니다.');
    }
  };

  // PDF 생성 및 제출
  const handleGeneratePDF = async () => {
    if (filteredAppointments.length === 0) {
      alert('⚠️ No appointments to generate PDF.');
      return;
    }

    try {
      setPdfLoading(true);
      setSubmitStatus('Saving...');
      setProgress(10);

      // PDF용 데이터 준비
      setSubmitStatus('Generating PDF...');
      setProgress(30);
      
      const pdfData = {
        startDate,
        endDate,
        selectedOffice: selectedOffice || 'All Offices',
        appointments: filteredAppointments,
        generatedBy: name || 'Supervisor',
        timestamp: new Date().toISOString()
      };

      // API로 PDF 생성
      const response = await fetch('/api/generate-show-check-pdf', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ showCheckData: pdfData }),
      });

      if (response.ok) {
        // PDF blob 받기
        setSubmitStatus('Processing PDF...');
        setProgress(60);
        const blob = await response.blob();
        
        // 파일명 생성
        const dateRange = startDate === endDate ? startDate : `${startDate}_to_${endDate}`;
        const office = selectedOffice || 'All';
        const supervisor = name || 'Supervisor';
        const filename = `${dateRange}_${office}_${supervisor}_Show_Check_Report.pdf`;
        
        setSubmitStatus('Cleaning up...');
        setProgress(80);
        
        // PDF 생성 후 해당 데이터 삭제
        try {
          await deleteProcessedAppointments();
        } catch (deleteError) {
          // Production에서는 에러 로깅 비활성화
          if (process.env.NODE_ENV !== 'production') {
            console.error('Error deleting processed data:', deleteError);
          }
        }
        
        setSubmitStatus('Complete!');
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
        
        // 2초 후 모달 닫기
        setTimeout(() => {
          setPdfLoading(false);
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
        console.error('PDF generation error:', error);
      }
      setSubmitStatus('❌ PDF generation failed: ' + ((error as any).message || 'Unknown error'));
      setProgress(0);
      setTimeout(() => {
        setPdfLoading(false);
        setSubmitStatus('');
      }, 3000);
    }
  };

  // 처리된 약속 데이터 삭제 함수
  const deleteProcessedAppointments = async () => {
    try {
      // 현재 필터링된 약속들의 document ID별로 그룹화
      const documentsToCheck = new Map();
      
      filteredAppointments.forEach(appointment => {
        const docId = appointment.docId;
        if (!documentsToCheck.has(docId)) {
          documentsToCheck.set(docId, []);
        }
        documentsToCheck.get(docId).push(appointment.rowIndex);
      });

      const deletedDocuments = [];
      const updatedDocuments = [];

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
            // 약속이 있었고 모든 약속이 처리됨 → 전체 document 삭제 (빈 row들도 함께)
            await deleteDoc(docRef);
            deletedDocuments.push(docId);
            // Production에서는 로깅 비활성화
            if (process.env.NODE_ENV !== 'production') {
              console.log(`🗑️ Document ${docId} completely deleted (all ${allAppointmentRows.length} appointments processed, including empty rows)`);
            }
          } else if (allAppointmentRows.length === 0) {
            // 애초에 약속이 없는 document → 삭제하지 않음
            if (process.env.NODE_ENV !== 'production') {
              console.log(`📋 Document ${docId} has no appointments, keeping as is`);
            }
          } else {
            // 아직 처리 안 된 약속이 있음 → 처리된 것들만 빈 상태로 초기화
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
                  other_duty: '',
                  showStatus: 'pending'
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
            if (process.env.NODE_ENV !== 'production') {
              console.log(`📝 Document ${docId} updated (${unprocessedAppointments.length} appointments remaining)`);
            }
          }
        }
      }

      // 로컬 상태 업데이트 - 처리된 약속들 제거
      setAppointments(prevAppointments => 
        prevAppointments.filter(apt => !filteredAppointments.some(filtered => 
          filtered.docId === apt.docId && filtered.rowIndex === apt.rowIndex
        ))
      );
      
      setFilteredAppointments([]);
      
      if (process.env.NODE_ENV !== 'production') {
        console.log(`✅ Processing complete: ${deletedDocuments.length} documents deleted, ${updatedDocuments.length} documents updated`);
      }

    } catch (error) {
      // Production에서는 에러 로깅 비활성화
      if (process.env.NODE_ENV !== 'production') {
        console.error('Error deleting processed appointments:', error);
      }
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
                  ⚠️ Please do not close
                </p>
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
                onChange={(e) => setName(e.target.value)}
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
                onChange={(e) => setStartDate(e.target.value)}
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
                onChange={(e) => setEndDate(e.target.value)}
                style={inputStyle}
              />
            </div>

            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Office:
              </label>
              <select
                value={selectedOffice}
                onChange={(e) => setSelectedOffice(e.target.value)}
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
          <button 
              onClick={handleGeneratePDF}
              disabled={pdfLoading}
              style={{
                backgroundColor: pdfLoading ? '#9ca3af' : '#3b82f6',
                color: 'white',
                border: 'none',
                padding: '15px 30px',
                borderRadius: '8px',
                fontSize: '16px',
                fontWeight: 'bold',
                cursor: pdfLoading ? 'not-allowed' : 'pointer',
                boxShadow: '0 4px 6px rgba(0, 0, 0, 0.1)',
                transition: 'all 0.2s ease'
              }}
            >
              {pdfLoading ? '📄 Generating PDF...' : '📄 Submit + Generate PDF'}
          </button>
        </div>
        )}
      </div>
    </div>
    </>
  );
}