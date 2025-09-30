'use client'

import React, { useState, useEffect } from "react";
import { doc, setDoc, collection, getDocs, updateDoc, deleteDoc } from "firebase/firestore";
import { db } from "@/lib/firebase.config";

export default function ShowCheckSystem() {
  // 상태 관리
  const [loading, setLoading] = useState(false);
  const [appointments, setAppointments] = useState<any[]>([]);
  const [filteredAppointments, setFilteredAppointments] = useState<any[]>([]);
  const [selectedDate, setSelectedDate] = useState(new Date().toISOString().split('T')[0]);
  const [selectedOffice, setSelectedOffice] = useState('');
  const [supervisorName, setSupervisorName] = useState('');
  const [pdfLoading, setPdfLoading] = useState(false);

  // Office 옵션
  const officeOptions = ['All', 'Ming', 'Bernard', 'Delano', 'Tulare', 'Visalia', 'Fresno', 'California', 'Ortho'];

  // 컴포넌트 마운트 시 데이터 로드
  useEffect(() => {
    loadAppointments();
  }, []);

  // 필터 변경 시 데이터 필터링
  useEffect(() => {
    filterAppointments();
  }, [appointments, selectedDate, selectedOffice]);

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
      console.error("Error loading appointments:", error);
      alert('❌ 데이터 로드 중 오류가 발생했습니다.');
    } finally {
      setLoading(false);
    }
  };

  // 필터링 함수
  const filterAppointments = () => {
    let filtered = appointments;
    
    // 날짜 필터
    if (selectedDate) {
      filtered = filtered.filter(apt => apt.appt_date === selectedDate);
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
      const docRef = doc(db, "patient-logs", appointment.docId);
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
        
        // Firebase 업데이트
        await updateDoc(docRef, {
          patientRows: updatedPatientRows,
          lastUpdated: new Date().toISOString()
        });
        
        // 로컬 상태 업데이트
        setAppointments(prev => 
          prev.map(apt => 
            apt.docId === appointment.docId && apt.rowIndex === appointment.rowIndex
              ? { ...apt, showStatus: newStatus }
              : apt
          )
        );
        
        console.log(`✅ ${appointment.name}의 상태가 ${newStatus}로 업데이트되었습니다.`);
      }
    } catch (error) {
      console.error("Error updating show status:", error);
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

      // PDF용 데이터 준비
      const pdfData = {
        selectedDate,
        selectedOffice: selectedOffice || 'All Offices',
        appointments: filteredAppointments,
        generatedBy: supervisorName || 'Supervisor',
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

      const result = await response.json();

      if (result.success) {
        // 파일명 생성
        const date = selectedDate || new Date().toISOString().split('T')[0];
        const office = selectedOffice || 'All';
        const supervisor = supervisorName || 'Supervisor';
        const filename = `${date}_${office}_${supervisor}_Show_Check_Report`;
        
        // 인쇄 창 열기
        const printWindow = window.open('', '_blank');
        if (printWindow) {
          printWindow.document.write(result.html);
          printWindow.document.close();
          
          printWindow.onload = function() {
            setTimeout(async () => {
              printWindow.print();
              
              // PDF 생성 후 해당 데이터 삭제
              try {
                await deleteProcessedAppointments();
                alert(`📄 ${filename}.pdf\n\nPlease select "Save as PDF" in the print dialog to save!\n\n✅ Processed appointment data has been deleted from the database.`);
              } catch (deleteError) {
                console.error('Error deleting processed data:', deleteError);
                alert(`📄 ${filename}.pdf\n\nPlease select "Save as PDF" in the print dialog to save!\n\n⚠️ PDF generated successfully, but failed to delete processed data.`);
              }
            }, 1000);
          };
        }
      } else {
        throw new Error(result.error || 'PDF generation failed');
      }

    } catch (error) {
      console.error('PDF generation error:', error);
      alert('PDF generation failed: ' + error.message);
    } finally {
      setPdfLoading(false);
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
        const docRef = doc(db, "patient-logs", docId);
        
        // 현재 document 데이터 가져오기
        const querySnapshot = await getDocs(collection(db, "patient-logs"));
        let currentData = null;
        
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
            console.log(`🗑️ Document ${docId} completely deleted (all ${allAppointmentRows.length} appointments processed, including empty rows)`);
          } else if (allAppointmentRows.length === 0) {
            // 애초에 약속이 없는 document → 삭제하지 않음
            console.log(`📋 Document ${docId} has no appointments, keeping as is`);
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

            await updateDoc(docRef, {
              patientRows: updatedPatientRows,
              lastUpdated: new Date().toISOString()
            });
            updatedDocuments.push(docId);
            console.log(`📝 Document ${docId} updated (${unprocessedAppointments.length} appointments remaining)`);
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
      
      console.log(`✅ Processing complete: ${deletedDocuments.length} documents deleted, ${updatedDocuments.length} documents updated`);

    } catch (error) {
      console.error('Error deleting processed appointments:', error);
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

  const getStatusColor = (status) => {
    switch (status) {
      case 'show': return '#d4edda';
      case 'no-show': return '#f8d7da';
      default: return '#fff3cd';
    }
  };

  const getStatusText = (status) => {
    switch (status) {
      case 'show': return 'Show ✅';
      case 'no-show': return 'No Show ❌';
      default: return 'Pending ⏳';
    }
  };

  return (
    <div style={bodyStyle}>
      <div style={containerStyle}>
        {/* 헤더 */}
        <h1 style={headerStyle}>📋 Show/No Show Check System</h1>

        {/* 필터 섹션 */}
        <div style={sectionStyle}>
          <h2 style={{ color: '#0077B6', marginBottom: '20px' }}>🔍 Filter Appointments</h2>
          
          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap', marginBottom: '20px' }}>
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Appointment Date:
              </label>
              <input
                type="date"
                value={selectedDate}
                onChange={(e) => setSelectedDate(e.target.value)}
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
                {officeOptions.map(office => (
                  <option key={office} value={office}>{office}</option>
                ))}
              </select>
            </div>
            
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Supervisor Name:
              </label>
              <input
                type="text"
                value={supervisorName}
                onChange={(e) => setSupervisorName(e.target.value)}
                placeholder="Enter supervisor name"
                style={inputStyle}
              />
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
            {selectedDate && ` | Date: ${selectedDate}`}
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
              <table style={tableStyle}>
                <thead style={{ backgroundColor: '#0077B6', color: 'white' }}>
                  <tr>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Patient Name</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Office</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Appt. Date</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Visit Type</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Time</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Staff</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Work Office</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Status</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredAppointments.map((appointment, index) => (
                    <tr 
                      key={`${appointment.docId}-${appointment.rowIndex}`} 
                      style={{ 
                        backgroundColor: index % 2 === 0 ? getStatusColor(appointment.showStatus) : 'white',
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
                      <td style={{ padding: '12px 8px', textAlign: 'center' }}>
                        {appointment.time || '-'}
                      </td>
                      <td style={{ padding: '12px 8px', textAlign: 'center' }}>
                        {appointment.userName}
                      </td>
                      <td style={{ padding: '12px 8px', textAlign: 'center' }}>
                        {appointment.workOffice}
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
                            Reset
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
  );
}
