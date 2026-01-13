'use client';

import React, { useState, useEffect } from 'react';
import { useRouter } from 'next/navigation';
import { collection, getDocs } from 'firebase/firestore';
import { db } from '@/lib/firebase.config';

interface TreatmentSheet {
  id: string;
  office: string;
  rdaName: string;
  date: string;
  lastUpdated: any;
  treatmentData: any[];
}

export default function ViewTreatmentSheets() {
  const router = useRouter();
  const [treatmentSheets, setTreatmentSheets] = useState<TreatmentSheet[]>([]);
  const [loading, setLoading] = useState(true);
  const [officeInput, setOfficeInput] = useState<string>('');
  const [dateFilter, setDateFilter] = useState<string>('');
  const [mounted, setMounted] = useState(false);

  // 클라이언트 사이드 마운트 확인
  useEffect(() => {
    setMounted(true);
  }, []);

  // Firebase에서 모든 treatment sheet 데이터 가져오기
  const fetchTreatmentSheets = async () => {
    try {
      setLoading(true);
      const requestsRef = collection(db, 'rda-treatment-sheets');
      const querySnapshot = await getDocs(requestsRef);
      
      const sheetsData: TreatmentSheet[] = [];
      querySnapshot.forEach((doc) => {
        const data = doc.data();
        
        // submitted: true인 데이터만 포함
        if (data.submitted === true) {
          sheetsData.push({
            id: doc.id,
            ...data
          } as TreatmentSheet);
        }
      });
      
      // 중복 제거: 같은 office, rdaName, date 조합의 데이터 중 가장 최신 것만 유지
      const uniqueSheets = new Map();
      sheetsData.forEach(sheet => {
        const key = `${sheet.office}_${sheet.rdaName}_${sheet.date}`;
        const existingSheet = uniqueSheets.get(key);
        
        if (!existingSheet) {
          uniqueSheets.set(key, sheet);
        } else {
          // 더 최신 데이터인지 확인
          const currentTime = sheet.lastUpdated?.toMillis ? sheet.lastUpdated.toMillis() : 
                            (sheet.lastUpdated ? new Date(sheet.lastUpdated).getTime() : 0);
          const existingTime = existingSheet.lastUpdated?.toMillis ? existingSheet.lastUpdated.toMillis() : 
                             (existingSheet.lastUpdated ? new Date(existingSheet.lastUpdated).getTime() : 0);
          
          if (currentTime > existingTime) {
            uniqueSheets.set(key, sheet);
          }
        }
      });
      
      // 중복 제거된 데이터를 배열로 변환하고 정렬
      const finalSheets = Array.from(uniqueSheets.values());
      finalSheets.sort((a, b) => {
        const timeA = a.lastUpdated?.toMillis ? a.lastUpdated.toMillis() : 
                     (a.lastUpdated ? new Date(a.lastUpdated).getTime() : 0);
        const timeB = b.lastUpdated?.toMillis ? b.lastUpdated.toMillis() : 
                     (b.lastUpdated ? new Date(b.lastUpdated).getTime() : 0);
        return timeB - timeA;
      });
      
      setTreatmentSheets(finalSheets);
      setLoading(false);
    } catch (error) {
      // 프로덕션에서는 상세한 에러 정보를 노출하지 않음
      if (process.env.NODE_ENV === 'development') {
        console.error('Failed to load treatment sheets:', error);
      }
      alert('Failed to load treatment sheets. Please try again.');
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchTreatmentSheets();
  }, []);

  // 입력된 office와 date로 필터링 및 정렬된 데이터
  // office가 입력되지 않으면 빈 배열 반환
  const filteredAndSortedSheets = treatmentSheets
    .filter(sheet => {
      // 공백 제거 및 대소문자 무시 비교
      const normalizedInputOffice = officeInput.trim();
      const normalizedSheetOffice = (sheet.office || '').trim();
      const officeMatch = normalizedInputOffice !== '' && 
                         normalizedSheetOffice.toLowerCase() === normalizedInputOffice.toLowerCase();
      
      const normalizedInputDate = dateFilter.trim();
      const dateMatch = normalizedInputDate === '' || sheet.date === normalizedInputDate;
      
      return officeMatch && dateMatch;
    })
    .sort((a, b) => {
      // 기본 정렬: date (내림차순), rdaName (오름차순)
      const dateCompare = b.date.localeCompare(a.date); // 최신 날짜가 먼저
      if (dateCompare !== 0) return dateCompare;
      
      return a.rdaName.localeCompare(b.rdaName); // rdaName 알파벳 순
    });

  const handleViewDetails = (sheet: TreatmentSheet) => {
    // 입력값 검증 및 정리
    const safeOffice = (sheet.office || '').trim().replace(/[^A-Za-z0-9]/g, '').substring(0, 10);
    const safeRdaName = (sheet.rdaName || '').trim().replace(/[^A-Za-z0-9\s]/g, '').substring(0, 50);
    const safeDate = (sheet.date || '').trim().replace(/[^0-9-]/g, '');
    
    // rda-treatment-sheet 페이지로 이동하면서 URL 파라미터로 데이터 전달
    const params = new URLSearchParams({
      office: safeOffice,
      rdaName: safeRdaName,
      date: safeDate,
      view: 'true'  // View Details로 접근했음을 표시
    });
    
    router.push(`/dashboard/forms/rda-treatment-sheet?${params.toString()}`);
  };

  const formatDate = (dateString: string) => {
    if (!dateString) return '';
    if (typeof window === 'undefined') return dateString; // 서버 사이드에서는 원본 반환
    
    try {
      // 날짜 문자열을 로컬 시간대로 파싱 (UTC로 해석하지 않음)
      const [year, month, day] = dateString.split('-').map(Number);
      const date = new Date(year, month - 1, day); // month는 0-based
      
      if (isNaN(date.getTime())) return dateString;
      
      // 캘리포니아 시간대로 날짜 포맷
      const formattedDate = date.toLocaleDateString('en-US', {
        timeZone: 'America/Los_Angeles',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit'
      });
      
      return formattedDate;
    } catch {
      return dateString;
    }
  };

  const formatLastUpdated = (timestamp: any) => {
    if (!timestamp) return 'Unknown';
    if (typeof window === 'undefined') return 'Unknown'; // 서버 사이드에서는 기본값 반환
    
    try {
      let date: Date;
      if (timestamp.toDate && typeof timestamp.toDate === 'function') {
        date = timestamp.toDate();
      } else if (timestamp.seconds) {
        date = new Date(timestamp.seconds * 1000);
      } else if (timestamp.toMillis && typeof timestamp.toMillis === 'function') {
        date = new Date(timestamp.toMillis());
      } else {
        date = new Date(timestamp);
      }
      
      if (isNaN(date.getTime())) return 'Unknown';
      
      // 캘리포니아 시간대로 날짜/시간 포맷
      return date.toLocaleString('en-US', {
        timeZone: 'America/Los_Angeles',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit'
      });
    } catch {
      return 'Unknown';
    }
  };

  const containerStyle: React.CSSProperties = {
    maxWidth: '1600px',
    margin: '0 auto',
    padding: '20px',
    fontFamily: 'Arial, sans-serif',
    backgroundColor: '#f8f9fa',
    minHeight: '100vh'
  };

  const headerStyle: React.CSSProperties = {
    textAlign: 'center',
    marginBottom: '30px',
    padding: '20px',
    backgroundColor: 'white',
    borderRadius: '10px',
    boxShadow: '0 2px 10px rgba(0,0,0,0.1)'
  };

  const searchContainerStyle: React.CSSProperties = {
    display: 'flex',
    gap: '15px',
    marginBottom: '20px',
    flexWrap: 'wrap',
    alignItems: 'center'
  };

  const inputStyle: React.CSSProperties = {
    padding: '10px 15px',
    border: '2px solid #ddd',
    borderRadius: '8px',
    fontSize: '16px',
    flex: '1',
    minWidth: '200px'
  };

  const selectStyle: React.CSSProperties = {
    padding: '10px 15px',
    border: '2px solid #ddd',
    borderRadius: '8px',
    fontSize: '16px',
    backgroundColor: 'white'
  };

  const buttonStyle: React.CSSProperties = {
    padding: '10px 20px',
    backgroundColor: '#007bff',
    color: 'white',
    border: 'none',
    borderRadius: '8px',
    fontSize: '16px',
    cursor: 'pointer',
    fontWeight: 'bold'
  };

  const tableContainerStyle: React.CSSProperties = {
    backgroundColor: 'white',
    borderRadius: '10px',
    boxShadow: '0 2px 10px rgba(0,0,0,0.1)',
    overflow: 'hidden'
  };

  const tableStyle: React.CSSProperties = {
    width: '100%',
    borderCollapse: 'collapse',
    fontSize: '14px'
  };

  const thStyle: React.CSSProperties = {
    backgroundColor: '#f8f9fa',
    padding: '15px 10px',
    textAlign: 'left',
    fontWeight: 'bold',
    borderBottom: '2px solid #dee2e6',
    cursor: 'pointer',
    userSelect: 'none'
  };

  const tdStyle: React.CSSProperties = {
    padding: '12px 10px',
    borderBottom: '1px solid #dee2e6',
    verticalAlign: 'top'
  };

  const rowStyle: React.CSSProperties = {
    cursor: 'pointer',
    transition: 'background-color 0.2s'
  };

  const rowHoverStyle: React.CSSProperties = {
    backgroundColor: '#f8f9fa'
  };

  const loadingStyle: React.CSSProperties = {
    textAlign: 'center',
    padding: '50px',
    fontSize: '18px',
    color: '#666'
  };

  const emptyStyle: React.CSSProperties = {
    textAlign: 'center',
    padding: '50px',
    fontSize: '18px',
    color: '#666',
    backgroundColor: 'white',
    borderRadius: '10px',
    boxShadow: '0 2px 10px rgba(0,0,0,0.1)'
  };

  if (loading) {
    return (
      <div style={containerStyle}>
        <div style={loadingStyle}>
          <div>⏳ Loading treatment sheets...</div>
        </div>
      </div>
    );
  }

  return (
    <div style={containerStyle}>
      <div style={headerStyle}>
        <h1 style={{
          color: '#2c3e50',
          textAlign: 'center',
          marginBottom: '30px',
          fontSize: '2.5em',
          fontWeight: 'bold',
          borderBottom: '3px solid #3498db',
          paddingBottom: '15px'
        }}>
          RDA/DA Treatment (Sealant) Review
        </h1>
      </div>

      {/* Office 및 Date 입력 컨트롤 */}
      <div style={searchContainerStyle}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '15px', flex: '1', flexWrap: 'wrap' }}>
          <label style={{
            fontSize: '16px',
            fontWeight: 'bold',
            color: '#333',
            whiteSpace: 'nowrap'
          }}>
            🏢 Enter Office:
          </label>
          <select
            value={officeInput}
            onChange={(e) => {
              // 입력값 검증: 알파벳과 숫자만 허용, 최대 10자
              const value = e.target.value.replace(/[^A-Za-z0-9]/g, '').substring(0, 10);
              setOfficeInput(value);
            }}
            style={{
              ...inputStyle,
              minWidth: '200px',
              fontSize: '16px',
              padding: '12px 15px'
            }}
          >
            <option value="">Select Office</option>
            {/* 기본 옵션들 */}
            <option value="A">A</option>
            <option value="B">B</option>
            <option value="C">C</option>
            <option value="D">D</option>
            <option value="E">E</option>
            <option value="F">F</option>
            <option value="G">G</option>
            <option value="H">H</option>
            {/* 데이터베이스에 존재하지만 기본 목록에 없는 office들 */}
            {treatmentSheets
              .map(s => s.office)
              .filter((office, index, self) => 
                office && 
                self.indexOf(office) === index && 
                !['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'].includes(office) &&
                /^[A-Za-z0-9]{1,10}$/.test(office) // 안전한 office 값만 허용
              )
              .map(office => (
                <option key={office} value={office}>{office}</option>
              ))
            }
          </select>
          
          {officeInput.trim() && (
            <>
              <label style={{
                fontSize: '16px',
                fontWeight: 'bold',
                color: '#333',
                whiteSpace: 'nowrap',
                marginLeft: '20px'
              }}>
                📅 Filter by Date:
              </label>
              <input
                type="date"
                value={dateFilter}
                onChange={(e) => {
                  // 날짜 형식 검증: YYYY-MM-DD 형식만 허용
                  const value = e.target.value.replace(/[^0-9-]/g, '');
                  if (value === '' || /^\d{4}-\d{2}-\d{2}$/.test(value)) {
                    setDateFilter(value);
                  }
                }}
                style={{
                  ...inputStyle,
                  minWidth: '200px',
                  fontSize: '16px',
                  padding: '12px 15px'
                }}
              />
            </>
          )}
          
          {officeInput.trim() && (
            <div style={{
              fontSize: '14px',
              color: '#666',
              padding: '10px 15px',
              backgroundColor: '#e9ecef',
              borderRadius: '8px',
              fontWeight: 'bold',
              marginLeft: '20px'
            }}>
              📊 Showing {filteredAndSortedSheets.length} sheet(s) for "{officeInput}" office
              {dateFilter && ` on ${dateFilter}`}
            </div>
          )}
        </div>
      </div>

      {/* 결과 표시 */}
      {treatmentSheets.length === 0 ? (
        <div style={emptyStyle}>
          <div>📭 No submitted treatment sheets found</div>
        </div>
      ) : officeInput.trim() === '' ? (
        <div style={emptyStyle}>
          <div>🏢 Please select an office to view treatment sheets</div>
        </div>
      ) : filteredAndSortedSheets.length === 0 ? (
        <div style={emptyStyle}>
          <div>
            🔍 No treatment sheets found for "{officeInput}" office
            {dateFilter && ` on ${dateFilter}`}
          </div>
          <div style={{ fontSize: '14px', marginTop: '10px' }}>
            Try selecting a different office{dateFilter && ' or date'}
          </div>
        </div>
      ) : (
        <div style={tableContainerStyle}>
          <table style={tableStyle}>
            <thead>
              <tr>
                <th style={thStyle}>
                  📅 Date
                </th>
                <th style={thStyle}>
                  🏢 Office
                </th>
                <th style={thStyle}>
                  👤 RDA/DA Name
                </th>
                <th style={thStyle}>
                  📊 Patients
                </th>
                <th style={thStyle}>
                  🕒 Last Updated
                </th>
                <th style={thStyle}>
                  🔗 Action
                </th>
              </tr>
            </thead>
            <tbody>
              {filteredAndSortedSheets.map((sheet, index) => {
                const patientCount = sheet.treatmentData.filter(row => 
                  row.patientName && row.patientName.trim() !== ''
                ).length;
                
                return (
                  <tr
                    key={sheet.id}
                    style={rowStyle}
                    onMouseEnter={(e) => {
                      e.currentTarget.style.backgroundColor = '#f8f9fa';
                    }}
                    onMouseLeave={(e) => {
                      e.currentTarget.style.backgroundColor = 'white';
                    }}
                  >
                    <td style={tdStyle}>
                      <strong>{formatDate(sheet.date)}</strong>
                    </td>
                    <td style={tdStyle}>
                      {sheet.office}
                    </td>
                    <td style={tdStyle}>
                      <strong>{sheet.rdaName}</strong>
                    </td>
                    <td style={tdStyle}>
                      <span style={{
                        backgroundColor: patientCount > 0 ? '#d4edda' : '#f8d7da',
                        color: patientCount > 0 ? '#155724' : '#721c24',
                        padding: '4px 8px',
                        borderRadius: '12px',
                        fontSize: '12px',
                        fontWeight: 'bold'
                      }}>
                        {patientCount} patients
                      </span>
                    </td>
                    <td style={tdStyle}>
                      <small style={{ color: '#666' }}>
                        {formatLastUpdated(sheet.lastUpdated)}
                      </small>
                    </td>
                    <td style={tdStyle}>
                      <button
                        onClick={() => handleViewDetails(sheet)}
                        style={{
                          ...buttonStyle,
                          padding: '8px 16px',
                          fontSize: '14px',
                          backgroundColor: '#28a745'
                        }}
                      >
                        👁️ View Details
                      </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

    </div>
  );
}
