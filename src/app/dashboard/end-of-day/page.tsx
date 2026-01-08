'use client';

import React, { useState, useEffect } from 'react';
import { db } from '@/lib/firebase.config';
import { collection, query, where, getDocs, Timestamp } from 'firebase/firestore';

export default function EndOfDay() {
  const [selectedOffice, setSelectedOffice] = useState('');
  const [pdfs, setPdfs] = useState<any[]>([]);
  const [loading, setLoading] = useState(false);
  const [filteredPdfs, setFilteredPdfs] = useState<any[]>([]);
  const [merging, setMerging] = useState(false);
  const [downloadingZip, setDownloadingZip] = useState(false);
  const [monthFilter, setMonthFilter] = useState('');

  // 오피스 옵션
  const officeOptions = ['Bernard', 'California', 'Call Center', 'Delano', 'Fresno', 'Janitor', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

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

  // PDF 목록 로드
  const loadPdfs = async () => {
    if (!selectedOffice) {
      setPdfs([]);
      setFilteredPdfs([]);
      return;
    }

    setLoading(true);
    try {
      // orderBy 없이 쿼리 (인덱스 불필요)
      const q = query(
        collection(db, 'pdf-documents'),
        where('office', '==', selectedOffice)
      );
      
      const querySnapshot = await getDocs(q);
      const pdfList: any[] = [];
      
      querySnapshot.forEach((doc) => {
        const data = doc.data();
        pdfList.push({
          id: doc.id,
          ...data,
          // createdAt이 없으면 현재 시간을 사용
          createdAt: data.createdAt || new Date()
        });
      });
      
      // 클라이언트 사이드에서 정렬 (파일명의 번호 순서)
      pdfList.sort((a, b) => {
        // 파일명에서 번호 추출 함수
        const extractNumber = (filename: string): number => {
          if (!filename) return 9999; // 번호가 없으면 뒤로
          // "1) ", "2) ", "8) " 형식의 번호 추출
          const match = filename.match(/^(\d+)\)\s/);
          return match ? parseInt(match[1], 10) : 9999;
        };
        
        const aNum = extractNumber(a.filename || '');
        const bNum = extractNumber(b.filename || '');
        
        // 번호가 같으면 createdAt 기준 최신순
        if (aNum === bNum) {
          if (a.createdAt?.toMillis && b.createdAt?.toMillis) {
            return b.createdAt.toMillis() - a.createdAt.toMillis();
          }
          const aTime = a.createdAt ? new Date(a.createdAt).getTime() : 0;
          const bTime = b.createdAt ? new Date(b.createdAt).getTime() : 0;
          return bTime - aTime;
        }
        
        return aNum - bNum; // 번호 순서대로 정렬
      });
      
      setPdfs(pdfList);
      setFilteredPdfs(pdfList);
    } catch (error: any) {
      console.error('Error loading PDFs:', error);
      const errorMessage = error?.message || '알 수 없는 오류';
      alert(`Error loading PDFs: ${errorMessage}`);
      setPdfs([]);
      setFilteredPdfs([]);
    } finally {
      setLoading(false);
    }
  };

  // Office 변경 시 PDF 목록 다시 로드
  useEffect(() => {
    loadPdfs();
  }, [selectedOffice]);


  // 캘리포니아 시간대의 오늘 날짜 가져오기
  const getCurrentCaliforniaDate = () => {
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

  // 날짜 필터링 (초기값: 캘리포니아 시간대의 오늘 날짜)
  const [dateFilter, setDateFilter] = useState(() => getCurrentCaliforniaDate());
  
  // Type 필터링
  const [typeFilter, setTypeFilter] = useState('');
  
  // Type 옵션
  const typeOptions = [
    'End of Day Fax Cover',
    'Daily Office Duty',
    'Attendance Tract',
    'Add On Treatment',
    'Lobby Inspection',
    'Restroom Inspection',
    'Patient Log',
    'RDA/DA (Sealant) Treatment',
    'Janitorial Daily Duty'
  ];
  
  useEffect(() => {
    let filtered = pdfs;
    
    // Date 필터 적용
    if (dateFilter) {
      filtered = filtered.filter(pdf => pdf.date === dateFilter);
    }
    
    // Month 필터 적용 (YYYY-MM 형식)
    if (monthFilter) {
      filtered = filtered.filter(pdf => {
        if (!pdf.date) return false;
        return pdf.date.startsWith(monthFilter);
      });
    }
    
    // Type 필터 적용
    if (typeFilter) {
      filtered = filtered.filter(pdf => pdf.type === typeFilter);
    }
    
    setFilteredPdfs(filtered);
  }, [dateFilter, monthFilter, typeFilter, pdfs]);

  // 시간 형식 변환
  const formatTime = (timestamp: any) => {
    if (!timestamp) return '';
    let date: Date;
    if (timestamp instanceof Timestamp) {
      date = timestamp.toDate();
    } else if (timestamp.toDate) {
      date = timestamp.toDate();
    } else {
      date = new Date(timestamp);
    }
    return date.toLocaleString('en-US', { 
      timeZone: 'America/Los_Angeles',
      month: '2-digit',
      day: '2-digit',
      year: 'numeric', 
      hour: '2-digit',
      minute: '2-digit'
    });
  };

  // PDF 병합 함수
  const handleMergePdfs = async () => {
    if (filteredPdfs.length === 0) {
      alert('No PDFs to merge.');
      return; 
    }

    setMerging(true);
    try {
      // 모든 필터링된 PDF의 URL 수집
      const pdfUrls = filteredPdfs.map(pdf => pdf.url).filter(url => url);
      
      if (pdfUrls.length === 0) {
        alert('No valid PDF URLs.');
        setMerging(false);
        return;
      }

      // 병합 API 호출
      const response = await fetch('/api/merge-pdfs', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ pdfUrls }),
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'PDF 병합 실패');
      }

      // 병합된 PDF blob 받기
      const blob = await response.blob();
      
      // 파일명 생성
      const dateStr = dateFilter || new Date().toISOString().split('T')[0];
      const typeStr = typeFilter ? `_${typeFilter.replace(/\s+/g, '_')}` : '';
      const filename = `${dateStr}_${selectedOffice}${typeStr}_End of Day Daily Check Out.pdf`;
      
      // 다운로드
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);

      alert(`${filteredPdfs.length} PDFs merged successfully.`);
    } catch (error: any) {
      console.error('Error merging PDFs:', error);
      alert(`Error merging PDFs: ${error?.message || 'Unknown error'}`);
    } finally {
      setMerging(false);
    }
  };

  // ZIP으로 모든 PDF 다운로드 함수
  const handleDownloadZip = async () => {
    if (filteredPdfs.length === 0) {
      alert('No PDFs to download.');
      return;
    }

    setDownloadingZip(true);
    try {
      // 모든 필터링된 PDF의 URL 수집
      const pdfUrls = filteredPdfs.map(pdf => pdf.url).filter(url => url);
      
      if (pdfUrls.length === 0) {
        alert('No valid PDF URLs.');
        setDownloadingZip(false);
        return;
      }

      // ZIP 다운로드 API 호출
      const response = await fetch('/api/download-pdfs-zip', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ pdfUrls }),
      });

      if (!response.ok) {
        let errorMessage = 'ZIP download failed';
        try {
          const errorData = await response.json();
          errorMessage = errorData.error || errorMessage;
        } catch (e) {
          errorMessage = `HTTP error! status: ${response.status}`;
        }
        throw new Error(errorMessage);
      }

      // ZIP blob 받기
      const zipBlob = await response.blob();
      
      // 파일명 생성
      const dateStr = monthFilter || dateFilter || new Date().toISOString().split('T')[0];
      const typeStr = typeFilter ? `_${typeFilter.replace(/\s+/g, '_')}` : '';
      const filename = `${dateStr}_${selectedOffice}${typeStr}.zip`;
      
      // 다운로드
      const url = window.URL.createObjectURL(zipBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);

      alert(`${filteredPdfs.length} PDFs downloaded to ZIP file.`);
    } catch (error: any) {
      console.error('Error downloading ZIP:', error);
      alert(`Error downloading ZIP: ${error?.message || 'Unknown error'}`);
    } finally {
      setDownloadingZip(false);
    }
  };

  // 월 목록 생성 (1~12)
  const getMonthOptions = (): Array<{ value: string; label: string }> => {
    const months: Array<{ value: string; label: string }> = [];
    const currentYear = new Date().getFullYear();
    for (let i = 1; i <= 12; i++) {
      const month = String(i).padStart(2, '0');
      const monthValue = `${currentYear}-${month}`;
      months.push({ value: monthValue, label: String(i) });
    }
    return months;
  };

  const styles = {
    body: {
      fontFamily: "'Roboto', sans-serif",
      background: '#ffe4e6',
      color: '#333',
      lineHeight: '1.6',
      minHeight: '100vh',
      margin: 0,
      padding: '20px'
    },
    container: {
      width: '67%',
      maxWidth: 'none',
      margin: '0 auto',
      background: 'white',
      borderRadius: '20px',
      boxShadow: '0 20px 40px rgba(0, 0, 0, 0.1)',
      padding: '30px'
    },
    header: {
      fontSize: '2rem',
      fontWeight: '700',
      marginBottom: '30px',
      color: '#4a6fa1',
      textAlign: 'center' as const
    },
    filterSection: {
      display: 'flex',
      gap: '20px',
      alignItems: 'center',
      marginBottom: '30px',
      padding: '20px',
      background: 'linear-gradient(135deg, rgba(74, 111, 161, 0.05) 0%, rgba(46, 58, 78, 0.05) 100%)',
      borderRadius: '15px'
    },
    label: {
      fontWeight: '600',
      color: '#4a6fa1',
      fontSize: '16px'
    },
    select: {
      width: '200px',
      padding: '10px 15px',
      border: '2px solid rgba(74, 111, 161, 0.2)',
      borderRadius: '10px',
      background: 'rgba(255, 255, 255, 0.9)',
      fontSize: '14px',
      fontWeight: '500'
    },
    input: {
      width: '160px',
      padding: '10px 15px',
      border: '2px solid rgba(74, 111, 161, 0.2)',
      borderRadius: '10px',
      background: 'rgba(255, 255, 255, 0.9)',
      fontSize: '14px',
      fontWeight: '500'
    },
    pdfList: {
      display: 'grid',
      gridTemplateColumns: 'repeat(auto-fill, minmax(300px, 1fr))',
      gap: '20px',
      marginTop: '20px'
    },
    pdfCard: {
      border: '1px solid #e1e5ea',
      borderRadius: '12px',
      padding: '20px',
      background: 'white',
      boxShadow: '0 4px 12px rgba(0, 0, 0, 0.1)',
      transition: 'transform 0.2s ease, box-shadow 0.2s ease',
      cursor: 'pointer'
    },
    pdfTitle: {
      fontSize: '16px',
      fontWeight: '600',
      color: '#2e3a4e',
      marginBottom: '10px',
      wordBreak: 'break-word' as const
    },
    pdfInfo: {
      fontSize: '14px',
      color: '#666',
      marginBottom: '5px'
    },
    emptyState: {
      textAlign: 'center' as const,
      padding: '60px 20px',
      color: '#666'
    },
    loadingState: {
      textAlign: 'center' as const,
      padding: '60px 20px',
      color: '#4a6fa1',
      fontSize: '18px'
    }
  };

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap');
      `}</style>
      <div style={styles.body}>
        <div style={styles.container}>
          <h2 style={styles.header}>
            End of Day Daily Check Out
          </h2>

          {/* 필터 섹션 */}
          <div style={styles.filterSection}>
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

            <label style={styles.label} htmlFor="dateFilter">📅 Date:</label>
            <input
              type="date"
              id="dateFilter"
              value={dateFilter}
              onChange={(e: any) => {
                setDateFilter(e.target.value);
                setMonthFilter(''); // 날짜 필터 선택 시 월 필터 초기화
              }}
              style={styles.input}
            />

            {selectedOffice && (
              <>
                <label style={styles.label} htmlFor="typeFilter">📄 Type:</label>
                <select
                  id="typeFilter"
                  value={typeFilter}
                  onChange={(e: any) => setTypeFilter(e.target.value)}
                  style={styles.select}
                >
                  <option value="">--All Types--</option>
                  {typeOptions.map(type => (
                    <option key={type} value={type}>{type}</option>
                  ))}
                </select>
                
                <label style={styles.label} htmlFor="monthFilter">📆 Month:</label>
                <select
                  id="monthFilter"
                  value={monthFilter}
                  onChange={(e: any) => {
                    setMonthFilter(e.target.value);
                    setDateFilter(''); // 월 필터 선택 시 날짜 필터 초기화
                  }}
                  style={styles.select}
                >
                  <option value="">--All Months--</option>
                  {getMonthOptions().map(month => (
                    <option key={month.value} value={month.value}>{month.label}</option>
                  ))}
                </select>
                {(dateFilter || monthFilter || typeFilter) && (
                  <button
                    onClick={() => {
                      setDateFilter('');
                      setMonthFilter('');
                      setTypeFilter('');
                    }}
                    style={{
                      padding: '10px 15px',
                      background: '#ff6b6b',
                      color: 'white',
                      border: 'none',
                      borderRadius: '8px',
                      cursor: 'pointer',
                      fontSize: '14px',
                      fontWeight: '600'
                    }}
                  >
                    Clear Filters
                  </button>
                )}
                {filteredPdfs.length > 0 && (
                  <>
                    <button
                      onClick={handleDownloadZip}
                      disabled={downloadingZip}
                      style={{
                        padding: '10px 15px',
                        background: downloadingZip ? '#95a5a6' : '#27ae60',
                        color: 'white',
                        border: 'none',
                        borderRadius: '8px',
                        cursor: downloadingZip ? 'not-allowed' : 'pointer',
                        fontSize: '14px',
                        fontWeight: '600',
                        marginLeft: '10px'
                      }}
                    >
                      {downloadingZip ? 'Downloading...' : `📥 Download ZIP (${filteredPdfs.length})`}
                    </button>
                    <button
                      onClick={handleMergePdfs}
                      disabled={merging}
                      style={{
                        padding: '10px 15px',
                        background: merging ? '#95a5a6' : '#4a90e2',
                        color: 'white',
                        border: 'none',
                        borderRadius: '8px',
                        cursor: merging ? 'not-allowed' : 'pointer',
                        fontSize: '14px',
                        fontWeight: '600',
                        marginLeft: '10px'
                      }}
                    >
                      {merging ? 'Merging...' : `📄 Merge (${filteredPdfs.length})`}
                    </button>
                  </>
                )}
              </>
            )}
        </div>

          {/* PDF 목록 */}
          {loading ? (
            <div style={styles.loadingState}>
              <div style={{fontSize: '24px', marginBottom: '10px'}}>⏳</div>
              Loading PDFs...
        </div>
          ) : !selectedOffice ? (
            null
          ) : filteredPdfs.length === 0 ? (
            null
          ) : (
            <div style={styles.pdfList}>
              {filteredPdfs.map((pdf, index) => (
                <div
                  key={pdf.id}
                  style={styles.pdfCard}
                  onClick={() => {
                    // p_route.ts에서 반환한 HTML이 있는 경우
                    if (pdf.html) {
                      try {
                        let htmlContent = pdf.html;
                        
                        // 🔒 보안: HTML 검증
                        if (typeof htmlContent !== 'string') {
                          alert('Invalid HTML format');
                          return;
                        }
                        
                        // HTML이 이스케이프되어 있는 경우 디코딩 시도
                        // &lt; -> <, &gt; -> > 등
                        if (htmlContent.includes('&lt;') || htmlContent.includes('&gt;')) {
                          // HTML 엔티티 디코딩
                          const textarea = document.createElement('textarea');
                          textarea.innerHTML = htmlContent;
                          htmlContent = textarea.value;
                        }
                        
                        // HTML이 올바른 형식인지 확인
                        const trimmedHtml = htmlContent.trim();
                        if (!trimmedHtml.startsWith('<!DOCTYPE') && !trimmedHtml.startsWith('<html')) {
                          // HTML이 JSON 문자열로 저장된 경우 파싱 시도
                          try {
                            const parsed = JSON.parse(htmlContent);
                            if (parsed && typeof parsed === 'string') {
                              htmlContent = parsed;
                            }
                          } catch (e) {
                            // JSON이 아니면 그대로 사용
                          }
                          
                          // 다시 확인
                          const reTrimmed = htmlContent.trim();
                          if (!reTrimmed.startsWith('<!DOCTYPE') && !reTrimmed.startsWith('<html')) {
                            alert('HTML 형식이 올바르지 않습니다.');
                            return;
                          }
                        }
                        
                        // 새 창 생성 및 HTML 직접 작성
                        const newWindow = window.open('', '_blank');
                        if (newWindow) {
                          // HTML을 직접 작성 (이스케이프하지 않음)
                          newWindow.document.open();
                          newWindow.document.write(htmlContent);
                          newWindow.document.close();
                        } else {
                          alert('팝업이 차단되었습니다. 팝업 차단을 해제해주세요.');
                        }
  } catch (error) {
                        console.error('Error opening PDF:', error);
                        alert('PDF를 열 수 없습니다.');
                      }
                    } else if (pdf.url && pdf.url.startsWith('http')) {
                      // 기존 URL 방식
                      window.open(pdf.url, '_blank');
                    }
                  }}
                  onMouseEnter={(e) => {
                    e.currentTarget.style.transform = 'translateY(-2px)';
                    e.currentTarget.style.boxShadow = '0 8px 20px rgba(0, 0, 0, 0.15)';
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.transform = 'translateY(0)';
                    e.currentTarget.style.boxShadow = '0 4px 12px rgba(0, 0, 0, 0.1)';
                  }}
                >
                  <div style={styles.pdfTitle}>{pdf.filename || 'PDF Document'}</div>
                  {pdf.createdAt && (
                    <div style={styles.pdfInfo}>
                      <strong>Created:</strong> {formatTime(pdf.createdAt)}
                    </div>
                  )}
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </>
  );
}
