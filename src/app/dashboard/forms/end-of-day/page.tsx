'use client';

import React, { useState, useEffect } from 'react';
import { db, auth } from '@/lib/firebase.config';
import { collection, query, where, getDocs, Timestamp, doc, getDoc } from 'firebase/firestore';
import { onAuthStateChanged } from 'firebase/auth';

export default function EndOfDay() {
  const [selectedOffice, setSelectedOffice] = useState('');
  const [pdfs, setPdfs] = useState<any[]>([]);
  const [loading, setLoading] = useState(false);
  const [filteredPdfs, setFilteredPdfs] = useState<any[]>([]);
  const [monthFilter, setMonthFilter] = useState('');
  const [typeOptions, setTypeOptions] = useState<string[]>([]);
  const [viewerOpen, setViewerOpen] = useState(false);
  const [currentPdfIndex, setCurrentPdfIndex] = useState(0);
  const [isAuthorized, setIsAuthorized] = useState<boolean | null>(null); // null: 확인 중, true: 인증됨, false: 인증 실패

  // 오피스 옵션
  const officeOptions = ['Appointment Show', 'Bernard', 'California', 'Call_Center', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

  // 보안: Office 값 검증
  const isValidOffice = (office: string): boolean => {
    return officeOptions.includes(office);
  };

  // Office 변경 처리
  const handleOfficeChange = (newOffice: string) => {
    // 빈 값으로 선택하면 비밀번호 없이 변경 허용 (초기화)
    if (newOffice === '') {
      setSelectedOffice('');
      return;
    }
    
    // 보안: 허용된 Office 값만 허용
    if (!isValidOffice(newOffice)) {
      alert('Invalid office selection.');
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
      
      // 실제 데이터에서 고유한 type 값들 추출
      const uniqueTypes = Array.from(new Set(
        pdfList
          .map(pdf => pdf.type)
          .filter(type => type && typeof type === 'string')
      )).sort();
      
      setTypeOptions(uniqueTypes);
      setPdfs(pdfList);
      setFilteredPdfs(pdfList);
    } catch (error: any) {
      // 보안: 상세한 에러 메시지 노출 최소화
      alert('Error, please try again.');
      setPdfs([]);
      setFilteredPdfs([]);
    } finally {
      setLoading(false);
    }
  };

  // 컴포넌트 마운트 시 사용자 인증 및 role 확인
  useEffect(() => {
    // 프로덕션 환경에서 HTTPS 강제 (클라이언트 사이드)
    if (process.env.NODE_ENV === 'production' && 
        typeof window !== 'undefined' && 
        window.location.protocol !== 'https:') {
      // HTTP로 접속한 경우 HTTPS로 리다이렉트
      window.location.href = window.location.href.replace('http:', 'https:');
      return;
    }

    // Firebase Auth 상태 변경 감지
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      try {
        if (!currentUser) {
          setIsAuthorized(false);
          if (typeof window !== 'undefined') {
            window.location.href = '/';
          }
          return;
        }

        // Firestore에서 사용자 role 확인
        const userDoc = await getDoc(doc(db, 'users', currentUser.uid));
        if (!userDoc.exists()) {
          setIsAuthorized(false);
          if (typeof window !== 'undefined') {
            window.location.href = '/';
          }
          return;
        }

        const userData = userDoc.data();

        if (userData?.role !== 'manager') {
          setIsAuthorized(false);
          if (typeof window !== 'undefined') {
            window.location.href = '/';
          }
          return;
        }

        setIsAuthorized(true);
      } catch (error: any) {
        setIsAuthorized(false);
        if (typeof window !== 'undefined') {
          window.location.href = '/';
        }
      }
    });

    // cleanup 함수
    return () => {
      unsubscribe();
    };
  }, []);

  // Office 변경 시 PDF 목록 다시 로드 (인증된 경우에만)
  useEffect(() => {
    if (isAuthorized === true) {
      loadPdfs();
    }
  }, [selectedOffice, isAuthorized]);


  // 키보드 이벤트 처리 (PDF 뷰어에서 좌우 화살표 키로 이동)
  useEffect(() => {
    if (!viewerOpen) return;

    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'ArrowLeft') {
        e.preventDefault();
        if (currentPdfIndex > 0) {
          setCurrentPdfIndex(currentPdfIndex - 1);
        }
      } else if (e.key === 'ArrowRight') {
        e.preventDefault();
        if (currentPdfIndex < filteredPdfs.length - 1) {
          setCurrentPdfIndex(currentPdfIndex + 1);
        }
      } else if (e.key === 'Escape') {
        e.preventDefault();
        setViewerOpen(false);
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [viewerOpen, currentPdfIndex, filteredPdfs.length]);


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
      filtered = filtered.filter(pdf => {
        if (!pdf.type) return false;
        // 정확히 일치하거나, 공백을 제거한 후 비교
        const pdfType = String(pdf.type).trim();
        const filterType = String(typeFilter).trim();
        return pdfType === filterType;
      });
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

  // 보안: URL 검증 함수
  const isValidFirebaseStorageUrl = (url: string): boolean => {
    if (typeof url !== 'string') return false;
    // Firebase Storage URL만 허용
    if (!url.startsWith('https://firebasestorage.googleapis.com')) return false;
    // URL 형식 검증
    try {
      const urlObj = new URL(url);
      // 도메인 검증
      if (urlObj.hostname !== 'firebasestorage.googleapis.com') return false;
      // 경로 검증 (path traversal 방지)
      if (urlObj.pathname.includes('..') || urlObj.pathname.includes('//')) return false;
      // 쿼리 파라미터 검증 (XSS 방지)
      if (urlObj.search && urlObj.search.length > 1000) return false;
      // URL 길이 제한
      if (url.length > 2048) return false;
      return true;
    } catch {
      return false;
    }
  };

  // PDF 뷰어 열기
  const openPdfViewer = (index: number) => {
    if (filteredPdfs[index] && filteredPdfs[index].url && isValidFirebaseStorageUrl(filteredPdfs[index].url)) {
      setCurrentPdfIndex(index);
      setViewerOpen(true);
    } else {
      alert('Cannot open the file.');
    }
  };

  // 이전 PDF로 이동
  const goToPreviousPdf = () => {
    if (currentPdfIndex > 0) {
      setCurrentPdfIndex(currentPdfIndex - 1);
    }
  };

  // 다음 PDF로 이동
  const goToNextPdf = () => {
    if (currentPdfIndex < filteredPdfs.length - 1) {
      setCurrentPdfIndex(currentPdfIndex + 1);
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
      fontFamily: "sans-serif",
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

  // 인증 확인 중이거나 인증 실패 시 처리
  if (isAuthorized === null) {
    return (
      <div style={styles.body}>
        <div style={styles.container}>
          <div style={styles.loadingState}>
            <div style={{fontSize: '24px', marginBottom: '10px'}}>⏳</div>
            Verifying authentication...
          </div>
        </div>
      </div>
    );
  }

  if (isAuthorized === false) {
    return null; // 리다이렉트 중이므로 아무것도 렌더링하지 않음
  }

  return (
    <>
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
            <div style={styles.emptyState}>
              <div style={{fontSize: '24px', marginBottom: '10px'}}>🏢</div>
              Please select an office to view PDFs.
            </div>
          ) : filteredPdfs.length === 0 ? (
            <div style={styles.emptyState}>
              <div style={{fontSize: '24px', marginBottom: '10px'}}>📭</div>
              No PDFs found for the selected filters.
            </div>
          ) : (
            <div style={styles.pdfList}>
              {filteredPdfs.map((pdf, index) => (
                <div
                  key={pdf.id}
                  style={styles.pdfCard}
                  onClick={() => {
                    const index = filteredPdfs.findIndex(p => p.id === pdf.id);
                    if (index !== -1) {
                      openPdfViewer(index);
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
                  <div style={styles.pdfTitle}>
                    {/* 보안: XSS 방지 - React가 자동으로 이스케이프하지만 명시적으로 처리 */}
                    {pdf.filename || 'PDF Document'}
                  </div>
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

      {/* PDF 뷰어 모달 */}
      {viewerOpen && filteredPdfs[currentPdfIndex] && (
        <div
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: 'rgba(0, 0, 0, 0.9)',
            zIndex: 1000,
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center'
          }}
          onClick={(e) => {
            if (e.target === e.currentTarget) {
              setViewerOpen(false);
            }
          }}
        >
          {/* 닫기 버튼 */}
          <button
            onClick={() => setViewerOpen(false)}
            style={{
              position: 'absolute',
              top: '20px',
              right: '20px',
              background: '#ff6b6b',
              color: 'white',
              border: 'none',
              borderRadius: '50%',
              width: '40px',
              height: '40px',
              fontSize: '20px',
              cursor: 'pointer',
              zIndex: 1001,
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              fontWeight: 'bold'
            }}
          >
            ×
          </button>

          {/* PDF 정보 */}
          <div
            style={{
              position: 'absolute',
              top: '20px',
              left: '50%',
              transform: 'translateX(-50%)',
              background: 'rgba(255, 255, 255, 0.9)',
              padding: '10px 20px',
              borderRadius: '8px',
              color: '#333',
              fontSize: '16px',
              fontWeight: '600',
              zIndex: 1001
            }}
          >
            {filteredPdfs[currentPdfIndex].filename || 'PDF Document'} ({currentPdfIndex + 1} / {filteredPdfs.length})
          </div>

          {/* 이전 버튼 */}
          {currentPdfIndex > 0 && (
            <button
              onClick={(e) => {
                e.stopPropagation();
                goToPreviousPdf();
              }}
              style={{
                position: 'absolute',
                left: '20px',
                top: '50%',
                transform: 'translateY(-50%)',
                background: 'rgba(255, 255, 255, 0.9)',
                color: '#333',
                border: 'none',
                borderRadius: '50%',
                width: '50px',
                height: '50px',
                fontSize: '24px',
                cursor: 'pointer',
                zIndex: 1001,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                fontWeight: 'bold'
              }}
            >
              ‹
            </button>
          )}

          {/* 다음 버튼 */}
          {currentPdfIndex < filteredPdfs.length - 1 && (
            <button
              onClick={(e) => {
                e.stopPropagation();
                goToNextPdf();
              }}
              style={{
                position: 'absolute',
                right: '20px',
                top: '50%',
                transform: 'translateY(-50%)',
                background: 'rgba(255, 255, 255, 0.9)',
                color: '#333',
                border: 'none',
                borderRadius: '50%',
                width: '50px',
                height: '50px',
                fontSize: '24px',
                cursor: 'pointer',
                zIndex: 1001,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                fontWeight: 'bold'
              }}
            >
              ›
            </button>
          )}

          {/* PDF iframe viewer */}
          {filteredPdfs[currentPdfIndex]?.url ? (
            <iframe
              src={`${filteredPdfs[currentPdfIndex].url}#toolbar=0`}
              style={{
                width: '90%',
                height: '90%',
                border: 'none',
                borderRadius: '8px',
                background: 'white'
              }}
              title="PDF Viewer"
              allow="fullscreen"
            />
          ) : (
            <div style={{
              width: '90%',
              height: '90%',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              color: 'white',
              fontSize: '18px'
            }}>
              Cannot open the file.
            </div>
          )}
        </div>
      )}
    </>
  );
}

