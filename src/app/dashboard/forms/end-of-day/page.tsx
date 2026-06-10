'use client';

import React, { useState, useEffect, useRef } from 'react';
import { db, auth } from '@/lib/firebase.config';
import { doc, getDoc } from 'firebase/firestore';
import { getStorage, ref, listAll, getMetadata, getDownloadURL, getBlob } from 'firebase/storage';
import JSZip from 'jszip';
import { onAuthStateChanged } from 'firebase/auth';

function extractSubmissionDateFromStoragePath(fullPath: string, office: string): string | null {
  if (!fullPath || !office) return null;
  const parts = fullPath.split('/').filter(Boolean);
  if (parts.length < 4) return null;
  if (parts[0] !== 'endofday-pdfs' || parts[1] !== office) return null;
  const folder = parts[2];
  return /^\d{4}-\d{2}-\d{2}$/.test(folder) ? folder : null;
}

export default function EndOfDay() {
  const [selectedOffice, setSelectedOffice] = useState('');
  const [pdfs, setPdfs] = useState<any[]>([]);
  const [loading, setLoading] = useState(false);
  const [filteredPdfs, setFilteredPdfs] = useState<any[]>([]);
  const [viewerOpen, setViewerOpen] = useState(false);
  const [viewingPdf, setViewingPdf] = useState<any | null>(null);
  const [pdfBlobUrl, setPdfBlobUrl] = useState<string | null>(null);
  const [pdfLoading, setPdfLoading] = useState(false);
  const [pdfError, setPdfError] = useState(false);
  const pdfBlobUrlRef = useRef<string | null>(null);
  const [isAuthorized, setIsAuthorized] = useState<boolean | null>(null); // null: 확인 중, true: 인증됨, false: 인증 실패
  const [zipDownloading, setZipDownloading] = useState(false);

  // 오피스 옵션
  const officeOptions = ['Bernard', 'Ming', 'Delano', 'Janitor', 'California', 'Fresno', 'Ortho', 'Tulare', 'Visalia', 'Call_Center'];


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

    setSelectedOffice(newOffice);
  };

  // PDF 목록 로드: Storage에서 직접 목록 가져오기 (path가 항상 올바름)
  const loadPdfs = async () => {
    if (!selectedOffice) {
      setPdfs([]);
      setFilteredPdfs([]);
      return;
    }

    setLoading(true);
    try {
      const storage = getStorage();
      const listRef = ref(storage, `endofday-pdfs/${selectedOffice}/`);
      const result = await listAll(listRef);

      // Office 하위 폴더를 재귀 탐색해 모든 PDF 수집
      const collectPdfItems = async (
        currentRef: ReturnType<typeof ref>
      ): Promise<(typeof result.items)[number][]> => {
        const current = await listAll(currentRef);
        const directPdfItems = current.items.filter((item) =>
          item.name.toLowerCase().endsWith('.pdf')
        );
        const nestedItems = await Promise.all(
          current.prefixes.map((prefix) => collectPdfItems(prefix))
        );
        return [...directPdfItems, ...nestedItems.flat()];
      };

      const pdfItems = await collectPdfItems(listRef);

      const pdfList: any[] = await Promise.all(
        pdfItems.map(async (item) => {
          const meta = await getMetadata(item);
          const createdAt = meta.timeCreated ? new Date(meta.timeCreated) : new Date();
          const submissionDate = extractSubmissionDateFromStoragePath(
            item.fullPath,
            selectedOffice
          );
          const dateStr =
            submissionDate ?? createdAt.toISOString().slice(0, 10);
          return {
            path: item.fullPath,
            filename: item.name,
            createdAt,
            date: dateStr
          };
        })
      );

      pdfList.sort((a, b) => {
        const extractNumber = (filename: string): number => {
          if (!filename) return 9999;
          const match = filename.match(/^(\d+)\)\s/);
          return match ? parseInt(match[1], 10) : 9999;
        };
        const aNum = extractNumber(a.filename || '');
        const bNum = extractNumber(b.filename || '');
        if (aNum === bNum) {
          const aTime = a.createdAt ? new Date(a.createdAt).getTime() : 0;
          const bTime = b.createdAt ? new Date(b.createdAt).getTime() : 0;
          return bTime - aTime;
        }
        return aNum - bNum;
      });

      setPdfs(pdfList);
      setFilteredPdfs(pdfList);
    } catch (error) {
      console.error('Failed to load PDFs from Storage:', error);
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

        if (!['Director', 'HR', 'Manager'].includes(userData?.role)) {
          setIsAuthorized(false);
          if (typeof window !== 'undefined') {
            window.location.href = '/';
          }
          return;
        }

        setIsAuthorized(true);
      } catch {
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

  // PDF 로드: path에서 download URL을 받아 바로 표시
  useEffect(() => {
    if (!viewerOpen || !viewingPdf) {
      if (pdfBlobUrlRef.current) {
        if (pdfBlobUrlRef.current.startsWith('blob:')) {
          URL.revokeObjectURL(pdfBlobUrlRef.current);
        }
        pdfBlobUrlRef.current = null;
      }
      setPdfBlobUrl(null);
      setPdfLoading(false);
      setPdfError(false);
      return;
    }

    const pdf = viewingPdf;
    const usePath = pdf?.path && isValidStoragePath(pdf.path);

    if (!usePath) {
      setPdfError(true);
      setPdfLoading(false);
      return;
    }

    setPdfLoading(true);
    setPdfError(false);
    let cancelled = false;

    const setViewerSource = (sourceUrl: string, revokePreviousBlob = true) => {
      if (cancelled || !sourceUrl) return;
      if (pdfBlobUrlRef.current) {
        if (revokePreviousBlob && pdfBlobUrlRef.current.startsWith('blob:')) {
          URL.revokeObjectURL(pdfBlobUrlRef.current);
        }
        pdfBlobUrlRef.current = null;
      }
      pdfBlobUrlRef.current = sourceUrl;
      setPdfBlobUrl(sourceUrl);
      setPdfLoading(false);
    };

    const onError = (error?: unknown) => {
      if (!cancelled) {
        if (error) {
          console.error('Failed to open PDF:', error);
        }
        setPdfError(true);
        setPdfLoading(false);
      }
    };

    const storagePath = normalizeStoragePath(pdf.path);
    if (!storagePath) {
      onError();
      return () => {
        cancelled = true;
        setPdfBlobUrl(null);
      };
    }
    const storage = getStorage();
    const storageRef = ref(storage, storagePath);
    getDownloadURL(storageRef)
      .then((downloadUrl) => setViewerSource(downloadUrl, false))
      .catch((urlError) => {
        console.error('Failed to get download URL:', urlError);
        onError(urlError);
      });

    return () => {
      cancelled = true;
      if (pdfBlobUrlRef.current) {
        if (pdfBlobUrlRef.current.startsWith('blob:')) {
          URL.revokeObjectURL(pdfBlobUrlRef.current);
        }
        pdfBlobUrlRef.current = null;
      }
      setPdfBlobUrl(null);
    };
  }, [viewerOpen, viewingPdf]);

  const [dateFilter, setDateFilter] = useState('');
  
  useEffect(() => {
    let filtered = pdfs;
    
    if (dateFilter) {
      filtered = filtered.filter((pdf) => pdf.date === dateFilter);
    }
    
    setFilteredPdfs(filtered);
  }, [dateFilter, pdfs]);

  const formatTime = (timestamp: Date | { toDate?: () => Date } | number) => {
    if (!timestamp) return '';
    const date =
      timestamp instanceof Date
        ? timestamp
        : typeof timestamp === 'number'
          ? new Date(timestamp)
          : (timestamp as { toDate?: () => Date }).toDate?.() ?? new Date();
    return date.toLocaleString('en-US', { 
      timeZone: 'America/Los_Angeles',
      month: '2-digit',
      day: '2-digit',
      year: 'numeric', 
      hour: '2-digit',
      minute: '2-digit'
    });
  };

  // Storage path 정규화: 앞 슬래시 제거, 보안 검사 후 반환 (Firebase ref용)
  const normalizeStoragePath = (path: string): string | null => {
    if (typeof path !== 'string' || !path.trim()) return null;
    const p = path.trim().replace(/^\/*/, ''); // 앞 슬래시 제거
    if (!p) return null;
    if (p.includes('..') || p.includes('//')) return null;
    if (p.length > 1024) return null;
    return p;
  };

  // path가 있으면 true (정규화 가능한 경로)
  const isValidStoragePath = (path: string): boolean => {
    return normalizeStoragePath(path) !== null;
  };

  // PDF 뷰어 열기 (path만 사용)
  const openPdfViewer = (pdf: any) => {
    const hasPath = pdf?.path && isValidStoragePath(pdf.path);
    if (hasPath) {
      setViewingPdf(pdf);
      setViewerOpen(true);
    } else {
      alert('Cannot open the file.');
    }
  };

  const closeViewer = () => {
    setViewerOpen(false);
    setViewingPdf(null);
  };

  const goToPrevPdf = () => {
    if (!viewingPdf?.path) return;
    const i = filteredPdfs.findIndex((p) => p.path === viewingPdf.path);
    if (i <= 0) return;
    const prev = filteredPdfs[i - 1];
    if (prev?.path && isValidStoragePath(prev.path)) setViewingPdf(prev);
  };

  const goToNextPdf = () => {
    if (!viewingPdf?.path) return;
    const i = filteredPdfs.findIndex((p) => p.path === viewingPdf.path);
    if (i < 0 || i >= filteredPdfs.length - 1) return;
    const next = filteredPdfs[i + 1];
    if (next?.path && isValidStoragePath(next.path)) setViewingPdf(next);
  };

  const handleDownloadZip = async () => {
    if (!selectedOffice || filteredPdfs.length === 0 || zipDownloading) return;
    setZipDownloading(true);
    try {
      const zip = new JSZip();
      const storage = getStorage();
      const usedEntryNames = new Set<string>();

      for (let i = 0; i < filteredPdfs.length; i++) {
        const pdf = filteredPdfs[i];
        const storagePath = normalizeStoragePath(pdf.path);
        if (!storagePath) continue;
        const blob = await getBlob(ref(storage, storagePath));
        let entryName = (pdf.filename || `file-${i}.pdf`).replace(/[/\\]/g, '_');
        if (!entryName.toLowerCase().endsWith('.pdf')) entryName += '.pdf';
        let unique = entryName;
        let suffix = 0;
        while (usedEntryNames.has(unique)) {
          suffix += 1;
          const base = entryName.replace(/\.pdf$/i, '');
          unique = `${base}_${suffix}.pdf`;
        }
        usedEntryNames.add(unique);
        zip.file(unique, blob);
      }

      if (usedEntryNames.size === 0) {
        alert('No files could be added to the zip.');
        return;
      }

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const suffix = dateFilter ? dateFilter : 'all-dates';
      const filename = `endofday-pdfs-${selectedOffice}-${suffix}.zip`;
      const url = URL.createObjectURL(zipBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      a.rel = 'noopener';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('ZIP download failed:', err);
      alert('Failed to download zip. Please try again.');
    } finally {
      setZipDownloading(false);
    }
  };

  // 키보드: Escape 닫기, 좌우 화살표로 이전·다음 PDF
  useEffect(() => {
    if (!viewerOpen) return;
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        closeViewer();
        return;
      }
      if (e.key === 'ArrowLeft') {
        e.preventDefault();
        const i = viewingPdf?.path
          ? filteredPdfs.findIndex((p) => p.path === viewingPdf.path)
          : -1;
        if (i > 0) {
          const prev = filteredPdfs[i - 1];
          if (prev?.path && isValidStoragePath(prev.path)) setViewingPdf(prev);
        }
        return;
      }
      if (e.key === 'ArrowRight') {
        e.preventDefault();
        const i = viewingPdf?.path
          ? filteredPdfs.findIndex((p) => p.path === viewingPdf.path)
          : -1;
        if (i >= 0 && i < filteredPdfs.length - 1) {
          const next = filteredPdfs[i + 1];
          if (next?.path && isValidStoragePath(next.path)) setViewingPdf(next);
        }
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [viewerOpen, viewingPdf, filteredPdfs]);

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
      display: 'flex',
      flexDirection: 'column' as const,
      gap: '12px',
      marginTop: '20px'
    },
    pdfCard: {
      width: '100%',
      boxSizing: 'border-box' as const,
      border: '1px solid #e1e5ea',
      borderRadius: '12px',
      padding: '16px 20px',
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

  const pdfNavIndex =
    viewerOpen && viewingPdf?.path
      ? filteredPdfs.findIndex((p) => p.path === viewingPdf.path)
      : -1;
  const pdfNavShowPrev = viewerOpen && pdfNavIndex > 0;
  const pdfNavShowNext =
    viewerOpen && pdfNavIndex >= 0 && pdfNavIndex < filteredPdfs.length - 1;

  const navArrowButton: React.CSSProperties = {
    position: 'absolute',
    top: '50%',
    transform: 'translateY(-50%)',
    width: '48px',
    height: '48px',
    borderRadius: '50%',
    border: 'none',
    background: 'rgba(255, 255, 255, 0.95)',
    color: '#4a6fa1',
    fontSize: '28px',
    lineHeight: 1,
    cursor: 'pointer',
    zIndex: 1001,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontWeight: 'bold',
    boxShadow: '0 4px 14px rgba(0, 0, 0, 0.25)'
  };

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
              onChange={(e: any) => setDateFilter(e.target.value)}
              style={styles.input}
            />

            {selectedOffice && (
              <button
                type="button"
                disabled={zipDownloading || filteredPdfs.length === 0}
                onClick={handleDownloadZip}
                style={{
                  padding: '10px 15px',
                  background:
                    zipDownloading || filteredPdfs.length === 0 ? '#b8c5d6' : '#4a6fa1',
                  color: 'white',
                  border: 'none',
                  borderRadius: '8px',
                  cursor:
                    zipDownloading || filteredPdfs.length === 0 ? 'not-allowed' : 'pointer',
                  fontSize: '14px',
                  fontWeight: '600',
                  whiteSpace: 'nowrap' as const
                }}
              >
                {zipDownloading ? 'Preparing ZIP…' : '📦 Download ZIP'}
              </button>
            )}

            {selectedOffice && dateFilter && (
              <button
                onClick={() => setDateFilter('')}
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
                  key={`pdf-${index}-${pdf.filename ?? ''}`}
                  style={styles.pdfCard}
                  onClick={() => openPdfViewer(pdf)}
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

      {/* PDF 뷰어 모달 (클릭한 해당 PDF만 표시, iframe 없음) */}
      {viewerOpen && viewingPdf && (
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
              closeViewer();
            }
          }}
        >
          {pdfNavShowPrev && (
            <button
              type="button"
              aria-label="Previous PDF"
              onClick={(e) => {
                e.stopPropagation();
                goToPrevPdf();
              }}
              style={{ ...navArrowButton, left: '16px' }}
            >
              ‹
            </button>
          )}
          {pdfNavShowNext && (
            <button
              type="button"
              aria-label="Next PDF"
              onClick={(e) => {
                e.stopPropagation();
                goToNextPdf();
              }}
              style={{ ...navArrowButton, right: '16px' }}
            >
              ›
            </button>
          )}

          {/* 닫기 버튼 */}
          <button
            onClick={closeViewer}
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
            {viewingPdf.filename || 'PDF Document'}
          </div>

          {/* PDF 뷰어: blob URL로만 표시 (iframe 없음, Storage URL을 DOM에 노출하지 않음) */}
          {pdfLoading ? (
            <div style={{
              width: '90%',
              height: '90%',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              color: 'white',
              fontSize: '18px'
            }}>
              Loading PDF...
            </div>
          ) : pdfError || !pdfBlobUrl ? (
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
          ) : (
            <object
              data={pdfBlobUrl}
              type="application/pdf"
              style={{
                width: '90%',
                height: '90%',
                border: 'none',
                borderRadius: '8px',
                background: 'white'
              }}
              title="PDF Document"
            >
              <p style={{ padding: '20px', color: '#333' }}>
                Your browser does not support viewing PDFs. Please download the file.
              </p>
            </object>
          )}
        </div>
      )}
    </>
  );
}