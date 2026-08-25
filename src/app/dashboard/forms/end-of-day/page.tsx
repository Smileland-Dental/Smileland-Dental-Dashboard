'use client';

import React, { useState, useEffect, useRef } from 'react';
import { db, auth, storage } from '@/lib/firebase.config';
import { doc, getDoc } from 'firebase/firestore';
import {ref, listAll, getMetadata, getDownloadURL, deleteObject, type StorageReference} from 'firebase/storage';
import { onAuthStateChanged } from 'firebase/auth';

function getTodayInLA(): string {
  return new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' });
}

async function collectPdfItemsUnder(
  folderRef: StorageReference
): Promise<StorageReference[]> {
  const current = await listAll(folderRef);
  const directPdfItems = current.items.filter((item) =>
    item.name.toLowerCase().endsWith('.pdf')
  );
  const nestedItems = await Promise.all(
    current.prefixes.map((prefix) => collectPdfItemsUnder(prefix))
  );
  return [...directPdfItems, ...nestedItems.flat()];
}

export default function EndOfDay() {
  const todayStr = getTodayInLA();
  const [selectedOffice, setSelectedOffice] = useState('');
  const [yearMonthFilter, setYearMonthFilter] = useState(todayStr.slice(0, 7));
  const [dateFilter, setDateFilter] = useState(todayStr);
  const [pdfs, setPdfs] = useState<any[]>([]);
  const [loading, setLoading] = useState(false);
  const [viewerOpen, setViewerOpen] = useState(false);
  const [viewingPdf, setViewingPdf] = useState<any | null>(null);
  const [pdfBlobUrl, setPdfBlobUrl] = useState<string | null>(null);
  const [pdfLoading, setPdfLoading] = useState(false);
  const [pdfError, setPdfError] = useState(false);
  const pdfBlobUrlRef = useRef<string | null>(null);
  const [isAuthorized, setIsAuthorized] = useState<boolean | null>(null);
  const [selectedPaths, setSelectedPaths] = useState<string[]>([]);
  const [deleting, setDeleting] = useState(false);

  const officeOptions = ['Bernard', 'California', 'Call_Center', 'Delano', 'Fresno', 'Janitor', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

  const isValidOffice = (office: string): boolean => {
    return officeOptions.includes(office);
  };

  const handleOfficeChange = (newOffice: string) => {
    if (newOffice === '') {
      setSelectedOffice('');
      return;
    }

    if (!isValidOffice(newOffice)) {
      alert('Invalid office selection.');
      return;
    }

    setSelectedOffice(newOffice);
  };

  const handleYearMonthChange = (ym: string) => {
    if (!ym) return;
    setYearMonthFilter(ym);
    if (!dateFilter.startsWith(ym)) {
      const today = getTodayInLA();
      setDateFilter(today.startsWith(ym) ? today : `${ym}-01`);
    }
  };

  const handleDateChange = (d: string) => {
    if (!d) return;
    setDateFilter(d);
    setYearMonthFilter(d.slice(0, 7));
  };

  const loadPdfs = async () => {
    if (!selectedOffice || !dateFilter) {
      setPdfs([]);
      return;
    }

    setLoading(true);
    try {
      const listRef = ref(storage, `endofday-pdfs/${selectedOffice}/${dateFilter}/`);

      let pdfItems: StorageReference[] = [];
      try {
        pdfItems = await collectPdfItemsUnder(listRef);
      } catch {
        pdfItems = [];
      }

      const pdfList: any[] = await Promise.all(
        pdfItems.map(async (item) => {
          const meta = await getMetadata(item);
          const createdAt = meta.timeCreated ? new Date(meta.timeCreated) : new Date();
          return {
            path: item.fullPath,
            filename: item.name,
            createdAt
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
    } catch (error) {
      console.error('Failed to load PDFs from Storage:', error);
      alert('Error, please try again.');
      setPdfs([]);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (process.env.NODE_ENV === 'production' &&
        typeof window !== 'undefined' &&
        window.location.protocol !== 'https:') {
      window.location.href = window.location.href.replace('http:', 'https:');
      return;
    }

    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      try {
        if (!currentUser) {
          setIsAuthorized(false);
          if (typeof window !== 'undefined') {
            window.location.href = '/';
          }
          return;
        }

        const userDoc = await getDoc(doc(db, 'users', currentUser.uid));
        if (!userDoc.exists()) {
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

    return () => {
      unsubscribe();
    };
  }, []);

  useEffect(() => {
    setSelectedPaths([]);
  }, [selectedOffice, dateFilter, yearMonthFilter]);

  useEffect(() => {
    if (isAuthorized === true) {
      loadPdfs();
    }
  }, [selectedOffice, dateFilter, isAuthorized]);

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

  const formatTime = (timestamp: Date) => {
    if (!timestamp) return '';
    return timestamp.toLocaleString('en-US', {
      timeZone: 'America/Los_Angeles',
      month: '2-digit',
      day: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });
  };

  const normalizeStoragePath = (path: string): string | null => {
    if (typeof path !== 'string' || !path.trim()) return null;
    const p = path.trim().replace(/^\/*/, '');
    if (!p) return null;
    if (p.includes('..') || p.includes('//')) return null;
    if (p.length > 1024) return null;
    return p;
  };

  const isValidStoragePath = (path: string): boolean => {
    return normalizeStoragePath(path) !== null;
  };

  const isAllowedDeletePath = (path: string, office: string): boolean => {
    const p = normalizeStoragePath(path);
    if (!p || !isValidOffice(office)) return false;
    if (!p.startsWith(`endofday-pdfs/${office}/`)) return false;
    return p.toLowerCase().endsWith('.pdf');
  };

  const deleteStoragePaths = async (paths: string[]) => {
    const uniquePaths = Array.from(new Set(paths.filter(Boolean)));
    let deleted = 0;
    for (const path of uniquePaths) {
      if (!isAllowedDeletePath(path, selectedOffice)) continue;
      const storagePath = normalizeStoragePath(path);
      if (!storagePath) continue;
      try {
        await deleteObject(ref(storage, storagePath));
        deleted += 1;
      } catch (err) {
        console.error('Failed to delete:', storagePath, err);
      }
    }
    return deleted;
  };

  const toggleSelectPath = (path: string) => {
    if (!path) return;
    setSelectedPaths((prev) =>
      prev.includes(path) ? prev.filter((p) => p !== path) : [...prev, path]
    );
  };

  const toggleSelectAll = () => {
    const allPaths = pdfs.map((p) => p.path).filter(Boolean);
    const allSelected =
      allPaths.length > 0 && allPaths.every((p) => selectedPaths.includes(p));
    setSelectedPaths(allSelected ? [] : allPaths);
  };

  const handleDeleteSelected = async () => {
    if (!selectedOffice || deleting || selectedPaths.length === 0) return;
    const ok = window.confirm(
      `Delete ${selectedPaths.length} selected PDF(s)? This cannot be undone.`
    );
    if (!ok) return;

    setDeleting(true);
    try {
      const deleted = await deleteStoragePaths(selectedPaths);
      if (viewingPdf?.path && selectedPaths.includes(viewingPdf.path)) {
        closeViewer();
      }
      setSelectedPaths([]);
      await loadPdfs();
      alert(deleted > 0 ? `Deleted ${deleted} file(s).` : 'No files were deleted.');
    } catch (err) {
      console.error('Delete selected failed:', err);
      alert('Failed to delete selected files. Please try again.');
    } finally {
      setDeleting(false);
    }
  };

  const handleDeleteMonth = async () => {
    if (!selectedOffice || !yearMonthFilter || deleting) return;
    const ok = window.confirm(
      `Delete ALL PDFs for ${selectedOffice} in ${yearMonthFilter}? This cannot be undone.`
    );
    if (!ok) return;

    setDeleting(true);
    try {
      const officeRef = ref(storage, `endofday-pdfs/${selectedOffice}/`);
      const listed = await listAll(officeRef);
      const dayPrefixes = listed.prefixes.filter((prefix) => {
        const folder = prefix.name;
        return (
          folder.startsWith(`${yearMonthFilter}-`) &&
          /^\d{4}-\d{2}-\d{2}$/.test(folder)
        );
      });

      const allItems = (
        await Promise.all(dayPrefixes.map((prefix) => collectPdfItemsUnder(prefix)))
      ).flat();

      const deleted = await deleteStoragePaths(allItems.map((item) => item.fullPath));
      closeViewer();
      setSelectedPaths([]);
      await loadPdfs();
      alert(
        deleted > 0
          ? `Deleted ${deleted} file(s) for ${yearMonthFilter}.`
          : 'No files found to delete.'
      );
    } catch (err) {
      console.error('Delete month failed:', err);
      alert('Failed to delete month files. Please try again.');
    } finally {
      setDeleting(false);
    }
  };

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
    const i = pdfs.findIndex((p) => p.path === viewingPdf.path);
    if (i <= 0) return;
    const prev = pdfs[i - 1];
    if (prev?.path && isValidStoragePath(prev.path)) setViewingPdf(prev);
  };

  const goToNextPdf = () => {
    if (!viewingPdf?.path) return;
    const i = pdfs.findIndex((p) => p.path === viewingPdf.path);
    if (i < 0 || i >= pdfs.length - 1) return;
    const next = pdfs[i + 1];
    if (next?.path && isValidStoragePath(next.path)) setViewingPdf(next);
  };

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
        goToPrevPdf();
        return;
      }
      if (e.key === 'ArrowRight') {
        e.preventDefault();
        goToNextPdf();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [viewerOpen, viewingPdf, pdfs]);

  const monthDays = (() => {
    if (!yearMonthFilter || !/^\d{4}-\d{2}$/.test(yearMonthFilter)) {
      return { min: undefined as string | undefined, max: undefined as string | undefined };
    }
    const [y, m] = yearMonthFilter.split('-').map(Number);
    const lastDay = new Date(y, m, 0).getDate();
    return {
      min: `${yearMonthFilter}-01`,
      max: `${yearMonthFilter}-${String(lastDay).padStart(2, '0')}`
    };
  })();

  const styles = {
    body: {
      fontFamily: 'sans-serif',
      background: '#fff5f6',
      color: '#333',
      lineHeight: '1.6',
      minHeight: '100vh',
      margin: 0,
      padding: '20px'
    },
    container: {
      width: '90%',
      maxWidth: '1200px',
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
    layout: {
      display: 'flex',
      gap: '24px',
      alignItems: 'flex-start'
    },
    sidebar: {
      width: '240px',
      flexShrink: 0,
      display: 'flex',
      flexDirection: 'column' as const,
      gap: '18px',
      padding: '20px',
      background: 'linear-gradient(135deg, rgba(74, 111, 161, 0.05) 0%, rgba(46, 58, 78, 0.05) 100%)',
      borderRadius: '15px',
      boxSizing: 'border-box' as const
    },
    filterGroup: {
      display: 'flex',
      flexDirection: 'column' as const,
      gap: '8px'
    },
    main: {
      flex: 1,
      minWidth: 0
    },
    label: {
      fontWeight: '600',
      color: '#4a6fa1',
      fontSize: '16px'
    },
    select: {
      width: '100%',
      boxSizing: 'border-box' as const,
      padding: '10px 15px',
      border: '2px solid rgba(74, 111, 161, 0.2)',
      borderRadius: '10px',
      background: 'rgba(255, 255, 255, 0.9)',
      fontSize: '14px',
      fontWeight: '500'
    },
    input: {
      width: '100%',
      boxSizing: 'border-box' as const,
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
      marginTop: '0'
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
    return null;
  }

  const pdfNavIndex =
    viewerOpen && viewingPdf?.path
      ? pdfs.findIndex((p) => p.path === viewingPdf.path)
      : -1;
  const pdfNavShowPrev = viewerOpen && pdfNavIndex > 0;
  const pdfNavShowNext =
    viewerOpen && pdfNavIndex >= 0 && pdfNavIndex < pdfs.length - 1;

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

          <div style={styles.layout}>
            <aside style={styles.sidebar}>
              <div style={styles.filterGroup}>
                <label style={styles.label} htmlFor="office">🏢 Office</label>
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
              </div>

              <div style={styles.filterGroup}>
                <label style={styles.label} htmlFor="yearMonthFilter">📆 Year/Month</label>
                <input
                  type="month"
                  id="yearMonthFilter"
                  value={yearMonthFilter}
                  onChange={(e: any) => handleYearMonthChange(e.target.value)}
                  style={styles.input}
                  required
                />
              </div>

              <div style={styles.filterGroup}>
                <label style={styles.label} htmlFor="dateFilter">📅 Date</label>
                <input
                  type="date"
                  id="dateFilter"
                  value={dateFilter}
                  min={monthDays.min}
                  max={monthDays.max}
                  onChange={(e: any) => handleDateChange(e.target.value)}
                  style={styles.input}
                  required
                />
              </div>

              {selectedOffice && yearMonthFilter && (
                <button
                  type="button"
                  disabled={deleting}
                  onClick={handleDeleteMonth}
                  style={{
                    padding: '10px 15px',
                    background: deleting ? '#f0a0a0' : '#c0392b',
                    color: 'white',
                    border: 'none',
                    borderRadius: '8px',
                    cursor: deleting ? 'not-allowed' : 'pointer',
                    fontSize: '14px',
                    fontWeight: '600',
                    whiteSpace: 'nowrap' as const
                  }}
                >
                  {deleting ? 'Deleting…' : '🗑️ Delete Month'}
                </button>
              )}
            </aside>

            <div style={styles.main}>
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
              ) : !dateFilter ? (
                <div style={styles.emptyState}>
                  <div style={{fontSize: '24px', marginBottom: '10px'}}>📅</div>
                  Please select a date to view PDFs.
                </div>
              ) : pdfs.length === 0 ? (
                <div style={styles.emptyState}>
                  <div style={{fontSize: '24px', marginBottom: '10px'}}>📭</div>
                  No PDFs found for the selected filters.
                </div>
              ) : (
                <>
                  <div
                    style={{
                      display: 'flex',
                      gap: '12px',
                      alignItems: 'center',
                      marginBottom: '14px',
                      flexWrap: 'wrap' as const
                    }}
                  >
                    <label
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        gap: '8px',
                        fontSize: '14px',
                        fontWeight: '600',
                        color: '#4a6fa1',
                        cursor: 'pointer'
                      }}
                    >
                      <input
                        type="checkbox"
                        checked={
                          pdfs.length > 0 &&
                          pdfs.every((p) => selectedPaths.includes(p.path))
                        }
                        onChange={toggleSelectAll}
                        disabled={deleting}
                      />
                      Select all
                    </label>
                    <button
                      type="button"
                      disabled={deleting || selectedPaths.length === 0}
                      onClick={handleDeleteSelected}
                      style={{
                        padding: '8px 14px',
                        background:
                          deleting || selectedPaths.length === 0 ? '#f0a0a0' : '#e74c3c',
                        color: 'white',
                        border: 'none',
                        borderRadius: '8px',
                        cursor:
                          deleting || selectedPaths.length === 0
                            ? 'not-allowed'
                            : 'pointer',
                        fontSize: '14px',
                        fontWeight: '600'
                      }}
                    >
                      {deleting
                        ? 'Deleting…'
                        : `🗑️ Delete Selected (${selectedPaths.length})`}
                    </button>
                  </div>
                  <div style={styles.pdfList}>
                    {pdfs.map((pdf, index) => {
                      const checked = selectedPaths.includes(pdf.path);
                      return (
                        <div
                          key={`pdf-${index}-${pdf.filename ?? ''}`}
                          style={{
                            ...styles.pdfCard,
                            display: 'flex',
                            gap: '14px',
                            alignItems: 'flex-start',
                            borderColor: checked ? '#4a6fa1' : '#e1e5ea',
                            background: checked ? 'rgba(74, 111, 161, 0.06)' : 'white'
                          }}
                          onClick={() => openPdfViewer(pdf)}
                          onMouseEnter={(e) => {
                            e.currentTarget.style.transform = 'translateY(-2px)';
                            e.currentTarget.style.boxShadow =
                              '0 8px 20px rgba(0, 0, 0, 0.15)';
                          }}
                          onMouseLeave={(e) => {
                            e.currentTarget.style.transform = 'translateY(0)';
                            e.currentTarget.style.boxShadow =
                              '0 4px 12px rgba(0, 0, 0, 0.1)';
                          }}
                        >
                          <input
                            type="checkbox"
                            checked={checked}
                            disabled={deleting}
                            onClick={(e) => e.stopPropagation()}
                            onChange={(e) => {
                              e.stopPropagation();
                              toggleSelectPath(pdf.path);
                            }}
                            style={{
                              marginTop: '4px',
                              width: '18px',
                              height: '18px',
                              cursor: 'pointer',
                              flexShrink: 0
                            }}
                            aria-label={`Select ${pdf.filename || 'PDF'}`}
                          />
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={styles.pdfTitle}>
                              {pdf.filename || 'PDF Document'}
                            </div>
                            {pdf.createdAt && (
                              <div style={styles.pdfInfo}>
                                <strong>Created:</strong> {formatTime(pdf.createdAt)}
                              </div>
                            )}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </>
              )}
            </div>
          </div>
        </div>
      </div>

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


