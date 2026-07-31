'use client'

import React, { useState, useEffect, useCallback } from "react";
import { doc, setDoc, getDoc, deleteDoc, onSnapshot } from "firebase/firestore";
import { db, auth } from "@/lib/firebase.config";
import { getStorage, ref, uploadBytes } from 'firebase/storage';
import { onAuthStateChanged } from 'firebase/auth';
import { pdf, Document, Page, View, Text, StyleSheet } from '@react-pdf/renderer';

export default function RestroomInspection() {
  const [loading, setLoading] = useState(false);
  const [submitStatus, setSubmitStatus] = useState('');
  const [progress, setProgress] = useState(0);
  const [isUpdatingFromFirebase, setIsUpdatingFromFirebase] = useState(false);
  const [userOfficesOptions, setUserOfficesOptions] = useState<string[]>([]);
  
  const [userSessionId] = useState(() => Math.random().toString(36).substr(2, 9));
  
  const [lastSavedData, setLastSavedData] = useState({});
  
  const getCurrentCaliforniaTime = () => {
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
  
  const [inspectionDate, setInspectionDate] = useState('');

  const [selectedOffice, setSelectedOffice] = useState('');
  const [selectedRestroom, setSelectedRestroom] = useState('');
  
  const officeOptions = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];
  const restroomOptions = ['1', '2', '3'];

  const OFFICE_CHECK_OPTIONS = {
    "Bernard": ["Carmen", "Elisa", "Ranjit"],
    "California": ["Helen", "Kindal", "Liz", "Alba"],
    "Delano": ["Helen", "Jasmine", "Leana"],
    "Fresno": ["Cynthia"],
    "Ming": ["Hopie", "Kindal", "Marbella"],
    "Ortho": ["Kindal"],
    "Tulare": ["Crystal", "Dianne", "Melissa"],
    "Visalia": ["Abby", "Dianne", "Jessica", "Alondra", 'Renee']
  };

  const RESTROOM_TASK_COLUMN_START = 3;
  const RESTROOM_TASK_COLUMN_COUNT = 9;

  const hasCheckColumnValue = (value: unknown): boolean =>
    value !== false && value !== null && value !== undefined && String(value).trim() !== '';

  const setRowTaskCheckboxes = (data: Record<string, unknown>, rowNumber: string, checked: boolean) => {
    for (let i = 0; i < RESTROOM_TASK_COLUMN_COUNT; i++) {
      data[`Row${rowNumber}_Col${RESTROOM_TASK_COLUMN_START + i}`] = checked;
    }
  };

  const FIRESTORE_META_KEYS = new Set([
    'inspectionDate', 'selectedOffice', 'selectedRestroom', 'timestamp', 'autoSaved', 'lastUpdatedBy',
  ]);

  const filterFirebaseData = (data: Record<string, unknown>): Record<string, unknown> => {
    const filtered: Record<string, unknown> = {};
    const rowFieldPattern = /^Row\d+_(Check|CheckedTime|Col\d+)$/;
    for (const [key, value] of Object.entries(data)) {
      if (FIRESTORE_META_KEYS.has(key) || !rowFieldPattern.test(key)) continue;
      filtered[key] = value;
    }
    return filtered;
  };

  const COLUMN_NAMES = [
    'Time', 'Check', 'Pick up Paper', 'Wipe Sinks and Mirrors', 'Wipe Toilets',
    'Wipe Baby Table', 'Empty Trash', 'Toilet Paper', 'Soap',
    'Toilet Seat Covers', 'Refresh Spray', 'Checked Time'
  ];

  const ROW_HEADERS = [
    'Manager Inspection',
    '8 am', '9 am', '10 am', '11 am',
    'Manager Inspection',
    '12 pm', '1 pm', 'Sweep/Mop', '2 pm', '3 pm',
    'Manager Inspection',
    '4 pm', '5 pm', '6 pm', 'Sweep/Mop', '7 pm',
    'Deep Clean Manager Inspection'
  ];

  const pdfStyles = StyleSheet.create({
    page: { padding: 22, fontFamily: 'Helvetica', fontSize: 8 },
    header: { marginBottom: 10, borderBottomWidth: 2, borderColor: '#333', paddingBottom: 6, alignItems: 'center' },
    headerTitle: { fontSize: 14, fontWeight: 'bold', marginBottom: 4 },
    headerSub: { fontSize: 10 },
    row: { flexDirection: 'row', borderBottomWidth: 0.5, borderColor: '#333' },
    cell: { padding: 3, fontSize: 6, flex: 1, borderRightWidth: 0.5, borderColor: '#333', justifyContent: 'center', alignItems: 'center' },
    cellFlex2: { flex: 2 },
    cellFlex4: { flex: 4 },
    cellFlex5: { flex: 5 },
    cellFlex1: { flex: 1 },
    cellBold: { fontWeight: 'bold' },
    cellGray: { backgroundColor: '#f0f0f0' },
    footer: { marginTop: 15, paddingTop: 8, borderTopWidth: 1, borderColor: '#ddd', alignItems: 'center', fontSize: 7, color: '#666' },
  });

  function safeStr(v: unknown, max: number): string {
    if (v == null) return '';
    return String(v).trim().slice(0, max).replace(/[<>]/g, '');
  }

  function isChecked(v: unknown): boolean {
    return v === true || v === 1 || (typeof v === 'string' && (v === 'true' || v === '1'));
  }

  function createRestroomPDFDocument(props: {
    safeInspectionDate: string;
    safeSelectedOffice: string;
    safeSelectedRestroom: string;
    restroomData: Record<string, unknown>;
    generatedDate: string;
  }) {
    const { safeInspectionDate, safeSelectedOffice, safeSelectedRestroom, restroomData, generatedDate } = props;
    const s = pdfStyles;

    const row1 = React.createElement(View, { key: 'r1', style: [s.row, s.cellGray] },
      React.createElement(View, { style: [s.cell, s.cellFlex2] }, React.createElement(Text, null, '')),
      React.createElement(View, { style: [s.cell, s.cellFlex4, s.cellBold] }, React.createElement(Text, null, 'SPOTLESS')),
      React.createElement(View, { style: [s.cell, s.cellFlex5, s.cellBold] }, React.createElement(Text, null, 'STOCKED')),
      React.createElement(View, { style: [s.cell, s.cellFlex1] }, React.createElement(Text, null, '')),
    );

    const row2 = React.createElement(View, { key: 'r2', style: [s.row, s.cellGray] },
      ...COLUMN_NAMES.map((name, i) =>
        React.createElement(View, { key: i, style: s.cell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, name))
      ),
    );

    const row3 = React.createElement(View, { key: 'r3', style: [s.row, s.cellGray] },
      React.createElement(View, { style: [s.cell, s.cellFlex2] }, React.createElement(Text, null, '')),
      React.createElement(View, { style: [s.cell, s.cellFlex4] }, React.createElement(Text, null, 'Perform each hour')),
      React.createElement(View, { style: [s.cell, s.cellFlex5] }, React.createElement(Text, null, 'Replenish as needed')),
      React.createElement(View, { style: [s.cell, s.cellFlex1] }, React.createElement(Text, null, '')),
    );

    const dataRows = ROW_HEADERS.map((header, rowIndex) => {
      const isMopRow = header.toLowerCase().includes('sweep/mop');
      const checkValue = safeStr(restroomData[`Row${rowIndex + 1}_Check`], 50);
      const checkedTime = safeStr(restroomData[`Row${rowIndex + 1}_CheckedTime`], 50);
      const taskCells = COLUMN_NAMES.slice(2, -1).map((_, colIndex) => {
        if (isMopRow) return React.createElement(View, { key: colIndex, style: s.cell }, React.createElement(Text, null, ''));
        const v = restroomData[`Row${rowIndex + 1}_Col${colIndex + 3}`];
        return React.createElement(View, { key: colIndex, style: s.cell }, React.createElement(Text, null, isChecked(v) ? 'O' : ''));
      });
      return React.createElement(View, { key: rowIndex, style: s.row },
        React.createElement(View, { style: s.cell }, React.createElement(Text, null, header)),
        React.createElement(View, { style: s.cell }, React.createElement(Text, null, checkValue)),
        ...taskCells,
        React.createElement(View, { style: s.cell }, React.createElement(Text, null, checkedTime)),
      );
    });

    const table = React.createElement(View, null, row1, row2, row3, ...dataRows);

    const header = React.createElement(View, { style: s.header },
      React.createElement(Text, { style: s.headerTitle }, 'Restroom Inspection Log'),
      React.createElement(Text, { style: s.headerSub }, `${safeSelectedOffice} - Restroom ${safeSelectedRestroom} (${safeInspectionDate})`),
    );

    const footer = React.createElement(View, { style: s.footer }, React.createElement(Text, null, `Generated: ${generatedDate}`));

    return React.createElement(Document, null,
      React.createElement(Page, { size: 'A4', orientation: 'landscape', style: s.page }, header, table, footer),
    );
  }

  const [restroomData, setRestroomData] = useState<{ [key: string]: any }>({});

  const getCurrentTime12Hour = () => {
    const now = new Date();
    const timeFormatter = new Intl.DateTimeFormat('en-US', {
      timeZone: 'America/Los_Angeles',
      hour: 'numeric',
      minute: '2-digit',
      hour12: true
    });
    const timeString = timeFormatter.format(now);
    
    const match = timeString.match(/(\d+):(\d+)\s*(AM|PM)/i);
    if (match) {
      const hours = match[1];
      const minutes = match[2];
      const ampm = match[3].toUpperCase();
      return `${hours}:${minutes} ${ampm}`;
    }
    
    const laTime = new Date(now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
    let hours = laTime.getHours();
    const minutes = laTime.getMinutes();
    const ampm = hours >= 12 ? 'PM' : 'AM';
    
    hours = hours % 12;
    hours = hours ? hours : 12;
    
    return `${hours}:${minutes.toString().padStart(2, '0')} ${ampm}`;
  };

  const autoSave = useCallback(async () => {
    if (!inspectionDate || !selectedOffice || !selectedRestroom || isUpdatingFromFirebase) return;

    const hasChanges = JSON.stringify(restroomData) !== JSON.stringify(lastSavedData);
    if (!hasChanges) return;

    try {
      const dataToSave = {
        inspectionDate,
        selectedOffice,
        selectedRestroom,
        ...restroomData,
        timestamp: new Date().toISOString(),
        autoSaved: true,
        lastUpdatedBy: userSessionId
      };

      const docId = `${inspectionDate}_${selectedOffice}_${selectedRestroom}`;
      await setDoc(doc(db, "restroom-inspections", docId), dataToSave);
      
      setLastSavedData({ ...restroomData });
      
    } catch {
    }
  }, [inspectionDate, selectedOffice, selectedRestroom, restroomData, lastSavedData, isUpdatingFromFirebase, userSessionId]);

  useEffect(() => {
    if (Object.values(restroomData).some(value => value !== '')) {
      autoSave();
    }
  }, [restroomData]);

  const handleRestroomChange = (newRestroom: string) => {
    setIsUpdatingFromFirebase(true);
    setRestroomData({});
    setLastSavedData({});
    setSelectedRestroom(newRestroom);
  };

  // 실시간 데이터 동기화
  useEffect(() => {
    if (!inspectionDate || !selectedOffice || !selectedRestroom) return;

    const docId = `${inspectionDate}_${selectedOffice}_${selectedRestroom}`;
    const docRef = doc(db, "restroom-inspections", docId);
    
    const unsubscribe = onSnapshot(docRef, (docSnap) => {
      if (docSnap.exists()) {
        const safeData = filterFirebaseData(docSnap.data());

        setIsUpdatingFromFirebase(true);

        setRestroomData((prevData) => ({
          ...prevData,
          ...safeData,
        }));

        setLastSavedData({ ...safeData });

        setTimeout(() => {
          setIsUpdatingFromFirebase(false);
        }, 100);
      } else {
        // 다른 탭에서 제출되어 문서가 삭제된 경우 폼 초기화
        setIsUpdatingFromFirebase(true);
        setRestroomData({});
        setLastSavedData({});
        setTimeout(() => {
          setIsUpdatingFromFirebase(false);
        }, 100);
      }
    }, () => {
      // Real-time listener error silently handled
    });

    return () => {
      unsubscribe();
    };
  }, [inspectionDate, selectedOffice, selectedRestroom]);

  const updateRestroomData = (field: string, value: any) => {
    setRestroomData((prev: any) => {
      const newData = { ...prev, [field]: value };

      if (field.includes('_Check')) {
        const rowNumber = field.match(/Row(\d+)_Check/)?.[1];
        if (rowNumber) {
          const timeField = `Row${rowNumber}_CheckedTime`;
          const filled = hasCheckColumnValue(value);

          if (filled) {
            setRowTaskCheckboxes(newData, rowNumber, true);
            const wasEmpty = !hasCheckColumnValue(prev[field]);
            if (wasEmpty && !newData[timeField]) {
              newData[timeField] = getCurrentTime12Hour();
            }
          } else {
            setRowTaskCheckboxes(newData, rowNumber, false);
            newData[timeField] = '';
          }
        }
      }

      return newData;
    });
  };

  const handleSubmit = async () => {
    if (!selectedOffice || !selectedRestroom) {
      alert('Please select an office and restroom first.');
      return;
    }

    const confirmed = confirm('Would you like to submit?');
    if (!confirmed) {
      return;
    }

    try {
      setLoading(true);
      setSubmitStatus('Saving...');
      setProgress(10);

      if (!inspectionDate || !selectedOffice || !selectedRestroom || !restroomData) {
        throw new Error('Please fill out all required fields.');
      }

      if (
        typeof inspectionDate !== 'string' ||
        typeof selectedOffice !== 'string' ||
        typeof selectedRestroom !== 'string' ||
        typeof restroomData !== 'object' ||
        restroomData === null ||
        Array.isArray(restroomData)
      ) {
        throw new Error('Invalid input format');
      }

      if (!/^\d{4}-\d{2}-\d{2}$/.test(inspectionDate)) {
        throw new Error('Invalid date format');
      }

      const dateObj = new Date(inspectionDate + 'T00:00:00');
      if (isNaN(dateObj.getTime())) {
        throw new Error('Invalid date format');
      }

      const [dateYear, dateMonth, dateDay] = inspectionDate.split('-').map(Number);
      if (dateYear < 1900 || dateYear > 2100 || dateMonth < 1 || dateMonth > 12 || dateDay < 1 || dateDay > 31) {
        throw new Error('Invalid date format');
      }

      const reconstructedDate = `${dateYear}-${String(dateMonth).padStart(2, '0')}-${String(dateDay).padStart(2, '0')}`;
      if (reconstructedDate !== inspectionDate) {
        throw new Error('Invalid date format');
      }

      if (!['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'].includes(selectedOffice)) {
        throw new Error('Invalid office selection');
      }

      if (!['1', '2', '3'].includes(selectedRestroom)) {
        throw new Error('Invalid restroom selection');
      }

      setSubmitStatus('Processing...');
      setProgress(60);

      const safeInspectionDate = inspectionDate.trim().slice(0, 20).replace(/[<>]/g, '');
      const safeSelectedOffice = selectedOffice.trim().slice(0, 100).replace(/[<>]/g, '');
      const safeSelectedRestroom = selectedRestroom.trim().slice(0, 50).replace(/[<>]/g, '');

      const generatedDate = new Date().toLocaleDateString('en-US', {
        year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit', hour12: true,
      });
      
      const pdfDoc = createRestroomPDFDocument({
        safeInspectionDate,
        safeSelectedOffice,
        safeSelectedRestroom,
        restroomData: restroomData as Record<string, unknown>,
        generatedDate,
      });

      const blob = await pdf(pdfDoc).toBlob();
        
      setSubmitStatus('Saving...');
      setProgress(70);
      try {
        const storage = getStorage();
        const now = new Date();
        const laTime = new Date(now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
        let hours = laTime.getHours();
        const minutes = laTime.getMinutes();
        const ampm = hours >= 12 ? 'pm' : 'am';
        hours = hours % 12;
        hours = hours ? hours : 12;
        const timeStamp = `${hours}${minutes.toString().padStart(2, '0')}${ampm}`;
        
        const filename = `6) ${inspectionDate}_${selectedOffice}_Restroom_${selectedRestroom}_Inspection_Log_${timeStamp}.pdf`;
        const storageRef = ref(storage, `endofday-pdfs/${selectedOffice}/${inspectionDate}/${filename}`);
        
        await uploadBytes(storageRef, blob);
        
      } catch (storageError: any) {
        const errorMsg = storageError?.message || 'Error';
        alert(`An error occurred while submitting. Please try again.: ${errorMsg}`);
        throw storageError;
      }
      
      setSubmitStatus('Cleaning up...');
      setProgress(80);
      const docId = `${inspectionDate}_${selectedOffice}_${selectedRestroom}`;
      await deleteDoc(doc(db, "restroom-inspections", docId));
      
      setRestroomData({});
      setLastSavedData({});

      setSubmitStatus('Complete!');
      setProgress(100);
      
      setTimeout(() => {
        setLoading(false);
        setSubmitStatus('');
        setProgress(0);
        alert('Submitted successfully!');
      }, 2000);

    } catch (error: any) {
      setSubmitStatus('❌ Submission failed: ' + error.message);
      setProgress(0);
      setTimeout(() => {
        setLoading(false);
        setSubmitStatus('');
        alert('❌ Submission failed: ' + error.message);
      }, 3000);
    }
  };

  const styles: { [key: string]: React.CSSProperties } = {
    body: {
      fontFamily: "Arial, sans-serif",
      backgroundColor: '#C2E6E6',
      color: '#333',
      lineHeight: '1.6',
      overflowX: 'hidden',
      minHeight: '100vh',
      margin: 0,
      padding: 0
    },
    container: {
      maxWidth: '95vw',
      minWidth: '320px',
      width: '100%',
      margin: '20px auto',
      backgroundColor: '#ffffff',
      borderRadius: '20px',
      boxShadow: '0 20px 40px rgba(0, 0, 0, 0.1), 0 0 0 1px rgba(255, 255, 255, 0.1)',
      padding: '25px 2vw 30px 2vw',
      overflowX: 'auto',
      backdropFilter: 'blur(10px)',
      border: '1px solid rgba(255, 255, 255, 0.2)'
    },
    header: {
      fontSize: '1.8rem',
      fontWeight: '700',
      marginBottom: '10px',
      color: '#4a6fa1',
      textAlign: 'center',
      textShadow: '0 2px 4px rgba(0, 0, 0, 0.1)',
      letterSpacing: '-0.5px'
    },
    infoRow: {
      display: 'flex',
      gap: '18px',
      alignItems: 'center',
      justifyContent: 'center',
      marginBottom: '30px',
      flexWrap: 'wrap',
      backgroundColor: 'rgba(194, 230, 230, 0.5)',
      padding: '20px',
      borderRadius: '15px',
      border: '1px solid rgba(74, 111, 161, 0.1)'
    },
    label: {
      fontWeight: '600',
      color: '#4a6fa1',
      fontSize: '16px'
    },
    input: {
      width: '160px',
      padding: '10px 15px',
      border: '2px solid rgba(74, 111, 161, 0.2)',
      borderRadius: '10px',
      backgroundColor: 'rgba(255, 255, 255, 0.9)',
      fontSize: '14px',
      fontWeight: '500',
      transition: 'all 0.3s ease'
    },
    select: {
      width: '100px',
      padding: '10px 15px',
      border: '2px solid rgba(74, 111, 161, 0.2)',
      borderRadius: '10px',
      backgroundColor: 'rgba(255, 255, 255, 0.9)',
      fontSize: '14px',
      fontWeight: '500',
      transition: 'all 0.3s ease'
    },
    tableContainer: {
      overflowX: 'auto',
      width: '100%',
      margin: '0 auto',
      borderRadius: '12px',
      backgroundColor: '#ffffff',
      boxShadow: '0 8px 32px rgba(74, 111, 161, 0.1)'
    },
    table: {
      width: '100%',
      maxWidth: '100%',
      minWidth: '1400px',
      tableLayout: 'auto',
      margin: '0 auto',
      overflowX: 'auto',
      borderRadius: '12px',
      overflow: 'hidden',
      boxShadow: '0 8px 32px rgba(74, 111, 161, 0.1)',
      borderCollapse: 'collapse'
    },
    th: {
      padding: '12px 15px',
      backgroundColor: '#2e3a4e',
      color: '#fff',
      fontWeight: '500',
      textAlign: 'center'
    },
    td: {
      padding: '12px 15px',
      borderBottom: '1px solid #e1e5ea',
      textAlign: 'center',
      verticalAlign: 'middle'
    },
    checkbox: {
      margin: 0,
      position: 'relative',
      left: 0
    },
    textInput: {
      width: '100%',
      padding: '6px 10px',
      border: '1px solid #c4cdd5',
      borderRadius: '4px',
      fontSize: '12px'
    },
    submitButton: {
      display: 'inline-block',
      padding: '12px 32px',
      backgroundColor: '#C2E6E6',
      color: '#333',
      fontSize: '1.1rem',
      fontWeight: '600',
      border: 'none',
      borderRadius: '50px',
      cursor: 'pointer',
      transition: 'all 0.3s ease',
      boxShadow: '0 8px 25px rgba(74, 111, 161, 0.3)',
      position: 'relative',
      overflow: 'hidden',
      letterSpacing: '0.5px'
    },
  };

  // 로그인한 사용자의 Office 옵션 불러오기
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      if (!currentUser) return;

      try {
        const userDoc = await getDoc(doc(db, 'users', currentUser.uid));
        if (!userDoc.exists()) return;

        const userData = userDoc.data();

        if (userData?.offices) {
          const officesArray = Array.isArray(userData.offices)
            ? userData.offices
            : [userData.offices];

          const validOptions = officesArray.filter((g: string) => officeOptions.includes(g));

          if (validOptions.length > 0) {
            setUserOfficesOptions(validOptions);
            if (validOptions.length === 1) {
              setSelectedOffice(validOptions[0]);
            }
          }
        }
      } catch {
        // 사이트 레벨 인증을 사용하므로 여기서는 조용히 처리
      }
    });

    return () => unsubscribe();
  }, []);

  useEffect(() => {
    setInspectionDate(getCurrentCaliforniaTime());
  }, []);

  useEffect(() => {
    if (process.env.NODE_ENV === 'production' &&
        typeof window !== 'undefined' &&
        window.location.protocol !== 'https:') {
      window.location.href = window.location.href.replace('http:', 'https:');
    }
  }, []);

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
    <div style={styles.body}>
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
          } as React.CSSProperties}>
            <div style={{
              backgroundColor: "white",
              padding: "40px",
              borderRadius: "20px",
              textAlign: "center",
              boxShadow: "0 20px 40px rgba(0, 0, 0, 0.3)",
              maxWidth: "400px",
              width: "90%"
            } as React.CSSProperties}>
              <div style={{
                fontSize: "18px",
                fontWeight: "600",
                color: "#4a6fa1",
                marginBottom: "20px"
              } as React.CSSProperties}>
                {submitStatus}
              </div>
              {progress > 0 && (
                <div style={{
                  width: "100%",
                  height: "8px",
                  backgroundColor: "#f0f0f0",
                  borderRadius: "4px",
                  overflow: "hidden",
                  marginBottom: "10px"
                } as React.CSSProperties}>
                  <div style={{
                    width: `${progress}%`,
                    height: "100%",
                    background: "linear-gradient(90deg, #4a90e2, #51cf66)",
                    transition: "width 0.3s ease",
                    borderRadius: "4px"
                  } as React.CSSProperties} />
                </div>
              )}
              <div style={{
                fontSize: "14px",
                color: "#666",
                marginTop: "10px"
              } as React.CSSProperties}>
                {progress}%
              </div>
            </div>
          </div>
        )}

        <div style={styles.container}>
          <h2 style={styles.header}>
            <span style={{fontSize:'2.8em',verticalAlign:'middle',textShadow:'0 2px 4px rgba(0,0,0,0.1)',margin:'0 12px'} as React.CSSProperties}>🚻</span> 
            Restroom Inspection Log 
            <span style={{fontSize:'2.8em',verticalAlign:'middle',textShadow:'0 2px 4px rgba(0,0,0,0.1)',margin:'0 12px'} as React.CSSProperties}>🚻</span>
          </h2>
          
          <div style={{textAlign:'center', fontSize:'14px', color:'#4a6fa1', marginBottom:'20px', fontStyle:'italic',fontWeight:'500',textShadow:'0 1px 2px rgba(0,0,0,0.05)'} as React.CSSProperties}>
            Only perform tasks as needed.
          </div>

          <div style={styles.infoRow}>
            <label style={styles.label} htmlFor="date">📅 Date:</label>
            <input
              type="date"
              id="date"
              value={inspectionDate}
              onChange={(e: any) => setInspectionDate(e.target.value)}
              style={styles.input}
            />
            {userOfficesOptions.length > 0 && (
              <>
                <label style={styles.label} htmlFor="office">🏢 Office:</label>
                {userOfficesOptions.length === 1 ? (
                  <span style={{
                    ...styles.select,
                    display: 'inline-flex',
                    alignItems: 'center',
                    backgroundColor: '#e9ecef',
                    fontWeight: '600',
                    color: '#4a6fa1'
                  } as React.CSSProperties}>
                    {selectedOffice}
                  </span>
                ) : (
                  <select
                    id="office"
                    value={selectedOffice}
                    onChange={(e: any) => setSelectedOffice(e.target.value)}
                    style={styles.select}
                  >
                    <option value="">--Select Office--</option>
                    {userOfficesOptions.map(office => (
                      <option key={office} value={office}>{office}</option>
                    ))}
                  </select>
                )}
              </>
            )}
            <label style={styles.label} htmlFor="restroom">🚻 Restroom:</label>
            <select
              id="restroom"
              value={selectedRestroom}
              onChange={(e: any) => handleRestroomChange(e.target.value)}
              style={styles.select}
            >
              <option value="" disabled hidden>Select Restroom</option>
              {restroomOptions.map(restroom => (
                <option key={restroom} value={restroom}>{restroom}</option>
              ))}
            </select>
          </div>

          {selectedOffice && selectedRestroom && (
            <div style={styles.tableContainer}>
              <table style={styles.table}>
                <thead>
                  <tr>
                    <th colSpan={2}></th>
                    <th colSpan={4} style={{backgroundColor:'#e3e8f0',fontWeight:'700'} as React.CSSProperties}>
                      <span style={{color:'#4a6fa1'} as React.CSSProperties}>SPOTLESS</span>
                    </th>
                    <th colSpan={5} style={{backgroundColor:'#f3f7fa',fontWeight:'700'} as React.CSSProperties}>
                      <span style={{color:'#4a6fa1'} as React.CSSProperties}>STOCKED</span>
                    </th>
                    <th></th>
                  </tr>
                  <tr>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Time</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Check</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Pick up Paper</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Wipe Sinks and Mirrors</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Wipe Toilets</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Wipe Baby Table</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Empty Trash</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Toilet Paper</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Soap</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Toilet Seat Covers</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Refresh Spray</th>
                    <th style={{...styles.th, backgroundColor:'#2e3a4e', color:'white'} as React.CSSProperties}>Checked Time</th>
                  </tr>
                  <tr>
                    <th colSpan={2}></th>
                    <th colSpan={4} style={{backgroundColor:'#f7fafd',fontSize:'13px',fontWeight:'400',color:'#4a6fa1',textAlign:'center'} as React.CSSProperties}>
                      Perform each hour
                    </th>
                    <th colSpan={5} style={{backgroundColor:'#fafdff',fontSize:'13px',fontWeight:'400',color:'#4a6fa1',textAlign:'center'} as React.CSSProperties}>
                      Replenish as needed
                    </th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {ROW_HEADERS.map((header: string, rowIndex: number) => {
                    const isMopRow = header.toLowerCase().includes('sweep/mop');
                    const isManagerInspection = header.includes('Manager Inspection');
                    const checkOptions = (OFFICE_CHECK_OPTIONS as any)[selectedOffice] || [];
                    
                    return (
                      <tr key={rowIndex}>
                        <td style={styles.td}>{header}</td>
                        <td style={styles.td}>
                          {isManagerInspection ? (
                            checkOptions.length > 0 ? (
                              <select
                                value={restroomData[`Row${rowIndex + 1}_Check`] || ''}
                                onChange={(e: any) => updateRestroomData(`Row${rowIndex + 1}_Check`, e.target.value)}
                                style={styles.textInput}
                              >
                                <option value=""></option>
                                {checkOptions.map((opt: any) => (
                                  <option key={opt} value={opt}>{opt}</option>
                                ))}
                              </select>
                            ) : (
                              <input
                                type="checkbox"
                                checked={restroomData[`Row${rowIndex + 1}_Check`] === true}
                                onChange={(e: any) => updateRestroomData(`Row${rowIndex + 1}_Check`, e.target.checked)}
                                style={styles.checkbox}
                              />
                            )
                          ) : (
                            <input
                              type="text"
                              value={restroomData[`Row${rowIndex + 1}_Check`] || ''}
                              onChange={(e: any) => updateRestroomData(`Row${rowIndex + 1}_Check`, e.target.value)}
                              style={styles.textInput}
                            />
                          )}
                        </td>
                        {COLUMN_NAMES.slice(2, -1).map((_, colIndex) => {
                          if (isMopRow) {
                            return <td key={colIndex} style={styles.td}></td>;
                          }
                          return (
                            <td key={colIndex} style={styles.td}>
                              <input
                                type="checkbox"
                                checked={restroomData[`Row${rowIndex + 1}_Col${colIndex + 3}`] === true}
                                onChange={(e: any) => updateRestroomData(`Row${rowIndex + 1}_Col${colIndex + 3}`, e.target.checked)}
                                style={styles.checkbox}
                              />
                            </td>
                          );
                        })}
                        <td style={styles.td}>
                          <span>{restroomData[`Row${rowIndex + 1}_CheckedTime`] || ''}</span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}

          {selectedOffice && selectedRestroom && (
            <div style={{textAlign:'center',marginTop:'30px'} as React.CSSProperties}>
              <button
                type="button"
                onClick={handleSubmit}
                disabled={loading || !selectedOffice}
                style={{
                  ...styles.submitButton,
                  backgroundColor: loading || !selectedOffice ? '#bdc3c7' : '#C2E6E6',
                  cursor: loading || !selectedOffice ? 'not-allowed' : 'pointer'
                } as React.CSSProperties}
              >
                🚀 Submit
              </button>
            </div>
          )}

        </div>
      </div>
  );
}
