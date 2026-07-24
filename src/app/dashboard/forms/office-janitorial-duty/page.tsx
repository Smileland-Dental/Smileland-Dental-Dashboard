'use client';

import React, { useState, useEffect, useCallback } from 'react';
import { db, auth } from '@/lib/firebase.config';
import { doc, setDoc, getDoc, onSnapshot, deleteDoc } from 'firebase/firestore';
import { getStorage, ref, uploadBytes } from 'firebase/storage';
import { onAuthStateChanged } from 'firebase/auth';
import { pdf, Document, Page, View, Text, StyleSheet } from '@react-pdf/renderer';

const commonDuties = [
  'Pick up trash outside building and in parking lot',
  'Pressure wash outside walkways/entrance areas',
  'Switch out mop heads (wash at Corp)',
  'Clean out shredder/Oil',
  'Office Supply Delivery',
  'Toy Delivery',
  'Dust Light Fixtures',
  'Dust/Clean Vents',
  'Deep Clean Bathrooms (Floors/Walls/Lights etc.)',
  'Windows and Window Sills',
  'Clean Wet Floor Signs and Mop Bucket',
  'Sweep',
  'Mop',
  'Pour Water Down Compressor Drain',
  'Wipe/Clean Baseboards and Trim',
  'Cardboard Disposal',
  'Wipe Down Trash Cans'
];

const customDutyCount = 3;
const weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];

function getAllDuties(dutyData: Record<string, unknown>): string[] {
  const custom = Array.from({ length: customDutyCount }, (_, i) => {
    const value = dutyData[`Duty${commonDuties.length + i + 1}_Custom`];
    if (value == null) return '';
    return String(value).trim().slice(0, 200).replace(/[<>]/g, '');
  });
  return [...commonDuties, ...custom];
}

const ui = {
  page: {
    minHeight: '100vh',
    margin: 0,
    padding: '20px',
    overflowX: 'hidden' as const,
    background: 'linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%)',
    fontFamily: 'Arial, sans-serif',
    color: '#1a202c',
    lineHeight: 1.6,
  },
  container: {
    width: '80%',
    maxWidth: '1600px',
    minWidth: '900px',
    margin: '20px auto',
    background: '#ffffff',
    borderRadius: '24px',
    boxShadow: '0 20px 60px rgba(0, 0, 0, 0.1)',
    padding: '50px',
    border: '1px solid rgba(0, 0, 0, 0.05)',
    position: 'relative' as const,
  },
  title: {
    fontSize: '2.5rem',
    fontWeight: 700,
    marginBottom: '40px',
    color: '#1a202c',
    textAlign: 'center' as const,
    letterSpacing: '-0.5px',
  },
  dateRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    marginBottom: '40px',
    background: 'linear-gradient(135deg, #90cdf4 0%, #63b3ed 100%)',
    padding: '25px',
    borderRadius: '16px',
    gap: '25px',
    boxShadow: '0 10px 30px rgba(144, 205, 244, 0.2)',
  },
  dateLabel: {
    fontWeight: 600,
    color: '#ffffff',
    fontSize: '1.1rem',
  },
  dateInput: {
    padding: '14px 24px',
    border: '2px solid rgba(255, 255, 255, 0.3)',
    borderRadius: '12px',
    fontSize: '1rem',
    fontWeight: 500,
    background: 'rgba(255, 255, 255, 0.95)',
    color: '#1a202c',
    cursor: 'pointer',
    minWidth: '150px',
  },
  daySelect: {
    padding: '12px 20px',
    border: '2px solid #dee2e6',
    borderRadius: '8px',
    fontSize: '1rem',
    fontWeight: 500,
    background: '#f8f9fa',
    color: '#6c757d',
    cursor: 'not-allowed',
    minWidth: '150px',
  },
  table: {
    width: '100%',
    borderCollapse: 'separate' as const,
    borderSpacing: 0,
    marginBottom: '30px',
    borderRadius: '16px',
    overflow: 'hidden',
    boxShadow: '0 4px 20px rgba(0, 0, 0, 0.08)',
    background: '#ffffff',
  },
  theadRow: {
    background: 'linear-gradient(135deg, #90cdf4 0%, #63b3ed 100%)',
  },
  th: {
    padding: '18px 24px',
    textAlign: 'left' as const,
    borderBottom: '1px solid #e2e8f0',
    color: '#1a202c',
    fontWeight: 600,
    fontSize: '0.95rem',
    textTransform: 'uppercase' as const,
    letterSpacing: '0.8px',
  },
  td: {
    padding: '12px',
    borderBottom: '1px solid #e2e8f0',
  },
  tdCenter: {
    padding: '12px',
    textAlign: 'center' as const,
    borderBottom: '1px solid #e2e8f0',
  },
  timeInput: {
    width: '100%',
    padding: '6px',
    border: '1px solid #c4cdd5',
    borderRadius: '4px',
  },
  dutyInput: {
    width: '100%',
    padding: '6px',
    border: '1px solid #c4cdd5',
    borderRadius: '4px',
    fontSize: '0.95rem',
    fontFamily: 'Arial, sans-serif',
  },
  byGroup: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '4px',
  },
  byOption: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    fontSize: '0.9rem',
    cursor: 'pointer',
    whiteSpace: 'nowrap' as const,
  },
  submitBtn: {
    display: 'block',
    width: '100%',
    maxWidth: '320px',
    margin: '40px auto',
    padding: '18px 40px',
    fontSize: '1.1rem',
    fontWeight: 600,
    background: 'linear-gradient(135deg, #90cdf4 0%, #63b3ed 100%)',
    color: '#fff',
    border: 'none',
    borderRadius: '14px',
    cursor: 'pointer',
    boxShadow: '0 10px 30px rgba(144, 205, 244, 0.3)',
  },
  submitBtnDisabled: {
    opacity: 0.6,
    cursor: 'not-allowed',
  },
  statusToast: {
    position: 'fixed' as const,
    top: '20px',
    right: '20px',
    padding: '12px 20px',
    color: 'white',
    borderRadius: '25px',
    fontSize: '14px',
    fontWeight: 'bold' as const,
    boxShadow: '0 4px 12px rgba(0,0,0,0.3)',
    zIndex: 1000,
    maxWidth: '300px',
    textAlign: 'center' as const,
  },
  authScreen: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    height: '100vh',
    background: 'linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%)',
    fontFamily: 'Arial, sans-serif',
  },
  authInner: {
    textAlign: 'center' as const,
  },
  authError: {
    fontSize: '18px',
    color: '#d32f2f',
    marginBottom: '10px',
  },
  overlay: {
    position: 'fixed' as const,
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    backgroundColor: 'rgba(0, 0, 0, 0.7)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    zIndex: 9999,
  },
  overlayCard: {
    backgroundColor: 'white',
    padding: '40px',
    borderRadius: '20px',
    textAlign: 'center' as const,
    boxShadow: '0 20px 40px rgba(0, 0, 0, 0.3)',
    maxWidth: '400px',
    width: '90%',
  },
  overlayStatus: {
    fontSize: '18px',
    fontWeight: 600,
    color: '#4a6fa1',
    marginBottom: '20px',
  },
  progressTrack: {
    width: '100%',
    height: '8px',
    backgroundColor: '#f0f0f0',
    borderRadius: '4px',
    overflow: 'hidden',
    marginBottom: '10px',
  },
  progressBar: {
    height: '100%',
    background: 'linear-gradient(90deg, #4a90e2, #51cf66)',
    borderRadius: '4px',
  },
  progressLabel: {
    fontSize: '14px',
    color: '#666',
    marginTop: '10px',
  },
};

function getTimeLabel(index: number): string {
  return `${index})`;
}

export default function JanitorialDutyPage() {
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

  const getInitialDateAndDay = () => {
    const now = new Date();
    const laTimeString = now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' });
    const laDate = new Date(laTimeString);
    const dayOfWeek = laDate.getDay();
    const pageForToday = dayOfWeek === 0 || dayOfWeek === 6 ? 1 : dayOfWeek;
    return {
      date: getCurrentCaliforniaTime(),
      day: pageForToday
    };
  };

  const initialValues = getInitialDateAndDay();

  const getDayFromDate = (dateString: string): number => {
    const date = new Date(dateString + 'T00:00:00');
    const dayOfWeek = date.getDay();
    return dayOfWeek === 0 || dayOfWeek === 6 ? 1 : dayOfWeek;
  };

  const [selectedDate, setSelectedDate] = useState(initialValues.date);
  const [selectedDay, setSelectedDay] = useState(initialValues.day);
  const [offices, setOffices] = useState<string[]>([]);
  const [selectedOffice, setSelectedOffice] = useState('');
  const [dutyData, setDutyData] = useState<any>({});
  const [loading, setLoading] = useState(false);
  const [isInitialLoad, setIsInitialLoad] = useState(true);
  const [autoSaveStatus, setAutoSaveStatus] = useState('');
  const [submitStatus, setSubmitStatus] = useState('');
  const [progress, setProgress] = useState(0);
  const [isUpdatingFromFirebase, setIsUpdatingFromFirebase] = useState(false);
  const [userSessionId] = useState(() => Math.random().toString(36).substr(2, 9));
  const [lastSavedData, setLastSavedData] = useState<any>({});
  const [isLoggedIn, setIsLoggedIn] = useState<boolean | null>(null);

  const sanitizeOffice = (office: string): string => {
    return String(office).trim().replace(/[\/\\#?[\]\s]+/g, '_').slice(0, 100);
  };

  const getDocId = (date: string, day: number, office: string): string => {
    const dayName = weekdays[day - 1];
    return `${date}_${dayName}_${sanitizeOffice(office)}_janitorial`;
  };

  const pdfStyles = StyleSheet.create({
    page: { padding: 22, fontFamily: 'Helvetica', fontSize: 8 },
    header: { marginBottom: 10, borderBottomWidth: 2, borderColor: '#333', paddingBottom: 6, alignItems: 'center' },
    headerTitle: { fontSize: 14, fontWeight: 'bold', marginBottom: 4 },
    headerSub: { fontSize: 10 },
    row: { flexDirection: 'row', borderBottomWidth: 0.5, borderColor: '#333' },
    cell: { padding: 3, fontSize: 6, flex: 1, borderRightWidth: 0.5, borderColor: '#333', justifyContent: 'center', alignItems: 'center' },
    cellDuty: { flex: 5 },
    cellTime: { flex: 1.5 },
    cellBy: { flex: 1.5 },
    cellCheckedBy: { flex: 1.5 },
    cellBold: { fontWeight: 'bold' },
    cellGray: { backgroundColor: '#f0f0f0' },
    footer: { marginTop: 15, paddingTop: 8, borderTopWidth: 1, borderColor: '#ddd', alignItems: 'center', fontSize: 7, color: '#666' },
  });

  function safeStr(v: unknown, max: number): string {
    if (v == null) return '';
    return String(v).trim().slice(0, max).replace(/[<>]/g, '');
  }

  function createJanitorialPDFDocument(props: {
    dayName: string;
    formattedDate: string;
    duties: string[];
    dutyData: Record<string, unknown>;
    selectedDay: number;
    generatedDate: string;
    office: string;
  }) {
    const { dayName, formattedDate, duties, dutyData, generatedDate, office } = props;
    const s = pdfStyles;

    // Header row
    const headerRow = React.createElement(View, { key: 'header', style: [s.row, s.cellGray] },
      React.createElement(View, { style: s.cell }, React.createElement(Text, { style: s.cellBold }, '#')),
      React.createElement(View, { style: [s.cell, s.cellDuty] }, React.createElement(Text, { style: s.cellBold }, 'Daily Duties')),
      React.createElement(View, { style: [s.cell, s.cellBy] }, React.createElement(Text, { style: s.cellBold }, 'Completed By')),
      React.createElement(View, { style: [s.cell, s.cellTime] }, React.createElement(Text, { style: s.cellBold }, 'Time')),
      React.createElement(View, { style: [s.cell, s.cellCheckedBy] }, React.createElement(Text, { style: s.cellBold }, 'Checked By')),
    );

    const dataRows = duties.map((duty, idx) => {
      const index = idx + 1;
      const timeLabel = getTimeLabel(index);
      const timeValue = safeStr(dutyData[`Duty${index}_Time`], 50);
      const byValue = dutyData[`Duty${index}_By`];
      const byStr = Array.isArray(byValue) ? byValue.join(', ') : safeStr(byValue, 50);
      const checkedByValue = safeStr(dutyData[`Duty${index}_CheckedBy`], 50);

      return React.createElement(View, { key: index, style: s.row },
        React.createElement(View, { style: s.cell }, React.createElement(Text, null, timeLabel)),
        React.createElement(View, { style: [s.cell, s.cellDuty] }, React.createElement(Text, null, duty)),
        React.createElement(View, { style: [s.cell, s.cellBy] }, React.createElement(Text, null, byStr)),
        React.createElement(View, { style: [s.cell, s.cellTime] }, React.createElement(Text, null, timeValue)),
        React.createElement(View, { style: [s.cell, s.cellCheckedBy] }, React.createElement(Text, null, checkedByValue)),
      );
    });

    const table = React.createElement(View, null, headerRow, ...dataRows);

    const header = React.createElement(View, { style: s.header },
      React.createElement(Text, { style: s.headerTitle }, `${office} Janitorial Duty`),
      React.createElement(Text, { style: s.headerSub }, `${dayName} (${formattedDate})`),
    );

    const footer = React.createElement(View, { style: s.footer }, React.createElement(Text, null, `Generated: ${generatedDate}`));

    return React.createElement(Document, null,
      React.createElement(Page, { size: 'A4', orientation: 'landscape', style: s.page }, header, table, footer),
    );
  }

  const getCurrentTime12Hour = () => {
    const now = new Date();
    const laTime = new Date(now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
    let hours = laTime.getHours();
    const minutes = laTime.getMinutes();
    const ampm = hours >= 12 ? 'pm' : 'am';
    hours = hours % 12;
    hours = hours ? hours : 12;
    
    return `${hours}:${minutes.toString().padStart(2, '0')}${ampm}`;
  };

  const autoSave = useCallback(async () => {
    if (!selectedDate || !selectedDay || !selectedOffice || isUpdatingFromFirebase) return;

    const hasChanges = JSON.stringify(dutyData) !== JSON.stringify(lastSavedData);
    if (!hasChanges) return;

    try {
      const dataToSave = {
        selectedDate,
        selectedDay,
        selectedOffice,
        ...dutyData,
        timestamp: new Date().toISOString(),
        autoSaved: true,
        lastUpdatedBy: userSessionId
      };

      const docId = getDocId(selectedDate, selectedDay, selectedOffice);
      await setDoc(doc(db, "janitorial-duties", docId), dataToSave);
      
      setLastSavedData({ ...dutyData });

    } catch (error) {
    }
  }, [selectedDate, selectedDay, selectedOffice, dutyData, lastSavedData, isUpdatingFromFirebase, userSessionId]);

  // 데이터 변경 시에만 자동 저장
  useEffect(() => {
    if (Object.values(dutyData).some(value => value !== '' && value !== null && value !== undefined)) {
      autoSave();
    }
  }, [dutyData]);

  // 데이터 로드 함수
  const loadData = async () => {
    if (!selectedDate || !selectedDay || !selectedOffice) return Promise.resolve();

    try {
      setSubmitStatus('Loading data...');
      
      const docId = getDocId(selectedDate, selectedDay, selectedOffice);
      const docSnap = await getDoc(doc(db, "janitorial-duties", docId));
      
      if (docSnap.exists()) {
        const data = docSnap.data();
        
        setIsUpdatingFromFirebase(true);
        
        setDutyData((prevData: any) => ({
          ...prevData,
          ...data
        }));
        
        setLastSavedData({ ...data });
        
        setTimeout(() => {
          setIsUpdatingFromFirebase(false);
        }, 100);
        
        setSubmitStatus('Data loaded successfully');
        setTimeout(() => setSubmitStatus(''), 2000);
      } else {
        setDutyData({});
        setLastSavedData({});
        setSubmitStatus('No data found - initialized empty form');
        setTimeout(() => setSubmitStatus(''), 2000);
      }
      
      return Promise.resolve();
    } catch (error: any) {
      setSubmitStatus('Error loading data: ' + error.message);
      setTimeout(() => setSubmitStatus(''), 3000);
      return Promise.resolve();
    }
  };

  // 날짜, 요일, 오피스 변경 시 데이터 로드
  useEffect(() => {
    if (selectedDate && selectedDay && selectedOffice) {
      loadData().then(() => {
        setIsInitialLoad(false);
      });
    }
  }, [selectedDate, selectedDay, selectedOffice]);

  useEffect(() => {
    if (!selectedDate || !selectedDay || !selectedOffice) return;

    const docId = getDocId(selectedDate, selectedDay, selectedOffice);
    const docRef = doc(db, "janitorial-duties", docId);
    
    setDutyData({});
    setLastSavedData({});
    
    const unsubscribe = onSnapshot(docRef, (docSnap: any) => {
      if (docSnap.exists()) {
        const data = docSnap.data();
        
        const { selectedDate: _, selectedDay: __, selectedOffice: ___, timestamp, autoSaved, lastUpdatedBy, ...dutyFields } = data;
        
        const hasChanges = JSON.stringify(dutyFields) !== JSON.stringify(dutyData);
        if (!hasChanges) return;
        
        setIsUpdatingFromFirebase(true);
        
        setDutyData(dutyFields);
        setLastSavedData(dutyFields);
        
        setTimeout(() => {
          setIsUpdatingFromFirebase(false);
        }, 100);
        
        if (data.timestamp && 
            new Date(data.timestamp).getTime() > Date.now() - 5000 && 
            data.lastUpdatedBy && 
            data.lastUpdatedBy !== userSessionId) {
          setAutoSaveStatus('🔄 Updated by another user.');
          setTimeout(() => setAutoSaveStatus(''), 3000);
        }
      } else {
        setDutyData({});
        setLastSavedData({});
      }
    }, (error: any) => {
      setAutoSaveStatus('❌ 연결 오류 - 실시간 동기화가 중단되었습니다');
      setTimeout(() => setAutoSaveStatus(''), 3000);
    });

    return () => {
      unsubscribe();
    };
  }, [selectedDate, selectedDay, selectedOffice, userSessionId]);

  const updateDutyData = (field: string, value: any) => {
    setDutyData((prev: any) => {
      const newData = { ...prev, [field]: value };
      
      if (field.includes('_By') && value && value.length > 0) {
        const rowNumber = field.match(/Duty(\d+)_By/)?.[1];
        if (rowNumber) {
          const timeField = `Duty${rowNumber}_Time`;
          if (!newData[timeField]) {
            newData[timeField] = getCurrentTime12Hour();
          }
        }
      }
      
      return newData;
    });
  };

  const handleSubmit = async () => {
    if (!selectedDate || !selectedDay || !selectedOffice) {
      alert('Please select a date, day, and office first.');
      return;
    }

    const confirmed = confirm('Would you like to submit?');
    if (!confirmed) return;

    setLoading(true);
    setSubmitStatus('Saving...');
    setProgress(10);

    try {
      setSubmitStatus('Processing...');
      setProgress(30);
      
      const dayName = weekdays[selectedDay - 1];
      const duties = getAllDuties(dutyData as Record<string, unknown>);

      const dateObj = new Date(selectedDate + 'T00:00:00');
      const laDate = new Date(dateObj.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
      const month = String(laDate.getMonth() + 1).padStart(2, '0');
      const day = String(laDate.getDate()).padStart(2, '0');
      const year = laDate.getFullYear();
      const formattedDate = `${month}/${day}/${year}`;

      const generatedDate = new Date().toLocaleDateString('en-US', {
        year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit', hour12: true,
      });

      setSubmitStatus('Processing...');
      setProgress(60);

      const pdfDoc = createJanitorialPDFDocument({
        dayName,
        formattedDate,
        duties,
        dutyData: dutyData as Record<string, unknown>,
        selectedDay,
        generatedDate,
        office: selectedOffice,
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
        
        const filename = `${selectedDate}_${dayName}_${selectedOffice}_Janitorial Duty_${timeStamp}.pdf`;
        const storageRef = ref(storage, `endofday-pdfs/${selectedOffice}/${selectedDate}/${filename}`);
        
        await uploadBytes(storageRef, blob);
        
      } catch (storageError: any) {
        const errorMsg = storageError?.message || 'Error';
        alert(`An error occurred while submitting. Please try again.: ${errorMsg}`);
        throw storageError;
      }
      
      setSubmitStatus('Cleaning up...');
      setProgress(80);
      const docId = getDocId(selectedDate, selectedDay, selectedOffice);
      await deleteDoc(doc(db, "janitorial-duties", docId));
      
      setDutyData({});
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

  const renderDutyRows = () => {
    const totalRows = commonDuties.length + customDutyCount;

    return Array.from({ length: totalRows }, (_, idx) => {
      const index = idx + 1;
      const isCustom = index > commonDuties.length;
      const duty = isCustom ? '' : commonDuties[idx];
      const customKey = `Duty${index}_Custom`;
      const timeLabel = getTimeLabel(index);
      const byKey = `Duty${index}_By`;
      const timeKey = `Duty${index}_Time`;
      const checkedByKey = `Duty${index}_CheckedBy`;
      const isEven = idx % 2 === 1;
      
      return (
        <tr 
          key={index}
          style={{
            background: isEven ? '#f7fafc' : '#ffffff',
          }}
        >
          <td style={ui.tdCenter}>{timeLabel}</td>
          <td style={ui.td}>
            {isCustom ? (
              <input
                type="text"
                name={customKey}
                value={dutyData[customKey] || ''}
                maxLength={200}
                onChange={(e) => updateDutyData(customKey, e.target.value)}
                style={ui.dutyInput}
              />
            ) : (
              duty
            )}
          </td>
          <td style={ui.td}>
            <div style={ui.byGroup}>
              {['A', 'M', 'Other'].map((name) => {
                const selected = Array.isArray(dutyData[byKey]) ? dutyData[byKey] : [];
                const checked = selected.includes(name);
                return (
                  <label key={name} style={ui.byOption}>
                    <input
                      type="checkbox"
                      checked={checked}
                      onChange={(e) => {
                        const next = e.target.checked
                          ? [...selected, name]
                          : selected.filter((v: string) => v !== name);
                        updateDutyData(byKey, next);
                      }}
                    />
                    {name}
                  </label>
                );
              })}
            </div>
          </td>
          <td style={ui.td}>
            <input
              type="text"
              name={timeKey}
              value={dutyData[timeKey] || ''}
              readOnly
              style={ui.timeInput}
            />
          </td>
          <td style={ui.td}>
            <input
              type="text"
              name={checkedByKey}
              value={dutyData[checkedByKey] || ''}
              maxLength={50}
              onChange={(e) => updateDutyData(checkedByKey, e.target.value)}
              style={ui.timeInput}
            />
          </td>
        </tr>
      );
    });
  };

  useEffect(() => {
    const timer = setTimeout(() => {
      setIsInitialLoad(false);
    }, 200);
    return () => clearTimeout(timer);
  }, []);

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      if (!currentUser) {
        alert('Please log in.');
        setIsLoggedIn(false);
        setOffices([]);
        setSelectedOffice('');
        return;
      }

      try {
        const userDoc = await getDoc(doc(db, 'users', currentUser.uid));
        const rawOffices = userDoc.exists() ? userDoc.data()?.offices : [];
        const officeList = Array.isArray(rawOffices)
          ? rawOffices
              .map((item: unknown) => String(item ?? '').trim())
              .filter((item: string) => item.length > 0)
          : [];

        setOffices(officeList);
        if (officeList.length === 1) {
          setSelectedOffice(officeList[0]);
        } else {
          setSelectedOffice('');
        }
        setIsLoggedIn(true);
      } catch (error) {
        setOffices([]);
        setSelectedOffice('');
        setIsLoggedIn(true);
      }
    });

    if (process.env.NODE_ENV === 'production' && 
        typeof window !== 'undefined' && 
        window.location.protocol !== 'https:') {
      window.location.href = window.location.href.replace('http:', 'https:');
    }

    return () => {
      unsubscribe();
    };
  }, []);

  const safeProgress = Math.min(100, Math.max(0, Number(progress) || 0));

  if (isLoggedIn === null) {
    return null;
  }

  if (isLoggedIn === false) {
    return (
      <div style={ui.authScreen}>
        <div style={ui.authInner}>
          <div style={ui.authError}>Please log in to access this page.</div>
        </div>
      </div>
    );
  }

  return (
    <div style={ui.page}>
      <div style={{
        ...ui.container,
        visibility: isInitialLoad ? 'hidden' : 'visible',
      }}>
        {autoSaveStatus && (
          <div style={{
            ...ui.statusToast,
            backgroundColor: autoSaveStatus.includes('❌') ? '#ef4444' : 
                            autoSaveStatus.includes('🔄') ? '#4a90e2' : '#51cf66',
          }}>
            {autoSaveStatus}
          </div>
        )}

        <h2 style={ui.title}>Janitorial Duty</h2>
        
        <div style={ui.dateRow}>
          <label htmlFor="date" style={ui.dateLabel}>📅 Date:</label>
          <input
            type="date"
            id="date"
            value={selectedDate}
            onChange={(e) => {
              const newDate = e.target.value;
              setSelectedDate(newDate);
              const newDay = getDayFromDate(newDate);
              setSelectedDay(newDay);
            }}
            style={ui.dateInput}
          />
          <label htmlFor="day" style={ui.dateLabel}>📆 Day:</label>
          <select
            id="day"
            value={selectedDay}
            disabled
            style={ui.daySelect}
          >
            {weekdays.map((day, idx) => (
              <option key={idx + 1} value={idx + 1}>{day}</option>
            ))}
          </select>
          <label htmlFor="office" style={ui.dateLabel}>🏢 Office:</label>
          <select
            id="office"
            value={selectedOffice}
            disabled={offices.length === 0}
            onChange={(e) => setSelectedOffice(e.target.value)}
            style={offices.length === 0 ? ui.daySelect : ui.dateInput}
          >
            <option value="">
              {offices.length === 0 ? 'No office available' : 'Select office'}
            </option>
            {offices.map((office) => (
              <option key={office} value={office}>{office}</option>
            ))}
          </select>
        </div>

        {selectedDate && selectedDay && selectedOffice && (
          <>
            <table style={ui.table}>
              <thead>
                <tr style={ui.theadRow}>
                  <th style={ui.th}>#</th>
                  <th style={ui.th}>Daily Duties</th>
                  <th style={ui.th}>Completed By</th>
                  <th style={ui.th}>Time</th>
                  <th style={ui.th}>Checked By</th>
                </tr>
              </thead>
              <tbody>
                {renderDutyRows()}
              </tbody>
            </table>

            <button
              onClick={handleSubmit}
              disabled={loading}
              style={{
                ...ui.submitBtn,
                ...(loading ? ui.submitBtnDisabled : {}),
              }}
            >
              {loading ? 'Submitting...' : 'Submit'}
            </button>
          </>
        )}

        {loading && (
          <div style={ui.overlay}>
            <div style={ui.overlayCard}>
              <div style={ui.overlayStatus}>
                {submitStatus}
              </div>
              {safeProgress > 0 && (
                <div style={ui.progressTrack}>
                  <div style={{
                    ...ui.progressBar,
                    width: `${safeProgress}%`,
                  }} />
                </div>
              )}
              <div style={ui.progressLabel}>
                {safeProgress}%
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
