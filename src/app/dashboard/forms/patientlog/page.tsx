'use client'

import React, { useState, useEffect, useCallback, useRef } from "react";
import { doc, setDoc, collection, getDocs, deleteDoc, getDoc, query, where, arrayUnion } from "firebase/firestore";
import { db, auth } from "@/lib/firebase.config";
import { getAuth, onAuthStateChanged } from 'firebase/auth';
import { getStorage, ref, uploadBytes } from 'firebase/storage';
import { pdf, Document, Page, View, Text, StyleSheet } from '@react-pdf/renderer';

interface PatientRowProps {
  row: any;
  updatePatientRow: (id: number, field: string, value: any) => void;
  removePatientRow: (id: number) => void;
  patientOfficeOptions: string[];
  getVisitTypeOptions: (office: string) => string[];
  remarkOptions: string[];
  sourceOptions: string[];
  otherDutyOptions: string[];
  inputStyle: React.CSSProperties;
  buttonStyle: React.CSSProperties;
}

const PatientRow = React.memo(({ 
  row, 
  updatePatientRow, 
  removePatientRow, 
  patientOfficeOptions, 
  getVisitTypeOptions, 
  remarkOptions, 
  sourceOptions,
  otherDutyOptions,
  inputStyle,
  buttonStyle
}: PatientRowProps) => {
  const visitTypeOptions = getVisitTypeOptions(row.office);
  
  return (
    <tr style={{ backgroundColor: row.id % 2 === 0 ? '#f9f9f9' : 'white' }}>
      <td style={{ padding: '8px', textAlign: 'center', fontWeight: 'bold' }}>
        {row.id}
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
          {visitTypeOptions.map(type => (
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
        {row.isOther ? (
          <input
            type="text"
            autoFocus
            value={row.source}
            onChange={(e) =>
              updatePatientRow(row.id, 'source', e.target.value)
            }
            placeholder="Enter source"
            style={{
              ...inputStyle,
              margin: 0,
              fontSize: '14px',
            }}
          />
        ) : (
          <select
            value={row.source}
            onChange={(e) => {
              const value = e.target.value;

              if (value === '__OTHER__') {
                updatePatientRow(row.id, 'source', '');
                updatePatientRow(row.id, 'isOther', true);
              } else {
                updatePatientRow(row.id, 'source', value);
                updatePatientRow(row.id, 'isOther', false);
              }
            }}
            style={{
              ...inputStyle,
              margin: 0,
              fontSize: '14px',
            }}
          >
            <option value=""></option>

            {sourceOptions.map(source => (
              <option key={source} value={source}>
                {source}
              </option>
            ))}

            <option value="__OTHER__">Other</option>
          </select>
        )}
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
      </td>
    </tr>
  );
}, (prevProps, nextProps) => {
  return (
    prevProps.row.id === nextProps.row.id &&
    prevProps.row.name === nextProps.row.name &&
    prevProps.row.office === nextProps.row.office &&
    prevProps.row.appt_date === nextProps.row.appt_date &&
    prevProps.row.visit_type === nextProps.row.visit_type &&
    prevProps.row.call_in === nextProps.row.call_in &&
    prevProps.row.call_out === nextProps.row.call_out &&
    prevProps.row.time === nextProps.row.time &&
    prevProps.row.remark === nextProps.row.remark &&
    prevProps.row.source === nextProps.row.source &&
    prevProps.row.isOther === nextProps.row.isOther &&
    prevProps.row.other_duty === nextProps.row.other_duty
  );
});

function sanitizeFirebaseDataClient(data: any, depth: number = 0): any {
  if (depth > 20) return null;
  
  if (data === null || data === undefined) return null;
  
  if (typeof data !== 'object') {
    if (typeof data === 'string') {
      if (data.length > 900 * 1024) {
        return data.slice(0, 900 * 1024);
      }
      return data.replace(/[\x00-\x1F\x7F-\x9F]/g, '').trim();
    }
    if (typeof data === 'number' && (isNaN(data) || !isFinite(data))) {
      return 0;
    }
    return data;
  }
  
  if (Array.isArray(data)) {
    if (data.length > 10000) {
      return data.slice(0, 10000).map(item => sanitizeFirebaseDataClient(item, depth + 1));
    }
    return data.map(item => sanitizeFirebaseDataClient(item, depth + 1));
  }
  
  const sanitized: any = {};
  let keyCount = 0;
  const maxKeys = 1000;
  
  for (const key in data) {
    if (Object.prototype.hasOwnProperty.call(data, key)) {
      if (keyCount >= maxKeys) break;
      
      if (key.length > 1500 || key.length === 0) continue;
      
      const safeKey = key.replace(/[.$[\]#\/]/g, '_').slice(0, 1500);
      
      sanitized[safeKey] = sanitizeFirebaseDataClient(data[key], depth + 1);
      keyCount++;
    }
  }
  
  return sanitized;
}

function sanitizeFirebaseDocIdClient(docId: string): string {
  return docId
    .replace(/[\/\s]/g, '_') 
    .replace(/[^a-zA-Z0-9_-]/g, '') 
    .slice(0, 1500); 
}

function safeStr(v: unknown, max: number): string {
  if (v == null) return '';
  return String(v).trim().slice(0, max).replace(/[<>]/g, '');
}

function convertTo12Hour(timeStr: string): string {
  if (!timeStr || timeStr === '-') return '-';
  try {
    const [hours, minutes] = timeStr.split(':');
    const hour = parseInt(hours);
    const min = minutes || '00';
    if (hour === 0) return `12:${min} AM`;
    if (hour < 12) return `${hour}:${min} AM`;
    if (hour === 12) return `12:${min} PM`;
    return `${hour - 12}:${min} PM`;
  } catch {
    return timeStr;
  }
}

function createPatientLogPDFDocument(props: {
  safeDutyDate: string;
  safeUserName: string;
  safeWorkOffice: string;
  safeWorkHoursFrom: string;
  safeWorkHoursTo: string;
  safeDailyWorkReport: string;
  patientList: any[];
  totalAppointments: number;
  incomingCalls: number;
  outgoingCalls: number;
  generatedDate: string;
}) {
  const { 
    safeDutyDate, 
    safeUserName, 
    safeWorkOffice, 
    safeWorkHoursFrom, 
    safeWorkHoursTo,
    safeDailyWorkReport,
    patientList,
    totalAppointments,
    incomingCalls,
    outgoingCalls,
    generatedDate 
  } = props;

  const pdfStyles = StyleSheet.create({
    page: { padding: 20, fontFamily: 'Helvetica', fontSize: 9 },
    header: { marginBottom: 15, borderBottomWidth: 2, borderColor: '#333', paddingBottom: 8, alignItems: 'center' },
    headerTitle: { fontSize: 18, fontWeight: 'bold', marginBottom: 4 },
    infoSection: { marginBottom: 15, padding: 8, borderWidth: 1, borderColor: '#ccc', flexDirection: 'row', justifyContent: 'space-between', flexWrap: 'wrap' },
    infoItem: { fontSize: 8, marginBottom: 4 },
    stats: { marginBottom: 15, padding: 8, borderWidth: 1, borderColor: '#ccc', backgroundColor: '#f9f9f9', flexDirection: 'row', justifyContent: 'center', gap: 30 },
    statItem: { alignItems: 'center' },
    statValue: { fontSize: 14, fontWeight: 'bold', marginBottom: 2 },
    statLabel: { fontSize: 9, fontWeight: 'bold' },
    table: { marginTop: 10 },
    tableRow: { flexDirection: 'row', borderBottomWidth: 0.5, borderColor: '#333' },
    tableHeader: { flexDirection: 'row', borderBottomWidth: 1, borderColor: '#333', backgroundColor: '#f0f0f0', fontWeight: 'bold' },
    tableCell: { padding: 4, fontSize: 7, flex: 1, borderRightWidth: 0.5, borderColor: '#333', justifyContent: 'center', alignItems: 'center' },
    dailyReport: { marginTop: 15, padding: 8, borderWidth: 1, borderColor: '#ccc' },
    dailyReportTitle: { fontSize: 12, fontWeight: 'bold', marginBottom: 4 },
    dailyReportContent: { fontSize: 8, lineHeight: 1.4 },
    footer: { marginTop: 15, paddingTop: 8, borderTopWidth: 1, borderColor: '#ddd', alignItems: 'center', fontSize: 8, color: '#666' },
  });

  const s = pdfStyles;

  const header = React.createElement(View, { style: s.header },
    React.createElement(Text, { style: s.headerTitle }, 'Patient Log'),
  );

  const infoSection = React.createElement(View, { style: s.infoSection },
    React.createElement(Text, { style: s.infoItem }, `Duty Date: ${safeDutyDate || '-'}`),
    React.createElement(Text, { style: s.infoItem }, `Name: ${safeUserName || '-'}`),
    React.createElement(Text, { style: s.infoItem }, `Work Office: ${safeWorkOffice || '-'}`),
    React.createElement(Text, { style: s.infoItem }, `Work Hours: ${convertTo12Hour(safeWorkHoursFrom)} - ${convertTo12Hour(safeWorkHoursTo)}`),
  );

  const stats = React.createElement(View, { style: s.stats },
    React.createElement(View, { style: s.statItem },
      React.createElement(Text, { style: s.statValue }, String(totalAppointments)),
      React.createElement(Text, { style: s.statLabel }, 'Total Appointments'),
    ),
    React.createElement(View, { style: s.statItem },
      React.createElement(Text, { style: s.statValue }, String(incomingCalls)),
      React.createElement(Text, { style: s.statLabel }, 'Incoming Calls'),
    ),
    React.createElement(View, { style: s.statItem },
      React.createElement(Text, { style: s.statValue }, String(outgoingCalls)),
      React.createElement(Text, { style: s.statLabel }, 'Outgoing Calls'),
    ),
  );

  const tableHeader = React.createElement(View, { style: s.tableHeader },
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'No.')),
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Patient\'s Name')),
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Office')),
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Appt. Date')),
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Visit Type')),
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Call In')),
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Call Out')),
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Time')),
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Remark')),
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Source')),
    React.createElement(View, { style: s.tableCell }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Other Duty')),
  );

  const tableRows = patientList.map((row: any, index: number) => {
    const safeName = safeStr(row?.name, 50);
    const safeOffice = safeStr(row?.office, 50);
    const safeApptDate = safeStr(row?.appt_date || row?.apptDate, 20);
    const safeVisitType = safeStr(row?.visit_type || row?.visitType, 50);
    const safeTime = safeStr(row?.time, 20);
    const safeRemark = safeStr(row?.remark, 100);
    const safeSource = safeStr(row?.source, 100);
    const safeOtherDuty = safeStr(row?.other_duty || row?.otherDuty, 100);
    const callIn = row?.call_in === true || row?.callIn === true;
    const callOut = row?.call_out === true || row?.callOut === true;

    return React.createElement(View, { key: index, style: s.tableRow },
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, String(index + 1))),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeName || '-')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeOffice || '-')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeApptDate || '-')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeVisitType || '-')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, callIn ? 'O' : '')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, callOut ? 'O' : '')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, convertTo12Hour(safeTime))),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeRemark || '-')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeSource || '-')),
      React.createElement(View, { style: s.tableCell }, React.createElement(Text, null, safeOtherDuty || '-')),
    );
  });

  const table = patientList.length > 0
    ? React.createElement(View, { style: s.table }, tableHeader, ...tableRows)
    : React.createElement(View, { style: { padding: 40, alignItems: 'center' } },
        React.createElement(Text, { style: { fontSize: 10, color: '#666' } }, 'No patient data recorded.'),
      );

  const dailyReport = safeDailyWorkReport
    ? React.createElement(View, { style: s.dailyReport },
        React.createElement(Text, { style: s.dailyReportTitle }, 'Daily Work Report'),
        React.createElement(Text, { style: s.dailyReportContent }, safeDailyWorkReport),
      )
    : null;

  const footer = React.createElement(View, { style: s.footer },
    React.createElement(Text, null, `Generated: ${generatedDate}`),
  );

  return React.createElement(Document, null,
    React.createElement(Page, { size: 'A4', orientation: 'portrait', style: s.page }, 
      header, 
      infoSection, 
      stats, 
      table, 
      dailyReport,
      footer
    ),
  );
}

export default function PatientLogSystem(): React.ReactElement {
  
  const [loading, setLoading] = useState(false);
  const [autoSaveStatus, setAutoSaveStatus] = useState('');
  const [submitStatus, setSubmitStatus] = useState('');
  const [progress, setProgress] = useState(0);
  const [previousFormData, setPreviousFormData] = useState<{
    dutyDate: string;
    userName: string;
    workOffice: string;
    workHoursFrom: string;
    workHoursTo: string;
  } | null>(null); 
  const previousFormDataRef = useRef<{
    dutyDate: string;
    userName: string;
    workOffice: string;
    workHoursFrom: string;
    workHoursTo: string;
  } | null>(null);
  const tableSaveEnabledRef = useRef(false);
  const submitInProgressRef = useRef(false); 
  const pendingAutoSaveRef = useRef<Promise<void> | null>(null); 
  const [isUnlocked, setIsUnlocked] = useState(false); 
  const [userOfficesOptions, setuserOfficesOptions] = useState<string[]>([]); 

  const [formData, setFormData] = useState({
    dutyDate: new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' }),
    userName: '',
    workOffice: '',
    workHoursFrom: '',
    workHoursTo: '',
    dailyWorkReport: ''
  });

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
      source: '',
      isOther: '',
      other_duty: ''
    }));
  });

  const appointments = patientRows.filter(row => row.appt_date && row.name).length;
  const incomingCalls = patientRows.filter(row => row.call_in).length;
  const outgoingCalls = patientRows.filter(row => row.call_out).length;

  const workOfficeOptions = ['Bernard', 'Call Center', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];
  const patientOfficeOptions = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

  const getVisitTypeOptions = useCallback((office: string): string[] => {
    if (office === 'Ortho') {
      return ['Adjustment', 'Bonding', 'Consult', 'FMS', 'Full Deband', 'Partial Deband', 'Records', 'Retainer Check', 'RPE Check', 'Seals'];
    } else if (office === 'California') {
      return ['Consult', 'Crown', 'Emergency', 'FMS', 'New Patient', 'RCRA', 'RCT', 'Recall', 'Seals', 'Tx'];
    } else {
      return ['Emergency', 'FMS', 'New Patient', 'RCRA', 'Recall', 'Seals', 'Tx'];
    }
  }, []);



  const remarkOptions = ['Disc', 'Elsewhere', 'LMA', 'LMW', 'NA', 'Not Interested', 'Will Call Back', 'Wrong'];

  const sourceOptions = ['Google', 'Social Media', 'Physician', 'Friend/Family', 'Insurance', 'Passed By'];

  const otherDutyOptions = [
    'Accounts with Balances', 'Booking ASL Interpreters', 'Break', 'Confirming', 
    'Incoming Call Report', 'Insurance Verifications', 'Lunch', 'Marketing Data', 
    'Medi-cal Eligibility', 'Monthly Report', 'Nintendo Switch Raffle', 'One Week', 
    'Other', 'Reconfirming', 'Postcards', 'Refer a friend', 'Reviews', 'Routing Slips', 
    'Sending Replacement Staff', 'Training', 'Working on Schedule'
  ];

  const generateDocId = (dutyDate: string, userName: string, workOffice: string, workHoursFrom: string, workHoursTo: string): string => {
    return `${dutyDate}_${userName}_${workOffice}_${workHoursFrom}_${workHoursTo}`.replace(/[^a-zA-Z0-9_-]/g, '_');
  };

  const hasPatientTableInput = (rows: typeof patientRows): boolean => {
    return rows.some(row =>
      Boolean(
        (typeof row.name === 'string' && row.name.trim()) ||
        row.office ||
        row.appt_date ||
        row.visit_type ||
        row.call_in ||
        row.call_out ||
        row.time ||
        row.remark ||
        row.source ||
        row.isOther ||
        row.other_duty
      )
    );
  };

  const hasDailyWorkReportInput = (report: string): boolean => {
    return typeof report === 'string' && report.trim().length > 0;
  };

  const hasSaveableContent = (): boolean => {
    return hasPatientTableInput(patientRows) || hasDailyWorkReportInput(formData.dailyWorkReport);
  };

  const isBasicInfoComplete = () => {
    return formData.dutyDate && 
           formData.userName && 
           formData.workOffice && 
           formData.workHoursFrom && 
           formData.workHoursTo;
  };

  const autoSave = useCallback(async () => {
    if (submitInProgressRef.current) return;
    if (!isUnlocked) return;
    
    if (!isBasicInfoComplete()) return;

    if (formData.userName.trim().endsWith(' ')) {
      return;
    }

    if (!tableSaveEnabledRef.current || !hasSaveableContent()) {
      return;
    }

    const saveTask = (async () => {
      const currentDocId = sanitizeFirebaseDocIdClient(generateDocId(formData.dutyDate, formData.userName, formData.workOffice, formData.workHoursFrom, formData.workHoursTo));
      
      let rowsToSave = patientRows;
      let previousDocIdToDelete: string | null = null;
      
      const prevFormData = previousFormDataRef.current;
      if (prevFormData && 
          (prevFormData.dutyDate !== formData.dutyDate ||
           prevFormData.userName !== formData.userName ||
           prevFormData.workOffice !== formData.workOffice ||
           prevFormData.workHoursFrom !== formData.workHoursFrom ||
           prevFormData.workHoursTo !== formData.workHoursTo)) {
        try {
          const prevDocId = sanitizeFirebaseDocIdClient(generateDocId(
            prevFormData.dutyDate,
            prevFormData.userName,
            prevFormData.workOffice,
            prevFormData.workHoursFrom,
            prevFormData.workHoursTo
          ));
          
          const previousDocRef = doc(db, "patient-logs", prevDocId);
          const previousDocSnap = await getDoc(previousDocRef);
          
          if (previousDocSnap.exists()) {
            const previousData = previousDocSnap.data();
            if (previousData.patientRows && Array.isArray(previousData.patientRows)) {
              rowsToSave = previousData.patientRows;
              if (!submitInProgressRef.current) {
                setPatientRows(previousData.patientRows);
              }
            }
            previousDocIdToDelete = prevDocId;
          }
        } catch (error) {
        }
      }

      if (submitInProgressRef.current) return;

      const dataToSave = {
        ...formData,
        patientRows: rowsToSave,
        timestamp: new Date().toISOString(),
        autoSaved: true
      };

      const safeDataToSave = sanitizeFirebaseDataClient(dataToSave);
      await setDoc(doc(db, "patient-logs", currentDocId), safeDataToSave);

      if (submitInProgressRef.current) return;

      if (previousDocIdToDelete && previousDocIdToDelete !== currentDocId) {
        try {
          await deleteDoc(doc(db, "patient-logs", previousDocIdToDelete));
        } catch (deleteError) {
        }
      }

      if (submitInProgressRef.current) return;

      const newPreviousFormData = {
        dutyDate: formData.dutyDate,
        userName: formData.userName,
        workOffice: formData.workOffice,
        workHoursFrom: formData.workHoursFrom,
        workHoursTo: formData.workHoursTo
      };
      setPreviousFormData(newPreviousFormData);
      previousFormDataRef.current = newPreviousFormData;
    })();

    pendingAutoSaveRef.current = saveTask;
    saveTask.catch(() => {
      if (!submitInProgressRef.current) {
        setAutoSaveStatus('❌ Save failed');
        setTimeout(() => {
          setAutoSaveStatus('');
        }, 2000);
      }
    }).finally(() => {
      if (pendingAutoSaveRef.current === saveTask) {
        pendingAutoSaveRef.current = null;
      }
    });
  }, [formData, patientRows, isUnlocked]);

  useEffect(() => {
    if (submitInProgressRef.current) return;
    if (!isUnlocked) return;
    
    if (!isBasicInfoComplete()) return;
    
    if (formData.userName.trim().endsWith(' ')) {
      return;
    }

    if (!hasSaveableContent() || !tableSaveEnabledRef.current) {
      return;
    }

    const timeoutId = setTimeout(() => {
      autoSave();
    }, 500);

    return () => clearTimeout(timeoutId);
  }, [formData, patientRows, isUnlocked, autoSave]);
  
  const handleFieldBlur = useCallback(() => {
    if (submitInProgressRef.current) return;
    if (!isUnlocked) return;
    
    const isComplete = formData.dutyDate && 
                       formData.userName && 
                       formData.workOffice && 
                       formData.workHoursFrom && 
                       formData.workHoursTo;
    if (isComplete) {
      autoSave();
    }
  }, [formData, autoSave, isUnlocked]);

  const loadExistingData = async () => {
    if (!isBasicInfoComplete()) {
      return;
    }

    try {
      const currentDocId = sanitizeFirebaseDocIdClient(generateDocId(formData.dutyDate, formData.userName, formData.workOffice, formData.workHoursFrom, formData.workHoursTo));
      
      if (previousFormData && 
          (previousFormData.dutyDate !== formData.dutyDate ||
           previousFormData.userName !== formData.userName ||
           previousFormData.workOffice !== formData.workOffice ||
           previousFormData.workHoursFrom !== formData.workHoursFrom ||
           previousFormData.workHoursTo !== formData.workHoursTo)) {
        try {
          const prevDocId = sanitizeFirebaseDocIdClient(generateDocId(
            previousFormData.dutyDate,
            previousFormData.userName,
            previousFormData.workOffice,
            previousFormData.workHoursFrom,
            previousFormData.workHoursTo
          ));
          
          const previousDocRef = doc(db, "patient-logs", prevDocId);
          const previousDocSnap = await getDoc(previousDocRef);
          
          if (previousDocSnap.exists()) {
            const previousData = previousDocSnap.data();

            if (previousData.patientRows && Array.isArray(previousData.patientRows)) {
              const loadedRows = previousData.patientRows.map((row: any, index: number) => ({
                ...row,
                id: index + 1
              }));
              
              const minRows = Math.max(loadedRows.length, 30);
              const newRows = Array.from({ length: minRows }, (_, index) => {
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
                  source: '',
                  isOther: '',
                  other_duty: ''
                };
              });
              
              setPatientRows(newRows);
            }
          }
        } catch (error) {
        }
      }
      
      const docRef = doc(db, "patient-logs", currentDocId);
      
      const docSnap = await getDoc(docRef);

      if (docSnap.exists()) {
        const matchingLog = docSnap.data();

        const newPreviousFormData = {
          dutyDate: formData.dutyDate,
          userName: formData.userName,
          workOffice: formData.workOffice,
          workHoursFrom: formData.workHoursFrom,
          workHoursTo: formData.workHoursTo
        };
        setPreviousFormData(newPreviousFormData);
        previousFormDataRef.current = newPreviousFormData;
        
        if (matchingLog && matchingLog.patientRows) {
          const loadedRows = matchingLog.patientRows.map((row: any, index: number) => ({
            ...row,
            id: index + 1
          }));
          
          const minRows = Math.max(loadedRows.length, 30);
          const newRows = Array.from({ length: minRows }, (_, index) => {
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
              source: '',
              isOther: '',
              other_duty: ''
            };
          });
          
          setPatientRows(newRows);
          
          if (matchingLog.dailyWorkReport) {
            setFormData(prev => ({
              ...prev,
              dailyWorkReport: matchingLog.dailyWorkReport
            }));
          }
        }
      } else {
        const newPreviousFormData = {
          dutyDate: formData.dutyDate,
          userName: formData.userName,
          workOffice: formData.workOffice,
          workHoursFrom: formData.workHoursFrom,
          workHoursTo: formData.workHoursTo
        };
        setPreviousFormData(newPreviousFormData);
        previousFormDataRef.current = newPreviousFormData;
      }
    } catch (error) {
    }
  };

  const checkUnsubmittedData = useCallback(async () => {
    if (!formData.userName || !formData.workOffice) return;

    try {
      const logsQuery = query(
        collection(db, "patient-logs"),
        where("userName", "==", formData.userName),
        where("workOffice", "==", formData.workOffice)
      );
      const logsSnapshot = await getDocs(logsQuery);
      const otherDates: string[] = [];

      for (const logDoc of logsSnapshot.docs) {
        const logDate = logDoc.data().dutyDate;
        if (logDate && logDate !== formData.dutyDate) otherDates.push(logDate);
      }

      if (otherDates.length > 0) {
        const datesList = [...new Set(otherDates)].sort().join(', ');
        alert(`You have data for these other dates: ${datesList}`);
      }
    } catch (error) {
    }
  }, [formData.userName, formData.workOffice, formData.dutyDate]);

  useEffect(() => {
    if (!isUnlocked) return;
    
    const timeoutId = setTimeout(() => {
      loadExistingData();
    }, 50);

    return () => clearTimeout(timeoutId);
  }, [formData.dutyDate, formData.userName, formData.workOffice, formData.workHoursFrom, formData.workHoursTo, isUnlocked]);

  useEffect(() => {
    if (!isUnlocked) return;
    
    const prevFormData = previousFormDataRef.current;
    if (!prevFormData) return;
    
    if (prevFormData.dutyDate !== formData.dutyDate ||
        prevFormData.userName !== formData.userName ||
        prevFormData.workOffice !== formData.workOffice ||
        prevFormData.workHoursFrom !== formData.workHoursFrom ||
        prevFormData.workHoursTo !== formData.workHoursTo) {
      tableSaveEnabledRef.current = false;
      setIsUnlocked(false);
      
      setPatientRows(() => {
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
          source: '',
          isOther: '',
          other_duty: ''
        }));
      });
      
      setFormData(prev => ({
        ...prev,
        dailyWorkReport: ''
      }));
      
      setPreviousFormData(null);
      previousFormDataRef.current = null;
    }
  }, [formData.dutyDate, formData.userName, formData.workOffice, formData.workHoursFrom, formData.workHoursTo, isUnlocked]);

  useEffect(() => {
    const timeoutId = setTimeout(() => {
      checkUnsubmittedData();
    }, 500); 

    return () => clearTimeout(timeoutId);
  }, [formData.userName, formData.workOffice, checkUnsubmittedData]);


  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      try {
        if (!currentUser) {
          setuserOfficesOptions([]);
          return;
        }

        const userDoc = await getDoc(doc(db, 'users', currentUser.uid));
        if (!userDoc.exists()) {
          setuserOfficesOptions([]);
          return;
        }

        const userData = userDoc.data();
        if (userData?.offices) {
          const officesArray = Array.isArray(userData.offices)
            ? userData.offices
            : [userData.offices];

          const validOptions = officesArray.filter((g: string) => workOfficeOptions.includes(g));

          if (validOptions.length > 0) {
            setuserOfficesOptions(validOptions);
            if (validOptions.length === 1) {
              setFormData(prev => ({ ...prev, workOffice: validOptions[0] }));
            }
          } else {
            setuserOfficesOptions([]);
          }
        } else {
          setuserOfficesOptions([]);
        }
      } catch (error) {
        setuserOfficesOptions([]);
      }
    });

    return () => unsubscribe();
  }, []);

  useEffect(() => {
    if (process.env.NODE_ENV === 'production' && 
        typeof window !== 'undefined' && 
        window.location.protocol !== 'https:') {
      window.location.href = window.location.href.replace('http:', 'https:');
    }
  }, []);

  const validateInput = useCallback((field: string, value: any): any => {
    if (typeof value === 'string') {
      const maxLengths: { [key: string]: number } = {
        dutyDate: 50,
        userName: 100,
        workOffice: 100,
        workHoursFrom: 20,
        workHoursTo: 20,
        dailyWorkReport: 2000,
        name: 100,
        office: 50,
        appt_date: 20,
        visit_type: 50,
        time: 20,
        remark: 200,
        source: 200,
        other_duty: 200
      };
      
      const maxLength = maxLengths[field] || 500;
      if (value.length > maxLength) {
        return value.slice(0, maxLength);
      }
      
      return value.replace(/[\x00-\x1F\x7F-\x9F]/g, '');
    }
    
    return value;
  }, []);

  const updateFormData = useCallback((field: string, value: any) => {
    const validatedValue = validateInput(field, value);
    setFormData(prev => {
      if ((prev as any)[field] === validatedValue) return prev;
      if (field === 'dailyWorkReport' && isUnlocked) {
        tableSaveEnabledRef.current = true;
      }
      return { ...prev, [field]: validatedValue };
    });
  }, [validateInput, isUnlocked]);


  const addPatientRow = useCallback(() => {
    setPatientRows(prevRows => {
      const newId = Math.max(...prevRows.map(row => row.id)) + 1;
      return [...prevRows, {
        id: newId,
        name: '',
        office: '',
        appt_date: '',
        visit_type: '',
        call_in: false,
        call_out: false,
        time: '',
        remark: '',
        source: '',
        isOther: '',
        other_duty: ''
      }];
    });
  }, []);

  const removePatientRow = useCallback((id: number) => {
    tableSaveEnabledRef.current = true;
    setPatientRows(prevRows => {
      if (prevRows.length > 1) {
        return prevRows.filter(row => row.id !== id);
      }
      return prevRows;
    });
  }, []);

  const updatePatientRow = useCallback((id: number, field: string, value: any) => {
    const validatedValue = validateInput(field, value);
    
    setPatientRows(prevRows => {
      const rowIndex = prevRows.findIndex(row => row.id === id);
      
      if (rowIndex === -1) {
        return prevRows;
      }
      
      const row = prevRows[rowIndex];
      
      if ((row as any)[field] === validatedValue) {
        return prevRows;
      }

      tableSaveEnabledRef.current = true;
      
      const updatedRow = { ...row, [field]: validatedValue };
      
      if (field === 'office' && row.office !== validatedValue) {
        updatedRow.visit_type = '';
      }
      if ((field === 'call_in' || field === 'call_out') && validatedValue === true) {
        const now = new Date();
        const timeString = now.toTimeString().slice(0, 5);
        updatedRow.time = timeString;
      }
      if ((field === 'call_in' && validatedValue === false && !row.call_out) || 
          (field === 'call_out' && validatedValue === false && !row.call_in)) {
        updatedRow.time = '';
      }
      
      const newRows = [...prevRows];
      newRows[rowIndex] = updatedRow;
      return newRows;
    });
  }, [validateInput]);

  const resetForm = () => {
    setFormData({
      dutyDate: new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' }),
      userName: '',
      workOffice: '',
      workHoursFrom: '',
      workHoursTo: '',
      dailyWorkReport: ''
    });
    setPatientRows([{
      id: 1,
      name: '',
      office: '',
      appt_date: '',
      visit_type: '',
      call_in: false,
      call_out: false,
      time: '',
      remark: '',
      source: '',
      isOther: '',
      other_duty: ''
    }]);
    setPreviousFormData(null);
    previousFormDataRef.current = null;
    tableSaveEnabledRef.current = false;
    setIsUnlocked(false);
  };

  const handleSubmit = async () => {
    if (loading) {
      return;
    }

    if (!isBasicInfoComplete()) {
      alert('⚠️ Please fill in all Basic Information fields.');
      return;
    }

    const auth = getAuth();
    const currentUser = auth.currentUser;
    if (!currentUser) {
      alert('⚠️ Please log in.');
      return;
    }

    submitInProgressRef.current = true;
    try {
      if (pendingAutoSaveRef.current) {
        try {
          await pendingAutoSaveRef.current;
        } catch {
        }
      }

      setLoading(true);
      setSubmitStatus('Saving...');
      setProgress(10);

      const currentDocId = sanitizeFirebaseDocIdClient(generateDocId(formData.dutyDate, formData.userName, formData.workOffice, formData.workHoursFrom, formData.workHoursTo));

      setSubmitStatus('Submitting...');
      setProgress(30);
      
      const patientListForPdf = patientRows.filter(row => 
        row.name || row.office || row.appt_date || row.visit_type || 
        row.call_in || row.call_out || row.time || row.remark || row.source || row.other_duty
      );
      
      const totalAppointments = patientListForPdf.filter(row => row.appt_date && row.name).length;
      const incomingCalls = patientListForPdf.filter(row => row.call_in).length;
      const outgoingCalls = patientListForPdf.filter(row => row.call_out).length;
      
      const generatedDate = new Date().toLocaleDateString('en-US', { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        hour12: true
      });
      
      const safeDutyDate = (formData.dutyDate || '').trim().slice(0, 50).replace(/[<>]/g, '');
      const safeUserName = (formData.userName || '').trim().slice(0, 100).replace(/[<>]/g, '');
      const safeWorkOffice = (formData.workOffice || '').trim().slice(0, 100).replace(/[<>]/g, '');
      const safeWorkHoursFrom = (formData.workHoursFrom || '').trim().slice(0, 20).replace(/[<>]/g, '');
      const safeWorkHoursTo = (formData.workHoursTo || '').trim().slice(0, 20).replace(/[<>]/g, '');
      const safeDailyWorkReport = (formData.dailyWorkReport || '').trim().slice(0, 2000).replace(/[<>]/g, '');
      
      setSubmitStatus('Processing...');
      setProgress(50);
      
      const pdfDoc = createPatientLogPDFDocument({
        safeDutyDate,
        safeUserName,
        safeWorkOffice,
        safeWorkHoursFrom,
        safeWorkHoursTo,
        safeDailyWorkReport,
        patientList: patientListForPdf,
        totalAppointments,
        incomingCalls,
        outgoingCalls,
        generatedDate,
      });
      
      setSubmitStatus('Processing...');
      setProgress(60);
      
      const pdfBlob = await pdf(pdfDoc).toBlob();
      
      const safeDate = (formData.dutyDate || new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' }))
        .replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 50);
      const safeName = (formData.userName || 'Unknown')
        .replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 50);
      const safeOffice = (formData.workOffice || 'Unknown')
        .replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 50);
      const tsNow = new Date();
      const tsLaTime = new Date(tsNow.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
      let tsHours = tsLaTime.getHours();
      const tsMinutes = tsLaTime.getMinutes();
      const tsAmpm = tsHours >= 12 ? 'pm' : 'am';
      tsHours = tsHours % 12;
      tsHours = tsHours ? tsHours : 12;
      const timeStamp = `${tsHours}${tsMinutes.toString().padStart(2, '0')}${tsAmpm}`;
      const filename = `7) ${safeDate}_${safeOffice}_${safeName}_Patient Log_${timeStamp}.pdf`.slice(0, 255);
      
      setSubmitStatus('Saving...');
      setProgress(70);
      
      try {
        const storage = getStorage();
        const storageRef = ref(storage, `endofday-pdfs/Call_Center/${safeDate}/${filename}`);
        
        await uploadBytes(storageRef, pdfBlob);
        
        try {
          await deleteDoc(doc(db, 'patient-logs', currentDocId));
        } catch (deleteError) {
        }
        
        const rowsWithDate = patientListForPdf.filter(row => row.appt_date && row.appt_date.trim() !== '');
        const byApptDate = new Map<string, { name: string; office: string }[]>();
        for (const row of rowsWithDate) {
          const d = (row.appt_date || '').trim();
          if (!byApptDate.has(d)) byApptDate.set(d, []);
          byApptDate.get(d)!.push({
            name: safeStr(row.name, 100),
            office: safeStr(row.office, 100),
          });
        }

        for (const [apptDate, patients] of byApptDate) {
          const showDocId = apptDate.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 1500);
          const docRef = doc(db, 'show-noshow', showDocId);
          const sanitizedPatients = patients.map((p) => sanitizeFirebaseDataClient(p));
          await setDoc(docRef, {
            ...sanitizeFirebaseDataClient({ appt_date: apptDate }),
            patients: arrayUnion(...sanitizedPatients),
          }, { merge: true });
        }
        if (byApptDate.size === 0) {
          const showDocId = formData.dutyDate.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 1500);
          const docRef = doc(db, 'show-noshow', showDocId);
          await setDoc(docRef, sanitizeFirebaseDataClient({
            appt_date: formData.dutyDate,
          }), { merge: true });
        }
        
        setSubmitStatus('Submitted Successfully!');
        setProgress(100);
        
        resetForm();
        
        setTimeout(() => {
          setLoading(false);
          setSubmitStatus('');
          setProgress(0);
        }, 2000);
      } catch (storageError: any) {
        const errorMessage = storageError?.message || 'Error';
        
        const sensitiveKeywords = ['password', 'token', 'secret', 'key', 'credential', 'auth', 'login', 'session', 'cookie', 'bearer', 'jwt', 'api', 'apikey'];
        const hasSensitiveInfo = sensitiveKeywords.some(keyword => 
          errorMessage.toLowerCase().includes(keyword.toLowerCase())
        );
        
        const safeErrorMessage = hasSensitiveInfo 
          ? 'Error.' 
          : (errorMessage.length > 100 ? errorMessage.substring(0, 100) + '...' : errorMessage).replace(/[<>\"'&]/g, '');
        
        alert(`Error: ${safeErrorMessage}`);
        setLoading(false);
        setSubmitStatus('');
        setProgress(0);
      }

    } catch (error) {
      const errorMessage = (error as any).message || 'Error';
      
      const sensitiveKeywords = ['password', 'token', 'secret', 'key', 'credential', 'auth', 'login', 'session', 'cookie', 'bearer', 'jwt', 'api', 'apikey'];
      const hasSensitiveInfo = sensitiveKeywords.some(keyword => 
        errorMessage.toLowerCase().includes(keyword.toLowerCase())
      );
      
      const safeErrorMessage = hasSensitiveInfo 
        ? 'An error occurred while submitting.' 
        : (errorMessage.length > 100 ? errorMessage.substring(0, 100) + '...' : errorMessage).replace(/[<>\"'&]/g, '');
      
      setSubmitStatus('❌ Submission failed: ' + safeErrorMessage);
      setProgress(0);
      
      alert('❌ Submission failed. Please try again.');
      
      setTimeout(() => {
        setLoading(false);
        setSubmitStatus('');
        setProgress(0);
      }, 5000);
    } finally {
      submitInProgressRef.current = false;
    }
  };

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

  const containerStyle = {
    maxWidth: '2000px',
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

  return (
    <>
      <div style={bodyStyle}>
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
          }}>
            <div style={{
              backgroundColor: "white",
              padding: "40px",
              borderRadius: "20px",
              textAlign: "center",
              boxShadow: "0 20px 40px rgba(0, 0, 0, 0.3)",
              maxWidth: "400px",
              width: "90%"
            }}>
              <div style={{
                fontSize: "18px",
                fontWeight: "600",
                color: "#4a6fa1",
                marginBottom: "20px"
              }}>
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
                }}>
                  <div style={{
                    width: `${progress}%`,
                    height: "100%",
                    background: "linear-gradient(90deg, #4a90e2, #51cf66)",
                    transition: "width 0.3s ease",
                    borderRadius: "4px"
                  }} />
                </div>
              )}
              <div style={{
                fontSize: "14px",
                color: "#666",
                marginTop: "10px"
              }}>
                {progress}%
              </div>
            </div>
          </div>
        )}

        <div style={containerStyle}>
        <div style={{ position: 'relative' }}>
        <h1 style={{ 
          color: '#0077B6', 
          textAlign: 'center', 
          marginBottom: '20px', 
          fontSize: '2.5rem', 
          fontWeight: 'bold',
          textShadow: '2px 2px 4px rgba(0,0,0,0.1)'
        }}>Patient Log</h1>
          {autoSaveStatus && (
            <div style={{
              position: 'absolute',
              top: '10px',
              right: '10px',
              padding: '8px 16px',
              backgroundColor: autoSaveStatus.includes('❌') ? '#ff6b6b' : '#51cf66',
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

        <div style={sectionStyle}>
          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap', marginBottom: '20px' }}>
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Duty Date:
              </label>
              <input
                type="date"
                value={formData.dutyDate}
                onChange={(e) => updateFormData('dutyDate', e.target.value)}
                onBlur={handleFieldBlur}
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
                onBlur={handleFieldBlur}
                placeholder="Enter your full name"
                style={inputStyle}
                required
              />
            </div>

            {userOfficesOptions.length > 0 && (
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Work Office:
              </label>
              {userOfficesOptions.length === 1 ? (
                <div style={{
                  ...inputStyle,
                  display: 'flex',
                  alignItems: 'center',
                  backgroundColor: '#e9ecef',
                  fontWeight: '600',
                  color: '#023047'
                }}>
                  {formData.workOffice}
                </div>
              ) : (
                <select
                  value={formData.workOffice}
                  onChange={(e) => updateFormData('workOffice', e.target.value)}
                  onBlur={handleFieldBlur}
                  style={inputStyle}
                  required
                >
                  <option value="">Select Office</option>
                  {userOfficesOptions.map(office => (
                    <option key={office} value={office}>{office}</option>
                  ))}
                </select>
              )}
            </div>
            )}
          </div>

          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap', marginBottom: '20px' }}>
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Work Hours From:
              </label>
              <input
                type="time"
                value={formData.workHoursFrom}
                onChange={(e) => updateFormData('workHoursFrom', e.target.value)}
                onBlur={handleFieldBlur}
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
                onBlur={handleFieldBlur}
                style={inputStyle}
                required
              />
            </div>
          </div>

          {!isUnlocked && (
            <div style={{ display: 'flex', justifyContent: 'center', marginTop: '20px' }}>
              <button
                onClick={async () => {
                  if (isBasicInfoComplete()) {
                    tableSaveEnabledRef.current = false;
                    await loadExistingData();
                    setIsUnlocked(true);
                  } else {
                    alert('⚠️ Please fill in all Basic Information fields.');
                  }
                }}
                style={{
                  ...buttonStyle,
                  backgroundColor: '#0077B6',
                  padding: '12px 30px',
                  fontSize: '1.1em'
                }}
              >
                🔓 Unlock
              </button>
            </div>
          )}
        </div>

        <div style={sectionStyle}>
          {formData.userName && formData.userName.trim().endsWith(' ') && (
            <div style={{ marginBottom: '20px' }}>
              <span style={{ 
                fontSize: '12px', 
                color: '#ffc107',
                fontWeight: 'bold',
                padding: '4px 8px',
                backgroundColor: '#fff3cd',
                borderRadius: '12px',
                border: '1px solid #ffc107',
              }}>
                ⏳ Typing...
              </span>
            </div>
          )}

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
                {appointments}
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
                {incomingCalls}
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
                {outgoingCalls}
              </span>
            </div>
          </div>

          {!isUnlocked ? (
            <div style={{
              padding: '40px',
              textAlign: 'center',
              backgroundColor: '#f8f9fa',
              border: '2px dashed #dee2e6',
              borderRadius: '8px',
              margin: '20px 0'
            }}>
              <div style={{ fontSize: '48px', marginBottom: '20px' }}>🔒</div>
            </div>
          ) : (
            <>
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', margin: '10px 0' }}>
              <thead style={{ backgroundColor: '#0077B6', color: 'white' }}>
                <tr>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '60px' }}>#</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '150px' }}>Name</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '100px' }}>Office</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '120px' }}>Appt. Date</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '120px' }}>Type of Visit</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '80px' }}>Call In</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '80px' }}>Call Out</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '80px' }}>Time</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '150px' }}>Remark</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '150px' }}>Source</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '150px' }}>Other Duty</th>
                  <th style={{ padding: '12px 8px', textAlign: 'center', minWidth: '80px' }}>Action</th>
                </tr>
              </thead>
              <tbody>
                {patientRows.map((row) => (
                  <PatientRow
                    key={row.id}
                    row={row}
                    updatePatientRow={updatePatientRow}
                    removePatientRow={removePatientRow}
                    patientOfficeOptions={patientOfficeOptions}
                    getVisitTypeOptions={getVisitTypeOptions}
                    remarkOptions={remarkOptions}
                    sourceOptions={sourceOptions}
                    otherDutyOptions={otherDutyOptions}
                    inputStyle={inputStyle}
                    buttonStyle={buttonStyle}
                  />
                ))}
              </tbody>
            </table>
            </div>

            <div style={{ display: 'flex', justifyContent: 'center', marginTop: '20px' }}>
              <button onClick={addPatientRow} style={buttonStyle}>
                + Add
              </button>
            </div>
            </>
          )}
        </div>

        <div style={sectionStyle}>
          <h2 style={{ color: '#0077B6', marginBottom: '15px' }}>Daily Work Report</h2>
          {!isUnlocked ? (
            <div style={{
              padding: '20px',
              textAlign: 'center',
              backgroundColor: '#f8f9fa',
              border: '2px dashed #dee2e6',
              borderRadius: '8px',
              margin: '10px 0'
            }}>
              <div style={{ fontSize: '24px', marginBottom: '10px' }}>🔒</div>
            </div>
          ) : (
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
          )}
        </div>

        <div style={{ display: 'flex', justifyContent: 'center', gap: '20px', marginBottom: '30px' }}>
          {!isUnlocked ? (
            <div style={{
              padding: '20px',
              textAlign: 'center',
              backgroundColor: '#f8f9fa',
              border: '2px dashed #dee2e6',
              borderRadius: '8px',
              margin: '10px 0'
            }}>
              <div style={{ fontSize: '24px', marginBottom: '10px' }}>🔒</div>
            </div>
          ) : (
            <button 
              onClick={handleSubmit} 
              disabled={loading} 
              style={{ ...buttonStyle, backgroundColor: '#28a745' }}
            >
              {loading ? 'Submitting...' : 'Submit'}
            </button>
          )}
        </div>

      </div>
    </div>
    </>
  );
}
