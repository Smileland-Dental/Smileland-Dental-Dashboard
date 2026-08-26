'use client';

import React, { useState, useEffect, useCallback, useRef } from "react";
import { doc, setDoc, getDoc, deleteDoc } from "firebase/firestore";
import { db, auth } from "@/lib/firebase.config";
import { getStorage, ref, uploadBytes } from 'firebase/storage';
import { onAuthStateChanged } from 'firebase/auth';
import { pdf, Document, Page, View, Text, StyleSheet } from '@react-pdf/renderer';

const validateInput = (value: string, maxLength: number = 500): string => {
  if (typeof value !== 'string') return '';
  if (value.length > maxLength) {
    return value.substring(0, maxLength);
  }
  return value;
};

const validateDate = (date: string): boolean => {
  if (!date || typeof date !== 'string') return false;
  const dateRegex1 = /^\d{2}\/\d{2}\/\d{4}$/;
  const dateRegex2 = /^\d{4}-\d{2}-\d{2}$/;
  return dateRegex1.test(date) || dateRegex2.test(date);
};

const OFFICE_OPTIONS = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

const validateOffice = (office: string): boolean => {
  if (!office || typeof office !== 'string') return false;
  return OFFICE_OPTIONS.includes(office);
};

const FIXED_FORM_NAMES = [
            'Daily Front Office Duties',
            'Attendance Tract Sheet',
            'Clock Adjustment',
            'Excuse Note',
            'Time Off Request',
            'Request for Sched. Change',
            'Written Warning',
            'Record of Conversation',
            'Incident Notice',
            'Restroom Log',
            'Add On Treatment Log',
            'Scheduled Appts Log',
            'Mileage Log',
            'Lobby Inspection Log',
            'RDA Sheets',
            'X-Ray/IOPs Before Treatment',
            'Covid-19 Daily Screening Log',
            'Other:',
            'Other:'
          ];

const getCaliforniaFormattedDate = (): string => {
  const californiaTime = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
  return californiaTime.toLocaleDateString('en-US', {
    month: '2-digit',
    day: '2-digit',
    year: 'numeric',
  });
};

const convertDateToISO = (dateStr: string): string => {
  if (!dateStr) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return dateStr;
  const parts = dateStr.split('/');
  if (parts.length === 3) {
    const month = parts[0].padStart(2, '0');
    const day = parts[1].padStart(2, '0');
    const year = parts[2];
    return `${year}-${month}-${day}`;
  }
  return '';
};

const getCaliforniaISODate = (): string => convertDateToISO(getCaliforniaFormattedDate());

const formatDateForDisplay = (dateStr: string): string => {
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    const [year, month, day] = dateStr.split('-');
    return `${month}/${day}/${year}`;
  }
  return dateStr;
};

const convertTo12Hour = (timeStr: string): string => {
  if (!timeStr) return '12:00 AM';
  if (timeStr.includes('AM') || timeStr.includes('PM')) return timeStr;
  const parts = timeStr.split(':');
  if (parts.length === 2) {
    let hour = parseInt(parts[0], 10);
    const minute = parts[1].trim();
    const ampm = hour >= 12 ? 'PM' : 'AM';
    hour = hour % 12;
    if (hour === 0) hour = 12;
    return `${hour.toString().padStart(2, '0')}:${minute} ${ampm}`;
  }
  return '12:00 AM';
};

const convertTo24Hour = (timeStr: string): string => {
  if (!timeStr || !timeStr.trim()) return '';
  const trimmed = timeStr.trim();

  if (trimmed.includes('AM') || trimmed.includes('PM')) {
    const parts = trimmed.split(' ');
    const timePart = parts[0] || '';
    const ampm = parts[1] || 'AM';
    const [hourStr, minute = '00'] = timePart.split(':');
    let hour = parseInt(hourStr, 10);
    if (ampm === 'PM' && hour !== 12) hour += 12;
    if (ampm === 'AM' && hour === 12) hour = 0;
    return `${hour.toString().padStart(2, '0')}:${minute.padStart(2, '0')}`;
  }

  const timeParts = trimmed.split(':');
  if (timeParts.length === 2) {
    const hour = parseInt(timeParts[0], 10);
    const minute = timeParts[1].trim();
    if (Number.isNaN(hour)) return '';
    return `${hour.toString().padStart(2, '0')}:${minute.padStart(2, '0')}`;
  }

  return '';
};

const PRODUCTION_ROW_COUNT = 2;

const createEmptyProductionRow = () => ({ date: '', note: '', status: '' });

const createInitialData = () => ({
  formData: {
    date: getCaliforniaISODate(),
    officeTimeCheckIn: '',
    officeName: '',
    timeCheckOut: '',
    name: '',
  },
  tableData: FIXED_FORM_NAMES.map(nameOfForm => ({
    nameOfForm,
    otherText: '',
    qty: '',
  })),
  productionData: Array.from({ length: PRODUCTION_ROW_COUNT }, createEmptyProductionRow),
  todayData: { addOns: '', noShows: '', seen: '' },
  nextDayData: { opener: '', closer: '' },
  callLogData: { whoCalled: '', appointmentsMade: '' },
  supervisorData: { officeSupervisorManager: '', checkOutBy: '' },
});

type FormDataState = ReturnType<typeof createInitialData>['formData'];
type TableRow = ReturnType<typeof createInitialData>['tableData'][number];
type ProductionRow = ReturnType<typeof createInitialData>['productionData'][number];
type TodayState = ReturnType<typeof createInitialData>['todayData'];
type NextDayState = ReturnType<typeof createInitialData>['nextDayData'];
type CallLogState = ReturnType<typeof createInitialData>['callLogData'];
type SupervisorState = ReturnType<typeof createInitialData>['supervisorData'];

type FormState = {
  formData: FormDataState;
  tableData: TableRow[];
  productionData: ProductionRow[];
  todayData: TodayState;
  nextDayData: NextDayState;
  callLogData: CallLogState;
  supervisorData: SupervisorState;
};

const serializeFormState = (state: FormState) => JSON.stringify(state);

const buildDocId = (selectedOffice: string, isoDate: string) => {
  const safeOffice = selectedOffice.replace(/[^a-zA-Z0-9_-]/g, '');
  const safeIsoDate = isoDate.replace(/[^a-zA-Z0-9_-]/g, '');
  if (!safeOffice || !safeIsoDate) return null;
  const docId = `${safeIsoDate}_${safeOffice}_fax_cover`;
  if (docId.length > 1500) return null;
  return { docId, safeOffice, safeIsoDate };
};

const validateAllFormData = (params: {
  formData: FormDataState;
  tableData: TableRow[];
  productionData: ProductionRow[];
  todayData: TodayState;
  nextDayData: NextDayState;
  callLogData: CallLogState;
  supervisorData: SupervisorState;
}) => ({
  validatedFormData: {
    date: validateInput(params.formData.date, 20),
    officeTimeCheckIn: validateInput(params.formData.officeTimeCheckIn, 20),
    officeName: validateInput(params.formData.officeName, 100),
    timeCheckOut: validateInput(params.formData.timeCheckOut, 20),
    name: validateInput(params.formData.name, 100),
  },
  validatedTableData: params.tableData.map(row => ({
    nameOfForm: validateInput(row.nameOfForm, 200),
    otherText: validateInput(row.otherText, 200),
    qty: validateInput(row.qty, 20),
  })),
  validatedProductionData: params.productionData.map(row => ({
    date: validateInput(row.date, 20),
    note: validateInput(row.note, 500),
    status: validateInput(row.status, 50),
  })),
  validatedTodayData: {
    addOns: validateInput(params.todayData.addOns, 500),
    noShows: validateInput(params.todayData.noShows, 500),
    seen: validateInput(params.todayData.seen, 500),
  },
  validatedNextDayData: {
    opener: validateInput(params.nextDayData.opener, 200),
    closer: validateInput(params.nextDayData.closer, 200),
  },
  validatedCallLogData: {
    whoCalled: validateInput(params.callLogData.whoCalled, 500),
    appointmentsMade: validateInput(params.callLogData.appointmentsMade, 500),
  },
  validatedSupervisorData: {
    officeSupervisorManager: validateInput(params.supervisorData.officeSupervisorManager, 200),
    checkOutBy: validateInput(params.supervisorData.checkOutBy, 200),
  },
});

const parseFirestoreDocument = (data: Record<string, unknown>): FormState => {
  const initial = createInitialData();
  const dateValue = convertDateToISO((data.date as string) || '') || convertDateToISO((data.faxDate as string) || '') || getCaliforniaISODate();

  const tableData = data.tableData && Array.isArray(data.tableData)
    ? FIXED_FORM_NAMES.map((fixedName, index) => {
        const savedRow = (data.tableData as Array<{ otherText?: string; qty?: string }>)[index];
        return {
          nameOfForm: fixedName,
          otherText: savedRow?.otherText || '',
          qty: savedRow?.qty || '',
        };
      }).slice(0, 19)
    : initial.tableData;

  return {
    formData: {
      date: dateValue,
      officeTimeCheckIn: data.officeTimeCheckIn ? convertTo24Hour(data.officeTimeCheckIn as string) : '',
      officeName: (data.officeName as string) || '',
      timeCheckOut: data.timeCheckOut ? convertTo24Hour(data.timeCheckOut as string) : '',
      name: (data.name as string) || '',
    },
    tableData,
    productionData: data.productionData && Array.isArray(data.productionData)
      ? Array.from({ length: PRODUCTION_ROW_COUNT }, (_, i) => {
          const savedRow = (data.productionData as ProductionRow[])[i];
          return {
            date: savedRow?.date || '',
            note: savedRow?.note || '',
            status: savedRow?.status || '',
          };
        })
      : initial.productionData,
    todayData: data.todayData
      ? {
          addOns: (data.todayData as TodayState).addOns || '',
          noShows: (data.todayData as TodayState).noShows || '',
          seen: (data.todayData as TodayState).seen || '',
        }
      : initial.todayData,
    nextDayData: data.nextDayData
      ? {
          opener: (data.nextDayData as NextDayState).opener || '',
          closer: (data.nextDayData as NextDayState).closer || '',
        }
      : initial.nextDayData,
    callLogData: data.callLogData
      ? {
          whoCalled: (data.callLogData as CallLogState).whoCalled || '',
          appointmentsMade: (data.callLogData as CallLogState).appointmentsMade || '',
        }
      : initial.callLogData,
    supervisorData: data.supervisorData
      ? {
          officeSupervisorManager: (data.supervisorData as SupervisorState).officeSupervisorManager || '',
          checkOutBy: (data.supervisorData as SupervisorState).checkOutBy || '',
        }
      : initial.supervisorData,
  };
};

export default function FaxCoverPage() {
  const pdfStyles = StyleSheet.create({
    page: { padding: 15, fontFamily: 'Helvetica', fontSize: 8 },
    header: { marginBottom: 6, borderBottomWidth: 1, borderColor: '#333', paddingBottom: 3, alignItems: 'center' },
    headerTitle: { fontSize: 16, fontWeight: 'bold', marginBottom: 2 },
    headerSubtitle: { fontSize: 9, color: '#666' },
    infoSection: { flexDirection: 'row', justifyContent: 'space-between', marginBottom: 5, fontSize: 8 },
    infoItem: { flexDirection: 'row', gap: 3 },
    infoLabel: { fontWeight: 'bold' },
    table: { marginTop: 5, marginBottom: 5 },
    tableRow: { flexDirection: 'row', borderBottomWidth: 0.5, borderColor: '#333' },
    tableCell: { padding: 2.5, fontSize: 7, flex: 1, borderRightWidth: 0.5, borderColor: '#333', justifyContent: 'center', alignItems: 'center' },
    tableCellNo: { flex: 0.3 },
    tableCellName: { flex: 2 },
    tableCellQty: { flex: 0.8 },
    tableHeader: { backgroundColor: '#f0f0f0', fontWeight: 'bold' },
    section: { marginTop: 6, marginBottom: 4 },
    sectionTitle: { fontSize: 10, fontWeight: 'bold', marginBottom: 2, borderBottomWidth: 1, borderColor: '#ddd', paddingBottom: 1 },
    sectionContent: { fontSize: 8, marginBottom: 2 },
    sectionRow: { flexDirection: 'row', marginBottom: 2 },
    sectionLabel: { fontWeight: 'bold', marginRight: 4 },
    sideBySideRow: { flexDirection: 'row', gap: 10, marginTop: 6, marginBottom: 4 },
    sideBySideColumn: { flex: 1 },
    footer: { marginTop: 8, paddingTop: 5, borderTopWidth: 1, borderColor: '#ddd', alignItems: 'center', fontSize: 7, color: '#666' },
  });

  function safeStr(v: unknown, max: number): string {
    if (v == null) return '';
    return String(v).trim().slice(0, max).replace(/[<>]/g, '');
  }

  function createFaxCoverPDFDocument(props: {
    safeSelectedOffice: string;
    formData: {
      date: string;
      officeTimeCheckIn: string;
      officeName: string;
      timeCheckOut: string;
      name: string;
    };
    tableData: Array<{
      nameOfForm: string;
      otherText: string;
      qty: string;
    }>;
    productionData: Array<{
      date: string;
      note: string;
      status: string;
    }>;
    todayData: {
      addOns: string;
      noShows: string;
      seen: string;
    };
    nextDayData: {
      opener: string;
      closer: string;
    };
    callLogData: {
      whoCalled: string;
      appointmentsMade: string;
    };
    supervisorData: {
      officeSupervisorManager: string;
      checkOutBy: string;
    };
    generatedDate: string;
  }) {
    const { safeSelectedOffice, formData, tableData, productionData, todayData, nextDayData, callLogData, supervisorData, generatedDate } = props;
    const s = pdfStyles;

    const header = React.createElement(View, { style: s.header },
      React.createElement(Text, { style: s.headerTitle }, 'End of Day Fax Cover'),
      React.createElement(Text, { style: s.headerSubtitle }, '(Check out only when leaving the office)'),
    );

    const infoSection1 = React.createElement(View, { style: s.infoSection },
      React.createElement(View, { style: s.infoItem },
        React.createElement(Text, { style: s.infoLabel }, 'Date: '),
        React.createElement(Text, null, safeStr(formatDateForDisplay(formData.date), 20)),
      ),
      React.createElement(View, { style: s.infoItem },
        React.createElement(Text, { style: s.infoLabel }, 'Office: '),
        React.createElement(Text, null, safeSelectedOffice),
      ),
    );

    const infoSection2 = React.createElement(View, { style: s.infoSection },
      React.createElement(View, { style: s.infoItem },
        React.createElement(Text, { style: s.infoLabel }, 'Time Check In: '),
        React.createElement(Text, null, safeStr(formData.officeTimeCheckIn ? convertTo12Hour(formData.officeTimeCheckIn) : '', 20)),
      ),
      React.createElement(View, { style: s.infoItem },
        React.createElement(Text, { style: s.infoLabel }, 'Name: '),
        React.createElement(Text, null, safeStr(formData.officeName, 100)),
      ),
      React.createElement(View, { style: s.infoItem },
        React.createElement(Text, { style: s.infoLabel }, 'Time Check Out: '),
        React.createElement(Text, null, safeStr(formData.timeCheckOut ? convertTo12Hour(formData.timeCheckOut) : '', 20)),
      ),
      React.createElement(View, { style: s.infoItem },
        React.createElement(Text, { style: s.infoLabel }, 'Name: '),
        React.createElement(Text, null, safeStr(formData.name, 100)),
      ),
    );

    const tableHeader = React.createElement(View, { style: [s.tableRow, s.tableHeader] },
      React.createElement(View, { style: [s.tableCell, s.tableCellNo] }, React.createElement(Text, null, 'No.')),
      React.createElement(View, { style: [s.tableCell, s.tableCellName] }, React.createElement(Text, null, 'Name of Form')),
      React.createElement(View, { style: [s.tableCell, s.tableCellQty] }, React.createElement(Text, null, 'Qty')),
    );

    const tableRows = tableData.map((row, index) => {
      const formName = row.nameOfForm === 'Other:' 
        ? `Other: ${safeStr(row.otherText, 200)}`
        : safeStr(row.nameOfForm, 200);
      return React.createElement(View, { key: index, style: s.tableRow },
        React.createElement(View, { style: [s.tableCell, s.tableCellNo] }, React.createElement(Text, null, String(index + 1))),
        React.createElement(View, { style: [s.tableCell, s.tableCellName] }, React.createElement(Text, null, formName)),
        React.createElement(View, { style: [s.tableCell, s.tableCellQty] }, React.createElement(Text, null, safeStr(row.qty, 20))),
      );
    });

    const totalPages = tableData.reduce((sum, row) => {
      const qty = parseFloat(safeStr(row.qty, 20)) || 0;
      return Math.max(0, Math.min(sum + qty, 999999));
    }, 0);
    const totalRow = React.createElement(View, { style: [s.tableRow, { backgroundColor: '#f8f9fa' }] },
      React.createElement(View, { style: [s.tableCell, s.tableCellNo, { flex: 2.3 }] }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Total Pages')),
      React.createElement(View, { style: [s.tableCell, s.tableCellQty] }, React.createElement(Text, { style: { fontWeight: 'bold' } }, String(totalPages))),
    );

    const table = React.createElement(View, { style: s.table }, tableHeader, ...tableRows, totalRow);

    const productionHeader = React.createElement(View, { style: [s.tableRow, s.tableHeader] },
      React.createElement(View, { style: [s.tableCell, { flex: 1 }] }, React.createElement(Text, null, 'Date')),
      React.createElement(View, { style: [s.tableCell, { flex: 2 }] }, React.createElement(Text, null, 'Note')),
      React.createElement(View, { style: [s.tableCell, { flex: 0.8 }] }, React.createElement(Text, null, 'Status')),
    );
    const productionRows = productionData.map((row, index) => 
      React.createElement(View, { key: index, style: s.tableRow },
        React.createElement(View, { style: [s.tableCell, { flex: 1 }] }, React.createElement(Text, null, safeStr(row.date, 20))),
        React.createElement(View, { style: [s.tableCell, { flex: 2 }] }, React.createElement(Text, null, safeStr(row.note, 500))),
        React.createElement(View, { style: [s.tableCell, { flex: 0.8 }] }, React.createElement(Text, null, safeStr(row.status, 50))),
      )
    );
    const productionTable = React.createElement(View, { style: s.table }, productionHeader, ...productionRows);

    const todayNextDaySection = React.createElement(View, { style: s.sideBySideRow },
      React.createElement(View, { style: s.sideBySideColumn },
        React.createElement(Text, { style: s.sectionTitle }, 'Today'),
        React.createElement(View, { style: s.sectionContent },
          React.createElement(View, { style: s.sectionRow },
            React.createElement(Text, { style: s.sectionLabel }, 'Add On\'s: '),
            React.createElement(Text, null, safeStr(todayData.addOns, 500)),
          ),
          React.createElement(View, { style: s.sectionRow },
            React.createElement(Text, { style: s.sectionLabel }, 'No Shows: '),
            React.createElement(Text, null, safeStr(todayData.noShows, 500)),
          ),
          React.createElement(View, { style: s.sectionRow },
            React.createElement(Text, { style: s.sectionLabel }, 'Seen: '),
            React.createElement(Text, null, safeStr(todayData.seen, 500)),
          ),
        ),
      ),
      React.createElement(View, { style: s.sideBySideColumn },
        React.createElement(Text, { style: s.sectionTitle }, 'Next Day'),
        React.createElement(View, { style: s.sectionContent },
          React.createElement(View, { style: s.sectionRow },
            React.createElement(Text, { style: s.sectionLabel }, 'Opener: '),
            React.createElement(Text, null, safeStr(nextDayData.opener, 200)),
          ),
          React.createElement(View, { style: s.sectionRow },
            React.createElement(Text, { style: s.sectionLabel }, 'Closer: '),
            React.createElement(Text, null, safeStr(nextDayData.closer, 200)),
          ),
        ),
      ),
    );

    const callLogSection = React.createElement(View, { style: s.section },
      React.createElement(Text, { style: s.sectionTitle }, 'Call Log'),
      React.createElement(View, { style: s.sectionContent },
        React.createElement(View, { style: s.sectionRow },
          React.createElement(Text, { style: s.sectionLabel }, 'Who called: '),
          React.createElement(Text, null, safeStr(callLogData.whoCalled, 500)),
        ),
        React.createElement(View, { style: s.sectionRow },
          React.createElement(Text, { style: s.sectionLabel }, 'Appointments made: '),
          React.createElement(Text, null, safeStr(callLogData.appointmentsMade, 500)),
        ),
      ),
    );

    const supervisorSection = React.createElement(View, { style: s.section },
      React.createElement(Text, { style: s.sectionTitle }, 'Office Supervisor/Manager'),
      React.createElement(View, { style: s.sectionContent },
        React.createElement(View, { style: s.sectionRow },
          React.createElement(Text, { style: s.sectionLabel }, 'Supervisor/Manager: '),
          React.createElement(Text, null, safeStr(supervisorData.officeSupervisorManager, 200)),
        ),
        React.createElement(View, { style: s.sectionRow },
          React.createElement(Text, { style: s.sectionLabel }, 'Check out by: '),
          React.createElement(Text, null, safeStr(supervisorData.checkOutBy, 200)),
        ),
      ),
    );

    const footer = React.createElement(View, { style: s.footer }, React.createElement(Text, null, `Generated: ${generatedDate}`));

    return React.createElement(Document, null,
      React.createElement(Page, { size: 'A4', orientation: 'portrait', style: s.page }, 
        header, 
        infoSection1, 
        infoSection2, 
        table, 
        productionTable, 
        todayNextDaySection, 
        callLogSection, 
        supervisorSection, 
        footer
      ),
    );
  }

  const initialData = createInitialData();

  const [loading, setLoading] = useState(false);
  const [submitStatus, setSubmitStatus] = useState('');
  const [progress, setProgress] = useState(0);
  const [isUpdatingFromFirebase, setIsUpdatingFromFirebase] = useState(false);
  const [userOfficeBasedOptions, setuserOfficeBasedOptions] = useState<string[]>([]);

  const [faxDate, setFaxDate] = useState(() => {
    return convertDateToISO(getCaliforniaFormattedDate());
  });

  const [selectedOffice, setSelectedOffice] = useState('');

  const [formData, setFormData] = useState(initialData.formData);

  const [productionData, setProductionData] = useState(initialData.productionData);
  const [todayData, setTodayData] = useState(initialData.todayData);
  const [nextDayData, setNextDayData] = useState(initialData.nextDayData);
  const [callLogData, setCallLogData] = useState(initialData.callLogData);
  const [supervisorData, setSupervisorData] = useState(initialData.supervisorData);
  const [tableData, setTableData] = useState(initialData.tableData);

  const [lastSavedSnapshot, setLastSavedSnapshot] = useState(() => serializeFormState(initialData));
  const isSwitchingDateRef = useRef(false);

  const getCurrentFormState = useCallback((): FormState => ({
    formData,
    tableData,
    productionData,
    todayData,
    nextDayData,
    callLogData,
    supervisorData,
  }), [formData, tableData, productionData, todayData, nextDayData, callLogData, supervisorData]);

  const isFormDirty = useCallback(
    () => serializeFormState(getCurrentFormState()) !== lastSavedSnapshot,
    [getCurrentFormState, lastSavedSnapshot]
  );

  const applyFormSnapshot = useCallback((snapshot: FormState) => {
    setIsUpdatingFromFirebase(true);
    setFormData(prev => ({ ...prev, ...snapshot.formData }));
    setTableData(snapshot.tableData);
    setProductionData(snapshot.productionData);
    setTodayData(snapshot.todayData);
    setNextDayData(snapshot.nextDayData);
    setCallLogData(snapshot.callLogData);
    setSupervisorData(snapshot.supervisorData);
    setLastSavedSnapshot(serializeFormState(snapshot));
    setTimeout(() => setIsUpdatingFromFirebase(false), 100);
  }, []);

  const resetToInitialData = useCallback((preserveDate?: string) => {
    const initial = createInitialData();
    if (preserveDate) {
      initial.formData.date = preserveDate;
    }
    setFormData(initial.formData);
    setTableData(initial.tableData);
    setProductionData(initial.productionData);
    setTodayData(initial.todayData);
    setNextDayData(initial.nextDayData);
    setCallLogData(initial.callLogData);
    setSupervisorData(initial.supervisorData);
    setLastSavedSnapshot(serializeFormState(initial));
  }, []);

  const autoSave = useCallback(async () => {
    if (isSwitchingDateRef.current) return;
    if (!faxDate || isUpdatingFromFirebase || !selectedOffice || !validateOffice(selectedOffice)) return;

    const current = getCurrentFormState();
    if (current.formData.date && !validateDate(current.formData.date)) return;

    const currentIso = convertDateToISO(current.formData.date);
    if (currentIso && currentIso !== faxDate) return;
    if (!isFormDirty()) return;

    try {
      const validated = validateAllFormData(current);
      const docInfo = buildDocId(selectedOffice, faxDate);
      if (!docInfo) return;

      await setDoc(doc(db, "fax-cover", docInfo.docId), {
        faxDate,
        selectedOffice,
        ...validated.validatedFormData,
        tableData: validated.validatedTableData,
        productionData: validated.validatedProductionData,
        todayData: validated.validatedTodayData,
        nextDayData: validated.validatedNextDayData,
        callLogData: validated.validatedCallLogData,
        supervisorData: validated.validatedSupervisorData,
      });

      await setDoc(doc(db, "simple-forms", `${docInfo.safeIsoDate}_${docInfo.safeOffice}`), {
        productionSideMetrics: {
          add: validated.validatedTodayData.addOns,
          noShow: validated.validatedTodayData.noShows,
          seen: validated.validatedTodayData.seen,
        },
        checkIn: validated.validatedFormData.officeTimeCheckIn,
        checkOut: validated.validatedFormData.timeCheckOut,
        closer: validated.validatedFormData.name,
      },
      { merge: true }
    );

      setLastSavedSnapshot(serializeFormState(getCurrentFormState()));
    } catch (error) {
    }
  }, [faxDate, selectedOffice, getCurrentFormState, isFormDirty, isUpdatingFromFirebase]);

  const hasUserInput = useCallback(() => {
    return Boolean(
      formData.officeTimeCheckIn ||
      formData.officeName ||
      formData.timeCheckOut ||
      formData.name ||
      tableData.some(row => row.otherText || row.qty) ||
      productionData.some(row => row.date || row.note || row.status) ||
      Object.values(todayData).some(value => value !== '') ||
      Object.values(nextDayData).some(value => value !== '') ||
      Object.values(callLogData).some(value => value !== '') ||
      Object.values(supervisorData).some(value => value !== '')
    );
  }, [formData, tableData, productionData, todayData, nextDayData, callLogData, supervisorData]);

  useEffect(() => {
    if (hasUserInput() && isFormDirty()) {
      autoSave();
    }
  }, [formData, tableData, productionData, todayData, nextDayData, callLogData, supervisorData, hasUserInput, isFormDirty, autoSave]);

  useEffect(() => {
    if (formData.date && selectedOffice) {
      const isoDate = convertDateToISO(formData.date);
      if (isoDate) {
        setFaxDate(isoDate);
      }
    }
  }, [formData.date, selectedOffice]);

  useEffect(() => {
    if (!selectedOffice) return;

    const isoDate = faxDate || convertDateToISO(formData.date);
    if (!isoDate) return;

    const docInfo = buildDocId(selectedOffice, isoDate);
    if (!docInfo) return;

    let cancelled = false;

    const loadData = async () => {
      try {
        setIsUpdatingFromFirebase(true);
        const docSnap = await getDoc(doc(db, "fax-cover", docInfo.docId));
        if (cancelled) return;

        if (docSnap.exists()) {
          applyFormSnapshot(parseFirestoreDocument(docSnap.data()));
        } else {
          resetToInitialData(isoDate);
          setTimeout(() => setIsUpdatingFromFirebase(false), 100);
        }
      } catch (error) {
        if (!cancelled) setIsUpdatingFromFirebase(false);
      } finally {
        if (!cancelled) isSwitchingDateRef.current = false;
      }
    };

    loadData();

    return () => {
      cancelled = true;
    };
  }, [faxDate, selectedOffice, applyFormSnapshot, resetToInitialData]);

  useEffect(() => {
    const loadUserOffice = async (currentUser: { uid: string } | null) => {
      if (!currentUser) return;

      try {
        const userDoc = await getDoc(doc(db, 'users', currentUser.uid));
        if (!userDoc.exists()) return;

        const userData = userDoc.data();

        if (userData?.offices) {
          const officeBasedArray = Array.isArray(userData.offices)
            ? userData.offices
            : [userData.offices];

          const validOptions = officeBasedArray.filter((g: string) => OFFICE_OPTIONS.includes(g));

          if (validOptions.length > 0) {
            setuserOfficeBasedOptions(validOptions);
            
            if (validOptions.length === 1) {
              setSelectedOffice(validOptions[0]);
            }
          }
        }
      } catch (error) {
      }
    };

    const unsubscribe = onAuthStateChanged(auth, loadUserOffice);

    if (process.env.NODE_ENV === 'production' &&
        typeof window !== 'undefined' &&
        window.location.protocol !== 'https:') {
      window.location.href = window.location.href.replace('http:', 'https:');
    }

    return () => {
      unsubscribe();
    };
  }, []);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value } = e.target;
    setFormData((prev) => ({
      ...prev,
      [name]: value,
    }));
  };

  const handleTableChange = (index: number, field: 'otherText' | 'qty', value: string) => {
    setTableData(prev => {
      const newData = [...prev];
      newData[index] = {
        ...newData[index],
        [field]: value
      };
      return newData;
    });
  };

  const handleSubmit = async () => {
    if (!selectedOffice) {
      alert('Please select an office');
      return;
    }

    if (!validateOffice(selectedOffice)) {
      alert('Invalid office value');
      return;
    }

    if (formData.date && !validateDate(formData.date)) {
      alert('Invalid date format');
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

      const validated = validateAllFormData({
        formData,
        tableData,
        productionData,
        todayData,
        nextDayData,
        callLogData,
        supervisorData,
      });

      const isoDate = convertDateToISO(validated.validatedFormData.date) || faxDate;
      const docInfo = buildDocId(selectedOffice, isoDate);
      if (!docInfo) {
        alert('Invalid document information');
        setLoading(false);
        return;
      }

      const { safeOffice, safeIsoDate, docId } = docInfo;
      const californiaTime = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
      const generatedDate = californiaTime.toLocaleString('en-US', {
        month: '2-digit',
        day: '2-digit',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        hour12: true,
      });

      setSubmitStatus('Generating PDF...');
      setProgress(30);

      try {
        const blob = await pdf(createFaxCoverPDFDocument({
          safeSelectedOffice: safeOffice,
          formData: validated.validatedFormData,
          tableData: validated.validatedTableData,
          productionData: validated.validatedProductionData,
          todayData: validated.validatedTodayData,
          nextDayData: validated.validatedNextDayData,
          callLogData: validated.validatedCallLogData,
          supervisorData: validated.validatedSupervisorData,
          generatedDate,
        })).toBlob();

        setSubmitStatus('Processing...');
        setProgress(60);

        setSubmitStatus('Saving...');
        setProgress(70);

        const storage = getStorage();
        const tsNow = new Date();
        const tsLaTime = new Date(tsNow.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
        let tsHours = tsLaTime.getHours();
        const tsMinutes = tsLaTime.getMinutes();
        const tsAmpm = tsHours >= 12 ? 'pm' : 'am';
        tsHours = tsHours % 12;
        tsHours = tsHours ? tsHours : 12;
        const timeStamp = `${tsHours}${tsMinutes.toString().padStart(2, '0')}${tsAmpm}`;

        const filename = `1) ${safeIsoDate}_${safeOffice || 'Unknown'}_End of Day Fax Cover_${timeStamp}.pdf`;
        const storageRef = ref(storage, `endofday-pdfs/${safeOffice}/${safeIsoDate}/${filename}`);

        await uploadBytes(storageRef, blob);

        await setDoc(doc(db, "simple-forms", `${safeIsoDate}_${safeOffice}`), {
          productionSideMetrics: {
            add: validated.validatedTodayData.addOns,
            noShow: validated.validatedTodayData.noShows,
            seen: validated.validatedTodayData.seen,
          },
          checkIn: validated.validatedFormData.officeTimeCheckIn,
          checkOut: validated.validatedFormData.timeCheckOut,
          closer: validated.validatedFormData.name,
        },
        { merge: true }
      );

        setSubmitStatus('✅ Complete!');

        setSubmitStatus('Cleaning up...');
        setProgress(80);

        setIsUpdatingFromFirebase(true);

        try {
          await deleteDoc(doc(db, "fax-cover", docId));
        } catch (deleteError) {
        }

        resetToInitialData();

        setSubmitStatus('Complete!');
        setProgress(100);

        setTimeout(() => {
          setIsUpdatingFromFirebase(false);
        }, 2000);

        setTimeout(() => {
          setLoading(false);
          setSubmitStatus('');
          setProgress(0);
        }, 2000);
      } catch (submitError) {
        alert('Submission failed. Please try again.');
        setSubmitStatus('❌ Submission failed. Please try again.');
        setProgress(0);
        setTimeout(() => {
          setLoading(false);
          setSubmitStatus('');
        }, 3000);
      }
    } catch (error: any) {
      setSubmitStatus('❌ Submission failed. Please try again.');
      setProgress(0);
      setTimeout(() => {
        setLoading(false);
        setSubmitStatus('');
      }, 3000);
    }
  };

  const styles: {
    body: React.CSSProperties;
    container: React.CSSProperties;
    header: React.CSSProperties;
    subtitle: React.CSSProperties;
    formGroup: React.CSSProperties;
    label: React.CSSProperties;
    input: React.CSSProperties;
    submitButton: React.CSSProperties;
    statusMessage: React.CSSProperties;
    table: React.CSSProperties;
    th: React.CSSProperties;
    td: React.CSSProperties;
  } = {
    body: {
      fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
      padding: '20px',
      background: 'linear-gradient(135deg, #f5f7fa 0%, #e8ecf1 100%)',
      color: '#2c3e50',
      lineHeight: '1.6',
      minHeight: '100vh'
    },
    container: {
      maxWidth: '67%',
      width: '67%',
      margin: '20px auto',
      padding: '30px',
      backgroundColor: 'white',
      borderRadius: '8px',
      boxShadow: '0 4px 8px rgba(0, 0, 0, 0.1)',
      position: 'relative'
    },
    header: {
      color: '#2c3e50',
      textAlign: 'center',
      marginBottom: '10px',
      paddingBottom: '10px',
      borderBottom: '2px solid #e9ecef',
      fontSize: '2em',
      fontWeight: 'bold'
    },
    subtitle: {
      fontSize: '0.875rem',
      color: '#666',
      textAlign: 'center',
      marginBottom: '30px',
      fontStyle: 'italic'
    },
    formGroup: {
      marginBottom: '25px'
    },
    label: {
      display: 'block',
      fontWeight: '500',
      color: '#2c3e50',
      marginBottom: '8px'
    },
    input: {
      width: '100%',
      padding: '10px 12px',
      fontSize: '14px',
      border: '1px solid #e9ecef',
      borderRadius: '4px',
      backgroundColor: 'white',
      boxSizing: 'border-box',
      transition: 'border-color 0.2s'
    },
    submitButton: {
      display: 'block',
      width: '150px',
      margin: '30px auto 0 auto',
      padding: '12px 20px',
      backgroundColor: '#3498db',
      color: 'white',
      border: 'none',
      borderRadius: '4px',
      cursor: 'pointer',
      fontSize: '16px',
      transition: 'background-color 0.2s'
    },
    statusMessage: {
      marginTop: '15px',
      fontWeight: 'bold',
      textAlign: 'center',
      padding: '10px',
      borderRadius: '4px'
    },
    table: {
      borderCollapse: 'collapse',
      width: '100%',
      marginTop: '20px',
      backgroundColor: 'white',
      boxShadow: '0 1px 3px rgba(0, 0, 0, 0.1)'
    },
    th: {
      border: '1px solid #e9ecef',
      padding: '10px',
      textAlign: 'left',
      verticalAlign: 'top',
      backgroundColor: '#2c3e50',
      color: 'white',
      fontWeight: '500'
    },
    td: {
      border: '1px solid #e9ecef',
      padding: '10px',
      textAlign: 'left',
      verticalAlign: 'top'
    }
  };

  useEffect(() => {
    const handleBeforeUnload = (e: BeforeUnloadEvent) => {
      if (loading) {
        e.preventDefault();
        e.returnValue = '';
      }
    };

    const handlePopState = (e: PopStateEvent) => {
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
    <>
      <style>{`
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
      `}</style>
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
              {!submitStatus && 'Processing... Please wait'}
            </p>
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
                ⚠️ Please do not close.
              </p>
            </div>
          </div>
        </div>
      )}

      <div style={styles.container}>
        <h2 style={styles.header}>End of Day Fax Cover</h2>
        <p style={styles.subtitle}>(Check out only when leaving the office)</p>

        <div style={{ display: 'flex', gap: '20px', marginBottom: '20px', flexWrap: 'wrap' }}>
          <div style={{ ...styles.formGroup, flex: '1', minWidth: '200px' }}>
            <label style={styles.label} htmlFor="faxDate">Date:</label>
            <input
              type="date"
              id="faxDate"
              value={convertDateToISO(formData.date)}
              onChange={(e) => {
                const nextDate = e.target.value;
                if (!nextDate || !validateDate(nextDate)) return;
                if (nextDate === convertDateToISO(formData.date)) return;
                isSwitchingDateRef.current = true;
                setFormData(prev => ({ ...prev, date: nextDate }));
              }}
              style={styles.input}
            />
          </div>
          {userOfficeBasedOptions.length > 0 && (
            <div style={{ ...styles.formGroup, flex: '1', minWidth: '200px' }}>
              <label style={styles.label} htmlFor="selectedOffice">Office:</label>
              {userOfficeBasedOptions.length === 1 ? (
                <div style={{
                  ...styles.input,
                  display: 'flex',
                  alignItems: 'center',
                  backgroundColor: '#e9ecef',
                  fontWeight: '600',
                  color: '#2c3e50'
                }}>
                  {selectedOffice}
                </div>
              ) : (
                <select
                  id="selectedOffice"
                  value={selectedOffice}
                  onChange={(e) => setSelectedOffice(e.target.value)}
                  style={styles.input}
                >
                  <option value="">-- Select Office --</option>
                  {userOfficeBasedOptions.map(office => (
                    <option key={office} value={office}>{office}</option>
                  ))}
                </select>
              )}
            </div>
          )}
        </div>

        {selectedOffice && (
        <form style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap' }}>
            <div style={{ ...styles.formGroup, flex: '1', minWidth: '150px' }}>
              <label htmlFor="officeTimeCheckIn" style={styles.label}>
                Time Check In
              </label>
              <input
                type="time"
                id="officeTimeCheckIn"
                name="officeTimeCheckIn"
                value={formData.officeTimeCheckIn}
                onChange={(e) => {
                  setFormData(prev => ({ ...prev, officeTimeCheckIn: e.target.value }));
                }}
                style={styles.input}
              />
            </div>

            <div style={{ ...styles.formGroup, flex: '1', minWidth: '150px' }}>
              <label htmlFor="officeName" style={styles.label}>
                Name
              </label>
              <input
                type="text"
                id="officeName"
                name="officeName"
                value={formData.officeName}
                onChange={handleChange}
                style={styles.input}
              />
            </div>

            <div style={{ ...styles.formGroup, flex: '1', minWidth: '150px' }}>
              <label htmlFor="timeCheckOut" style={styles.label}>
                Time Check Out
              </label>
              <input
                type="time"
                id="timeCheckOut"
                name="timeCheckOut"
                value={formData.timeCheckOut}
                onChange={(e) => {
                  setFormData(prev => ({ ...prev, timeCheckOut: e.target.value }));
                }}
                style={styles.input}
              />
            </div>

            <div style={{ ...styles.formGroup, flex: '1', minWidth: '150px' }}>
              <label htmlFor="name" style={styles.label}>
                Name
              </label>
              <input
                type="text"
                id="name"
                name="name"
                value={formData.name}
                onChange={handleChange}
                style={styles.input}
              />
            </div>
          </div>
        </form>
        )}

        {selectedOffice && (
        <table style={styles.table}>
          <thead>
            <tr>
              <th style={{ ...styles.th, width: '60px' }}>No.</th>
              <th style={styles.th}>Name of Form</th>
              <th style={styles.th}>Qty</th>
            </tr>
          </thead>
          <tbody>
            {tableData.map((row, index) => (
              <tr key={index}>
                <td style={{ ...styles.td, textAlign: 'center', fontWeight: 'bold' }}>
                  {index + 1}
                </td>
                <td style={styles.td}>
                  {row.nameOfForm === 'Other:' ? (
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                      {/* "Other:" 텍스트 */}
                      <span style={{
                        fontSize: '13px',
                        color: '#2c3e50',
                        fontWeight: '500',
                        whiteSpace: 'nowrap'
                      }}>
                        Other:
                      </span>
                    <input
                      type="text"
                      value={row.otherText}
                      onChange={(e) => {
                        const validated = validateInput(e.target.value, 200);
                        handleTableChange(index, 'otherText', validated);
                      }}
                      maxLength={200}
                      style={{
                        ...styles.input,
                        margin: 0,
                        fontSize: '13px',
                        padding: '6px 8px',
                        flex: 1,
                        boxSizing: 'border-box'
                      }}
                    />
                    </div>
                  ) : (
                    <div style={{
                      padding: '6px 8px',
                      fontSize: '13px',
                      color: '#2c3e50',
                      fontWeight: row.nameOfForm ? '500' : 'normal'
                    }}>
                      {row.nameOfForm || '\u00A0'}
                    </div>
                  )}
                </td>
                <td style={styles.td}>
                  <input
                    type="text"
                    value={row.qty}
                    onChange={(e) => {
                      const validated = validateInput(e.target.value, 20);
                      handleTableChange(index, 'qty', validated);
                    }}
                    maxLength={20}
                    style={{
                      ...styles.input,
                      margin: 0,
                      fontSize: '13px',
                      padding: '6px 8px',
                      width: '100%',
                      boxSizing: 'border-box'
                    }}
                  />
                </td>
              </tr>
            ))}
            <tr style={{ backgroundColor: '#f8f9fa', fontWeight: 'bold' }}>
              <td style={{ ...styles.td, textAlign: 'center' }} colSpan={2}>
                Total Pages
              </td>
              <td style={{ ...styles.td, textAlign: 'center', fontWeight: 'bold', fontSize: '16px' }}>
                {tableData.reduce((sum, row) => {
                  const qty = parseFloat(row.qty) || 0;
                  return sum + qty;
                }, 0)}
              </td>
            </tr>
          </tbody>
        </table>
        )}
        {selectedOffice && (
        <div style={{ marginTop: '40px', paddingTop: '20px', borderTop: '2px solid #e9ecef' }}>
          <h3 style={{
            color: '#2c3e50',
            fontSize: '1.5em',
            fontWeight: 'bold',
            marginBottom: '20px'
          }}>
            Production
          </h3>
          {productionData.map((row, index) => (
            <div key={index} style={{ ...styles.formGroup, marginBottom: '20px' }}>
              <div style={{ display: 'flex', gap: '20px', alignItems: 'flex-start', flexWrap: 'wrap' }}>
                {/* Date 입력 */}
                <div style={{ flex: '1', minWidth: '150px' }}>
                  <label style={styles.label}>Date</label>
                  <input
                    type="date"
                    value={row.date}
                    onChange={(e) => {
                      const newData = [...productionData];
                      newData[index] = { ...newData[index], date: e.target.value };
                      setProductionData(newData);
                    }}
                    style={styles.input}
                  />
                </div>
                <div style={{ flex: '2', minWidth: '200px' }}>
                  <label style={styles.label}>Note</label>
                  <input
                    type="text"
                    value={row.note}
                    onChange={(e) => {
                      const newData = [...productionData];
                      newData[index] = { ...newData[index], note: e.target.value };
                      setProductionData(newData);
                    }}
                    style={styles.input}
                  />
                </div>
                <div style={{ flex: '1', minWidth: '150px' }}>
                  <label style={styles.label}>Status</label>
                  <div style={{ display: 'flex', gap: '15px', alignItems: 'center', marginTop: '8px' }}>
                    <label style={{ display: 'flex', alignItems: 'center', gap: '6px', cursor: 'pointer' }}>
                      <input
                        type="radio"
                        name={`productionStatus_${index}`}
                        value="final"
                        checked={row.status === 'final'}
                        onChange={(e) => {
                          const newData = [...productionData];
                          newData[index] = { ...newData[index], status: e.target.value };
                          setProductionData(newData);
                        }}
                        style={{ cursor: 'pointer' }}
                      />
                      <span style={{ fontSize: '14px', color: '#2c3e50' }}>Final</span>
                    </label>
                    <label style={{ display: 'flex', alignItems: 'center', gap: '6px', cursor: 'pointer' }}>
                      <input
                        type="radio"
                        name={`productionStatus_${index}`}
                        value="not final"
                        checked={row.status === 'not final'}
                        onChange={(e) => {
                          const newData = [...productionData];
                          newData[index] = { ...newData[index], status: e.target.value };
                          setProductionData(newData);
                        }}
                        style={{ cursor: 'pointer' }}
                      />
                      <span style={{ fontSize: '14px', color: '#2c3e50' }}>Not Final</span>
                    </label>
                  </div>
                </div>
              </div>
            </div>
          ))}
        </div>
        )}

        {selectedOffice && (
        <div style={{ marginTop: '40px', paddingTop: '20px', borderTop: '2px solid #e9ecef' }}>
          <h3 style={{
            color: '#2c3e50',
            fontSize: '1.5em',
            fontWeight: 'bold',
            marginBottom: '20px'
          }}>
            Today
          </h3>

          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap' }}>
            <div style={{ ...styles.formGroup, flex: '1', minWidth: '150px' }}>
              <label htmlFor="addOns" style={styles.label}>
                Add On's
              </label>
              <input
                type="text"
                id="addOns"
                name="addOns"
                value={todayData.addOns}
                onChange={(e) => {
                  const validated = validateInput(e.target.value, 500);
                  setTodayData(prev => ({ ...prev, addOns: validated }));
                }}
                maxLength={500}
                style={styles.input}
              />
            </div>

            <div style={{ ...styles.formGroup, flex: '1', minWidth: '150px' }}>
              <label htmlFor="noShows" style={styles.label}>
                No Shows
              </label>
              <input
                type="text"
                id="noShows"
                name="noShows"
                value={todayData.noShows}
                onChange={(e) => {
                  const validated = validateInput(e.target.value, 500);
                  setTodayData(prev => ({ ...prev, noShows: validated }));
                }}
                style={styles.input}
              />
            </div>

            <div style={{ ...styles.formGroup, flex: '1', minWidth: '150px' }}>
              <label htmlFor="seen" style={styles.label}>
                Seen
              </label>
              <input
                type="text"
                id="seen"
                name="seen"
                value={todayData.seen}
                onChange={(e) => {
                  const validated = validateInput(e.target.value, 500);
                  setTodayData(prev => ({ ...prev, seen: validated }));
                }}
                maxLength={500}
                style={styles.input}
              />
            </div>
          </div>
        </div>
        )}

        {selectedOffice && (
        <div style={{ marginTop: '40px', paddingTop: '20px', borderTop: '2px solid #e9ecef' }}>
          <h3 style={{
            color: '#2c3e50',
            fontSize: '1.5em',
            fontWeight: 'bold',
            marginBottom: '20px'
          }}>
            Next Day
          </h3>

          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap' }}>
            <div style={{ ...styles.formGroup, flex: '1', minWidth: '150px' }}>
              <label htmlFor="opener" style={styles.label}>
                Opener
              </label>
              <input
                type="text"
                id="opener"
                name="opener"
                value={nextDayData.opener}
                onChange={(e) => {
                  const validated = validateInput(e.target.value, 200);
                  setNextDayData(prev => ({ ...prev, opener: validated }));
                }}
                style={styles.input}
              />
            </div>

            <div style={{ ...styles.formGroup, flex: '1', minWidth: '150px' }}>
              <label htmlFor="closer" style={styles.label}>
                Closer
              </label>
              <input
                type="text"
                id="closer"
                name="closer"
                value={nextDayData.closer}
                onChange={(e) => {
                  const validated = validateInput(e.target.value, 200);
                  setNextDayData(prev => ({ ...prev, closer: validated }));
                }}
                maxLength={200}
                style={styles.input}
              />
            </div>
          </div>
        </div>
        )}

        {selectedOffice && (
        <div style={{ marginTop: '40px', paddingTop: '20px', borderTop: '2px solid #e9ecef' }}>
          <h3 style={{
            color: '#2c3e50',
            fontSize: '1.5em',
            fontWeight: 'bold',
            marginBottom: '20px'
          }}>
            Call Log
          </h3>

          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap' }}>
            <div style={{ ...styles.formGroup, flex: '1', minWidth: '150px' }}>
              <label htmlFor="whoCalled" style={styles.label}>
                Who called
              </label>
              <input
                type="text"
                id="whoCalled"
                name="whoCalled"
                value={callLogData.whoCalled}
                onChange={(e) => {
                  const validated = validateInput(e.target.value, 500);
                  setCallLogData(prev => ({ ...prev, whoCalled: validated }));
                }}
                style={styles.input}
              />
            </div>

            <div style={{ ...styles.formGroup, flex: '1', minWidth: '150px' }}>
              <label htmlFor="appointmentsMade" style={styles.label}>
                How many appointments made
              </label>
              <input
                type="text"
                id="appointmentsMade"
                name="appointmentsMade"
                value={callLogData.appointmentsMade}
                onChange={(e) => {
                  const validated = validateInput(e.target.value, 500);
                  setCallLogData(prev => ({ ...prev, appointmentsMade: validated }));
                }}
                maxLength={500}
                style={styles.input}
              />
            </div>
          </div>
        </div>
        )}

        {selectedOffice && (
        <div style={{ marginTop: '40px', paddingTop: '20px', borderTop: '2px solid #e9ecef' }}>
          <h3 style={{
            color: '#2c3e50',
            fontSize: '1.5em',
            fontWeight: 'bold',
            marginBottom: '20px'
          }}>
            (For Corporate Use Only)
          </h3>

          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap' }}>
            <div style={{ ...styles.formGroup, flex: '1', minWidth: '150px' }}>
              <label htmlFor="officeSupervisorManager" style={styles.label}>
                Office Supervisor/Manager
              </label>
              <input
                type="text"
                id="officeSupervisorManager"
                name="officeSupervisorManager"
                value={supervisorData.officeSupervisorManager}
                onChange={(e) => {
                  const validated = validateInput(e.target.value, 200);
                  setSupervisorData(prev => ({ ...prev, officeSupervisorManager: validated }));
                }}
                maxLength={200}
                style={styles.input}
              />
            </div>

            <div style={{ ...styles.formGroup, flex: '1', minWidth: '150px' }}>
              <label htmlFor="checkOutBy" style={styles.label}>
                Check out by
              </label>
              <input
                type="text"
                id="checkOutBy"
                name="checkOutBy"
                value={supervisorData.checkOutBy}
                onChange={(e) => {
                  const validated = validateInput(e.target.value, 200);
                  setSupervisorData(prev => ({ ...prev, checkOutBy: validated }));
                }}
                maxLength={200}
                style={styles.input}
              />
            </div>
          </div>
        </div>
        )}

        {selectedOffice && (
        <div style={{ textAlign: 'center', marginTop: '20px' }}>
          <button
            onClick={handleSubmit}
            disabled={loading}
            style={{
              ...styles.submitButton,
              backgroundColor: loading ? '#bdc3c7' : '#3498db',
              cursor: loading ? 'not-allowed' : 'pointer'
            }}
          >
            {loading ? 'Submitting...' : 'Submit Button'}
          </button>
        </div>
        )}

        {submitStatus && (
          <div style={{
            ...styles.statusMessage,
            backgroundColor: submitStatus.includes('failed') || submitStatus.includes('Error') ? '#f8d7da' : 
                           submitStatus.includes('successfully') || submitStatus.includes('Complete') ? '#d4edda' : '#d1ecf1',
            color: submitStatus.includes('failed') || submitStatus.includes('Error') ? '#721c24' : 
                   submitStatus.includes('successfully') || submitStatus.includes('Complete') ? '#155724' : '#0c5460',
            border: submitStatus.includes('failed') || submitStatus.includes('Error') ? '1px solid #f5c6cb' : 
                    submitStatus.includes('successfully') || submitStatus.includes('Complete') ? '1px solid #c3e6cb' : '1px solid #bee5eb'
          }}>
            {submitStatus}
          </div>
        )}

      </div>
    </div>
    </>
  );
}
