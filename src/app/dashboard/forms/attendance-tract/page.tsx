'use client';

import React, { useState, useEffect, useCallback, useRef } from 'react';
import { db } from '@/lib/firebase.config';
import { doc, setDoc, getDoc, Timestamp, onSnapshot, deleteDoc } from 'firebase/firestore';
import { getStorage, ref, uploadBytes, getDownloadURL } from 'firebase/storage';

interface StaffMember {
  no: number;
  name: string;
  active?: boolean;
}

interface StaffList {
  [position: string]: StaffMember[];
}

interface DoctorMember {
  no: number;
  name: string;
}

interface AttendanceRow {
  date: string;
  filledBy: string;
  checkedBy: string;
  position: string;
  count: number;
  no: number;
  name: string;
  present: boolean;
  startTardy?: string;
  lateLunch?: string;
  needsAdj?: boolean;
  overtime?: string;
  otCorp?: string;
  subAnother?: boolean;
  incident?: string;
  notes?: string;
}

interface DoctorRow {
  date: string;
  filledBy: string;
  checkedBy: string;
  position: string;
  count: number;
  no: number;
  name: string;
  present: boolean;
  checkIn?: string;
  lunchOut?: string;
  lunchIn?: string;
  checkOut?: string;
}

export default function AttendanceTrack() {
  const [loading, setLoading] = useState(false);
  const [staffList, setStaffList] = useState<{ staff: StaffList; doctors: DoctorMember[] } | null>(null);
  const [attendanceData, setAttendanceData] = useState<{ staffData: AttendanceRow[]; doctorData: DoctorRow[] }>({
    staffData: [],
    doctorData: []
  });

  // 마지막 저장된 데이터 추적 (자동 저장 최적화용)
  const lastSavedDataRef = useRef<string>('');
  // 초기 로드 완료 플래그 (초기 로드 시 자동 저장 방지)
  const isInitialLoadRef = useRef<boolean>(true);
  // 🔒 보안: Rate limiting을 위한 ref
  const lastAutoSaveTimeRef = useRef<number>(0);
  const lastApiCallTimeRef = useRef<number>(0);
  const autoSaveCountRef = useRef<number>(0);
  const autoSaveResetTimeRef = useRef<number>(Date.now());

  // 🔒 보안: 입력 검증 및 XSS 방지 함수
  const validateInput = (value: string | undefined | null, maxLength: number = 500): string => {
    if (!value || typeof value !== 'string') return '';
    // XSS 방지: 위험한 문자 제거
    let sanitized = value
      .replace(/[<>\"']/g, '') // HTML 태그 및 따옴표 제거
      .replace(/javascript:/gi, '') // javascript: 프로토콜 제거
      .replace(/on\w+=/gi, ''); // 이벤트 핸들러 제거 (onclick, onerror 등)
    // 길이 제한
    if (sanitized.length > maxLength) {
      sanitized = sanitized.substring(0, maxLength);
    }
    return sanitized;
  };

  // 🔒 보안: 날짜 형식 검증
  const validateDate = (date: string): boolean => {
    const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
    if (!dateRegex.test(date)) return false;
    const parsedDate = new Date(date);
    return !isNaN(parsedDate.getTime());
  };

  // 🔒 보안: 오피스 값 검증
  const validateOffice = (office: string): boolean => {
    const allowedOffices = ['Ming', 'Bernard', 'Delano', 'Tulare', 'Visalia', 'Fresno', 'California', 'Ortho'];
    return allowedOffices.includes(office);
  };

  // 🔒 보안: Position 값 검증
  const validatePosition = (position: string): boolean => {
    const allowedPositions = ['Front Office', 'Biller', 'Dental Assistant', 'RDA', 'Sub', 'Extern'];
    return allowedPositions.includes(position);
  };

  // 🔒 보안: Incident 값 검증
  const validateIncident = (incident: string): boolean => {
    const allowedIncidents = ['', 'Late In', 'Early Out', 'Long Lunch', 'Leave and Come Back', 'Voluntary Early Out'];
    return allowedIncidents.includes(incident);
  };

  // 날짜 상태 (캘리포니아 시간대)
  const [trackDate, setTrackDate] = useState(() => {
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
  });

  const [selectedOffice, setSelectedOffice] = useState(''); // 오피스 선택 필요
  const [officePasswordVerified, setOfficePasswordVerified] = useState(false); // 오피스 비밀번호 확인 상태
  const [filledBy, setFilledBy] = useState('');
  const [checkedBy, setCheckedBy] = useState('');

  // 테이블 데이터 상태 (staff-list 기반)
  const [tableRows, setTableRows] = useState<Array<{
    id: string;
    position: string;
    no: number;
    name: string;
    present: boolean;
    startTardy: string;
    lateLunch: string;
    needsAdj: boolean;
    overtime: string;
    otCorp: string;
    subAnother: boolean;
    incident: string;
    notes: string;
  }>>([]);

  // Doctor 테이블 데이터 상태
  const [doctorRows, setDoctorRows] = useState<Array<{
    id: string;
    no: number;
    name: string;
    present: boolean;
    checkIn: string;
    lunchOut: string;
    lunchIn: string;
    checkOut: string;
  }>>([]);

  // Office 옵션
  const officeOptions = ['Ming', 'Bernard', 'Delano', 'Tulare', 'Visalia', 'Fresno', 'California', 'Ortho'];
  const incidentOptions = ['', 'Late In', 'Early Out', 'Long Lunch', 'Leave and Come Back', 'Voluntary Early Out'];

  // 오피스 선택 핸들러 (비밀번호 확인 포함)
  const handleOfficeChange = (newOffice: string) => {
    // 빈 값으로 선택하면 비밀번호 없이 변경 허용 (초기화)
    if (newOffice === '') {
      setSelectedOffice('');
      setOfficePasswordVerified(false);
      return;
    }
    
    // 🔒 보안: 오피스 값 검증
    if (!validateOffice(newOffice)) {
      alert('Invalid office value');
      return;
    }
    
    // 선택된 office의 첫 알파벳 대문자를 비밀번호로 사용
    const officePassword = newOffice.charAt(0).toUpperCase();
    const password = prompt(`Enter password to access office ${newOffice}:`);
    if (password === null) return; // 사용자가 취소한 경우
    if (password !== officePassword) {
      alert("Incorrect password. Office access denied.");
      return;
    }
    setSelectedOffice(newOffice);
    setOfficePasswordVerified(true);
  };

  // Staff List를 기반으로 테이블 행 업데이트 (기존 입력 데이터 보존)
  const updateTableRowsFromStaffList = useCallback((staffListData: { staff: StaffList; doctors: DoctorMember[] }, preserveExistingData: boolean = true) => {
    setTableRows(prevRows => {
      
      // 타입 정의
      type TableRow = {
        id: string;
        position: string;
        no: number;
        name: string;
        present: boolean;
        startTardy: string;
        lateLunch: string;
        needsAdj: boolean;
        overtime: string;
        otCorp: string;
        subAnother: boolean;
        incident: string;
        notes: string;
      };
      
      // 임시 row는 preserveExistingData가 true일 때만 보존
      const tempRows = preserveExistingData ? prevRows.filter(row => row.id.startsWith('temp-')) : [];
      
      // 저장된 출석 데이터를 가져오기 (attendanceData에서)
      const savedRowsMap = new Map<string, TableRow>();
      if (attendanceData.staffData && attendanceData.staffData.length > 0) {
        attendanceData.staffData.forEach((row: AttendanceRow) => {
          const id = `${row.position}-${row.no}`;
          savedRowsMap.set(id, {
            id,
            position: row.position,
            no: row.no,
            name: row.name,
            present: row.present || false,
            startTardy: row.startTardy || '',
            lateLunch: row.lateLunch || '',
            needsAdj: row.needsAdj || false,
            overtime: row.overtime || '',
            otCorp: row.otCorp || '',
            subAnother: row.subAnother || false,
            incident: row.incident || '',
            notes: row.notes || ''
          });
        });
      }
      
      // staff-list를 기반으로 테이블 행 생성
      const newRows: TableRow[] = [];
      
      // Position 순서 정의
      const positionOrder = ['Front Office', 'Biller', 'Dental Assistant', 'RDA', 'Sub', 'Extern'];
      
      // 기존 테이블 행 데이터를 맵으로 변환 (입력 데이터 보존용)
      const existingRowsMap = new Map<string, TableRow>();
      if (preserveExistingData) {
        prevRows.forEach(row => {
          existingRowsMap.set(row.id, row);
        });
      }
      
      // 저장된 출석 데이터도 맵에 병합 (저장된 데이터가 우선)
      savedRowsMap.forEach((savedRow, id) => {
        if (existingRowsMap.has(id)) {
          // 기존 행과 저장된 데이터 병합 (저장된 데이터 우선)
          existingRowsMap.set(id, { ...existingRowsMap.get(id)!, ...savedRow });
        }
      });
      
      positionOrder.forEach(position => {
        const members = staffListData.staff[position] || [];
        
        // Active가 true인 것만, 또는 Sub/Extern은 항상 포함
        const activeMembers = members.filter(m => {
          // Sub/Extern은 항상 포함
          if (position === 'Sub' || position === 'Extern') {
            return true;
          }
          // 다른 포지션은 active가 명시적으로 true인 것만 포함
          return m.active === true;
        });
        
        activeMembers.forEach(member => {
          const rowId = `${position}-${member.no}`;
          const existingRow = existingRowsMap.get(rowId);
          
          // 기존 행이 있으면 입력 데이터를 보존, 저장된 데이터가 있으면 병합
          const savedRow = savedRowsMap.get(rowId);
          
          if (existingRow) {
            // 기존 행과 저장된 데이터 병합 (저장된 데이터가 우선)
            const mergedRow = savedRow 
              ? { ...existingRow, ...savedRow, name: member.name || existingRow.name || savedRow.name || '' }
              : { ...existingRow, name: member.name || existingRow.name || '' };
            newRows.push(mergedRow);
          } else if (savedRow) {
            // 저장된 데이터가 있으면 그것을 사용, staff-list의 이름으로 업데이트
            newRows.push({
              ...savedRow,
              name: member.name || savedRow.name || '',
              position, // position은 확실히 설정
              no: member.no // no도 확실히 설정
            });
          } else {
            newRows.push({
              id: rowId,
              position,
              no: member.no,
              name: member.name || '',
              present: false,
              startTardy: '',
              lateLunch: '',
              needsAdj: false,
              overtime: '',
              otCorp: '',
              subAnother: false,
              incident: '',
              notes: ''
            });
          }
        });
      });
      
      // staff-list에서 제거된 행은 newRows에 포함되지 않으므로 자동으로 제거됨
      // 하지만 임시 row는 보존
      tempRows.forEach(tempRow => {
        // 이미 포함되어 있지 않으면 추가
        if (!newRows.find(r => r.id === tempRow.id)) {
          newRows.push(tempRow);
        }
      });
      
      // position별로 정렬 (임시 row는 각 position의 마지막에 위치)
      const sortedRows = positionOrder.flatMap(position => {
        const positionRows = newRows.filter(r => r.position === position);
        const regularRows = positionRows.filter(r => !r.id.startsWith('temp-'));
        const tempRowsForPosition = positionRows.filter(r => r.id.startsWith('temp-'));
        return [...regularRows, ...tempRowsForPosition];
      });
      
      return sortedRows;
    });
  }, [attendanceData]);

  // 임시 row 추가 (특정 position에)
  const addTempRow = (position: string) => {
    // 🔒 보안: Position 값 검증
    if (!validatePosition(position)) {
      return;
    }
    const positionRows = tableRows.filter(r => r.position === position);
    const tempRows = positionRows.filter(r => r.id.startsWith('temp-'));
    const tempNo = 1000 + tempRows.length; // 임시 row는 1000부터 시작
    
    const newTempRow = {
      id: `temp-${position}-${Date.now()}`,
      position,
      no: tempNo,
      name: '',
      present: false,
      startTardy: '',
      lateLunch: '',
      needsAdj: false,
      overtime: '',
      otCorp: '',
      subAnother: false,
      incident: '',
      notes: ''
    };
    
    // 해당 position의 마지막에 추가
    const positionIndex = tableRows.findIndex(r => r.position === position);
    if (positionIndex === -1) {
      // position이 없으면 맨 뒤에 추가
      setTableRows([...tableRows, newTempRow]);
    } else {
      // position의 마지막 row 다음에 추가
      let insertIndex = positionIndex;
      for (let i = positionIndex; i < tableRows.length; i++) {
        if (tableRows[i].position === position) {
          insertIndex = i + 1;
        } else {
          break;
        }
      }
      const newRows = [...tableRows];
      newRows.splice(insertIndex, 0, newTempRow);
      setTableRows(newRows);
    }
  };

  // Doctor 임시 row 추가
  const addDoctor = () => {
    const tempDoctors = doctorRows.filter(r => r.id.startsWith('temp-'));
    const tempNo = 1000 + tempDoctors.length; // 임시 row는 1000부터 시작
    
    const newDoctor = {
      id: `temp-doctor-${Date.now()}`,
      no: tempNo,
      name: '',
      present: false,
      checkIn: '',
      lunchOut: '',
      lunchIn: '',
      checkOut: ''
    };
    
    setDoctorRows([...doctorRows, newDoctor]);
  };

  // Doctor 테이블 행 업데이트
  const updateDoctorRowsFromStaffList = useCallback((doctors: DoctorMember[], preserveExistingData: boolean = true) => {
    setDoctorRows(prevRows => {
      // 임시 row는 preserveExistingData가 true일 때만 보존
      const tempRows = preserveExistingData ? prevRows.filter(row => row.id.startsWith('temp-')) : [];
      
      // 저장된 출석 데이터에서 doctor 데이터 가져오기
      const savedDoctorsMap = new Map<string, {
        id: string;
        no: number;
        name: string;
        present: boolean;
        checkIn: string;
        lunchOut: string;
        lunchIn: string;
        checkOut: string;
      }>();
      
      if (attendanceData.doctorData && attendanceData.doctorData.length > 0) {
        attendanceData.doctorData.forEach((row: DoctorRow) => {
          const id = `doctor-${row.no}`;
          savedDoctorsMap.set(id, {
            id,
            no: row.no,
            name: row.name,
            present: row.present || false,
            checkIn: row.checkIn || '',
            lunchOut: row.lunchOut || '',
            lunchIn: row.lunchIn || '',
            checkOut: row.checkOut || ''
          });
        });
      }
      
      // 기존 행 데이터 맵 (임시 row 제외)
      const existingRowsMap = new Map<string, typeof savedDoctorsMap extends Map<string, infer V> ? V : never>();
      prevRows.forEach(row => {
        if (!row.id.startsWith('temp-')) {
          existingRowsMap.set(row.id, row);
        }
      });
      
      const newRows = doctors.map(doctor => {
        const rowId = `doctor-${doctor.no}`;
        const existingRow = existingRowsMap.get(rowId);
        const savedRow = savedDoctorsMap.get(rowId);
        
        if (existingRow) {
          // 기존 행과 저장된 데이터 병합
          return savedRow 
            ? { ...existingRow, ...savedRow, name: doctor.name || existingRow.name || savedRow.name || '' }
            : { ...existingRow, name: doctor.name || existingRow.name || '' };
        } else if (savedRow) {
          // 저장된 데이터 사용
          return {
            ...savedRow,
            name: doctor.name || savedRow.name || ''
          };
        } else {
          // 새 행 생성
          return {
            id: rowId,
            no: doctor.no,
            name: doctor.name || '',
            present: false,
            checkIn: '',
            lunchOut: '',
            lunchIn: '',
            checkOut: ''
          };
        }
      });
      
      // 임시 row 추가 (preserveExistingData가 true일 때만)
      tempRows.forEach(tempRow => {
        if (!newRows.find(r => r.id === tempRow.id)) {
          newRows.push(tempRow);
        }
      });
      
      return newRows;
    });
  }, [attendanceData]);

  // 출석 데이터 불러오기
  const loadAttendanceData = useCallback(async (date: string) => {
    if (!date || !selectedOffice) {
      // 오피스가 선택되지 않았으면 빈 데이터로 초기화
      setAttendanceData({ staffData: [], doctorData: [] });
      setFilledBy('');
      setCheckedBy('');
      return;
    }
    
    // 🔒 보안: 날짜 형식 검증
    if (!validateDate(date)) {
      console.error('Invalid date format:', date);
      return;
    }

    // 🔒 보안: 오피스 값 검증
    if (!validateOffice(selectedOffice)) {
      console.error('Invalid office value:', selectedOffice);
      return;
    }
    
    try {
      setLoading(true);
      // 🔒 보안: 문서 ID 검증 (특수문자 제거)
      const safeDocId = `${date}_${selectedOffice}`.replace(/[^a-zA-Z0-9_-]/g, '');
      const docRef = doc(db, 'attendance-data', safeDocId);
      const docSnap = await getDoc(docRef);
      
      if (docSnap.exists()) {
        const docData = docSnap.data();
        setAttendanceData({
          staffData: docData.staffData || [],
          doctorData: docData.doctorData || []
        });
        setFilledBy(docData.filledBy || '');
        setCheckedBy(docData.checkedBy || '');
        
        // Doctor 데이터가 있으면 doctorRows 업데이트
        if (docData.doctorData && Array.isArray(docData.doctorData) && docData.doctorData.length > 0) {
          const doctorRowsData = docData.doctorData.map((row: DoctorRow) => ({
            id: `doctor-${row.no}`,
            no: row.no,
            name: row.name,
            present: row.present || false,
            checkIn: row.checkIn || '',
            lunchOut: row.lunchOut || '',
            lunchIn: row.lunchIn || '',
            checkOut: row.checkOut || ''
          }));
          setDoctorRows(doctorRowsData);
        }
        
        // 저장된 데이터를 임시로 저장 (staff-list가 로드된 후에 적용)
        if (docData.staffData && Array.isArray(docData.staffData)) {
          const savedRows = docData.staffData.map((row: AttendanceRow) => ({
            id: `${row.position}-${row.no}`,
            position: row.position,
            no: row.no,
            name: row.name,
            present: row.present || false,
            startTardy: row.startTardy || '',
            lateLunch: row.lateLunch || '',
            needsAdj: row.needsAdj || false,
            overtime: row.overtime || '',
            otCorp: row.otCorp || '',
            subAnother: row.subAnother || false,
            incident: row.incident || '',
            notes: row.notes || ''
          }));
          
          // staff-list 기반 테이블 행이 생성된 후에 저장된 데이터 적용
          // updateTableRowsFromStaffList에서 처리하도록 수정
          setTableRows(prevRows => {
            if (prevRows.length === 0) {
              // 테이블 행이 아직 없으면 staff-list가 먼저 로드되어야 함
              return prevRows;
            }
            
            // 기존 테이블 행과 병합 (위치는 유지, 값만 업데이트)
            const mergedRows = prevRows.map(prevRow => {
              const savedRow = savedRows.find(sr => sr.id === prevRow.id);
              if (savedRow) {
                return { ...prevRow, ...savedRow };
              }
              return prevRow;
            });
            
            return mergedRows;
          });
        }
      } else {
        // 새 데이터 초기화
        setAttendanceData({ staffData: [], doctorData: [] });
        setFilledBy('');
        setCheckedBy('');
        // 마지막 저장된 데이터도 초기화
        lastSavedDataRef.current = '';
      }
    } catch (error) {
      console.error('Error loading attendance data:', error);
      // 🔒 보안: 상세한 에러 메시지 노출 최소화
      alert('Error loading attendance data. Please try again.');
    } finally {
      setLoading(false);
    }
  }, [selectedOffice]);

  // 출석 데이터 저장
  const saveAttendanceData = useCallback(async (silent: boolean = false) => {
    // 🔒 보안: Rate limiting - 자동 저장은 최소 2초 간격, 분당 최대 30회
    const now = Date.now();
    if (silent) {
      // 자동 저장의 경우 rate limiting 적용
      if (now - lastAutoSaveTimeRef.current < 2000) {
        return; // 최소 2초 간격
      }
      // 분당 30회 제한
      if (now - autoSaveResetTimeRef.current > 60000) {
        autoSaveCountRef.current = 0;
        autoSaveResetTimeRef.current = now;
      }
      if (autoSaveCountRef.current >= 30) {
        return; // 분당 30회 초과
      }
      autoSaveCountRef.current++;
      lastAutoSaveTimeRef.current = now;
    }

    if (!trackDate) {
      if (!silent) alert('Please select a date');
      return;
    }

    // 🔒 보안: 오피스 선택 확인
    if (!selectedOffice) {
      if (!silent) alert('Please select an office');
      return;
    }

    // 🔒 보안: 날짜 형식 검증
    if (!validateDate(trackDate)) {
      if (!silent) alert('Invalid date format');
      return;
    }

    // 🔒 보안: 오피스 값 검증
    if (!validateOffice(selectedOffice)) {
      if (!silent) alert('Invalid office value');
      return;
    }

    try {
      // 자동 저장 시에는 loading 상태를 변경하지 않음 (깜빡임 방지)
      if (!silent) {
        setLoading(true);
      }
      
      // 🔒 보안: 입력 검증 및 정리
      const safeFilledBy = validateInput(filledBy, 100);
      const safeCheckedBy = validateInput(checkedBy, 100);
      
      // 현재 테이블에서 데이터 추출 (입력 검증 포함)
      const staffData: AttendanceRow[] = tableRows.map(row => {
        const validatedPosition = validatePosition(row.position) ? row.position : '';
        const validatedIncident = validateIncident(row.incident || '') ? (row.incident || '') : '';
        return {
          date: trackDate,
          filledBy: safeFilledBy,
          checkedBy: safeCheckedBy,
          position: validatedPosition,
          count: 0, // position별 present 개수는 저장 시 계산
          no: typeof row.no === 'number' && row.no >= 0 && row.no <= 9999 ? row.no : 0,
          name: validateInput(row.name, 100),
          present: typeof row.present === 'boolean' ? row.present : false,
          startTardy: validateInput(row.startTardy, 50),
          lateLunch: validateInput(row.lateLunch, 50),
          needsAdj: typeof row.needsAdj === 'boolean' ? row.needsAdj : false,
          overtime: validateInput(row.overtime, 50),
          otCorp: validateInput(row.otCorp, 50),
          subAnother: typeof row.subAnother === 'boolean' ? row.subAnother : false,
          incident: validatedIncident,
          notes: validateInput(row.notes, 500)
        };
      });

      // position별 present 개수 계산
      const positionCounts: { [key: string]: number } = {};
      staffData.forEach(row => {
        if (!positionCounts[row.position]) positionCounts[row.position] = 0;
        if (row.present) positionCounts[row.position]++;
      });
      staffData.forEach(row => {
        row.count = positionCounts[row.position] || 0;
      });

      // Doctor 데이터 추출 (입력 검증 포함)
      const doctorData: DoctorRow[] = doctorRows.map(row => ({
        date: trackDate,
        filledBy: safeFilledBy,
        checkedBy: safeCheckedBy,
        position: 'Doctor',
        count: doctorRows.filter(r => r.present).length,
        no: typeof row.no === 'number' ? row.no : 0,
        name: validateInput(row.name, 100),
        present: typeof row.present === 'boolean' ? row.present : false,
        checkIn: validateInput(row.checkIn, 10),
        lunchOut: validateInput(row.lunchOut, 10),
        lunchIn: validateInput(row.lunchIn, 10),
        checkOut: validateInput(row.checkOut, 10)
      }));

      const data = {
        date: trackDate,
        office: selectedOffice,
        filledBy: safeFilledBy,
        checkedBy: safeCheckedBy,
        staffData,
        doctorData,
        createdAt: Timestamp.now()
      };

      // 🔒 보안: 문서 ID 검증 (특수문자 제거)
      const safeDocId = `${trackDate}_${selectedOffice}`.replace(/[^a-zA-Z0-9_-]/g, '');
      await setDoc(doc(db, 'attendance-data', safeDocId), data);
      
      // 마지막 저장된 데이터 업데이트 (자동 저장 최적화용)
      lastSavedDataRef.current = JSON.stringify({ tableRows, doctorRows, filledBy, checkedBy });
      
    } catch (error) {
      console.error('Error saving attendance data:', error);
      if (!silent) {
        // 🔒 보안: 상세한 에러 메시지 노출 최소화
        alert('Error saving attendance data. Please try again.');
      }
    } finally {
      if (!silent) {
        setLoading(false);
      }
    }
  }, [trackDate, filledBy, checkedBy, tableRows, doctorRows, selectedOffice]);

  // Submit (데이터 저장 + PDF 생성)
  const handleSubmit = useCallback(async () => {
    if (!trackDate) {
      alert('Please select a date');
      return;
    }

    // 🔒 보안: 오피스 선택 확인
    if (!selectedOffice) {
      alert('Please select an office');
      return;
    }

    // 🔒 보안: 날짜 형식 검증
    if (!validateDate(trackDate)) {
      alert('Invalid date format');
      return;
    }

    // 🔒 보안: 오피스 값 검증
    if (!validateOffice(selectedOffice)) {
      alert('Invalid office value');
      return;
    }

    try {
      setLoading(true);
      
      // 🔒 보안: 입력 검증 및 정리
      const safeFilledBy = validateInput(filledBy, 100);
      const safeCheckedBy = validateInput(checkedBy, 100);
      
      // 🔒 보안: Rate limiting - API 호출은 최소 5초 간격
      const now = Date.now();
      if (now - lastApiCallTimeRef.current < 5000) {
        alert('Please wait a moment before submitting again.');
        setLoading(false);
        return;
      }
      lastApiCallTimeRef.current = now;

      // 1. 현재 테이블 데이터를 Firestore에 저장 (입력 검증 포함)
      const staffData: AttendanceRow[] = tableRows.map(row => {
        const validatedPosition = validatePosition(row.position) ? row.position : '';
        const validatedIncident = validateIncident(row.incident || '') ? (row.incident || '') : '';
        return {
          date: trackDate,
          filledBy: safeFilledBy,
          checkedBy: safeCheckedBy,
          position: validatedPosition,
          count: 0,
          no: typeof row.no === 'number' && row.no >= 0 && row.no <= 9999 ? row.no : 0,
          name: validateInput(row.name, 100),
          present: typeof row.present === 'boolean' ? row.present : false,
          startTardy: validateInput(row.startTardy, 50),
          lateLunch: validateInput(row.lateLunch, 50),
          needsAdj: typeof row.needsAdj === 'boolean' ? row.needsAdj : false,
          overtime: validateInput(row.overtime, 50),
          otCorp: validateInput(row.otCorp, 50),
          subAnother: typeof row.subAnother === 'boolean' ? row.subAnother : false,
          incident: validatedIncident,
          notes: validateInput(row.notes, 500)
        };
      });

      // position별 present 개수 계산
      const positionCounts: { [key: string]: number } = {};
      staffData.forEach(row => {
        if (!positionCounts[row.position]) positionCounts[row.position] = 0;
        if (row.present) positionCounts[row.position]++;
      });
      staffData.forEach(row => {
        row.count = positionCounts[row.position] || 0;
      });

      // Doctor 데이터 추출 (입력 검증 포함)
      const doctorData: DoctorRow[] = doctorRows.map(row => ({
        date: trackDate,
        filledBy: safeFilledBy,
        checkedBy: safeCheckedBy,
        position: 'Doctor',
        count: doctorRows.filter(r => r.present).length,
        no: typeof row.no === 'number' ? row.no : 0,
        name: validateInput(row.name, 100),
        present: typeof row.present === 'boolean' ? row.present : false,
        checkIn: validateInput(row.checkIn, 10),
        lunchOut: validateInput(row.lunchOut, 10),
        lunchIn: validateInput(row.lunchIn, 10),
        checkOut: validateInput(row.checkOut, 10)
      }));

      // 🔒 보안: 문서 ID 검증 (특수문자 제거)
      const safeDocId = `${trackDate}_${selectedOffice}`.replace(/[^a-zA-Z0-9_-]/g, '');
      await setDoc(doc(db, 'attendance-data', safeDocId), {
        date: trackDate,
        office: selectedOffice,
        filledBy: safeFilledBy,
        checkedBy: safeCheckedBy,
        staffData,
        doctorData,
        createdAt: Timestamp.now()
      });

      // 2. PDF 생성
      const response = await fetch('/api/generate-attendance-pdf', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          date: trackDate,
          office: selectedOffice,
          filledBy: safeFilledBy,
          checkedBy: safeCheckedBy,
          staffData,
          doctorData
        }),
      });

      if (response.ok) {
        // PDF blob 받기
        const blob = await response.blob();
        
        if (!blob || blob.size === 0) {
          throw new Error('PDF blob is empty');
        }
        
        // PDF를 Firebase Storage에 저장
        try {
          const storage = getStorage();
          
          // 🔒 보안: 파일명 및 경로 검증 (특수문자 제거)
          const safeDate = trackDate.replace(/[^a-zA-Z0-9_-]/g, '');
          const safeOffice = selectedOffice.replace(/[^a-zA-Z0-9_-]/g, '');
          const filename = `3) ${safeDate}_${safeOffice}_Attendance Tract.pdf`;
          const safeStoragePath = `endofday-pdfs/${safeOffice}/${safeDate}/${filename}`.replace(/[^a-zA-Z0-9_/.-]/g, '');
          const storageRef = ref(storage, safeStoragePath);
          
          // PDF 업로드
          await uploadBytes(storageRef, blob);
          
          // 다운로드 URL 가져오기
          const downloadUrl = await getDownloadURL(storageRef);
          
          // 🔒 보안: 문서 ID 검증
          const safePdfDocId = `${safeDate}_${safeOffice}_attendance-tract_${Date.now()}`.replace(/[^a-zA-Z0-9_-]/g, '');
          await setDoc(doc(db, 'pdf-documents', safePdfDocId), {
            filename,
            office: selectedOffice,
            date: trackDate,
            type: 'Attendance Tract',
            url: downloadUrl,
            storagePath: safeStoragePath,
            createdAt: Timestamp.now(),
          });
          
        } catch (storageError: any) {
          console.error('Storage error:', storageError);
          const errorMsg = storageError?.message || '알 수 없는 오류';
          throw new Error(`PDF 저장 중 오류가 발생했습니다: ${errorMsg}`);
        }
      } else {
        // 에러 처리
        let errorMessage = 'PDF generation failed';
        try {
          const errorData = await response.json();
          errorMessage = errorData.error || errorMessage;
        } catch {
          try {
            const errorText = await response.text();
            errorMessage = errorText || errorMessage;
          } catch (e) {
            errorMessage = `Status: ${response.status}`;
          }
        }
        console.error('PDF generation failed:', response.status, errorMessage);
        throw new Error(`PDF generation failed: ${response.status} - ${errorMessage}`);
      }

      alert('PDF generated and saved successfully to End of Day!');
      
      // 3. database에서 attendance-data 삭제
      // 🔒 보안: 문서 ID 검증 (위에서 선언한 safeDocId 재사용)
      await deleteDoc(doc(db, 'attendance-data', safeDocId));
      
      // 폼 리셋
      setAttendanceData({ staffData: [], doctorData: [] });
      setFilledBy('');
      setCheckedBy('');
      
      // Office 및 비밀번호 상태 초기화
      setSelectedOffice('');
      setOfficePasswordVerified(false);
      
      // 테이블 데이터 리셋 (staff-list 기반으로 다시 초기화)
      // 자동 저장이 트리거되지 않도록 먼저 플래그 설정
      isInitialLoadRef.current = true;
      lastSavedDataRef.current = '';
      
      // doctorRows와 tableRows를 명시적으로 빈 배열로 초기화
      setDoctorRows([]);
      setTableRows([]);
      
      // staffList가 있으면 나중에 다시 로드될 때 자동으로 채워질 것임
      // 여기서는 명시적으로 빈 배열로 초기화만 함
      
      // 리셋 완료 후 자동 저장 방지를 위한 처리
      // isInitialLoadRef가 true로 설정되어 있으므로 자동 저장 useEffect가 실행되지 않음
      // 자동 저장 useEffect의 초기 로드 로직이 1초 후에 현재 상태를 lastSavedDataRef에 저장함
      // 따라서 별도로 lastSavedDataRef를 업데이트할 필요 없음
    } catch (error) {
      console.error('Error submitting:', error);
      // 🔒 보안: 상세한 에러 메시지 노출 최소화
      alert('Error submitting. Please try again.');
    } finally {
      setLoading(false);
    }
  }, [trackDate, filledBy, checkedBy, tableRows, doctorRows, selectedOffice, staffList, updateTableRowsFromStaffList, updateDoctorRowsFromStaffList]);

  // Staff List 실시간 감지 및 테이블 업데이트
  useEffect(() => {
    // 🔒 보안: 오피스가 선택되지 않았거나 비밀번호가 확인되지 않았으면 리스너 설정하지 않음
    if (!selectedOffice || !officePasswordVerified || !validateOffice(selectedOffice)) {
      setStaffList({ staff: {}, doctors: [] });
      setTableRows([]);
      setDoctorRows([]);
      return;
    }

    const staffListDocRef = doc(db, 'staff-list', selectedOffice);
    
    const unsubscribe = onSnapshot(staffListDocRef, (docSnap) => {
      if (docSnap.exists()) {
        const data = docSnap.data();
        
        const staffListData = {
          staff: data.staff || {},
          doctors: data.doctors || []
        };
        
        setStaffList(staffListData);
        
        // 테이블 행 업데이트 (기존 입력 데이터 보존)
        updateTableRowsFromStaffList(staffListData, true);
        
        // Doctor 테이블 행 업데이트
        if (staffListData.doctors && staffListData.doctors.length > 0) {
          updateDoctorRowsFromStaffList(staffListData.doctors, true);
        }
      } else {
        // Staff List가 없으면 빈 데이터로 초기화
        setStaffList({
          staff: {},
          doctors: []
        });
        setTableRows([]);
      }
    }, (error) => {
      console.error('Error listening to staff list:', error);
      alert('Error loading staff list. Please try again.');
    });
    
    return () => {
      unsubscribe();
    };
  }, [updateTableRowsFromStaffList, selectedOffice, officePasswordVerified]);

  useEffect(() => {
    if (trackDate && officePasswordVerified) {
      // 오피스가 변경되면 초기 로드 플래그 리셋
      isInitialLoadRef.current = true;
      lastSavedDataRef.current = '';
      loadAttendanceData(trackDate);
    } else if (!officePasswordVerified) {
      // 비밀번호가 확인되지 않으면 데이터 초기화
      setAttendanceData({ staffData: [], doctorData: [] });
      setFilledBy('');
      setCheckedBy('');
      setTableRows([]);
      setDoctorRows([]);
    }
  }, [trackDate, selectedOffice, officePasswordVerified, loadAttendanceData]);

  // 자동 저장 (debounce 적용, 깜빡임 방지)
  useEffect(() => {
    // 오피스가 선택되지 않았거나 초기 로드 시에는 저장하지 않음
    if (!selectedOffice || (tableRows.length === 0 && doctorRows.length === 0)) return;
    
    // 초기 로드 시에는 자동 저장하지 않음 (데이터 로드 후 한 번만 설정)
    if (isInitialLoadRef.current) {
      // 초기 로드 완료 후 현재 데이터를 마지막 저장된 데이터로 설정
      setTimeout(() => {
        lastSavedDataRef.current = JSON.stringify({ tableRows, doctorRows, filledBy, checkedBy });
        isInitialLoadRef.current = false;
      }, 1000);
      return;
    }
    
    // 데이터가 실제로 변경되었는지 확인
    const currentData = JSON.stringify({ tableRows, doctorRows, filledBy, checkedBy });
    if (currentData === lastSavedDataRef.current) {
      return; // 변경사항이 없으면 저장하지 않음
    }
    
    // debounce: 3초 후에 저장 (깜빡임 최소화)
    const timer = setTimeout(() => {
      if (trackDate && selectedOffice) {
        // 자동 저장은 조용히 수행 (silent: true, loading 상태 변경 안 함)
        saveAttendanceData(true);
      }
    }, 3000);

    return () => clearTimeout(timer);
  }, [tableRows, doctorRows, filledBy, checkedBy, trackDate, selectedOffice, officePasswordVerified, saveAttendanceData]);

  return (
    <div style={{ padding: '20px', fontFamily: 'Arial, sans-serif' }}>
      <h1 style={{ fontSize: '3rem', fontWeight: 'bold', color: '#2D3748', marginBottom: '20px', textAlign: 'center' }}>
        Attendance Tract
      </h1>
      
      <div style={{ marginTop: '20px', marginBottom: '20px', fontSize: '15px', textAlign: 'center', color: '#666', fontStyle: 'italic', lineHeight: '1.5' }}>
        Management is required to review all staff members' times on Time Clock to fill out the Attendance Tract Sheet accurately and to request any necessary clock adjustments.
      </div>

      <div style={{ 
        marginBottom: '20px', 
        display: 'flex', 
        alignItems: 'center', 
        justifyContent: 'center',
        gap: '20px',
        flexWrap: 'wrap'
      }}>
        <label style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          <strong>Office:</strong>
          <select
            value={selectedOffice}
            onChange={(e) => handleOfficeChange(e.target.value)}
            style={{ padding: '5px', borderRadius: '4px', border: '1px solid #ddd' }}
          >
            <option value="">Select Office</option>
            {officeOptions.map(office => (
              <option key={office} value={office}>{office}</option>
            ))}
          </select>
        </label>
        
        <label style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          <strong>Date:</strong>
          <input
            type="date"
            value={trackDate}
            onChange={(e) => setTrackDate(e.target.value)}
            style={{ padding: '5px', borderRadius: '4px', border: '1px solid #ddd' }}
          />
        </label>
        
        <label style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          <strong>Filled Out By:</strong>
          <input
            type="text"
            value={filledBy}
            onChange={(e) => {
              // 🔒 보안: 입력 검증
              const validatedValue = validateInput(e.target.value, 100);
              setFilledBy(validatedValue);
            }}
            style={{ padding: '5px', borderRadius: '4px', border: '1px solid #ddd', minWidth: '150px' }}
          />
        </label>
        
        <label style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          <strong>Management that Checked Times Today on Time Clock:</strong>
          <input
            type="text"
            value={checkedBy}
            onChange={(e) => {
              // 🔒 보안: 입력 검증
              const validatedValue = validateInput(e.target.value, 100);
              setCheckedBy(validatedValue);
            }}
            style={{ padding: '5px', borderRadius: '4px', border: '1px solid #ddd', minWidth: '150px' }}
          />
        </label>
      </div>

      {loading && <div>Loading...</div>}

      {/* Staff 테이블 */}
      {staffList && (
        <div style={{ marginTop: '20px', overflowX: 'auto' }}>
          <div style={{ 
            background: '#f9fbfd',
            borderRadius: '12px',
            overflow: 'hidden',
            boxShadow: '0 2px 8px rgba(0,0,0,0.1)'
          }}>
            {/* 고정 헤더 테이블 */}
            <div style={{ position: 'sticky', top: 0, zIndex: 100, background: '#f9fbfd' }}>
              <table style={{ 
                width: '100%', 
                tableLayout: 'fixed',
                borderCollapse: 'separate',
                borderSpacing: 0,
                background: '#f9fbfd'
              }}>
                <colgroup>
                  <col style={{ width: '5%' }} />
                  <col style={{ width: '12%' }} />
                  <col style={{ width: '6%' }} />
                  <col style={{ width: '10%' }} />
                  <col style={{ width: '10%' }} />
                  <col style={{ width: '8%' }} />
                  <col style={{ width: '8%' }} />
                  <col style={{ width: '10%' }} />
                  <col style={{ width: '8%' }} />
                  <col style={{ width: '13%' }} />
                  <col style={{ width: '10%' }} />
                </colgroup>
                <thead>
                  <tr style={{ background: 'linear-gradient(90deg, #2D3748 0%, #4A5568 100%)', color: '#fff' }}>
                    <th style={{ padding: '14px 8px', border: '1px solid #ddd', fontSize: '1.15rem', fontWeight: '700', background: 'linear-gradient(90deg, #2D3748 0%, #4A5568 100%)' }}>No.</th>
                    <th style={{ padding: '14px 8px', border: '1px solid #ddd', fontSize: '1.15rem', fontWeight: '700', background: 'linear-gradient(90deg, #2D3748 0%, #4A5568 100%)' }}>Name</th>
                    <th style={{ padding: '14px 8px', border: '1px solid #ddd', fontSize: '1.15rem', fontWeight: '700', background: 'linear-gradient(90deg, #2D3748 0%, #4A5568 100%)' }}>Present</th>
                    <th style={{ padding: '14px 8px', border: '1px solid #ddd', fontSize: '1.15rem', fontWeight: '700', background: 'linear-gradient(90deg, #2D3748 0%, #4A5568 100%)' }}>Start Shift<br />Tardy (Min)</th>
                    <th style={{ padding: '14px 8px', border: '1px solid #ddd', fontSize: '1.15rem', fontWeight: '700', background: 'linear-gradient(90deg, #2D3748 0%, #4A5568 100%)' }}>Late from<br />Lunch (Min)</th>
                    <th style={{ padding: '14px 8px', border: '1px solid #ddd', fontSize: '1.15rem', fontWeight: '700', background: 'linear-gradient(90deg, #2D3748 0%, #4A5568 100%)' }}>Needs Clock Adj.</th>
                    <th style={{ padding: '14px 8px', border: '1px solid #ddd', fontSize: '1.15rem', fontWeight: '700', background: 'linear-gradient(90deg, #2D3748 0%, #4A5568 100%)' }}>Overtime</th>
                    <th style={{ padding: '14px 8px', border: '1px solid #ddd', fontSize: '1.15rem', fontWeight: '700', background: 'linear-gradient(90deg, #2D3748 0%, #4A5568 100%)' }}>OT Corp Authorized<br />By (Initials)</th>
                    <th style={{ padding: '14px 8px', border: '1px solid #ddd', fontSize: '1.15rem', fontWeight: '700', background: 'linear-gradient(90deg, #2D3748 0%, #4A5568 100%)' }}>Sub. at Another<br />Office</th>
                    <th style={{ padding: '14px 8px', border: '1px solid #ddd', fontSize: '1.15rem', fontWeight: '700', background: 'linear-gradient(90deg, #2D3748 0%, #4A5568 100%)' }}>Incident Description</th>
                    <th style={{ padding: '14px 8px', border: '1px solid #ddd', fontSize: '1.15rem', fontWeight: '700', background: 'linear-gradient(90deg, #2D3748 0%, #4A5568 100%)' }}>Notes</th>
                  </tr>
                </thead>
              </table>
            </div>
            {/* 스크롤 가능한 본문 */}
            <div style={{ maxHeight: '70vh', overflowY: 'auto' }}>
              <table style={{ 
                width: '100%', 
                tableLayout: 'fixed',
                borderCollapse: 'separate',
                borderSpacing: 0,
                background: '#f9fbfd'
              }}>
                <colgroup>
                  <col style={{ width: '5%' }} />
                  <col style={{ width: '12%' }} />
                  <col style={{ width: '6%' }} />
                  <col style={{ width: '10%' }} />
                  <col style={{ width: '10%' }} />
                  <col style={{ width: '8%' }} />
                  <col style={{ width: '8%' }} />
                  <col style={{ width: '10%' }} />
                  <col style={{ width: '8%' }} />
                  <col style={{ width: '13%' }} />
                  <col style={{ width: '10%' }} />
                </colgroup>
                <tbody>
              {(() => {
                const positionOrder = ['Front Office', 'Biller', 'Dental Assistant', 'RDA', 'Sub', 'Extern'];
                const rows: React.ReactElement[] = [];
                
                if (tableRows.length === 0) {
                  return (
                    <tr>
                      <td colSpan={11} style={{ padding: '20px', textAlign: 'center' }}>
                        No staff members found. Please add staff in Staff List Management.
                      </td>
                    </tr>
                  );
                }
                
                // Position별로 그룹화하여 헤더와 버튼 추가
                positionOrder.forEach(position => {
                  const positionRows = tableRows.filter(r => r.position === position);
                  if (positionRows.length === 0) return;
                  
                  const presentCount = positionRows.filter(r => r.present).length;
                  
                  // Position 헤더 행
                  rows.push(
                    <tr key={`header-${position}`} style={{ background: '#EDF2F7', fontWeight: 'bold' }}>
                      <td colSpan={11} style={{ padding: '10px', textAlign: 'center', border: '1px solid #ddd', position: 'relative' }}>
                        <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', position: 'relative' }}>
                          <span style={{ textAlign: 'center' }}>{position} <span style={{ marginLeft: '10px' }}>{presentCount}</span></span>
                          <button
                            onClick={() => addTempRow(position)}
                            style={{
                              padding: '4px 8px',
                              background: '#4CAF50',
                              color: 'white',
                              border: 'none',
                              borderRadius: '4px',
                              cursor: 'pointer',
                              fontSize: '12px',
                              fontWeight: 'normal',
                              position: 'absolute',
                              right: '10px'
                            }}
                          >
                            Add Staff
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                  
                  // Position별 행들 추가
                  positionRows.forEach(row => {
                  
                  // 데이터 행
                  rows.push(
                    <tr key={row.id} style={{ background: '#fff' }}>
                      <td style={{ padding: '10px 6px', border: '1px solid #ddd', textAlign: 'center' }}>{row.no}</td>
                      <td style={{ padding: '10px 6px', border: '1px solid #ddd', fontWeight: '600' }}>
                        {(row.position === 'Sub' || row.position === 'Extern' || row.id.startsWith('temp-')) ? (
                          <input
                            type="text"
                            value={row.name}
                            onChange={(e) => {
                              // 🔒 보안: 입력 검증
                              const validatedValue = validateInput(e.target.value, 100);
                              const newRows = [...tableRows];
                              const rowIndex = newRows.findIndex(r => r.id === row.id);
                              if (rowIndex !== -1) {
                                newRows[rowIndex].name = validatedValue;
                                setTableRows(newRows);
                              }
                            }}
                            style={{ width: '100%', border: 'none', background: 'transparent', padding: '0', fontWeight: '600' }}
                            placeholder="Enter name"
                          />
                        ) : (
                          row.name
                        )}
                      </td>
                      <td style={{ padding: '10px 6px', border: '1px solid #ddd', textAlign: 'center' }}>
                        <input
                          type="checkbox"
                          checked={row.present}
                          onChange={(e) => {
                            const newRows = [...tableRows];
                            const rowIndex = newRows.findIndex(r => r.id === row.id);
                            if (rowIndex !== -1) {
                              newRows[rowIndex].present = e.target.checked;
                              setTableRows(newRows);
                            }
                          }}
                          style={{ width: '20px', height: '20px', accentColor: '#2D3748' }}
                        />
                      </td>
                      <td style={{ padding: '10px 6px', border: '1px solid #ddd' }}>
                        <input
                          type="text"
                          value={row.startTardy}
                            onChange={(e) => {
                              // 🔒 보안: 입력 검증
                              const validatedValue = validateInput(e.target.value, 50);
                              const newRows = [...tableRows];
                              const rowIndex = newRows.findIndex(r => r.id === row.id);
                              if (rowIndex !== -1) {
                                newRows[rowIndex].startTardy = validatedValue;
                                setTableRows(newRows);
                              }
                            }}
                          style={{ width: '100%', border: 'none', background: 'transparent', padding: '0' }}
                        />
                      </td>
                      <td style={{ padding: '10px 6px', border: '1px solid #ddd' }}>
                        <input
                          type="text"
                          value={row.lateLunch}
                            onChange={(e) => {
                              // 🔒 보안: 입력 검증
                              const validatedValue = validateInput(e.target.value, 50);
                              const newRows = [...tableRows];
                              const rowIndex = newRows.findIndex(r => r.id === row.id);
                              if (rowIndex !== -1) {
                                newRows[rowIndex].lateLunch = validatedValue;
                                setTableRows(newRows);
                              }
                            }}
                          style={{ width: '100%', border: 'none', background: 'transparent', padding: '0' }}
                        />
                      </td>
                      <td style={{ padding: '10px 6px', border: '1px solid #ddd', textAlign: 'center' }}>
                        <input
                          type="checkbox"
                          checked={row.needsAdj}
                          onChange={(e) => {
                            const newRows = [...tableRows];
                            const rowIndex = newRows.findIndex(r => r.id === row.id);
                            if (rowIndex !== -1) {
                              newRows[rowIndex].needsAdj = e.target.checked;
                              setTableRows(newRows);
                            }
                          }}
                          style={{ width: '20px', height: '20px', accentColor: '#2D3748' }}
                        />
                      </td>
                      <td style={{ padding: '10px 6px', border: '1px solid #ddd' }}>
                        <input
                          type="text"
                          value={row.overtime}
                            onChange={(e) => {
                              // 🔒 보안: 입력 검증
                              const validatedValue = validateInput(e.target.value, 50);
                              const newRows = [...tableRows];
                              const rowIndex = newRows.findIndex(r => r.id === row.id);
                              if (rowIndex !== -1) {
                                newRows[rowIndex].overtime = validatedValue;
                                setTableRows(newRows);
                              }
                            }}
                          style={{ width: '100%', border: 'none', background: 'transparent', padding: '0' }}
                        />
                      </td>
                      <td style={{ padding: '10px 6px', border: '1px solid #ddd' }}>
                        <input
                          type="text"
                          value={row.otCorp}
                            onChange={(e) => {
                              // 🔒 보안: 입력 검증
                              const validatedValue = validateInput(e.target.value, 50);
                              const newRows = [...tableRows];
                              const rowIndex = newRows.findIndex(r => r.id === row.id);
                              if (rowIndex !== -1) {
                                newRows[rowIndex].otCorp = validatedValue;
                                setTableRows(newRows);
                              }
                            }}
                          style={{ width: '100%', border: 'none', background: 'transparent', padding: '0' }}
                        />
                      </td>
                      <td style={{ padding: '10px 6px', border: '1px solid #ddd', textAlign: 'center' }}>
                        <input
                          type="checkbox"
                          checked={row.subAnother}
                          onChange={(e) => {
                            const newRows = [...tableRows];
                            const rowIndex = newRows.findIndex(r => r.id === row.id);
                            if (rowIndex !== -1) {
                              newRows[rowIndex].subAnother = e.target.checked;
                              setTableRows(newRows);
                            }
                          }}
                          style={{ width: '20px', height: '20px', accentColor: '#2D3748' }}
                        />
                      </td>
                      <td style={{ padding: '10px 6px', border: '1px solid #ddd' }}>
                        <select
                          value={row.incident || ''}
                          onChange={(e) => {
                            // 🔒 보안: Incident 값 검증
                            const value = e.target.value;
                            if (!validateIncident(value)) {
                              return; // 유효하지 않은 값은 무시
                            }
                            const newRows = [...tableRows];
                            const rowIndex = newRows.findIndex(r => r.id === row.id);
                            if (rowIndex !== -1) {
                              newRows[rowIndex].incident = value;
                              setTableRows(newRows);
                            }
                          }}
                          style={{ width: '100%', border: 'none', background: 'transparent', padding: '2px', fontSize: '14px' }}
                        >
                          {incidentOptions.map(option => (
                            <option key={option} value={option}>
                              {option || 'Select...'}
                            </option>
                          ))}
                        </select>
                      </td>
                      <td style={{ padding: '10px 6px', border: '1px solid #ddd' }}>
                        <input
                          type="text"
                          value={row.notes}
                            onChange={(e) => {
                              // 🔒 보안: 입력 검증
                              const validatedValue = validateInput(e.target.value, 500);
                              const newRows = [...tableRows];
                              const rowIndex = newRows.findIndex(r => r.id === row.id);
                              if (rowIndex !== -1) {
                                newRows[rowIndex].notes = validatedValue;
                                setTableRows(newRows);
                              }
                            }}
                          style={{ width: '100%', border: 'none', background: 'transparent', padding: '0' }}
                        />
                      </td>
                    </tr>
                  );
                  });
                });
                
                return rows;
              })()}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* Doctor 테이블 */}
      {staffList && officePasswordVerified && staffList.doctors && staffList.doctors.length > 0 && (
        <div style={{ marginTop: '40px', overflowX: 'auto' }}>
          <table style={{ 
            width: '100%', 
            borderCollapse: 'collapse',
            background: '#f9fbfd',
            borderRadius: '12px',
            overflow: 'hidden',
            boxShadow: '0 2px 8px rgba(0,0,0,0.1)'
          }}>
            <thead>
              <tr style={{ background: 'linear-gradient(90deg, #2D3748 0%, #4A5568 100%)', color: '#fff' }}>
                <th style={{ padding: '14px 8px', border: '1px solid #ddd', fontSize: '1.15rem', fontWeight: '700', background: 'linear-gradient(90deg, #2D3748 0%, #4A5568 100%)' }}>No.</th>
                <th style={{ padding: '14px 8px', border: '1px solid #ddd', fontSize: '1.15rem', fontWeight: '700', background: 'linear-gradient(90deg, #2D3748 0%, #4A5568 100%)' }}>Name</th>
                <th style={{ padding: '14px 8px', border: '1px solid #ddd', fontSize: '1.15rem', fontWeight: '700', background: 'linear-gradient(90deg, #2D3748 0%, #4A5568 100%)' }}>Present</th>
                <th style={{ padding: '14px 8px', border: '1px solid #ddd', fontSize: '1.15rem', fontWeight: '700', background: 'linear-gradient(90deg, #2D3748 0%, #4A5568 100%)' }}>Check In</th>
                <th style={{ padding: '14px 8px', border: '1px solid #ddd', fontSize: '1.15rem', fontWeight: '700', background: 'linear-gradient(90deg, #2D3748 0%, #4A5568 100%)' }}>Lunch Out</th>
                <th style={{ padding: '14px 8px', border: '1px solid #ddd', fontSize: '1.15rem', fontWeight: '700', background: 'linear-gradient(90deg, #2D3748 0%, #4A5568 100%)' }}>Lunch In</th>
                <th style={{ padding: '14px 8px', border: '1px solid #ddd', fontSize: '1.15rem', fontWeight: '700', background: 'linear-gradient(90deg, #2D3748 0%, #4A5568 100%)' }}>Check Out</th>
              </tr>
              <tr style={{ background: '#EDF2F7', fontWeight: 'bold' }}>
                <td colSpan={7} style={{ padding: '10px', textAlign: 'center', border: '1px solid #ddd', position: 'relative' }}>
                  <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', position: 'relative' }}>
                    <span style={{ textAlign: 'center' }}>Doctor <span style={{ marginLeft: '10px' }}>{doctorRows.filter(r => r.present).length}</span></span>
                    <button
                      onClick={addDoctor}
                      style={{
                        padding: '4px 8px',
                        background: '#4CAF50',
                        color: 'white',
                        border: 'none',
                        borderRadius: '4px',
                        cursor: 'pointer',
                        fontSize: '12px',
                        fontWeight: 'normal',
                        position: 'absolute',
                        right: '10px'
                      }}
                    >
                      Add Doctor
                    </button>
                  </div>
                </td>
              </tr>
            </thead>
            <tbody>
              {doctorRows.map((row) => (
                <tr key={row.id} style={{ background: '#fff' }}>
                  <td style={{ padding: '10px 6px', border: '1px solid #ddd', textAlign: 'center' }}>{row.no}</td>
                  <td style={{ padding: '10px 6px', border: '1px solid #ddd', fontWeight: '600' }}>
                    {(row.id.startsWith('temp-')) ? (
                      <input
                        type="text"
                        value={row.name}
                        onChange={(e) => {
                          // 🔒 보안: 입력 검증 (길이 제한)
                          const validatedValue = validateInput(e.target.value, 100);
                          const newRows = [...doctorRows];
                          const rowIndex = newRows.findIndex(r => r.id === row.id);
                          if (rowIndex !== -1) {
                            newRows[rowIndex].name = validatedValue;
                            setDoctorRows(newRows);
                          }
                        }}
                        style={{ width: '100%', border: 'none', background: 'transparent', padding: '0', fontWeight: '600' }}
                        placeholder="Enter name"
                        maxLength={100}
                      />
                    ) : (
                      row.name
                    )}
                  </td>
                  <td style={{ padding: '10px 6px', border: '1px solid #ddd', textAlign: 'center' }}>
                    <input
                      type="checkbox"
                      checked={row.present}
                      onChange={(e) => {
                        const newRows = [...doctorRows];
                        const rowIndex = newRows.findIndex(r => r.id === row.id);
                        if (rowIndex !== -1) {
                          newRows[rowIndex].present = e.target.checked;
                          setDoctorRows(newRows);
                        }
                      }}
                      style={{ width: '20px', height: '20px', accentColor: '#1976D2' }}
                    />
                  </td>
                  <td style={{ padding: '10px 6px', border: '1px solid #ddd' }}>
                    <input
                      type="time"
                      value={row.checkIn}
                      onChange={(e) => {
                        const newRows = [...doctorRows];
                        const rowIndex = newRows.findIndex(r => r.id === row.id);
                        if (rowIndex !== -1) {
                          newRows[rowIndex].checkIn = e.target.value;
                          setDoctorRows(newRows);
                        }
                      }}
                      style={{ width: '100%', border: 'none', background: 'transparent', padding: '0' }}
                    />
                  </td>
                  <td style={{ padding: '10px 6px', border: '1px solid #ddd' }}>
                    <input
                      type="time"
                      value={row.lunchOut}
                      onChange={(e) => {
                        const newRows = [...doctorRows];
                        const rowIndex = newRows.findIndex(r => r.id === row.id);
                        if (rowIndex !== -1) {
                          newRows[rowIndex].lunchOut = e.target.value;
                          setDoctorRows(newRows);
                        }
                      }}
                      style={{ width: '100%', border: 'none', background: 'transparent', padding: '0' }}
                    />
                  </td>
                  <td style={{ padding: '10px 6px', border: '1px solid #ddd' }}>
                    <input
                      type="time"
                      value={row.lunchIn}
                      onChange={(e) => {
                        const newRows = [...doctorRows];
                        const rowIndex = newRows.findIndex(r => r.id === row.id);
                        if (rowIndex !== -1) {
                          newRows[rowIndex].lunchIn = e.target.value;
                          setDoctorRows(newRows);
                        }
                      }}
                      style={{ width: '100%', border: 'none', background: 'transparent', padding: '0' }}
                    />
                  </td>
                  <td style={{ padding: '10px 6px', border: '1px solid #ddd' }}>
                    <input
                      type="time"
                      value={row.checkOut}
                      onChange={(e) => {
                        const newRows = [...doctorRows];
                        const rowIndex = newRows.findIndex(r => r.id === row.id);
                        if (rowIndex !== -1) {
                          newRows[rowIndex].checkOut = e.target.value;
                          setDoctorRows(newRows);
                        }
                      }}
                      style={{ width: '100%', border: 'none', background: 'transparent', padding: '0' }}
                    />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <div style={{ marginTop: '20px', textAlign: 'center' }}>
        <button
          onClick={handleSubmit}
          disabled={loading}
          style={{ padding: '10px 20px', background: '#2D3748', color: 'white', border: 'none', borderRadius: '5px', cursor: 'pointer', fontSize: '16px', fontWeight: '600' }}
        >
          {loading ? 'Submitting...' : 'Submit'}
        </button>
      </div>
    </div>
  );
}

