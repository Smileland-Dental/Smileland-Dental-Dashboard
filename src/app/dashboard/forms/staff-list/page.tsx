'use client';

import React, { useState, useEffect, useCallback, useRef } from 'react';
import { db, auth } from '@/lib/firebase.config';
// Firebase 인증 직접 사용
import { onAuthStateChanged } from 'firebase/auth';
import { doc, getDoc, setDoc, Timestamp } from 'firebase/firestore';

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

interface StaffListData {
  staff: StaffList;
  doctors: DoctorMember[];
}

export default function StaffListManagement() {
  const [loading, setLoading] = useState(false);
  const [saveStatus, setSaveStatus] = useState('');
  const [selectedOffice, setSelectedOffice] = useState('');
  const [staffListData, setStaffListData] = useState<StaffListData>({
    staff: {},
    doctors: []
  });
  const [isAuthorized, setIsAuthorized] = useState<boolean | null>(null); // null: 확인 중, true: 인증됨, false: 인증 실패

  // Rate limiting을 위한 ref
  const lastUpdateStaffMemberCall = useRef<number>(0);
  const lastUpdateDoctorCall = useRef<number>(0);
  const lastSaveCall = useRef<number>(0);

  const officeOptions = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

  // 🔒 보안: 입력 검증 함수
  const validateInput = (value: string | undefined | null, maxLength: number = 100): string => {
    if (!value || typeof value !== 'string') return '';
    // 길이 제한
    if (value.length > maxLength) {
      return value.substring(0, maxLength);
    }
    return value;
  };

  // 🔒 보안: 오피스 값 검증
  const validateOffice = (office: string): boolean => {
    return officeOptions.includes(office);
  };

  // 🔒 보안: Position 값 검증
  const validatePosition = (position: string): boolean => {
    return positionOptions.includes(position);
  };

  const positionOptions = [
    'Front Office',
    'Biller',
    'Dental Assistant',
    'RDA',
    'Sub',
    'Extern'
  ];

  // Staff List 불러오기
  const loadStaffList = useCallback(async () => {
    // 오피스가 선택되지 않았으면 빈 데이터로 초기화
    if (!selectedOffice) {
      setStaffListData({ staff: {}, doctors: [] });
      return;
    }

    // 🔒 보안: 오피스 값 검증
    if (!validateOffice(selectedOffice)) {
      alert('Invalid office value');
      return;
    }

    try {
      setLoading(true);
      // 🔒 보안: 문서 ID 검증 (특수문자 제거)
      const safeOffice = selectedOffice.replace(/[^a-zA-Z0-9_-]/g, '');
      const staffListDoc = await getDoc(doc(db, 'staff-list', safeOffice));
      if (staffListDoc.exists()) {
        const data = staffListDoc.data();
        setStaffListData({
          staff: data.staff || {},
          doctors: data.doctors || []
        });
      } else {
        // 문서가 없으면 빈 데이터로 초기화
        setStaffListData({
          staff: {},
          doctors: []
        });
      }
    } catch (error) {
      alert('Error loading staff list: error');
    } finally {
      setLoading(false);
    }
  }, [selectedOffice]);

  // Staff List 저장
  const saveStaffList = useCallback(async () => {
    // Rate limiting: 최근 3초 내 호출 방지 (저장은 중요한 작업)
    const now = Date.now();
    if (now - lastSaveCall.current < 3000) {
      alert('⚠️ 저장은 3초에 한 번만 가능합니다. 잠시 후 다시 시도해주세요.');
      return;
    }
    lastSaveCall.current = now;

    // 🔒 보안: 오피스 선택 확인
    if (!selectedOffice) {
      alert('Please select an office');
      return;
    }

    // 🔒 보안: 오피스 값 검증
    if (!validateOffice(selectedOffice)) {
      alert('Invalid office value');
      return;
    }

    try {
      setLoading(true);
      setSaveStatus('Saving...');
      
      // 🔒 보안: 데이터 검증 및 정리
      const validatedStaff: StaffList = {};
      Object.keys(staffListData.staff).forEach(position => {
        if (validatePosition(position)) {
          validatedStaff[position] = (staffListData.staff[position] || []).map(member => ({
            no: typeof member.no === 'number' && member.no > 0 ? member.no : 1,
            name: validateInput(member.name, 100),
            active: typeof member.active === 'boolean' ? member.active : true
          }));
        }
      });

      const validatedDoctors = (staffListData.doctors || []).map(doctor => ({
        no: typeof doctor.no === 'number' && doctor.no > 0 ? doctor.no : 1,
        name: validateInput(doctor.name, 100)
      }));
      
      const dataToSave = {
        staff: validatedStaff,
        doctors: validatedDoctors,
        updatedAt: Timestamp.now()
      };
      
      // 🔒 보안: 문서 ID 검증
      const safeOffice = selectedOffice.replace(/[^a-zA-Z0-9_-]/g, '');
      await setDoc(doc(db, 'staff-list', safeOffice), dataToSave);
      setSaveStatus('Saved successfully!');
      setTimeout(() => setSaveStatus(''), 2000);
    } catch (error) {
      setSaveStatus('Error saving');
      alert('Error saving staff list: error');
    } finally {
      setLoading(false);
    }
  }, [staffListData, selectedOffice]);

  // Staff 추가
  const addStaffMember = (position: string) => {
    // 🔒 보안: Position 값 검증
    if (!validatePosition(position)) {
      alert('Invalid position');
      return;
    }

    const positionStaff = staffListData.staff[position] || [];
    const maxNo = positionStaff.length > 0 
      ? Math.max(...positionStaff.map(s => s.no)) 
      : 0;
    
    const newStaff: StaffMember = {
      no: maxNo + 1,
      name: '',
      active: position === 'Sub' || position === 'Extern' ? false : true
    };

    setStaffListData(prev => ({
      ...prev,
      staff: {
        ...prev.staff,
        [position]: [...(prev.staff[position] || []), newStaff]
      }
    }));
  };

  // Staff 삭제
  const removeStaffMember = async (position: string, index: number) => {
    // 🔒 보안: Position 값 검증
    if (!validatePosition(position)) {
      alert('Invalid position');
      return;
    }

    // 🔒 보안: 인덱스 검증
    if (typeof index !== 'number' || index < 0) {
      alert('Invalid index');
      return;
    }

    const memberToRemove = staffListData.staff[position]?.[index];
    const memberName = validateInput(memberToRemove?.name, 100) || 'this member';
    
    // 확인 대화상자
    if (!confirm(`Are you sure you want to remove ${memberName} from ${position}?`)) {
      return;
    }
    
    // 삭제 후 나머지 멤버들의 no를 1부터 순차적으로 재정렬
    const filteredStaff = staffListData.staff[position].filter((_, i) => i !== index);
    const reorderedStaff = filteredStaff.map((member, i) => ({
      ...member,
      no: i + 1
    }));
    
    const updatedData = {
      ...staffListData,
      staff: {
        ...staffListData.staff,
        [position]: reorderedStaff
      }
    };
    
    setStaffListData(updatedData);
    
    // 자동 저장
    try {
      setLoading(true);
      setSaveStatus('Saving...');
      
      // 🔒 보안: 데이터 검증 및 정리
      const validatedStaff: StaffList = {};
      Object.keys(updatedData.staff).forEach(position => {
        if (validatePosition(position)) {
          validatedStaff[position] = (updatedData.staff[position] || []).map(member => ({
            no: typeof member.no === 'number' && member.no > 0 ? member.no : 1,
            name: validateInput(member.name, 100),
            active: typeof member.active === 'boolean' ? member.active : true
          }));
        }
      });

      const validatedDoctors = (updatedData.doctors || []).map(doctor => ({
        no: typeof doctor.no === 'number' && doctor.no > 0 ? doctor.no : 1,
        name: validateInput(doctor.name, 100)
      }));
      
      const dataToSave = {
        staff: validatedStaff,
        doctors: validatedDoctors,
        updatedAt: Timestamp.now()
      };
      
      // 🔒 보안: 문서 ID 검증
      const safeOffice = selectedOffice.replace(/[^a-zA-Z0-9_-]/g, '');
      await setDoc(doc(db, 'staff-list', safeOffice), dataToSave);
      
      setSaveStatus('Removed and saved successfully!');
      setTimeout(() => setSaveStatus(''), 2000);
    } catch (error) {
      setSaveStatus('Error saving');
      alert('Error saving after removal: error');
    } finally {
      setLoading(false);
    }
  };

  // Staff 업데이트
  const updateStaffMember = async (position: string, index: number, field: keyof StaffMember, value: any) => {
    // Rate limiting: 입력 반응성을 위해 완화된 제한 적용
    const now = Date.now();
    const fieldKey = `lastUpdateStaff_${position}_${index}_${field}`;
    const lastCall = (window as any)[fieldKey] || 0;

    // 전역 rate limiting: 모든 업데이트에 대해 50ms 제한 (입력 반응성 향상)
    if (now - lastUpdateStaffMemberCall.current < 50) {
      return;
    }
    lastUpdateStaffMemberCall.current = now;

    // 개별 필드 rate limiting: 동일 필드에 대해 100ms 제한 (입력 반응성 향상)
    if (now - lastCall < 100) {
      return;
    }
    (window as any)[fieldKey] = now;

    // 🔒 보안: Position 값 검증
    if (!validatePosition(position)) {
      alert('Invalid position');
      return;
    }

    // 🔒 보안: 인덱스 검증
    if (typeof index !== 'number' || index < 0) {
      alert('Invalid index');
      return;
    }

    // 🔒 보안: 필드 값 검증 및 정리
    let safeValue: any = value;
    if (field === 'name') {
      safeValue = validateInput(value, 100);
    } else if (field === 'no') {
      safeValue = typeof value === 'number' && value > 0 ? value : 1;
    } else if (field === 'active') {
      safeValue = typeof value === 'boolean' ? value : true;
    }

    const updatedData = {
      ...staffListData,
      staff: {
        ...staffListData.staff,
        [position]: staffListData.staff[position].map((member, i) => 
          i === index ? { ...member, [field]: safeValue } : member
        )
      }
    };
    
    setStaffListData(updatedData);
    
    // active 필드 변경 시 자동 저장
    if (field === 'active') {
      try {
        setLoading(true);
        setSaveStatus('Saving...');
        
        // 🔒 보안: 데이터 검증
        const validatedStaff: StaffList = {};
        Object.keys(updatedData.staff).forEach(pos => {
          if (validatePosition(pos)) {
            validatedStaff[pos] = (updatedData.staff[pos] || []).map(member => ({
              no: typeof member.no === 'number' && member.no > 0 ? member.no : 1,
              name: validateInput(member.name, 100),
              active: typeof member.active === 'boolean' ? member.active : true
            }));
          }
        });

        const validatedDoctors = (updatedData.doctors || []).map(doctor => ({
          no: typeof doctor.no === 'number' && doctor.no > 0 ? doctor.no : 1,
          name: validateInput(doctor.name, 100)
        }));
        
        const dataToSave = {
          staff: validatedStaff,
          doctors: validatedDoctors,
          updatedAt: Timestamp.now()
        };
        
        // 🔒 보안: 문서 ID 검증
        const safeOffice = selectedOffice.replace(/[^a-zA-Z0-9_-]/g, '');
        await setDoc(doc(db, 'staff-list', safeOffice), dataToSave);
        
        setSaveStatus('Saved!');
        setTimeout(() => setSaveStatus(''), 2000);
      } catch (error) {
        setSaveStatus('Error saving');
        alert('Error saving active status: error');
      } finally {
        setLoading(false);
      }
    }
  };

  // Doctor 추가
  const addDoctor = () => {
    const maxNo = staffListData.doctors.length > 0
      ? Math.max(...staffListData.doctors.map(d => d.no))
      : 0;
    
    const newDoctor: DoctorMember = {
      no: maxNo + 1,
      name: ''
    };

    setStaffListData(prev => ({
      ...prev,
      doctors: [...prev.doctors, newDoctor]
    }));
  };

  // Doctor 삭제
  const removeDoctor = async (index: number) => {
    // 🔒 보안: 인덱스 검증
    if (typeof index !== 'number' || index < 0 || index >= staffListData.doctors.length) {
      alert('Invalid index');
      return;
    }

    const doctorToRemove = staffListData.doctors[index];
    const doctorName = validateInput(doctorToRemove?.name, 100) || 'this doctor';
    
    // 확인 대화상자
    if (!confirm(`Are you sure you want to remove ${doctorName}?`)) {
      return;
    }
    
    // 삭제 후 나머지 doctor들의 no를 1부터 순차적으로 재정렬
    const filteredDoctors = staffListData.doctors.filter((_, i) => i !== index);
    const reorderedDoctors = filteredDoctors.map((doctor, i) => ({
      ...doctor,
      no: i + 1
    }));
    
    const updatedData = {
      ...staffListData,
      doctors: reorderedDoctors
    };
    
    setStaffListData(updatedData);
    
    // 자동 저장
    try {
      setLoading(true);
      setSaveStatus('Saving...');
      
      // 🔒 보안: 데이터 검증 및 정리
      const validatedStaff: StaffList = {};
      Object.keys(updatedData.staff).forEach(position => {
        if (validatePosition(position)) {
          validatedStaff[position] = (updatedData.staff[position] || []).map(member => ({
            no: typeof member.no === 'number' && member.no > 0 ? member.no : 1,
            name: validateInput(member.name, 100),
            active: typeof member.active === 'boolean' ? member.active : true
          }));
        }
      });

      const validatedDoctors = (updatedData.doctors || []).map(doctor => ({
        no: typeof doctor.no === 'number' && doctor.no > 0 ? doctor.no : 1,
        name: validateInput(doctor.name, 100)
      }));
      
      const dataToSave = {
        staff: validatedStaff,
        doctors: validatedDoctors,
        updatedAt: Timestamp.now()
      };
      
      // 🔒 보안: 문서 ID 검증
      const safeOffice = selectedOffice.replace(/[^a-zA-Z0-9_-]/g, '');
      await setDoc(doc(db, 'staff-list', safeOffice), dataToSave);
      
      setSaveStatus('Removed and saved successfully!');
      setTimeout(() => setSaveStatus(''), 2000);
    } catch (error) {
      setSaveStatus('Error saving');
      alert('Error saving after removal: error');
    } finally {
      setLoading(false);
    }
  };

  // Doctor 업데이트
  const updateDoctor = (index: number, field: keyof DoctorMember, value: any) => {
    // Rate limiting: 입력 반응성을 위해 완화된 제한 적용
    const now = Date.now();
    const fieldKey = `lastUpdateDoctor_${index}_${field}`;
    const lastCall = (window as any)[fieldKey] || 0;

    // 전역 rate limiting: 모든 업데이트에 대해 50ms 제한 (입력 반응성 향상)
    if (now - lastUpdateDoctorCall.current < 50) {
      return;
    }
    lastUpdateDoctorCall.current = now;

    // 개별 필드 rate limiting: 동일 필드에 대해 100ms 제한 (입력 반응성 향상)
    if (now - lastCall < 100) {
      return;
    }
    (window as any)[fieldKey] = now;

    // 🔒 보안: 인덱스 검증
    if (typeof index !== 'number' || index < 0 || index >= staffListData.doctors.length) {
      alert('Invalid index');
      return;
    }

    // 🔒 보안: 필드 값 검증 및 정리
    let safeValue: any = value;
    if (field === 'name') {
      safeValue = validateInput(value, 100);
    } else if (field === 'no') {
      safeValue = typeof value === 'number' && value > 0 ? value : 1;
    }

    setStaffListData(prev => {
      const doctors = [...prev.doctors];
      doctors[index] = {
        ...doctors[index],
        [field]: safeValue
      };
      return {
        ...prev,
        doctors
      };
    });
  };

  // 컴포넌트 마운트 시 사용자 인증 및 role 확인
  useEffect(() => {
    // Firebase Auth 상태 변경 감지
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      try {
        if (!currentUser) {
          // 로그인되지 않은 경우 로그인 페이지로 리다이렉트
          if (typeof window !== 'undefined') {
            window.location.href = '/';
          }
          return;
        }

        // Firestore에서 사용자 role 확인
        const userDoc = await getDoc(doc(db, 'users', currentUser.uid));
        if (!userDoc.exists()) {
          // 사용자 정보가 없는 경우 로그인 페이지로 리다이렉트
          if (typeof window !== 'undefined') {
            window.location.href = '/';
          }
          return;
        }

        const userData = userDoc.data();

        if (userData?.role !== 'HR' && userData?.role !== 'Manager') {
          // 관리자 권한이 없는 경우 로그인 페이지로 리다이렉트
          if (typeof window !== 'undefined') {
            window.location.href = '/';
          }
          return;
        }

        setIsAuthorized(true);
      } catch (error: any) {
        // 오류 발생 시 로그인 페이지로 리다이렉트
        if (typeof window !== 'undefined') {
          window.location.href = '/';
        }
      }
    });

    // 프로덕션 환경에서 HTTPS 강제 (클라이언트 사이드)
    if (process.env.NODE_ENV === 'production' && 
        typeof window !== 'undefined' && 
        window.location.protocol !== 'https:') {
      // HTTP로 접속한 경우 HTTPS로 리다이렉트
      window.location.href = window.location.href.replace('http:', 'https:');
    }

    // cleanup 함수
    return () => {
      unsubscribe();
    };
  }, []);

  // 오피스 변경 시 staff list 다시 불러오기
  useEffect(() => {
    loadStaffList();
  }, [loadStaffList]);

  // 오피스 선택 핸들러
  const handleOfficeChange = (office: string) => {
    // 🔒 보안: 오피스 값 검증
    if (validateOffice(office)) {
      setSelectedOffice(office);
    } else {
      alert('Invalid office value');
    }
  };

  // 인증 확인 중이거나 인증 실패 시 로딩 화면 표시
  if (isAuthorized === null) {
    return (
      <div style={{
        display: 'flex',
        justifyContent: 'center',
        alignItems: 'center',
        height: '100vh',
        background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
        fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
      }}>
        <div style={{ textAlign: 'center' }}>
          <div style={{ fontSize: '24px', marginBottom: '20px' }}>🔐</div>
          <div style={{ fontSize: '18px', color: '#fff' }}>Loading..</div>
        </div>
      </div>
    );
  }


  const styles = {
    container: {
      padding: '20px',
      fontFamily: 'Arial, sans-serif',
      maxWidth: '1200px',
      margin: '0 auto'
    },
    header: {
      fontSize: '2rem',
      fontWeight: 'bold',
      color: '#1976D2',
      marginBottom: '20px'
    },
    section: {
      marginBottom: '40px',
      background: '#fff',
      padding: '20px',
      borderRadius: '8px',
      boxShadow: '0 2px 4px rgba(0,0,0,0.1)'
    },
    sectionTitle: {
      fontSize: '1.5rem',
      fontWeight: 'bold',
      marginBottom: '15px',
      color: '#333'
    },
    table: {
      width: '100%',
      borderCollapse: 'collapse' as const,
      marginBottom: '15px'
    },
    th: {
      background: '#1976D2',
      color: '#fff',
      padding: '10px',
      textAlign: 'left' as const,
      border: '1px solid #ddd'
    },
    td: {
      padding: '8px',
      border: '1px solid #ddd'
    },
    input: {
      width: '100%',
      padding: '5px',
      border: '1px solid #ddd',
      borderRadius: '4px'
    },
    checkbox: {
      width: '20px',
      height: '20px'
    },
    button: {
      padding: '8px 16px',
      margin: '5px',
      border: 'none',
      borderRadius: '4px',
      cursor: 'pointer',
      fontSize: '14px'
    },
    addButton: {
      background: '#4CAF50',
      color: '#fff'
    },
    removeButton: {
      background: '#f44336',
      color: '#fff'
    },
    saveButton: {
      background: '#1976D2',
      color: '#fff',
      padding: '12px 24px',
      fontSize: '16px',
      fontWeight: 'bold'
    },
    status: {
      padding: '10px',
      marginBottom: '10px',
      borderRadius: '4px',
      background: '#4CAF50',
      color: '#fff'
    }
  };

  return (
    <div style={styles.container}>
      <h1 style={styles.header}>Staff List Management</h1>

      {/* 오피스 선택 */}
      <div style={{ marginBottom: '20px', display: 'flex', alignItems: 'center', gap: '10px' }}>
        <label style={{ fontSize: '16px', fontWeight: 'bold', color: '#333' }}>
          Office:
        </label>
        <select
          value={selectedOffice}
          onChange={(e) => handleOfficeChange(e.target.value)}
          style={{
            padding: '8px 16px',
            fontSize: '16px',
            border: '1px solid #ddd',
            borderRadius: '4px',
            cursor: 'pointer',
            background: '#fff'
          }}
        >
          <option value="">Select Office</option>
          {officeOptions.map(office => (
            <option key={office} value={office}>
              {office}
            </option>
          ))}
        </select>
      </div>

      {saveStatus && (
        <div style={styles.status}>{saveStatus}</div>
      )}

      {/* Staff 섹션 */}
      <div style={styles.section}>
        <h2 style={styles.sectionTitle}>Staff Members</h2>
        {positionOptions.map(position => (
          <div key={position} style={{ marginBottom: '30px' }}>
            <h3 style={{ fontSize: '1.2rem', marginBottom: '10px', color: '#555' }}>
              {position}
              <button
                onClick={() => addStaffMember(position)}
                style={{ ...styles.button, ...styles.addButton, marginLeft: '10px' }}
              >
                + Add {position}
              </button>
            </h3>
            <table style={styles.table}>
              <thead>
                <tr>
                  <th style={styles.th}>No.</th>
                  <th style={styles.th}>Name</th>
                  <th style={styles.th}>Active</th>
                  <th style={styles.th}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {(staffListData.staff[position] || []).map((member, index) => (
                  <tr key={index}>
                    <td style={styles.td}>{member.no}</td>
                    <td style={styles.td}>
                      <input
                        type="text"
                        value={member.name}
                        onChange={(e) => updateStaffMember(position, index, 'name', e.target.value)}
                        style={styles.input}
                        placeholder="Enter name"
                      />
                    </td>
                    <td style={styles.td}>
                      <input
                        type="checkbox"
                        checked={member.active !== false}
                        onChange={(e) => updateStaffMember(position, index, 'active', e.target.checked)}
                        style={styles.checkbox}
                      />
                    </td>
                    <td style={styles.td}>
                      <button
                        onClick={() => removeStaffMember(position, index)}
                        style={{ ...styles.button, ...styles.removeButton }}
                      >
                        Remove
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ))}
      </div>

      {/* Doctor 섹션 */}
      <div style={styles.section}>
        <h2 style={styles.sectionTitle}>
          Doctors
          <button
            onClick={addDoctor}
            style={{ ...styles.button, ...styles.addButton, marginLeft: '10px' }}
          >
            + Add Doctor
          </button>
        </h2>
        <table style={styles.table}>
          <thead>
            <tr>
              <th style={styles.th}>No.</th>
              <th style={styles.th}>Name</th>
              <th style={styles.th}>Actions</th>
            </tr>
          </thead>
          <tbody>
            {staffListData.doctors.map((doctor, index) => (
              <tr key={index}>
                <td style={styles.td}>{doctor.no}</td>
                <td style={styles.td}>
                  <input
                    type="text"
                    value={doctor.name}
                    onChange={(e) => updateDoctor(index, 'name', e.target.value)}
                    style={styles.input}
                    placeholder="Enter doctor name"
                  />
                </td>
                <td style={styles.td}>
                  <button
                    onClick={() => removeDoctor(index)}
                    style={{ ...styles.button, ...styles.removeButton }}
                  >
                    Remove
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* 저장 버튼 */}
      <div style={{ textAlign: 'center', marginTop: '30px' }}>
        <button
          onClick={saveStaffList}
          disabled={loading}
          style={styles.saveButton}
        >
          {loading ? 'Saving...' : 'Save Staff List'}
        </button>
      </div>
    </div>
  );
}
