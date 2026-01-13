'use client'

import React, { useState, useEffect, useCallback, useRef } from "react";
import { enableAllSecurityMeasures, sanitizeFirebaseDataClient } from "@/lib/security-client";
import { db } from '@/lib/firebase.config';
import { doc, setDoc, getDoc, onSnapshot, addDoc, collection, deleteDoc } from 'firebase/firestore';
import { getStorage, ref, uploadBytes, getDownloadURL } from 'firebase/storage';

// 개별 테이블 행 컴포넌트
const TreatmentRow = ({ 
  rowIndex,
  treatmentData,
  updateTreatment,
  updateSealantDetail,
  addSealantDetail,
  inputStyle,
  addingSealant,
  isInputDisabled,
  isViewMode
}: {
  rowIndex: number;
  treatmentData: any;
  updateTreatment: (rowIndex: number, field: string, value: string | string[] | boolean) => void;
  updateSealantDetail: (rowIndex: number, detailIndex: number, field: string, value: string) => void;
  addSealantDetail: (rowIndex: number) => void;
  inputStyle: any;
  addingSealant: Set<number>;
  isInputDisabled: boolean;
  isViewMode: boolean;
}) => {
  return (
    <>
      <tr style={{ backgroundColor: rowIndex % 2 === 0 ? '#f9f9f9' : 'white', height: '40px' }}>
        <td style={{ padding: '4px', textAlign: 'center', fontWeight: 'bold', minWidth: '60px' }}>
          {rowIndex + 1}
        </td>
        <td style={{ padding: '0' }}>
          <input
            type="text"
            value={treatmentData.patientName || ''}
            onChange={(e) => updateTreatment(rowIndex, 'patientName', e.target.value)}
            style={{...inputStyle, backgroundColor: isInputDisabled ? '#f5f5f5' : 'white', cursor: isInputDisabled ? 'not-allowed' : 'text'}}
            placeholder="Patient Name"
            disabled={isInputDisabled}
            maxLength={100}
          />
        </td>
        <td style={{ padding: '0' }}>
          <input
            type="time"
            value={treatmentData.startTime || ''}
            onChange={(e) => updateTreatment(rowIndex, 'startTime', e.target.value)}
            style={{...inputStyle, backgroundColor: isInputDisabled ? '#f5f5f5' : 'white', cursor: isInputDisabled ? 'not-allowed' : 'text'}}
            disabled={isInputDisabled}
          />
        </td>
        <td style={{ padding: '0' }}>
          <select
            value={treatmentData.roomNumber || ''}
            onChange={(e) => updateTreatment(rowIndex, 'roomNumber', e.target.value)}
            style={{...inputStyle, backgroundColor: isInputDisabled ? '#f5f5f5' : 'white', cursor: isInputDisabled ? 'not-allowed' : 'pointer'}}
            disabled={isInputDisabled}
          >
            <option value="">Select Room</option>
            <option value="1">Room 1</option>
            <option value="2">Room 2</option>
            <option value="3">Room 3</option>
            <option value="4">Room 4</option>
            <option value="5">Room 5</option>
            <option value="6">Room 6</option>
            <option value="7">Room 7</option>
            <option value="8">Room 8</option>
          </select>
        </td>
        <td style={{ padding: '0' }}>
          <div style={{ 
            border: '1px solid #BDE0FE', 
            borderRadius: '4px', 
            backgroundColor: 'white',
            minHeight: '32px'
          }}>
            <button
              type="button"
              onClick={() => {
                if (!isInputDisabled) {
                  updateTreatment(rowIndex, 'showServices', !treatmentData.showServices);
                }
            }}
            style={{
                width: '100%',
                padding: '4px 6px',
                border: 'none',
                backgroundColor: 'transparent',
                cursor: isInputDisabled ? 'not-allowed' : 'pointer',
                textAlign: 'left',
              fontSize: '13px',
                color: isInputDisabled ? '#999' : '#023047',
                height: '32px'
              }}
              disabled={isInputDisabled}
            >
              {(treatmentData.services || []).length > 0 
                ? (treatmentData.services || []).join(', ')
                : 'Click to select services...'}
            </button>
            {treatmentData.showServices && (
              <div style={{ 
                padding: '8px', 
                borderTop: '1px solid #BDE0FE',
                maxHeight: '120px',
                overflowY: 'auto'
              }}>
                <div style={{ fontSize: '12px', color: '#666', marginBottom: '8px' }}>
                  Select services:
                </div>
                {[
                  'Xray',
                  'Intra Oral Pictures', 
                  'Caries Risk Assessments',
                  'Prophy',
                  'Flouride',
                  'Varnish',
                  'Sealant',
                  'Main assistant during treatment',
                  'Holding/stabilizing the patient\'s head',
                  'Dismissed patient consists of: explained treatment, gave post op instructions, gave tooth brush bag and points',
                  'Postcard'
                ].map((service) => (
                  <label key={service} style={{ 
                    display: 'block', 
                    fontSize: '12px', 
                    marginBottom: '4px',
                    cursor: 'pointer',
                    padding: '2px 0'
                  }}>
                    <input
                      type="checkbox"
                      checked={(treatmentData.services || []).includes(service)}
                      onChange={(e) => {
                        const currentServices = treatmentData.services || [];
                        let newServices;
                        if (e.target.checked) {
                          newServices = [...currentServices, service];
                        } else {
                          newServices = currentServices.filter((s: string) => s !== service);
                        }
                        updateTreatment(rowIndex, 'services', newServices);
                      }}
                      style={{ marginRight: '6px' }}
                    />
                    {service}
                  </label>
                ))}
              </div>
            )}
          </div>
        </td>
        <td style={{ padding: '0' }}>
          <textarea
            value={treatmentData.explanation || ''}
            onChange={(e) => updateTreatment(rowIndex, 'explanation', e.target.value)}
            style={{
              ...inputStyle,
              minHeight: '60px',
              resize: 'vertical',
              fontSize: '13px',
              padding: '4px 6px',
              backgroundColor: isInputDisabled ? '#f5f5f5' : 'white',
              cursor: isInputDisabled ? 'not-allowed' : 'text'
            }}
            disabled={isInputDisabled}
            maxLength={1000}
          />
        </td>
      </tr>
      {treatmentData.services?.includes('Sealant') && (
        <tr>
          <td colSpan={6} style={{ padding: '0', backgroundColor: '#f8f9fa' }}>
            <div style={{ padding: '15px', borderTop: '2px solid #0077B6' }}>
              <h4 style={{ color: '#0077B6', marginBottom: '10px', fontSize: '16px' }}>
                🦷 Sealant Details
              </h4>
              <table style={{ width: '100%', borderCollapse: 'collapse', backgroundColor: 'white', borderRadius: '4px', overflow: 'hidden' }}>
                <thead>
                  <tr style={{ backgroundColor: '#e9ecef' }}>
                    <th style={{ padding: '8px', border: '1px solid #dee2e6', fontSize: '12px', fontWeight: 'bold' }}>PT Name</th>
                    <th style={{ padding: '8px', border: '1px solid #dee2e6', fontSize: '12px', fontWeight: 'bold' }}>Chart #</th>
                    <th style={{ padding: '8px', border: '1px solid #dee2e6', fontSize: '12px', fontWeight: 'bold' }}>DOB</th>
                    <th style={{ padding: '8px', border: '1px solid #dee2e6', fontSize: '12px', fontWeight: 'bold' }}>Tooth #</th>
                    <th style={{ padding: '8px', border: '1px solid #dee2e6', fontSize: '12px', fontWeight: 'bold' }}>Redo</th>
                    <th style={{ padding: '8px', border: '1px solid #dee2e6', fontSize: '12px', fontWeight: 'bold' }}>Acct Type</th>
                    <th style={{ padding: '8px', border: '1px solid #dee2e6', fontSize: '12px', fontWeight: 'bold' }}>Payable</th>
                    <th style={{ padding: '8px', border: '1px solid #dee2e6', fontSize: '12px', fontWeight: 'bold' }}>DX Dr.</th>
                    <th style={{ padding: '8px', border: '1px solid #dee2e6', fontSize: '12px', fontWeight: 'bold' }}>DR $</th>
                    <th style={{ padding: '8px', border: '1px solid #dee2e6', fontSize: '12px', fontWeight: 'bold' }}>RDA $</th>
                  </tr>
                </thead>
                <tbody>
                  {(treatmentData.sealantDetails && treatmentData.sealantDetails.length > 0 ? treatmentData.sealantDetails : [{ ptName: treatmentData.patientName || '' }]).map((detail: any, detailIndex: number) => (
                    <tr key={detailIndex}>
                      <td style={{ padding: '0', border: '1px solid #dee2e6' }}>
                        <input
                          type="text"
                          value={detail.ptName || ''}
                          onChange={(e) => updateSealantDetail(rowIndex, detailIndex, 'ptName', e.target.value)}
                          style={{ ...inputStyle, border: 'none', fontSize: '12px', padding: '6px', backgroundColor: isInputDisabled ? '#f5f5f5' : 'white', cursor: isInputDisabled ? 'not-allowed' : 'text' }}
                          disabled={isInputDisabled}
                        />
                      </td>
                      <td style={{ padding: '0', border: '1px solid #dee2e6' }}>
                        <input
                          type="text"
                          value={detail.chartNumber || ''}
                          onChange={(e) => updateSealantDetail(rowIndex, detailIndex, 'chartNumber', e.target.value)}
                          style={{ ...inputStyle, border: 'none', fontSize: '12px', padding: '6px', backgroundColor: isInputDisabled ? '#f5f5f5' : 'white', cursor: isInputDisabled ? 'not-allowed' : 'text' }}
                          disabled={isInputDisabled}
                        />
                      </td>
                      <td style={{ padding: '0', border: '1px solid #dee2e6' }}>
                        <input
                          type="date"
                          value={detail.dob || ''}
                          onChange={(e) => updateSealantDetail(rowIndex, detailIndex, 'dob', e.target.value)}
                          style={{ ...inputStyle, border: 'none', fontSize: '12px', padding: '6px', backgroundColor: isInputDisabled ? '#f5f5f5' : 'white', cursor: isInputDisabled ? 'not-allowed' : 'text' }}
                          disabled={isInputDisabled}
                        />
                      </td>
                      <td style={{ padding: '0', border: '1px solid #dee2e6' }}>
                        <select
                          value={detail.toothNumber || ''}
                          onChange={(e) => updateSealantDetail(rowIndex, detailIndex, 'toothNumber', e.target.value)}
                          style={{ ...inputStyle, border: 'none', fontSize: '12px', padding: '6px', backgroundColor: isInputDisabled ? '#f5f5f5' : 'white', cursor: isInputDisabled ? 'not-allowed' : 'pointer' }}
                          disabled={isInputDisabled}
                        >
                          <option value=""></option>
                          <option value="2">2</option>
                          <option value="3">3</option>
                          <option value="14">14</option>
                          <option value="15">15</option>
                          <option value="18">18</option>
                          <option value="19">19</option>
                          <option value="30">30</option>
                          <option value="31">31</option>
                        </select>
                      </td>
                      <td style={{ padding: '0', border: '1px solid #dee2e6' }}>
                        <select
                          value={detail.redo || ''}
                          onChange={(e) => updateSealantDetail(rowIndex, detailIndex, 'redo', e.target.value)}
                          style={{ ...inputStyle, border: 'none', fontSize: '12px', padding: '6px', backgroundColor: isInputDisabled ? '#f5f5f5' : 'white', cursor: isInputDisabled ? 'not-allowed' : 'pointer' }}
                          disabled={isInputDisabled}
                        >
                          <option value=""></option>
                          <option value="Yes">Yes</option>
                          <option value="No">No</option>
                        </select>
                      </td>
                      <td style={{ padding: '0', border: '1px solid #dee2e6' }}>
                        <select
                          value={detail.acctType || ''}
                          onChange={(e) => updateSealantDetail(rowIndex, detailIndex, 'acctType', e.target.value)}
                          style={{ ...inputStyle, border: 'none', fontSize: '12px', padding: '6px', backgroundColor: isInputDisabled ? '#f5f5f5' : 'white', cursor: isInputDisabled ? 'not-allowed' : 'pointer' }}
                          disabled={isInputDisabled}
                        >
                          <option value=""></option>
                          <option value="D">D</option>
                          <option value="I">I</option>
                          <option value="PV">PV</option>
                        </select>
                      </td>
                      <td style={{ padding: '0', border: '1px solid #dee2e6' }}>
                        <select
                          value={detail.payable || ''}
                          onChange={(e) => updateSealantDetail(rowIndex, detailIndex, 'payable', e.target.value)}
                          style={{ ...inputStyle, border: 'none', fontSize: '12px', padding: '6px', backgroundColor: isInputDisabled ? '#f5f5f5' : 'white', cursor: isInputDisabled ? 'not-allowed' : 'pointer' }}
                          disabled={isInputDisabled}
                        >
                          <option value=""></option>
                          <option value="Y">Y</option>
                          <option value="N">N</option>
                        </select>
                      </td>
                      <td style={{ padding: '0', border: '1px solid #dee2e6' }}>
                        <input
                          type="text"
                          value={detail.dxDr || ''}
                          onChange={(e) => updateSealantDetail(rowIndex, detailIndex, 'dxDr', e.target.value)}
                          style={{ ...inputStyle, border: 'none', fontSize: '12px', padding: '6px', backgroundColor: isInputDisabled ? '#f5f5f5' : 'white', cursor: isInputDisabled ? 'not-allowed' : 'text' }}
                          disabled={isInputDisabled}
                        />
                      </td>
                      <td style={{ padding: '0', border: '1px solid #dee2e6' }}>
                        <input
                          type="text"
                          inputMode="decimal"
                          value={detail.drAmount || ''}
                          onChange={(e) => {
                            // 숫자와 소수점만 허용
                            const value = e.target.value;
                            if (value === '' || /^\d*\.?\d*$/.test(value)) {
                              updateSealantDetail(rowIndex, detailIndex, 'drAmount', value);
                            }
                          }}
                          style={{ 
                            ...inputStyle, 
                            border: 'none', 
                            fontSize: '12px',
                            padding: '6px',
                            backgroundColor: isInputDisabled ? '#f5f5f5' : 'white',
                            cursor: isInputDisabled ? 'not-allowed' : 'text'
                          }}
                          disabled={isInputDisabled}
                        />
                      </td>
                      <td style={{ padding: '0', border: '1px solid #dee2e6' }}>
                        <input
                          type="text"
                          inputMode="decimal"
                          value={detail.rdaAmount || ''}
                          onChange={(e) => {
                            // 숫자와 소수점만 허용
                            const value = e.target.value;
                            if (value === '' || /^\d*\.?\d*$/.test(value)) {
                              updateSealantDetail(rowIndex, detailIndex, 'rdaAmount', value);
                            }
                          }}
                          style={{ 
                            ...inputStyle, 
                            border: 'none', 
                            fontSize: '12px',
                            padding: '6px',
                            backgroundColor: isInputDisabled ? '#f5f5f5' : 'white',
                            cursor: isInputDisabled ? 'not-allowed' : 'text'
                          }}
                          disabled={isInputDisabled}
                        />
                      </td>
                    </tr>
                  ))}
                  {!isViewMode && (
                    <tr>
                      <td colSpan={10} style={{ padding: '8px', textAlign: 'center', border: '1px solid #dee2e6', backgroundColor: '#f8f9fa' }}>
                        <button
                          disabled={addingSealant.has(rowIndex) || isInputDisabled}
                          onClick={(e) => {
                            if (!isInputDisabled) {
                              e.preventDefault();
                              e.stopPropagation();
                              addSealantDetail(rowIndex);
                            }
                          }}
                          style={{
                            backgroundColor: (addingSealant.has(rowIndex) || isInputDisabled) ? '#6c757d' : '#28a745',
                            color: 'white',
                            border: 'none',
                            padding: '8px 16px',
                            borderRadius: '6px',
                            fontSize: '13px',
                            cursor: (addingSealant.has(rowIndex) || isInputDisabled) ? 'not-allowed' : 'pointer',
                            fontWeight: 'bold',
                            boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
                            transition: 'all 0.2s ease',
                            opacity: (addingSealant.has(rowIndex) || isInputDisabled) ? 0.6 : 1
                          }}
                          onMouseOver={(e) => {
                            e.currentTarget.style.backgroundColor = '#218838';
                            e.currentTarget.style.transform = 'translateY(-1px)';
                          }}
                          onMouseOut={(e) => {
                            e.currentTarget.style.backgroundColor = '#28a745';
                            e.currentTarget.style.transform = 'translateY(0)';
                          }}
                        >
                          ➕ Add Sealant Detail
                        </button>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </td>
        </tr>
      )}
    </>
  );
};

export default function RDATreatmentSheetSystem() {
  // 중복 실행 방지를 위한 ref
  const addingSealantRef = useRef<Set<number>>(new Set());
  const [addingSealant, setAddingSealant] = useState<Set<number>>(new Set());
  const isProcessingRef = useRef<boolean>(false);
  
  // Firebase 관련 상태
  const [autoSaveStatus, setAutoSaveStatus] = useState('');
  const [isUpdatingFromFirebase, setIsUpdatingFromFirebase] = useState(false);
  const [userSessionId] = useState(() => Math.random().toString(36).substr(2, 9));
  const [lastSavedData, setLastSavedData] = useState<any>({});
  const [isSubmitted, setIsSubmitted] = useState(false); // 제출 상태 추적
  const pdfGeneratedRef = useRef(false); // Generate PDF 버튼을 눌렀는지 추적
  
  // Hydration 오류 방지를 위한 클라이언트 마운트 상태
  const [isClient, setIsClient] = useState(false);
  
  // 이전 조합 추적 (변경 감지 및 이전 조합 저장용)
  const previousCombinationRef = useRef<{office: string, rdaName: string, date: string} | null>(null);
  const previousTreatmentDataRef = useRef<any[]>([]);
  
  // 조합 변경 중인지 추적 (autoSave 방지용)
  const isChangingCombinationRef = useRef<boolean>(false);
  
  // Debounce 타이머 (빠른 타이핑 시 중간 값 무시용)
  const combinationChangeTimerRef = useRef<NodeJS.Timeout | null>(null);
  
  // URL 파라미터에서 초기값 가져오기
  const getInitialValues = () => {
    // 캘리포니아 시간대의 현재 날짜를 YYYY-MM-DD 형식으로 반환
    const getPacificDate = () => {
      const now = new Date();
      // 캘리포니아 시간대의 날짜 정보를 직접 가져오기
      const pacificDate = new Date(now.toLocaleString("en-US", {timeZone: "America/Los_Angeles"}));
      const year = pacificDate.getFullYear();
      const month = String(pacificDate.getMonth() + 1).padStart(2, '0');
      const day = String(pacificDate.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    };

    if (typeof window !== 'undefined') {
      const urlParams = new URLSearchParams(window.location.search);
      return {
        office: urlParams.get('office') || '',
        rdaName: urlParams.get('rdaName') || '',
        date: urlParams.get('date') || getPacificDate()
      };
    }
    return {
      office: '',
      rdaName: '',
      date: getPacificDate()
    };
  };

  const initialValues = getInitialValues();

  // 기본 정보 상태
  const [office, setOffice] = useState(initialValues.office);
  const [rdaName, setRdaName] = useState(initialValues.rdaName);
  const [date, setDate] = useState(initialValues.date);
  
  // Office별 비밀번호 잠금 상태
  const [unlockedOffices, setUnlockedOffices] = useState<Set<string>>(new Set());
  const [showPasswordModal, setShowPasswordModal] = useState(false);
  const [passwordInput, setPasswordInput] = useState('');
  const [selectedOfficeForPassword, setSelectedOfficeForPassword] = useState('');
  
  // 테이블 잠금 해제 상태 (Unlock 버튼으로 제어)
  const [isUnlocked, setIsUnlocked] = useState(false);
  const [unlockedCombination, setUnlockedCombination] = useState<{office: string, rdaName: string, date: string} | null>(null);
  // Unlock된 조합의 treatmentData를 추적
  const unlockedTreatmentDataRef = useRef<{combination: {office: string, rdaName: string, date: string}, data: any[]} | null>(null);
  
  // 치료 데이터 상태 (20행으로 시작)
  const [treatmentData, setTreatmentData] = useState<any[]>(() => {
    const initialData: any[] = [];
    for (let i = 0; i < 20; i++) {
      initialData.push({
        patientName: '',
        startTime: '',
        roomNumber: '',
        services: [],
        explanation: '',
        showServices: false,
        sealantDetails: []
      });
    }
    return initialData;
  });

  // 최신 값들을 ref로 추적 (무한 루프 방지)
  const officeRef = useRef(office);
  const rdaNameRef = useRef(rdaName);
  const dateRef = useRef(date);
  const treatmentDataRef = useRef(treatmentData);
  
  // ref 값들 업데이트
  officeRef.current = office;
  rdaNameRef.current = rdaName;
  dateRef.current = date;
  treatmentDataRef.current = treatmentData;

  // 로딩 상태
  const [loading, setLoading] = useState(false);

  // Office 잠금 해제 확인
  const isOfficeUnlocked = unlockedOffices.has(office);
  
  // 입력 비활성화 조건 확인 (Office가 잠금 해제되고 Unlock 버튼을 눌러야 함)
  const isInputDisabled = !office || !rdaName || !date || !isOfficeUnlocked || !isUnlocked;
  
  // 필드 변경 시 잠금 처리
  const handleDateChange = (newDate: string) => {
    setDate(newDate);
    setIsUnlocked(false);
    setUnlockedCombination(null);
    unlockedTreatmentDataRef.current = null;
    // 테이블 데이터 초기화
    const emptyRows = Array.from({ length: 20 }, () => ({
      patientName: '',
      startTime: '',
      roomNumber: '',
      services: [],
      explanation: '',
      showServices: false,
      sealantDetails: []
    }));
    setTreatmentData(emptyRows);
  };
  
  const handleRdaNameChange = (newRdaName: string) => {
    setRdaName(newRdaName);
    setIsUnlocked(false);
    setUnlockedCombination(null);
    unlockedTreatmentDataRef.current = null;
    // 테이블 데이터 초기화
    const emptyRows = Array.from({ length: 20 }, () => ({
      patientName: '',
      startTime: '',
      roomNumber: '',
      services: [],
      explanation: '',
      showServices: false,
      sealantDetails: []
    }));
    setTreatmentData(emptyRows);
  };
  
  // Unlock 버튼 핸들러
  const handleUnlock = async () => {
    if (!office || !rdaName || !date || !isOfficeUnlocked) {
      return;
    }
    const currentCombination = { office, rdaName, date };
    const previousUnlockedCombination = unlockedCombination;
    const previousUnlockedData = unlockedTreatmentDataRef.current;
    
    // 이전에 Unlock된 조합이 있고, 조합이 변경되었다면 먼저 저장
    if (previousUnlockedCombination && 
        previousUnlockedData &&
        (previousUnlockedCombination.office !== currentCombination.office ||
         previousUnlockedCombination.rdaName !== currentCombination.rdaName ||
         previousUnlockedCombination.date !== currentCombination.date)) {
      // 이전 조합의 데이터 저장 (현재 treatmentData가 아닌 이전 조합의 데이터 사용)
      await savePreviousCombination(
        previousUnlockedCombination.office,
        previousUnlockedCombination.rdaName,
        previousUnlockedCombination.date,
        previousUnlockedData.data
      );
    }
    
    setIsUnlocked(true);
    setUnlockedCombination(currentCombination);
    // previousCombinationRef 업데이트
    previousCombinationRef.current = { ...currentCombination };
    // Unlock 시 데이터 로드
    await loadData();
    // loadData 완료 후 useEffect에서 unlockedTreatmentDataRef가 자동으로 업데이트됨
  };
  
  // Office 선택 핸들러 (비밀번호 확인)
  const handleOfficeChange = (selectedOffice: string) => {
    // 테이블 데이터 초기화
    const emptyRows = Array.from({ length: 20 }, () => ({
      patientName: '',
      startTime: '',
      roomNumber: '',
      services: [],
      explanation: '',
      showServices: false,
      sealantDetails: []
    }));
    
    if (!selectedOffice) {
      setOffice('');
      setIsUnlocked(false);
      setUnlockedCombination(null);
      unlockedTreatmentDataRef.current = null;
      setTreatmentData(emptyRows);
      return;
    }
    
    // 이미 잠금 해제된 Office면 바로 설정
    if (unlockedOffices.has(selectedOffice)) {
      setOffice(selectedOffice);
      setIsUnlocked(false);
      setUnlockedCombination(null);
      unlockedTreatmentDataRef.current = null;
      setTreatmentData(emptyRows);
      return;
    }
    
    // 잠금 해제되지 않은 Office면 비밀번호 모달 표시
    setSelectedOfficeForPassword(selectedOffice);
    setPasswordInput('');
    setShowPasswordModal(true);
    setIsUnlocked(false);
    setUnlockedCombination(null);
    unlockedTreatmentDataRef.current = null;
    setTreatmentData(emptyRows);
  };
  
  // 비밀번호 확인 핸들러
  const handlePasswordSubmit = () => {
    if (passwordInput.trim().toUpperCase() === selectedOfficeForPassword.toUpperCase()) {
      // 비밀번호가 맞으면 Office 잠금 해제 (테이블은 아직 잠김)
      setUnlockedOffices(prev => new Set(prev).add(selectedOfficeForPassword));
      setOffice(selectedOfficeForPassword);
      setShowPasswordModal(false);
      setPasswordInput('');
      setSelectedOfficeForPassword('');
      setIsUnlocked(false);
      setUnlockedCombination(null);
      unlockedTreatmentDataRef.current = null;
    } else {
      setPasswordInput('');
    }
  };
  
  // 비밀번호 모달 닫기
  const handlePasswordModalClose = () => {
    setShowPasswordModal(false);
    setPasswordInput('');
    setSelectedOfficeForPassword('');
  };

  // 안전한 docId 생성 함수 (입력 검증 포함)
  const createSafeDocId = useCallback((date: string, office: string, rdaName: string): string => {
    // 입력값 정리 및 검증
    const safeDate = date.trim().replace(/[^0-9-]/g, '');
    const safeOffice = office.trim().replace(/[^A-Za-z0-9]/g, '').substring(0, 10);
    const safeRdaName = rdaName.trim().replace(/[^A-Za-z0-9\s]/g, '').substring(0, 50);
    
    return `${safeDate}_${safeOffice}_${safeRdaName}_rda-treatment`;
  }, []);

  // 자동 저장 함수
  // Office, RDA/DA Name, Date 조합으로 고유한 문서 ID로 저장
  // 각 조합마다 별도의 데이터로 저장됨
  const autoSave = useCallback(async () => {
    if (!office || !rdaName || !date || isUpdatingFromFirebase || !isUnlocked) {
      return;
    }

    // treatmentData에 실제 데이터가 있는지 확인
    const hasData = treatmentData.some(row => 
      row.patientName?.trim() || 
      row.startTime?.trim() || 
      row.roomNumber?.trim() || 
      row.services?.length > 0 || 
      row.explanation?.trim() ||
      (row.sealantDetails && row.sealantDetails.length > 0 && row.sealantDetails.some((detail: any) => 
        detail.ptName?.trim() || 
        detail.chartNumber?.trim() || 
        detail.dob?.trim() || 
        detail.toothNumber?.trim() || 
        detail.redo?.trim() || 
        detail.acctType?.trim() || 
        detail.payable?.trim() || 
        detail.dxDr?.trim() || 
        detail.drAmount || 
        detail.rdaAmount
      ))
    );
    
    // 데이터가 없으면 저장하지 않음
    if (!hasData) {
      return;
    }

    // 데이터가 실제로 변경되었는지 확인
    const currentData = { office, rdaName, date, treatmentData };
    const hasChanges = JSON.stringify(currentData) !== JSON.stringify(lastSavedData);
    if (!hasChanges) {
      return;
    }

    try {
      const docId = createSafeDocId(date, office, rdaName);
      setAutoSaveStatus('💾 Saving...');
      
      // 캘리포니아 시간대로 현재 시간 생성
      const currentTime = new Date();
      const pacificDateTime = new Date(currentTime.toLocaleString("en-US", {timeZone: "America/Los_Angeles"}));
      
      const dataToSave = {
        office,
        rdaName,
        date,
        treatmentData,
        timestamp: pacificDateTime.toISOString(),
        autoSaved: true,
        lastUpdatedBy: userSessionId,
        submitted: isSubmitted // 기존 제출 상태 유지 (view 모드에서 수정 시에도 submitted: true 유지)
      };

      // 데이터 검증 및 sanitization
      const sanitizedData = sanitizeFirebaseDataClient(dataToSave);

      // sealant details가 포함된 데이터 확인
      const sealantRows = treatmentData.filter((row: any) => row.sealantDetails && row.sealantDetails.length > 0);

      // Office, RDA/DA Name, Date 조합으로 고유한 문서 ID 생성
      // 이 조합이 변경되면 새로운 문서로 저장됨
      await setDoc(doc(db, "rda-treatment-sheets", docId), sanitizedData);
      
      // 저장 성공 시 마지막 저장된 데이터 업데이트 (깊은 복사)
      const deepCopy = JSON.parse(JSON.stringify(currentData));
      setLastSavedData(deepCopy);
      
      setAutoSaveStatus('💾 Saved ✅');
      setTimeout(() => setAutoSaveStatus(''), 2000);
      
    } catch (error) {
      // 프로덕션에서는 상세한 에러 정보를 노출하지 않음
      if (process.env.NODE_ENV === 'development') {
        console.error("Auto-save error:", error);
      }
      setAutoSaveStatus('💾 Save failed ❌');
      setTimeout(() => setAutoSaveStatus(''), 3000);
    }
  }, [office, rdaName, date, treatmentData, lastSavedData, isUpdatingFromFirebase, userSessionId, isSubmitted, isUnlocked, createSafeDocId]);

  // 데이터 변경 시에만 자동 저장
  useEffect(() => {
    // Firebase 업데이트 중이거나 필수 필드가 없거나 Unlock되지 않았으면 저장하지 않음
    if (isUpdatingFromFirebase || !office || !rdaName || !date || !isUnlocked) {
      return;
    }

    // Unlock된 조합과 현재 조합이 일치하는지 확인
    if (!unlockedCombination || 
        unlockedCombination.office !== office || 
        unlockedCombination.rdaName !== rdaName || 
        unlockedCombination.date !== date) {
      return;
    }

    // View Details로 접근했는지 확인 (URL에 view=true가 있는지)
    const isViewMode = typeof window !== 'undefined' && 
      new URLSearchParams(window.location.search).get('view') === 'true';
    
    // 조합이 변경되었는지 확인 (이전 조합과 비교)
    const currentCombination = { office, rdaName, date };
    const previousCombination = previousCombinationRef.current;
    const combinationChanged = !previousCombination || 
      previousCombination.office !== office || 
      previousCombination.rdaName !== rdaName || 
      previousCombination.date !== date;
    
    // view 모드이거나 조합이 변경되지 않은 경우 (같은 데이터 수정) autoSave 실행
    // view 모드에서는 항상 자동 저장 활성화
    if (isViewMode || !combinationChanged) {
      // 데이터가 있는 경우에만 저장 (빈 데이터도 저장 가능하도록 수정)
      autoSave();
    }
  }, [treatmentData, autoSave, office, rdaName, date, isUpdatingFromFirebase, isUnlocked, unlockedCombination]);

  // 치료 데이터 업데이트
  const updateTreatment = useCallback((rowIndex: number, field: string, value: string | string[] | boolean) => {
    setTreatmentData(prev => {
      const newData = [...prev];
      
      // showServices가 변경될 때 다른 모든 행의 showServices를 false로 설정
      if (field === 'showServices') {
        newData.forEach((row, index) => {
          if (index === rowIndex) {
            newData[index] = {
              ...newData[index],
              [field]: value
            };
          } else {
            newData[index] = {
              ...newData[index],
              showServices: false
            };
          }
        });
      } else if (field === 'patientName') {
        // Patient Name이 변경되면 Sealant Details의 모든 PT Name도 업데이트
      newData[rowIndex] = {
        ...newData[rowIndex],
        [field]: value
      };
        
        if (newData[rowIndex].sealantDetails) {
          newData[rowIndex].sealantDetails = newData[rowIndex].sealantDetails.map((detail: any) => ({
            ...detail,
            ptName: value
          }));
        }
      } else {
        newData[rowIndex] = {
          ...newData[rowIndex],
          [field]: value
        };
      }
      
      return newData;
    });
  }, []);

  // Sealant 상세 정보 업데이트
  const updateSealantDetail = useCallback((rowIndex: number, detailIndex: number, field: string, value: string) => {
    setTreatmentData(prev => {
      const newData = [...prev];
      if (!newData[rowIndex].sealantDetails) {
        newData[rowIndex].sealantDetails = [];
      }
      const newDetails = [...newData[rowIndex].sealantDetails];
      
      // 기존 데이터가 없으면 기본값으로 초기화
      if (!newDetails[detailIndex]) {
        newDetails[detailIndex] = {
          ptName: prev[rowIndex].patientName || '',
          chartNumber: '',
          dob: '',
          toothNumber: '',
          redo: '',
          acctType: '',
          payable: '',
          dxDr: '',
          drAmount: '',
          rdaAmount: ''
        };
      }
      
      newDetails[detailIndex] = {
        ...newDetails[detailIndex],
        [field]: value
      };
      
      // 첫 번째 줄의 특정 필드가 변경되면 다른 모든 줄도 동일하게 업데이트
      if (detailIndex === 0 && (field === 'chartNumber' || field === 'dob')) {
        newDetails.forEach((detail, index) => {
          if (index > 0) {
            newDetails[index] = {
              ...newDetails[index],
              [field]: value
            };
          }
        });
      }
      
      newData[rowIndex].sealantDetails = newDetails;
      return newData;
    });
  }, []);

  // Sealant 상세 정보 추가
  const addSealantDetail = useCallback((rowIndex: number) => {
    // 전역 처리 중 체크
    if (isProcessingRef.current) {
      return;
    }
    
    // 중복 실행 방지
    if (addingSealant.has(rowIndex)) {
      return;
    }
    
    // 전역 처리 플래그 설정
    isProcessingRef.current = true;
    setAddingSealant(prev => new Set(prev).add(rowIndex));
    
    setTreatmentData(prev => {
      // 이미 처리된 경우 중복 실행 방지
      if (isProcessingRef.current === false) {
        return prev;
      }
      
      const newData = [...prev];
      if (!newData[rowIndex].sealantDetails) {
        newData[rowIndex].sealantDetails = [];
      }
      
      // 현재 행의 Patient Name과 첫 번째 Chart #, DOB 가져오기
      const currentPatientName = newData[rowIndex].patientName || '';
      const firstChartNumber = newData[rowIndex].sealantDetails && newData[rowIndex].sealantDetails.length > 0 
        ? newData[rowIndex].sealantDetails[0].chartNumber || '' 
        : '';
      const firstDob = newData[rowIndex].sealantDetails && newData[rowIndex].sealantDetails.length > 0 
        ? newData[rowIndex].sealantDetails[0].dob || '' 
        : '';
      
      // 새로운 detail 객체 생성
      const newDetail = {
        ptName: currentPatientName,
        chartNumber: firstChartNumber,
        dob: firstDob,
        toothNumber: '',
        redo: '',
        acctType: '',
        payable: '',
        dxDr: '',
        drAmount: '',
        rdaAmount: ''
      };
      
      // 기존 배열에 새 detail 추가
      newData[rowIndex].sealantDetails = [...newData[rowIndex].sealantDetails, newDetail];
      
      // 즉시 플래그 해제
      isProcessingRef.current = false;
      
      // 실행 완료 후 상태에서 제거
      setTimeout(() => {
        setAddingSealant(prev => {
          const newSet = new Set(prev);
          newSet.delete(rowIndex);
          return newSet;
        });
      }, 100);
      
      return newData;
    });
  }, [addingSealant]);

  // 이전 조합으로 현재 데이터 저장 (조합 변경 시 호출)
  const savePreviousCombination = useCallback(async (
    prevOffice: string, 
    prevRdaName: string, 
    prevDate: string, 
    prevTreatmentData: any[]
  ) => {
    if (!prevOffice || !prevRdaName || !prevDate) {
      return;
    }
    
    // 데이터가 실제로 있는지 확인
    const hasData = prevTreatmentData.some(row => 
      row.patientName || row.startTime || row.roomNumber || row.services?.length > 0 || row.explanation
    );
    
    if (!hasData) {
      return; // 데이터가 없으면 저장하지 않음
    }
    
    try {
      // 캘리포니아 시간대로 현재 시간 생성
      const currentTime = new Date();
      const pacificDateTime = new Date(currentTime.toLocaleString("en-US", {timeZone: "America/Los_Angeles"}));
      
      const dataToSave = {
        office: prevOffice,
        rdaName: prevRdaName,
        date: prevDate,
        treatmentData: prevTreatmentData,
        timestamp: pacificDateTime.toISOString(),
        autoSaved: true,
        lastUpdatedBy: userSessionId,
        submitted: isSubmitted // 제출 상태 유지
      };

      const prevDocId = createSafeDocId(prevDate, prevOffice, prevRdaName);
      await setDoc(doc(db, "rda-treatment-sheets", prevDocId), dataToSave);
    } catch (error) {
      // 프로덕션에서는 상세한 에러 정보를 노출하지 않음
      if (process.env.NODE_ENV === 'development') {
        console.error("Error saving previous combination:", error);
      }
    }
  }, [userSessionId, isSubmitted, createSafeDocId]);

  // 데이터 로드 함수 (기존 데이터가 있으면 로드, submitted된 데이터는 view mode에서만 접근 가능)
  const loadData = useCallback(async () => {
    if (!office || !rdaName || !date) {
      return Promise.resolve();
    }

    try {
      // Generate PDF 버튼을 눌렀으면 데이터를 로드하지 않음
      if (pdfGeneratedRef.current) {
        return Promise.resolve();
      }

      // URL에서 view 파라미터 확인
      const isViewMode = typeof window !== 'undefined' && 
        new URLSearchParams(window.location.search).get('view') === 'true';
      
      // 기존 데이터가 있는지 확인 (view mode가 아니어도 확인)
      setAutoSaveStatus('Checking for existing data...');
      
      const docId = createSafeDocId(date, office, rdaName);
      const docSnap = await getDoc(doc(db, "rda-treatment-sheets", docId));
      
      if (docSnap.exists()) {
        const data = docSnap.data();
        
        // submitted 상태 저장 (view 모드에서 수정 시에도 제출 상태 유지)
        setIsSubmitted(data.submitted === true);
        
        // submitted가 true인 경우 view mode에서만 로드 가능
        // 일반 사용자가 제출된 데이터 조합을 입력하면 데이터를 로드하지 않고 빈 시트로 초기화
        if (data.submitted === true && !isViewMode) {
          setAutoSaveStatus('This data has been submitted. Only accessible via View page.');
          // 새 시트 초기화 (제출된 데이터는 로드하지 않음)
          setIsUpdatingFromFirebase(true);
          const emptyRows = Array.from({ length: 20 }, () => ({
            patientName: '',
            startTime: '',
            roomNumber: '',
            services: [],
            explanation: '',
            showServices: false,
            sealantDetails: []
          }));
          
          setTreatmentData(emptyRows);
          // 깊은 복사로 lastSavedData 업데이트
          const emptyDataCopy = JSON.parse(JSON.stringify({ 
            office: office, 
            rdaName: rdaName, 
            date: date, 
            treatmentData: emptyRows 
          }));
          setLastSavedData(emptyDataCopy);
          previousTreatmentDataRef.current = [...emptyRows];
          setIsSubmitted(false); // 새 시트는 제출되지 않은 상태
          
          setTimeout(() => {
            setIsUpdatingFromFirebase(false);
          }, 100);
          
          setTimeout(() => setAutoSaveStatus(''), 3000);
          return;
        }
        
        // submitted가 false이거나 undefined인 경우 (아직 제출되지 않은 데이터) - 로드 가능
        setAutoSaveStatus('Loading existing data...');
        setIsUpdatingFromFirebase(true);
        
        // 치료 데이터를 정규화하여 sealantDetails가 없는 경우 빈 배열로 초기화
        const normalizedTreatmentData = (data.treatmentData || []).map((row: any) => ({
          patientName: row.patientName || '',
          startTime: row.startTime || '',
          roomNumber: row.roomNumber || '',
          services: row.services || [],
          explanation: row.explanation || '',
          showServices: row.showServices || false,
          sealantDetails: row.sealantDetails || []
        }));
        
        // 20개 행을 유지하도록 빈 행 추가
        const paddedData = [...normalizedTreatmentData];
        while (paddedData.length < 20) {
          paddedData.push({
            patientName: '',
            startTime: '',
            roomNumber: '',
            services: [],
            explanation: '',
            showServices: false,
            sealantDetails: []
          });
        }
        
        setTreatmentData(paddedData);
        // 깊은 복사로 lastSavedData 업데이트 (sealantDetails 포함)
        const savedDataCopy = JSON.parse(JSON.stringify({ 
          office: office, 
          rdaName: rdaName, 
          date: date, 
          treatmentData: paddedData 
        }));
        setLastSavedData(savedDataCopy);
        previousTreatmentDataRef.current = [...paddedData];
        
        setTimeout(() => {
          setIsUpdatingFromFirebase(false);
        }, 100);
      } else {
        // 데이터가 없는 경우 - 새 시트 초기화
        setAutoSaveStatus('Preparing new sheet...');
        
        setIsUpdatingFromFirebase(true);
        const emptyRows = Array.from({ length: 20 }, () => ({
          patientName: '',
          startTime: '',
          roomNumber: '',
          services: [],
          explanation: '',
          showServices: false,
          sealantDetails: []
        }));
        
        setTreatmentData(emptyRows);
        // 깊은 복사로 lastSavedData 업데이트
        const newSheetDataCopy = JSON.parse(JSON.stringify({ 
          office: office, 
          rdaName: rdaName, 
          date: date, 
          treatmentData: emptyRows 
        }));
        setLastSavedData(newSheetDataCopy);
        previousTreatmentDataRef.current = [...emptyRows];
        setIsSubmitted(false); // 새 시트는 제출되지 않은 상태
        
        setTimeout(() => {
          setIsUpdatingFromFirebase(false);
        }, 100);
        
        setAutoSaveStatus('New sheet ready ✅');
        setTimeout(() => setAutoSaveStatus(''), 2000);
      }
    } catch (error: any) {
      // 프로덕션에서는 상세한 에러 정보를 노출하지 않음
      if (process.env.NODE_ENV === 'development') {
        console.error("Error loading data:", error);
      }
      setAutoSaveStatus('Error loading data');
      setTimeout(() => setAutoSaveStatus(''), 3000);
    }
  }, [office, rdaName, date]);

  // treatmentData가 변경될 때마다 현재 조합의 데이터를 추적
  useEffect(() => {
    if (office && rdaName && date) {
      // 현재 조합과 이전 조합이 같으면 treatmentData 업데이트
      if (previousCombinationRef.current &&
          previousCombinationRef.current.office === office && 
          previousCombinationRef.current.rdaName === rdaName && 
          previousCombinationRef.current.date === date) {
        previousTreatmentDataRef.current = [...treatmentData];
      }
      
      // Unlock된 조합과 현재 조합이 일치하면 unlockedTreatmentDataRef 업데이트
      if (isUnlocked && unlockedCombination &&
          unlockedCombination.office === office &&
          unlockedCombination.rdaName === rdaName &&
          unlockedCombination.date === date) {
        unlockedTreatmentDataRef.current = {
          combination: { ...unlockedCombination },
          data: JSON.parse(JSON.stringify(treatmentData)) // 깊은 복사
        };
      }
    }
  }, [treatmentData, office, rdaName, date, isUnlocked, unlockedCombination]);

  // 조합 변경 처리 함수 (debounce 없이 직접 호출)
  const handleCombinationChange = useCallback(() => {
    // 최신 값들을 ref에서 가져오기 (무한 루프 방지)
    const currentOffice = officeRef.current;
    const currentRdaName = rdaNameRef.current;
    const currentDate = dateRef.current;
    const currentTreatmentData = treatmentDataRef.current;
    
    if (!currentOffice || !currentRdaName || !currentDate) return;
    
    const currentCombination = { office: currentOffice, rdaName: currentRdaName, date: currentDate };
    const previousCombination = previousCombinationRef.current;
    
    // 조합이 변경되었는지 확인
    const combinationChanged = !previousCombination || 
      previousCombination.office !== currentOffice || 
      previousCombination.rdaName !== currentRdaName || 
      previousCombination.date !== currentDate;
    
    if (combinationChanged && previousCombination) {
      // 조합 변경 중 플래그 설정 (autoSave 방지)
      isChangingCombinationRef.current = true;
      
      // 조합이 변경되었고 이전 조합이 있으면
      // 현재 treatmentData는 아직 이전 조합의 데이터이므로 그대로 저장
      // 이전 조합 저장 후 새 조합 로드
      savePreviousCombination(
        previousCombination.office,
        previousCombination.rdaName,
        previousCombination.date,
        [...currentTreatmentData] // 현재 상태의 데이터를 이전 조합으로 저장
      ).then(() => {
        // 이전 조합 저장 완료 후 새 조합 로드
        previousCombinationRef.current = { ...currentCombination };
        
        loadData().then(() => {
          // 새 조합 로드 완료 후 플래그 해제
          isChangingCombinationRef.current = false;
        });
      });
      
      return;
    }
    
    // 조합이 변경되지 않았거나 첫 로드인 경우
    // 이전 조합 업데이트 (treatmentData는 loadData 완료 후 또는 별도 useEffect에서 업데이트)
    previousCombinationRef.current = { ...currentCombination };
    
    // 새로운 조합의 데이터 로드 (이때 treatmentData가 교체되고, loadData에서 ref도 업데이트함)
    loadData();
  }, [savePreviousCombination, loadData]);

  // Office, RDA Name, Date 변경 시 잠금 처리 (자동 로드 제거)
  // Unlock 버튼을 눌러야만 데이터 로드
  useEffect(() => {
    // 필수 필드 중 하나라도 비어있으면 빈 테이블로 초기화 (20개 행)
    if (!office || !rdaName || !date) {
      setIsUpdatingFromFirebase(true);
      const emptyRows = Array.from({ length: 20 }, () => ({
        patientName: '',
        startTime: '',
        roomNumber: '',
        services: [],
        explanation: '',
        showServices: false,
        sealantDetails: []
      }));
      setTreatmentData(emptyRows);
      // 깊은 복사로 lastSavedData 업데이트
      const resetDataCopy = JSON.parse(JSON.stringify({ office: '', rdaName: '', date: '', treatmentData: emptyRows }));
      setLastSavedData(resetDataCopy);
      previousCombinationRef.current = null;
      previousTreatmentDataRef.current = [];
      setIsSubmitted(false); // 초기화 시 제출 상태도 초기화
      setIsUnlocked(false);
      setUnlockedCombination(null);
      
      setTimeout(() => {
        setIsUpdatingFromFirebase(false);
      }, 100);
    } else {
      // 필드가 모두 채워져도 자동으로 잠금 해제하지 않음 (Unlock 버튼 필요)
      // 조합이 변경되면 잠금 처리
      const currentCombination = { office, rdaName, date };
      if (unlockedCombination && (
        unlockedCombination.office !== office || 
        unlockedCombination.rdaName !== rdaName || 
        unlockedCombination.date !== date
      )) {
        setIsUnlocked(false);
        setUnlockedCombination(null);
      }
    }
  }, [office, rdaName, date, unlockedCombination]);

  // 폼 저장 및 view 페이지로 이동
  const handleSave = useCallback(async () => {
    if (!office.trim() || !rdaName.trim()) {
      alert('⚠️ Please fill in Office and RDA/DA Name!');
      return;
    }

    try {
      setLoading(true);
      
      // 캘리포니아 시간대로 현재 시간 생성
      const currentTime = new Date();
      const pacificDateTime = new Date(currentTime.toLocaleString("en-US", {timeZone: "America/Los_Angeles"}));
      
      // 선택된 날짜가 캘리포니아 시간대인지 확인하고 필요시 조정
      const selectedDateObj = new Date(date + 'T00:00:00'); // 로컬 시간으로 해석
      const pacificDateString = pacificDateTime.toISOString().split('T')[0];
      
      const formData = {
        office: office.trim(),
        rdaName: rdaName.trim(),
        date: date, // 이미 캘리포니아 시간대로 설정된 날짜
        treatmentData: treatmentData,
        createdAt: pacificDateTime.toISOString(),
        lastUpdated: pacificDateTime.toISOString()
      };

      // 데이터 검증
      const sanitizedData = sanitizeFirebaseDataClient(formData);
      
      // autoSave와 동일한 문서 ID 사용 (중복 저장 방지)
      const docId = createSafeDocId(date, office, rdaName);
      
      // Firebase에 저장 (setDoc 사용으로 기존 문서 업데이트)
      const finalData = {
        ...sanitizedData,
        lastUpdated: new Date(),
        submitted: true
      };
      
      await setDoc(doc(db, 'rda-treatment-sheets', docId), finalData);
      
      // 제출 상태 업데이트
      setIsSubmitted(true);
      
      alert('✅ Treatment sheet submitted successfully!');
      
      // 성공적으로 저장된 후 화면 초기화
      setOffice('');
      setRdaName('');
      const now = new Date();
      const pacificTime = new Date(now.toLocaleString("en-US", {timeZone: "America/Los_Angeles"}));
      setDate(pacificTime.toISOString().split('T')[0]);
      setTreatmentData(() => {
        const initialData: any[] = [];
        for (let i = 0; i < 20; i++) {
          initialData.push({
            patientName: '',
            startTime: '',
            roomNumber: '',
            services: [],
            explanation: '',
            showServices: false,
            sealantDetails: []
          });
        }
        return initialData;
      });
      
    } catch (error) {
      // 프로덕션에서는 상세한 에러 정보를 노출하지 않음
      if (process.env.NODE_ENV === 'development') {
        console.error('Error saving treatment sheet:', error);
      }
      alert('❌ Failed to save treatment sheet. Please try again.');
    } finally {
      setLoading(false);
    }
  }, [office, rdaName, date, treatmentData]);

  // 폼 초기화
  const handleClear = useCallback(() => {
    if (confirm('Are you sure you want to clear all data?')) {
      setOffice('');
      setRdaName('');
      // 캘리포니아 시간대의 현재 날짜를 YYYY-MM-DD 형식으로 반환
      const now = new Date();
      const pacificDate = new Date(now.toLocaleString("en-US", {timeZone: "America/Los_Angeles"}));
      const year = pacificDate.getFullYear();
      const month = String(pacificDate.getMonth() + 1).padStart(2, '0');
      const day = String(pacificDate.getDate()).padStart(2, '0');
      setDate(`${year}-${month}-${day}`);
      setTreatmentData(() => {
        const initialData: any[] = [];
        for (let i = 0; i < 20; i++) {
          initialData.push({
            patientName: '',
            startTime: '',
            roomNumber: '',
            services: [],
            explanation: '',
            showServices: false,
            sealantDetails: []
          });
        }
        return initialData;
      });
    }
  }, []);

  // 행 추가 함수 (5개 행씩 추가)
  const handleAddRow = useCallback(() => {
    const newRows = Array.from({ length: 5 }, () => ({
      patientName: '',
      startTime: '',
      roomNumber: '',
      services: [],
      explanation: '',
      showServices: false,
      sealantDetails: []
    }));
    
    setTreatmentData(prevData => [...prevData, ...newRows]);
  }, []);

  // URL 파라미터 변경 감지 및 뒤로가기 처리
  const isInitialLoadRef = useRef(true);
  useEffect(() => {
    const handleURLChange = (isPopState = false) => {
      if (typeof window !== 'undefined') {
        const urlParams = new URLSearchParams(window.location.search);
        const urlOffice = urlParams.get('office');
        const urlRdaName = urlParams.get('rdaName');
        const urlDate = urlParams.get('date');
        const isViewMode = urlParams.get('view') === 'true';
        
        // Generate PDF 후에는 데이터를 로드하지 않음
        if (pdfGeneratedRef.current) {
          isInitialLoadRef.current = false;
          return;
        }
        
        // 초기 로드이고 view 모드이면 데이터 로드 (뒤로가기가 아닌 경우)
        if (isInitialLoadRef.current && urlOffice && urlRdaName && urlDate && isViewMode) {
          setOffice(urlOffice);
          setRdaName(urlRdaName);
          setDate(urlDate);
          isInitialLoadRef.current = false;
          return;
        }
        
        // 초기 로드이고 view 모드가 아니면 URL 파라미터만 제거 (데이터는 초기화하지 않음)
        if (isInitialLoadRef.current && !isViewMode && (urlOffice || urlRdaName || urlDate)) {
          window.history.replaceState({}, '', window.location.pathname);
          isInitialLoadRef.current = false;
          return;
        }
        
        // 뒤로가기로 인한 URL 변경인 경우에만 처리
        if (isPopState) {
          // Generate PDF 후 뒤로가기로 돌아온 경우 데이터를 로드하지 않음
          if (pdfGeneratedRef.current) {
            // URL 파라미터 제거하고 상태 리셋
            window.history.replaceState({}, '', window.location.pathname);
            setOffice('');
            setRdaName('');
            const now = new Date();
            const pacificDate = new Date(now.toLocaleString("en-US", {timeZone: "America/Los_Angeles"}));
            const year = pacificDate.getFullYear();
            const month = String(pacificDate.getMonth() + 1).padStart(2, '0');
            const day = String(pacificDate.getDate()).padStart(2, '0');
            setDate(`${year}-${month}-${day}`);
            const emptyRows = Array.from({ length: 20 }, () => ({
              patientName: '',
              startTime: '',
              roomNumber: '',
              services: [],
              explanation: '',
              showServices: false,
              sealantDetails: []
            }));
            setTreatmentData(emptyRows);
            const resetDataCopy = JSON.parse(JSON.stringify({ office: '', rdaName: '', date: '', treatmentData: emptyRows }));
            setLastSavedData(resetDataCopy);
            setIsSubmitted(false);
            return;
          }
          
          // view 모드가 아닌데 URL 파라미터가 있으면 제거 (뒤로가기로 돌아온 경우)
          if (!isViewMode && (urlOffice || urlRdaName || urlDate)) {
            // URL 파라미터 제거하고 상태 리셋
            window.history.replaceState({}, '', window.location.pathname);
            setOffice('');
            setRdaName('');
            // 캘리포니아 시간대의 현재 날짜로 리셋
            const now = new Date();
            const pacificDate = new Date(now.toLocaleString("en-US", {timeZone: "America/Los_Angeles"}));
            const year = pacificDate.getFullYear();
            const month = String(pacificDate.getMonth() + 1).padStart(2, '0');
            const day = String(pacificDate.getDate()).padStart(2, '0');
            setDate(`${year}-${month}-${day}`);
            
            // treatmentData도 초기화
            const emptyRows = Array.from({ length: 20 }, () => ({
              patientName: '',
              startTime: '',
              roomNumber: '',
              services: [],
              explanation: '',
              showServices: false,
              sealantDetails: []
            }));
            setTreatmentData(emptyRows);
            const resetDataCopy = JSON.parse(JSON.stringify({ office: '', rdaName: '', date: '', treatmentData: emptyRows }));
            setLastSavedData(resetDataCopy);
            setIsSubmitted(false);
            return;
          }
          
          // view 모드이고 URL 파라미터가 있으면 데이터 로드 (뒤로가기로 돌아온 경우)
          if (urlOffice && urlRdaName && urlDate && isViewMode) {
            setOffice(urlOffice);
            setRdaName(urlRdaName);
            setDate(urlDate);
          }
        }
      }
    };

    // 초기 로드 시 URL 파라미터 확인
    handleURLChange(false);
    isInitialLoadRef.current = false;
    
    // URL 변경 감지를 위한 이벤트 리스너 (뒤로가기만 감지)
    const handlePopState = () => {
      handleURLChange(true);
    };
    window.addEventListener('popstate', handlePopState);
    
    return () => {
      window.removeEventListener('popstate', handlePopState);
    };
  }, []);

  // 클라이언트 마운트 상태 설정 (Hydration 오류 방지)
  useEffect(() => {
    setIsClient(true);
  }, []);
  
  // View 모드에서는 자동으로 Unlock (Office 비밀번호 검증 건너뛰기)
  useEffect(() => {
    if (isClient && typeof window !== 'undefined') {
      const isViewMode = new URLSearchParams(window.location.search).get('view') === 'true';
      if (isViewMode && office && rdaName && date && !isUnlocked) {
        // View 모드에서는 Office 비밀번호 검증 없이 자동으로 잠금 해제
        if (!isOfficeUnlocked && office) {
          setUnlockedOffices(prev => new Set(prev).add(office));
        }
        // 자동으로 Unlock
        setIsUnlocked(true);
        setUnlockedCombination({ office, rdaName, date });
        previousCombinationRef.current = { office, rdaName, date };
        loadData();
      }
    }
  }, [isClient, office, rdaName, date, isOfficeUnlocked, isUnlocked, loadData]);

  // 보안 조치 활성화
  useEffect(() => {
    enableAllSecurityMeasures({
      disableConsole: true,
      disableRightClick: true,
      disableShortcuts: true,
      disableCopy: false,
      disableSelection: false,
      monitorDevTools: true
    });
  }, []);

  // 스타일 정의
  const bodyStyle: React.CSSProperties = {
    fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
    backgroundColor: '#f8f9fa',
    margin: 0,
    padding: '20px',
    minHeight: '100vh'
  };

  const containerStyle: React.CSSProperties = {
    maxWidth: '95%',
    margin: '0 auto',
    backgroundColor: 'white',
    borderRadius: '8px',
    boxShadow: '0 2px 10px rgba(0,0,0,0.1)',
    padding: '30px'
  };

  const headerStyle: React.CSSProperties = {
    color: '#2c3e50',
    textAlign: 'center',
    marginBottom: '30px',
    fontSize: '2.5em',
    fontWeight: 'bold',
    borderBottom: '3px solid #3498db',
    paddingBottom: '15px'
  };

  const sectionStyle: React.CSSProperties = {
    backgroundColor: 'white',
    padding: '25px',
    borderRadius: '8px',
    boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
    marginBottom: '20px',
    borderWidth: '1px',
    borderStyle: 'solid',
    borderColor: '#e9ecef'
  };

  const inputStyle: React.CSSProperties = {
    padding: '8px 12px',
    borderWidth: '1px',
    borderStyle: 'solid',
    borderColor: '#BDE0FE',
    borderRadius: '4px',
    fontSize: '1em',
    backgroundColor: 'white',
    color: '#023047',
    width: '100%'
  };

  const buttonStyle: React.CSSProperties = {
    backgroundColor: '#0077B6',
    color: 'white',
    borderWidth: '0',
    borderStyle: 'none',
    borderColor: 'transparent',
    padding: '12px 24px',
    borderRadius: '6px',
    fontSize: '1em',
    fontWeight: 'bold',
    cursor: 'pointer',
    margin: '5px',
    transition: 'all 0.3s ease'
  };

  const tableStyle: React.CSSProperties = {
    width: '100%',
    borderCollapse: 'collapse',
    backgroundColor: 'white',
    boxShadow: '0 1px 3px rgba(0,0,0,0.1)',
    fontSize: '15px'
  };

  const excelInputStyle: React.CSSProperties = {
    width: '100%',
    padding: '4px 6px',
    borderWidth: '0',
    borderStyle: 'none',
    borderColor: 'transparent',
    backgroundColor: 'transparent',
    fontSize: '14px',
    fontFamily: 'inherit',
    outline: 'none',
    boxSizing: 'border-box',
    height: '32px'
  };

  return (
    <div style={bodyStyle}>
      {/* 숫자 입력 필드의 화살표(spinner) 제거 */}
      <style>{`
        input.no-spinner::-webkit-inner-spin-button,
        input.no-spinner::-webkit-outer-spin-button {
          -webkit-appearance: none;
          margin: 0;
        }
        input.no-spinner {
          -moz-appearance: textfield;
        }
      `}</style>
      {/* 자동 저장 상태 표시 */}
      {autoSaveStatus && (
        <div style={{
          position: 'fixed',
          top: '20px',
          right: '20px',
          padding: '12px 20px',
          backgroundColor: autoSaveStatus.includes('❌') ? '#ff6b6b' : 
                          autoSaveStatus.includes('🔄') ? '#4a90e2' : 
                          autoSaveStatus.includes('💾') ? '#51cf66' : '#51cf66',
          color: 'white',
          borderRadius: '25px',
          fontSize: '14px',
          fontWeight: 'bold',
          boxShadow: '0 4px 12px rgba(0,0,0,0.3)',
          zIndex: 1000,
          maxWidth: '300px',
          textAlign: 'center',
          backdropFilter: 'blur(10px)',
          border: '1px solid rgba(255, 255, 255, 0.2)'
        }}>
          {autoSaveStatus}
        </div>
      )}

      {/* Office 비밀번호 입력 모달 */}
      {showPasswordModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          backgroundColor: 'rgba(0, 0, 0, 0.5)',
          display: 'flex',
          justifyContent: 'center',
          alignItems: 'center',
          zIndex: 2000
        }}
        onClick={handlePasswordModalClose}
        >
          <div style={{
            backgroundColor: 'white',
            padding: '30px',
            borderRadius: '8px',
            boxShadow: '0 4px 20px rgba(0,0,0,0.3)',
            maxWidth: '300px',
            width: '90%',
            zIndex: 2001
          }}
          onClick={(e) => e.stopPropagation()}
          >
            <h2 style={{
              marginTop: 0,
              marginBottom: '20px',
              color: '#2c3e50',
              fontSize: '20px',
              textAlign: 'center',
              fontWeight: 'bold'
            }}>
              Password
            </h2>
            <input
              type="text"
              value={passwordInput}
              onChange={(e) => setPasswordInput(e.target.value)}
              onKeyPress={(e) => {
                if (e.key === 'Enter') {
                  handlePasswordSubmit();
                }
              }}
              placeholder="Enter password"
              style={{
                ...inputStyle,
                marginBottom: '20px',
                fontSize: '18px',
                textAlign: 'center',
                letterSpacing: '4px',
                textTransform: 'uppercase'
              }}
              autoFocus
            />
            <div style={{
              display: 'flex',
              gap: '10px',
              justifyContent: 'center'
            }}>
              <button
                onClick={handlePasswordSubmit}
                style={{
                  ...buttonStyle,
                  backgroundColor: '#0077B6',
                  minWidth: '100px'
                }}
              >
                Submit
              </button>
              <button
                onClick={handlePasswordModalClose}
                style={{
                  ...buttonStyle,
                  backgroundColor: '#6c757d',
                  minWidth: '100px'
                }}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      <div style={containerStyle}>
        {/* 헤더 */}
        <h1 style={headerStyle}>🦷 RDA/DA Treatment</h1>

        {/* 기본 정보 섹션 */}
        <div style={sectionStyle}>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '20px' }}>
            <div>
              <label style={{ display: 'block', marginBottom: '8px', fontWeight: 'bold', color: '#495057' }}>
                📅 Date:
              </label>
              <input
                type="date"
                value={date}
                onChange={(e) => handleDateChange(e.target.value)}
                style={inputStyle}
              />
            </div>
            <div>
              <label style={{ display: 'block', marginBottom: '8px', fontWeight: 'bold', color: '#495057' }}>
                🏢 Office:
              </label>
              <select
                value={office}
                onChange={(e) => handleOfficeChange(e.target.value)}
                style={inputStyle}
              >
                <option value="">Select Office</option>
                <option value="Bernard">Bernard</option>
                <option value="California">California</option>
                <option value="Delano">Delano</option>
                <option value="Fresno">Fresno</option>
                <option value="Ming">Ming</option>
                <option value="Ortho">Ortho</option>
                <option value="Tulare">Tulare</option>
                <option value="Visalia">Visalia</option>
              </select>
            </div>
            <div>
              <label style={{ display: 'block', marginBottom: '8px', fontWeight: 'bold', color: '#495057' }}>
                👨‍⚕️ RDA/DA Name:
              </label>
              <input
                type="text"
                value={rdaName}
                onChange={(e) => handleRdaNameChange(e.target.value)}
                placeholder="Enter RDA/DA name"
                style={inputStyle}
                maxLength={50}
              />
            </div>
          </div>
          {/* Unlock 버튼 */}
          {office && rdaName && date && isOfficeUnlocked && !isUnlocked && (
            <div style={{ textAlign: 'center', marginTop: '20px' }}>
              <button
                onClick={handleUnlock}
                style={{
                  ...buttonStyle,
                  backgroundColor: '#28a745',
                  fontSize: '18px',
                  padding: '16px 32px',
                  fontWeight: 'bold',
                  boxShadow: '0 4px 8px rgba(0,0,0,0.1)'
                }}
              >
                🔓 Unlock
              </button>
            </div>
          )}
        </div>

        {/* 치료 기록 테이블 (Unlock 시에만 표시) */}
        {isUnlocked && (
        <div style={sectionStyle}>
          <div style={{ overflowX: 'auto', maxHeight: '1000px', overflowY: 'auto' }}>
            <table style={tableStyle}>
              <thead>
                <tr style={{ backgroundColor: '#2c3e50', color: 'white' }}>
                  <th style={{ 
                    padding: '12px', 
                    textAlign: 'center', 
                    minWidth: '60px', 
                    border: '1px solid #d0d0d0', 
                    fontWeight: 'bold',
                    position: 'sticky',
                    top: 0,
                    backgroundColor: '#2c3e50',
                    zIndex: 10
                  }}>
                    #
                  </th>
                  <th style={{ 
                    padding: '12px', 
                    textAlign: 'center', 
                    minWidth: '300px', 
                    border: '1px solid #d0d0d0', 
                    fontWeight: 'bold',
                    position: 'sticky',
                    top: 0,
                    backgroundColor: '#2c3e50',
                    zIndex: 10
                  }}>
                    Patient's Name
                  </th>
                  <th style={{ 
                    padding: '12px', 
                    textAlign: 'center', 
                    minWidth: '150px', 
                    border: '1px solid #d0d0d0', 
                    fontWeight: 'bold',
                    position: 'sticky',
                    top: 0,
                    backgroundColor: '#2c3e50',
                    zIndex: 10
                  }}>
                    Start Time
                  </th>
                  <th style={{ 
                    padding: '12px', 
                    textAlign: 'center', 
                    minWidth: '150px', 
                    border: '1px solid #d0d0d0', 
                    fontWeight: 'bold',
                    position: 'sticky',
                    top: 0,
                    backgroundColor: '#2c3e50',
                    zIndex: 10
                  }}>
                    Room #
                  </th>
                  <th style={{ 
                    padding: '12px', 
                    textAlign: 'center', 
                    minWidth: '400px', 
                    border: '1px solid #d0d0d0', 
                    fontWeight: 'bold',
                    position: 'sticky',
                    top: 0,
                    backgroundColor: '#2c3e50',
                    zIndex: 10
                  }}>
                    Treatment or Services Performed
                  </th>
                  <th style={{ 
                    padding: '12px', 
                    textAlign: 'center', 
                    minWidth: '300px', 
                    border: '1px solid #d0d0d0', 
                    fontWeight: 'bold',
                    position: 'sticky',
                    top: 0,
                    backgroundColor: '#2c3e50',
                    zIndex: 10
                  }}>
                    Explanation of Treatment/Services or Amount Performed
                  </th>
                </tr>
              </thead>
              <tbody>
                {treatmentData.map((treatment, index) => (
                  <TreatmentRow
                    key={index}
                    rowIndex={index}
                    treatmentData={treatment}
                    updateTreatment={updateTreatment}
                    updateSealantDetail={updateSealantDetail}
                    addSealantDetail={addSealantDetail}
                    inputStyle={excelInputStyle}
                    addingSealant={addingSealant}
                    isInputDisabled={isInputDisabled}
                    isViewMode={isClient && typeof window !== 'undefined' && new URLSearchParams(window.location.search).get('view') === 'true'}
                  />
                ))}
              </tbody>
            </table>
          </div>
          
          {/* Add Row 버튼 (view mode가 아닐 때만 표시) */}
          {isClient && typeof window !== 'undefined' && new URLSearchParams(window.location.search).get('view') !== 'true' && (
            <div style={{ textAlign: 'center', padding: '20px 0' }}>
              <button
                onClick={handleAddRow}
                style={{
                  ...buttonStyle,
                  backgroundColor: '#17a2b8',
                  fontSize: '16px',
                  padding: '12px 24px',
                  fontWeight: 'bold',
                  boxShadow: '0 2px 4px rgba(0,0,0,0.1)'
                }}
              >
                ➕ Add 5 Rows
              </button>
            </div>
          )}
        </div>
        )}

        {/* 액션 버튼들 (Unlock 시에만 표시) */}
        {isUnlocked && (
        <div style={{ textAlign: 'center', padding: '30px' }}>
          {/* View Details로 접근한 경우 PDF 버튼 표시 */}
          {isClient && typeof window !== 'undefined' && new URLSearchParams(window.location.search).get('view') === 'true' ? (
            <button
              onClick={async () => {
                // 바로 PDF 생성
                setLoading(true);
                setAutoSaveStatus('Generating PDF...');

                try {
                  // PDF 생성 API 호출
                  const response = await fetch('/api/generate-rda-pdf', {
                    method: 'POST',
                    headers: {
                      'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({
                      office,
                      rdaName,
                      date,
                      treatmentData
                    }),
                  });

                  if (response.ok) {
                    setAutoSaveStatus('Processing PDF...');
                    const blob = await response.blob();
                    
                    // PDF를 Firebase Storage에 저장
                    setAutoSaveStatus('Saving PDF to archive...');
                    try {
                      const storage = getStorage();
                      const filename = `8) ${date}_${office}_${rdaName}_RDA/DA Treatment(Sealant) Sheet.pdf`;
                      const storageRef = ref(storage, `endofday-pdfs/${office}/${date}/${filename}`);
                      
                      // PDF 업로드
                      await uploadBytes(storageRef, blob);
                      
                      // 다운로드 URL 가져오기
                      const downloadUrl = await getDownloadURL(storageRef);
                      
                      // Firestore에 메타데이터 저장
                      await setDoc(doc(db, 'pdf-documents', `${date}_${office}_${rdaName}_rda-treatment_${Date.now()}`), {
                        filename,
                        office: office,
                        date: date,
                        type: 'RDA Treatment Sheet',
                        url: downloadUrl,
                        storagePath: `endofday-pdfs/${office}/${date}/${filename}`,
                        createdAt: new Date(),
                      });
                      
                      setAutoSaveStatus('✅ PDF saved to archive successfully!');
                    } catch (storageError: any) {
                      // 프로덕션에서는 상세한 에러 정보를 노출하지 않음
                      if (process.env.NODE_ENV === 'development') {
                        console.error('Storage error:', storageError);
                      }
                      alert('PDF 저장 중 오류가 발생했습니다.');
                      setAutoSaveStatus('❌ PDF 저장 실패');
                    }
                    
                    // PDF 생성 성공 후 데이터베이스에서 삭제
                    try {
                      setAutoSaveStatus('Deleting data from database...');
                      const docId = createSafeDocId(date, office, rdaName);
                      await deleteDoc(doc(db, 'rda-treatment-sheets', docId));
                      setAutoSaveStatus('✅ PDF generated and saved to archive successfully!');
                    } catch (deleteError) {
                      if (process.env.NODE_ENV === 'development') {
                        console.error('Error deleting document:', deleteError);
                      }
                      setAutoSaveStatus('⚠️ PDF generated but failed to delete data from database');
                    }
                    
                    // Generate PDF 완료 플래그 설정
                    pdfGeneratedRef.current = true;
                    
                    // 모든 상태 완전히 리셋
                    setOffice('');
                    setRdaName('');
                    const now = new Date();
                    const pacificTime = new Date(now.toLocaleString("en-US", {timeZone: "America/Los_Angeles"}));
                    setDate(pacificTime.toISOString().split('T')[0]);
                    const emptyRows = Array.from({ length: 20 }, () => ({
                      patientName: '',
                      startTime: '',
                      roomNumber: '',
                      services: [],
                      explanation: '',
                      showServices: false,
                      sealantDetails: []
                    }));
                    setTreatmentData(emptyRows);
                    const resetDataCopy = JSON.parse(JSON.stringify({ office: '', rdaName: '', date: '', treatmentData: emptyRows }));
                    setLastSavedData(resetDataCopy);
                    setIsSubmitted(false);
                    previousCombinationRef.current = null;
                    previousTreatmentDataRef.current = [];
                    
                    setTimeout(() => {
                      setLoading(false);
                      setAutoSaveStatus('');
                      // PDF 생성 및 데이터 삭제 완료 후 view 페이지로 리다이렉트 (히스토리 교체)
                      // 완전히 새로운 히스토리 항목을 만들어서 뒤로가기 방지
                      window.history.pushState(null, '', '/dashboard/forms/rda-treatment-sheet/view');
                      window.location.replace('/dashboard/forms/rda-treatment-sheet/view');
                    }, 3000);
                  } else {
                    let errorMessage = 'PDF generation failed';
                    try {
                      const errorData = await response.json();
                      errorMessage = errorData.error || errorMessage;
                    } catch (e) {
                      errorMessage = `HTTP ${response.status}: ${response.statusText}`;
                    }
                    throw new Error(errorMessage);
                  }

                } catch (error: any) {
                  // 프로덕션에서는 상세한 에러 정보를 노출하지 않음
                  if (process.env.NODE_ENV === 'development') {
                    console.error('PDF generation error:', error);
                  }
                  setAutoSaveStatus('❌ PDF generation failed');
                  setTimeout(() => {
                    setLoading(false);
                    setAutoSaveStatus('');
                  }, 3000);
                }
              }}
              disabled={loading}
              style={{
                ...buttonStyle,
                backgroundColor: '#6f42c1',
                fontSize: '18px',
                padding: '16px 32px',
                fontWeight: 'bold',
                boxShadow: '0 4px 8px rgba(0,0,0,0.1)',
                marginRight: '20px'
              }}
            >
              {loading ? '⏳ Generating PDF...' : '📄 Generate PDF'}
            </button>
          ) : (
            <>
              <button
                onClick={handleSave}
                disabled={loading}
                style={{
                  ...buttonStyle,
                  backgroundColor: '#28a745',
                  fontSize: '18px',
                  padding: '16px 32px',
                  fontWeight: 'bold',
                  boxShadow: '0 4px 8px rgba(0,0,0,0.1)',
                  marginRight: '20px'
                }}
              >
                {loading ? '⏳ Submitting...' : '📤 Submit'}
              </button>
              <button
                onClick={handleClear}
                disabled={loading}
                style={{
                  ...buttonStyle,
                  backgroundColor: '#dc3545',
                  fontSize: '18px',
                  padding: '16px 32px',
                  fontWeight: 'bold',
                  boxShadow: '0 4px 8px rgba(0,0,0,0.1)'
                }}
              >
                🗑️ Clear All
              </button>
            </>
          )}
        </div>
        )}
      </div>
    </div>
  );
}
