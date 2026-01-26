'use client'

import React, { useState, useEffect, useCallback, useRef } from "react";
import { db, auth } from '@/lib/firebase.config';
import { doc, setDoc, getDoc, deleteDoc } from 'firebase/firestore';
import { onAuthStateChanged } from 'firebase/auth';
import { getStorage, ref, uploadBytes, getDownloadURL } from 'firebase/storage';
import { pdf, Document, Page, View, Text, StyleSheet } from '@react-pdf/renderer';

// 서비스 목록 상수
const SERVICE_OPTIONS = [
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
];

// 🔒 보안: 입력 검증 함수
function safeStr(v: unknown, max: number): string {
  if (v == null) return '';
  return String(v).trim().slice(0, max).replace(/[<>]/g, '');
}

// 🔒 보안: 데이터 저장 전 검증 및 sanitization 함수
function sanitizeDataForFirebase(data: any): any {
  if (data === null || data === undefined) {
    return null;
  }
  
  // 원시 타입 검증
  if (typeof data === 'string') {
    // 문자열 길이 제한 (최대 10000자)
    const sanitized = data.trim().substring(0, 10000);
    // 위험한 문자 제거 (XSS 방지)
    return sanitized.replace(/[<>\"']/g, '');
  }
  
  if (typeof data === 'number') {
    // 숫자 범위 검증 (안전한 범위 내)
    if (!isFinite(data) || data > Number.MAX_SAFE_INTEGER || data < Number.MIN_SAFE_INTEGER) {
      return 0;
    }
    return data;
  }
  
  if (typeof data === 'boolean') {
    return data;
  }
  
  // 배열 처리
  if (Array.isArray(data)) {
    // 배열 크기 제한 (최대 10000개)
    const limitedArray = data.slice(0, 10000);
    return limitedArray.map(item => sanitizeDataForFirebase(item));
  }
  
  // 객체 처리
  if (typeof data === 'object') {
    const sanitized: any = {};
    const keys = Object.keys(data);
    // 객체 키 개수 제한 (최대 1000개)
    const limitedKeys = keys.slice(0, 1000);
    
    for (const key of limitedKeys) {
      // 키 이름 검증 (알파벳, 숫자, 언더스코어만 허용, 최대 100자)
      const safeKey = key.replace(/[^A-Za-z0-9_]/g, '').substring(0, 100);
      if (safeKey) {
        sanitized[safeKey] = sanitizeDataForFirebase(data[key]);
      }
    }
    return sanitized;
  }
  
  return null;
}

// 🔒 보안: 금액 필드 검증 (숫자만 허용)
function sanitizeAmount(value: string): string {
  if (!value || typeof value !== 'string') return '';
  // 숫자, 소수점, 음수 기호만 허용
  const sanitized = value.replace(/[^0-9.-]/g, '');
  // 최대 길이 제한
  return sanitized.substring(0, 20);
}

// 24시간제를 12시간제로 변환하는 함수
const formatTime12Hour = (time24: string): string => {
  if (!time24 || typeof time24 !== 'string' || time24.trim() === '') return '';
  
  // 입력 길이 제한
  const limitedTime = time24.trim().substring(0, 10);
  
  // 시간 형식이 HH:MM 또는 H:MM인지 확인
  const timeMatch = limitedTime.match(/^(\d{1,2}):(\d{2})$/);
  if (!timeMatch) return '';
  
  const hours = parseInt(timeMatch[1], 10);
  const minutes = timeMatch[2];
  
  // 시간 범위 검증
  if (hours < 0 || hours > 23 || parseInt(minutes, 10) < 0 || parseInt(minutes, 10) > 59) {
    return '';
  }
  
  if (hours === 0) {
    return `12:${minutes} AM`;
  } else if (hours < 12) {
    return `${hours}:${minutes} AM`;
  } else if (hours === 12) {
    return `12:${minutes} PM`;
  } else {
    return `${hours - 12}:${minutes} PM`;
  }
};

// 날짜 포맷팅
const formatDate = (dateString: string) => {
  if (!dateString || typeof dateString !== 'string') return '';
  
  // 날짜 형식 검증
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateString)) return '';
  
  const [year, month, day] = dateString.split('-').map(Number);
  
  // 날짜 범위 검증
  if (isNaN(year) || isNaN(month) || isNaN(day) || 
      year < 1900 || year > 2100 || 
      month < 1 || month > 12 || 
      day < 1 || day > 31) {
    return '';
  }
  
  const date = new Date(year, month - 1, day);
  
  // 유효한 날짜인지 확인
  if (isNaN(date.getTime())) return '';
  
  return date.toLocaleDateString('en-US', {
    timeZone: 'America/Los_Angeles',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  });
};

// RDA Treatment PDF 생성 함수
function createRDATreatmentPDFDocument(props: {
  safeDate: string;
  safeOffice: string;
  safeRdaName: string;
  filteredData: any[];
  allSealantDetails: any[];
  serviceCounts: Array<[string, number]>;
  generatedDate: string;
}) {
  const { safeDate, safeOffice, safeRdaName, filteredData, allSealantDetails, serviceCounts, generatedDate } = props;

  const pdfStyles = StyleSheet.create({
    page: {
      padding: 20,
      fontFamily: 'Helvetica',
      fontSize: 8
    },
    header: {
      marginBottom: 15,
      borderBottomWidth: 2,
      borderColor: '#000',
      paddingBottom: 8,
      alignItems: 'center'
    },
    headerTitle: {
      fontSize: 16,
      fontWeight: 'bold',
      marginBottom: 4
    },
    infoSection: {
      marginBottom: 15,
      padding: 8,
      borderWidth: 1,
      borderColor: '#ccc',
      flexDirection: 'row',
      justifyContent: 'space-between',
      flexWrap: 'wrap',
      backgroundColor: '#f8f9fa'
    },
    infoItem: {
      fontSize: 9,
      marginBottom: 4
    },
    serviceCounts: {
      marginBottom: 12,
      fontSize: 8,
      flexDirection: 'row',
      flexWrap: 'wrap'
    },
    serviceCountItem: {
      marginRight: 8,
      marginBottom: 2,
      padding: 2,
      backgroundColor: '#f0f0f0',
      borderRadius: 3
    },
    table: {
      marginTop: 10
    },
    tableHeader: {
      flexDirection: 'row',
      borderBottomWidth: 1,
      borderColor: '#000',
      backgroundColor: '#f8f9fa',
      fontWeight: 'bold'
    },
    tableRow: {
      flexDirection: 'row',
      borderBottomWidth: 0.5,
      borderColor: '#000'
    },
    tableCell: {
      padding: 4,
      fontSize: 7,
      borderRightWidth: 0.5,
      borderColor: '#000',
      justifyContent: 'flex-start',
      alignItems: 'flex-start'
    },
    tableCellChart: {
      padding: 4,
      fontSize: 7,
      flex: 0.09,
      borderRightWidth: 0.5,
      borderColor: '#000',
      justifyContent: 'flex-start',
      alignItems: 'flex-start'
    },
    tableCellDob: {
      padding: 4,
      fontSize: 7,
      flex: 0.09,
      borderRightWidth: 0.5,
      borderColor: '#000',
      justifyContent: 'flex-start',
      alignItems: 'flex-start'
    },
    tableCellTooth: {
      padding: 4,
      fontSize: 7,
      flex: 0.08,
      borderRightWidth: 0.5,
      borderColor: '#000',
      justifyContent: 'flex-start',
      alignItems: 'flex-start'
    },
    tableCellRedo: {
      padding: 4,
      fontSize: 7,
      flex: 0.08,
      borderRightWidth: 0.5,
      borderColor: '#000',
      justifyContent: 'flex-start',
      alignItems: 'flex-start'
    },
    tableCellAcct: {
      padding: 4,
      fontSize: 7,
      flex: 0.09,
      borderRightWidth: 0.5,
      borderColor: '#000',
      justifyContent: 'flex-start',
      alignItems: 'flex-start'
    },
    tableCellPayable: {
      padding: 4,
      fontSize: 7,
      flex: 0.09,
      borderRightWidth: 0.5,
      borderColor: '#000',
      justifyContent: 'flex-start',
      alignItems: 'flex-start'
    },
    tableCellDxDr: {
      padding: 4,
      fontSize: 7,
      flex: 0.10,
      borderRightWidth: 0.5,
      borderColor: '#000',
      justifyContent: 'flex-start',
      alignItems: 'flex-start'
    },
    tableCellDrAmount: {
      padding: 4,
      fontSize: 7,
      flex: 0.09,
      borderRightWidth: 0.5,
      borderColor: '#000',
      justifyContent: 'flex-start',
      alignItems: 'flex-start'
    },
    tableCellRdaAmount: {
      padding: 4,
      fontSize: 7,
      flex: 0.09,
      borderRightWidth: 0.5,
      borderColor: '#000',
      justifyContent: 'flex-start',
      alignItems: 'flex-start'
    },
    tableCellNo: {
      padding: 4,
      fontSize: 8,
      flex: 0.05,
      borderRightWidth: 0.5,
      borderColor: '#000',
      justifyContent: 'center',
      alignItems: 'center',
      fontWeight: 'bold'
    },
    tableCellName: {
      padding: 4,
      fontSize: 7,
      flex: 0.12,
      borderRightWidth: 0.5,
      borderColor: '#000',
      justifyContent: 'flex-start',
      alignItems: 'flex-start'
    },
    tableCellTime: {
      padding: 4,
      fontSize: 7,
      flex: 0.10,
      borderRightWidth: 0.5,
      borderColor: '#000',
      justifyContent: 'flex-start',
      alignItems: 'flex-start'
    },
    tableCellRoom: {
      padding: 4,
      fontSize: 7,
      flex: 0.10,
      borderRightWidth: 0.5,
      borderColor: '#000',
      justifyContent: 'flex-start',
      alignItems: 'flex-start'
    },
    tableCellServices: {
      padding: 4,
      fontSize: 6,
      flex: 0.28,
      borderRightWidth: 0.5,
      borderColor: '#000',
      justifyContent: 'flex-start',
      alignItems: 'flex-start'
    },
    tableCellExplanation: {
      padding: 4,
      fontSize: 6,
      flex: 0.35,
      borderRightWidth: 0.5,
      borderColor: '#000',
      justifyContent: 'flex-start',
      alignItems: 'flex-start'
    },
    sealantSection: {
      marginTop: 20
    },
    sealantTitle: {
      fontSize: 10,
      fontWeight: 'bold',
      marginBottom: 10,
      borderBottomWidth: 1,
      borderColor: '#000',
      paddingBottom: 4
    },
    footer: {
      fontSize: 7,
      color: '#666',
      marginTop: 15,
      textAlign: 'center',
      borderTopWidth: 1,
      borderColor: '#ccc',
      paddingTop: 8
    }
  });

  const s = pdfStyles;
  const formattedDate = formatDate(safeDate);

  // 헤더
  const header = React.createElement(View, { style: s.header },
    React.createElement(Text, { style: s.headerTitle }, 'RDA/DA Treatment(Sealant)')
  );

  // 정보 섹션
  const infoSection = React.createElement(View, { style: s.infoSection },
    React.createElement(Text, { style: s.infoItem }, `Date: ${formattedDate}`),
    React.createElement(Text, { style: s.infoItem }, `Office: ${safeOffice}`),
    React.createElement(Text, { style: s.infoItem }, `RDA/DA Name: ${safeRdaName}`)
  );

  // 서비스 카운트
  const serviceCountsSection = serviceCounts.length > 0
    ? React.createElement(View, { style: s.serviceCounts },
        ...serviceCounts.map(([service, count]) =>
          React.createElement(View, { key: service, style: s.serviceCountItem },
            React.createElement(Text, null, `${safeStr(service, 100)}: ${count}`)
          )
        )
      )
    : null;

  // 메인 테이블 헤더
  const mainTableHeader = React.createElement(View, { style: s.tableHeader },
    React.createElement(View, { style: s.tableCellNo }, React.createElement(Text, { style: { fontWeight: 'bold' } }, '#')),
    React.createElement(View, { style: s.tableCellName }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'PT Name')),
    React.createElement(View, { style: s.tableCellTime }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Start Time')),
    React.createElement(View, { style: s.tableCellRoom }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Room #')),
    React.createElement(View, { style: s.tableCellServices }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Treatment or Services Performed')),
    React.createElement(View, { style: s.tableCellExplanation }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Explanation of Treatment/Services or Amount Performed'))
  );

  // 메인 테이블 행
  const mainTableRows = filteredData.map((row, index) => {
    const safePatientName = safeStr(row.patientName, 100);
    const safeStartTime = formatTime12Hour(String(row.startTime || ''));
    const safeRoomNumber = safeStr(row.roomNumber, 10);
    const safeExplanation = safeStr(row.explanation, 500);

    // 서비스 목록
    const servicesList = row.services && Array.isArray(row.services) && row.services.length > 0
      ? row.services.slice(0, 50).filter((service: any) => 
          typeof service === 'string' && SERVICE_OPTIONS.includes(service)
        ).map((service: string, idx: number) =>
          React.createElement(Text, { key: idx, style: { marginBottom: 2 } }, `• ${safeStr(service, 200)}`)
        )
      : [React.createElement(Text, { key: 0 }, '-')];

    return React.createElement(View, { key: index, style: s.tableRow },
      React.createElement(View, { style: s.tableCellNo }, React.createElement(Text, null, String(index + 1))),
      React.createElement(View, { style: s.tableCellName }, React.createElement(Text, null, safePatientName || '-')),
      React.createElement(View, { style: s.tableCellTime }, React.createElement(Text, null, safeStartTime || '-')),
      React.createElement(View, { style: s.tableCellRoom }, React.createElement(Text, null, safeRoomNumber || '-')),
      React.createElement(View, { style: s.tableCellServices }, ...servicesList),
      React.createElement(View, { style: s.tableCellExplanation }, React.createElement(Text, null, safeExplanation || '-'))
    );
  });

  // 메인 테이블
  const mainTable = React.createElement(View, { style: s.table },
    mainTableHeader,
    ...mainTableRows
  );

  // Sealant 테이블 헤더
  const sealantTableHeader = allSealantDetails.length > 0
    ? React.createElement(View, { style: s.tableHeader },
        React.createElement(View, { style: s.tableCellNo }, React.createElement(Text, { style: { fontWeight: 'bold' } }, '#')),
        React.createElement(View, { style: s.tableCellName }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'PT Name')),
        React.createElement(View, { style: s.tableCellChart }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Chart #')),
        React.createElement(View, { style: s.tableCellDob }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'DOB')),
        React.createElement(View, { style: s.tableCellTooth }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Tooth #')),
        React.createElement(View, { style: s.tableCellRedo }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Redo')),
        React.createElement(View, { style: s.tableCellAcct }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Acct Type')),
        React.createElement(View, { style: s.tableCellPayable }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'Payable')),
        React.createElement(View, { style: s.tableCellDxDr }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'DX Dr.')),
        React.createElement(View, { style: s.tableCellDrAmount }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'DR $')),
        React.createElement(View, { style: s.tableCellRdaAmount }, React.createElement(Text, { style: { fontWeight: 'bold' } }, 'RDA $'))
      )
    : null;

  // Sealant 테이블 행
  const sealantTableRows = allSealantDetails.map((item, index) => {
    const safePatientName = safeStr(item.patientName, 100);
    const safeChartNumber = safeStr(item.detail.chartNumber, 50);
    const safeDob = safeStr(item.detail.dob, 20);
    const safeToothNumber = safeStr(item.detail.toothNumber, 10);
    const safeRedo = safeStr(item.detail.redo, 10);
    const safeAcctType = safeStr(item.detail.acctType, 10);
    const safePayable = safeStr(item.detail.payable, 10);
    const safeDxDr = safeStr(item.detail.dxDr, 50);
    const safeDrAmount = safeStr(item.detail.drAmount, 20);
    const safeRdaAmount = safeStr(item.detail.rdaAmount, 20);

    return React.createElement(View, { key: index, style: s.tableRow },
      React.createElement(View, { style: s.tableCellNo }, React.createElement(Text, null, String(index + 1))),
      React.createElement(View, { style: s.tableCellName }, React.createElement(Text, null, safePatientName || '-')),
      React.createElement(View, { style: s.tableCellChart }, React.createElement(Text, null, safeChartNumber || '-')),
      React.createElement(View, { style: s.tableCellDob }, React.createElement(Text, null, safeDob || '-')),
      React.createElement(View, { style: s.tableCellTooth }, React.createElement(Text, null, safeToothNumber || '-')),
      React.createElement(View, { style: s.tableCellRedo }, React.createElement(Text, null, safeRedo || '-')),
      React.createElement(View, { style: s.tableCellAcct }, React.createElement(Text, null, safeAcctType || '-')),
      React.createElement(View, { style: s.tableCellPayable }, React.createElement(Text, null, safePayable || '-')),
      React.createElement(View, { style: s.tableCellDxDr }, React.createElement(Text, null, safeDxDr || '-')),
      React.createElement(View, { style: s.tableCellDrAmount }, React.createElement(Text, null, safeDrAmount || '-')),
      React.createElement(View, { style: s.tableCellRdaAmount }, React.createElement(Text, null, safeRdaAmount || '-'))
    );
  });

  // Sealant 테이블
  const sealantTable = allSealantDetails.length > 0
    ? React.createElement(View, { style: s.sealantSection },
        React.createElement(Text, { style: s.sealantTitle }, 'Sealant Details'),
        React.createElement(View, { style: s.table },
          sealantTableHeader,
          ...sealantTableRows
        )
      )
    : null;

  // 푸터
  const footer = React.createElement(View, { style: s.footer },
    React.createElement(Text, null, `Report generated on ${generatedDate} PDT`)
  );

  return React.createElement(Document, null,
    React.createElement(Page, { 
      size: 'A4', 
      orientation: 'portrait', 
      style: s.page
    }, 
      header, 
      infoSection,
      serviceCountsSection,
      mainTable,
      sealantTable,
      footer
    )
  );
}

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
  isViewMode,
  isSealantDetailsEditable
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
  isSealantDetailsEditable: boolean;
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
            onChange={(e) => {
              const value = e.target.value;
              // 최대 길이 제한
              if (value.length <= 100) {
                updateTreatment(rowIndex, 'patientName', value);
              }
            }}
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
                {SERVICE_OPTIONS.map((service) => (
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
                        if (!isInputDisabled) {
                          const currentServices = treatmentData.services || [];
                          let newServices;
                          if (e.target.checked) {
                            // 허용된 서비스만 추가
                            if (SERVICE_OPTIONS.includes(service)) {
                              newServices = [...currentServices, service];
                            } else {
                              return; // 허용되지 않은 서비스는 무시
                            }
                          } else {
                            newServices = currentServices.filter((s: string) => s !== service);
                          }
                          updateTreatment(rowIndex, 'services', newServices);
                        }
                      }}
                      disabled={isInputDisabled}
                      style={{ marginRight: '6px', cursor: isInputDisabled ? 'not-allowed' : 'pointer' }}
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
            onChange={(e) => {
              const value = e.target.value;
              // 최대 길이 제한
              if (value.length <= 1000) {
                updateTreatment(rowIndex, 'explanation', value);
              }
            }}
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
                Sealant Details
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
                          style={{ ...inputStyle, border: 'none', fontSize: '12px', padding: '6px', backgroundColor: !isSealantDetailsEditable ? '#f5f5f5' : 'white', cursor: !isSealantDetailsEditable ? 'not-allowed' : 'text' }}
                          disabled={!isSealantDetailsEditable}
                        />
                      </td>
                      <td style={{ padding: '0', border: '1px solid #dee2e6' }}>
                        <input
                          type="text"
                          value={detail.chartNumber || ''}
                          onChange={(e) => updateSealantDetail(rowIndex, detailIndex, 'chartNumber', e.target.value)}
                          style={{ ...inputStyle, border: 'none', fontSize: '12px', padding: '6px', backgroundColor: !isSealantDetailsEditable ? '#f5f5f5' : 'white', cursor: !isSealantDetailsEditable ? 'not-allowed' : 'text' }}
                          disabled={!isSealantDetailsEditable}
                        />
                      </td>
                      <td style={{ padding: '0', border: '1px solid #dee2e6' }}>
                        <input
                          type="date"
                          value={detail.dob || ''}
                          onChange={(e) => updateSealantDetail(rowIndex, detailIndex, 'dob', e.target.value)}
                          style={{ ...inputStyle, border: 'none', fontSize: '12px', padding: '6px', backgroundColor: !isSealantDetailsEditable ? '#f5f5f5' : 'white', cursor: !isSealantDetailsEditable ? 'not-allowed' : 'text' }}
                          disabled={!isSealantDetailsEditable}
                        />
                      </td>
                      <td style={{ padding: '0', border: '1px solid #dee2e6' }}>
                        <select
                          value={detail.toothNumber || ''}
                          onChange={(e) => updateSealantDetail(rowIndex, detailIndex, 'toothNumber', e.target.value)}
                          style={{ ...inputStyle, border: 'none', fontSize: '12px', padding: '6px', backgroundColor: !isSealantDetailsEditable ? '#f5f5f5' : 'white', cursor: !isSealantDetailsEditable ? 'not-allowed' : 'pointer' }}
                          disabled={!isSealantDetailsEditable}
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
                          style={{ ...inputStyle, border: 'none', fontSize: '12px', padding: '6px', backgroundColor: !isSealantDetailsEditable ? '#f5f5f5' : 'white', cursor: !isSealantDetailsEditable ? 'not-allowed' : 'pointer' }}
                          disabled={!isSealantDetailsEditable}
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
                          style={{ ...inputStyle, border: 'none', fontSize: '12px', padding: '6px', backgroundColor: !isSealantDetailsEditable ? '#f5f5f5' : 'white', cursor: !isSealantDetailsEditable ? 'not-allowed' : 'pointer' }}
                          disabled={!isSealantDetailsEditable}
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
                          style={{ ...inputStyle, border: 'none', fontSize: '12px', padding: '6px', backgroundColor: !isSealantDetailsEditable ? '#f5f5f5' : 'white', cursor: !isSealantDetailsEditable ? 'not-allowed' : 'pointer' }}
                          disabled={!isSealantDetailsEditable}
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
                          style={{ ...inputStyle, border: 'none', fontSize: '12px', padding: '6px', backgroundColor: !isSealantDetailsEditable ? '#f5f5f5' : 'white', cursor: !isSealantDetailsEditable ? 'not-allowed' : 'text' }}
                          disabled={!isSealantDetailsEditable}
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
                            backgroundColor: !isSealantDetailsEditable ? '#f5f5f5' : 'white',
                            cursor: !isSealantDetailsEditable ? 'not-allowed' : 'text'
                          }}
                          disabled={!isSealantDetailsEditable}
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
                            backgroundColor: !isSealantDetailsEditable ? '#f5f5f5' : 'white',
                            cursor: !isSealantDetailsEditable ? 'not-allowed' : 'text'
                          }}
                          disabled={!isSealantDetailsEditable}
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
                  {isViewMode && (
                    <tr>
                      <td colSpan={10} style={{ padding: '8px', textAlign: 'center', border: '1px solid #dee2e6', backgroundColor: '#f8f9fa' }}>
                        <button
                          disabled={addingSealant.has(rowIndex) || !isSealantDetailsEditable}
                          onClick={(e) => {
                            if (isSealantDetailsEditable) {
                              e.preventDefault();
                              e.stopPropagation();
                              addSealantDetail(rowIndex);
                            }
                          }}
                          style={{
                            backgroundColor: (addingSealant.has(rowIndex) || !isSealantDetailsEditable) ? '#6c757d' : '#28a745',
                            color: 'white',
                            border: 'none',
                            padding: '8px 16px',
                            borderRadius: '6px',
                            fontSize: '13px',
                            cursor: (addingSealant.has(rowIndex) || !isSealantDetailsEditable) ? 'not-allowed' : 'pointer',
                            fontWeight: 'bold',
                            boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
                            transition: 'all 0.2s ease',
                            opacity: (addingSealant.has(rowIndex) || !isSealantDetailsEditable) ? 0.6 : 1
                          }}
                          onMouseOver={(e) => {
                            if (!(addingSealant.has(rowIndex) || !isSealantDetailsEditable)) {
                              e.currentTarget.style.backgroundColor = '#218838';
                              e.currentTarget.style.transform = 'translateY(-1px)';
                            }
                          }}
                          onMouseOut={(e) => {
                            if (!(addingSealant.has(rowIndex) || !isSealantDetailsEditable)) {
                              e.currentTarget.style.backgroundColor = '#28a745';
                              e.currentTarget.style.transform = 'translateY(0)';
                            }
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
  // 보안 강화: 더 안전한 세션 ID 생성 (crypto API 사용)
  const [userSessionId] = useState(() => {
    if (typeof window !== 'undefined' && window.crypto && window.crypto.getRandomValues) {
      const array = new Uint32Array(4);
      window.crypto.getRandomValues(array);
      return Array.from(array).map((val: number) => val.toString(36)).join('');
    }
    // Fallback: 더 긴 랜덤 문자열
    return Math.random().toString(36).substring(2, 15) + 
           Math.random().toString(36).substring(2, 15) + 
           Date.now().toString(36);
  });
  const [lastSavedData, setLastSavedData] = useState<any>({});
  const [isSubmitted, setIsSubmitted] = useState(false); // 제출 상태 추적
  const pdfGeneratedRef = useRef(false); // Generate PDF 버튼을 눌렀는지 추적
  
  // Rate limiting을 위한 ref
  const lastAutoSaveTimeRef = useRef<number>(0);
  const lastApiCallTimeRef = useRef<number>(0);
  const autoSaveAttemptsRef = useRef<number>(0);
  const lastUpdateTreatmentCall = useRef<number>(0);
  const lastUpdateSealantDetailCall = useRef<number>(0);
  const lastAddSealantDetailCall = useRef<number>(0);
  const lastFieldUpdateTimeRef = useRef<Map<string, number>>(new Map());
  const autoSaveTimerRef = useRef<NodeJS.Timeout | null>(null);
  const pendingAutoSaveRef = useRef<boolean>(false);
  
  // Hydration 오류 방지를 위한 클라이언트 마운트 상태
  const [isClient, setIsClient] = useState(false);
  
  // 인증 상태
  const [isAuthorized, setIsAuthorized] = useState<boolean | null>(null); // null: 확인 중, true: 인증됨, false: 인증 실패
  
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
      // URL 파라미터 검증 및 sanitization
      const urlOffice = urlParams.get('office') || '';
      const urlRdaName = urlParams.get('rdaName') || '';
      const urlDate = urlParams.get('date') || '';
      
      // Office 검증: 알파벳과 숫자만 허용, 최대 10자
      const safeOffice = urlOffice.trim().replace(/[^A-Za-z0-9]/g, '').substring(0, 10);
      
      // RDA Name 검증: 알파벳, 숫자, 공백만 허용, 최대 50자
      const safeRdaName = urlRdaName.trim().replace(/[^A-Za-z0-9\s]/g, '').substring(0, 50);
      
      // Date 검증: YYYY-MM-DD 형식만 허용 (보안 강화)
      let safeDate = urlDate.trim().replace(/[^0-9-]/g, '');
      if (safeDate && !/^\d{4}-\d{2}-\d{2}$/.test(safeDate)) {
        safeDate = '';
      }
      // 날짜 범위 검증 (1900-01-01 ~ 2100-12-31)
      if (safeDate) {
        const dateObj = new Date(safeDate + 'T00:00:00');
        const year = dateObj.getFullYear();
        if (isNaN(year) || year < 1900 || year > 2100) {
          safeDate = '';
        }
      }
      
      return {
        office: safeOffice,
        rdaName: safeRdaName,
        date: safeDate || getPacificDate()
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
  // View 모드에서는 sealant details만 수정 가능하도록 별도 처리
  const isViewModeFromUrl = isClient && typeof window !== 'undefined' && new URLSearchParams(window.location.search).get('view') === 'true';
  const baseIsInputDisabled = !office || !rdaName || !date || !isOfficeUnlocked || !isUnlocked;
  // View 모드일 때는 일반 필드는 항상 비활성화, sealant details만 수정 가능
  const isInputDisabled = isViewModeFromUrl ? true : baseIsInputDisabled;
  const isSealantDetailsEditable = isViewModeFromUrl ? true : !baseIsInputDisabled;
  
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
  
  // 비밀번호 확인 핸들러 (보안 강화)
  const handlePasswordSubmit = () => {
    // 입력값 검증 및 sanitization
    const sanitizedPassword = passwordInput.trim().toUpperCase();
    const sanitizedOffice = selectedOfficeForPassword.trim().toUpperCase();
    
    // 빈 값 체크
    if (!sanitizedPassword || !sanitizedOffice) {
      setPasswordInput('');
      return;
    }
    
    // 길이 제한 (보안 강화)
    if (sanitizedPassword.length > 50 || sanitizedOffice.length > 10) {
      setPasswordInput('');
      return;
    }
    
    // 알파벳과 숫자만 허용
    if (!/^[A-Z0-9]+$/.test(sanitizedPassword) || !/^[A-Z0-9]+$/.test(sanitizedOffice)) {
      setPasswordInput('');
      return;
    }
    
    // 비밀번호 검증 (타이밍 공격 방지를 위한 일정한 시간 소요)
    const startTime = Date.now();
    const isValid = sanitizedPassword === sanitizedOffice;
    
    // 최소 처리 시간 보장 (타이밍 공격 방지)
    const minProcessingTime = 100; // 100ms
    const processingTime = Date.now() - startTime;
    if (processingTime < minProcessingTime) {
      setTimeout(() => {
        if (isValid) {
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
      }, minProcessingTime - processingTime);
    } else {
      if (isValid) {
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
    
    // Rate limiting: 최소 2초 간격으로 저장 (서버 부하 방지)
    const now = Date.now();
    const timeSinceLastSave = now - lastAutoSaveTimeRef.current;
    if (timeSinceLastSave < 2000) {
      return; // 너무 빈번한 저장 방지
    }
    
    // 연속 저장 시도 제한 (1분에 30회 이상 시도 시 차단)
    if (timeSinceLastSave < 60000) {
      autoSaveAttemptsRef.current++;
      if (autoSaveAttemptsRef.current > 30) {
        return;
      }
    } else {
      autoSaveAttemptsRef.current = 0; // 1분 경과 시 카운터 리셋
    }
    
    lastAutoSaveTimeRef.current = now;

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
      setAutoSaveStatus('Saving...');
      
      // 캘리포니아 시간대로 현재 시간 생성
      const currentTime = new Date();
      const pacificDateTime = new Date(currentTime.toLocaleString("en-US", {timeZone: "America/Los_Angeles"}));
      
      // Sealant가 없는 행의 sealantDetails를 빈 배열로 명시적으로 설정
      // 🔒 보안: 배열 데이터 검증 및 정규화
      const normalizedTreatmentData = treatmentData
        .slice(0, 10000) // 최대 10000개 행만 허용
        .map((row: any) => {
          if (!row || typeof row !== 'object') {
            return {
              patientName: '',
              startTime: '',
              roomNumber: '',
              services: [],
              explanation: '',
              showServices: false,
              sealantDetails: []
            };
          }
          
          const hasSealant = row.services && Array.isArray(row.services) && row.services.includes('Sealant');
          
          // sealantDetails 검증 및 정규화
          let safeSealantDetails: any[] = [];
          if (hasSealant && Array.isArray(row.sealantDetails)) {
            safeSealantDetails = row.sealantDetails
              .slice(0, 100) // 최대 100개만 허용
              .map((detail: any) => {
                if (!detail || typeof detail !== 'object') {
                  return {
                    ptName: row.patientName || '',
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
                return {
                  ptName: String(detail.ptName || '').substring(0, 100).replace(/[<>\"']/g, ''),
                  chartNumber: String(detail.chartNumber || '').substring(0, 50).replace(/[<>\"']/g, ''),
                  dob: String(detail.dob || '').substring(0, 20).replace(/[<>\"']/g, ''),
                  toothNumber: String(detail.toothNumber || '').substring(0, 10).replace(/[<>\"']/g, ''),
                  redo: String(detail.redo || '').substring(0, 10).replace(/[<>\"']/g, ''),
                  acctType: String(detail.acctType || '').substring(0, 10).replace(/[<>\"']/g, ''),
                  payable: String(detail.payable || '').substring(0, 10).replace(/[<>\"']/g, ''),
                  dxDr: String(detail.dxDr || '').substring(0, 50).replace(/[<>\"']/g, ''),
                  drAmount: sanitizeAmount(String(detail.drAmount || '')),
                  rdaAmount: sanitizeAmount(String(detail.rdaAmount || ''))
                };
              });
          }
          
          return {
            patientName: String(row.patientName || '').substring(0, 100).replace(/[<>\"']/g, ''),
            startTime: String(row.startTime || '').substring(0, 10).replace(/[<>\"']/g, ''),
            roomNumber: String(row.roomNumber || '').substring(0, 10).replace(/[<>\"']/g, ''),
            services: Array.isArray(row.services) 
              ? row.services
                  .filter((s: any) => typeof s === 'string' && SERVICE_OPTIONS.includes(s))
                  .slice(0, 50)
              : [],
            explanation: String(row.explanation || '').substring(0, 500).replace(/[<>\"']/g, ''),
            showServices: Boolean(row.showServices),
            sealantDetails: safeSealantDetails
          };
        });
      
      const dataToSave = {
        office: office.trim().substring(0, 10),
        rdaName: rdaName.trim().substring(0, 50),
        date: date.trim().substring(0, 20),
        treatmentData: normalizedTreatmentData,
        timestamp: pacificDateTime.toISOString(),
        autoSaved: true,
        lastUpdatedBy: userSessionId.substring(0, 200),
        submitted: isSubmitted // 기존 제출 상태 유지 (view 모드에서 수정 시에도 submitted: true 유지)
      };

      // 🔒 보안: 데이터 저장 전 검증 및 sanitization
      const sanitizedData = sanitizeDataForFirebase(dataToSave);

      // sealant details가 포함된 데이터 확인
      const sealantRows = treatmentData.filter((row: any) => row.sealantDetails && row.sealantDetails.length > 0);

      // Office, RDA/DA Name, Date 조합으로 고유한 문서 ID 생성
      // 이 조합이 변경되면 새로운 문서로 저장됨
      await setDoc(doc(db, "rda-treatment-sheets", docId), sanitizedData);
      
      // 저장 성공 시 마지막 저장된 데이터 업데이트 (깊은 복사)
      const deepCopy = JSON.parse(JSON.stringify(currentData));
      setLastSavedData(deepCopy);
      
      setAutoSaveStatus('Saved ✅');
      setTimeout(() => setAutoSaveStatus(''), 2000);
      
    } catch (error: any) {
      // 🔒 보안: 프로덕션에서는 상세한 에러 정보를 노출하지 않음
      console.error('Auto-save error:', error);
      setAutoSaveStatus('Save failed. Please try again.');
      setTimeout(() => setAutoSaveStatus(''), 3000);
    }
  }, [office, rdaName, date, treatmentData, lastSavedData, isUpdatingFromFirebase, userSessionId, isSubmitted, isUnlocked, createSafeDocId]);

  // 데이터 변경 시에만 자동 저장 (Debounce 패턴 적용)
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
      // 기존 타이머가 있으면 취소
      if (autoSaveTimerRef.current) {
        clearTimeout(autoSaveTimerRef.current);
      }
      
      // Debounce: 사용자가 입력을 멈춘 후 1초 후에 저장
      autoSaveTimerRef.current = setTimeout(() => {
        autoSave();
        // 저장 후 pending 상태 확인하여 추가 저장 필요 시 실행
        if (pendingAutoSaveRef.current) {
          setTimeout(() => {
            pendingAutoSaveRef.current = false;
            autoSave();
          }, 1000);
        }
      }, 1000);
    }
    
    // cleanup: 컴포넌트 언마운트 시 타이머 정리
    return () => {
      if (autoSaveTimerRef.current) {
        clearTimeout(autoSaveTimerRef.current);
      }
    };
  }, [treatmentData, autoSave, office, rdaName, date, isUpdatingFromFirebase, isUnlocked, unlockedCombination]);

  // 치료 데이터 업데이트
  const updateTreatment = useCallback((rowIndex: number, field: string, value: string | string[] | boolean) => {
    // Rate limiting: 최소화 (1ms - 빠른 입력 허용하면서도 과도한 호출 방지)
    const now = Date.now();
    if (now - lastUpdateTreatmentCall.current < 1) {
      return;
    }
    lastUpdateTreatmentCall.current = now;
    
    // 개별 필드 throttle 제거 - 빠른 타이핑 시에도 모든 입력이 반영되도록 함
    
    setTreatmentData(prev => {
      // rowIndex 범위 검증 (보안 강화: 음수 및 범위 초과 방지)
      if (!Number.isInteger(rowIndex) || rowIndex < 0 || rowIndex >= prev.length || rowIndex >= 10000) {
        return prev;
      }
      
      // field 검증 (허용된 필드만 처리)
      const allowedFields = ['patientName', 'startTime', 'roomNumber', 'services', 'explanation', 'showServices'];
      if (!allowedFields.includes(field)) {
        return prev;
      }
      
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
      } else if (field === 'services') {
        // Services 배열 검증
        if (Array.isArray(value)) {
          // 보안 강화: rowIndex 범위 재검증 및 prev[rowIndex] 존재 확인
          if (!prev[rowIndex] || typeof prev[rowIndex] !== 'object') {
            return prev;
          }
          
          // 변경 전 Sealant 상태 확인 (업데이트 전에 확인)
          const hadSealant = Array.isArray(prev[rowIndex].services) && prev[rowIndex].services.includes('Sealant');
          
          // 허용된 서비스만 필터링
          const validServices = value.filter((s: any) => 
            typeof s === 'string' && SERVICE_OPTIONS.includes(s)
          );
          // 중복 제거
          const uniqueServices = Array.from(new Set(validServices));
          // 최대 50개로 제한
          const limitedServices = uniqueServices.slice(0, 50);
          
          // 변경 후 Sealant 상태 확인
          const hasSealant = limitedServices.includes('Sealant');
          
          // Sealant 상태에 따라 sealantDetails 처리
          if (hadSealant && !hasSealant) {
            // Sealant가 제거되면 sealantDetails도 완전히 삭제
            newData[rowIndex] = {
              ...newData[rowIndex],
              [field]: limitedServices,
              sealantDetails: [] // 명시적으로 빈 배열로 설정
            };
          } else if (!hadSealant && hasSealant) {
            // Sealant가 새로 추가되면 빈 배열로 초기화 (이전 데이터 복원 방지)
            newData[rowIndex] = {
              ...newData[rowIndex],
              [field]: limitedServices,
              sealantDetails: [] // 명시적으로 빈 배열로 설정
            };
          } else {
            // Sealant 상태가 변경되지 않은 경우
            newData[rowIndex] = {
              ...newData[rowIndex],
              [field]: limitedServices
            };
          }
        }
      } else if (field === 'patientName') {
        // 입력값 검증 및 sanitization
        let sanitizedValue = typeof value === 'string' ? value : String(value || '');
        // 최대 길이 제한
        sanitizedValue = sanitizedValue.substring(0, 100);
        // 위험한 문자 제거 (XSS 방지)
        sanitizedValue = sanitizedValue.replace(/[<>\"']/g, '');
        
        // Patient Name이 변경되면 Sealant Details의 모든 PT Name도 업데이트
        newData[rowIndex] = {
          ...newData[rowIndex],
          [field]: sanitizedValue
        };
        
        if (newData[rowIndex].sealantDetails) {
          newData[rowIndex].sealantDetails = newData[rowIndex].sealantDetails.map((detail: any) => ({
            ...detail,
            ptName: sanitizedValue
          }));
        }
      } else if (field === 'explanation') {
        // Explanation 필드 검증 및 sanitization
        let sanitizedValue = typeof value === 'string' ? value : String(value || '');
        // 최대 길이 제한 (500자)
        sanitizedValue = sanitizedValue.substring(0, 500);
        // 위험한 문자 제거 (XSS 방지)
        sanitizedValue = sanitizedValue.replace(/[<>\"']/g, '');
        
        newData[rowIndex] = {
          ...newData[rowIndex],
          [field]: sanitizedValue
        };
      } else if (field === 'startTime' || field === 'roomNumber') {
        // 시간 및 방 번호 필드 검증
        let sanitizedValue = typeof value === 'string' ? value : String(value || '');
        // 최대 길이 제한
        const maxLength = field === 'startTime' ? 10 : 10;
        sanitizedValue = sanitizedValue.substring(0, maxLength);
        // 위험한 문자 제거
        sanitizedValue = sanitizedValue.replace(/[<>\"']/g, '');
        
        newData[rowIndex] = {
          ...newData[rowIndex],
          [field]: sanitizedValue
        };
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
    // Rate limiting: 최소화 (1ms - 빠른 입력 허용하면서도 과도한 호출 방지)
    const now = Date.now();
    if (now - lastUpdateSealantDetailCall.current < 1) {
      return;
    }
    lastUpdateSealantDetailCall.current = now;
    
    // 개별 필드 throttle 제거 - 빠른 타이핑 시에도 모든 입력이 반영되도록 함
    
    setTreatmentData(prev => {
      // rowIndex 범위 검증
      if (rowIndex < 0 || rowIndex >= prev.length) {
        return prev;
      }
      
      // 허용된 필드만 처리
      const allowedFields = ['ptName', 'chartNumber', 'dob', 'toothNumber', 'redo', 'acctType', 'payable', 'dxDr', 'drAmount', 'rdaAmount'];
      if (!allowedFields.includes(field)) {
        return prev;
      }
      
      // value가 문자열인지 확인
      let safeValue = typeof value === 'string' ? value : String(value || '');
      
      // 금액 필드는 숫자만 허용 (소수점 포함)
      if (field === 'drAmount' || field === 'rdaAmount') {
        safeValue = sanitizeAmount(safeValue);
      } else {
        // XSS 방지: 위험한 문자 제거
        safeValue = safeValue.replace(/[<>\"']/g, '');
      }
      
      // 필드별 길이 제한
      const maxLengths: { [key: string]: number } = {
        ptName: 100,
        chartNumber: 50,
        dob: 20,
        toothNumber: 10,
        redo: 10,
        acctType: 10,
        payable: 10,
        dxDr: 50,
        drAmount: 20,
        rdaAmount: 20
      };
      const limitedValue = safeValue.substring(0, maxLengths[field] || 1000);
      
      const newData = [...prev];
      
      // 보안 강화: newData[rowIndex] 존재 및 유효성 검증
      if (!newData[rowIndex] || typeof newData[rowIndex] !== 'object') {
        return prev;
      }
      
      if (!Array.isArray(newData[rowIndex].sealantDetails)) {
        newData[rowIndex].sealantDetails = [];
      }
      
      // detailIndex 범위 검증 및 최대 개수 제한
      if (!Number.isInteger(detailIndex) || detailIndex < 0 || detailIndex >= 100) {
        return prev;
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
        [field]: limitedValue
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
    // Rate limiting: 최소 500ms 간격으로 추가 (중복 추가 방지)
    const now = Date.now();
    if (now - lastAddSealantDetailCall.current < 500) {
      return;
    }
    lastAddSealantDetailCall.current = now;
    
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
      
      // Sealant가 없는 행의 sealantDetails를 빈 배열로 명시적으로 설정
      const normalizedTreatmentData = prevTreatmentData.map((row: any) => {
        const hasSealant = row.services && Array.isArray(row.services) && row.services.includes('Sealant');
        return {
          ...row,
          sealantDetails: hasSealant ? (row.sealantDetails || []) : []
        };
      });
      
      const dataToSave = {
        office: prevOffice,
        rdaName: prevRdaName,
        date: prevDate,
        treatmentData: normalizedTreatmentData,
        timestamp: pacificDateTime.toISOString(),
        autoSaved: true,
        lastUpdatedBy: userSessionId,
        submitted: isSubmitted // 제출 상태 유지
      };

      const prevDocId = createSafeDocId(prevDate, prevOffice, prevRdaName);
      await setDoc(doc(db, "rda-treatment-sheets", prevDocId), dataToSave);
    } catch (error) {
      // 프로덕션에서는 상세한 에러 정보를 노출하지 않음
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
          setAutoSaveStatus('This data has been submitted.');
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
        setAutoSaveStatus('Loading data...');
        setIsUpdatingFromFirebase(true);
        
        // 치료 데이터를 정규화하여 sealantDetails가 없는 경우 빈 배열로 초기화
        // Sealant가 없는 행의 sealantDetails는 빈 배열로 설정
        const normalizedTreatmentData = (data.treatmentData || []).map((row: any) => {
          const hasSealant = row.services && Array.isArray(row.services) && row.services.includes('Sealant');
          return {
            patientName: row.patientName || '',
            startTime: row.startTime || '',
            roomNumber: row.roomNumber || '',
            services: row.services || [],
            explanation: row.explanation || '',
            showServices: row.showServices || false,
            sealantDetails: hasSealant ? (row.sealantDetails || []) : [] // Sealant가 없으면 빈 배열
          };
        });
        
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
      setAutoSaveStatus('error');
      setTimeout(() => setAutoSaveStatus(''), 3000);
    }
  }, [office, rdaName, date]);

  // treatmentData가 변경될 때마다 Sealant가 없는 행의 sealantDetails 정리
  // 보안 강화: ref를 사용하여 무한 루프 방지 및 성능 최적화
  const lastCleanupRef = useRef<string>('');
  useEffect(() => {
    // Firebase 업데이트 중이면 정리하지 않음 (무한 루프 방지)
    if (isUpdatingFromFirebase) {
      return;
    }
    
    // treatmentData의 해시를 생성하여 실제 변경이 있었는지 확인 (성능 최적화)
    const dataHash = JSON.stringify(treatmentData.map((row: any) => ({
      hasSealant: row.services && Array.isArray(row.services) && row.services.includes('Sealant'),
      hasSealantDetails: row.sealantDetails && row.sealantDetails.length > 0
    })));
    
    // 이전에 정리한 데이터와 동일하면 스킵 (무한 루프 방지)
    if (lastCleanupRef.current === dataHash) {
      return;
    }
    
    // Sealant가 없는 행의 sealantDetails가 있는지 확인
    const needsCleanup = treatmentData.some((row: any) => {
      // 입력값 검증: row가 유효한 객체인지 확인
      if (!row || typeof row !== 'object') {
        return false;
      }
      const hasSealant = row.services && Array.isArray(row.services) && row.services.includes('Sealant');
      return !hasSealant && row.sealantDetails && Array.isArray(row.sealantDetails) && row.sealantDetails.length > 0;
    });
    
    // 정리가 필요한 경우에만 업데이트
    if (needsCleanup) {
      setTreatmentData(prev => {
        // 입력값 검증: prev가 유효한 배열인지 확인
        if (!Array.isArray(prev)) {
          return prev;
        }
        
        const cleaned = prev.map((row: any, index: number) => {
          // 행 데이터 검증
          if (!row || typeof row !== 'object' || !Number.isInteger(index) || index < 0 || index >= 10000) {
            return row;
          }
          
          // services 배열 검증
          const services = Array.isArray(row.services) ? row.services : [];
          const hasSealant = services.includes('Sealant');
          
          // sealantDetails 검증 및 정리
          if (!hasSealant) {
            const sealantDetails = Array.isArray(row.sealantDetails) ? row.sealantDetails : [];
            if (sealantDetails.length > 0) {
              // Sealant가 없으면 빈 배열로 설정 (다른 필드는 그대로 유지)
              return {
                ...row,
                sealantDetails: []
              };
            }
          }
          return row;
        });
        
        // 정리된 데이터의 해시 저장
        const cleanedHash = JSON.stringify(cleaned.map((row: any) => ({
          hasSealant: row.services && Array.isArray(row.services) && row.services.includes('Sealant'),
          hasSealantDetails: row.sealantDetails && row.sealantDetails.length > 0
        })));
        lastCleanupRef.current = cleanedHash;
        
        return cleaned;
      });
    } else {
      // 정리가 필요 없어도 해시 업데이트 (다음 변경 감지용)
      lastCleanupRef.current = dataHash;
    }
  }, [treatmentData, isUpdatingFromFirebase]);

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
    
    // Rate limiting: 중복 제출 방지
    if (loading) {
      return;
    }
    
    // Rate limiting: 최소 3초 간격으로 저장 (서버 부하 방지)
    const now = Date.now();
    const timeSinceLastSave = now - lastAutoSaveTimeRef.current;
    if (timeSinceLastSave < 3000) {
      alert('⚠️ Please try again.');
      return;
    }
    lastAutoSaveTimeRef.current = now;

    try {
      setLoading(true);
      
      // 캘리포니아 시간대로 현재 시간 생성
      const currentTime = new Date();
      const pacificDateTime = new Date(currentTime.toLocaleString("en-US", {timeZone: "America/Los_Angeles"}));
      
      // 선택된 날짜가 캘리포니아 시간대인지 확인하고 필요시 조정
      const selectedDateObj = new Date(date + 'T00:00:00'); // 로컬 시간으로 해석
      const pacificDateString = pacificDateTime.toISOString().split('T')[0];
      
      // Sealant가 없는 행의 sealantDetails를 빈 배열로 명시적으로 설정
      // 🔒 보안: 배열 데이터 검증 및 정규화 (autoSave와 동일한 검증 적용)
      const normalizedTreatmentData = treatmentData
        .slice(0, 10000) // 최대 10000개 행만 허용
        .map((row: any) => {
          if (!row || typeof row !== 'object') {
            return {
              patientName: '',
              startTime: '',
              roomNumber: '',
              services: [],
              explanation: '',
              showServices: false,
              sealantDetails: []
            };
          }
          
          const hasSealant = row.services && Array.isArray(row.services) && row.services.includes('Sealant');
          
          // sealantDetails 검증 및 정규화
          let safeSealantDetails: any[] = [];
          if (hasSealant && Array.isArray(row.sealantDetails)) {
            safeSealantDetails = row.sealantDetails
              .slice(0, 100) // 최대 100개만 허용
              .map((detail: any) => {
                if (!detail || typeof detail !== 'object') {
                  return {
                    ptName: row.patientName || '',
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
                return {
                  ptName: String(detail.ptName || '').substring(0, 100).replace(/[<>\"']/g, ''),
                  chartNumber: String(detail.chartNumber || '').substring(0, 50).replace(/[<>\"']/g, ''),
                  dob: String(detail.dob || '').substring(0, 20).replace(/[<>\"']/g, ''),
                  toothNumber: String(detail.toothNumber || '').substring(0, 10).replace(/[<>\"']/g, ''),
                  redo: String(detail.redo || '').substring(0, 10).replace(/[<>\"']/g, ''),
                  acctType: String(detail.acctType || '').substring(0, 10).replace(/[<>\"']/g, ''),
                  payable: String(detail.payable || '').substring(0, 10).replace(/[<>\"']/g, ''),
                  dxDr: String(detail.dxDr || '').substring(0, 50).replace(/[<>\"']/g, ''),
                  drAmount: sanitizeAmount(String(detail.drAmount || '')),
                  rdaAmount: sanitizeAmount(String(detail.rdaAmount || ''))
                };
              });
          }
          
          return {
            patientName: String(row.patientName || '').substring(0, 100).replace(/[<>\"']/g, ''),
            startTime: String(row.startTime || '').substring(0, 10).replace(/[<>\"']/g, ''),
            roomNumber: String(row.roomNumber || '').substring(0, 10).replace(/[<>\"']/g, ''),
            services: Array.isArray(row.services) 
              ? row.services
                  .filter((s: any) => typeof s === 'string' && SERVICE_OPTIONS.includes(s))
                  .slice(0, 50)
              : [],
            explanation: String(row.explanation || '').substring(0, 500).replace(/[<>\"']/g, ''),
            showServices: Boolean(row.showServices),
            sealantDetails: safeSealantDetails
          };
        });
      
      const formData = {
        office: office.trim().substring(0, 10),
        rdaName: rdaName.trim().substring(0, 50),
        date: date.trim().substring(0, 20),
        treatmentData: normalizedTreatmentData,
        createdAt: pacificDateTime.toISOString(),
        lastUpdated: pacificDateTime.toISOString()
      };

      // 🔒 보안: 데이터 저장 전 검증 및 sanitization
      const sanitizedFormData = sanitizeDataForFirebase(formData);

      // autoSave와 동일한 문서 ID 사용 (중복 저장 방지)
      const docId = createSafeDocId(date, office, rdaName);
      
      // Firebase에 저장 (setDoc 사용으로 기존 문서 업데이트)
      const finalData = {
        ...sanitizedFormData,
        lastUpdated: new Date(),
        submitted: true
      };
      
      await setDoc(doc(db, 'rda-treatment-sheets', docId), finalData);
      
      // 제출 상태 업데이트
      setIsSubmitted(true);
      
      alert('Submitted successfully!');
      
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
      
    } catch (error: any) {
      // 🔒 보안: 프로덕션에서는 상세한 에러 정보를 노출하지 않음
      console.error('Submit error:', error);
      alert('Submission failed. Please check your data and try again.');
    } finally {
      setLoading(false);
    }
  }, [office, rdaName, date, treatmentData]);

  // 폼 초기화
  const handleClear = useCallback(() => {
    if (confirm('Would you like to submit?')) {
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
        const urlOffice = urlParams.get('office') || '';
        const urlRdaName = urlParams.get('rdaName') || '';
        const urlDate = urlParams.get('date') || '';
        const isViewMode = urlParams.get('view') === 'true';
        
        // URL 파라미터 검증 및 sanitization (보안 강화)
        const safeOffice = urlOffice.trim().replace(/[^A-Za-z0-9]/g, '').substring(0, 10);
        const safeRdaName = urlRdaName.trim().replace(/[^A-Za-z0-9\s]/g, '').substring(0, 50);
        let safeDate = urlDate.trim().replace(/[^0-9-]/g, '');
        if (safeDate && !/^\d{4}-\d{2}-\d{2}$/.test(safeDate)) {
          safeDate = '';
        }
        // 날짜 범위 검증 (1900-01-01 ~ 2100-12-31)
        if (safeDate) {
          const dateObj = new Date(safeDate + 'T00:00:00');
          const year = dateObj.getFullYear();
          if (isNaN(year) || year < 1900 || year > 2100) {
            safeDate = '';
          }
        }
        
        // Generate PDF 후에는 데이터를 로드하지 않음
        if (pdfGeneratedRef.current) {
          isInitialLoadRef.current = false;
          return;
        }
        
        // 초기 로드이고 view 모드이면 데이터 로드 (뒤로가기가 아닌 경우)
        if (isInitialLoadRef.current && safeOffice && safeRdaName && safeDate && isViewMode) {
          setOffice(safeOffice);
          setRdaName(safeRdaName);
          setDate(safeDate);
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

  // 인증 확인 중이거나 인증 실패 시 처리
  if (isAuthorized === null) {
    return (
      <div style={bodyStyle}>
        <div style={{ textAlign: 'center', padding: '50px', fontSize: '18px' }}>
          Verifying authentication...
        </div>
      </div>
    );
  }

  if (isAuthorized === false) {
    return null; // 리다이렉트 중이므로 아무것도 렌더링하지 않음
  }

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
      {/* 자동 저장 상태 표시 - 에러 메시지만 표시 */}
      {autoSaveStatus && autoSaveStatus.includes('❌') && (
        <div style={{
          position: 'fixed',
          top: '20px',
          right: '20px',
          padding: '12px 20px',
          backgroundColor: '#ff6b6b',
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
        <h1 style={headerStyle}>RDA/DA Treatment</h1>

        {/* 기본 정보 섹션 */}
        <div style={sectionStyle}>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '20px' }}>
            <div>
              <label style={{ display: 'block', marginBottom: '8px', fontWeight: 'bold', color: '#495057' }}>
                Date:
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
                Office:
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
                RDA/DA Name:
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
                    isViewMode={isViewModeFromUrl}
                    isSealantDetailsEditable={isSealantDetailsEditable}
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

        {/* 서비스별 체크 개수 요약 (Unlock 시에만 표시) */}
        {isUnlocked && (() => {
          // 서비스별 체크 개수 계산
          const serviceCounts: { [key: string]: number } = {};
          SERVICE_OPTIONS.forEach(service => {
            serviceCounts[service] = 0;
          });
          
          treatmentData.forEach(row => {
            if (row.services && Array.isArray(row.services)) {
              row.services.forEach((service: string) => {
                if (serviceCounts.hasOwnProperty(service)) {
                  serviceCounts[service]++;
                }
              });
            }
          });
          
          // 체크된 서비스만 필터링
          const checkedServices = Object.entries(serviceCounts)
            .filter(([_, count]) => count > 0)
            .sort((a, b) => b[1] - a[1]); // 개수 순으로 정렬
          
          if (checkedServices.length === 0) {
            return null;
          }
          
          return (
            <div style={{
              marginBottom: '15px',
              backgroundColor: '#f8f9fa',
              border: '1px solid #0077B6',
              borderRadius: '6px',
              padding: '12px'
            }}>
              <h3 style={{
                color: '#0077B6',
                marginBottom: '8px',
                marginTop: '0',
                fontSize: '14px',
                fontWeight: 'bold',
                textAlign: 'center'
              }}>
                Treatment or Services Performed
              </h3>
              <div style={{
                display: 'grid',
                gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
                gap: '6px',
                padding: '0'
              }}>
                {checkedServices.map(([service, count]) => (
                  <div
                    key={service}
                    style={{
                      backgroundColor: 'white',
                      padding: '6px 10px',
                      borderRadius: '4px',
                      border: '1px solid #dee2e6',
                      display: 'flex',
                      justifyContent: 'space-between',
                      alignItems: 'center',
                      boxShadow: '0 1px 2px rgba(0,0,0,0.05)'
                    }}
                  >
                    <span style={{
                      fontSize: '12px',
                      color: '#495057',
                      flex: 1,
                      marginRight: '8px',
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                      whiteSpace: 'nowrap'
                    }}>
                      {service}
                    </span>
                    <span style={{
                      fontSize: '13px',
                      fontWeight: 'bold',
                      color: '#0077B6',
                      backgroundColor: '#e7f3ff',
                      padding: '2px 8px',
                      borderRadius: '12px',
                      minWidth: '30px',
                      textAlign: 'center',
                      flexShrink: 0
                    }}>
                      {count}
                    </span>
                  </div>
                ))}
              </div>
            </div>
          );
        })()}

        {/* 액션 버튼들 (Unlock 시에만 표시) */}
        {isUnlocked && (
        <div style={{ textAlign: 'center', padding: '30px' }}>
          {/* View Details로 접근한 경우 PDF 버튼 표시 */}
          {isClient && typeof window !== 'undefined' && new URLSearchParams(window.location.search).get('view') === 'true' ? (
            <button
              onClick={async () => {
                // Rate limiting: 최소 5초 간격으로 PDF 생성 (서버 부하 방지)
                const now = Date.now();
                const timeSinceLastCall = now - lastApiCallTimeRef.current;
                if (timeSinceLastCall < 5000) {
                  alert('⚠️ Please wait a moment before generating another PDF.');
                  return;
                }
                lastApiCallTimeRef.current = now;
                
                // 바로 PDF 생성
                setLoading(true);
                setAutoSaveStatus('Processing...');

                try {
                  // 🔒 보안: 입력값 검증 및 정리
                  const safeOffice = String(office).trim().replace(/[^A-Za-z0-9]/g, '').substring(0, 10);
                  const safeRdaName = String(rdaName).trim().replace(/[^A-Za-z0-9\s]/g, '').substring(0, 50);
                  const safeDate = String(date).trim().replace(/[^0-9-]/g, '');
                  
                  // 날짜 형식 검증 (YYYY-MM-DD)
                  if (!/^\d{4}-\d{2}-\d{2}$/.test(safeDate)) {
                    alert('Invalid date format');
                    setLoading(false);
                    setAutoSaveStatus('');
                    return;
                  }
                  
                  // 데이터 타입 검증
                  if (!Array.isArray(treatmentData)) {
                    alert('Invalid treatment data format');
                    setLoading(false);
                    setAutoSaveStatus('');
                    return;
                  }
                  
                  // 배열 크기 제한 (DoS 공격 방지)
                  if (treatmentData.length > 1000) {
                    alert('Treatment data array too large');
                    setLoading(false);
                    setAutoSaveStatus('');
                    return;
                  }

                  // Sealant가 없는 행의 sealantDetails를 빈 배열로 명시적으로 설정
                  const normalizedTreatmentData = treatmentData.map((row: any) => {
                    const hasSealant = row.services && Array.isArray(row.services) && row.services.includes('Sealant');
                    return {
                      ...row,
                      sealantDetails: hasSealant ? (row.sealantDetails || []) : []
                    };
                  });

                  // 치료 데이터 필터링 및 검증 (빈 행 제외)
                  const filteredData = normalizedTreatmentData
                    .filter((row: any) => {
                      if (!row || typeof row !== 'object' || Array.isArray(row)) {
                        return false;
                      }
                      return row.patientName || row.startTime || row.roomNumber || 
                             (row.services && Array.isArray(row.services) && row.services.length > 0) || 
                             row.explanation ||
                             (row.sealantDetails && Array.isArray(row.sealantDetails) && row.sealantDetails.length > 0);
                    })
                    .slice(0, 1000);

                  // Sealant Details 수집
                  const allSealantDetails: Array<{
                    patientName: string;
                    detail: any;
                    rowIndex: number;
                  }> = [];

                  filteredData.forEach((row, rowIndex) => {
                    if (row.sealantDetails && Array.isArray(row.sealantDetails) && row.sealantDetails.length > 0) {
                      const limitedDetails = row.sealantDetails.slice(0, 100);
                      limitedDetails.forEach((detail: any) => {
                        if (detail && typeof detail === 'object' && !Array.isArray(detail) && detail.constructor === Object) {
                          const sanitizedDetail: any = {};
                          const allowedFields = ['ptName', 'chartNumber', 'dob', 'toothNumber', 'redo', 'acctType', 'payable', 'dxDr', 'drAmount', 'rdaAmount'];
                          
                          allowedFields.forEach(field => {
                            if (detail.hasOwnProperty(field)) {
                              const value = detail[field];
                              if (value === null || value === undefined) {
                                sanitizedDetail[field] = '';
                              } else if (typeof value === 'string' || typeof value === 'number') {
                                sanitizedDetail[field] = String(value);
                              } else {
                                sanitizedDetail[field] = '';
                              }
                            } else {
                              sanitizedDetail[field] = '';
                            }
                          });
                          
                          allSealantDetails.push({
                            patientName: row.patientName || '',
                            detail: sanitizedDetail,
                            rowIndex: rowIndex
                          });
                        }
                      });
                    }
                  });

                  // 서비스별 체크 개수 계산
                  const calculateServiceCounts = () => {
                    const serviceCounts: { [key: string]: number } = {};
                    SERVICE_OPTIONS.forEach(service => {
                      serviceCounts[service] = 0;
                    });
                    
                    filteredData.forEach((row: any) => {
                      if (row && typeof row === 'object' && row.services && Array.isArray(row.services)) {
                        row.services.forEach((service: any) => {
                          if (typeof service === 'string' && serviceCounts.hasOwnProperty(service)) {
                            serviceCounts[service]++;
                          }
                        });
                      }
                    });
                    
                    return Object.entries(serviceCounts)
                      .filter(([_, count]) => count > 0)
                      .sort((a, b) => b[1] - a[1]);
                  };

                  const serviceCounts = calculateServiceCounts();
                  const generatedDate = new Date().toLocaleString('en-US', { timeZone: 'America/Los_Angeles' });

                  // PDF 생성 (클라이언트 사이드)
                  setAutoSaveStatus('Processing...');
                  const pdfDoc = createRDATreatmentPDFDocument({
                    safeDate,
                    safeOffice,
                    safeRdaName,
                    filteredData,
                    allSealantDetails,
                    serviceCounts,
                    generatedDate
                  });

                  const pdfBlob = await pdf(pdfDoc).toBlob();

                  if (!pdfBlob || pdfBlob.size === 0) {
                    throw new Error('error');
                  }
                    
                    // PDF를 Firebase Storage에 저장
                    setAutoSaveStatus('Submitting...');
                    try {
                      const storage = getStorage();
                      const filename = `8) ${safeDate}_${safeOffice}_${safeRdaName}_RDA/DA Treatment(Sealant) Sheet.pdf`;
                      const storageRef = ref(storage, `endofday-pdfs/${safeOffice}/${safeDate}/${filename}`);
                      
                      // PDF 업로드
                      await uploadBytes(storageRef, pdfBlob);
                      
                      // 다운로드 URL 가져오기
                      const downloadUrl = await getDownloadURL(storageRef);
                      
                      // Firestore에 메타데이터 저장
                      await setDoc(doc(db, 'pdf-documents', `${safeDate}_${safeOffice}_${safeRdaName}_rda-treatment_${Date.now()}`), {
                        filename,
                        office: safeOffice,
                        date: safeDate,
                        type: 'RDA Treatment Sheet',
                        url: downloadUrl,
                        storagePath: `endofday-pdfs/${safeOffice}/${safeDate}/${filename}`,
                        createdAt: new Date(),
                      });
                      
                      setAutoSaveStatus('✅ PDF saved to archive successfully!');
                    } catch (storageError: any) {
                      // 프로덕션에서는 상세한 에러 정보를 노출하지 않음
                      alert('error');
                      setAutoSaveStatus('error');
                    }
                    
                    // PDF 생성 성공 후 데이터베이스에서 삭제
                    try {
                      setAutoSaveStatus('Processing...');
                      const docId = createSafeDocId(safeDate, safeOffice, safeRdaName);
                      await deleteDoc(doc(db, 'rda-treatment-sheets', docId));
                      setAutoSaveStatus('Submitted successfully!');
                    } catch (deleteError) {
                      setAutoSaveStatus('error');
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
                } catch (error: any) {
                  // 프로덕션에서는 상세한 에러 정보를 노출하지 않음
                  alert('error');
                  setAutoSaveStatus('error');
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
              {loading ? 'Submitting...' : '📄 Submit'}
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
                {loading ? 'Submitting...' : '📤 Submit'}
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
