'use client'

import React, { useState, useEffect, useRef } from "react";
import { doc, setDoc, collection, getDocs, getDoc, updateDoc, deleteDoc, query, where } from "firebase/firestore";
import { db } from "@/lib/firebase.config";
import { 
  enableAllSecurityMeasures, 
  sanitizeFirebaseDataClient, 
  sanitizeFirebaseDocIdClient,
  secureFetch,
  checkAuthState,
  getCurrentUser
} from "@/lib/security-client";

export default function ShowCheckSystem() {
  // 보안 조치 활성화
  useEffect(() => {
    enableAllSecurityMeasures({
      disableConsole: true,
      disableRightClick: true,
      disableShortcuts: true,
      disableCopy: false,
      disableSelection: false,
      monitorDevTools: false
    });
  }, []);

  // 상태 관리
  const [loading, setLoading] = useState(false);
  const [appointments, setAppointments] = useState<any[]>([]);
  const appointmentsRef = useRef<any[]>([]); // 최신 appointments 추적
  const [filteredAppointments, setFilteredAppointments] = useState<any[]>([]);
  const [startDate, setStartDate] = useState(new Date().toISOString().split('T')[0]);
  const [endDate, setEndDate] = useState(new Date().toISOString().split('T')[0]);
  const [selectedOffice, setSelectedOffice] = useState('');
  const [name, setName] = useState('');
  const [pdfLoading, setPdfLoading] = useState(false);
  const [submitStatus, setSubmitStatus] = useState('');
  const [progress, setProgress] = useState(0);

  // Office 옵션
  const officeOptions = ['All', 'Bernard', 'Call Center', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

  // 컴포넌트 마운트 시 데이터 로드 (Firebase Auth 초기화 대기)
  useEffect(() => {
    // Firebase Auth 초기화를 기다린 후 데이터 로드
    const initAndLoad = async () => {
      // 짧은 지연 후 로드 (Firebase Auth 초기화 시간 확보)
      await new Promise(resolve => setTimeout(resolve, 100));
      await loadAppointments();
    };
    
    initAndLoad();
  }, []);

  // 필터 변경 시 데이터 필터링
  useEffect(() => {
    filterAppointments();
  }, [appointments, startDate, endDate, selectedOffice]);

  // 제출 중 브라우저 네비게이션 방지
  useEffect(() => {
    const handleBeforeUnload = (e: any) => {
      if (pdfLoading) {
        e.preventDefault();
        e.returnValue = '';
      }
    };

    const handlePopState = (e: any) => {
      if (pdfLoading) {
        e.preventDefault();
        window.history.pushState(null, '', window.location.href);
      }
    };

    if (pdfLoading) {
      window.addEventListener('beforeunload', handleBeforeUnload);
      window.addEventListener('popstate', handlePopState);
      window.history.pushState(null, '', window.location.href);
    }

    return () => {
      window.removeEventListener('beforeunload', handleBeforeUnload);
      window.removeEventListener('popstate', handlePopState);
    };
  }, [pdfLoading]);

  // Firebase에서 모든 환자 로그 불러오기 (보안 강화)
  const loadAppointments = async () => {
    try {
      // 인증 상태 확인 (보안 강화 - 초기화 대기 포함)
      const isAuthenticated = await checkAuthState(true); // waitForInit = true
      if (!isAuthenticated) {
        // 알림을 표시하지 않고 조용히 실패 (새로고침 시 Firebase Auth 초기화 중일 수 있음)
        // 사용자가 실제로 로그인하지 않은 경우에만 알림 표시
        // Firebase Security Rules가 이미 접근을 제어하므로 여기서는 조용히 실패
        return;
      }

      // 현재 사용자 정보 가져오기
      const currentUser = await getCurrentUser();
      if (!currentUser) {
        alert('⚠️ 사용자 정보를 가져올 수 없습니다. 로그인 상태를 확인해주세요.');
        return;
      }

      setLoading(true);
      const querySnapshot = await getDocs(collection(db, "patient-logs"));
      
      // 제출된 PDF 목록 가져오기 (제출된 것만 필터링하기 위해)
      const pdfQuerySnapshot = await getDocs(collection(db, "pdf-documents"));
      const submittedDocs = new Set<string>();
      
      pdfQuerySnapshot.forEach((pdfDoc) => {
        const pdfData = pdfDoc.data();
        if (pdfData.type === 'Patient Log' && pdfData.date && pdfData.name && pdfData.office) {
          // 제출된 문서의 키 생성: date_name_office
          const key = `${pdfData.date}_${pdfData.name}_${pdfData.office}`;
          submittedDocs.add(key);
        }
      });
      
      const allAppointments: any[] = [];
      
      querySnapshot.forEach((doc) => {
        const data = doc.data();
        
        // 보안 강화: 데이터 소유권 확인 (Firebase Security Rules와 함께)
        if (currentUser && data.userId && data.userId !== currentUser.uid) {
          // 다른 사용자의 데이터는 로드하지 않음 (Security Rules에서도 차단됨)
          return;
        }
        
        // 제출된 데이터만 필터링
        if (data.dutyDate && data.userName && data.workOffice) {
          const docKey = `${data.dutyDate}_${data.userName}_${data.workOffice}`;
          if (!submittedDocs.has(docKey)) {
            // 제출되지 않은 데이터는 건너뛰기
            return;
          }
        }
        
        if (data.patientRows) {
          data.patientRows.forEach((row: any, index: number) => {
            if (row.appt_date && row.name) {
              allAppointments.push({
            ...row,
                docId: doc.id,
                rowIndex: index,
                dutyDate: data.dutyDate,
                userName: data.userName,
                workOffice: data.workOffice,
                workHoursFrom: data.workHoursFrom,
                workHoursTo: data.workHoursTo,
                showStatus: row.showStatus || 'pending' // pending, show, no-show
              });
            }
          });
        }
      });
      
      setAppointments(allAppointments);
      appointmentsRef.current = allAppointments;
    } catch (error) {
      // Production에서는 에러 로깅 비활성화
      if (process.env.NODE_ENV !== 'production') {
        console.error("Error loading appointments:", error);
      }
      alert('❌ 데이터 로드 중 오류가 발생했습니다.');
    } finally {
      setLoading(false);
    }
  };

  // 필터링 함수
  const filterAppointments = () => {
    let filtered = appointments;
    
    // 날짜 범위 필터
    if (startDate && endDate) {
      filtered = filtered.filter(apt => {
        const aptDate = new Date(apt.appt_date);
        const start = new Date(startDate);
        const end = new Date(endDate);
        return aptDate >= start && aptDate <= end;
      });
    }
    
    // 오피스 필터
    if (selectedOffice && selectedOffice !== 'All') {
      filtered = filtered.filter(apt => apt.office === selectedOffice);
    }
    
    setFilteredAppointments(filtered);
  };

  // Show/No Show 상태 업데이트 (보안 강화)
  const updateShowStatus = async (appointment: any, newStatus: string) => {
    try {
      // 인증 상태 확인 (보안 강화)
      const isAuthenticated = await checkAuthState();
      if (!isAuthenticated) {
        alert('⚠️ 로그인이 필요합니다. 로그인 후 다시 시도해주세요.');
        return;
      }

      // 현재 사용자 정보 가져오기
      const currentUser = await getCurrentUser();
      if (!currentUser) {
        alert('⚠️ 사용자 정보를 가져올 수 없습니다. 로그인 상태를 확인해주세요.');
        return;
      }

      // 해당 document를 다시 가져와서 patientRows 업데이트
      const safeDocId = sanitizeFirebaseDocIdClient(appointment.docId);
      const docRef = doc(db, "patient-logs", safeDocId);
      const querySnapshot = await getDocs(collection(db, "patient-logs"));
      let currentData: any = null;
      
      querySnapshot.forEach((document) => {
        if (document.id === appointment.docId) {
          currentData = document.data();
        }
      });
      
      // 보안 강화: 데이터 소유권 확인
      if (currentData && currentData.userId && currentData.userId !== currentUser.uid) {
        alert('⚠️ 권한이 없습니다. 다른 사용자의 데이터를 수정할 수 없습니다.');
        return;
      }
      
      if (currentData && currentData.patientRows) {
        // 해당 row의 showStatus 업데이트
        const updatedPatientRows = currentData.patientRows.map((row: any, index: number) => {
          if (index === appointment.rowIndex) {
            return { ...row, showStatus: newStatus };
          }
          return row;
        });
        
        // Firebase 업데이트 (보안 검증 적용 + 사용자 정보 강제 추가)
        const safeUpdateData = sanitizeFirebaseDataClient({
          patientRows: updatedPatientRows,
          lastUpdated: new Date().toISOString(),
          // 보안 강화: 사용자 정보 강제 추가
          userId: currentUser.uid, // 항상 현재 사용자 ID로 설정
          ...(currentUser.email && { userEmail: currentUser.email }),
          updatedAt: new Date().toISOString()
        });
        await updateDoc(docRef, safeUpdateData);
        
        // 로컬 상태 업데이트
        setAppointments(prev => {
          const updated = prev.map((apt: any) => 
            apt.docId === appointment.docId && apt.rowIndex === appointment.rowIndex
              ? { ...apt, showStatus: newStatus }
              : apt
          );
          
          // appointmentsRef도 업데이트 (최신 값 보장)
          appointmentsRef.current = updated;
          
          // filteredAppointments도 즉시 업데이트 (useEffect 대기 없이)
          // filterAppointments 함수 로직을 재사용
          let filtered = updated;
          
          // 날짜 범위 필터
          if (startDate && endDate) {
            filtered = filtered.filter(apt => {
              const aptDate = new Date(apt.appt_date);
              const start = new Date(startDate);
              const end = new Date(endDate);
              return aptDate >= start && aptDate <= end;
            });
          }
          
          // 오피스 필터
          if (selectedOffice && selectedOffice !== 'All') {
            filtered = filtered.filter(apt => apt.office === selectedOffice);
          }
          
          // 즉시 filteredAppointments 업데이트
          setFilteredAppointments(filtered);
          
          return updated;
        });
        
        // Production에서는 로깅 비활성화
        if (process.env.NODE_ENV !== 'production') {
          console.log(`✅ ${appointment.name}의 상태가 ${newStatus}로 업데이트되었습니다.`);
        }
      }
    } catch (error) {
      // Production에서는 에러 로깅 비활성화
      if (process.env.NODE_ENV !== 'production') {
        console.error("Error updating show status:", error);
      }
      alert('❌ 상태 업데이트 중 오류가 발생했습니다.');
    }
  };

  // PDF 생성 및 제출 (보안 강화)
  const handleGeneratePDF = async () => {
    if (filteredAppointments.length === 0) {
      alert('⚠️ No appointments to generate PDF.');
      return;
    }

    // pending 상태가 있는지 체크 (actions 또는 showStatus)
    const hasPendingStatus = filteredAppointments.some(apt => 
      apt.actions === 'pending' || 
      apt.showStatus === 'pending' || 
      (!apt.actions && !apt.showStatus) ||
      (apt.actions !== 'show' && apt.actions !== 'no-show' && apt.showStatus !== 'show' && apt.showStatus !== 'no-show')
    );
    
    if (hasPendingStatus) {
      alert('⚠️ Please select Show or No Show for all appointments before submitting.');
      return;
    }

    // 인증 상태 확인 (보안 강화)
    const isAuthenticated = await checkAuthState();
    if (!isAuthenticated) {
      alert('⚠️ 로그인이 필요합니다. 로그인 후 다시 시도해주세요.');
      return;
    }

    // 현재 사용자 정보 가져오기 (토큰 갱신 포함)
    const currentUser = await getCurrentUser();
    if (!currentUser) {
      alert('⚠️ 사용자 정보를 가져올 수 없습니다. 로그인 상태를 확인해주세요.');
      return;
    }
    
    // 토큰 미리 갱신 (서버 요청 전에 토큰이 최신인지 확인)
    try {
      const { getAuth } = await import('firebase/auth');
      const auth = getAuth();
      if (auth.currentUser) {
        // 토큰을 미리 갱신하여 최신 상태로 유지
        await auth.currentUser.getIdToken(true);
      }
    } catch (tokenError) {
      // 토큰 갱신 실패는 무시 (secureFetch에서 재시도할 것)
      if (process.env.NODE_ENV !== 'production') {
        console.warn('Token refresh warning:', tokenError);
      }
    }

    try {
      setPdfLoading(true);
      setSubmitStatus('Saving...');
      setProgress(10);

      // PDF용 데이터 준비
      setSubmitStatus('Generating PDF...');
      setProgress(30);
      
      // 최신 appointments state를 기반으로 필터링 (클로저 문제 방지)
      // appointmentsRef.current를 사용하여 최신 상태 보장
      const latestAppointments = appointmentsRef.current.length > 0 ? appointmentsRef.current : appointments;
      
      // 날짜 범위 필터
      let pdfFilteredAppointments = latestAppointments;
      if (startDate && endDate) {
        pdfFilteredAppointments = pdfFilteredAppointments.filter(apt => {
          const aptDate = new Date(apt.appt_date);
          const start = new Date(startDate);
          const end = new Date(endDate);
          return aptDate >= start && aptDate <= end;
        });
      }
      
      // 오피스 필터
      if (selectedOffice && selectedOffice !== 'All') {
        pdfFilteredAppointments = pdfFilteredAppointments.filter(apt => apt.office === selectedOffice);
      }
      
      // pending 상태의 appointment는 PDF에서 제외
      pdfFilteredAppointments = pdfFilteredAppointments.filter(apt => 
        apt.showStatus === 'show' || apt.showStatus === 'no-show'
      );
      
      // PDF 생성에 사용할 데이터 저장 (삭제 시 사용)
      const appointmentsForPdf = pdfFilteredAppointments;
      
      // PDF 생성 전에 데이터가 있는지 확인
      if (appointmentsForPdf.length === 0) {
        alert('⚠️ PDF에 포함할 appointment가 없습니다. Show 또는 No Show를 선택한 appointment가 필요합니다.');
        setPdfLoading(false);
        setSubmitStatus('');
        setProgress(0);
        return;
      }
      
      const pdfData = {
        startDate,
        endDate,
        selectedOffice: selectedOffice || 'All Offices',
        appointments: appointmentsForPdf,
        generatedBy: name || 'Supervisor',
        timestamp: new Date().toISOString(),
        // 보안 강화: 사용자 정보 추가
        userId: currentUser.uid,
        ...(currentUser.email && { userEmail: currentUser.email })
      };

      // API로 PDF 생성 (보안 강화 - secureFetch 사용, 인증 필수)
      let response: Response;
      try {
        // 인증 상태 재확인 (요청 전)
        const isAuthReady = await checkAuthState(true);
        if (!isAuthReady) {
          throw new Error('인증이 필요합니다. 로그인 후 다시 시도해주세요.');
        }
        
        // 토큰 미리 갱신 (요청 전에 최신 토큰 확보)
        const { getAuth } = await import('firebase/auth');
        const auth = getAuth();
        if (auth.currentUser) {
          await auth.currentUser.getIdToken(true); // 강제 갱신
        }
        
        response = await secureFetch('/api/generate-show-check-pdf', {
          method: 'POST',
          body: JSON.stringify({ showCheckData: pdfData }),
        }, true); // requireAuth = true
      } catch (fetchError: any) {
        // secureFetch 자체에서 발생한 에러 (인증 실패 등)
        const errorMsg = fetchError.message || '요청 전송 중 오류가 발생했습니다.';
        
        // 인증 관련 에러인지 확인
        if (errorMsg.includes('Authentication') || 
            errorMsg.includes('token') || 
            errorMsg.includes('인증') ||
            errorMsg.includes('로그인')) {
          // 인증 에러인 경우 재시도 로직으로 넘어가도록 특별 처리
          // response 객체를 생성하여 401 상태로 처리
          response = new Response(
            JSON.stringify({ 
              success: false, 
              error: '인증이 필요합니다. 로그인 후 다시 시도해주세요.' 
            }),
            { 
              status: 401,
              statusText: 'Unauthorized',
              headers: { 'Content-Type': 'application/json' }
            }
          );
        } else {
          // 인증이 아닌 다른 에러
          throw new Error(errorMsg);
        }
      }

      if (response.ok) {
        // HTML 받기
        setSubmitStatus('Opening print view...');
        setProgress(60);
        const htmlContent = await response.text();
        
        setSubmitStatus('Complete!');
        setProgress(100);
        
        // 새 창에서 HTML 열기 (인쇄 대화상자가 자동으로 열림)
        const printWindow = window.open('', '_blank');
        if (printWindow) {
          printWindow.document.write(htmlContent);
          printWindow.document.close();
          
          // 인쇄 완료 메시지 리스너 등록
          const handlePrintComplete = async (event: MessageEvent) => {
            if (event.data === 'print-completed') {
              // 인쇄 완료 후 데이터 삭제
              setSubmitStatus('Cleaning up...');
              setProgress(80);
              
              try {
                // PDF 생성에 사용된 데이터를 전달하여 삭제
                await deleteProcessedAppointments(appointmentsForPdf);
              } catch (deleteError) {
                // Production에서는 에러 로깅 비활성화
                if (process.env.NODE_ENV !== 'production') {
                  console.error('Error deleting processed data:', deleteError);
                }
              }
              
              // 리스너 제거
              window.removeEventListener('message', handlePrintComplete);
              
              // 2초 후 모달 닫기
              setTimeout(() => {
                setPdfLoading(false);
                setSubmitStatus('');
                setProgress(0);
              }, 2000);
            }
          };
          
          window.addEventListener('message', handlePrintComplete);
          
          // 창이 닫히면 리스너 제거 (인쇄하지 않고 닫은 경우)
          const checkClosed = setInterval(() => {
            if (printWindow.closed) {
              clearInterval(checkClosed);
              window.removeEventListener('message', handlePrintComplete);
              // 창이 닫혔지만 인쇄하지 않았을 수 있으므로 데이터는 유지
              setPdfLoading(false);
              setSubmitStatus('');
              setProgress(0);
            }
          }, 1000);
        } else {
          alert('⚠️ 팝업이 차단되었습니다. 팝업을 허용한 후 다시 시도해주세요.');
          setPdfLoading(false);
          setSubmitStatus('');
          setProgress(0);
        }
      } else {
        // 에러 응답 처리 (보안 강화)
        let errorMessage = 'PDF 생성 중 오류가 발생했습니다.';
        let isAuthError = false;
        
        // 먼저 에러 메시지 파싱 시도
        try {
          const contentType = response.headers.get('content-type');
          if (contentType && contentType.includes('application/json')) {
            const errorData = await response.json();
            errorMessage = errorData.error || errorMessage;
            
            // 인증 관련 에러 메시지인지 확인
            if (errorMessage.includes('인증') || 
                errorMessage.includes('Authentication') ||
                errorMessage.includes('token') ||
                errorMessage.includes('로그인') ||
                errorMessage.includes('인증 토큰')) {
              isAuthError = true;
            }
          }
        } catch (parseError) {
          // JSON 파싱 실패 시 상태 코드로 판단
          if (response.status === 401) {
            isAuthError = true;
          }
        }
        
        // 401 에러이거나 인증 관련 에러인 경우 특별 처리 (강화된 재시도)
        if (response.status === 401 || isAuthError) {
          // 토큰 갱신을 여러 번 시도 (최대 3회)
          let retrySuccess = false;
          const maxRetries = 3;
          
          for (let retryCount = 0; retryCount < maxRetries && !retrySuccess; retryCount++) {
            try {
              const { getAuth } = await import('firebase/auth');
              const auth = getAuth();
              
              // 인증 상태 재확인
              if (!auth.currentUser) {
                // 사용자가 로그아웃된 경우
                errorMessage = '로그인 세션이 만료되었습니다. 다시 로그인해주세요.';
                break;
              }
              
              // 토큰 강제 갱신 (매번 새로 갱신)
              const newToken = await auth.currentUser.getIdToken(true);
              if (!newToken) {
                if (retryCount < maxRetries - 1) {
                  // 재시도 전 짧은 대기
                  await new Promise(resolve => setTimeout(resolve, 500 * (retryCount + 1)));
                  continue;
                }
                errorMessage = '토큰을 가져올 수 없습니다. 다시 로그인해주세요.';
                break;
              }
              
              // 갱신된 토큰으로 재시도
              setSubmitStatus(`재시도 중... (${retryCount + 1}/${maxRetries})`);
              const retryResponse = await secureFetch('/api/generate-show-check-pdf', {
                method: 'POST',
                body: JSON.stringify({ showCheckData: pdfData }),
              }, true);
              
              if (retryResponse.ok) {
                // 재시도 성공 - 원래 로직 계속
                retrySuccess = true;
                const htmlContent = await retryResponse.text();
                
                setSubmitStatus('Complete!');
                setProgress(100);
                
                // 새 창에서 HTML 열기 (인쇄 대화상자가 자동으로 열림)
                const printWindow = window.open('', '_blank');
                if (printWindow) {
                  printWindow.document.write(htmlContent);
                  printWindow.document.close();
                  
                  // 인쇄 완료 메시지 리스너 등록
                  const handlePrintComplete = async (event: MessageEvent) => {
                    if (event.data === 'print-completed') {
                      // 인쇄 완료 후 데이터 삭제
                      setSubmitStatus('Cleaning up...');
                      setProgress(80);
                      
                      try {
                        // PDF 생성에 사용된 데이터를 전달하여 삭제
                        await deleteProcessedAppointments(appointmentsForPdf);
                      } catch (deleteError) {
                        if (process.env.NODE_ENV !== 'production') {
                          console.error('Error deleting processed data:', deleteError);
                        }
                      }
                      
                      // 리스너 제거
                      window.removeEventListener('message', handlePrintComplete);
                      
                      setTimeout(() => {
                        setPdfLoading(false);
                        setSubmitStatus('');
                        setProgress(0);
                      }, 2000);
                    }
                  };
                  
                  window.addEventListener('message', handlePrintComplete);
                  
                  // 창이 닫히면 리스너 제거
                  const checkClosed = setInterval(() => {
                    if (printWindow.closed) {
                      clearInterval(checkClosed);
                      window.removeEventListener('message', handlePrintComplete);
                      setPdfLoading(false);
                      setSubmitStatus('');
                      setProgress(0);
                    }
                  }, 1000);
                } else {
                  alert('⚠️ 팝업이 차단되었습니다. 팝업을 허용한 후 다시 시도해주세요.');
                  setPdfLoading(false);
                  setSubmitStatus('');
                  setProgress(0);
                }
                return; // 성공적으로 완료
              } else if (retryResponse.status !== 401) {
                // 401이 아닌 다른 에러면 재시도 중단
                const contentType = retryResponse.headers.get('content-type');
                if (contentType && contentType.includes('application/json')) {
                  try {
                    const errorData = await retryResponse.json();
                    errorMessage = errorData.error || errorMessage;
                  } catch {
                    errorMessage = 'PDF 생성 중 오류가 발생했습니다.';
                  }
                }
                break;
              }
              // 401이면 계속 재시도
              
            } catch (retryError: any) {
              // 재시도 실패 - 다음 시도로
              if (process.env.NODE_ENV !== 'production') {
                console.warn(`Token refresh retry ${retryCount + 1} failed:`, retryError);
              }
              
              if (retryCount < maxRetries - 1) {
                // 재시도 전 대기
                await new Promise(resolve => setTimeout(resolve, 500 * (retryCount + 1)));
              } else {
                // 모든 재시도 실패
                errorMessage = '인증 오류가 지속됩니다. 페이지를 새로고침한 후 다시 시도해주세요.';
              }
            }
          }
          
          // 모든 재시도 실패
          if (!retrySuccess) {
            errorMessage = errorMessage || '인증이 만료되었습니다. 페이지를 새로고침한 후 다시 시도해주세요.';
          }
        } else {
          // 401이 아닌 다른 에러 처리
          try {
            const contentType = response.headers.get('content-type');
            if (contentType && contentType.includes('application/json')) {
              const errorData = await response.json();
              errorMessage = errorData.error || errorMessage;
              
              // 인증 관련 에러 메시지인지 확인
              if (errorMessage.includes('인증') || 
                  errorMessage.includes('Authentication') ||
                  errorMessage.includes('token') ||
                  errorMessage.includes('로그인') ||
                  errorMessage.includes('인증 토큰')) {
                // 인증 에러인 경우 재시도 로직으로 처리
                // 401 상태로 변경하여 재시도 로직 실행
                isAuthError = true;
                response = new Response(
                  JSON.stringify({ 
                    success: false, 
                    error: errorMessage 
                  }),
                  { 
                    status: 401,
                    statusText: 'Unauthorized',
                    headers: { 'Content-Type': 'application/json' }
                  }
                );
                // 재시도 로직으로 넘어가도록 다시 처리
                // (아래 재시도 로직이 실행되도록)
              }
            } else {
              // JSON이 아닌 경우 상태 코드에 따른 메시지
              if (response.status === 403) {
                errorMessage = '권한이 없습니다. 접근이 거부되었습니다.';
              } else if (response.status === 429) {
                errorMessage = '요청이 너무 많습니다. 잠시 후 다시 시도해주세요.';
              } else if (response.status >= 500) {
                errorMessage = '서버 오류가 발생했습니다. 잠시 후 다시 시도해주세요.';
              } else if (response.status === 400) {
                errorMessage = '잘못된 요청입니다. 입력 데이터를 확인해주세요.';
              }
            }
          } catch (parseError) {
            // JSON 파싱 실패 시 상태 코드 기반 메시지 사용
            if (response.status === 403) {
              errorMessage = '권한이 없습니다. 접근이 거부되었습니다.';
            }
          }
          
          // 인증 에러로 감지된 경우 재시도 로직 실행
          if (isAuthError) {
            // 토큰 갱신을 여러 번 시도 (최대 3회)
            let retrySuccess = false;
            const maxRetries = 3;
            
            for (let retryCount = 0; retryCount < maxRetries && !retrySuccess; retryCount++) {
              try {
                const { getAuth } = await import('firebase/auth');
                const auth = getAuth();
                
                // 인증 상태 재확인
                if (!auth.currentUser) {
                  errorMessage = '로그인 세션이 만료되었습니다. 다시 로그인해주세요.';
                  break;
                }
                
                // 토큰 강제 갱신
                const newToken = await auth.currentUser.getIdToken(true);
                if (!newToken) {
                  if (retryCount < maxRetries - 1) {
                    await new Promise(resolve => setTimeout(resolve, 500 * (retryCount + 1)));
                    continue;
                  }
                  errorMessage = '토큰을 가져올 수 없습니다. 다시 로그인해주세요.';
                  break;
                }
                
                // 갱신된 토큰으로 재시도
                setSubmitStatus(`재시도 중... (${retryCount + 1}/${maxRetries})`);
                const retryResponse = await secureFetch('/api/generate-show-check-pdf', {
                  method: 'POST',
                  body: JSON.stringify({ showCheckData: pdfData }),
                }, true);
                
                if (retryResponse.ok) {
                  retrySuccess = true;
                  const htmlContent = await retryResponse.text();
                  
                  setSubmitStatus('Complete!');
                  setProgress(100);
                  
                  // 새 창에서 HTML 열기 (인쇄 대화상자가 자동으로 열림)
                  const printWindow = window.open('', '_blank');
                  if (printWindow) {
                    printWindow.document.write(htmlContent);
                    printWindow.document.close();
                    
                    // 인쇄 완료 메시지 리스너 등록
                    const handlePrintComplete = async (event: MessageEvent) => {
                      if (event.data === 'print-completed') {
                        // 인쇄 완료 후 데이터 삭제
                        setSubmitStatus('Cleaning up...');
                        setProgress(80);
                        
                        try {
                          // PDF 생성에 사용된 데이터를 전달하여 삭제
                          await deleteProcessedAppointments(appointmentsForPdf);
                        } catch (deleteError) {
                          if (process.env.NODE_ENV !== 'production') {
                            console.error('Error deleting processed data:', deleteError);
                          }
                        }
                        
                        // 리스너 제거
                        window.removeEventListener('message', handlePrintComplete);
                        
                        setTimeout(() => {
                          setPdfLoading(false);
                          setSubmitStatus('');
                          setProgress(0);
                        }, 2000);
                      }
                    };
                    
                    window.addEventListener('message', handlePrintComplete);
                    
                    // 창이 닫히면 리스너 제거
                    const checkClosed = setInterval(() => {
                      if (printWindow.closed) {
                        clearInterval(checkClosed);
                        window.removeEventListener('message', handlePrintComplete);
                        setPdfLoading(false);
                        setSubmitStatus('');
                        setProgress(0);
                      }
                    }, 1000);
                  } else {
                    alert('⚠️ 팝업이 차단되었습니다. 팝업을 허용한 후 다시 시도해주세요.');
                    setPdfLoading(false);
                    setSubmitStatus('');
                    setProgress(0);
                  }
                  return; // 성공적으로 완료
                } else if (retryResponse.status !== 401) {
                  const contentType = retryResponse.headers.get('content-type');
                  if (contentType && contentType.includes('application/json')) {
                    try {
                      const errorData = await retryResponse.json();
                      errorMessage = errorData.error || errorMessage;
                    } catch {
                      errorMessage = 'PDF 생성 중 오류가 발생했습니다.';
                    }
                  }
                  break;
                }
              } catch (retryError: any) {
                if (process.env.NODE_ENV !== 'production') {
                  console.warn(`Token refresh retry ${retryCount + 1} failed:`, retryError);
                }
                
                if (retryCount < maxRetries - 1) {
                  await new Promise(resolve => setTimeout(resolve, 500 * (retryCount + 1)));
                } else {
                  errorMessage = '인증 오류가 지속됩니다. 페이지를 새로고침한 후 다시 시도해주세요.';
                }
              }
            }
            
            if (!retrySuccess) {
              errorMessage = errorMessage || '인증이 만료되었습니다. 페이지를 새로고침한 후 다시 시도해주세요.';
            }
          }
        }
        
        throw new Error(errorMessage);
      }

    } catch (error) {
      // Production에서는 에러 로깅 비활성화
      if (process.env.NODE_ENV !== 'production') {
        console.error('PDF generation error:', error);
      }
      setSubmitStatus('❌ PDF generation failed: ' + ((error as any).message || 'Unknown error'));
      setProgress(0);
      setTimeout(() => {
        setPdfLoading(false);
        setSubmitStatus('');
      }, 3000);
    }
  };

  // 처리된 약속 데이터 삭제 함수
  // appointmentsToDelete: PDF 생성에 사용된 appointment 배열 (선택적)
  const deleteProcessedAppointments = async (appointmentsToDelete?: any[]) => {
    try {
      // PDF 생성에 사용된 appointment가 있으면 그것을 사용, 없으면 현재 filteredAppointments 사용
      const appointmentsToProcess = appointmentsToDelete || filteredAppointments;
      
      // 현재 필터링된 약속들의 document ID별로 그룹화
      const documentsToCheck = new Map();
      
      appointmentsToProcess.forEach(appointment => {
        const docId = appointment.docId;
        if (!documentsToCheck.has(docId)) {
          documentsToCheck.set(docId, []);
        }
        documentsToCheck.get(docId).push(appointment.rowIndex);
      });

      const deletedDocuments = [];
      const updatedDocuments = [];

      // 각 document 확인 및 처리
      for (const [docId, processedRowIndices] of documentsToCheck) {
        const safeDocId = sanitizeFirebaseDocIdClient(docId);
        const docRef = doc(db, "patient-logs", safeDocId);
        
        // 현재 document 데이터 가져오기
        const querySnapshot = await getDocs(collection(db, "patient-logs"));
        let currentData: any = null;
        
        querySnapshot.forEach((document) => {
          if (document.id === docId) {
            currentData = document.data();
          }
        });

        if (currentData && currentData.patientRows) {
          // 모든 약속(appt_date가 있는 것들) 찾기
          const allAppointmentRows = currentData.patientRows
            .map((row: any, index: number) => ({ ...row, originalIndex: index }))
            .filter((row: any) => row.appt_date && row.name);

          // 처리되지 않은 약속들 찾기 (현재 필터링된 것들 제외)
          const unprocessedAppointments = allAppointmentRows.filter((row: any) => 
            !processedRowIndices.includes(row.originalIndex)
          );

          if (unprocessedAppointments.length === 0 && allAppointmentRows.length > 0) {
            // 약속이 있었고 모든 약속이 처리됨 → 전체 document 삭제
            await deleteDoc(docRef);
            deletedDocuments.push(docId);
            // Production에서는 로깅 비활성화
            if (process.env.NODE_ENV !== 'production') {
              console.log(`🗑️ Document ${docId} completely deleted (all ${allAppointmentRows.length} appointments processed)`);
            }
          } else if (allAppointmentRows.length === 0) {
            // 애초에 약속이 없는 document → 삭제하지 않음
            if (process.env.NODE_ENV !== 'production') {
              console.log(`📋 Document ${docId} has no appointments, keeping as is`);
            }
          } else {
            // 아직 처리 안 된 약속이 있음 → 처리된 것들만 빈 상태로 초기화
            // showStatus를 pending으로 설정하지 않고 필드를 제거하거나 undefined로 설정
            const updatedPatientRows = currentData.patientRows.map((row: any, index: number) => {
              if (processedRowIndices.includes(index)) {
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
                  other_duty: ''
                  // showStatus 필드를 제거 (pending으로 설정하지 않음)
                };
              }
              return row;
            });

            const safeUpdateData = sanitizeFirebaseDataClient({
              patientRows: updatedPatientRows,
              lastUpdated: new Date().toISOString()
            });
            await updateDoc(docRef, safeUpdateData);
            updatedDocuments.push(docId);
            if (process.env.NODE_ENV !== 'production') {
              console.log(`📝 Document ${docId} updated (${unprocessedAppointments.length} appointments remaining)`);
            }
          }
        }
      }

      // 로컬 상태 업데이트 - 처리된 약속들 제거
      setAppointments(prevAppointments => {
        const updated = prevAppointments.filter(apt => !filteredAppointments.some(filtered => 
          filtered.docId === apt.docId && filtered.rowIndex === apt.rowIndex
        ));
        appointmentsRef.current = updated; // appointmentsRef도 업데이트
        return updated;
      });
      
      setFilteredAppointments([]);
      
      if (process.env.NODE_ENV !== 'production') {
        console.log(`✅ Processing complete: ${deletedDocuments.length} documents deleted, ${updatedDocuments.length} documents updated`);
      }

    } catch (error) {
      // Production에서는 에러 로깅 비활성화
      if (process.env.NODE_ENV !== 'production') {
        console.error('Error deleting processed appointments:', error);
      }
      throw error;
    }
  };

  // 스타일 정의
  const containerStyle = {
    maxWidth: '1400px',
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

  const headerStyle = {
    color: '#0077B6',
    textAlign: 'center',
    marginBottom: '30px',
    paddingBottom: '10px',
    borderBottom: '2px solid #BDE0FE',
    fontSize: '2.5em',
    fontWeight: 'bold'
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
    padding: '8px 16px',
    borderRadius: '6px',
    fontSize: '14px',
    fontWeight: 'bold',
    cursor: 'pointer',
    margin: '2px',
    border: 'none',
    transition: 'all 0.3s ease'
  };

  const showButtonStyle = {
    ...buttonStyle,
    backgroundColor: '#28a745',
    color: 'white'
  };

  const noShowButtonStyle = {
    ...buttonStyle,
    backgroundColor: '#dc3545',
    color: 'white'
  };

  const pendingButtonStyle = {
    ...buttonStyle,
    backgroundColor: '#ffc107',
    color: '#212529'
  };

  const tableStyle = {
    width: '100%',
    borderCollapse: 'collapse',
    marginTop: '20px',
    backgroundColor: 'white',
    borderRadius: '8px',
    overflow: 'hidden',
    boxShadow: '0 4px 8px rgba(0, 0, 0, 0.1)'
  };

  const getStatusColor = (status: string, index: number): string => {
    switch (status) {
      case 'show': return '#d4edda'; // 초록색 배경
      case 'no-show': return '#f8d7da'; // 빨간색 배경
      case 'pending': return '#fff3cd'; // 노란색 배경
      default: return index % 2 === 0 ? '#f9f9f9' : 'white'; // 기본 교대로 배경색
    }
  };

  const getStatusText = (status: string): string => {
    switch (status) {
      case 'show': return 'Show ✅';
      case 'no-show': return 'No Show ❌';
      default: return 'Pending ⏳';
    }
  };

  return (
    <>
      <style>{`
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
      `}</style>
      <div style={bodyStyle}>
        {/* 로딩 모달 */}
        {pdfLoading && (
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
                {submitStatus === 'Saving...'}
                {submitStatus === 'Generating PDF...'}
                {submitStatus === 'Processing PDF...'}
                {submitStatus === 'Cleaning up...'}
                {submitStatus === 'Complete!'}
                {!submitStatus && 'Processing... Please wait'}
              </p>
              {/* 진행률 바 */}
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
                  ⚠️ Please do not close
                </p>
              </div>
            </div>
          </div>
        )}

        <div style={containerStyle}>
        {/* 헤더 */}
        <h1 style={{ 
          color: '#0077B6', 
          textAlign: 'center', 
          marginBottom: '20px', 
          fontSize: '2.5rem', 
          fontWeight: 'bold',
          textShadow: '2px 2px 4px rgba(0,0,0,0.1)'
        }}>📋 Appointment Show/No Show Check</h1>

        {/* Name 입력 섹션 */}
        <div style={sectionStyle}>
          <h2 style={{ color: '#0077B6', marginBottom: '20px' }}>👤 Name</h2>
          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap', marginBottom: '20px' }}>
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Name:
              </label>
              <input
                type="text"
                value={name}
                onChange={(e) => setName(e.target.value)}
                placeholder="Enter your name"
                style={inputStyle}
              />
            </div>
          </div>
        </div>

        {/* 필터 섹션 */}
        <div style={sectionStyle}>
          <h2 style={{ color: '#0077B6', marginBottom: '20px' }}>🔍 Filter Appointments</h2>
          
          <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap', marginBottom: '20px' }}>
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Start Date:
              </label>
              <input
                type="date"
                value={startDate}
                onChange={(e) => setStartDate(e.target.value)}
                style={inputStyle}
              />
            </div>
            
            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                End Date:
              </label>
              <input
                type="date"
                value={endDate}
                onChange={(e) => setEndDate(e.target.value)}
                style={inputStyle}
              />
            </div>

            <div style={{ flex: '1', minWidth: '200px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>
                Office:
              </label>
              <select
                value={selectedOffice}
                onChange={(e) => setSelectedOffice(e.target.value)}
                style={inputStyle}
              >
                {officeOptions.map((office: string) => (
                  <option key={office} value={office}>{office}</option>
                ))}
              </select>
            </div>
            

            <div style={{ flex: '1', minWidth: '200px', display: 'flex', alignItems: 'end' }}>
              <button 
                onClick={loadAppointments}
                disabled={loading}
                style={{
                  ...buttonStyle,
                  backgroundColor: '#0077B6',
                  color: 'white',
                  width: '100%'
                }}
              >
                {loading ? 'Loading...' : '🔄 Refresh Data'}
              </button>
            </div>
          </div>

          <div style={{ 
            padding: '10px', 
            backgroundColor: '#e3f2fd',
            borderRadius: '5px', 
            fontSize: '14px',
            color: '#1565c0'
          }}>
            📊 <strong>Total Appointments:</strong> {filteredAppointments.length}
            {startDate && endDate && ` | Date Range: ${startDate} to ${endDate}`}
            {selectedOffice && selectedOffice !== 'All' && ` | Office: ${selectedOffice}`}
          </div>
            </div>

        {/* 약속 목록 테이블 */}
        <div style={sectionStyle}>
          <h2 style={{ color: '#0077B6', marginBottom: '20px' }}>👥 Appointments to Check</h2>
          
          {loading ? (
            <div style={{ textAlign: 'center', padding: '40px', color: '#666' }}>
              Loading appointments...
            </div>
          ) : filteredAppointments.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '40px', color: '#666' }}>
              No appointments found for the selected criteria.
            </div>
          ) : (
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', margin: '10px 0' }}>
              <thead style={{ backgroundColor: '#0077B6', color: 'white' }}>
                <tr>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Patient Name</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Office</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Appt. Date</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Visit Type</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Status</th>
                    <th style={{ padding: '12px 8px', textAlign: 'center' }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                  {filteredAppointments.map((appointment: any, index: number) => (
                    <tr 
                      key={`${appointment.docId}-${appointment.rowIndex}`} 
                      style={{ 
                        backgroundColor: getStatusColor(appointment.showStatus, index),
                        opacity: appointment.showStatus === 'pending' ? 1 : 0.8
                      }}
                    >
                      <td style={{ padding: '12px 8px', fontWeight: 'bold' }}>
                        {appointment.name}
                    </td>
                      <td style={{ padding: '12px 8px', textAlign: 'center' }}>
                        {appointment.office}
                    </td>
                      <td style={{ padding: '12px 8px', textAlign: 'center' }}>
                        {appointment.appt_date}
                    </td>
                      <td style={{ padding: '12px 8px', textAlign: 'center' }}>
                        {appointment.visit_type}
                    </td>
                      <td style={{ padding: '12px 8px', textAlign: 'center', fontWeight: 'bold' }}>
                        {getStatusText(appointment.showStatus)}
                    </td>
                      <td style={{ padding: '12px 8px', textAlign: 'center' }}>
                        <div style={{ display: 'flex', gap: '4px', justifyContent: 'center' }}>
                          <button
                            onClick={() => updateShowStatus(appointment, 'show')}
                            style={showButtonStyle}
                            disabled={appointment.showStatus === 'show'}
                          >
                            Show
                          </button>
                          <button
                            onClick={() => updateShowStatus(appointment, 'no-show')}
                            style={noShowButtonStyle}
                            disabled={appointment.showStatus === 'no-show'}
                          >
                            No Show
                          </button>
                        <button
                            onClick={() => updateShowStatus(appointment, 'pending')}
                            style={pendingButtonStyle}
                            disabled={appointment.showStatus === 'pending'}
                          >
                            Pending
                        </button>
                        </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          )}
        </div>

        {/* 통계 섹션 */}
        {filteredAppointments.length > 0 && (
        <div style={sectionStyle}>
            <h2 style={{ color: '#0077B6', marginBottom: '20px' }}>📊 Statistics</h2>
            <div style={{ display: 'flex', gap: '20px', flexWrap: 'wrap' }}>
              {(() => {
                const showCount = filteredAppointments.filter(apt => apt.showStatus === 'show').length;
                const noShowCount = filteredAppointments.filter(apt => apt.showStatus === 'no-show').length;
                const pendingCount = filteredAppointments.filter(apt => apt.showStatus === 'pending').length;
                const showRate = filteredAppointments.length > 0 ? ((showCount / (showCount + noShowCount)) * 100).toFixed(1) : 0;
                
                return (
                  <>
                    <div style={{ 
                      flex: '1', 
                      minWidth: '150px', 
                      padding: '15px', 
                      backgroundColor: '#d4edda', 
                      borderRadius: '8px',
                      textAlign: 'center'
                    }}>
                      <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#155724' }}>{showCount}</div>
                      <div style={{ color: '#155724' }}>Show</div>
                    </div>
                    <div style={{ 
                      flex: '1', 
                      minWidth: '150px', 
                      padding: '15px', 
                      backgroundColor: '#f8d7da', 
                      borderRadius: '8px',
                      textAlign: 'center'
                    }}>
                      <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#721c24' }}>{noShowCount}</div>
                      <div style={{ color: '#721c24' }}>No Show</div>
                    </div>
                    <div style={{ 
                      flex: '1', 
                      minWidth: '150px', 
                      padding: '15px', 
                      backgroundColor: '#fff3cd', 
                      borderRadius: '8px',
                      textAlign: 'center'
                    }}>
                      <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#856404' }}>{pendingCount}</div>
                      <div style={{ color: '#856404' }}>Pending</div>
                    </div>
                    <div style={{ 
                      flex: '1', 
                      minWidth: '150px', 
                      padding: '15px', 
                      backgroundColor: '#e3f2fd', 
                      borderRadius: '8px',
                      textAlign: 'center'
                    }}>
                      <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#1565c0' }}>{showRate}%</div>
                      <div style={{ color: '#1565c0' }}>Show Rate</div>
                    </div>
                  </>
                );
              })()}
            </div>
        </div>
        )}

        {/* Submit 버튼 */}
        {filteredAppointments.length > 0 && (
          <div style={{textAlign: 'center', margin: '30px 0'}}>
          {(() => {
            // pending 상태가 있는지 체크 (actions 또는 showStatus)
            const hasPendingStatus = filteredAppointments.some(apt => 
              apt.actions === 'pending' || 
              apt.showStatus === 'pending' || 
              (!apt.actions && !apt.showStatus) ||
              (apt.actions !== 'show' && apt.actions !== 'no-show' && apt.showStatus !== 'show' && apt.showStatus !== 'no-show')
            );
            
            return (
              <button 
                onClick={handleGeneratePDF}
                disabled={pdfLoading || hasPendingStatus}
                style={{
                  backgroundColor: (pdfLoading || hasPendingStatus) ? '#9ca3af' : '#3b82f6',
                  color: 'white',
                  border: 'none',
                  padding: '15px 30px',
                  borderRadius: '8px',
                  fontSize: '16px',
                  fontWeight: 'bold',
                  cursor: (pdfLoading || hasPendingStatus) ? 'not-allowed' : 'pointer',
                  boxShadow: '0 4px 6px rgba(0, 0, 0, 0.1)',
                  transition: 'all 0.2s ease'
                }}
                title={hasPendingStatus ? '⚠️ Please select Show or No Show for all appointments before submitting.' : ''}
              >
                {pdfLoading ? '📄 Generating PDF...' : '📄 Submit + Generate PDF'}
              </button>
            );
          })()}
        </div>
        )}
      </div>
    </div>
    </>
  );
}
