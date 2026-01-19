'use client'

import React, { useState, useEffect, useCallback, useRef } from "react";

// 타입 정의
interface ReferralData {
  id: string;
  date: string;
  office: string;
  patientName: string;
  dob: string;
  type: string;
  insurance: string;
  behavior: string;
  medicalCondition: string;
  selectedNumbers: string;
  remarks: string;
  status: 'pending' | 'endo' | 'crown';
  createdAt: string;
  updatedAt?: string;
  uploadedFiles?: Array<{
    name: string;
    size: number;
    url: string;
  }>;
  generatedPdfUrl?: string; // 생성된 PDF URL 추가
  insuranceOther?: string;
  medicalConditionDetails?: string;
}

const BernardReferralShow = React.memo(() => {
  const [referrals, setReferrals] = useState<ReferralData[]>([]);
  const [loading, setLoading] = useState(true);
  const [initialized, setInitialized] = useState(false);
  const [filterStatus, setFilterStatus] = useState('all');
  const [filterOffice, setFilterOffice] = useState('all');
  const [filterDate, setFilterDate] = useState('');
  const [searchTerm, setSearchTerm] = useState('');
  const [exportLoading, setExportLoading] = useState(false);
  const [exportMonth, setExportMonth] = useState('');
  const [exportYear, setExportYear] = useState('');
  
  // 무한 루프 방지를 위한 ref
  const loadingRef = useRef(false);

  // Bernard Referral 데이터 로드
  const loadReferrals = useCallback(async () => {
    // 이미 로딩 중이면 중복 실행 방지
    if (loadingRef.current) {
      return;
    }
    
    try {
      loadingRef.current = true;
      setLoading(true);
      
      // 클라이언트 사이드에서만 실행
      if (typeof window === 'undefined') {
        console.warn("Firebase not available on server side");
        setReferrals([]);
        setLoading(false);
        setInitialized(true);
        return;
      }

      // Firebase 동적 import 시도
      try {
        const { initializeApp } = await import('firebase/app');
        const { getFirestore, collection, getDocs } = await import('firebase/firestore');
        
        // Firebase 설정 (환경변수 또는 기본값 사용)
        const firebaseConfig = {
          apiKey: process.env.NEXT_PUBLIC_FIREBASE_API_KEY || "demo-key",
          authDomain: process.env.NEXT_PUBLIC_FIREBASE_AUTH_DOMAIN || "demo-project.firebaseapp.com",
          projectId: process.env.NEXT_PUBLIC_FIREBASE_PROJECT_ID || "demo-project",
          storageBucket: process.env.NEXT_PUBLIC_FIREBASE_STORAGE_BUCKET || "demo-project.appspot.com",
          messagingSenderId: process.env.NEXT_PUBLIC_FIREBASE_MESSAGING_SENDER_ID || "123456789",
          appId: process.env.NEXT_PUBLIC_FIREBASE_APP_ID || "demo-app-id"
        };

        const app = initializeApp(firebaseConfig);
        const db = getFirestore(app);
        
        const querySnapshot = await getDocs(collection(db, "bernard-referral"));
        const referralData: ReferralData[] = [];
        
        querySnapshot.forEach((doc) => {
          referralData.push({
            id: doc.id,
            ...doc.data()
          } as ReferralData);
        });
        
        // 날짜순으로 정렬 (최신순)
        referralData.sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());
        setReferrals(referralData);
        
      } catch (firebaseError) {
        console.warn("Firebase not configured, using demo data:", firebaseError);
        // Firebase 설정이 없을 경우 데모 데이터 사용
        setReferrals([
          {
            id: "demo-1",
            date: "2024-01-15",
            office: "Bernard",
            patientName: "John Doe",
            dob: "01/15/1980",
            type: "Crown",
            insurance: "MC",
            behavior: "Good",
            medicalCondition: "No",
            selectedNumbers: "14,15",
            remarks: "Demo referral",
            status: "pending",
            createdAt: new Date().toISOString()
          }
        ]);
      }
    } catch (error) {
      console.error("Error loading referrals:", error);
      // 에러 발생 시 빈 배열로 설정하여 페이지가 로딩되도록 함
      setReferrals([]);
    } finally {
      setLoading(false);
      setInitialized(true);
      loadingRef.current = false;
    }
  }, []);

  // 저장된 PDF 열기 또는 새로 생성
  const generatePDF = useCallback(async (referral: ReferralData) => {
    try {
      console.log('=== PDF 열기 요청 ===');
      console.log('환자명:', referral.patientName);
      console.log('Document ID:', referral.id);
      console.log('저장된 PDF URL:', referral.generatedPdfUrl);
      console.log('첨부 파일 개수:', referral.uploadedFiles ? referral.uploadedFiles.length : 0);
      
      // 저장된 PDF가 있는 경우 직접 열기
      if (referral.generatedPdfUrl) {
        console.log('저장된 PDF 열기:', referral.generatedPdfUrl);
        window.open(referral.generatedPdfUrl, '_blank');
        return;
      }
      
      // 저장된 PDF가 없는 경우 새로 생성
      console.log('저장된 PDF가 없어서 새로 생성합니다.');
      
      if (referral.uploadedFiles && referral.uploadedFiles.length > 0) {
        referral.uploadedFiles.forEach((file, index) => {
          console.log(`첨부 파일 ${index + 1}:`, {
            name: file.name,
            url: file.url ? file.url.substring(0, 100) + '...' : 'No URL',
            size: file.size
          });
        });
      }
      
      // 데모 데이터의 경우 간단한 PDF 생성
      if (referral.id.startsWith('demo-')) {
        const demoData = {
          ...referral,
          uploadedFiles: [] // 데모 데이터에는 첨부 파일 없음
        };
        
        console.log('데모 데이터로 PDF 생성');
        
        const response = await fetch('/api/generate-bernard-referral-pdf', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(demoData)
        });

        if (response.ok) {
          const blob = await response.blob();
          const url = window.URL.createObjectURL(blob);
          window.open(url, '_blank');
          setTimeout(() => {
            window.URL.revokeObjectURL(url);
          }, 5000);
        } else {
          alert('Failed to generate PDF.');
        }
        return;
      }

      // 실제 데이터 처리 - Firebase Storage URL을 Base64로 변환
      const pdfData = {
        ...referral,
        uploadedFiles: referral.uploadedFiles ? await Promise.all(referral.uploadedFiles.map(async f => {
          console.log(`첨부 파일 매핑: ${f.name}`, {
            hasUrl: !!f.url,
            urlType: typeof f.url,
            urlPreview: f.url ? f.url.substring(0, 100) + '...' : 'No URL'
          });
          
          let fileData = f.url; // 기본값은 Firebase Storage URL
          
          // Firebase Storage URL이 있는 경우 Base64로 변환 시도
          if (f.url && f.url.includes('firebasestorage.googleapis.com')) {
            try {
              console.log(`Firebase Storage URL을 Base64로 변환 시도: ${f.name}`);
              
              // URL에서 토큰 제거하여 공개 URL로 변환
              const url = new URL(f.url);
              url.searchParams.delete('token');
              const publicUrl = url.toString();
              
              console.log(`공개 URL로 변환: ${publicUrl}`);
              
              // 파일 다운로드 시도
              const response = await fetch(publicUrl, {
                headers: {
                  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                }
              });
              
              if (response.ok) {
                const arrayBuffer = await response.arrayBuffer();
                const base64 = btoa(String.fromCharCode(...new Uint8Array(arrayBuffer)));
                const mimeType = response.headers.get('content-type') || 
                  (f.name.toLowerCase().endsWith('.pdf') ? 'application/pdf' : 'image/jpeg');
                fileData = `data:${mimeType};base64,${base64}`;
                console.log(`Base64 변환 완료: ${f.name}, 크기: ${base64.length} bytes`);
              } else {
                console.log(`Firebase Storage URL 접근 실패: ${f.name}, ${response.status} ${response.statusText}`);
                // 실패 시 원본 URL 사용
                fileData = f.url;
              }
            } catch (error) {
              console.log(`Firebase Storage URL 처리 오류: ${f.name}`, error.message);
              // 오류 시 원본 URL 사용
              fileData = f.url;
            }
          }
          
          return {
            name: f.name,
            size: f.size,
            data: fileData, // Base64 변환된 데이터 또는 원본 URL
            type: f.name.toLowerCase().endsWith('.pdf') ? 'application/pdf' : 'image'
          };
        })) : []
      };
      
      console.log('PDF 생성 요청 전송');
      
      const response = await fetch('/api/generate-bernard-referral-pdf', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(pdfData)
      });

      if (response.ok) {
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        // 새 탭에서 PDF 열기
        window.open(url, '_blank');
        // 5초 후 URL 정리 (메모리 누수 방지)
        setTimeout(() => {
          window.URL.revokeObjectURL(url);
        }, 5000);
      } else {
        alert('Failed to generate PDF.');
      }
    } catch (error) {
      console.error('Error generating PDF:', error);
      alert('An error occurred while generating PDF.');
    }
  }, []);

  useEffect(() => {
    // 초기화가 되지 않은 경우에만 로드
    if (!initialized && !loadingRef.current) {
      loadReferrals();
      
      // 기본값을 현재 월/년도로 설정
      const currentDate = new Date();
      setExportYear(currentDate.getFullYear().toString());
      setExportMonth((currentDate.getMonth() + 1).toString());
    }
  }, [initialized, loadReferrals]);



  useEffect(() => {
    // URL 파라미터에서 PDF 생성 요청 확인 (클라이언트 사이드에서만)
    if (typeof window !== 'undefined' && referrals.length > 0) {
      const urlParams = new URLSearchParams(window.location.search);
      const generatePDFId = urlParams.get('generatePDF');
      if (generatePDFId) {
        // 해당 ID의 referral 찾기
        const targetReferral = referrals.find(r => r.id === generatePDFId);
        if (targetReferral) {
          generatePDF(targetReferral);
          // URL에서 파라미터 제거
          window.history.replaceState({}, document.title, window.location.pathname);
        }
      }
    }
  }, [referrals, generatePDF]);

  // 상태 필터링
  const filteredReferrals = referrals.filter(referral => {
    const statusMatch = filterStatus === 'all' || referral.status === filterStatus;
    const officeMatch = filterOffice === 'all' || referral.office === filterOffice;
    const dateMatch = !filterDate || referral.date === filterDate;
    const searchMatch = !searchTerm || 
      referral.patientName.toLowerCase().includes(searchTerm.toLowerCase()) ||
      referral.type.toLowerCase().includes(searchTerm.toLowerCase());
    
    return statusMatch && officeMatch && dateMatch && searchMatch;
  });

  // 상태 업데이트
  const updateStatus = useCallback(async (id: string, newStatus: 'pending' | 'endo' | 'crown') => {
    try {
      // Firebase가 설정되지 않은 경우 로컬 상태만 업데이트
      if (id.startsWith('demo-')) {
        setReferrals(prev => prev.map(ref => 
          ref.id === id ? { ...ref, status: newStatus } : ref
        ));
        return;
      }

      // Firebase 업데이트 시도
      try {
        const { getFirestore, doc, updateDoc } = await import('firebase/firestore');
        const db = getFirestore();
        
        await updateDoc(doc(db, "bernard-referral", id), {
          status: newStatus,
          updatedAt: new Date().toISOString()
        });
        
        // 성공 시 로컬 상태도 업데이트
        setReferrals(prev => prev.map(ref => 
          ref.id === id ? { ...ref, status: newStatus } : ref
        ));
      } catch (firebaseError) {
        console.warn("Firebase update failed, updating local state only:", firebaseError);
        // Firebase 실패 시 로컬 상태만 업데이트
        setReferrals(prev => prev.map(ref => 
          ref.id === id ? { ...ref, status: newStatus } : ref
        ));
      }
    } catch (error) {
      console.error("Error updating status:", error);
    }
  }, []);

  // 삭제
  const deleteReferral = useCallback(async (id: string) => {
    if (confirm("Are you sure you want to delete this submission?")) {
      try {
        // Firebase가 설정되지 않은 경우 로컬 상태만 업데이트
        if (id.startsWith('demo-')) {
          setReferrals(prev => prev.filter(ref => ref.id !== id));
          return;
        }

        // Firebase 삭제 시도
        try {
          const { getFirestore, doc, deleteDoc } = await import('firebase/firestore');
          const db = getFirestore();
          
          await deleteDoc(doc(db, "bernard-referral", id));
          
          // 성공 시 로컬 상태도 업데이트
          setReferrals(prev => prev.filter(ref => ref.id !== id));
        } catch (firebaseError) {
          console.warn("Firebase delete failed, updating local state only:", firebaseError);
          // Firebase 실패 시 로컬 상태만 업데이트
          setReferrals(prev => prev.filter(ref => ref.id !== id));
        }
      } catch (error) {
        console.error("Error deleting referral:", error);
      }
    }
  }, []);

  // Excel 다운로드 기능
  const exportToExcel = async () => {
    try {
      setExportLoading(true);
      
      // 선택된 월과 년도로 데이터 필터링
      const selectedMonth = exportMonth || (new Date().getMonth() + 1);
      const selectedYear = exportYear || new Date().getFullYear();
      
      const monthlyReferrals = referrals.filter(referral => {
        const referralDate = new Date(referral.date);
        return referralDate.getMonth() + 1 === parseInt(selectedMonth.toString()) && 
               referralDate.getFullYear() === parseInt(selectedYear.toString());
      });

      // Excel 데이터 생성
      const excelData = monthlyReferrals.map((referral, index) => ({
        'No': index + 1,
        'Date': referral.date,
        'Office': referral.office,
        'Patient Name': referral.patientName,
        'DOB': referral.dob,
        'Type': referral.type,
        'Insurance': referral.insurance + (referral.insuranceOther ? ` (${referral.insuranceOther})` : ''),
        'Behavior': referral.behavior,
        'Medical Condition': referral.medicalCondition + (referral.medicalConditionDetails ? ` - ${referral.medicalConditionDetails}` : ''),
        'Selected Teeth': referral.selectedNumbers,
        'Remarks': referral.remarks,
        'Status': referral.status,
        'Created At': referral.createdAt ? new Date(referral.createdAt).toLocaleString() : '',
        'PDF URL': `${window.location.origin}/dashboard/forms/bernard-referral-show?generatePDF=${referral.id}`
      }));

      // CSV 형태로 변환
      const headers = Object.keys(excelData[0] || {});
      const csvContent = [
        headers.join(','),
        ...excelData.map(row => 
          headers.map(header => {
            const value = row[header] || '';
            // CSV에서 쉼표나 따옴표가 포함된 경우 처리
            return `"${value.toString().replace(/"/g, '""')}"`;
          }).join(',')
        )
      ].join('\n');

      // BOM 추가 (한글 지원)
      const BOM = '\uFEFF';
      const blob = new Blob([BOM + csvContent], { type: 'text/csv;charset=utf-8;' });
      
      // 파일 다운로드
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Bernard_Referral_${selectedYear}_${selectedMonth.toString().padStart(2, '0')}.csv`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
      
    } catch (error) {
      console.error('Excel export failed:', error);
      alert('An error occurred while downloading Excel file.');
    } finally {
      setExportLoading(false);
    }
  };

  // 상태별 색상
  const getStatusColor = (status: string) => {
    switch (status) {
      case 'pending': return 'bg-yellow-100 text-yellow-800';
      case 'endo': return 'bg-blue-100 text-blue-800';
      case 'crown': return 'bg-green-100 text-green-800';
      default: return 'bg-gray-100 text-gray-800';
    }
  };

  if (!initialized || loading) {
    return (
      <div className="flex justify-center items-center h-64">
        <div className="flex flex-col items-center gap-4">
          <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
          <div className="text-lg text-gray-600">Loading Referral data...</div>
        </div>
      </div>
    );
  }

  return (
    <div className="max-w-7xl mx-auto p-6">
      <div className="mb-8">
        <div className="flex justify-between items-center">
          <div>
            <h1 className="text-3xl font-bold text-gray-900 mb-2">Referral Submissions</h1>
          </div>
          <div className="flex items-center gap-3">
            <div className="flex items-center gap-2">
              <select
                value={exportYear}
                onChange={(e) => setExportYear(e.target.value)}
                className="px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-green-500"
              >
                <option value="">Year</option>
                {Array.from({ length: 5 }, (_, i) => {
                  const year = new Date().getFullYear() - i;
                  return <option key={year} value={year}>{year}</option>;
                })}
              </select>
              <select
                value={exportMonth}
                onChange={(e) => setExportMonth(e.target.value)}
                className="px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-green-500"
              >
                <option value="">Month</option>
                {Array.from({ length: 12 }, (_, i) => {
                  const month = i + 1;
                  const monthName = new Date(2024, i).toLocaleString('en-US', { month: 'long' });
                  return <option key={month} value={month}>{monthName}</option>;
                })}
              </select>
            </div>
            <button
              onClick={exportToExcel}
              disabled={exportLoading || referrals.length === 0}
              className="bg-green-600 hover:bg-green-700 disabled:bg-gray-400 text-white px-4 py-2 rounded-lg font-medium flex items-center gap-2"
            >
              {exportLoading ? (
                <>
                  <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-white"></div>
                  Exporting...
                </>
              ) : (
                <>
                  📊 Export to Excel
                </>
              )}
            </button>
          </div>
        </div>
      </div>

      {/* 필터 및 검색 */}
      <div className="bg-white p-6 rounded-lg shadow-sm border border-gray-200 mb-6">
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Status</label>
            <select
              value={filterStatus}
              onChange={(e) => setFilterStatus(e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
            >
              <option value="all">All</option>
              <option value="pending">Pending</option>
              <option value="endo">Endo</option>
              <option value="crown">Crown</option>
            </select>
          </div>
          
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Office</label>
            <select
              value={filterOffice}
              onChange={(e) => setFilterOffice(e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
            >
              <option value="all">All</option>
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
            <label className="block text-sm font-medium text-gray-700 mb-1">Date</label>
            <input
              type="date"
              value={filterDate}
              onChange={(e) => setFilterDate(e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>
          
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Search</label>
            <input
              type="text"
              placeholder="Search by patient name or type"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>
        </div>
      </div>

      {/* 통계 */}
      <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-6">
        <div className="bg-white p-4 rounded-lg shadow-sm border border-gray-200">
          <div className="text-2xl font-bold text-blue-600">{referrals.length}</div>
          <div className="text-sm text-gray-600">Total Submissions</div>
        </div>
        <div className="bg-white p-4 rounded-lg shadow-sm border border-gray-200">
          <div className="text-2xl font-bold text-yellow-600">
            {referrals.filter(r => r.status === 'pending').length}
          </div>
          <div className="text-sm text-gray-600">Pending</div>
        </div>
        <div className="bg-white p-4 rounded-lg shadow-sm border border-gray-200">
          <div className="text-2xl font-bold text-blue-600">
            {referrals.filter(r => r.status === 'endo').length}
          </div>
          <div className="text-sm text-gray-600">Endo</div>
        </div>
        <div className="bg-white p-4 rounded-lg shadow-sm border border-gray-200">
          <div className="text-2xl font-bold text-green-600">
            {referrals.filter(r => r.status === 'crown').length}
          </div>
          <div className="text-sm text-gray-600">Crown</div>
        </div>
      </div>

      {/* 제출 목록 */}
      <div className="bg-white rounded-lg shadow-sm border border-gray-200">
        <div className="overflow-x-auto">
          <table className="min-w-full divide-y divide-gray-200">
            <thead className="bg-gray-50">
              <tr>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Patient Name</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Type</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Office</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Date</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Status</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Actions</th>
              </tr>
            </thead>
            <tbody className="bg-white divide-y divide-gray-200">
              {filteredReferrals.map((referral) => (
                <tr key={referral.id} className="hover:bg-gray-50">
                  <td className="px-6 py-4 whitespace-nowrap">
                    <button
                      onClick={() => generatePDF(referral)}
                      className="text-sm font-medium text-blue-600 hover:text-blue-900 hover:underline cursor-pointer"
                    >
                      {referral.patientName}
                    </button>
                    <div className="text-sm text-gray-500">DOB: {referral.dob}</div>
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap">
                    <div className="text-sm text-gray-900">{referral.type}</div>
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap">
                    <div className="text-sm text-gray-900">{referral.office}</div>
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap">
                    <div className="text-sm text-gray-900">{referral.date}</div>
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap">
                    <span className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${getStatusColor(referral.status)}`}>
                      {referral.status === 'pending' ? 'Pending' : 
                       referral.status === 'endo' ? 'Endo' : 
                       referral.status === 'crown' ? 'Crown' : referral.status}
                    </span>
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap text-sm font-medium">
                    <div className="flex space-x-2">
                      <button
                        onClick={() => generatePDF(referral)}
                        className="text-blue-600 hover:text-blue-900"
                      >
                        PDF
                      </button>
                      <select
                        value={referral.status}
                        onChange={(e) => updateStatus(referral.id, e.target.value as 'pending' | 'endo' | 'crown')}
                        className="text-sm border border-gray-300 rounded px-2 py-1"
                      >
                        <option value="pending">Pending</option>
                        <option value="endo">Endo</option>
                        <option value="crown">Crown</option>
                      </select>
                      <button
                        onClick={() => deleteReferral(referral.id)}
                        className="text-red-600 hover:text-red-900"
                      >
                        Delete
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        
        {filteredReferrals.length === 0 && (
          <div className="text-center py-8 text-gray-500">
            {referrals.length === 0 ? 'No Bernard Referral submissions found.' : 'No submissions match the filter criteria.'}
          </div>
        )}
      </div>
    </div>
  );
});

export default BernardReferralShow;