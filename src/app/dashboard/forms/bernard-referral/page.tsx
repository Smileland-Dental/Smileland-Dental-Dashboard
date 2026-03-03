'use client'

import React, { useState, useEffect } from "react";
import { doc, setDoc, collection, getDocs } from "firebase/firestore";
import { ref, uploadBytes, getDownloadURL } from "firebase/storage";
import { db, storage } from "@/lib/firebase.config";

export default function BernardReferralForm() {
  // 폼 데이터 상태
  const [formData, setFormData] = useState({
    date: new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' }),
    office: '',
    type: '',
    patientName: '',
    dob: '',
    insurance: '',
    insuranceOther: '',
    behavior: '',
    medicalCondition: '',
    medicalConditionDetails: '',
    selectedNumbers: '',
    remarks: '',
  });

  const [uploadedFiles, setUploadedFiles] = useState([]);
  const [loading, setLoading] = useState(false);
  const [submitStatus, setSubmitStatus] = useState('');
  const [selectedNumbers, setSelectedNumbers] = useState(new Set());
  const [progress, setProgress] = useState(0);

  // 페이지 이탈 방지 (제출 중일 때)
  useEffect(() => {
    const handleBeforeUnload = (e) => {
      if (loading) {
        e.preventDefault();
        e.returnValue = '제출이 진행 중입니다. 정말로 페이지를 떠나시겠습니까?';
        return '제출이 진행 중입니다. 정말로 페이지를 떠나시겠습니까?';
      }
    };

    const handlePopState = (e) => {
      if (loading) {
        e.preventDefault();
        window.history.pushState(null, '', window.location.href);
        alert('제출이 진행 중입니다. 완료 후에 페이지를 떠나주세요.');
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

  // Document ID 생성 함수 (고유성 보장)
  const generateDocId = (date, patientName, type) => {
    const timestamp = Date.now();
    const randomId = Math.random().toString(36).substring(2, 8);
    return `${date}_${patientName}_${type}_${timestamp}_${randomId}`.replace(/[^a-zA-Z0-9_-]/g, '_');
  };

  // 자동 저장 기능 제거됨 - Submit 버튼으로만 제출

  // 폼 데이터 업데이트
  const updateFormData = (field, value) => {
    setFormData(prev => ({
      ...prev,
      [field]: value
    }));
  };

  // 치아 번호 선택 처리
  const handleNumberClick = (number) => {
    const newSelectedNumbers = new Set(selectedNumbers);
    if (newSelectedNumbers.has(number)) {
      newSelectedNumbers.delete(number);
    } else {
      newSelectedNumbers.add(number);
    }
    setSelectedNumbers(newSelectedNumbers);
    updateFormData('selectedNumbers', Array.from(newSelectedNumbers).sort((a, b) => a - b).join(','));
  };

  // 파일 업로드 처리
  const handleFileChange = (e) => {
    const files = Array.from(e.target.files);
    
    files.forEach((file) => {
      if (file.type.startsWith('image/')) {
        if (file.size > 50 * 1024 * 1024) {
          alert(`파일이 너무 큽니다: ${file.name} (${(file.size / 1024 / 1024).toFixed(1)}MB)\n50MB 이하의 파일만 업로드 가능합니다.`);
          return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
          const fileData = {
            name: file.name,
            type: file.type,
            size: file.size,
            data: e.target.result,
            file: file
          };
          setUploadedFiles(prev => [...prev, fileData]);
        };
        reader.readAsDataURL(file);
      } else {
        alert(`지원하지 않는 파일 형식입니다: ${file.name}\n이미지 파일(JPG, PNG, GIF)만 업로드 가능합니다.`);
      }
    });
  };

  const removeFile = (index) => {
    setUploadedFiles(prev => prev.filter((_, i) => i !== index));
  };

  // Firebase에 저장
  const handleSave = async (docId?: string) => {
    try {
      const documentId = docId || generateDocId(formData.date, formData.patientName, formData.type);
      
      // 파일 업로드 (큰 파일 지원)
      const uploadedFileUrls = [];
      for (const fileData of uploadedFiles) {
        try {
          // 파일명에 타임스탬프 추가하여 고유성 보장
          const timestamp = Date.now();
          const fileExtension = fileData.name.split('.').pop();
          const fileNameWithoutExt = fileData.name.replace(/\.[^/.]+$/, "");
          const uniqueFileName = `${fileNameWithoutExt}_${timestamp}.${fileExtension}`;
          const fileRef = ref(storage, `bernard-referral/${documentId}/${uniqueFileName}`);
          
          // 파일 데이터 확인 및 로깅
          console.log(`파일 업로드 처리: ${fileData.name}`);
          console.log('파일 데이터 상태:', {
            hasData: !!fileData.data,
            hasFile: !!fileData.file,
            dataType: typeof fileData.data,
            fileType: typeof fileData.file,
            dataLength: fileData.data ? fileData.data.length : 0
          });
          
          // 파일 타입에 따른 처리
          let fileBlob;
          if (fileData.file && fileData.file instanceof File) {
            // File 객체가 있는 경우 (PDF 또는 이미지)
            console.log(`File 객체 사용: ${fileData.name} (타입: ${fileData.type})`);
            fileBlob = fileData.file;
          } else if (fileData.data && fileData.data.startsWith('data:')) {
            // Base64 데이터가 있는 경우 (이미지)
            console.log(`Base64 데이터 사용: ${fileData.name}`);
            fileBlob = await fetch(fileData.data).then(r => r.blob());
          } else if (fileData.data && fileData.data.startsWith('http')) {
            // URL 데이터인 경우
            console.log(`URL 데이터 사용: ${fileData.name}`);
            fileBlob = await fetch(fileData.data).then(r => r.blob());
          } else {
            console.warn(`No valid data available for file: ${fileData.name}`, {
              hasData: !!fileData.data,
              hasFile: !!fileData.file,
              dataType: typeof fileData.data,
              fileType: typeof fileData.file,
              dataPreview: fileData.data ? fileData.data.substring(0, 100) : 'No data'
            });
            continue;
          }
          
          // 큰 파일의 경우 청크 업로드 사용
          if (fileBlob.size > 32 * 1024 * 1024) { // 32MB 이상
            console.log(`Large file detected: ${fileData.name} (${(fileBlob.size / 1024 / 1024).toFixed(1)}MB)`);
          }
          
          console.log(`Firebase Storage 업로드 시작: ${fileData.name}`);
          await uploadBytes(fileRef, fileBlob);
          const downloadURL = await getDownloadURL(fileRef);
          console.log(`Firebase Storage 업로드 성공: ${fileData.name}`, {
            url: downloadURL.substring(0, 100) + '...',
            size: fileBlob.size
          });
          
          uploadedFileUrls.push({
            name: fileData.name, // 원본 파일명 유지
            url: downloadURL,
            size: fileData.size,
            uniqueFileName: uniqueFileName // 고유 파일명도 저장
          });
        } catch (error) {
          console.error(`Error uploading file ${fileData.name}:`, error);
          // 개별 파일 업로드 실패 시에도 계속 진행
        }
      }

      const documentData = {
        ...formData,
        selectedNumbers: Array.from(selectedNumbers).sort((a, b) => a - b).join(','),
        uploadedFiles: uploadedFileUrls,
        status: 'pending',
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      };

      console.log('=== 데이터 저장 정보 ===');
      console.log('Document ID:', documentId);
      console.log('환자명:', formData.patientName);
      console.log('첨부 파일 개수:', uploadedFileUrls.length);
      console.log('원본 업로드된 파일 개수:', uploadedFiles.length);
      
      uploadedFiles.forEach((file, index) => {
        console.log(`원본 파일 ${index + 1}:`, {
          name: file.name,
          type: file.type,
          size: file.size,
          hasData: !!file.data,
          hasFile: !!file.file,
          dataType: typeof file.data
        });
      });
      
      uploadedFileUrls.forEach((file, index) => {
        console.log(`업로드된 파일 ${index + 1}:`, {
          originalName: file.name,
          uniqueFileName: file.uniqueFileName,
          url: file.url.substring(0, 100) + '...',
          size: file.size
        });
      });

      await setDoc(doc(db, "bernard-referral", documentId), documentData);
    } catch (error) {
      console.error("Save failed:", error);
      throw error; // 에러를 다시 던져서 handleSubmit에서 처리할 수 있도록 함
    }
  };

  // PDF 생성 및 제출
  const handleSubmit = async () => {
    try {
      // 필수 필드 검증
      if (!formData.patientName.trim()) {
        alert('환자 이름을 입력해주세요.');
        return;
      }
      if (!formData.office) {
        alert('오피스를 선택해주세요.');
        return;
      }
      if (!formData.type) {
        alert('치료 유형을 선택해주세요.');
        return;
      }
      
      setLoading(true);
      setProgress(10);
      setSubmitStatus('Saving...');
      
      // Document ID 생성
      const docId = generateDocId(formData.date, formData.patientName, formData.type);
      
      // 먼저 데이터 저장
      await handleSave(docId);
      
      setProgress(30);
      setSubmitStatus('Generating PDF...');
      
      // PDF 생성 API 호출
      const response = await fetch('/api/generate-bernard-referral-pdf', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          ...formData,
          selectedNumbers: Array.from(selectedNumbers).sort((a, b) => a - b).join(','),
          uploadedFiles: uploadedFiles.map(f => ({ 
            name: f.name, 
            size: f.size, 
            data: f.data, // 이미지의 경우 base64, PDF의 경우 null
            type: f.type // 파일 타입 추가
          }))
        })
      });

      if (response.ok) {
        setProgress(60);
        setSubmitStatus('Uploading PDF...');
        const blob = await response.blob();
        
        // 생성된 PDF를 Firebase Storage에 저장
        const pdfFileName = `generated_pdf_${Date.now()}.pdf`;
        const pdfRef = ref(storage, `bernard-referral/${docId}/${pdfFileName}`);
        await uploadBytes(pdfRef, blob);
        const pdfUrl = await getDownloadURL(pdfRef);
        
        console.log('PDF가 Firebase Storage에 저장됨:', pdfUrl);
        
        setProgress(80);
        setSubmitStatus('Finalizing...');
        
        // Firestore에 PDF URL 업데이트
        await setDoc(doc(db, 'bernard-referral', docId), {
          generatedPdfUrl: pdfUrl
        }, { merge: true });
        
        setProgress(90);
        
        // 새 탭에서 PDF 열기
        const url = window.URL.createObjectURL(blob);
        window.open(url, '_blank');
        // 5초 후 URL 정리 (메모리 누수 방지)
        setTimeout(() => {
          window.URL.revokeObjectURL(url);
        }, 5000);
        
        setProgress(100);
        setSubmitStatus('Complete!');
        
        // 완료 메시지를 잠시 보여준 후 폼 초기화
        setTimeout(() => {
          // 폼 초기화
          setFormData({
            date: new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' }),
            office: '',
            type: '',
            patientName: '',
            dob: '',
            insurance: '',
            insuranceOther: '',
            behavior: '',
            medicalCondition: '',
            medicalConditionDetails: '',
            selectedNumbers: '',
            remarks: '',
          });
          setSelectedNumbers(new Set());
          setUploadedFiles([]);
          
          // 모달 닫기
          setLoading(false);
          setSubmitStatus('');
          setProgress(0);
        }, 2000); // 2초 후에 모달 닫기
      } else {
        const errorData = await response.json();
        console.error('PDF 생성 실패:', errorData);
        alert(`PDF 생성에 실패했습니다: ${errorData.error || '알 수 없는 오류'}`);
      }
    } catch (error) {
      console.error('Submit failed:', error);
      alert('제출 중 오류가 발생했습니다.');
      // 에러 발생 시에만 모달 닫기
      setLoading(false);
      setSubmitStatus('');
      setProgress(0);
    }
  };

  return (
    <div style={{
      fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
      backgroundColor: "#f0f8f8",
      minHeight: "100vh",
      display: "flex",
      justifyContent: "center",
      alignItems: "center",
      padding: "20px"
    }}>
      <div style={{
        background: "white",
        borderRadius: "15px",
        boxShadow: "0 15px 30px rgba(0, 0, 0, 0.1)",
        padding: "25px",
        width: "100%",
        maxWidth: "1200px",
        animation: "fadeInUp 0.6s ease-out"
      }}>
        <h1 style={{
          textAlign: "center",
          color: "#000000",
          fontSize: "2rem",
          fontWeight: "700",
          marginBottom: "25px"
        }}>Referral</h1>

        <form onSubmit={(e) => { e.preventDefault(); handleSubmit(); }}>
          <div style={{ display: "flex", gap: "12px", marginBottom: "18px" }}>
            <div style={{ flex: 1 }}>
              <label style={{
                display: "block",
                marginBottom: "6px",
                color: "#555",
                fontWeight: "600",
                fontSize: "0.9rem"
              }}>Date</label>
              <input
                type="date"
                value={formData.date}
                onChange={(e) => updateFormData('date', e.target.value)}
                style={{
                  width: "100%",
                  padding: "10px 12px",
                  border: "2px solid #e1e5e9",
                  borderRadius: "8px",
                  fontSize: "0.9rem",
                  background: "#f8f9fa"
                }}
                required
              />
            </div>

            <div style={{ flex: 1 }}>
              <label style={{
                display: "block",
                marginBottom: "6px",
                color: "#555",
                fontWeight: "600",
                fontSize: "0.9rem"
              }}>Office</label>
              <select
                value={formData.office}
                onChange={(e) => updateFormData('office', e.target.value)}
                style={{
                  width: "100%",
                  padding: "10px 12px",
                  border: "2px solid #e1e5e9",
                  borderRadius: "8px",
                  fontSize: "0.9rem",
                  background: "#f8f9fa"
                }}
                required
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

            <div style={{ flex: 1 }}>
              <label style={{
                display: "block",
                marginBottom: "6px",
                color: "#555",
                fontWeight: "600",
                fontSize: "0.9rem"
              }}>Type</label>
              <div style={{ display: "flex", gap: "10px", marginTop: "8px" }}>
                <div style={{ flex: 1, position: "relative" }}>
                  <input
                    type="radio"
                    id="Crown"
                    name="type"
                    value="Crown"
                    checked={formData.type === 'Crown'}
                    onChange={(e) => updateFormData('type', e.target.value)}
                    style={{ position: "absolute", opacity: 0, pointerEvents: "none" }}
                    required
                  />
                  <label htmlFor="Crown" style={{
                    display: "block",
                    padding: "10px 12px",
                    textAlign: "center",
                    background: formData.type === 'Crown' ? "linear-gradient(135deg, #4a90e2, #357abd)" : "#f8f9fa",
                    color: formData.type === 'Crown' ? "white" : "#555",
                    border: "2px solid #e1e5e9",
                    borderRadius: "8px",
                    cursor: "pointer",
                    fontWeight: "600",
                    fontSize: "0.9rem"
                  }}>Crown</label>
                </div>
                <div style={{ flex: 1, position: "relative" }}>
                  <input
                    type="radio"
                    id="rct"
                    name="type"
                    value="RCT"
                    checked={formData.type === 'RCT'}
                    onChange={(e) => updateFormData('type', e.target.value)}
                    style={{ position: "absolute", opacity: 0, pointerEvents: "none" }}
                    required
                  />
                  <label htmlFor="rct" style={{
                    display: "block",
                    padding: "10px 12px",
                    textAlign: "center",
                    background: formData.type === 'RCT' ? "linear-gradient(135deg, #4a90e2, #357abd)" : "#f8f9fa",
                    color: formData.type === 'RCT' ? "white" : "#555",
                    border: "2px solid #e1e5e9",
                    borderRadius: "8px",
                    cursor: "pointer",
                    fontWeight: "600",
                    fontSize: "0.9rem"
                  }}>RCT</label>
                </div>
              </div>
            </div>
          </div>

          <div style={{ display: "flex", gap: "12px", marginBottom: "18px" }}>
            <div style={{ flex: 1 }}>
              <label style={{
                display: "block",
                marginBottom: "6px",
                color: "#555",
                fontWeight: "600",
                fontSize: "0.9rem"
              }}>Patient Name</label>
              <input
                type="text"
                value={formData.patientName}
                onChange={(e) => updateFormData('patientName', e.target.value)}
                placeholder="Enter patient name"
                style={{
                  width: "100%",
                  padding: "10px 12px",
                  border: "2px solid #e1e5e9",
                  borderRadius: "8px",
                  fontSize: "0.9rem",
                  background: "#f8f9fa"
                }}
                required
              />
            </div>

            <div style={{ flex: 1 }}>
              <label style={{
                display: "block",
                marginBottom: "6px",
                color: "#555",
                fontWeight: "600",
                fontSize: "0.9rem"
              }}>DOB</label>
              <input
                type="text"
                value={formData.dob}
                onChange={(e) => updateFormData('dob', e.target.value)}
                placeholder="mm/dd/yyyy"
                style={{
                  width: "100%",
                  padding: "10px 12px",
                  border: "2px solid #e1e5e9",
                  borderRadius: "8px",
                  fontSize: "0.9rem",
                  background: "#f8f9fa"
                }}
                required
              />
            </div>
          </div>

          <div style={{ display: "flex", gap: "12px", marginBottom: "18px" }}>
            <div style={{ flex: 1 }}>
              <label style={{
                display: "block",
                marginBottom: "6px",
                color: "#555",
                fontWeight: "600",
                fontSize: "0.9rem"
              }}>Insurance</label>
              <div style={{ display: "flex", gap: "10px", marginTop: "8px" }}>
                <div style={{ flex: 1, position: "relative" }}>
                  <input
                    type="radio"
                    id="mc"
                    name="insurance"
                    value="MC"
                    checked={formData.insurance === 'MC'}
                    onChange={(e) => updateFormData('insurance', e.target.value)}
                    style={{ position: "absolute", opacity: 0, pointerEvents: "none" }}
                    required
                  />
                  <label htmlFor="mc" style={{
                    display: "block",
                    padding: "10px 12px",
                    textAlign: "center",
                    background: formData.insurance === 'MC' ? "linear-gradient(135deg, #4a90e2, #357abd)" : "#f8f9fa",
                    color: formData.insurance === 'MC' ? "white" : "#555",
                    border: "2px solid #e1e5e9",
                    borderRadius: "8px",
                    cursor: "pointer",
                    fontWeight: "600",
                    fontSize: "0.9rem"
                  }}>MC</label>
                </div>
                <div style={{ flex: 1, position: "relative" }}>
                  <input
                    type="radio"
                    id="other"
                    name="insurance"
                    value="Other"
                    checked={formData.insurance === 'Other'}
                    onChange={(e) => updateFormData('insurance', e.target.value)}
                    style={{ position: "absolute", opacity: 0, pointerEvents: "none" }}
                    required
                  />
                  <label htmlFor="other" style={{
                    display: "block",
                    padding: "10px 12px",
                    textAlign: "center",
                    background: formData.insurance === 'Other' ? "linear-gradient(135deg, #4a90e2, #357abd)" : "#f8f9fa",
                    color: formData.insurance === 'Other' ? "white" : "#555",
                    border: "2px solid #e1e5e9",
                    borderRadius: "8px",
                    cursor: "pointer",
                    fontWeight: "600",
                    fontSize: "0.9rem"
                  }}>Other</label>
                </div>
              </div>
              {formData.insurance === 'Other' && (
                <input
                  type="text"
                  value={formData.insuranceOther}
                  onChange={(e) => updateFormData('insuranceOther', e.target.value)}
                  placeholder="Enter insurance details"
                  style={{
                    width: "100%",
                    padding: "10px 12px",
                    border: "2px solid #e1e5e9",
                    borderRadius: "8px",
                    fontSize: "0.9rem",
                    background: "#f8f9fa",
                    marginTop: "10px"
                  }}
                  required
                />
              )}
            </div>

            <div style={{ flex: 1 }}>
              <label style={{
                display: "block",
                marginBottom: "6px",
                color: "#555",
                fontWeight: "600",
                fontSize: "0.9rem"
              }}>Behavior</label>
              <div style={{ display: "flex", gap: "10px", marginTop: "8px" }}>
                <div style={{ flex: 1, position: "relative" }}>
                  <input
                    type="radio"
                    id="good"
                    name="behavior"
                    value="Good"
                    checked={formData.behavior === 'Good'}
                    onChange={(e) => updateFormData('behavior', e.target.value)}
                    style={{ position: "absolute", opacity: 0, pointerEvents: "none" }}
                    required
                  />
                  <label htmlFor="good" style={{
                    display: "block",
                    padding: "10px 12px",
                    textAlign: "center",
                    background: formData.behavior === 'Good' ? "linear-gradient(135deg, #4a90e2, #357abd)" : "#f8f9fa",
                    color: formData.behavior === 'Good' ? "white" : "#555",
                    border: "2px solid #e1e5e9",
                    borderRadius: "8px",
                    cursor: "pointer",
                    fontWeight: "600",
                    fontSize: "0.9rem"
                  }}>Good</label>
                </div>
                <div style={{ flex: 1, position: "relative" }}>
                  <input
                    type="radio"
                    id="excellent"
                    name="behavior"
                    value="Excellent"
                    checked={formData.behavior === 'Excellent'}
                    onChange={(e) => updateFormData('behavior', e.target.value)}
                    style={{ position: "absolute", opacity: 0, pointerEvents: "none" }}
                    required
                  />
                  <label htmlFor="excellent" style={{
                    display: "block",
                    padding: "10px 12px",
                    textAlign: "center",
                    background: formData.behavior === 'Excellent' ? "linear-gradient(135deg, #4a90e2, #357abd)" : "#f8f9fa",
                    color: formData.behavior === 'Excellent' ? "white" : "#555",
                    border: "2px solid #e1e5e9",
                    borderRadius: "8px",
                    cursor: "pointer",
                    fontWeight: "600",
                    fontSize: "0.9rem"
                  }}>Excellent</label>
                </div>
              </div>
            </div>

            <div style={{ flex: 1 }}>
              <label style={{
                display: "block",
                marginBottom: "6px",
                color: "#555",
                fontWeight: "600",
                fontSize: "0.9rem"
              }}>Medical Condition</label>
              <div style={{ display: "flex", gap: "10px", marginTop: "8px" }}>
                <div style={{ flex: 1, position: "relative" }}>
                  <input
                    type="radio"
                    id="noCondition"
                    name="medicalCondition"
                    value="No"
                    checked={formData.medicalCondition === 'No'}
                    onChange={(e) => updateFormData('medicalCondition', e.target.value)}
                    style={{ position: "absolute", opacity: 0, pointerEvents: "none" }}
                    required
                  />
                  <label htmlFor="noCondition" style={{
                    display: "block",
                    padding: "10px 12px",
                    textAlign: "center",
                    background: formData.medicalCondition === 'No' ? "linear-gradient(135deg, #4a90e2, #357abd)" : "#f8f9fa",
                    color: formData.medicalCondition === 'No' ? "white" : "#555",
                    border: "2px solid #e1e5e9",
                    borderRadius: "8px",
                    cursor: "pointer",
                    fontWeight: "600",
                    fontSize: "0.9rem"
                  }}>No</label>
                </div>
                <div style={{ flex: 1, position: "relative" }}>
                  <input
                    type="radio"
                    id="yesCondition"
                    name="medicalCondition"
                    value="Yes"
                    checked={formData.medicalCondition === 'Yes'}
                    onChange={(e) => updateFormData('medicalCondition', e.target.value)}
                    style={{ position: "absolute", opacity: 0, pointerEvents: "none" }}
                    required
                  />
                  <label htmlFor="yesCondition" style={{
                    display: "block",
                    padding: "10px 12px",
                    textAlign: "center",
                    background: formData.medicalCondition === 'Yes' ? "linear-gradient(135deg, #4a90e2, #357abd)" : "#f8f9fa",
                    color: formData.medicalCondition === 'Yes' ? "white" : "#555",
                    border: "2px solid #e1e5e9",
                    borderRadius: "8px",
                    cursor: "pointer",
                    fontWeight: "600",
                    fontSize: "0.9rem"
                  }}>Yes</label>
                </div>
              </div>
              {formData.medicalCondition === 'Yes' && (
                <input
                  type="text"
                  value={formData.medicalConditionDetails}
                  onChange={(e) => updateFormData('medicalConditionDetails', e.target.value)}
                  placeholder="Enter medical condition details"
                  style={{
                    width: "100%",
                    padding: "10px 12px",
                    border: "2px solid #e1e5e9",
                    borderRadius: "8px",
                    fontSize: "0.9rem",
                    background: "#f8f9fa",
                    marginTop: "10px"
                  }}
                  required
                />
              )}
            </div>
          </div>

          <div style={{ marginBottom: "18px" }}>
            <label style={{
              display: "block",
              marginBottom: "6px",
              color: "#555",
              fontWeight: "600",
              fontSize: "0.9rem"
            }}>Select Tooth Number</label>
            <div style={{ marginTop: "10px", position: "relative" }}>
              <div style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                marginBottom: "8px",
                padding: "0 15px",
                position: "absolute",
                top: "50%",
                transform: "translateY(-50%)",
                width: "100%",
                pointerEvents: "none"
              }}>
                <span style={{ fontWeight: "700", fontSize: "1rem", color: "#4a90e2", margin: "0 6px" }}>R</span>
                <span style={{ fontWeight: "700", fontSize: "1rem", color: "#4a90e2", margin: "0 6px" }}>L</span>
              </div>
              <div style={{ marginBottom: "10px" }}>
                <div style={{ display: "flex", flexWrap: "wrap", gap: "5px", justifyContent: "center" }}>
                  {[1,2,3,4,5,6,7,8].map(num => (
                    <button
                      key={num}
                      type="button"
                      onClick={() => handleNumberClick(num.toString())}
                      style={{
                        width: "38px",
                        height: "38px",
                        border: "2px solid #e1e5e9",
                        background: selectedNumbers.has(num.toString()) ? "linear-gradient(135deg, #4a90e2, #357abd)" : "#f8f9fa",
                        color: selectedNumbers.has(num.toString()) ? "white" : "#555",
                        borderRadius: "6px",
                        fontSize: "13px",
                        fontWeight: "600",
                        cursor: "pointer",
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center"
                      }}
                    >
                      {num}
                    </button>
                  ))}
                  <div style={{ width: "2px", height: "38px", background: "#4a90e2", margin: "0 3px", borderRadius: "1px" }}></div>
                  {[9,10,11,12,13,14,15,16].map(num => (
                    <button
                      key={num}
                      type="button"
                      onClick={() => handleNumberClick(num.toString())}
                      style={{
                        width: "38px",
                        height: "38px",
                        border: "2px solid #e1e5e9",
                        background: selectedNumbers.has(num.toString()) ? "linear-gradient(135deg, #4a90e2, #357abd)" : "#f8f9fa",
                        color: selectedNumbers.has(num.toString()) ? "white" : "#555",
                        borderRadius: "6px",
                        fontSize: "13px",
                        fontWeight: "600",
                        cursor: "pointer",
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center"
                      }}
                    >
                      {num}
                    </button>
                  ))}
                </div>
              </div>
              <div style={{ width: "calc(16 * 38px + 15 * 5px + 2 * 2px)", height: "2px", background: "#4a90e2", margin: "6px auto", borderRadius: "1px" }}></div>
              <div style={{ marginBottom: "10px" }}>
                <div style={{ display: "flex", flexWrap: "wrap", gap: "5px", justifyContent: "center" }}>
                  {[32,31,30,29,28,27,26,25].map(num => (
                    <button
                      key={num}
                      type="button"
                      onClick={() => handleNumberClick(num.toString())}
                      style={{
                        width: "38px",
                        height: "38px",
                        border: "2px solid #e1e5e9",
                        background: selectedNumbers.has(num.toString()) ? "linear-gradient(135deg, #4a90e2, #357abd)" : "#f8f9fa",
                        color: selectedNumbers.has(num.toString()) ? "white" : "#555",
                        borderRadius: "6px",
                        fontSize: "13px",
                        fontWeight: "600",
                        cursor: "pointer",
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center"
                      }}
                    >
                      {num}
                    </button>
                  ))}
                  <div style={{ width: "2px", height: "38px", background: "#4a90e2", margin: "0 3px", borderRadius: "1px" }}></div>
                  {[24,23,22,21,20,19,18,17].map(num => (
                    <button
                      key={num}
                      type="button"
                      onClick={() => handleNumberClick(num.toString())}
                      style={{
                        width: "38px",
                        height: "38px",
                        border: "2px solid #e1e5e9",
                        background: selectedNumbers.has(num.toString()) ? "linear-gradient(135deg, #4a90e2, #357abd)" : "#f8f9fa",
                        color: selectedNumbers.has(num.toString()) ? "white" : "#555",
                        borderRadius: "6px",
                        fontSize: "13px",
                        fontWeight: "600",
                        cursor: "pointer",
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center"
                      }}
                    >
                      {num}
                    </button>
                  ))}
                </div>
              </div>
            </div>
          </div>

          <div style={{ marginBottom: "18px" }}>
            <label style={{
              display: "block",
              marginBottom: "6px",
              color: "#555",
              fontWeight: "600",
              fontSize: "0.9rem"
            }}>Remarks (Dr. please note)</label>
            <textarea
              value={formData.remarks}
              onChange={(e) => updateFormData('remarks', e.target.value)}
              rows={4}
              style={{
                width: "100%",
                padding: "10px 12px",
                border: "2px solid #e1e5e9",
                borderRadius: "8px",
                fontSize: "0.9rem",
                background: "#f8f9fa",
                resize: "vertical",
                minHeight: "100px"
              }}
            />
          </div>

          <div style={{ marginTop: "15px" }}>
            <label style={{
              display: "block",
              marginBottom: "6px",
              color: "#555",
              fontWeight: "600",
              fontSize: "0.9rem"
            }}>Attach Files (Images Only) - Max 50MB per file</label>
            <input
              type="file"
              onChange={handleFileChange}
              accept="image/*"
              multiple
              style={{ display: "none" }}
              id="photoUpload"
            />
            <button
              type="button"
              onClick={() => document.getElementById('photoUpload').click()}
              style={{
                display: "inline-block",
                padding: "10px 20px",
                background: "linear-gradient(135deg, #4a90e2, #357abd)",
                color: "white",
                border: "none",
                borderRadius: "8px",
                cursor: "pointer",
                fontSize: "0.9rem",
                fontWeight: "600",
                marginBottom: "10px"
              }}
            >
              📷 Choose Files
            </button>
            
            <div style={{ display: "flex", flexWrap: "wrap", gap: "10px", marginTop: "10px" }}>
              {uploadedFiles.map((file, index) => (
                <div key={index} style={{
                  position: "relative",
                  width: "120px",
                  height: "120px",
                  border: "2px solid #e1e5e9",
                  borderRadius: "8px",
                  overflow: "hidden",
                  background: "#f8f9fa"
                }}>
                  {file.type === 'application/pdf' ? (
                    <div style={{
                      width: "100%",
                      height: "100%",
                      display: "flex",
                      flexDirection: "column",
                      alignItems: "center",
                      justifyContent: "center",
                      background: "#dc3545",
                      color: "white"
                    }}>
                      <div style={{ fontSize: "24px", marginBottom: "5px" }}>📄</div>
                      <div style={{ fontSize: "10px", textAlign: "center", padding: "0 5px" }}>PDF</div>
                    </div>
                  ) : (
                    <img
                      src={file.data}
                      alt={file.name}
                      style={{ width: "100%", height: "100%", objectFit: "cover" }}
                    />
                  )}
                  <button
                    type="button"
                    onClick={() => removeFile(index)}
                    style={{
                      position: "absolute",
                      top: "5px",
                      right: "5px",
                      background: "rgba(220, 53, 69, 0.9)",
                      color: "white",
                      border: "none",
                      borderRadius: "50%",
                      width: "24px",
                      height: "24px",
                      cursor: "pointer",
                      fontSize: "12px",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center"
                    }}
                  >
                    ×
                  </button>
                  <div style={{
                    position: "absolute",
                    bottom: "0",
                    left: "0",
                    right: "0",
                    background: "rgba(0, 0, 0, 0.7)",
                    color: "white",
                    padding: "4px 8px",
                    fontSize: "10px",
                    textAlign: "center"
                  }}>
                    {file.name}
                  </div>
                </div>
              ))}
            </div>
          </div>

          <button
            type="submit"
            style={{
              width: "100%",
              padding: "12px",
              background: loading ? "linear-gradient(135deg, #6c757d, #495057)" : "linear-gradient(135deg, #4a90e2, #357abd)",
              color: "white",
              border: "none",
              borderRadius: "8px",
              fontSize: "1rem",
              fontWeight: "600",
              cursor: loading ? "not-allowed" : "pointer",
              marginTop: "15px",
              transition: "all 0.3s ease"
            }}
            disabled={loading}
          >
            {loading ? (submitStatus || "Processing...") : "Submit"}
          </button>
        </form>

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
                {submitStatus === 'Saving...'}
                {submitStatus === 'Generating PDF...'}
                {submitStatus === 'Uploading PDF...'}
                {submitStatus === 'Finalizing...'}
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
                  ⚠️ Please do not close.
                </p>
              </div>
            </div>
          </div>
        )}

      </div>

      <style jsx>{`
        @keyframes fadeInUp {
          from {
            opacity: 0;
            transform: translateY(30px);
          }
          to {
            opacity: 1;
            transform: translateY(0);
          }
        }
        
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
        
        @media (max-width: 600px) {
          .form-row {
            flex-direction: column;
            gap: 0;
          }
        }
      `}</style>
    </div>
  );
}