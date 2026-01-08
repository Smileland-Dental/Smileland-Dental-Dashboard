'use client'

import React, { useState, useEffect, useRef } from "react";
import { doc, setDoc, deleteDoc } from "firebase/firestore";
import { db } from "@/lib/firebase.config";
import { getStorage, ref as storageRef, uploadBytes, getDownloadURL } from "firebase/storage";
import { validateSignatureDataUrlClient, sanitizeSignatureDataUrlClient } from "@/lib/security-client";

export default function OrthodonticContract() {
  const [loading, setLoading] = useState(false);
  const [submitStatus, setSubmitStatus] = useState('');
  const [progress, setProgress] = useState(0);
  
  // 현재 날짜
  const [contractDate] = useState(() => {
    const now = new Date();
    const californiaTime = new Date(now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
    return californiaTime.toISOString().split('T')[0];
  });

  // 폼 데이터
  const [formData, setFormData] = useState({
    patientName: '',
    dob: '',
    responsibleParty: '',
    relationship: '',
    ssn: '',
    driversLicense: '',
    typeOfTreatment: '',
    servicesRequired: [] as string[],
    additionalAppliances: [] as string[],
    firstOption: {
      treatment: '',
      appliance: '',
      deposit: '800',
      subtotal: '',
      estimatedInsurance: '',
      netBalance: '',
      estimatedTreatmentPeriod: '',
      monthlyPayment: ''
    },
    secondOption: {
      treatment: '',
      appliance: '',
      deposit: '',
      subtotal: '',
      estimatedInsurance: '',
      netBalance: '',
      estimatedTreatmentPeriod: '',
      monthlyPayment: ''
    },
    quotePresentedBy: '',
    quotePresentedDate: '',
    signatureDate: '',
    initial1: '',
    initial2: '',
    initial3: '',
    initial4: '',
    initial5: '',
    initial6: '',
    initial7: '',
    unpaidBalance: '',
    paymentMonths: '',
    monthlyPaymentAmount: '',
    paymentBeginDate: '',
    responsiblePartyName: '',
    responsiblePartySignatureDate: ''
  });

  // 서명 캔버스 관련 (Quote Presented 서명)
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [isDrawing, setIsDrawing] = useState(false);
  const [signatureData, setSignatureData] = useState<string>('');
  
  // 약관 동의 서명 캔버스 관련
  const termsCanvasRef = useRef<HTMLCanvasElement>(null);
  const [isDrawingTerms, setIsDrawingTerms] = useState(false);
  const [termsSignatureData, setTermsSignatureData] = useState<string>('');

  // Treatment 옵션 배열
  const treatmentOptions = [
    { label: 'Cash', price: '5400' },
    { label: 'Dltappo', price: '4051' },
    { label: 'Dltapre', price: '5116' },
    { label: 'Metlife', price: '4700' },
    { label: 'Cigna', price: '4499' },
    { label: 'Aetna', price: '4121' },
    { label: 'Guard', price: '4080' },
    { label: 'FDH', price: '5183' },
    { label: 'Dentmax', price: '5213.89' },
    { label: 'Amerits', price: '4792' },
    { label: 'PremAcc', price: '2635' },
    { label: 'UntdCon', price: '4284' },
    { label: 'Connect', price: '3800' },
    { label: 'AnthmBC', price: '4000' },
    { label: 'BS of CA', price: '3971' },
    { label: 'UntdHlt', price: '5028' },
    { label: 'DHA', price: '4121' }
  ];

  // 데이터 업데이트 함수
  const updateFormData = (field: string, value: string) => {
    setFormData(prev => ({ ...prev, [field]: value }));
  };

  // 체크박스 배열 업데이트 함수
  const toggleArrayItem = (field: 'servicesRequired' | 'additionalAppliances', value: string) => {
    setFormData(prev => {
      const currentArray = prev[field];
      const newArray = currentArray.includes(value)
        ? currentArray.filter(item => item !== value)
        : [...currentArray, value];
      return { ...prev, [field]: newArray };
    });
  };

  // 서명 그리기 시작
  const startDrawing = (e: React.MouseEvent<HTMLCanvasElement>) => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    
    const rect = canvas.getBoundingClientRect();
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    // 스케일 팩터 계산
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;

    setIsDrawing(true);
    ctx.beginPath();
    ctx.moveTo((e.clientX - rect.left) * scaleX, (e.clientY - rect.top) * scaleY);
  };

  // 서명 그리기
  const draw = (e: React.MouseEvent<HTMLCanvasElement>) => {
    if (!isDrawing) return;
    
    const canvas = canvasRef.current;
    if (!canvas) return;
    
    const rect = canvas.getBoundingClientRect();
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    // 스케일 팩터 계산
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;

    ctx.lineTo((e.clientX - rect.left) * scaleX, (e.clientY - rect.top) * scaleY);
    ctx.strokeStyle = '#000';
    ctx.lineWidth = 2;
    ctx.lineCap = 'round';
    ctx.stroke();
  };

  // 서명 그리기 종료
  const stopDrawing = () => {
    if (isDrawing) {
      const canvas = canvasRef.current;
      if (canvas) {
        const dataUrl = canvas.toDataURL();
        // 보안 검증: 서명 데이터 URL 검증
        const validatedDataUrl = sanitizeSignatureDataUrlClient(dataUrl);
        if (validatedDataUrl) {
          setSignatureData(validatedDataUrl);
        } else {
          alert('Invalid signature data. Please try signing again.');
          clearSignature();
        }
      }
    }
    setIsDrawing(false);
  };

  // 서명 지우기
  const clearSignature = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    ctx.clearRect(0, 0, canvas.width, canvas.height);
    setSignatureData('');
  };

  // 터치 이벤트: 서명 그리기 시작
  const startDrawingTouch = (e: React.TouchEvent<HTMLCanvasElement>) => {
    e.preventDefault(); // 스크롤 방지
    const canvas = canvasRef.current;
    if (!canvas) return;
    
    const rect = canvas.getBoundingClientRect();
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const touch = e.touches[0];
    if (!touch) return;

    // 스케일 팩터 계산
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;

    setIsDrawing(true);
    ctx.beginPath();
    ctx.moveTo((touch.clientX - rect.left) * scaleX, (touch.clientY - rect.top) * scaleY);
  };

  // 터치 이벤트: 서명 그리기
  const drawTouch = (e: React.TouchEvent<HTMLCanvasElement>) => {
    e.preventDefault(); // 스크롤 방지
    if (!isDrawing) return;
    
    const canvas = canvasRef.current;
    if (!canvas) return;
    
    const rect = canvas.getBoundingClientRect();
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const touch = e.touches[0];
    if (!touch) return;

    // 스케일 팩터 계산
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;

    ctx.lineTo((touch.clientX - rect.left) * scaleX, (touch.clientY - rect.top) * scaleY);
    ctx.strokeStyle = '#000';
    ctx.lineWidth = 2;
    ctx.lineCap = 'round';
    ctx.stroke();
  };

  // 터치 이벤트: 서명 그리기 종료
  const stopDrawingTouch = (e: React.TouchEvent<HTMLCanvasElement>) => {
    e.preventDefault(); // 스크롤 방지
    stopDrawing();
  };

  // 터치 이벤트: 서명 그리기 취소 (터치가 중단될 때)
  const cancelDrawingTouch = (e: React.TouchEvent<HTMLCanvasElement>) => {
    e.preventDefault();
    setIsDrawing(false);
  };

  // 약관 서명 그리기 시작
  const startDrawingTerms = (e: React.MouseEvent<HTMLCanvasElement>) => {
    const canvas = termsCanvasRef.current;
    if (!canvas) return;
    
    const rect = canvas.getBoundingClientRect();
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    // 스케일 팩터 계산
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;

    setIsDrawingTerms(true);
    ctx.beginPath();
    ctx.moveTo((e.clientX - rect.left) * scaleX, (e.clientY - rect.top) * scaleY);
  };

  // 약관 서명 그리기
  const drawTerms = (e: React.MouseEvent<HTMLCanvasElement>) => {
    if (!isDrawingTerms) return;
    
    const canvas = termsCanvasRef.current;
    if (!canvas) return;
    
    const rect = canvas.getBoundingClientRect();
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    // 스케일 팩터 계산
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;

    ctx.lineTo((e.clientX - rect.left) * scaleX, (e.clientY - rect.top) * scaleY);
    ctx.strokeStyle = '#000';
    ctx.lineWidth = 2;
    ctx.lineCap = 'round';
    ctx.stroke();
  };

  // 약관 서명 그리기 종료
  const stopDrawingTerms = () => {
    if (isDrawingTerms) {
      const canvas = termsCanvasRef.current;
      if (canvas) {
        const dataUrl = canvas.toDataURL();
        // 보안 검증: 서명 데이터 URL 검증
        const validatedDataUrl = sanitizeSignatureDataUrlClient(dataUrl);
        if (validatedDataUrl) {
          setTermsSignatureData(validatedDataUrl);
        } else {
          alert('Invalid signature data. Please try signing again.');
          clearTermsSignature();
        }
      }
    }
    setIsDrawingTerms(false);
  };

  // 약관 서명 지우기
  const clearTermsSignature = () => {
    const canvas = termsCanvasRef.current;
    if (!canvas) return;
    
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    ctx.clearRect(0, 0, canvas.width, canvas.height);
    setTermsSignatureData('');
  };

  // 터치 이벤트: 약관 서명 그리기 시작
  const startDrawingTermsTouch = (e: React.TouchEvent<HTMLCanvasElement>) => {
    e.preventDefault(); // 스크롤 방지
    const canvas = termsCanvasRef.current;
    if (!canvas) return;
    
    const rect = canvas.getBoundingClientRect();
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const touch = e.touches[0];
    if (!touch) return;

    // 스케일 팩터 계산
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;

    setIsDrawingTerms(true);
    ctx.beginPath();
    ctx.moveTo((touch.clientX - rect.left) * scaleX, (touch.clientY - rect.top) * scaleY);
  };

  // 터치 이벤트: 약관 서명 그리기
  const drawTermsTouch = (e: React.TouchEvent<HTMLCanvasElement>) => {
    e.preventDefault(); // 스크롤 방지
    if (!isDrawingTerms) return;
    
    const canvas = termsCanvasRef.current;
    if (!canvas) return;
    
    const rect = canvas.getBoundingClientRect();
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const touch = e.touches[0];
    if (!touch) return;

    // 스케일 팩터 계산
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;

    ctx.lineTo((touch.clientX - rect.left) * scaleX, (touch.clientY - rect.top) * scaleY);
    ctx.strokeStyle = '#000';
    ctx.lineWidth = 2;
    ctx.lineCap = 'round';
    ctx.stroke();
  };

  // 터치 이벤트: 약관 서명 그리기 종료
  const stopDrawingTermsTouch = (e: React.TouchEvent<HTMLCanvasElement>) => {
    e.preventDefault(); // 스크롤 방지
    stopDrawingTerms();
  };

  // 터치 이벤트: 약관 서명 그리기 취소 (터치가 중단될 때)
  const cancelDrawingTermsTouch = (e: React.TouchEvent<HTMLCanvasElement>) => {
    e.preventDefault();
    setIsDrawingTerms(false);
  };

  // 제출 처리
  const handleSubmit = async () => {
    // 필수 필드 확인
    if (!formData.patientName || !formData.dob) {
      alert("Patient's Name and DOB are required fields.");
      return;
    }

    // 보안 검증: 서명 데이터 검증
    let validatedSignatureData = '';
    let validatedTermsSignatureData = '';
    
    if (signatureData) {
      validatedSignatureData = sanitizeSignatureDataUrlClient(signatureData);
      if (!validatedSignatureData) {
        alert('Invalid signature data. Please sign again.');
        return;
      }
    }
    
    if (termsSignatureData) {
      validatedTermsSignatureData = sanitizeSignatureDataUrlClient(termsSignatureData);
      if (!validatedTermsSignatureData) {
        alert('Invalid terms signature data. Please sign again.');
        return;
      }
    }

    const confirmSubmit = window.confirm("Are you sure you want to submit this contract? This will save the data and generate a PDF.");
    if (!confirmSubmit) return;

    try {
      setLoading(true);
      setSubmitStatus('Generating PDF...');
      setProgress(20);
      
      // 1. PDF 생성
      const response = await fetch('/api/generate-orthodontic-contract-pdf', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          contractDate,
          formData: formData,
          signatureData: validatedSignatureData,
          termsSignatureData: validatedTermsSignatureData
        }),
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'PDF generation failed');
      }

      const blob = await response.blob();
      
      setSubmitStatus('Saving to Firebase...');
      setProgress(40);

      // 2. Firestore에 데이터 저장 (서명 데이터는 제외 - PDF에만 포함됨)
      const timestamp = Date.now();
      const docId = `contract_${contractDate}_${timestamp}`;
      
      const dataToSave = {
        ...formData,
        contractDate,
        timestamp: new Date().toISOString(),
        submitted: true,
        approved: true
        // signatureData와 termsSignatureData는 PDF 생성에만 사용하고 Firestore에는 저장하지 않음
      };

      await setDoc(doc(db, "orthodontic-contracts", docId), dataToSave);
      
      setSubmitStatus('Uploading PDF to storage...');
      setProgress(60);
      
      // 3. Firebase Storage에 PDF 업로드
      const storage = getStorage();
      const fileName = `Orthodontic_Contract_${formData.patientName.replace(/\s+/g, '_')}_${contractDate}.pdf`;
      const pdfRef = storageRef(storage, `orthodontic-contracts/${docId}/${fileName}`);
      
      await uploadBytes(pdfRef, blob);
      const downloadURL = await getDownloadURL(pdfRef);
      
      // Firestore에 PDF URL 추가
      await setDoc(doc(db, "orthodontic-contracts", docId), {
        ...dataToSave,
        pdfUrl: downloadURL,
        pdfFileName: fileName
      });
      
      setSubmitStatus('Complete!');
      setProgress(100);
      
      // 폼 초기화
      setFormData({
        patientName: '',
        dob: '',
        responsibleParty: '',
        relationship: '',
        ssn: '',
        driversLicense: '',
        typeOfTreatment: '',
        servicesRequired: [],
        additionalAppliances: [],
        firstOption: {
          treatment: '',
          appliance: '',
          deposit: '800',
          subtotal: '',
          estimatedInsurance: '',
          netBalance: '',
          estimatedTreatmentPeriod: '',
          monthlyPayment: ''
        },
        secondOption: {
          treatment: '',
          appliance: '',
          deposit: '',
          subtotal: '',
          estimatedInsurance: '',
          netBalance: '',
          estimatedTreatmentPeriod: '',
          monthlyPayment: ''
        },
        quotePresentedBy: '',
        quotePresentedDate: '',
        signatureDate: '',
        initial1: '',
        initial2: '',
        initial3: '',
        initial4: '',
        initial5: '',
        initial6: '',
        initial7: '',
        unpaidBalance: '',
        paymentMonths: '',
        monthlyPaymentAmount: '',
        paymentBeginDate: '',
        responsiblePartyName: '',
        responsiblePartySignatureDate: ''
      });
      clearSignature();
      clearTermsSignature();
      
      setTimeout(() => {
        setLoading(false);
        setSubmitStatus('');
        setProgress(0);
        alert('✅ Contract has been saved successfully and sent to admin!');
      }, 1000);
      
    } catch (error) {
      console.error('Submit error:', error);
      setSubmitStatus('❌ Submission failed: ' + (error as Error).message);
      setProgress(0);
      setTimeout(() => {
        setLoading(false);
        setSubmitStatus('');
      }, 3000);
    }
  };

  // 스타일 정의
  const styles = {
    body: {
      fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
      padding: '15px',
      background: 'linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%)',
      color: '#2c3e50',
      lineHeight: '1.6',
      minHeight: '100vh'
    },
    container: {
      maxWidth: '95%',
      width: '100%',
      margin: '20px auto',
      padding: '30px',
      backgroundColor: 'white',
      borderRadius: '12px',
      boxShadow: '0 4px 12px rgba(0, 0, 0, 0.15)',
      position: 'relative' as const,
      boxSizing: 'border-box' as const
    },
    header: {
      color: '#1976d2',
      textAlign: 'center' as const,
      marginBottom: '30px',
      paddingBottom: '15px',
      borderBottom: '3px solid #1976d2',
      fontSize: '2.2em',
      fontWeight: 'bold'
    },
    formGroup: {
      marginBottom: '25px'
    },
    formRow: {
      display: 'grid',
      gridTemplateColumns: '1fr 1fr',
      gap: '20px',
      marginBottom: '25px'
    },
    formRowThree: {
      display: 'grid',
      gridTemplateColumns: '1fr 1fr 1fr',
      gap: '15px',
      marginBottom: '25px'
    },
    label: {
      display: 'block',
      fontWeight: '600',
      color: '#2c3e50',
      marginBottom: '8px',
      fontSize: '15px'
    },
    input: {
      width: '100%',
      padding: '12px 15px',
      fontSize: '15px',
      border: '2px solid #e0e0e0',
      borderRadius: '6px',
      backgroundColor: 'white',
      boxSizing: 'border-box' as const,
      transition: 'border-color 0.2s',
      outline: 'none'
    },
    radioGroup: {
      display: 'flex',
      gap: '20px',
      marginTop: '10px'
    },
    radioOption: {
      display: 'flex',
      alignItems: 'center',
      gap: '8px',
      fontSize: '15px'
    },
    checkboxGroup: {
      display: 'grid',
      gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))',
      gap: '15px',
      marginTop: '15px'
    },
    checkboxOption: {
      display: 'flex',
      alignItems: 'center',
      gap: '10px',
      fontSize: '15px',
      padding: '12px 15px',
      backgroundColor: '#f8f9fa',
      borderRadius: '6px',
      border: '2px solid #e0e0e0',
      cursor: 'pointer',
      transition: 'all 0.2s'
    },
    checkboxOptionChecked: {
      backgroundColor: '#e3f2fd',
      border: '2px solid #1976d2',
      fontWeight: '600'
    },
    servicePrice: {
      marginLeft: 'auto',
      color: '#1976d2',
      fontWeight: '600'
    },
    sectionTitle: {
      fontSize: '18px',
      fontWeight: '600',
      color: '#1976d2',
      marginTop: '30px',
      marginBottom: '15px',
      paddingBottom: '10px',
      borderBottom: '2px solid #e0e0e0'
    },
    submitButton: {
      display: 'block',
      width: '200px',
      margin: '40px auto 0 auto',
      padding: '15px 25px',
      backgroundColor: '#1976d2',
      color: 'white',
      border: 'none',
      borderRadius: '8px',
      cursor: 'pointer',
      fontSize: '17px',
      fontWeight: '600',
      transition: 'background-color 0.2s',
      boxShadow: '0 2px 8px rgba(25, 118, 210, 0.3)'
    },
    statusMessage: {
      marginTop: '20px',
      fontWeight: 'bold',
      textAlign: 'center' as const,
      padding: '12px',
      borderRadius: '6px'
    },
    footer: {
      marginTop: '40px',
      paddingTop: '20px',
      borderTop: '2px solid #e0e0e0',
      textAlign: 'center' as const,
      fontSize: '14px',
      color: '#6c757d'
    }
  };

  // Additional Appliance 선택 시 가격 합산하여 First Option의 Appliance 필드에 자동 입력
  useEffect(() => {
    const additionalApplianceOptions = [
      { label: 'Nance', price: 250 },
      { label: 'Quad helix', price: 250 },
      { label: 'LLA', price: 100 }
    ];

    const totalPrice = formData.additionalAppliances.reduce((sum, selectedValue) => {
      // selectedValue 형식: "Label|$Price"
      const priceMatch = selectedValue.match(/\$(\d+)/);
      if (priceMatch) {
        return sum + parseInt(priceMatch[1], 10);
      }
      return sum;
    }, 0);

    setFormData(prev => ({
      ...prev,
      firstOption: {
        ...(prev.firstOption || {}),
        appliance: totalPrice > 0 ? totalPrice.toString() : '',
        deposit: '800' // Deposit은 항상 800으로 유지
      }
    }));
  }, [formData.additionalAppliances]);

  // Deposit이 항상 800으로 유지되도록 보장
  useEffect(() => {
    if (formData.firstOption?.deposit !== '800') {
      setFormData(prev => ({
        ...prev,
        firstOption: {
          ...(prev.firstOption || {}),
          deposit: '800'
        }
      }));
    }
  }, [formData.firstOption?.deposit]);

  // Subtotal 자동 계산: Treatment + Appliance - Deposit
  useEffect(() => {
    const treatment = parseFloat(formData.firstOption?.treatment || '0') || 0;
    const appliance = parseFloat(formData.firstOption?.appliance || '0') || 0;
    const deposit = parseFloat(formData.firstOption?.deposit || '800') || 800;
    
    const subtotal = treatment + appliance - deposit;
    
    setFormData(prev => ({
      ...prev,
      firstOption: {
        ...(prev.firstOption || {}),
        subtotal: subtotal > 0 ? subtotal.toString() : '0'
      }
    }));
  }, [formData.firstOption?.treatment, formData.firstOption?.appliance, formData.firstOption?.deposit]);

  // Net Balance 자동 계산: Treatment + Appliance - Deposit - Estimated Insurance
  useEffect(() => {
    const treatment = parseFloat(formData.firstOption?.treatment || '0') || 0;
    const appliance = parseFloat(formData.firstOption?.appliance || '0') || 0;
    const deposit = parseFloat(formData.firstOption?.deposit || '800') || 800;
    const estimatedInsurance = parseFloat(formData.firstOption?.estimatedInsurance || '0') || 0;
    
    const netBalance = treatment + appliance - deposit - estimatedInsurance;
    
    setFormData(prev => ({
      ...prev,
      firstOption: {
        ...(prev.firstOption || {}),
        netBalance: netBalance > 0 ? netBalance.toString() : '0'
      }
    }));
  }, [formData.firstOption?.treatment, formData.firstOption?.appliance, formData.firstOption?.deposit, formData.firstOption?.estimatedInsurance]);

  // Monthly Payment 자동 계산: Net Balance / Estimated Treatment Period (24 months 초과 시 24로 나눔, 소수점 올림)
  useEffect(() => {
    const netBalance = parseFloat(formData.firstOption?.netBalance || '0') || 0;
    const estimatedTreatmentPeriod = parseFloat(formData.firstOption?.estimatedTreatmentPeriod || '0') || 0;
    
    if (netBalance > 0 && estimatedTreatmentPeriod > 0) {
      const divisor = estimatedTreatmentPeriod > 24 ? 24 : estimatedTreatmentPeriod;
      const monthlyPayment = Math.ceil(netBalance / divisor);
      
      setFormData(prev => ({
        ...prev,
        firstOption: {
          ...(prev.firstOption || {}),
          monthlyPayment: monthlyPayment.toString()
        }
      }));
    } else {
      setFormData(prev => ({
        ...prev,
        firstOption: {
          ...(prev.firstOption || {}),
          monthlyPayment: '0'
        }
      }));
    }
  }, [formData.firstOption?.netBalance, formData.firstOption?.estimatedTreatmentPeriod]);

  // Second Option - Subtotal 자동 계산: Treatment + Appliance - Deposit
  useEffect(() => {
    const treatment = parseFloat(formData.secondOption?.treatment || '0') || 0;
    const appliance = parseFloat(formData.secondOption?.appliance || '0') || 0;
    const deposit = parseFloat(formData.secondOption?.deposit || '0') || 0;
    
    const subtotal = treatment + appliance - deposit;
    
    setFormData(prev => ({
      ...prev,
      secondOption: {
        ...(prev.secondOption || {}),
        subtotal: subtotal > 0 ? subtotal.toString() : '0'
      }
    }));
  }, [formData.secondOption?.treatment, formData.secondOption?.appliance, formData.secondOption?.deposit]);

  // Second Option - Net Balance 자동 계산: Treatment + Appliance - Deposit - Estimated Insurance
  useEffect(() => {
    const treatment = parseFloat(formData.secondOption?.treatment || '0') || 0;
    const appliance = parseFloat(formData.secondOption?.appliance || '0') || 0;
    const deposit = parseFloat(formData.secondOption?.deposit || '0') || 0;
    const estimatedInsurance = parseFloat(formData.secondOption?.estimatedInsurance || '0') || 0;
    
    const netBalance = treatment + appliance - deposit - estimatedInsurance;
    
    setFormData(prev => ({
      ...prev,
      secondOption: {
        ...(prev.secondOption || {}),
        netBalance: netBalance > 0 ? netBalance.toString() : '0'
      }
    }));
  }, [formData.secondOption?.treatment, formData.secondOption?.appliance, formData.secondOption?.deposit, formData.secondOption?.estimatedInsurance]);

  // Second Option - Monthly Payment 자동 계산: Net Balance / Estimated Treatment Period (24 months 초과 시 24로 나눔, 소수점 올림)
  useEffect(() => {
    const netBalance = parseFloat(formData.secondOption?.netBalance || '0') || 0;
    const estimatedTreatmentPeriod = parseFloat(formData.secondOption?.estimatedTreatmentPeriod || '0') || 0;
    
    if (netBalance > 0 && estimatedTreatmentPeriod > 0) {
      const divisor = estimatedTreatmentPeriod > 24 ? 24 : estimatedTreatmentPeriod;
      const monthlyPayment = Math.ceil(netBalance / divisor);
      
      setFormData(prev => ({
        ...prev,
        secondOption: {
          ...(prev.secondOption || {}),
          monthlyPayment: monthlyPayment.toString()
        }
      }));
    } else {
      setFormData(prev => ({
        ...prev,
        secondOption: {
          ...(prev.secondOption || {}),
          monthlyPayment: '0'
        }
      }));
    }
  }, [formData.secondOption?.netBalance, formData.secondOption?.estimatedTreatmentPeriod]);

  // 제출 중 브라우저 네비게이션 방지
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
        * {
          box-sizing: border-box;
        }
        
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
        
        input:focus, select:focus {
          border-color: #1976d2 !important;
        }
        
        button:hover:not(:disabled) {
          background-color: #1565c0 !important;
          transform: translateY(-2px);
          box-shadow: 0 4px 12px rgba(25, 118, 210, 0.4) !important;
        }
        
        label.checkbox-label:hover {
          transform: translateY(-2px);
          box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        }
        
        /* 반응형 미디어 쿼리 */
        @media (max-width: 1024px) {
          .payment-options {
            grid-template-columns: 1fr !important;
          }
        }
        
        @media (max-width: 768px) {
          .form-row-responsive {
            grid-template-columns: 1fr !important;
          }
          
          .services-grid {
            grid-template-columns: 1fr !important;
          }
          
          canvas {
            width: 100% !important;
            height: 120px !important;
          }
        }
        
        @media (max-width: 480px) {
          h1 {
            font-size: 1.5em !important;
          }
        }
      `}</style>
      <div style={styles.body}>
      {/* 로딩 모달 */}
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
              borderTop: "4px solid #1976d2",
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
            {/* 진행률 바 */}
            <div style={{
              width: "100%",
              backgroundColor: "#e9ecef",
              borderRadius: "10px",
              overflow: "hidden",
              marginTop: "20px"
            }}>
              <div style={{
                width: `${progress}%`,
                height: "8px",
                backgroundColor: "#1976d2",
                borderRadius: "10px",
                transition: "width 0.3s ease"
              }}></div>
            </div>
            <p style={{
              color: "#495057",
              fontSize: "0.9rem",
              margin: "15px 0",
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

        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '20px', marginBottom: '30px' }}>
          <h1 style={{ ...styles.header, marginBottom: 0, paddingBottom: 0, borderBottom: 'none', flex: 1, textAlign: 'center' }}>Orthodontic Contract</h1>
          <a 
            href="/dashboard/forms/orthodonticcontract-spanish"
            style={{
              padding: '10px 20px',
              backgroundColor: '#28a745',
              color: 'white',
              textDecoration: 'none',
              borderRadius: '6px',
              fontSize: '14px',
              fontWeight: '600',
              transition: 'background-color 0.2s',
              cursor: 'pointer',
              border: 'none',
              boxShadow: '0 2px 8px rgba(40, 167, 69, 0.3)'
            }}
            onMouseEnter={(e) => e.currentTarget.style.backgroundColor = '#218838'}
            onMouseLeave={(e) => e.currentTarget.style.backgroundColor = '#28a745'}
          >
            Español
          </a>
        </div>

        {/* Patient's Name and DOB */}
        <div className="form-row-responsive" style={styles.formRow}>
          <div>
            <label style={styles.label} htmlFor="patientName">Patient's Name *</label>
            <input
              type="text"
              id="patientName"
              value={formData.patientName}
              onChange={(e) => updateFormData('patientName', e.target.value)}
              style={styles.input}
              placeholder="Enter patient's full name"
              required
            />
          </div>
          <div>
            <label style={styles.label} htmlFor="dob">Date of Birth (DOB) *</label>
            <input
              type="text"
              id="dob"
              value={formData.dob}
              onChange={(e) => updateFormData('dob', e.target.value)}
              style={styles.input}
              placeholder="MM/DD/YYYY"
              required
            />
          </div>
        </div>

        {/* Responsible Party and Relationship */}
        <div className="form-row-responsive" style={styles.formRow}>
          <div>
            <label style={styles.label} htmlFor="responsibleParty">Responsible Party</label>
            <input
              type="text"
              id="responsibleParty"
              value={formData.responsibleParty}
              onChange={(e) => updateFormData('responsibleParty', e.target.value)}
              style={styles.input}
              placeholder="Enter responsible party name"
            />
          </div>
          <div>
            <label style={styles.label} htmlFor="relationship">Relationship</label>
            <input
              type="text"
              id="relationship"
              value={formData.relationship}
              onChange={(e) => updateFormData('relationship', e.target.value)}
              style={styles.input}
            />
          </div>
        </div>

        {/* S.S.# and Driver's License */}
        <div className="form-row-responsive" style={styles.formRow}>
          <div>
            <label style={styles.label} htmlFor="ssn">S.S.#</label>
            <input
              type="text"
              id="ssn"
              value={formData.ssn}
              onChange={(e) => updateFormData('ssn', e.target.value)}
              style={styles.input}
              placeholder="XXX-XX-XXXX"
            />
          </div>
          <div>
            <label style={styles.label} htmlFor="driversLicense">Driver's License</label>
            <input
              type="text"
              id="driversLicense"
              value={formData.driversLicense}
              onChange={(e) => updateFormData('driversLicense', e.target.value)}
              style={styles.input}
              placeholder="Enter driver's license number"
            />
          </div>
        </div>

        {/* Type of Treatment */}
        <div style={styles.formGroup}>
          <label style={styles.label}>Type of Treatment</label>
          <div style={styles.radioGroup}>
            <label style={styles.radioOption}>
              <input
                type="radio"
                name="typeOfTreatment"
                value="Comprehensive"
                checked={formData.typeOfTreatment === 'Comprehensive'}
                onChange={(e) => updateFormData('typeOfTreatment', e.target.value)}
                style={{ cursor: 'pointer' }}
              />
              <span>Comprehensive</span>
            </label>
            <label style={styles.radioOption}>
              <input
                type="radio"
                name="typeOfTreatment"
                value="Limited"
                checked={formData.typeOfTreatment === 'Limited'}
                onChange={(e) => updateFormData('typeOfTreatment', e.target.value)}
                style={{ cursor: 'pointer' }}
              />
              <span>Limited</span>
            </label>
            <label style={styles.radioOption}>
              <input
                type="radio"
                name="typeOfTreatment"
                value="Phase I"
                checked={formData.typeOfTreatment === 'Phase I'}
                onChange={(e) => updateFormData('typeOfTreatment', e.target.value)}
                style={{ cursor: 'pointer' }}
              />
              <span>Phase I</span>
            </label>
            <label style={styles.radioOption}>
              <input
                type="radio"
                name="typeOfTreatment"
                value="Phase II"
                checked={formData.typeOfTreatment === 'Phase II'}
                onChange={(e) => updateFormData('typeOfTreatment', e.target.value)}
                style={{ cursor: 'pointer' }}
              />
              <span>Phase II</span>
            </label>
          </div>
        </div>

        {/* Services Required */}
        <div style={{ marginTop: '40px' }}>
          <div style={styles.sectionTitle}>Services Required</div>
          <div className="services-grid" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '15px', marginTop: '15px' }}>
            {[
              { label: 'Panoramic film' },
              { label: 'Cephalometric film' },
              { label: 'Diagnostic casts' },
              { label: 'Oral/Facial image' },
              { label: 'Orthodontic retention(U/L)' }
            ].map((service) => {
              const value = service.label;
              const isChecked = formData.servicesRequired.includes(value);
              return (
                <label
                  key={service.label}
                  className="checkbox-label"
                  style={{
                    ...styles.checkboxOption,
                    ...(isChecked ? styles.checkboxOptionChecked : {})
                  }}
                >
                  <input
                    type="checkbox"
                    checked={isChecked}
                    onChange={() => toggleArrayItem('servicesRequired', value)}
                    style={{ cursor: 'pointer', width: '18px', height: '18px' }}
                  />
                  <span style={{ flex: 1 }}>{service.label}</span>
                </label>
              );
            })}
          </div>
        </div>

        {/* Additional appliance (if necessary) */}
        <div style={{ marginTop: '40px' }}>
          <div style={styles.sectionTitle}>Additional Appliance (if necessary)</div>
          <div className="services-grid" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '15px', marginTop: '15px' }}>
            {[
              { label: 'Nance', price: 250 },
              { label: 'Quad helix', price: 250 },
              { label: 'LLA', price: 100 }
            ].map((appliance) => {
              const value = `${appliance.label}|$${appliance.price}`;
              const isChecked = formData.additionalAppliances.includes(value);
              return (
                <label
                  key={appliance.label}
                  className="checkbox-label"
                  style={{
                    ...styles.checkboxOption,
                    ...(isChecked ? styles.checkboxOptionChecked : {})
                  }}
                >
                  <input
                    type="checkbox"
                    checked={isChecked}
                    onChange={() => toggleArrayItem('additionalAppliances', value)}
                    style={{ cursor: 'pointer', width: '18px', height: '18px' }}
                  />
                  <span style={{ flex: 1 }}>{appliance.label}</span>
                  <span style={styles.servicePrice}>${appliance.price.toLocaleString()}</span>
                </label>
              );
            })}
          </div>
        </div>

        {/* Payment Options */}
        <div className="payment-options" style={{ marginTop: '50px', display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '30px' }}>
          {/* First Option */}
          <div style={{ border: '2px solid #2c3e50', borderRadius: '10px', padding: '20px', backgroundColor: '#f8f9fa' }}>
            <h3 style={{ marginTop: 0, marginBottom: '20px', color: '#2c3e50', textAlign: 'center' }}>First Option</h3>
            
            <div style={styles.formGroup}>
              <label style={styles.label}>Treatment:</label>
              <select
                value={formData.firstOption?.treatment || ''}
                onChange={(e) => {
                  setFormData({
                    ...formData,
                    firstOption: { ...(formData.firstOption || {}), treatment: e.target.value }
                  });
                }}
                style={{
                  ...styles.input,
                  padding: '12px 15px',
                  cursor: 'pointer',
                  appearance: 'auto'
                }}
              >
                <option value="">Select treatment option</option>
                {treatmentOptions.map((option, index) => (
                  <option key={index} value={option.price}>
                    {option.label} - ${option.price}
                  </option>
                ))}
              </select>
              {formData.firstOption?.treatment && (
                <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px', marginTop: '10px' }}>
                  <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                  <input
                    type="text"
                    value={formData.firstOption.treatment}
                    readOnly
                    style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%', backgroundColor: 'transparent' }}
                  />
                </div>
              )}
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Appliance:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.firstOption?.appliance || ''}
                  readOnly
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%', backgroundColor: '#f5f5f5', cursor: 'not-allowed' }}
                />
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Deposit:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  value={formData.firstOption?.deposit || '800'}
                  readOnly
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%', backgroundColor: '#f5f5f5', cursor: 'not-allowed' }}
                />
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Subtotal:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.firstOption?.subtotal || ''}
                  readOnly
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%', backgroundColor: '#f5f5f5', cursor: 'not-allowed' }}
                />
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Estimated Insurance:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.firstOption?.estimatedInsurance}
                  onChange={(e) => setFormData({
                    ...formData,
                    firstOption: { ...(formData.firstOption || {}), estimatedInsurance: e.target.value }
                  })}
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%' }}
                />
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Net Balance:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.firstOption?.netBalance || ''}
                  readOnly
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%', backgroundColor: '#f5f5f5', cursor: 'not-allowed' }}
                />
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Estimated Treatment Period:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <input
                  type="text"
                  placeholder="0"
                  value={formData.firstOption?.estimatedTreatmentPeriod || ''}
                  onChange={(e) => setFormData({
                    ...formData,
                    firstOption: { ...(formData.firstOption || {}), estimatedTreatmentPeriod: e.target.value }
                  })}
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%' }}
                />
                <span style={{ color: '#666', marginLeft: '5px', whiteSpace: 'nowrap' }}>months</span>
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Monthly Payment:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.firstOption?.monthlyPayment || ''}
                  readOnly
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%', backgroundColor: '#f5f5f5', cursor: 'not-allowed' }}
                />
              </div>
            </div>
          </div>

          {/* Second Option */}
          <div style={{ border: '2px solid #2c3e50', borderRadius: '10px', padding: '20px', backgroundColor: '#f8f9fa' }}>
            <h3 style={{ marginTop: 0, marginBottom: '20px', color: '#2c3e50', textAlign: 'center' }}>Second Option</h3>
            
            <div style={styles.formGroup}>
              <label style={styles.label}>Treatment:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.secondOption?.treatment}
                  onChange={(e) => setFormData({
                    ...formData,
                    secondOption: { ...(formData.secondOption || {}), treatment: e.target.value }
                  })}
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%' }}
                />
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Appliance:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.secondOption?.appliance}
                  onChange={(e) => setFormData({
                    ...formData,
                    secondOption: { ...(formData.secondOption || {}), appliance: e.target.value }
                  })}
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%' }}
                />
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Deposit:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.secondOption?.deposit}
                  onChange={(e) => setFormData({
                    ...formData,
                    secondOption: { ...(formData.secondOption || {}), deposit: e.target.value }
                  })}
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%' }}
                />
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Subtotal:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.secondOption?.subtotal || ''}
                  readOnly
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%', backgroundColor: '#f5f5f5', cursor: 'not-allowed' }}
                />
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Estimated Insurance:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.secondOption?.estimatedInsurance}
                  onChange={(e) => setFormData({
                    ...formData,
                    secondOption: { ...(formData.secondOption || {}), estimatedInsurance: e.target.value }
                  })}
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%' }}
                />
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Net Balance:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.secondOption?.netBalance || ''}
                  readOnly
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%', backgroundColor: '#f5f5f5', cursor: 'not-allowed' }}
                />
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Estimated Treatment Period:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <input
                  type="text"
                  placeholder="0"
                  value={formData.secondOption?.estimatedTreatmentPeriod || ''}
                  onChange={(e) => setFormData({
                    ...formData,
                    secondOption: { ...(formData.secondOption || {}), estimatedTreatmentPeriod: e.target.value }
                  })}
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%' }}
                />
                <span style={{ color: '#666', marginLeft: '5px', whiteSpace: 'nowrap' }}>months</span>
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Monthly Payment:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.secondOption?.monthlyPayment || ''}
                  readOnly
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%', backgroundColor: '#f5f5f5', cursor: 'not-allowed' }}
                />
              </div>
            </div>
          </div>
        </div>

        {/* Terms and Conditions */}
        <div style={{ marginTop: '30px', padding: '20px', backgroundColor: '#f8f9fa', borderRadius: '8px', border: '1px solid #dee2e6' }}>
          <div style={{ margin: 0, fontSize: '13px', lineHeight: '1.8', color: '#495057' }}>
            <div style={{ marginBottom: '8px' }}>* The charges for the procedures already performed on patient are not refundable.</div>
            <div style={{ marginBottom: '8px' }}>* We offer an 5% fee reduction for paying in full with cash at the beginning of orthodontic treatment and a 3% fee reduction if paid in full with credit card (Visa or MasterCard).</div>
            <div style={{ marginBottom: '8px' }}>* Clear brackets are available for an additional cost of $300 per arch.</div>
            <div style={{ marginBottom: '8px' }}>* Estimated insurance amount is not a guaranteed amount by no means and subject to change due to variable calculation formulas used by insurance carriers.</div>
            <div style={{ marginBottom: '8px' }}>* An administration fee of $100 will be charged should the account require a new contract.</div>
            <div style={{ marginBottom: '8px' }}>* This quote is valid for 30 days.</div>
          </div>
        </div>

        {/* Quote Presented By and Signature Section */}
        <div style={{ marginTop: '40px' }}>
          {/* Quote Presented By and Date */}
          <div className="form-row-responsive" style={styles.formRow}>
            <div>
              <label style={styles.label} htmlFor="quotePresentedBy">Quote Presented by:</label>
              <input
                type="text"
                id="quotePresentedBy"
                value={formData.quotePresentedBy}
                onChange={(e) => updateFormData('quotePresentedBy', e.target.value)}
                style={styles.input}
                placeholder="Enter full name"
              />
            </div>
            <div>
              <label style={styles.label} htmlFor="quotePresentedDate">Date:</label>
              <input
                type="text"
                id="quotePresentedDate"
                value={formData.quotePresentedDate}
                onChange={(e) => updateFormData('quotePresentedDate', e.target.value)}
                style={styles.input}
                placeholder="MM/DD/YYYY"
              />
            </div>
          </div>

          {/* Signature and Date */}
          <div className="form-row-responsive" style={{ ...styles.formRow, marginTop: '30px' }}>
            <div>
              <label style={styles.label}>Signature of Responsible Party:</label>
              <div style={{ position: 'relative' }}>
                <canvas
                  ref={canvasRef}
                  width={500}
                  height={150}
                  onMouseDown={startDrawing}
                  onMouseMove={draw}
                  onMouseUp={stopDrawing}
                  onMouseLeave={stopDrawing}
                  onTouchStart={startDrawingTouch}
                  onTouchMove={drawTouch}
                  onTouchEnd={stopDrawingTouch}
                  onTouchCancel={cancelDrawingTouch}
                  style={{
                    border: '2px solid #e0e0e0',
                    borderRadius: '6px',
                    cursor: 'crosshair',
                    backgroundColor: 'white',
                    width: '100%',
                    height: '150px',
                    touchAction: 'none' // 터치 스크롤 방지
                  }}
                />
                <button
                  type="button"
                  onClick={clearSignature}
                  style={{
                    position: 'absolute',
                    top: '5px',
                    right: '5px',
                    padding: '5px 10px',
                    fontSize: '12px',
                    backgroundColor: '#ff6b6b',
                    color: 'white',
                    border: 'none',
                    borderRadius: '4px',
                    cursor: 'pointer'
                  }}
                >
                  Clear
                </button>
              </div>
            </div>
            <div>
              <label style={styles.label} htmlFor="signatureDate">Date:</label>
              <input
                type="text"
                id="signatureDate"
                value={formData.signatureDate}
                onChange={(e) => updateFormData('signatureDate', e.target.value)}
                style={styles.input}
                placeholder="MM/DD/YYYY"
              />
            </div>
          </div>
        </div>

        {/* Terms and Conditions - Detailed */}
        <div style={{ marginTop: '50px', padding: '30px', backgroundColor: '#f8f9fa', borderRadius: '8px', border: '1px solid #dee2e6' }}>
          <h3 style={{ marginTop: 0, marginBottom: '20px', color: '#2c3e50', textAlign: 'center' }}>Terms and Conditions</h3>
          
          {/* Section 1: Payment Terms */}
          <div style={{ marginBottom: '25px', fontSize: '13px', lineHeight: '1.8', color: '#495057' }}>
            <p style={{ textAlign: 'justify', marginBottom: '15px', display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: '5px' }}>
              <span>The unpaid balance of $</span>
              <input
                type="text"
                value={formData.unpaidBalance}
                onChange={(e) => updateFormData('unpaidBalance', e.target.value)}
                style={{ width: '100px', padding: '4px 8px', border: '1px solid #333', borderRadius: '4px', fontSize: '13px', textAlign: 'center' }}
                placeholder=""
              />
              <span>will be paid in</span>
              <input
                type="text"
                value={formData.paymentMonths}
                onChange={(e) => updateFormData('paymentMonths', e.target.value)}
                style={{ width: '60px', padding: '4px 8px', border: '1px solid #333', borderRadius: '4px', fontSize: '13px', textAlign: 'center' }}
                placeholder=""
              />
              <span>months in the amount of $</span>
              <input
                type="text"
                value={formData.monthlyPaymentAmount}
                onChange={(e) => updateFormData('monthlyPaymentAmount', e.target.value)}
                style={{ width: '100px', padding: '4px 8px', border: '1px solid #333', borderRadius: '4px', fontSize: '13px', textAlign: 'center' }}
                placeholder=""
              />
              <span>, each due on the 1st or 15th day of the month, beginning on</span>
              <input
                type="text"
                value={formData.paymentBeginDate}
                onChange={(e) => updateFormData('paymentBeginDate', e.target.value)}
                style={{ width: '120px', padding: '4px 8px', border: '1px solid #333', borderRadius: '4px', fontSize: '13px', textAlign: 'center' }}
                placeholder="MM/DD/YYYY"
              />
              <span>and continued until the above balance is paid in full. $20 late fee will be applied to your account if your payment is not received within five days of the payment due date. We accept cash, credit card (Visa or MasterCard) or Care Credit as payment at the office location. Personal check can be accepted by mail only.</span>
            </p>
            <p style={{ textAlign: 'justify', marginBottom: '15px' }}>
              The unpaid balance has been divided into monthly payments only to assist you in the payment of the total balance, and has no correlation with your appointments. The total fee must be paid in full prior to removing the orthodontic appliances and placement of retainers. If the account becomes delinquent, patient will not be seen for any treatment except emergency palliative treatment only to relieve pain with a $50 charge. If an account is in default more than three months, the case will be terminated and referred to a collection agency and an additional debanding fee of $600 to be charged to the account.
            </p>
            <div style={{ display: 'flex', alignItems: 'center', marginTop: '10px' }}>
              <strong style={{ marginRight: '10px' }}>Initial:</strong>
              <input
                type="text"
                value={formData.initial1}
                onChange={(e) => updateFormData('initial1', e.target.value)}
                style={{ width: '100px', padding: '8px', border: '2px solid #e0e0e0', borderRadius: '4px', fontSize: '14px' }}
                placeholder=""
              />
            </div>
          </div>

          {/* Section 2: Insurance */}
          <div style={{ marginBottom: '25px', fontSize: '13px', lineHeight: '1.8', color: '#495057' }}>
            <p style={{ textAlign: 'justify', marginBottom: '10px' }}>
              If you have orthodontic insurance coverage, please understand that your insurance coverage is a contract between you and your insurance company. We will only assist you in billing your insurance company. The estimated insurance amount is a quote and is subject to change due to variable calculation formulas used by insurance carriers. If your insurance does not pay the estimated amount, you are personally responsible for the total cost of treatment. All scheduled monthly orthodontic visits are mandatory due to the requirements of most insurance carriers. If you miss any of your appointments you will be responsible for the amount which was not paid by your insurance carrier as well as your monthly payment.
            </p>
            <p style={{ textAlign: 'justify', marginBottom: '10px' }}>
              Although after billing the dental insurance, the insurance may send payments to the subscriber (to you or the primary insured) therefore it is your responsibility to bring the payment to the office to clear your outstanding balance. If this payment is sent under the subscriber's name, you may need to cash the check first than bring in the amount to our office. All amounts paid to the subscriber and not paid toward the patient's account will be considered delinquent and therefore can be referred to a collection agency.
            </p>
            <p style={{ textAlign: 'justify', marginBottom: '15px' }}>
              If the patient is under 18, the parents, custodial parent or guardian will be legally responsible for any outstanding balance.
            </p>
            <div style={{ display: 'flex', alignItems: 'center', marginTop: '10px' }}>
              <strong style={{ marginRight: '10px' }}>Initial:</strong>
              <input
                type="text"
                value={formData.initial2}
                onChange={(e) => updateFormData('initial2', e.target.value)}
                style={{ width: '100px', padding: '8px', border: '2px solid #e0e0e0', borderRadius: '4px', fontSize: '14px' }}
                placeholder=""
              />
            </div>
          </div>

          {/* Section 3: General Dentistry */}
          <div style={{ marginBottom: '25px', fontSize: '13px', lineHeight: '1.8', color: '#495057' }}>
            <p style={{ textAlign: 'justify', marginBottom: '15px' }}>
              The Orthodontic Treatment Fee does not include any general dentistry such as fillings, expose & bond, extractions, crowns, or dental cleaning. Regular dental check-up and all necessary dental work must be completed and documented by your dentist prior to orthodontic bonding. Please ask your dentist to consult with us should you require any extensive dental treatment. Good oral hygiene is imperative to the success of orthodontic treatment. Should the patient require additional treatment period due to lack of patient's cooperation, such as no elastic wear / missed appointment an additional fee of $100 per month will be charged.
            </p>
            <div style={{ display: 'flex', alignItems: 'center', marginTop: '10px' }}>
              <strong style={{ marginRight: '10px' }}>Initial:</strong>
              <input
                type="text"
                value={formData.initial3}
                onChange={(e) => updateFormData('initial3', e.target.value)}
                style={{ width: '100px', padding: '8px', border: '2px solid #e0e0e0', borderRadius: '4px', fontSize: '14px' }}
                placeholder=""
              />
            </div>
          </div>

          {/* Section 4: Broken Appliances */}
          <div style={{ marginBottom: '25px', fontSize: '13px', lineHeight: '1.8', color: '#495057' }}>
            <p style={{ textAlign: 'justify', marginBottom: '15px' }}>
              Replacement of lost or broken appliances will result in appropriate additional charge plus laboratory fee. Broken band metal brackets are $35 and clear brackets are $45 each.
            </p>
            <div style={{ display: 'flex', alignItems: 'center', marginTop: '10px' }}>
              <strong style={{ marginRight: '10px' }}>Initial:</strong>
              <input
                type="text"
                value={formData.initial4}
                onChange={(e) => updateFormData('initial4', e.target.value)}
                style={{ width: '100px', padding: '8px', border: '2px solid #e0e0e0', borderRadius: '4px', fontSize: '14px' }}
                placeholder=""
              />
            </div>
          </div>

          {/* Section 5: Missed Appointments */}
          <div style={{ marginBottom: '25px', fontSize: '13px', lineHeight: '1.8', color: '#495057' }}>
            <p style={{ textAlign: 'justify', marginBottom: '10px' }}>
              We will try our best to accommodate you when making appointments. <strong>Any missed months or cancellation without a 24 hour notice is subject to a charge of $20 per appointment</strong>; therefore it is important that you notify us immediately if you need to change your appointment. You will be responsible to contact our office and reschedule any missed appointments. If you arrive more than 15 minutes past your appointment time, you may be asked to reschedule your appointment. Due to heavy afternoon scheduling you may be seen in a timelier manner by accepting morning appointments.
            </p>
            <p style={{ textAlign: 'justify', marginBottom: '15px' }}>
              Our office strives to provide high-quality orthodontic treatment to patients to create a beautiful smile. To accomplish this, however, it will require a mutual commitment by both the dentist and the patient. Without the patient's cooperation, the completion of orthodontic treatment cannot be achieved in a timely manner or may never be achieved. Therefore, if the patient misses three (3) or more months in a twelve (12) month period, we will then assume our patient/orthodontic relationship has been terminated and that you will seek all future dental treatment at another dental office.
            </p>
            <div style={{ display: 'flex', alignItems: 'center', marginTop: '10px' }}>
              <strong style={{ marginRight: '10px' }}>Initial:</strong>
              <input
                type="text"
                value={formData.initial5}
                onChange={(e) => updateFormData('initial5', e.target.value)}
                style={{ width: '100px', padding: '8px', border: '2px solid #e0e0e0', borderRadius: '4px', fontSize: '14px' }}
                placeholder=""
              />
            </div>
          </div>

          {/* Section 6: Records and Discontinuation */}
          <div style={{ marginBottom: '25px', fontSize: '13px', lineHeight: '1.8', color: '#495057' }}>
            <p style={{ textAlign: 'justify', marginBottom: '10px' }}>
              Should it be necessary to duplicate your records for transferable purposes due to re-location, we will forward a copy of your records within 10 working days to your new orthodontist for a fee of $100.
            </p>
            <p style={{ textAlign: 'justify', marginBottom: '10px' }}>
              If for any reason you discontinue treatment before completion, you must pay your previous delinquent monthly payment and debanding fee of $600. Additionally any discount previously given on the account will be retroactively voided, you will be responsible for the full amount of charges.
            </p>
            <p style={{ textAlign: 'justify', marginBottom: '15px' }}>
              Administration fee of $50 applies to all patients who are contracted for treatment but do not begin their treatment regardless of reason.
            </p>
            <div style={{ display: 'flex', alignItems: 'center', marginTop: '10px', marginBottom: '25px' }}>
              <strong style={{ marginRight: '10px' }}>Initial:</strong>
              <input
                type="text"
                value={formData.initial6}
                onChange={(e) => updateFormData('initial6', e.target.value)}
                style={{ width: '100px', padding: '8px', border: '2px solid #e0e0e0', borderRadius: '4px', fontSize: '14px' }}
                placeholder=""
              />
            </div>
            <p style={{ textAlign: 'justify', marginBottom: '15px' }}>
              It is understood and agreed that orthodontic treatment is elective in nature and all orthodontic appliances may be removed at any time without refund due to the following reasons: Non-payment of this account's financial obligations, excessive breakage ({'>'}5) of orthodontic appliances, non-compliance, or poor oral hygiene.
            </p>
            <div style={{ display: 'flex', alignItems: 'center', marginTop: '10px' }}>
              <strong style={{ marginRight: '10px' }}>Initial:</strong>
              <input
                type="text"
                value={formData.initial7}
                onChange={(e) => updateFormData('initial7', e.target.value)}
                style={{ width: '100px', padding: '8px', border: '2px solid #e0e0e0', borderRadius: '4px', fontSize: '14px' }}
                placeholder=""
              />
            </div>
          </div>
          
          {/* 약관 동의 서명 섹션 */}
          <div style={{ marginTop: '40px', padding: '30px', border: '2px solid #2c3e50', borderRadius: '10px', backgroundColor: '#fff' }}>
            <h4 style={{ marginTop: 0, marginBottom: '20px', color: '#2c3e50', textAlign: 'center' }}>Responsible Party Agreement</h4>
            
            <div style={{ marginBottom: '25px' }}>
              <label style={{ ...styles.label, marginBottom: '10px' }}>Signature of Responsible Party:</label>
              <div style={{ position: 'relative' }}>
                <canvas
                  ref={termsCanvasRef}
                  width={700}
                  height={150}
                  onMouseDown={startDrawingTerms}
                  onMouseMove={drawTerms}
                  onMouseUp={stopDrawingTerms}
                  onMouseLeave={stopDrawingTerms}
                  onTouchStart={startDrawingTermsTouch}
                  onTouchMove={drawTermsTouch}
                  onTouchEnd={stopDrawingTermsTouch}
                  onTouchCancel={cancelDrawingTermsTouch}
                  style={{
                    border: '2px solid #e0e0e0',
                    borderRadius: '6px',
                    cursor: 'crosshair',
                    backgroundColor: 'white',
                    width: '100%',
                    height: '150px',
                    touchAction: 'none' // 터치 스크롤 방지
                  }}
                />
                <button
                  type="button"
                  onClick={clearTermsSignature}
                  style={{
                    position: 'absolute',
                    top: '5px',
                    right: '5px',
                    padding: '5px 10px',
                    fontSize: '12px',
                    backgroundColor: '#ff6b6b',
                    color: 'white',
                    border: 'none',
                    borderRadius: '4px',
                    cursor: 'pointer'
                  }}
                >
                  Clear
                </button>
              </div>
            </div>

            <div className="form-row-responsive" style={styles.formRow}>
              <div>
                <label style={styles.label} htmlFor="responsiblePartyName">Name of Responsible Party:</label>
                <input
                  type="text"
                  id="responsiblePartyName"
                  value={formData.responsiblePartyName}
                  onChange={(e) => updateFormData('responsiblePartyName', e.target.value)}
                  style={styles.input}
                  placeholder="Enter full name"
                />
              </div>
              <div>
                <label style={styles.label} htmlFor="responsiblePartySignatureDate">Date:</label>
                <input
                  type="text"
                  id="responsiblePartySignatureDate"
                  value={formData.responsiblePartySignatureDate}
                  onChange={(e) => updateFormData('responsiblePartySignatureDate', e.target.value)}
                  style={styles.input}
                  placeholder="MM/DD/YYYY"
                />
              </div>
            </div>
          </div>
        </div>

        {/* 제출 버튼 */}
        <div style={{ textAlign: 'center', marginTop: '40px' }}>
          <button
            onClick={handleSubmit}
            disabled={loading}
            style={{
              ...styles.submitButton,
              backgroundColor: loading ? '#bdc3c7' : '#1976d2',
              cursor: loading ? 'not-allowed' : 'pointer'
            }}
          >
            {loading ? 'Submitting...' : 'Submit'}
          </button>
        </div>

        {/* 상태 메시지 */}
        {submitStatus && !loading && (
          <div style={{
            ...styles.statusMessage,
            backgroundColor: submitStatus.includes('failed') || submitStatus.includes('Error') ? '#f8d7da' : 
                           submitStatus.includes('Complete') ? '#d4edda' : '#d1ecf1',
            color: submitStatus.includes('failed') || submitStatus.includes('Error') ? '#721c24' : 
                   submitStatus.includes('Complete') ? '#155724' : '#0c5460',
            border: submitStatus.includes('failed') || submitStatus.includes('Error') ? '1px solid #f5c6cb' : 
                    submitStatus.includes('Complete') ? '1px solid #c3e6cb' : '1px solid #bee5eb'
          }}>
            {submitStatus}
          </div>
        )}

        {/* Footer */}
        <div style={styles.footer}>
          <p style={{ marginBottom: '5px', fontWeight: '600' }}>Smileland Dental</p>
        </div>
      </div>
    </div>
    </>
  );
}

