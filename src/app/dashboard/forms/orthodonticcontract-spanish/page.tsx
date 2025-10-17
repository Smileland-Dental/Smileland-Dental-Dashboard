'use client'

import React, { useState, useEffect, useRef } from "react";
import { doc, setDoc, deleteDoc } from "firebase/firestore";
import { db } from "@/lib/firebase.config";
import { getStorage, ref as storageRef, uploadBytes, getDownloadURL } from "firebase/storage";

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
      deposit: '',
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
        setSignatureData(dataUrl);
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
        setTermsSignatureData(dataUrl);
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

  // 제출 처리
  const handleSubmit = async () => {
    // 필수 필드 확인
    if (!formData.patientName || !formData.dob) {
      alert("Nombre del Paciente y Fecha de Nacimiento son campos obligatorios.");
      return;
    }

    const confirmSubmit = window.confirm("¿Está seguro de que desea enviar este contrato? Esto guardará los datos y generará un PDF.");
    if (!confirmSubmit) return;

    try {
      setLoading(true);
      setSubmitStatus('Generando PDF...');
      setProgress(20);
      
      // 1. PDF 생성
      const response = await fetch('/api/generate-orthodontic-contract-pdf-spanish', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          contractDate,
          formData: formData,
          signatureData: signatureData,
          termsSignatureData: termsSignatureData
        }),
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'PDF generation failed');
      }

      const blob = await response.blob();
      
      setSubmitStatus('Guardando en Firebase...');
      setProgress(40);

      // 2. Firestore에 데이터 저장
      const timestamp = Date.now();
      const docId = `contract_spanish_${contractDate}_${timestamp}`;
      
      const dataToSave = {
        ...formData,
        contractDate,
        timestamp: new Date().toISOString(),
        submitted: true,
        approved: true,
        language: 'spanish',
        signatureData: signatureData,
        termsSignatureData: termsSignatureData
      };

      await setDoc(doc(db, "orthodontic-contracts", docId), dataToSave);
      
      setSubmitStatus('Subiendo PDF al almacenamiento...');
      setProgress(60);
      
      // 3. Firebase Storage에 PDF 업로드
      const storage = getStorage();
      const fileName = `Contrato_de_Ortodoncia_${formData.patientName.replace(/\s+/g, '_')}_${contractDate}.pdf`;
      const pdfRef = storageRef(storage, `orthodontic-contracts/${docId}/${fileName}`);
      
      await uploadBytes(pdfRef, blob);
      const downloadURL = await getDownloadURL(pdfRef);
      
      // Firestore에 PDF URL 추가
      await setDoc(doc(db, "orthodontic-contracts", docId), {
        ...dataToSave,
        pdfUrl: downloadURL,
        pdfFileName: fileName
      });
      
      setSubmitStatus('¡Completo!');
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
          deposit: '',
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
        alert('✅ ¡El contrato se ha guardado exitosamente y se ha enviado al administrador!');
      }, 1000);
      
    } catch (error) {
      console.error('Submit error:', error);
      setSubmitStatus('❌ Error al enviar: ' + (error as Error).message);
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
          <h1 style={{ ...styles.header, marginBottom: 0, paddingBottom: 0, borderBottom: 'none', flex: 1, textAlign: 'center' }}>Contrato de Ortodoncia</h1>
          <a 
            href="/dashboard/forms/orthodonticcontract"
            style={{
              padding: '10px 20px',
              backgroundColor: '#1976d2',
              color: 'white',
              textDecoration: 'none',
              borderRadius: '6px',
              fontSize: '14px',
              fontWeight: '600',
              transition: 'background-color 0.2s',
              cursor: 'pointer',
              border: 'none',
              boxShadow: '0 2px 8px rgba(25, 118, 210, 0.3)'
            }}
            onMouseEnter={(e) => e.currentTarget.style.backgroundColor = '#1565c0'}
            onMouseLeave={(e) => e.currentTarget.style.backgroundColor = '#1976d2'}
          >
            English
          </a>
        </div>

        {/* Patient's Name and DOB */}
        <div className="form-row-responsive" style={styles.formRow}>
          <div>
            <label style={styles.label} htmlFor="patientName">Nombre d. Paciente *</label>
            <input
              type="text"
              id="patientName"
              value={formData.patientName}
              onChange={(e) => updateFormData('patientName', e.target.value)}
              style={styles.input}
              placeholder="Ingrese el nombre completo del paciente"
              required
            />
          </div>
          <div>
            <label style={styles.label} htmlFor="dob">Fecha d. Naci. *</label>
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
            <label style={styles.label} htmlFor="responsibleParty">Persona Responsable</label>
            <input
              type="text"
              id="responsibleParty"
              value={formData.responsibleParty}
              onChange={(e) => updateFormData('responsibleParty', e.target.value)}
              style={styles.input}
              placeholder="Ingrese el nombre de la persona responsable"
            />
          </div>
          <div>
            <label style={styles.label} htmlFor="relationship">Relacion al Paciente</label>
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
            <label style={styles.label} htmlFor="ssn"># d. Seguro Social</label>
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
            <label style={styles.label} htmlFor="driversLicense"># Lic. d. Manejar</label>
            <input
              type="text"
              id="driversLicense"
              value={formData.driversLicense}
              onChange={(e) => updateFormData('driversLicense', e.target.value)}
              style={styles.input}
              placeholder="Ingrese el nombre que aparece en la licencia de conducir"
            />
          </div>
        </div>

        {/* Type of Treatment */}
        <div style={styles.formGroup}>
          <label style={styles.label}>Tipo de tratamiento</label>
          <div style={styles.radioGroup}>
            <label style={styles.radioOption}>
              <input
                type="radio"
                name="typeOfTreatment"
                value="Limited"
                checked={formData.typeOfTreatment === 'Limited'}
                onChange={(e) => updateFormData('typeOfTreatment', e.target.value)}
                style={{ cursor: 'pointer' }}
              />
              <span>Limitado</span>
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
              <span>Fase I</span>
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
              <span>Fase II</span>
            </label>
          </div>
        </div>

        {/* Services Required */}
        <div style={{ marginTop: '40px' }}>
          <div style={styles.sectionTitle}>Servicios Requeridos</div>
          <div className="services-grid" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '15px', marginTop: '15px' }}>
            {[
              { label: 'P Radiografía Panorámica', price: 120 },
              { label: 'Radiografía Lateral de Cráneo', price: 145 },
              { label: 'Moldes Diagnósticos', price: 120 },
              { label: 'Imágenes Faciales/Orales', price: 80 },
              { label: 'Tratamiento de Ortodoncia Periódico', price: 4235 },
              { label: 'Retenedores de Ortodoncia(sup/Inf)', price: 700 }
            ].map((service) => {
              const value = `${service.label}|$${service.price}`;
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
                  <span style={styles.servicePrice}>${service.price.toLocaleString()}</span>
                </label>
              );
            })}
          </div>
        </div>

        {/* Additional appliance (if necessary) */}
        <div style={{ marginTop: '40px' }}>
          <div style={styles.sectionTitle}>Aparatos (Si Necesario)</div>
          <div className="services-grid" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '15px', marginTop: '15px' }}>
            {[
              { label: 'Nance', price: 250 },
              { label: 'Quad Helix', price: 250 },
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
            <h3 style={{ marginTop: 0, marginBottom: '20px', color: '#2c3e50', textAlign: 'center' }}>Primera Opicón</h3>
            
            <div style={styles.formGroup}>
              <label style={styles.label}>Total de Servicios:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.firstOption?.treatment || ''}
                  onChange={(e) => setFormData({
                    ...formData,
                    firstOption: { ...(formData.firstOption || {}), treatment: e.target.value }
                  })}
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%' }}
                />
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Aparat:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.firstOption?.appliance || ''}
                  onChange={(e) => setFormData({
                    ...formData,
                    firstOption: { ...(formData.firstOption || {}), appliance: e.target.value }
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
                  value={formData.firstOption?.deposit}
                  onChange={(e) => setFormData({
                    ...formData,
                    firstOption: { ...(formData.firstOption || {}), deposit: e.target.value }
                  })}
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%' }}
                />
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Total Parcial:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.firstOption?.subtotal}
                  onChange={(e) => setFormData({
                    ...formData,
                    firstOption: { ...(formData.firstOption || {}), subtotal: e.target.value }
                  })}
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%' }}
                />
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Estimado de Seguro:</label>
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
              <label style={styles.label}>Balance Neto:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.firstOption?.netBalance}
                  onChange={(e) => setFormData({
                    ...formData,
                    firstOption: { ...(formData.firstOption || {}), netBalance: e.target.value }
                  })}
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%' }}
                />
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Periodo de Tratamiento Estimado:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.firstOption?.estimatedTreatmentPeriod}
                  onChange={(e) => setFormData({
                    ...formData,
                    firstOption: { ...(formData.firstOption || {}), estimatedTreatmentPeriod: e.target.value }
                  })}
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%' }}
                />
                <span style={{ color: '#666', marginLeft: '5px', whiteSpace: 'nowrap' }}>meses</span>
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Pago Mansual:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.firstOption?.monthlyPayment}
                  onChange={(e) => setFormData({
                    ...formData,
                    firstOption: { ...(formData.firstOption || {}), monthlyPayment: e.target.value }
                  })}
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%' }}
                />
              </div>
            </div>
          </div>

          {/* Second Option */}
          <div style={{ border: '2px solid #2c3e50', borderRadius: '10px', padding: '20px', backgroundColor: '#f8f9fa' }}>
            <h3 style={{ marginTop: 0, marginBottom: '20px', color: '#2c3e50', textAlign: 'center' }}>Segunda Opicón</h3>
            
            <div style={styles.formGroup}>
              <label style={styles.label}>Total de Servicios:</label>
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
              <label style={styles.label}>Aparato</label>
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
              <label style={styles.label}>Total Parcial:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.secondOption?.subtotal}
                  onChange={(e) => setFormData({
                    ...formData,
                    secondOption: { ...(formData.secondOption || {}), subtotal: e.target.value }
                  })}
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%' }}
                />
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Estimado de Seguro:</label>
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
              <label style={styles.label}>Balance Neto:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.secondOption?.netBalance}
                  onChange={(e) => setFormData({
                    ...formData,
                    secondOption: { ...(formData.secondOption || {}), netBalance: e.target.value }
                  })}
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%' }}
                />
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Periodo de Tratamiento Estimado:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.secondOption?.estimatedTreatmentPeriod}
                  onChange={(e) => setFormData({
                    ...formData,
                    secondOption: { ...(formData.secondOption || {}), estimatedTreatmentPeriod: e.target.value }
                  })}
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%' }}
                />
                <span style={{ color: '#666', marginLeft: '5px', whiteSpace: 'nowrap' }}>meses</span>
              </div>
            </div>

            <div style={styles.formGroup}>
              <label style={styles.label}>Pago Mansual:</label>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: '6px', backgroundColor: 'white', padding: '0 15px' }}>
                <span style={{ color: '#666', marginRight: '5px' }}>$</span>
                <input
                  type="text"
                  placeholder="0.00"
                  value={formData.secondOption?.monthlyPayment}
                  onChange={(e) => setFormData({
                    ...formData,
                    secondOption: { ...(formData.secondOption || {}), monthlyPayment: e.target.value }
                  })}
                  style={{ ...styles.input, border: 'none', padding: '12px 0', width: '100%' }}
                />
              </div>
            </div>
          </div>
        </div>

        {/* Terms and Conditions */}
        <div style={{ marginTop: '30px', padding: '20px', backgroundColor: '#f8f9fa', borderRadius: '8px', border: '1px solid #dee2e6' }}>
          <div style={{ margin: 0, fontSize: '13px', lineHeight: '1.8', color: '#495057' }}>
            <div style={{ marginBottom: '8px' }}>* Los cargos para procedimientos ya realizados en un paciente no son reembolsables.</div>
            <div style={{ marginBottom: '8px' }}>* Ofrecemos una reduccion de tarifa de 5% por pago completo en efectivo, y una reduccion de tarifa de 3% si paga por completo con tarjeta de crédito(Visa o Mastercard) al inicio del tratamiento.</div>
            <div style={{ marginBottom: '8px' }}>* Bracketes translucientes son disponibles por un costo adicional de $300 por arco.</div>
            <div style={{ marginBottom: '8px' }}>* Se cobrara una tarifa de administracion de $100 si la cuenta requiere un nuevo contracto.</div>
            <div style={{ marginBottom: '8px' }}>* Estimado de seguros no es una cantidad garantizada de ninguna manera y sujetas a cambios debido a las fórmulas de cálculo variable utilizada por las compañías de seguros.</div>
            <div style={{ marginBottom: '8px' }}>* Este estimado es valido por 30 días.</div>
          </div>
        </div>

        {/* Quote Presented By and Signature Section */}
        <div style={{ marginTop: '40px' }}>
          {/* Quote Presented By and Date */}
          <div className="form-row-responsive" style={styles.formRow}>
            <div>
              <label style={styles.label} htmlFor="quotePresentedBy">Esitmado Presenta do Por:</label>
              <input
                type="text"
                id="quotePresentedBy"
                value={formData.quotePresentedBy}
                onChange={(e) => updateFormData('quotePresentedBy', e.target.value)}
                style={styles.input}
                placeholder="Ingrese el nombre completo"
              />
            </div>
            <div>
              <label style={styles.label} htmlFor="quotePresentedDate">Fecha:</label>
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
              <label style={styles.label}>Firma de Persona Resporsable:</label>
              <div style={{ position: 'relative' }}>
                <canvas
                  ref={canvasRef}
                  width={500}
                  height={150}
                  onMouseDown={startDrawing}
                  onMouseMove={draw}
                  onMouseUp={stopDrawing}
                  onMouseLeave={stopDrawing}
                  style={{
                    border: '2px solid #e0e0e0',
                    borderRadius: '6px',
                    cursor: 'crosshair',
                    backgroundColor: 'white',
                    width: '100%',
                    height: '150px'
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
              <label style={styles.label} htmlFor="signatureDate">Fecha:</label>
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
          {/* Section 1: Payment Terms */}
          <div style={{ marginBottom: '25px', fontSize: '13px', lineHeight: '1.8', color: '#495057' }}>
            <p style={{ textAlign: 'justify', marginBottom: '15px', display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: '5px' }}>
              <span>El saldo pendiente de</span>
              <input
                type="text"
                value={formData.unpaidBalance}
                onChange={(e) => updateFormData('unpaidBalance', e.target.value)}
                style={{ width: '100px', padding: '4px 8px', border: '1px solid #333', borderRadius: '4px', fontSize: '13px', textAlign: 'center' }}
                placeholder=""
              />
              <span>será pagado en</span>
              <input
                type="text"
                value={formData.paymentMonths}
                onChange={(e) => updateFormData('paymentMonths', e.target.value)}
                style={{ width: '60px', padding: '4px 8px', border: '1px solid #333', borderRadius: '4px', fontSize: '13px', textAlign: 'center' }}
                placeholder=""
              />
              <span>meses en la cantidad de</span>
              <input
                type="text"
                value={formData.monthlyPaymentAmount}
                onChange={(e) => updateFormData('monthlyPaymentAmount', e.target.value)}
                style={{ width: '100px', padding: '4px 8px', border: '1px solid #333', borderRadius: '4px', fontSize: '13px', textAlign: 'center' }}
                placeholder=""
              />
              <span>cada uno, por el día 1o o 15 de cada mes, comenzando él</span>
              <input
                type="text"
                value={formData.paymentBeginDate}
                onChange={(e) => updateFormData('paymentBeginDate', e.target.value)}
                style={{ width: '120px', padding: '4px 8px', border: '1px solid #333', borderRadius: '4px', fontSize: '13px', textAlign: 'center' }}
                placeholder="MM/DD/YYYY"
              />
              <span>y continuando hasta que el saldo anterior sea pagado en su totalidad. Un cargo de $20.00 será aplicado a su cuenta si su pago no se recibe en 5 días de la fecha de vencimiento Smileland acepta dinero en efectivo, tarjetas de crédito y Care Credit como forma de pago. Cheques personales podrán ser aceptados solo por correo.</span>
            </p>
            <p style={{ textAlign: 'justify', marginBottom: '15px' }}>
              El saldo pendiente se ha dividido en pagos mensuales solamente como ayuda del saldo total y no esta correlacionada con sus citas. El costo total debe ser pagado por completo antes de retirar los aparatos de ortodoncia y la colocación de los retenedores. Aparatos de retención se pondrán a la conclusión del tratamiento ortodontico. Si la cuenta llega a estar atrasada, el paciente no será visto para ningún tratamiento solo para emergencia de tratamiento paliativo sólo para aliviar dolor con $50 cargo. Si su cuenta está en atraso más de 3 meses, el caso será terminado y será referido a la agencia de la colección y el honorario adicional será de $600 cargado a la cuenta.
            </p>
            <div style={{ display: 'flex', alignItems: 'center', marginTop: '10px' }}>
              <strong style={{ marginRight: '10px' }}>Iniciales:</strong>
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
              Si usted tiene cobertura de ortodoncia por favor entienda que la cobertura de seguro es un contrato entre usted y su compañía de seguros. Nosotros le ayudaremos en la facturación de su compañía de seguros. Si su compañía de seguros no paga su parte, será su responsabilidad el seguimiento con ellos directamente. Usted es personalmente responsable del costo total de tratamiento recibido. Si el seguro no paga la cuenta del tratamiento estimado, usted será responsable del saldo.
            </p>
            <p style={{ textAlign: 'justify', marginBottom: '10px' }}>
              Aunque después de Facturar el seguro dental, el seguro puede enviar pagos al suscriptor (a usted o al asegurado principal) por lo tanto es su responsabilidad de traer el pago a la oficina para eliminar su saldo pendiente. Si el pago se envía bajo el nombre de los suscriptores, es posible que tenga cobrar el cheque primero y luego traer la cantidad a nuestra oficina. Toda la cantidad pagada al suscriptor y no pagada hacia la cuenta del paciente será considerado delincuente y por lo tanto pueden ser referidos a una agencia de cobranzas.
            </p>
            <p style={{ textAlign: 'justify', marginBottom: '15px' }}>
              Si el paciente es menor de 18 años, los padres, padre o tutor serán legalmente responsables de cualquier saldo pendiente.
            </p>
            <div style={{ display: 'flex', alignItems: 'center', marginTop: '10px' }}>
              <strong style={{ marginRight: '10px' }}>Iniciales:</strong>
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
              Los Cargos De tratamiento de ortodoncia no incluyen ningún costo de odontología general, tales como rellenos, exponer y adherir, extracciones, coronas o limpiezas. Los chequeos dentales regulares y todo el trabajo dental necesario deben ser completados y documentados por su dentista antes de la adhesión de ortodoncia. Por favor pida que su dentista consulte con nosotros si requiere cualquier tratamiento dental extenso. La buena higiene oral es esencial para el éxito del tratamiento de ortodoncia. Si el paciente requiere un periodo de tratamiento adicional debido al mal higiene bucal, no usare elástico, y faltas de citas, un cargo de $100 por cada mes será agregado a su cuenta.
            </p>
            <div style={{ display: 'flex', alignItems: 'center', marginTop: '10px' }}>
              <strong style={{ marginRight: '10px' }}>Iniciales:</strong>
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
              El reemplazo de aplicaciones pérdidas o rotas resultara en un costo adicional más el costo del laboratorio. El costo de los soportes de metal es $35 y los soportes claros son $45 por cada uno.
            </p>
            <div style={{ display: 'flex', alignItems: 'center', marginTop: '10px' }}>
              <strong style={{ marginRight: '10px' }}>Iniciales:</strong>
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
              Trataremos de alojarle al hacer sus citas. <strong>Cualquier mes perdido o cancelación sin un aviso de 24 horas será sujeta a un cargo de $20 por cita; por lo tanto es importante que nos notifique inmediatamente si tiene que cambiar su cita.</strong> Será responsable de contactar a nuestra oficina y reprogramar cualquier cita perdida. Si llega más de 15 minutos tarde a su cita, es probable que tuviera que reprogramar su cita. Debido a muchas citas por la tarde, usted puede ser visto en un tiempo más razonable si acepta citas en la mañana.
            </p>
            <p style={{ textAlign: 'justify', marginBottom: '15px' }}>
              Nuestra oficina se esfuerza para proveer tratamiento de ortodoncia de alta calidad a pacientes para crear una sonrisa hermosa. Para llevar a cabo esto, requerirá un compromiso mutuo tanto del dentista como del paciente. Sin la cooperación de los pacientes, la finalización del tratamiento de ortodoncia no puede ser conseguida en una manera oportuna o nunca podría ser lograda. Por lo tanto, si el paciente pierde tres (3) o más meses en un período de doce (12) meses, supondremos entonces que nuestra relación de paciente y oficina de ortodoncia ha sido terminada y que buscará todo el futuro tratamiento dental en otro consultorio dental.
            </p>
            <div style={{ display: 'flex', alignItems: 'center', marginTop: '10px' }}>
              <strong style={{ marginRight: '10px' }}>Iniciales:</strong>
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
              Si es necesario duplicar sus expedientes para los propósitos de la transferencia debido a la relocalización, remitiremos una copia de sus expedientes dentro de 10 días laborables a su nuevo ortodontista con un honorario de $100.00.
            </p>
            <p style={{ textAlign: 'justify', marginBottom: '10px' }}>
              Si por alguna razón usted interrumpe el tratamiento antes de que todos los pagos se hayan hecho, usted debe pagar su antiguo balance y una cuota de $600 para quitar aparatos. El cargo administrativo de $50.00 se aplica a todos los pacientes que firmen su contrato para el tratamiento y por algún motivo no se presenten para comenzar su tratamiento.
            </p>
            <p style={{ textAlign: 'justify', marginBottom: '15px' }}>
              Cargo de administración de $50 se aplicara a todos los pacientes que son contratados para recibir tratamiento, pero no empiezan su tratamiento, por cualquier razón.
            </p>
            <div style={{ display: 'flex', alignItems: 'center', marginTop: '10px', marginBottom: '25px' }}>
              <strong style={{ marginRight: '10px' }}>Iniciales:</strong>
              <input
                type="text"
                value={formData.initial6}
                onChange={(e) => updateFormData('initial6', e.target.value)}
                style={{ width: '100px', padding: '8px', border: '2px solid #e0e0e0', borderRadius: '4px', fontSize: '14px' }}
                placeholder=""
              />
            </div>
            <p style={{ textAlign: 'justify', marginBottom: '15px' }}>
              Es entendido y acordado que el tratamiento de ortodoncia es electivo en la naturaleza y todos los aparatos ortodrómicos pueden ser removidos en cualquier momento sin reembolso por las siguientes razones: Falta de pago de obligaciones financieras de la cuenta, rotura excesiva (5 o más) de aparatos de ortodoncia, incumplimiento, o mala higiene bucal son medidas que serían perjudiciales para la función opcional o estética.
            </p>
            <div style={{ display: 'flex', alignItems: 'center', marginTop: '10px' }}>
              <strong style={{ marginRight: '10px' }}>Iniciales:</strong>
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
            <div style={{ marginBottom: '25px' }}>
              <label style={{ ...styles.label, marginBottom: '10px' }}>Firma de Persona Responsable:</label>
              <div style={{ position: 'relative' }}>
                <canvas
                  ref={termsCanvasRef}
                  width={700}
                  height={150}
                  onMouseDown={startDrawingTerms}
                  onMouseMove={drawTerms}
                  onMouseUp={stopDrawingTerms}
                  onMouseLeave={stopDrawingTerms}
                  style={{
                    border: '2px solid #e0e0e0',
                    borderRadius: '6px',
                    cursor: 'crosshair',
                    backgroundColor: 'white',
                    width: '100%',
                    height: '150px'
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
                <label style={styles.label} htmlFor="responsiblePartyName">Nombre de Persona Responsable:</label>
                <input
                  type="text"
                  id="responsiblePartyName"
                  value={formData.responsiblePartyName}
                  onChange={(e) => updateFormData('responsiblePartyName', e.target.value)}
                  style={styles.input}
                  placeholder="Ingrese el nombre completo"
                />
              </div>
              <div>
                <label style={styles.label} htmlFor="responsiblePartySignatureDate">Fecha:</label>
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
            {loading ? 'Enviando...' : 'Enviar'}
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

