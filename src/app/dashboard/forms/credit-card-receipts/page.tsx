'use client';

import { useState, useEffect } from 'react';
import { doc, setDoc } from 'firebase/firestore';
import { ref, uploadBytes, deleteObject } from 'firebase/storage';
import { db, storage } from '@/lib/firebase.config';

const RECEIPT_STORAGE_FOLDER = 'receipts';

type ReceiptUploadStatus = 'pending' | 'uploading' | 'failed';

// 🔒 보안: 입력 데이터 sanitization 함수
const sanitizeInput = (input: string, maxLength: number = 1000): string => {
  if (!input) return '';
  return String(input)
    .replace(/[<>]/g, '') // Remove < and >
    .substring(0, maxLength)
    .trim();
};

// 🔒 보안: 파일명 sanitization (path traversal 방지)
const sanitizeFileName = (fileName: string): string => {
  return fileName
    .replace(/[<>:"/\\|?*]/g, '') // Remove dangerous characters
    .replace(/\.\./g, '') // Remove path traversal attempts
    .substring(0, 255); // Limit length
};

const ALLOWED_RECEIPT_MIME = new Set(['image/jpeg', 'image/jpg', 'image/png']);
const ALLOWED_RECEIPT_EXT = /\.(jpe?g|png)$/i;

/** Receipt: JPG, JPEG, PNG only (MIME + extension if type missing) */
function isAllowedReceiptImageFile(file: File): boolean {
  const type = (file.type || '').toLowerCase().trim();
  if (type && ALLOWED_RECEIPT_MIME.has(type)) {
    return true;
  }
  if (type && type.startsWith('image/')) {
    return false;
  }
  return ALLOWED_RECEIPT_EXT.test(file.name || '');
}

// Interfaces
interface ReceiptFile {
  name: string;
  file?: File;
  uploadStatus?: ReceiptUploadStatus;
}

interface Purchase {
  date: string;
  vendor: string;
  reason: string;
  amount: string;
  description: string;
  receiptFiles: ReceiptFile[];
}

const CreditCardReceipts = () => {
  const [formData, setFormData] = useState({
    name: '',
    cardLastFour: '',
    date: new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' }),
    office: '' 
  });
  const [purchases, setPurchases] = useState<Purchase[]>([{
    date: '',
    vendor: '',
    reason: '',
    amount: '',
    description: '',
    receiptFiles: []
  }]);
  const [loading, setLoading] = useState(false);
  const [submitStatus, setSubmitStatus] = useState('');

  useEffect(() => {
    if (
      process.env.NODE_ENV === 'production' &&
      typeof window !== 'undefined' &&
      window.location.protocol !== 'https:'
    ) {
      window.location.href = window.location.href.replace('http:', 'https:');
    }
  }, []);

  const officeOptions = ['Bernard', 'California', 'Corporate', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

  const generateSubmissionId = () => {
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const randomSuffix = Math.random().toString(36).substring(2, 8);
    return `${timestamp}_${randomSuffix}`;
  };

  const deleteUploadedFiles = async (storageFileNames: string[]) => {
    if (storageFileNames.length === 0) {
      return;
    }

    await Promise.all(
      storageFileNames.map((fileName) =>
        deleteObject(ref(storage, `${RECEIPT_STORAGE_FOLDER}/${fileName}`)).catch(() => undefined)
      )
    );
  };

  const saveData = async (currentSubmissionId: string, purchaseList: Purchase[]) => {
    const sanitizedName = sanitizeInput(formData.name, 100);
    const sanitizedCardNumber = sanitizeInput(formData.cardLastFour, 4);
    const sanitizedOffice = sanitizeInput(formData.office, 50);
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const uniqueId = `${sanitizedName}_${sanitizedCardNumber}_${timestamp}`;

    const docRef = doc(db, 'credit-card-receipts', uniqueId);
    await setDoc(docRef, {
      name: sanitizedName,
      cardNumber: sanitizedCardNumber,
      date: formData.date,
      office: sanitizedOffice,
      submissionId: currentSubmissionId,
      data: collectFormData(purchaseList),
      lastUpdated: new Date(),
      createdAt: new Date()
    }, { merge: true });
  };

  const uploadAllReceiptFiles = async (
    purchaseList: Purchase[],
    currentSubmissionId: string,
    onFileStatusChange?: (purchaseIndex: number, fileIndex: number, status: ReceiptUploadStatus) => void
  ): Promise<Purchase[]> => {
    const sanitizedName = sanitizeInput(formData.name, 100);
    const sanitizedCardNumber = sanitizeInput(formData.cardLastFour, 4);
    const baseTimestamp = Date.now();
    const uploadedStorageFiles: string[] = [];
    const updatedPurchases: Purchase[] = [];

    try {
      for (let purchaseIndex = 0; purchaseIndex < purchaseList.length; purchaseIndex++) {
        const purchase = purchaseList[purchaseIndex];
        const uploadedFiles: ReceiptFile[] = [];

        for (let fileIndex = 0; fileIndex < purchase.receiptFiles.length; fileIndex++) {
          const receiptFile = purchase.receiptFiles[fileIndex];
          onFileStatusChange?.(purchaseIndex, fileIndex, 'uploading');

          const sequenceNumber = String(baseTimestamp + purchaseIndex * 1000 + fileIndex).padStart(15, '0');
          const fileName = `${sanitizedName}_${sanitizedCardNumber}_${currentSubmissionId}_purchase${purchaseIndex + 1}_${sequenceNumber}_${receiptFile.name}`;

          const uploadResult = await uploadFile(receiptFile.file!, fileName);
          if (!uploadResult?.success || !uploadResult.fileName) {
            onFileStatusChange?.(purchaseIndex, fileIndex, 'failed');
            throw new Error(uploadResult?.error || `Failed to upload ${receiptFile.name}`);
          }

          uploadedStorageFiles.push(uploadResult.fileName);
          uploadedFiles.push({ name: uploadResult.fileName });
        }

        updatedPurchases.push({ ...purchase, receiptFiles: uploadedFiles });
      }

      return updatedPurchases;
    } catch (error) {
      await deleteUploadedFiles(uploadedStorageFiles);
      throw error;
    }
  };

  const uploadFile = async (file: File, fileName: string) => {
    try {
      const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10MB
      if (file.size > MAX_FILE_SIZE) {
        return { success: false, error: 'File size exceeds limit' };
      }

      if (!isAllowedReceiptImageFile(file)) {
        return { success: false, error: 'Only JPG, JPEG, or PNG images are allowed' };
      }
      
      const sanitizedFileName = sanitizeFileName(fileName);
      
      const storageRef = ref(storage, `${RECEIPT_STORAGE_FOLDER}/${sanitizedFileName}`);
      await uploadBytes(storageRef, file);

      return {
        success: true,
        fileName: sanitizedFileName,
      };
    } catch {
      return { success: false, error: 'File upload failed' };
    }
  };

  const isSubmissionComplete = (): boolean => {
    if (!formData.office?.trim() || !formData.date?.trim() || !formData.name?.trim()) {
      return false;
    }
    if (!/^\d{4}$/.test((formData.cardLastFour || '').trim())) {
      return false;
    }
    for (const p of purchases) {
      if (!p.date?.trim() || !p.vendor?.trim() || !p.reason?.trim() || !p.description?.trim()) {
        return false;
      }
      if (!p.amount?.toString().trim()) {
        return false;
      }
      const amt = parseFloat(String(p.amount).replace(/,/g, ''));
      if (!isFinite(amt) || amt < 0) {
        return false;
      }
      if (!p.receiptFiles || p.receiptFiles.length === 0) {
        return false;
      }
    }
    return true;
  };

  const updateReceiptFileStatus = (
    purchaseIndex: number,
    fileIndex: number,
    status: ReceiptUploadStatus
  ) => {
    setPurchases((prev) => {
      const next = [...prev];
      const receiptFiles = [...next[purchaseIndex].receiptFiles];
      receiptFiles[fileIndex] = { ...receiptFiles[fileIndex], uploadStatus: status };
      next[purchaseIndex] = { ...next[purchaseIndex], receiptFiles };
      return next;
    });
  };

  const resetReceiptFileStatuses = () => {
    setPurchases((prev) =>
      prev.map((purchase) => ({
        ...purchase,
        receiptFiles: purchase.receiptFiles.map((file) => ({
          ...file,
          uploadStatus: undefined,
        })),
      }))
    );
  };

  const handleSubmit = async () => {
    if (!isSubmissionComplete()) {
      alert(
        'Please fill in all fields for every purchase, including at least one receipt image (JPG, JPEG, or PNG) for each purchase.'
      );
      return;
    }

    let uploadedStorageFiles: string[] = [];

    try {
      setLoading(true);
      setSubmitStatus('📤 Uploading receipts...');

      const currentSubmissionId = generateSubmissionId();
      const purchasesWithUploadedFiles = await uploadAllReceiptFiles(
        purchases,
        currentSubmissionId,
        updateReceiptFileStatus
      );

      uploadedStorageFiles = purchasesWithUploadedFiles.flatMap((purchase) =>
        purchase.receiptFiles.map((file) => file.name)
      );

      setSubmitStatus('📤 Saving submission...');
      await saveData(currentSubmissionId, purchasesWithUploadedFiles);

      setFormData({ name: '', cardLastFour: '', date: new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' }), office: '' });
      setPurchases([{ date: '', vendor: '', reason: '', amount: '', description: '', receiptFiles: [] }]);

      setSubmitStatus('✅ Submitted successfully!');
      
      setTimeout(() => setSubmitStatus(''), 7000);

    } catch (error) {
      console.error('Submit error:', error);

      if (uploadedStorageFiles.length > 0) {
        setSubmitStatus('↩️ Cleaning up uploaded files...');
        await deleteUploadedFiles(uploadedStorageFiles);
        resetReceiptFileStatuses();
      } else {
        setPurchases((prev) =>
          prev.map((purchase) => ({
            ...purchase,
            receiptFiles: purchase.receiptFiles.map((file) => ({
              ...file,
              uploadStatus: file.uploadStatus === 'failed' ? 'failed' : 'pending',
            })),
          }))
        );
      }

      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      setSubmitStatus(
        `❌ Submission failed: ${errorMessage}. No data was saved. Please try again.`
      );
      setTimeout(() => setSubmitStatus(''), 7000);
    } finally {
      setLoading(false);
    }
  };

  const collectFormData = (purchaseList: Purchase[]) => {
    return purchaseList.map((purchase: Purchase, index: number) => ({
      name: sanitizeInput(formData.name, 100),
      cardLastFour: sanitizeInput(formData.cardLastFour, 4),
      date: purchase.date, 
      office: sanitizeInput(formData.office, 50),
      vendor: sanitizeInput(purchase.vendor, 200),
      reason: sanitizeInput(purchase.reason, 500),
      amount: sanitizeInput(purchase.amount, 20),
      description: sanitizeInput(purchase.description, 200),
      receiptFiles: purchase.receiptFiles.map((file: ReceiptFile) => sanitizeFileName(file.name)).join(', '),
      index: index + 1
    }));
  };

  const addPurchaseRow = () => {
    setPurchases((prev) => [...prev, {
      date: '',
      vendor: '',
      reason: '',
      amount: '',
      description: '',
      receiptFiles: []
    }]);
  };

  const removePurchaseRow = (index: number) => {
    setPurchases((prev) => (prev.length > 1 ? prev.filter((_, i) => i !== index) : prev));
  };

  const updatePurchase = (index: number, field: string, value: unknown) => {
    setPurchases((prev) => {
      const next = [...prev];
      next[index] = { ...next[index], [field]: value };
      return next;
    });
  };

  const handleFileUpload = (index: number, files: FileList) => {
    if (!files || files.length === 0) {
      return;
    }

    const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10MB
    const MAX_FILES = 10;
    const fileArray = Array.from(files).slice(0, MAX_FILES);

    const imageFiles = fileArray.filter(file => {
      if (file.size > MAX_FILE_SIZE) {
        alert(`❌ File "${sanitizeFileName(file.name)}" exceeds the 10MB size limit.`);
        return false;
      }
      if (isAllowedReceiptImageFile(file)) {
        return true;
      }
      alert(
        `❌ Only JPG, JPEG, or PNG files are allowed. File "${sanitizeFileName(file.name)}" was rejected.`
      );
      return false;
    });

    if (imageFiles.length === 0) {
      alert('❌ No valid files selected. Please upload JPG, JPEG, or PNG images only.');
      return;
    }

    const pendingFiles: ReceiptFile[] = imageFiles.map((file) => ({
      name: sanitizeFileName(file.name),
      file,
      uploadStatus: 'pending',
    }));

    updatePurchase(index, 'receiptFiles', pendingFiles);
  };

  const styles = {
    body: {
      fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
      lineHeight: 1.6,
      color: '#333',
      background: 'linear-gradient(135deg, #a8edea 0%, #fed6e3 100%)',
      minHeight: '100vh',
      margin: 0,
      padding: 0
    },
    container: {
      maxWidth: '1600px',
      margin: '0 auto',
      padding: '15px',
      minHeight: '100vh',
      display: 'flex',
      flexDirection: 'column' as const,
      justifyContent: 'center',
      width: '100%',
      boxSizing: 'border-box' as const
    },
    header: {
      textAlign: 'center' as const,
      marginBottom: '30px',
      backgroundColor: 'rgba(255, 255, 255, 0.95)',
      padding: '25px',
      borderRadius: '15px',
      boxShadow: '0 8px 32px rgba(0, 0, 0, 0.1)'
    },
    title: {
      color: '#2c3e50',
      fontSize: '2.5em',
      fontWeight: 'bold',
      margin: '0 0 10px 0',
      textShadow: '2px 2px 4px rgba(0,0,0,0.1)'
    },
    form: {
      backgroundColor: 'rgba(255, 255, 255, 0.95)',
      padding: '30px',
      borderRadius: '15px',
      boxShadow: '0 8px 32px rgba(0, 0, 0, 0.1)',
      backdropFilter: 'blur(10px)'
    },
    formGroup: {
      marginBottom: '20px'
    },
    label: {
      display: 'block',
      marginBottom: '8px',
      fontWeight: '600',
      color: '#2c3e50',
      fontSize: '14px'
    },
    input: {
      width: '100%',
      padding: '12px 15px',
      border: '2px solid #e9ecef',
      borderRadius: '8px',
      fontSize: '16px',
      transition: 'all 0.3s ease',
      backgroundColor: '#fff',
      boxSizing: 'border-box' as const
    },
    button: {
      backgroundColor: '#4CAF50',
      color: 'white',
      padding: '12px 24px',
      border: 'none',
      borderRadius: '8px',
      cursor: 'pointer',
      fontSize: '16px',
      fontWeight: '600',
      transition: 'all 0.3s ease',
      marginRight: '10px',
      marginBottom: '10px'
    },
    buttonSecondary: {
      backgroundColor: '#6c757d',
      color: 'white',
      padding: '8px 16px',
      border: 'none',
      borderRadius: '6px',
      cursor: 'pointer',
      fontSize: '14px',
      fontWeight: '500',
      transition: 'all 0.3s ease'
    },
    buttonDanger: {
      backgroundColor: '#dc3545',
      color: 'white',
      padding: '8px 16px',
      border: 'none',
      borderRadius: '6px',
      cursor: 'pointer',
      fontSize: '14px',
      fontWeight: '500',
      transition: 'all 0.3s ease'
    },
    purchaseRow: {
      backgroundColor: '#f8f9fa',
      padding: '20px',
      borderRadius: '10px',
      marginBottom: '20px',
      border: '1px solid #e9ecef'
    },
    purchaseHeader: {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      marginBottom: '15px'
    },
    purchaseTitle: {
      fontSize: '18px',
      fontWeight: '600',
      color: '#2c3e50',
      margin: 0
    },
    status: {
      padding: '10px 15px',
      borderRadius: '8px',
      marginBottom: '20px',
      fontWeight: '500',
      textAlign: 'center' as const
    },
    statusSuccess: {
      backgroundColor: '#d4edda',
      color: '#155724',
      border: '1px solid #c3e6cb'
    },
    statusError: {
      backgroundColor: '#f8d7da',
      color: '#721c24',
      border: '1px solid #f5c6cb'
    },
    statusInfo: {
      backgroundColor: '#d1ecf1',
      color: '#0c5460',
      border: '1px solid #bee5eb'
    }
  };

  return (
    <div style={styles.body}>
      <style>{`
        /* Remove number input arrows/spinners */
        input[type=number]::-webkit-inner-spin-button,
        input[type=number]::-webkit-outer-spin-button {
          -webkit-appearance: none;
          margin: 0;
        }
        input[type=number] {
          -moz-appearance: textfield;
        }
      `}</style>
      <div style={styles.container}>
        <header style={styles.header}>
          <h1 style={styles.title}>Company Credit Card Receipt Form</h1>
        </header>

        <div style={{
          display: 'grid',
          gridTemplateColumns: '1fr 1fr',
          gap: '20px',
          marginBottom: '20px'
        }}>
          <div style={{
            backgroundColor: 'rgba(255, 255, 255, 0.95)',
            padding: '15px',
            borderRadius: '10px',
            boxShadow: '0 4px 16px rgba(0, 0, 0, 0.1)'
          }}>
            <h3 style={{
              color: '#2c3e50',
              fontSize: '16px',
              fontWeight: 'bold',
              marginBottom: '12px',
              borderBottom: '2px solid #8e44ad',
              paddingBottom: '6px'
            }}>
              Account Description
            </h3>
            <div style={{ lineHeight: '1.5', fontSize: '13px' }}>
              <div style={{ marginBottom: '8px' }}>
                <strong style={{ color: '#2c3e50' }}>Office Account:</strong> Office Bonus / Prize
              </div>
              <div style={{ marginBottom: '8px' }}>
                <strong style={{ color: '#2c3e50' }}>Holiday Budget:</strong> Holiday Decorating
              </div>
              <div style={{ marginBottom: '8px' }}>
                <strong style={{ color: '#2c3e50' }}>Office Supplies:</strong> Necessities (Distilled Water / Patient Water / etc.)
              </div>
              <div>
                <strong style={{ color: '#2c3e50' }}>Extra:</strong> Coffe / Food for Doctors. Explain
              </div>
            </div>
          </div>

          <div style={{
            backgroundColor: 'rgba(255, 255, 255, 0.95)',
            padding: '15px',
            borderRadius: '10px',
            boxShadow: '0 4px 16px rgba(0, 0, 0, 0.1)'
          }}>
            <div style={{...styles.formGroup, marginBottom: '12px'}}>
              <label style={{...styles.label, fontSize: '12px'}} htmlFor="office">Office *</label>
              <select
                id="office"
                value={formData.office}
                onChange={(e) => setFormData({...formData, office: e.target.value})}
                style={{...styles.input, padding: '8px 10px', fontSize: '14px'}}
                required
              >
                <option value="">Select Office</option>
                {officeOptions.map(office => (
                  <option key={office} value={office}>{office}</option>
                ))}
              </select>
            </div>
            <div style={{...styles.formGroup, marginBottom: '12px'}}>
              <label style={{...styles.label, fontSize: '12px'}} htmlFor="date">Date *</label>
              <input
                type="date"
                id="date"
                value={formData.date}
                onChange={(e) => setFormData({...formData, date: e.target.value})}
                style={{...styles.input, padding: '8px 10px', fontSize: '14px'}}
                required
              />
            </div>
            <div style={{...styles.formGroup, marginBottom: '12px'}}>
              <label style={{...styles.label, fontSize: '12px'}} htmlFor="name">Employee Name *</label>
              <input
                type="text"
                id="name"
                value={formData.name}
                onChange={(e) => setFormData({...formData, name: e.target.value})}
                style={{...styles.input, padding: '8px 10px', fontSize: '14px'}}
                placeholder="Enter your name"
                required
              />
            </div>
            <div style={styles.formGroup}>
              <label style={{...styles.label, fontSize: '12px'}} htmlFor="cardLastFour">Card Number *</label>
              <div style={{ display: 'flex', alignItems: 'center', gap: '5px', marginBottom: '5px' }}>
                <span style={{ color: '#6c757d', fontSize: '11px' }}>(Last 4 digits of the credit card)</span>
              </div>
              <div style={{ display: 'flex', alignItems: 'center' }}>
                <span style={{
                  padding: '6px 8px',
                  border: '2px solid #e9ecef',
                  borderRadius: '4px 0 0 4px',
                  backgroundColor: '#f8f9fa',
                  color: '#6c757d',
                  fontSize: '12px',
                  fontFamily: 'monospace',
                  whiteSpace: 'nowrap'
                }}>
                  XXXX-XXXX-XXXX-
                </span>
                <input
                  type="text"
                  id="cardLastFour"
                  value={formData.cardLastFour}
                  onChange={(e) => setFormData({...formData, cardLastFour: e.target.value})}
                  style={{
                    ...styles.input,
                    padding: '6px 8px',
                    borderRadius: '0 4px 4px 0',
                    borderLeft: 'none',
                    fontFamily: 'monospace',
                    fontSize: '12px',
                    width: '60px'
                  }}
                  maxLength={4}
                  pattern="[0-9]{4}"
                  placeholder="1234"
                  required
                />
              </div>
            </div>
          </div>
        </div>

        {submitStatus && (
          <div style={{
            ...styles.status,
            ...(submitStatus.includes('✅') ? styles.statusSuccess : 
               submitStatus.includes('❌') ? styles.statusError : styles.statusInfo)
          }}>
            {submitStatus}
          </div>
        )}


        <form style={styles.form} onSubmit={(e) => e.preventDefault()}>

          <div style={{marginTop: '20px'}}>
            {purchases.map((purchase, index) => (
              <div key={index} style={{
                ...styles.purchaseRow,
                opacity: (!formData.office || !formData.date || !formData.name || !formData.cardLastFour) ? 0.5 : 1,
                pointerEvents: (!formData.office || !formData.date || !formData.name || !formData.cardLastFour) ? 'none' : 'auto'
              }}>
                <div style={styles.purchaseHeader}>
                  <h4 style={styles.purchaseTitle}>Purchase {index + 1}</h4>
                  {purchases.length > 1 && (
                    <button
                      type="button"
                      onClick={() => removePurchaseRow(index)}
                      style={styles.buttonDanger}
                      disabled={!formData.office || !formData.date || !formData.name || !formData.cardLastFour}
                    >
                      Remove
                    </button>
                  )}
                </div>
                
                <div style={{
                  display: 'grid',
                  gridTemplateColumns: '1fr 1fr 1fr 1fr 1fr 1fr',
                  gap: '15px',
                  marginBottom: '15px'
                }}>
                  <div style={styles.formGroup}>
                    <label style={styles.label}>Date of Purchase *</label>
                    <input
                      type="date"
                      value={purchase.date}
                      onChange={(e) => updatePurchase(index, 'date', e.target.value)}
                      style={styles.input}
                      required
                    />
                  </div>
                  
                  <div style={styles.formGroup}>
                    <label style={styles.label}>Store/Website *</label>
                    <input
                      type="text"
                      list={`vendor-options-${index}`}
                      value={purchase.vendor}
                      onChange={(e) => updatePurchase(index, 'vendor', e.target.value)}
                      style={styles.input}
                      placeholder="Select or Type"
                      required
                    />
                    <datalist id={`vendor-options-${index}`}>
                      <option value="Amazon">Amazon</option>
                      <option value="Target">Target</option>
                      <option value="Smart and Final">Smart and Final</option>
                      <option value="Sams Club">Sams Club</option>
                      <option value="Costco">Costco</option>
                      <option value="Walmart">Walmart</option>
                      <option value="Home Depot">Home Depot</option>
                      <option value="Family Dollar">Family Dollar</option>
                      <option value="Dollar Tree">Dollar Tree</option>
                      <option value="Dollar General">Dollar General</option>
                      <option value="Post Office">Post Office</option>
                      <option value="Vons">Vons</option>
                      <option value="Bakersfield Rubber Stamps">Bakersfield Rubber Stamps</option>
                      <option value="Scholastic">Scholastic</option>
                      <option value="UPD">UPD</option>
                      <option value="Office Depot">Office Depot</option>
                      <option value="Scent Air">Scent Air</option>
                      <option value="Bill Wright Toyota">Bill Wright Toyota</option>
                      <option value="Loma Linda Spore Tests">Loma Linda Spore Tests</option>
                      <option value="CDA">CDA</option>
                    </datalist>
                  </div>
                  
                  <div style={styles.formGroup}>
                    <label style={styles.label}>Reason for Purchase *</label>
                    <input
                      type="text"
                      list={`reason-options-${index}`}
                      value={purchase.reason}
                      onChange={(e) => updatePurchase(index, 'reason', e.target.value)}
                      style={styles.input}
                      placeholder="Select or Type"
                      required
                    />
                    <datalist id={`reason-options-${index}`}>
                      <option value="Gas">Gas</option>
                      <option value="Toys">Toys</option>
                      <option value="Water">Water</option>
                      <option value="Event">Event</option>
                      <option value="Food for Staff">Food for Staff</option>
                      <option value="Stamps">Stamps</option>
                    </datalist>
                  </div>
                  
                  <div style={styles.formGroup}>
                    <label style={styles.label}>Amount *</label>
                    <input
                      type="number"
                      step="0.01"
                      value={purchase.amount}
                      onChange={(e) => updatePurchase(index, 'amount', e.target.value)}
                      style={styles.input}
                      required
                    />
                  </div>
                  
                  <div style={styles.formGroup}>
                    <label style={styles.label}>Account Description *</label>
                    <select
                      value={purchase.description}
                      onChange={(e) => updatePurchase(index, 'description', e.target.value)}
                      style={styles.input}
                      required
                    >
                      <option value="">Select Account</option>
                      <option value="Office Account">Office Account</option>
                      <option value="Holiday Budget">Holiday Budget</option>
                      <option value="Office Supplies">Office Supplies</option>
                      <option value="Extra">Extra</option>
                    </select>
                  </div>
                  
                  <div style={styles.formGroup}>
                    <label style={styles.label}>Receipt Images *</label>
                    <input
                      type="file"
                      accept=".jpg,.jpeg,.png,image/jpeg,image/png"
                      multiple
                      onChange={(e) => {
                        const files = e.target.files;
                        if (files && files.length > 0) handleFileUpload(index, files);
                      }}
                      style={styles.input}
                    />
                    {purchase.receiptFiles && purchase.receiptFiles.length > 0 && (
                      <div style={{marginTop: '5px'}}>
                        {purchase.receiptFiles.map((file, fileIndex) => {
                          const status = file.uploadStatus ?? 'pending';
                          const statusStyles = {
                            pending: { icon: '📎', color: '#6c757d', label: 'Selected' },
                            uploading: { icon: '⏳', color: '#0c5460', label: 'Uploading' },
                            failed: { icon: '❌', color: '#721c24', label: 'Upload failed' },
                          }[status];

                          return (
                            <p key={fileIndex} style={{ color: statusStyles.color, fontSize: '11px', margin: '2px 0' }}>
                              {statusStyles.icon} {file.name} ({statusStyles.label})
                            </p>
                          );
                        })}
                      </div>
                    )}
                  </div>
                </div>
              </div>
            ))}
            
            <button
              type="button"
              onClick={addPurchaseRow}
              style={{
                ...styles.buttonSecondary,
                opacity: (!formData.office || !formData.date || !formData.name || !formData.cardLastFour) ? 0.5 : 1,
                cursor: (!formData.office || !formData.date || !formData.name || !formData.cardLastFour) ? 'not-allowed' : 'pointer'
              }}
              disabled={!formData.office || !formData.date || !formData.name || !formData.cardLastFour}
            >
              + Add Another Purchase
            </button>
          </div>

          <div style={{marginTop: '30px', textAlign: 'center'}}>
            <button
              type="button"
              onClick={handleSubmit}
              disabled={loading || !isSubmissionComplete()}
              style={{
                ...styles.button,
                fontSize: '18px',
                padding: '15px 30px',
                opacity: loading || !isSubmissionComplete() ? 0.7 : 1,
                cursor: loading || !isSubmissionComplete() ? 'not-allowed' : 'pointer'
              }}
            >
              {loading ? '⏳ Processing...' : '📄 Submit'}
            </button>
          </div>
        </form>
      </div>

    </div>
  );
};

export default CreditCardReceipts;
