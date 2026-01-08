'use client';

import React, { useState, useEffect } from 'react';
import { doc, setDoc, getDoc, deleteDoc } from 'firebase/firestore';
import { ref, uploadBytes, getDownloadURL } from 'firebase/storage';
import { db, storage } from '@/lib/firebase.config';
import { enableAllSecurityMeasures } from '@/lib/security-client';

// Interfaces for type safety
interface ReceiptFile {
  name: string;
  url: string;
}

interface Purchase {
  date: string;
  vendor: string;
  reason: string;
  amount: string;
  description: string;
  receiptFiles: ReceiptFile[];
}

const ReimbursementRequest = () => {
  // State management
  const [formData, setFormData] = useState({
    name: '',
    date: new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' }), // 캘리포니아 시간대 오늘 날짜
    office: '' // Office 선택
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
  const [submissionId, setSubmissionId] = useState('');

  // 🔒 보안 조치 활성화
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

  // Office options
  const officeOptions = ['Bernard', 'California', 'Corporate', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];

  // Generate unique submission ID
  const generateSubmissionId = () => {
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const randomSuffix = Math.random().toString(36).substring(2, 8);
    return `${timestamp}_${randomSuffix}`;
  };

  // Save data to Firestore
  const saveData = async (currentSubmissionId: string) => {
    try {
      // Generate unique document ID with timestamp to prevent overwriting
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const uniqueId = `${formData.name}_${timestamp}`;
      
      const docRef = doc(db, 'reimbursement-requests', uniqueId);
      await setDoc(docRef, {
        name: formData.name,
        cardNumber: '', // Not used for reimbursement requests
        date: formData.date,
        office: formData.office,
        submissionId: currentSubmissionId,
        data: collectFormData(),
        lastUpdated: new Date(),
        createdAt: new Date()
      }, { merge: true });
      
    } catch (error) {
      console.error('Error saving data:', error);
    }
  };

  // Upload file to Firebase Storage
  const uploadFile = async (file: File, fileName: string) => {
    try {
      console.log('Starting file upload:', { fileName, fileSize: file.size, fileType: file.type });
      
      const storageRef = ref(storage, `reimbursement-receipts/${fileName}`);
      const snapshot = await uploadBytes(storageRef, file);
      const downloadURL = await getDownloadURL(snapshot.ref);
      
      console.log('File uploaded successfully:', downloadURL);
      
      return {
        success: true,
        fileName: fileName,
        downloadURL: downloadURL
      };
    } catch (error) {
      console.error('Error uploading file:', error);
      return { success: false, error: error instanceof Error ? error.message : 'Unknown error' };
    }
  };

  // Handle form submission
  const handleSubmit = async () => {
    try {
      setLoading(true);
      setSubmitStatus('📤 Submitting...');

      // Generate submission ID if not already set (in case no files were uploaded)
      let currentSubmissionId = submissionId;
      if (!currentSubmissionId) {
        currentSubmissionId = generateSubmissionId();
        setSubmissionId(currentSubmissionId);
        console.log('🆔 Generated submission ID during submit:', currentSubmissionId);
      }

      // Save data to Firestore (final save)
      await saveData(currentSubmissionId);
      
      // Reset form
      setFormData({ name: '', date: new Date().toLocaleDateString('en-CA', { timeZone: 'America/Los_Angeles' }), office: '' });
      setPurchases([{ date: '', vendor: '', reason: '', amount: '', description: '', receiptFiles: [] }]);
      setSubmissionId(''); // Reset submission ID for next submission

      setSubmitStatus('✅ Submitted successfully! Data saved for manager review.');
      
      setTimeout(() => setSubmitStatus(''), 7000);

    } catch (error) {
      console.error('Submit error:', error);
      setSubmitStatus('❌ Submission failed: ' + (error instanceof Error ? error.message : 'Unknown error'));
      setTimeout(() => setSubmitStatus(''), 7000);
    } finally {
      setLoading(false);
    }
  };

  // Collect form data
  const collectFormData = () => {
    return purchases.map((purchase: Purchase, index: number) => ({
      name: formData.name,
      cardLastFour: '', // Not used for reimbursement requests
      date: purchase.date,
      office: formData.office,
      vendor: purchase.vendor,
      reason: purchase.reason,
      amount: purchase.amount,
      description: '', // Account Description not used for reimbursement requests
      receiptFiles: purchase.receiptFiles.map((file: ReceiptFile) => file.name).join(', '),
      index: index + 1
    }));
  };

  // Add new purchase row
  const addPurchaseRow = () => {
    setPurchases([...purchases, {
      date: '',
      vendor: '',
      reason: '',
      amount: '',
      description: '',
      receiptFiles: []
    }]);
  };

  // Remove purchase row
  const removePurchaseRow = (index: number) => {
    if (purchases.length > 1) {
      const newPurchases = purchases.filter((_: Purchase, i: number) => i !== index);
      setPurchases(newPurchases);
    }
  };

  // Update purchase data
  const updatePurchase = (index: number, field: string, value: any) => {
    const newPurchases = [...purchases];
    newPurchases[index] = { ...newPurchases[index], [field]: value };
    setPurchases(newPurchases);
  };

  // Handle file upload
  const handleFileUpload = async (index: number, files: FileList) => {
    console.log('handleFileUpload called:', { index, fileCount: files.length });
    
    if (!files || files.length === 0) {
      console.log('No files to upload');
      return;
    }

    // Filter out PDF files - only allow image files
    const imageFiles = Array.from(files).filter(file => {
      const isImage = file.type.startsWith('image/');
      if (!isImage) {
        alert(`❌ PDF files are not allowed. Please upload image files only. File "${file.name}" was rejected.`);
      }
      return isImage;
    });

    if (imageFiles.length === 0) {
      alert('❌ No valid image files selected. Please upload image files only (PDF files are not allowed).');
      return;
    }

    // Generate submission ID if not already set - use immediately
    let currentSubmissionId = submissionId;
    if (!currentSubmissionId) {
      currentSubmissionId = generateSubmissionId();
      setSubmissionId(currentSubmissionId);
      console.log('🆔 Generated new submission ID:', currentSubmissionId);
    } else {
      console.log('🆔 Using existing submission ID:', currentSubmissionId);
    }

    console.log('Starting file upload process...');
    const uploadPromises = imageFiles.map(async (file: File, fileIndex: number) => {
      console.log(`Processing file ${fileIndex + 1}:`, { name: file.name, size: file.size, type: file.type });
      
      // Use the current submission ID for all files in this upload
      const sequenceNumber = String(Date.now() + fileIndex).padStart(15, '0');
      const fileName = `${formData.name}_${currentSubmissionId}_purchase${index + 1}_${sequenceNumber}_${file.name}`;
      console.log('📁 Generated fileName:', fileName);
      console.log('📁 Submission ID:', currentSubmissionId);
      console.log('📁 Purchase index:', index + 1);
      console.log('📁 File parts:', fileName.split('_'));
      
      const uploadResult = await uploadFile(file, fileName);
      console.log('Upload result:', uploadResult);
      
      if (uploadResult?.success) {
        console.log('File uploaded successfully:', fileName);
        return { name: fileName, url: uploadResult.downloadURL };
      } else {
        console.error('File upload failed:', uploadResult?.error);
        return null;
      }
    });

    console.log('Waiting for all uploads to complete...');
    const uploadedFiles = (await Promise.all(uploadPromises)).filter((file): file is ReceiptFile => file !== null);
    console.log('Uploaded files:', uploadedFiles);
    
    if (uploadedFiles.length > 0) {
      const currentFiles = purchases[index].receiptFiles || [];
      console.log('Current files:', currentFiles);
      console.log('Replacing files for purchase...');
      // Replace existing files with new ones (don't append)
      updatePurchase(index, 'receiptFiles', uploadedFiles);
      console.log('Files replaced for purchase');
    } else {
      console.log('No files were successfully uploaded');
    }
  };

  const styles = {
    body: {
      fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
      lineHeight: 1.6,
      color: '#333',
      background: 'linear-gradient(135deg, #e0c3fc 0%, #8ec5fc 100%)',
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
    textarea: {
      width: '100%',
      padding: '12px 15px',
      border: '2px solid #e9ecef',
      borderRadius: '8px',
      fontSize: '16px',
      transition: 'all 0.3s ease',
      backgroundColor: '#fff',
      boxSizing: 'border-box' as const,
      minHeight: '80px',
      resize: 'vertical' as const
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
    grid: {
      display: 'grid',
      gridTemplateColumns: 'repeat(auto-fit, minmax(250px, 1fr))',
      gap: '15px'
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
          <h1 style={styles.title}>Reimbursement Request</h1>
        </header>

        {/* Employee Info */}
        <div style={{
          display: 'flex',
          justifyContent: 'center',
          marginBottom: '20px'
        }}>
          <div style={{
            backgroundColor: 'rgba(255, 255, 255, 0.95)',
            padding: '15px',
            borderRadius: '10px',
            boxShadow: '0 4px 16px rgba(0, 0, 0, 0.1)',
            width: '100%',
            maxWidth: '1200px'
          }}>
            <div style={{
              display: 'grid',
              gridTemplateColumns: '1fr 1fr 1fr',
              gap: '15px',
              alignItems: 'end'
            }}>
              <div style={styles.formGroup}>
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
              <div style={styles.formGroup}>
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
              <div style={styles.formGroup}>
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
            </div>
          </div>
        </div>

        {/* Status Messages */}
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
          {/* Purchase Rows */}
          <div style={{marginTop: '20px'}}>
            {purchases.map((purchase, index) => (
              <div key={index} style={{
                ...styles.purchaseRow,
                opacity: (!formData.office || !formData.date || !formData.name) ? 0.5 : 1,
                pointerEvents: (!formData.office || !formData.date || !formData.name) ? 'none' : 'auto'
              }}>
                <div style={styles.purchaseHeader}>
                  <h4 style={styles.purchaseTitle}>Purchase {index + 1}</h4>
                  {purchases.length > 1 && (
                    <button
                      type="button"
                      onClick={() => removePurchaseRow(index)}
                      style={styles.buttonDanger}
                      disabled={!formData.office || !formData.date || !formData.name}
                    >
                      Remove
                    </button>
                  )}
                </div>
                
                <div style={{
                  display: 'grid',
                  gridTemplateColumns: '1fr 1fr 1fr 1fr 1fr',
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
                    <label style={styles.label}>Item *</label>
                    <input
                      type="text"
                      value={purchase.vendor}
                      onChange={(e) => updatePurchase(index, 'vendor', e.target.value)}
                      style={styles.input}
                      required
                    />
                  </div>
                  
                  <div style={styles.formGroup}>
                    <label style={styles.label}>Reason for Purchase *</label>
                    <input
                      type="text"
                      value={purchase.reason}
                      onChange={(e) => updatePurchase(index, 'reason', e.target.value)}
                      style={styles.input}
                      required
                    />
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
                    <label style={styles.label}>Receipt Images</label>
                    <input
                      type="file"
                      accept="image/*"
                      multiple
                      onChange={(e) => {
                        const files = e.target.files;
                        if (files && files.length > 0) handleFileUpload(index, files);
                      }}
                      style={styles.input}
                    />
                    {purchase.receiptFiles && purchase.receiptFiles.length > 0 && (
                      <div style={{marginTop: '5px'}}>
                        {purchase.receiptFiles.map((file, fileIndex) => (
                          <p key={fileIndex} style={{color: '#28a745', fontSize: '11px', margin: '2px 0'}}>
                            ✅ {file.name}
                          </p>
                        ))}
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
                opacity: (!formData.office || !formData.date || !formData.name) ? 0.5 : 1,
                cursor: (!formData.office || !formData.date || !formData.name) ? 'not-allowed' : 'pointer'
              }}
              disabled={!formData.office || !formData.date || !formData.name}
            >
              + Add Another Purchase
            </button>
          </div>

          {/* Submit Button */}
          <div style={{marginTop: '30px', textAlign: 'center'}}>
            <button
              type="button"
              onClick={handleSubmit}
              disabled={loading || !formData.office || !formData.date || !formData.name}
              style={{
                ...styles.button,
                fontSize: '18px',
                padding: '15px 30px',
                opacity: (loading || !formData.office || !formData.date || !formData.name) ? 0.7 : 1,
                cursor: (loading || !formData.office || !formData.date || !formData.name) ? 'not-allowed' : 'pointer'
              }}
            >
              {loading ? '⏳ Processing...' : '📄 Submit Report'}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};

export default ReimbursementRequest;

