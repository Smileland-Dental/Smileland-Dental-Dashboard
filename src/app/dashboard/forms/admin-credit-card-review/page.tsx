'use client';

import React, { useState, useEffect } from 'react';
import { collection, getDocs, deleteDoc, doc, getDoc, setDoc } from 'firebase/firestore';
import { ref, listAll, getDownloadURL, deleteObject } from 'firebase/storage';
import { onAuthStateChanged } from 'firebase/auth';
import { db, storage, auth } from '@/lib/firebase.config';

const sanitizeFirebaseDataClient = (data: any): any => {
  if (!data || typeof data !== 'object') {
    return data;
  }
  
  const sanitized: any = {};
  for (const [key, value] of Object.entries(data)) {
    if (!/^[a-zA-Z0-9_]+$/.test(key)) {
      continue; 
    }
    
    if (value === null || value === undefined) {
      sanitized[key] = value;
    } else if (typeof value === 'string') {
      sanitized[key] = value.substring(0, 10000).replace(/[<>]/g, '');
    } else if (typeof value === 'number') {
      sanitized[key] = isFinite(value) ? value : 0;
    } else if (value instanceof Date) {
      sanitized[key] = value;
    } else if (typeof value === 'boolean') {
      sanitized[key] = value;
    } else if (Array.isArray(value)) {
      sanitized[key] = value.slice(0, 1000);
    } else if (typeof value === 'object') {
      sanitized[key] = sanitizeFirebaseDataClient(value);
    }
  }
  
  return sanitized;
};

const ISO_CALENDAR_DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

function formatDatePacificLosAngeles(date: Date): string {
  const formatter = new Intl.DateTimeFormat('en-US', {
    timeZone: 'America/Los_Angeles',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  });
  const parts = formatter.formatToParts(date);
  const year = parts.find(p => p.type === 'year')?.value || '';
  const month = parts.find(p => p.type === 'month')?.value || '';
  const day = parts.find(p => p.type === 'day')?.value || '';
  return `${year}-${month}-${day}`;
}

function getSubmissionDateStringForFilter(submission: { date: string; submittedAt: Date }): string {
  if (submission.date && ISO_CALENDAR_DATE_RE.test(submission.date)) {
    return submission.date;
  }
  const at = submission.submittedAt instanceof Date ? submission.submittedAt : new Date(submission.submittedAt);
  return formatDatePacificLosAngeles(at);
}

function getSubmissionMonthKey(submission: { date: string; submittedAt: Date }): string {
  return getSubmissionDateStringForFilter(submission).slice(0, 7);
}

function getCurrentMonthKeyPacific(): string {
  return formatDatePacificLosAngeles(new Date()).slice(0, 7);
}

function formatMonthLabel(yyyyMm: string): string {
  const [year, month] = yyyyMm.split('-').map(Number);
  if (!year || !month) return yyyyMm;
  return new Date(year, month - 1, 1).toLocaleDateString('en-US', {
    year: 'numeric',
    month: 'long',
  });
}

function getSubmissionDateDisplayLabel(submission: { date: string; submittedAt: Date }): string {
  if (submission.date?.trim()) {
    return submission.date.trim();
  }
  const at = submission.submittedAt instanceof Date ? submission.submittedAt : new Date(submission.submittedAt);
  return formatDatePacificLosAngeles(at);
}

interface Purchase {
  date: string;
  vendor: string;
  reason: string;
  amount: string;
  description: string;
  receiptFiles: string[] | string;
}

interface Submission {
  id: string;
  employeeName: string;
  cardNumber: string;
  date: string;
  office: string;
  submissionId: string;
  purchases: Purchase[];
  totalAmount: string;
  submittedAt: Date;
  signed?: boolean; 
  addedOnNumbersChecked?: boolean;
  formType?: 'credit-card' | 'reimbursement'; 
  approved?: boolean;
}

interface ReceiptFile {
  name: string;
  url: string;
}

const AdminCreditCardReview = () => {
  const [submissions, setSubmissions] = useState<Submission[]>([]);
  const [selectedSubmission, setSelectedSubmission] = useState<Submission | null>(null);
  const [loading, setLoading] = useState(false);
  const [receiptFiles, setReceiptFiles] = useState<ReceiptFile[]>([]);
  const [loadingFiles, setLoadingFiles] = useState(false);
  const [filterMonth, setFilterMonth] = useState(getCurrentMonthKeyPacific);
  const [filterOffice, setFilterOffice] = useState('');
  const [checkedSubmissionIds, setCheckedSubmissionIds] = useState<Set<string>>(new Set());
  const [amountAdjustedTo, setAmountAdjustedTo] = useState('');
  const [reasonForAdjustment, setReasonForAdjustment] = useState('');
  const [pageReady, setPageReady] = useState(false);

  useEffect(() => {
    let cancelled = false;
    const goHome = () => {
      if (typeof window !== 'undefined') {
        window.location.replace('/');
      }
    };

    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      try {
        if (!currentUser) {
          goHome();
          return;
        }

        const userDoc = await getDoc(doc(db, 'users', currentUser.uid));
        if (!userDoc.exists()) {
          goHome();
          return;
        }

        const userData = userDoc.data();
        if (
          userData?.role !== 'HR' &&
          userData?.role !== 'Director' 
        ) {
          goHome();
          return;
        }

        if (!cancelled) {
          setPageReady(true);
        }
      } catch {
        goHome();
      }
    });

    if (
      process.env.NODE_ENV === 'production' &&
      typeof window !== 'undefined' &&
      window.location.protocol !== 'https:'
    ) {
      window.location.href = window.location.href.replace('http:', 'https:');
    }

    return () => {
      cancelled = true;
      unsubscribe();
    };
  }, []);

  const loadSubmissions = async () => {
    try {
      setLoading(true);
      
      const submissions: Submission[] = [];
      
      let retryCount = 0;
      const maxRetries = 3;
      let creditCardSnapshot: Awaited<ReturnType<typeof getDocs>> | null = null;
      
      while (retryCount < maxRetries && !creditCardSnapshot) {
        try {
          creditCardSnapshot = await getDocs(collection(db, 'credit-card-receipts'));
        } catch (firestoreError) {
          retryCount++;
          
          if (retryCount < maxRetries) {
            await new Promise(resolve => setTimeout(resolve, 1000 * retryCount));
          } else {
            break;
          }
        }
      }
      
      if (creditCardSnapshot) {
        creditCardSnapshot.forEach((doc) => {
          const data: any = doc.data();
          const totalAmount = data.data.reduce((sum: number, item: any) => sum + parseFloat(item.amount || 0), 0);
          
          submissions.push({
            id: doc.id,
            employeeName: data.name,
            cardNumber: data.cardNumber,
            date: data.date || data.data[0]?.date,
            office: data.office || 'N/A',
            submissionId: data.submissionId,
            purchases: data.data.map((item: any) => ({
              date: item.date,
              vendor: item.vendor,
              reason: item.reason,
              amount: item.amount,
              description: item.description,
              receiptFiles: item.receiptFiles && typeof item.receiptFiles === 'string' ? item.receiptFiles.split(', ') : (Array.isArray(item.receiptFiles) ? item.receiptFiles : [])
            })),
            totalAmount: totalAmount.toFixed(2),
            submittedAt: data.date ? new Date(data.date) : (data.createdAt?.toDate() || new Date()),
            signed: data.signed || false,
            addedOnNumbersChecked: data.addedOnNumbersChecked || false,
            formType: 'credit-card',
          });
        });
      }
      
      retryCount = 0;
      let reimbursementSnapshot: Awaited<ReturnType<typeof getDocs>> | null = null;
      
      while (retryCount < maxRetries && !reimbursementSnapshot) {
        try {
          reimbursementSnapshot = await getDocs(collection(db, 'reimbursement-requests'));
        } catch (firestoreError) {
          retryCount++;
          
          if (retryCount < maxRetries) {
            await new Promise(resolve => setTimeout(resolve, 1000 * retryCount));
          } else {
            break;
          }
        }
      }
      
      if (reimbursementSnapshot) {
        reimbursementSnapshot.forEach((doc) => {
          const data: any = doc.data();
          if (!data.data || !Array.isArray(data.data) || data.data.length === 0) {
            return;
          }
          
          const totalAmount = data.data.reduce((sum: number, item: any) => sum + parseFloat(item.amount || 0), 0);
          
          let submittedAtDate: Date;
          if (data.createdAt) {
            if (data.createdAt.toDate && typeof data.createdAt.toDate === 'function') {
              submittedAtDate = data.createdAt.toDate();
            } else if (data.createdAt instanceof Date) {
              submittedAtDate = data.createdAt;
            } else {
              submittedAtDate = new Date(data.createdAt);
            }
          } else {
            submittedAtDate = new Date();
          }
          
          submissions.push({
            id: doc.id,
            employeeName: data.name || 'Unknown',
            cardNumber: data.cardNumber || '',
            date: data.date || data.data[0]?.date || '',
            office: data.office || 'N/A',
            submissionId: data.submissionId || '',
            purchases: data.data.map((item: any) => ({
              date: item.date || '',
              vendor: item.vendor || '',
              reason: item.reason || '',
              amount: item.amount || '0',
              description: item.description || '',
              receiptFiles: item.receiptFiles && typeof item.receiptFiles === 'string' ? item.receiptFiles.split(', ') : (Array.isArray(item.receiptFiles) ? item.receiptFiles : [])
            })),
            totalAmount: totalAmount.toFixed(2),
            submittedAt: submittedAtDate,
            signed: data.signed || false,
            addedOnNumbersChecked: data.addedOnNumbersChecked || false,
            formType: 'reimbursement',
            approved: data.approved !== undefined ? data.approved : undefined,
          });
        });
      }
      
      submissions.sort((a: Submission, b: Submission) => {
        const aChecked = a.addedOnNumbersChecked || false;
        const bChecked = b.addedOnNumbersChecked || false;
        
        if (aChecked !== bChecked) {
          return aChecked ? 1 : -1;
        }
        
        return b.submittedAt.getTime() - a.submittedAt.getTime();
      });
      setSubmissions(submissions);
    } catch (error) {
      alert('Error loading submissions. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  const loadReceiptFiles = async (submission: Submission) => {
    try {
      setLoadingFiles(true);
      const files: ReceiptFile[] = [];
      
      const storagePath = submission.formType === 'reimbursement' ? 'reimbursement-receipts/' : 'receipts/';
      const listRef = ref(storage, storagePath);
      const result = await listAll(listRef);
      
      let submissionPrefix = '';
      
      if (submission.submissionId) {
        const cardNumberPart = submission.cardNumber && submission.cardNumber.trim() !== '' ? `${submission.cardNumber}_` : '';
        submissionPrefix = `${submission.employeeName}_${cardNumberPart}${submission.submissionId}`;
      } else {
        setReceiptFiles([]);
        setLoadingFiles(false);
        return;
      }
      
      for (const itemRef of result.items) {
        const matches = itemRef.name.startsWith(submissionPrefix);
        
        if (matches) {
          try {
            const downloadURL = await getDownloadURL(itemRef);
            files.push({
              name: itemRef.name,
              url: downloadURL
            });
          } catch (urlError) {
          }
        }
      }
      
      const filesByPurchase: Record<number, ReceiptFile[]> = {};
      
      files.forEach((file: ReceiptFile) => {
        const parts = file.name.split('_');
        let purchaseNum = 0;
        
        for (let i = 0; i < parts.length; i++) {
          const purchaseMatch = parts[i].match(/purchase(\d+)/i);
          if (purchaseMatch) {
            const potentialPurchaseNum = parseInt(purchaseMatch[1], 10);
            if (!isNaN(potentialPurchaseNum) && potentialPurchaseNum > 0) {
              purchaseNum = potentialPurchaseNum;
              break;
            }
          }
        }
        
        if (!filesByPurchase[purchaseNum]) {
          filesByPurchase[purchaseNum] = [];
        }
        filesByPurchase[purchaseNum].push(file);
      });
      
      Object.keys(filesByPurchase).forEach((purchaseNum: string) => {
        const purchaseNumInt = parseInt(purchaseNum, 10);
        filesByPurchase[purchaseNumInt].sort((a: ReceiptFile, b: ReceiptFile) => {
          const getSequenceNumber = (filename: string): number => {
            const parts = filename.split('_');
            for (let i = 0; i < parts.length; i++) {
              if (/^\d{10,}$/.test(parts[i])) {
                return parseInt(parts[i], 10);
              }
            }
            return 0;
          };
          
          const aSeq = getSequenceNumber(a.name);
          const bSeq = getSequenceNumber(b.name);
          return aSeq - bSeq;
        });
      });
      
      const sortedFiles: ReceiptFile[] = [];
      const purchaseNumbers = Object.keys(filesByPurchase).map(Number).sort((a: number, b: number) => a - b);
      
      purchaseNumbers.forEach((purchaseNum: number) => {
        sortedFiles.push(...filesByPurchase[purchaseNum]);
      });
      
      setReceiptFiles(sortedFiles);
    } catch (error) {
      setReceiptFiles([]);
    } finally {
      setLoadingFiles(false);
    }
  };

  useEffect(() => {
    if (!pageReady) return;
    loadSubmissions();
  }, [pageReady]);

  const getAvailableMonths = (): string[] => {
    const months = new Set<string>();
    submissions.forEach((submission: Submission) => {
      const monthKey = getSubmissionMonthKey(submission);
      if (monthKey && /^\d{4}-\d{2}$/.test(monthKey)) {
        months.add(monthKey);
      }
    });
    months.add(getCurrentMonthKeyPacific());
    return Array.from(months).sort((a, b) => b.localeCompare(a));
  };

  const getFilteredSubmissions = () => {
    let filtered = submissions;

    if (filterMonth) {
      filtered = filtered.filter(
        (submission: Submission) => getSubmissionMonthKey(submission) === filterMonth
      );
    }

    if (filterOffice) {
      filtered = filtered.filter((submission: Submission) => 
        submission.office === filterOffice
      );
    }

    return filtered;
  };

  const getUniqueOffices = (): string[] => {
    const offices = [...new Set(submissions.map((s: Submission) => s.office))];
    return offices.filter((office): office is string => Boolean(office) && office !== 'N/A');
  };

  const clearFilters = () => {
    setFilterMonth(getCurrentMonthKeyPacific());
    setFilterOffice('');
  };

  const toggleSubmissionChecked = (submissionId: string) => {
    setCheckedSubmissionIds((prev) => {
      const next = new Set(prev);
      if (next.has(submissionId)) {
        next.delete(submissionId);
      } else {
        next.add(submissionId);
      }
      return next;
    });
  };

  const toggleSelectAllFiltered = () => {
    const filtered = getFilteredSubmissions();
    const allSelected =
      filtered.length > 0 && filtered.every((submission) => checkedSubmissionIds.has(submission.id));

    setCheckedSubmissionIds((prev) => {
      const next = new Set(prev);
      if (allSelected) {
        filtered.forEach((submission) => next.delete(submission.id));
      } else {
        filtered.forEach((submission) => next.add(submission.id));
      }
      return next;
    });
  };

  const removeSubmissionFilesAndDoc = async (submission: Submission) => {
    const collectionName = submission.formType === 'reimbursement' ? 'reimbursement-requests' : 'credit-card-receipts';
    await deleteDoc(doc(db, collectionName, submission.id));

    const receiptStoragePath = submission.formType === 'reimbursement' ? 'reimbursement-receipts/' : 'receipts/';

    for (const purchase of submission.purchases) {
      if (purchase.receiptFiles) {
        let receiptFiles: string[] = [];

        if (typeof purchase.receiptFiles === 'string') {
          receiptFiles = purchase.receiptFiles.split(', ');
        } else if (Array.isArray(purchase.receiptFiles)) {
          receiptFiles = purchase.receiptFiles;
        }

        for (const fileName of receiptFiles) {
          if (fileName && typeof fileName === 'string' && fileName.trim()) {
            try {
              const receiptRef = ref(storage, `${receiptStoragePath}${fileName.trim()}`);
              await deleteObject(receiptRef);
            } catch (error) {
            }
          }
        }
      }
    }
  };

  const checkAddedOnNumbers = async () => {
    if (!selectedSubmission) {
      alert('No submission selected');
      return;
    }
    
    try {
      setLoading(true);
      
      const collectionName = selectedSubmission.formType === 'reimbursement' ? 'reimbursement-requests' : 'credit-card-receipts';
      const docRef = doc(db, collectionName, selectedSubmission.id);
      
      const updateData = sanitizeFirebaseDataClient({
        addedOnNumbersChecked: true,
        addedOnNumbersCheckedAt: new Date()
      });
      await setDoc(docRef, updateData, { merge: true });
      
      setSelectedSubmission({
        ...selectedSubmission,
        addedOnNumbersChecked: true,
      });
      
      setSubmissions(prev => prev.map((sub: Submission) => 
        sub.id === selectedSubmission.id 
          ? { ...sub, addedOnNumbersChecked: true }
          : sub
      ));
      
      alert('✅ Added on Numbers checked successfully!');
      
    } catch (error) {
      alert('Error updating status. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  const managerNotApprove = async () => {
    if (!selectedSubmission) {
      alert('No submission selected');
      return;
    }
    
    if (selectedSubmission.formType !== 'reimbursement') {
      alert('Not Approved option is only available for this.');
      return;
    }
    
    try {
      setLoading(true);
      
      const collectionName = 'reimbursement-requests';
      const docRef = doc(db, collectionName, selectedSubmission.id);
      
      let retryCount = 0;
      const maxRetries = 3;
      let firestoreSuccess = false;
      
      while (retryCount < maxRetries && !firestoreSuccess) {
        try {
          const updateData: any = {
            approved: false, 
            rejectionDate: new Date(),
            signed: true 
          };
          
          if (amountAdjustedTo) {
            const sanitizedAmount = parseFloat(amountAdjustedTo);
            if (!isNaN(sanitizedAmount) && sanitizedAmount >= 0 && sanitizedAmount <= 1000000) {
              updateData.amountAdjustedTo = sanitizedAmount.toFixed(2);
            }
          }
          if (reasonForAdjustment) {
            updateData.reasonForAdjustment = reasonForAdjustment.substring(0, 1000).replace(/[<>]/g, '');
          }
          
          const sanitizedData = sanitizeFirebaseDataClient(updateData);
          await setDoc(docRef, sanitizedData, { merge: true });
          firestoreSuccess = true;
        } catch (firestoreError) {
          retryCount++;
          
          if (retryCount < maxRetries) {
            await new Promise(resolve => setTimeout(resolve, 1000 * retryCount));
          }
        }
      }

      if (!firestoreSuccess) {
        alert('Error processing Not Approved. Please try again.');
        return;
      }
      
      alert('❌ Not approved!');
      
      setSelectedSubmission(null);
      setReceiptFiles([]);
      setAmountAdjustedTo('');
      setReasonForAdjustment('');
      loadSubmissions();
      
    } catch (error) {
      alert('Error processing Not Approved: ' + (error instanceof Error ? error.message : 'Unknown error'));
    } finally {
      setLoading(false);
    }
  };

  const managerApprove = async () => {
    if (!selectedSubmission) {
      alert('No submission selected');
      return;
    }
    
    try {
      setLoading(true);
      
      const collectionName = selectedSubmission.formType === 'reimbursement' ? 'reimbursement-requests' : 'credit-card-receipts';
      const docRef = doc(db, collectionName, selectedSubmission.id);
      
      let retryCount = 0;
      const maxRetries = 3;
      let firestoreSuccess = false;
      
      while (retryCount < maxRetries && !firestoreSuccess) {
        try {
          const updateData: any = {
            signed: true 
          };
          
          if (selectedSubmission.formType === 'reimbursement') {
            updateData.approved = true;
            if (amountAdjustedTo) {
              const sanitizedAmount = parseFloat(amountAdjustedTo);
              if (!isNaN(sanitizedAmount) && sanitizedAmount >= 0 && sanitizedAmount <= 1000000) {
                updateData.amountAdjustedTo = sanitizedAmount.toFixed(2);
              }
            }
            if (reasonForAdjustment) {
              updateData.reasonForAdjustment = reasonForAdjustment.substring(0, 1000).replace(/[<>]/g, '');
            }
          }
          
          const sanitizedData = sanitizeFirebaseDataClient(updateData);
          await setDoc(docRef, sanitizedData, { merge: true });
          firestoreSuccess = true;
        } catch (firestoreError) {
          retryCount++;
          
          if (retryCount < maxRetries) {
            await new Promise(resolve => setTimeout(resolve, 1000 * retryCount));
          }
        }
      }

      if (!firestoreSuccess) {
        alert('Error approving submission. Please try again.');
        return;
      }
      
      alert('✅ Approved!');
      
      setSelectedSubmission(null);
      setReceiptFiles([]);
      setAmountAdjustedTo('');
      setReasonForAdjustment('');
      loadSubmissions();
      
    } catch (error) {
      alert('Error approving submission. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  const resetAllData = async () => {
    if (checkedSubmissionIds.size === 0) {
      alert('Please select at least one submission to delete.');
      return;
    }

    const selectedCount = checkedSubmissionIds.size;
    if (!confirm(`⚠️ This will delete ${selectedCount} selected submission(s). 
      Are you sure you want to continue?`)) {
      return;
    }

    try {
      setLoading(true);

      const toDelete = submissions.filter((submission) => checkedSubmissionIds.has(submission.id));
      if (toDelete.length === 0) {
        alert('No selected submissions found to delete.');
        setLoading(false);
        return;
      }

      for (const submission of toDelete) {
        await removeSubmissionFilesAndDoc(submission);
      }

      if (selectedSubmission && checkedSubmissionIds.has(selectedSubmission.id)) {
        setSelectedSubmission(null);
        setReceiptFiles([]);
      }

      setCheckedSubmissionIds(new Set());
      await loadSubmissions();

      alert(`✅ Deleted ${toDelete.length} selected submission(s).`);
    } catch (error) {
      alert('Error deleting selected submissions. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  const styles = {
    body: {
      fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
      lineHeight: 1.6,
      color: '#333',
      background: 'linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%)',
      minHeight: '100vh',
      margin: 0,
      padding: 0
    },
    container: {
      maxWidth: '1600px',
      margin: '0 auto',
      padding: '20px',
      minHeight: '100vh'
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
    submissionsGrid: {
      display: 'grid',
      gridTemplateColumns: 'repeat(auto-fill, minmax(300px, 1fr))',
      gap: '20px',
      marginBottom: '30px'
    },
    submissionCard: {
      backgroundColor: 'rgba(255, 255, 255, 0.95)',
      padding: '20px',
      borderRadius: '15px',
      boxShadow: '0 8px 32px rgba(0, 0, 0, 0.1)',
      cursor: 'pointer',
      transition: 'all 0.3s ease',
      border: '2px solid transparent'
    },
    
    submissionCardSigned: {
      backgroundColor: 'rgba(173, 216, 230, 0.95)',
      padding: '20px',
      borderRadius: '15px',
      boxShadow: '0 8px 32px rgba(0, 0, 0, 0.1)',
      cursor: 'pointer',
      transition: 'all 0.3s ease',
      border: '2px solid #87CEEB'
    },
    submissionHeader: {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      marginBottom: '15px'
    },
    employeeName: {
      fontSize: '18px',
      fontWeight: 'bold',
      color: '#2c3e50'
    },
    submissionDate: {
      fontSize: '14px',
      color: '#666'
    },
    submissionDetails: {
      fontSize: '14px',
      color: '#555',
      lineHeight: '1.4'
    },
    modal: {
      position: 'fixed' as const,
      top: 0,
      left: 0,
      width: '100%',
      height: '100%',
      backgroundColor: 'rgba(0, 0, 0, 0.8)',
      display: 'flex',
      justifyContent: 'center',
      alignItems: 'center',
      zIndex: 1000
    },
    modalContent: {
      backgroundColor: 'white',
      borderRadius: '15px',
      padding: '30px',
      maxWidth: '90%',
      maxHeight: '90%',
      overflow: 'auto',
      boxShadow: '0 20px 60px rgba(0, 0, 0, 0.3)'
    },
    modalHeader: {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      marginBottom: '20px',
      paddingBottom: '15px',
      borderBottom: '2px solid #e9ecef'
    },
    modalTitle: {
      fontSize: '24px',
      fontWeight: 'bold',
      color: '#2c3e50'
    },
    closeButton: {
      backgroundColor: '#dc3545',
      color: 'white',
      border: 'none',
      borderRadius: '8px',
      padding: '8px 16px',
      cursor: 'pointer',
      fontSize: '16px'
    },
    purchaseTable: {
      width: '100%',
      borderCollapse: 'collapse' as const,
      marginBottom: '20px'
    },
    tableHeader: {
      backgroundColor: '#f8f9fa',
      padding: '12px',
      textAlign: 'left' as const,
      fontWeight: 'bold',
      border: '1px solid #dee2e6'
    },
    tableCell: {
      padding: '12px',
      border: '1px solid #dee2e6',
      fontSize: '14px'
    },
    signatureSection: {
      marginTop: '30px',
      padding: '20px',
      backgroundColor: '#f8f9fa',
      borderRadius: '10px',
      border: '2px solid #e9ecef'
    },
    signatureButtons: {
      marginTop: '15px',
      display: 'flex',
      gap: '10px'
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
      transition: 'all 0.3s ease'
    },
    loading: {
      textAlign: 'center' as const,
      padding: '50px',
      fontSize: '18px',
      color: '#666'
    }
  };

  if (!pageReady) {
    return (
      <div
        style={{
          minHeight: '100vh',
          background: 'linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%)',
        }}
      />
    );
  }

  return (
    <div style={styles.body}>
      <div style={styles.container}>
        <header style={styles.header}>
          <h1 style={styles.title}>Credit Card Receipts</h1>
          
          <div style={{
            display: 'flex',
            gap: '15px',
            justifyContent: 'center',
            marginTop: '20px',
            padding: '15px',
            backgroundColor: 'rgba(248, 249, 250, 0.8)',
            borderRadius: '10px',
            border: '1px solid #dee2e6'
          }}>
            <div style={{display: 'flex', flexDirection: 'column', gap: '5px'}}>
              <label style={{fontSize: '12px', fontWeight: 'bold', color: '#495057'}}>Submitted Month:</label>
              <select
                value={filterMonth}
                onChange={(e) => setFilterMonth(e.target.value)}
                style={{
                  padding: '8px 12px',
                  border: '1px solid #ced4da',
                  borderRadius: '4px',
                  fontSize: '14px',
                  minWidth: '160px'
                }}
              >
                {getAvailableMonths().map((month: string) => (
                  <option key={month} value={month}>{formatMonthLabel(month)}</option>
                ))}
              </select>
            </div>
            
            <div style={{display: 'flex', flexDirection: 'column', gap: '5px'}}>
              <label style={{fontSize: '12px', fontWeight: 'bold', color: '#495057'}}>Filter by Office:</label>
              <select
                value={filterOffice}
                onChange={(e) => setFilterOffice(e.target.value)}
                style={{
                  padding: '8px 12px',
                  border: '1px solid #ced4da',
                  borderRadius: '4px',
                  fontSize: '14px',
                  minWidth: '120px'
                }}
              >
                <option value="">All Offices</option>
                {getUniqueOffices().map((office: string) => (
                  <option key={office} value={office}>{office}</option>
                ))}
              </select>
            </div>
            
            <div style={{display: 'flex', flexDirection: 'column', gap: '5px'}}>
              <label style={{fontSize: '12px', fontWeight: 'bold', color: '#495057'}}>Actions:</label>
              <button
                style={{
                  padding: '8px 12px',
                  backgroundColor: '#6c757d',
                  color: 'white',
                  border: 'none',
                  borderRadius: '4px',
                  fontSize: '12px',
                  cursor: 'pointer'
                }}
                onClick={clearFilters}
              >
                Clear Filters
              </button>
            </div>
          </div>
          
          <div style={{
            marginTop: '10px',
            padding: '8px 15px',
            backgroundColor: 'rgba(23, 162, 184, 0.1)',
            borderRadius: '6px',
            border: '1px solid #17a2b8',
            fontSize: '14px',
            color: '#0c5460'
          }}>
            🔍 Showing {getFilteredSubmissions().length} of {submissions.length} submissions
            {filterMonth && ` • Month: ${formatMonthLabel(filterMonth)}`}
            {filterOffice && ` • Office: ${filterOffice}`}
          </div>
          
          <div style={{display: 'flex', gap: '10px', justifyContent: 'center', marginTop: '15px', flexWrap: 'wrap'}}>
            <button
              style={{
                ...styles.button,
                backgroundColor: '#6c757d',
                fontSize: '14px',
                padding: '10px 20px'
              }}
              onClick={toggleSelectAllFiltered}
              disabled={loading || getFilteredSubmissions().length === 0}
            >
              {getFilteredSubmissions().length > 0 &&
              getFilteredSubmissions().every((submission) => checkedSubmissionIds.has(submission.id))
                ? 'Deselect All'
                : 'Select All'}
            </button>
            <button
              style={{
                ...styles.button,
                backgroundColor: '#dc3545',
                fontSize: '14px',
                padding: '10px 20px',
                opacity: checkedSubmissionIds.size === 0 || loading ? 0.6 : 1
              }}
              onClick={resetAllData}
              disabled={loading || checkedSubmissionIds.size === 0}
            >
              Delete{checkedSubmissionIds.size > 0 ? ` (${checkedSubmissionIds.size})` : ''}
            </button>
          </div>
        </header>

        {loading && !selectedSubmission ? (
          <div style={styles.loading}>Loading submissions...</div>
        ) : getFilteredSubmissions().length === 0 ? (
          <div style={{
            textAlign: 'center',
            padding: '50px',
            backgroundColor: 'rgba(255, 255, 255, 0.95)',
            borderRadius: '15px',
            boxShadow: '0 8px 32px rgba(0, 0, 0, 0.1)'
          }}>
            <h3 style={{color: '#6c757d', marginBottom: '10px'}}>No submissions found</h3>
            <p style={{color: '#666', marginBottom: '20px'}}>
              {submissions.length === 0
                ? 'No submissions have been submitted yet.'
                : 'No submissions match your current filters. Try another month or office.'}
            </p>
            {submissions.length > 0 && (
              <button
                style={{
                  ...styles.button,
                  backgroundColor: '#6c757d',
                  fontSize: '14px',
                  padding: '10px 20px'
                }}
                onClick={clearFilters}
              >
                Clear Filters
              </button>
            )}
          </div>
        ) : (
          <div style={styles.submissionsGrid}>
            {getFilteredSubmissions().map((submission: Submission) => (
              <div
                key={submission.id}
                style={{
                  ...(submission.signed ? styles.submissionCardSigned : styles.submissionCard),
                  position: 'relative',
                  outline: checkedSubmissionIds.has(submission.id) ? '2px solid #dc3545' : undefined,
                }}
                onClick={() => {
                  setSelectedSubmission(submission);
                  loadReceiptFiles(submission);
                }}
              >
                <label
                  style={{
                    position: 'absolute',
                    top: '12px',
                    left: '12px',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '6px',
                    zIndex: 2,
                    cursor: 'pointer',
                    backgroundColor: 'rgba(255, 255, 255, 0.9)',
                    padding: '4px 8px',
                    borderRadius: '6px',
                    border: '1px solid #ced4da',
                    fontSize: '12px',
                    fontWeight: 600,
                    color: '#495057',
                  }}
                  onClick={(e) => e.stopPropagation()}
                >
                  <input
                    type="checkbox"
                    checked={checkedSubmissionIds.has(submission.id)}
                    onChange={() => toggleSubmissionChecked(submission.id)}
                    onClick={(e) => e.stopPropagation()}
                    style={{ width: '16px', height: '16px', cursor: 'pointer' }}
                  />
                  Select
                </label>
                <div style={{ ...styles.submissionHeader, marginTop: '28px' }}>
                  <div style={styles.employeeName}>
                    {submission.employeeName}
                    {submission.formType === 'reimbursement' && (
                      <span style={{
                        marginLeft: '12px',
                        fontSize: '11px',
                        fontWeight: 'bold',
                        padding: '4px 10px',
                        borderRadius: '12px',
                        backgroundColor: '#8e44ad',
                        color: '#ffffff',
                        textTransform: 'uppercase',
                        letterSpacing: '0.5px',
                        display: 'inline-block',
                        verticalAlign: 'middle'
                      }}>
                        Reimbursement Request
                      </span>
                    )}
                  </div>
                  <div style={styles.submissionDate}>
                    Submitted: {getSubmissionDateDisplayLabel(submission)}
                  </div>
                </div>
                <div style={styles.submissionDetails}>
                  <p><strong>Office:</strong> {submission.office}</p>
                  {submission.formType !== 'reimbursement' && (
                    <p><strong>Card:</strong> ****{submission.cardNumber}</p>
                  )}
                  <p><strong>Purchases:</strong> {submission.purchases.length} items</p>
                  <p><strong>Total Amount:</strong> ${submission.totalAmount}</p>
                  
                  <div style={{
                    marginTop: '8px',
                    padding: '8px 12px',
                    backgroundColor: '#f8f9fa',
                    borderRadius: '6px',
                    border: '1px solid #dee2e6',
                    fontSize: '12px'
                  }}>
                    <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'center'}}>
                      <span style={{color: '#6c757d'}}>
                       Added on Numbers:
                      </span>
                      {submission.addedOnNumbersChecked ? (
                        <span style={{color: '#28a745', fontWeight: 'bold'}}>
                          ✓ Checked
                        </span>
                      ) : (
                        <span style={{color: '#6c757d'}}>
                          ✗ X
                        </span>
                      )}
                    </div>
                  </div>
                  
                  {submission.signed && (
                    <p style={{color: '#0066cc', fontWeight: 'bold', marginTop: '8px'}}>
                      {submission.formType === 'reimbursement' && submission.approved === false 
                        ? '❌ Not Approved' 
                        : '✅ Approved'}
                    </p>
                  )}
                </div>
              </div>
            ))}
          </div>
        )}

        {selectedSubmission && (
          <div style={styles.modal}>
            <div style={styles.modalContent}>
              <div style={styles.modalHeader}>
                <h2 style={styles.modalTitle}>
                  Review Submission - {selectedSubmission.employeeName}
                </h2>
                <button
                  style={styles.closeButton}
                  onClick={() => setSelectedSubmission(null)}
                >
                  ✕ Close
                </button>
              </div>

              <div>
                <h3>Employee Information</h3>
                <p><strong>Name:</strong> {selectedSubmission.employeeName}</p>
                <p><strong>Office:</strong> {selectedSubmission.office}</p>
                {selectedSubmission.formType !== 'reimbursement' && (
                  <p><strong>Card Number:</strong> {selectedSubmission.cardNumber}</p>
                )}
                <p><strong>Submission Date:</strong> {getSubmissionDateDisplayLabel(selectedSubmission)}</p>

                <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '20px'}}>
                  <h3 style={{margin: 0}}>Purchase Details</h3>
                  {!selectedSubmission?.addedOnNumbersChecked && (
                    <button
                      style={{
                        backgroundColor: '#17a2b8',
                        color: 'white',
                        padding: '6px 12px',
                        border: 'none',
                        borderRadius: '4px',
                        cursor: 'pointer',
                        fontSize: '12px',
                        fontWeight: '500'
                      }}
                      onClick={checkAddedOnNumbers}
                      disabled={loading}
                    >
                      {loading ? 'Checking...' : '✓ Check Added on Numbers'}
                    </button>
                  )}
                  {selectedSubmission.addedOnNumbersChecked && (
                    <span style={{
                      color: '#28a745',
                      fontWeight: 'bold',
                      fontSize: '12px'
                    }}>
                      ✓ Numbers Checked
                    </span>
                  )}
                </div>
                <table style={styles.purchaseTable}>
                  <thead>
                    <tr>
                      <th style={styles.tableHeader}>Purchase #</th>
                      <th style={styles.tableHeader}>Date</th>
                      {selectedSubmission.formType !== 'reimbursement' && (
                        <th style={styles.tableHeader}>Store/Website</th>
                      )}
                      {selectedSubmission.formType !== 'credit-card' && (
                        <th style={styles.tableHeader}>Item</th>
                      )}
                      <th style={styles.tableHeader}>Reason</th>
                      <th style={styles.tableHeader}>Amount</th>
                      {selectedSubmission.formType !== 'reimbursement' && (
                        <th style={styles.tableHeader}>Account</th>
                      )}
                    </tr>
                  </thead>
                  <tbody>
                    {selectedSubmission.purchases.map((purchase: Purchase, index: number) => (
                      <tr key={index}>
                        <td style={{
                          ...styles.tableCell,
                          backgroundColor: '#f8f9fa',
                          fontWeight: 'bold',
                          textAlign: 'center',
                          color: '#17a2b8'
                        }}>
                          #{index + 1}
                        </td>
                        <td style={styles.tableCell}>{purchase.date}</td>
                        <td style={styles.tableCell}>{purchase.vendor}</td>
                        <td style={styles.tableCell}>{purchase.reason}</td>
                        <td style={styles.tableCell}>${purchase.amount}</td>
                        {selectedSubmission.formType !== 'reimbursement' && (
                          <td style={styles.tableCell}>{purchase.description}</td>
                        )}
                      </tr>
                    ))}
                  </tbody>
                </table>

                <div style={{marginTop: '30px', padding: '20px', backgroundColor: '#f8f9fa', borderRadius: '10px', border: '2px solid #e9ecef'}}>
                  <h3 style={{marginBottom: '15px', color: '#2c3e50'}}>📎 Receipt Files</h3>
                  {loadingFiles ? (
                    <p>Loading files...</p>
                  ) : receiptFiles.length > 0 ? (
                    <div style={{display: 'flex', flexDirection: 'column', gap: '25px'}}>
                      {(() => {
                        const filesByPurchase: Record<number, ReceiptFile[]> = {};
                        receiptFiles.forEach((file: ReceiptFile) => {
                          const parts = file.name.split('_');
                          let purchaseNum = 0;
                          
                          for (let i = 0; i < parts.length; i++) {
                            const purchaseMatch = parts[i].match(/purchase(\d+)/i);
                            if (purchaseMatch) {
                              const potentialPurchaseNum = parseInt(purchaseMatch[1], 10);
                              if (!isNaN(potentialPurchaseNum) && potentialPurchaseNum > 0) {
                                purchaseNum = potentialPurchaseNum;
                                break;
                              }
                            }
                          }
                          
                          if (!filesByPurchase[purchaseNum]) {
                            filesByPurchase[purchaseNum] = [];
                          }
                          filesByPurchase[purchaseNum].push(file);
                        });
                        
                        const purchaseNumbers = Object.keys(filesByPurchase).map(Number).sort((a: number, b: number) => a - b);
                        
                        return purchaseNumbers.map((purchaseNum: number) => (
                          <div key={purchaseNum} style={{
                            border: '2px solid #17a2b8',
                            borderRadius: '10px',
                            padding: '20px',
                            backgroundColor: 'white',
                            boxShadow: '0 2px 8px rgba(0,0,0,0.1)'
                          }}>
                            <h4 style={{
                              margin: '0 0 15px 0',
                              color: '#17a2b8',
                              fontSize: '18px',
                              fontWeight: 'bold',
                              borderBottom: '2px solid #17a2b8',
                              paddingBottom: '8px'
                            }}>
                              🛒 Purchase {purchaseNum}
                            </h4>
                            <div style={{display: 'flex', flexDirection: 'column', gap: '15px'}}>
                              {filesByPurchase[purchaseNum].map((file: ReceiptFile, index: number) => (
                                <div key={index} style={{
                                  border: '1px solid #dee2e6',
                                  borderRadius: '8px',
                                  padding: '15px',
                                  backgroundColor: '#f8f9fa',
                                  textAlign: 'center'
                                }}>
                                  {file.name.toLowerCase().includes('.png') || 
                                   file.name.toLowerCase().includes('.jpg') || 
                                   file.name.toLowerCase().includes('.jpeg') || 
                                   file.name.toLowerCase().includes('.gif') ? (
                                    <img 
                                      src={file.url} 
                                      alt="Receipt"
                                      style={{
                                        maxWidth: '100%',
                                        height: 'auto',
                                        borderRadius: '4px',
                                        border: '1px solid #dee2e6',
                                        boxShadow: '0 2px 4px rgba(0,0,0,0.1)'
                                      }}
                                      onError={(e: React.SyntheticEvent<HTMLImageElement>) => {
                                        const target = e.target as HTMLImageElement;
                                        target.style.display = 'none';
                                        const nextSibling = target.nextSibling as HTMLElement | null;
                                        if (nextSibling) {
                                          nextSibling.style.display = 'block';
                                        }
                                      }}
                                    />
                                  ) : file.name.toLowerCase().includes('.pdf') ? (
                                    <div style={{
                                      padding: '40px',
                                      backgroundColor: '#f8f9fa',
                                      borderRadius: '4px',
                                      border: '1px solid #dee2e6',
                                      marginBottom: '10px'
                                    }}>
                                      <div style={{fontSize: '48px', marginBottom: '10px'}}>📄</div>
                                      <p style={{margin: 0, color: '#666'}}>PDF Document</p>
                                    </div>
                                  ) : (
                                    <div style={{
                                      padding: '40px',
                                      backgroundColor: '#f8f9fa',
                                      borderRadius: '4px',
                                      border: '1px solid #dee2e6',
                                      marginBottom: '10px'
                                    }}>
                                      <div style={{fontSize: '48px', marginBottom: '10px'}}>📎</div>
                                      <p style={{margin: 0, color: '#666'}}>File Attachment</p>
                                    </div>
                                  )}
                                  <div style={{display: 'none', padding: '40px', backgroundColor: '#f8f9fa', borderRadius: '4px', border: '1px solid #dee2e6'}}>
                                    📄 File
                                  </div>
                                  <a 
                                    href={file.url} 
                                    target="_blank" 
                                    rel="noopener noreferrer"
                                    style={{
                                      display: 'inline-block',
                                      backgroundColor: '#007bff',
                                      color: 'white',
                                      padding: '8px 16px',
                                      borderRadius: '4px',
                                      textDecoration: 'none',
                                      fontSize: '14px',
                                      marginTop: '10px'
                                    }}
                                  >
                                    View Full Size
                                  </a>
                                </div>
                              ))}
                            </div>
                          </div>
                        ));
                      })()}
                    </div>
                  ) : (
                    <p style={{color: '#666', fontStyle: 'italic'}}>No receipt files found for this submission.</p>
                  )}
                </div>


                <div style={styles.signatureSection}>
                  {!selectedSubmission?.signed && selectedSubmission.formType === 'reimbursement' && (
                    <div style={{
                      marginBottom: '20px',
                      padding: '20px',
                      backgroundColor: '#f8f9fa',
                      borderRadius: '8px',
                      border: '1px solid #dee2e6'
                    }}>
                      <div style={{marginBottom: '15px'}}>
                        <label style={{
                          display: 'block',
                          marginBottom: '5px',
                          fontWeight: '600',
                          color: '#2c3e50',
                          fontSize: '14px'
                        }}>
                          Amount adjusted to (optional):
                        </label>
                        <input
                          type="number"
                          step="0.01"
                          value={amountAdjustedTo}
                          onChange={(e) => setAmountAdjustedTo(e.target.value)}
                          style={{
                            width: '100%',
                            padding: '10px',
                            border: '1px solid #ced4da',
                            borderRadius: '6px',
                            fontSize: '14px',
                            boxSizing: 'border-box'
                          }}
                        />
                      </div>
                      <div>
                        <label style={{
                          display: 'block',
                          marginBottom: '5px',
                          fontWeight: '600',
                          color: '#2c3e50',
                          fontSize: '14px'
                        }}>
                          Reason for Adjustment or Non-Approval (optional):
                        </label>
                        <textarea
                          value={reasonForAdjustment}
                          onChange={(e) => setReasonForAdjustment(e.target.value)}
                          rows={3}
                          style={{
                            width: '100%',
                            padding: '10px',
                            border: '1px solid #ced4da',
                            borderRadius: '6px',
                            fontSize: '14px',
                            resize: 'vertical',
                            boxSizing: 'border-box',
                            fontFamily: 'inherit'
                          }}
                        />
                      </div>
                    </div>
                  )}
                  
                  <div style={styles.signatureButtons}>
                    {selectedSubmission.signed ? (
                      selectedSubmission.formType === 'reimbursement' && selectedSubmission.approved === false ? (
                        <div style={{
                          padding: '15px',
                          backgroundColor: '#f8d7da',
                          borderRadius: '8px',
                          border: '2px solid #dc3545',
                          textAlign: 'center'
                        }}>
                          <h4 style={{margin: '0 0 10px 0', color: '#721c24'}}>❌ Not Approved</h4>
                        </div>
                      ) : (
                        <div style={{
                          padding: '15px',
                          backgroundColor: '#d4edda',
                          borderRadius: '8px',
                          border: '2px solid #28a745',
                          textAlign: 'center'
                        }}>
                          <h4 style={{margin: '0', color: '#155724', fontSize: '16px', fontWeight: 'bold'}}>
                            ✅ Approved
                          </h4>
                        </div>
                      )
                    ) : (
                      <>
                        {selectedSubmission.formType === 'reimbursement' ? (
                          <>
                            <button
                              style={{...styles.button, backgroundColor: '#28a745'}}
                              onClick={managerApprove}
                              disabled={loading}
                            >
                              {loading ? 'Approving...' : '✅ Approve'}
                            </button>
                            <button
                              style={{...styles.button, backgroundColor: '#dc3545'}}
                              onClick={managerNotApprove}
                              disabled={loading}
                            >
                              {loading ? 'Processing...' : '❌ Not Approve'}
                            </button>
                          </>
                        ) : (
                          <button
                            style={{...styles.button, backgroundColor: '#28a745'}}
                            onClick={managerApprove}
                            disabled={loading}
                          >
                            {loading ? 'Approving...' : '✅ Approve'}
                          </button>
                        )}
                      </>
                    )}
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default AdminCreditCardReview;
