'use client';

import React, { useState, useEffect, useRef } from 'react';
import { collection, getDocs, deleteDoc, doc, setDoc } from 'firebase/firestore';
import { ref, listAll, getDownloadURL, uploadBytes, deleteObject } from 'firebase/storage';
import { db, storage } from '@/lib/firebase.config';
import { enableAllSecurityMeasures, sanitizeCSVCell, sanitizeFirebaseDataClient } from '@/lib/security-client';

// Interfaces for type safety
interface Purchase {
  date: string;
  vendor: string;
  reason: string;
  amount: string;
  description: string;
  receiptFiles: string[];
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
  lastUpdated: Date;
  signatureURL?: string;
  signatureSavedAt?: Date;
  addedOnNumbersChecked?: boolean;
  addedOnNumbersCheckedAt?: Date;
  formType?: 'credit-card' | 'reimbursement'; // Form type to distinguish between credit card receipts and reimbursement requests
  amountAdjustedTo?: string;
  reasonForAdjustment?: string;
  approved?: boolean;
}

interface ReceiptFile {
  name: string;
  url: string;
  purchaseNumber: number;
  fullPath?: string;
}

const AdminCreditCardReview = () => {
  const [submissions, setSubmissions] = useState<Submission[]>([]);
  const [selectedSubmission, setSelectedSubmission] = useState<Submission | null>(null);
  const [loading, setLoading] = useState(false);
  const [signature, setSignature] = useState('');
  const [isDrawing, setIsDrawing] = useState(false);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [receiptFiles, setReceiptFiles] = useState<ReceiptFile[]>([]);
  const [loadingFiles, setLoadingFiles] = useState(false);
  const [savedSignature, setSavedSignature] = useState('');
  const [isSignatureSaved, setIsSignatureSaved] = useState(false);
  const [filterDateFrom, setFilterDateFrom] = useState('');
  const [filterDateTo, setFilterDateTo] = useState('');
  const [filterOffice, setFilterOffice] = useState('');
  const savedSignatureCanvasRef = useRef<HTMLCanvasElement>(null);
  const [amountAdjustedTo, setAmountAdjustedTo] = useState('');
  const [reasonForAdjustment, setReasonForAdjustment] = useState('');

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

  // Load all submissions from Firestore (both credit-card-receipts and reimbursement-requests)
  const loadSubmissions = async () => {
    try {
      setLoading(true);
      
      const submissions: Submission[] = [];
      
      // Load credit card receipts
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
            lastUpdated: data.lastUpdated?.toDate() || new Date(),
            signatureURL: data.signatureURL,
            signatureSavedAt: data.signatureSavedAt?.toDate(),
            addedOnNumbersChecked: data.addedOnNumbersChecked || false,
            addedOnNumbersCheckedAt: data.addedOnNumbersCheckedAt?.toDate(),
            formType: 'credit-card'
          });
        });
      }
      
      // Load reimbursement requests
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
          // Check if data.data exists and is an array
          if (!data.data || !Array.isArray(data.data) || data.data.length === 0) {
            return;
          }
          
          const totalAmount = data.data.reduce((sum: number, item: any) => sum + parseFloat(item.amount || 0), 0);
          
          // Handle createdAt - can be Firestore Timestamp or Date
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
          
          // Handle lastUpdated
          let lastUpdatedDate: Date;
          if (data.lastUpdated) {
            if (data.lastUpdated.toDate && typeof data.lastUpdated.toDate === 'function') {
              lastUpdatedDate = data.lastUpdated.toDate();
            } else if (data.lastUpdated instanceof Date) {
              lastUpdatedDate = data.lastUpdated;
            } else {
              lastUpdatedDate = new Date(data.lastUpdated);
            }
          } else {
            lastUpdatedDate = new Date();
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
            lastUpdated: lastUpdatedDate,
            signatureURL: data.signatureURL,
            signatureSavedAt: data.signatureSavedAt?.toDate ? data.signatureSavedAt.toDate() : (data.signatureSavedAt ? new Date(data.signatureSavedAt) : undefined),
            addedOnNumbersChecked: data.addedOnNumbersChecked || false,
            addedOnNumbersCheckedAt: data.addedOnNumbersCheckedAt?.toDate ? data.addedOnNumbersCheckedAt.toDate() : (data.addedOnNumbersCheckedAt ? new Date(data.addedOnNumbersCheckedAt) : undefined),
            formType: 'reimbursement',
            amountAdjustedTo: data.amountAdjustedTo || undefined,
            reasonForAdjustment: data.reasonForAdjustment || undefined,
            approved: data.approved !== undefined ? data.approved : undefined
          });
        });
      }
      
      // Sort by submission date (newest first)
      submissions.sort((a: Submission, b: Submission) => b.submittedAt.getTime() - a.submittedAt.getTime());
      setSubmissions(submissions);
    } catch (error) {
      alert('Error loading submissions: ' + (error instanceof Error ? error.message : 'Unknown error'));
    } finally {
      setLoading(false);
    }
  };

  // Load receipt files for selected submission
  const loadReceiptFiles = async (submission: Submission) => {
    try {
      setLoadingFiles(true);
      const files: ReceiptFile[] = [];
      
      // Get all files from appropriate folder based on form type
      const storagePath = submission.formType === 'reimbursement' ? 'reimbursement-receipts/' : 'receipts/';
      const listRef = ref(storage, storagePath);
      const result = await listAll(listRef);
      
      // Filter files that belong to this specific submission
      // Use submission ID for precise matching
      let submissionPrefix = '';
      
      if (submission.submissionId) {
        // Use submission ID for exact matching
        // For reimbursement requests, cardNumber may be empty, so handle it differently
        const cardNumberPart = submission.cardNumber && submission.cardNumber.trim() !== '' ? `${submission.cardNumber}_` : '';
        submissionPrefix = `${submission.employeeName}_${cardNumberPart}${submission.submissionId}`;
      } else {
        setReceiptFiles([]);
        setLoadingFiles(false);
        return;
      }
      
      for (const itemRef of result.items) {
        // Check if file starts with the submission prefix
        const matches = itemRef.name.startsWith(submissionPrefix);
        
        if (matches) {
          try {
            const downloadURL = await getDownloadURL(itemRef);
            files.push({
              name: itemRef.name,
              url: downloadURL,
              fullPath: itemRef.fullPath,
              purchaseNumber: 0 // Will be set later when grouping
            });
          } catch (urlError) {
            // Continue with other files even if one fails
          }
        }
      }
      
      // Group files by purchase number first, then sort within each group
      const filesByPurchase: Record<number, ReceiptFile[]> = {};
      
      // Group files by purchase number
      // New filename pattern: Name_CardNumber_Timestamp_purchaseX_SequenceNumber_Filename
      files.forEach((file: ReceiptFile) => {
        // Split filename and find purchase number
        const parts = file.name.split('_');
        let purchaseNum = 0;
        
        for (let i = 0; i < parts.length; i++) {
          // Check if this part contains "purchase" followed by a number
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
      
      // Sort files within each purchase group by sequence number
      Object.keys(filesByPurchase).forEach((purchaseNum: string) => {
        const purchaseNumInt = parseInt(purchaseNum, 10);
        filesByPurchase[purchaseNumInt].sort((a: ReceiptFile, b: ReceiptFile) => {
          const getSequenceNumber = (filename: string): number => {
            // New pattern: Name_CardNumber_Timestamp_purchaseX_SequenceNumber_Filename
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
      
      // Flatten back to single array in purchase order
      const sortedFiles: ReceiptFile[] = [];
      const purchaseNumbers = Object.keys(filesByPurchase).map(Number).sort((a: number, b: number) => a - b);
      
      purchaseNumbers.forEach((purchaseNum: number) => {
        sortedFiles.push(...filesByPurchase[purchaseNum]);
      });
      
      // Replace the original files array with sorted files
      files.length = 0;
      files.push(...sortedFiles);
      
      setReceiptFiles(files);
    } catch (error) {
      setReceiptFiles([]);
    } finally {
      setLoadingFiles(false);
    }
  };

  // Load submissions on component mount
  useEffect(() => {
    loadSubmissions();
  }, []);

  // Initialize canvas when modal opens
  useEffect(() => {
    if (selectedSubmission && !selectedSubmission.signatureURL && canvasRef.current) {
      const canvas = canvasRef.current;
      const ctx = canvas.getContext('2d');
      if (ctx) {
        // Initialize canvas with white background and drawing styles
        ctx.fillStyle = '#fff';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.strokeStyle = '#000';
        ctx.lineWidth = 2;
        ctx.lineCap = 'round';
        ctx.lineJoin = 'round';
      }
    }
  }, [selectedSubmission]);

  // Load saved signature to canvas
  useEffect(() => {
    if (selectedSubmission?.signatureURL && savedSignatureCanvasRef.current) {
      const canvas = savedSignatureCanvasRef.current;
      const ctx = canvas.getContext('2d');
      if (!ctx) return;

      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);

      const img = new Image();
      img.crossOrigin = 'anonymous';
      img.onload = () => {
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      };
      img.src = selectedSubmission.signatureURL;
    }
  }, [selectedSubmission]);

  // Filter submissions based on date range and office
  const getFilteredSubmissions = () => {
    let filtered = submissions;

    // Filter by date range (using submission date, not purchase date)
    if (filterDateFrom || filterDateTo) {
      filtered = filtered.filter((submission: Submission) => {
        // Use the date string directly if available (format: "YYYY-MM-DD")
        // This is the California date selected by the user
        let submissionDateStr = submission.date;
        
        // If date string is not in the expected format, try to convert from submittedAt
        if (!submissionDateStr || !/^\d{4}-\d{2}-\d{2}$/.test(submissionDateStr)) {
          let submissionDate: Date;
          
          if (submission.submittedAt instanceof Date) {
            submissionDate = submission.submittedAt;
          } else {
            submissionDate = new Date(submission.submittedAt);
          }
          
          // Get California time components using Intl.DateTimeFormat
          const formatter = new Intl.DateTimeFormat('en-US', {
            timeZone: 'America/Los_Angeles',
            year: 'numeric',
            month: '2-digit',
            day: '2-digit'
          });
          
          const parts = formatter.formatToParts(submissionDate);
          const year = parts.find(p => p.type === 'year')?.value || '';
          const month = parts.find(p => p.type === 'month')?.value || '';
          const day = parts.find(p => p.type === 'day')?.value || '';
          submissionDateStr = `${year}-${month}-${day}`;
        }
        
        // If only From date is set
        if (filterDateFrom && !filterDateTo) {
          return submissionDateStr >= filterDateFrom;
        }
        
        // If only To date is set
        if (!filterDateFrom && filterDateTo) {
          return submissionDateStr <= filterDateTo;
        }
        
        // If both dates are set
        if (filterDateFrom && filterDateTo) {
          return submissionDateStr >= filterDateFrom && submissionDateStr <= filterDateTo;
        }
        
        return true;
      });
    }

    // Filter by office
    if (filterOffice) {
      filtered = filtered.filter((submission: Submission) => 
        submission.office === filterOffice
      );
    }

    return filtered;
  };

  // Get unique offices for filter dropdown
  const getUniqueOffices = (): string[] => {
    const offices = [...new Set(submissions.map((s: Submission) => s.office))];
    return offices.filter((office): office is string => Boolean(office) && office !== 'N/A');
  };

  // Handle signature drawing
  const startDrawing = (e: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>) => {
    e.preventDefault();
    setIsDrawing(true);
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const rect = canvas.getBoundingClientRect();
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;
    
    let clientX, clientY;
    if ('touches' in e) {
      clientX = e.touches[0].clientX;
      clientY = e.touches[0].clientY;
    } else {
      clientX = e.clientX;
      clientY = e.clientY;
    }
    
    ctx.beginPath();
    ctx.moveTo((clientX - rect.left) * scaleX, (clientY - rect.top) * scaleY);
  };

  const draw = (e: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>) => {
    e.preventDefault();
    if (!isDrawing) return;
    
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const rect = canvas.getBoundingClientRect();
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;
    
    let clientX, clientY;
    if ('touches' in e) {
      clientX = e.touches[0].clientX;
      clientY = e.touches[0].clientY;
    } else {
      clientX = e.clientX;
      clientY = e.clientY;
    }
    
    ctx.lineTo((clientX - rect.left) * scaleX, (clientY - rect.top) * scaleY);
    ctx.stroke();
  };

  const stopDrawing = () => {
    setIsDrawing(false);
    const canvas = canvasRef.current;
    if (canvas) {
      setSignature(canvas.toDataURL());
    }
  };

  const clearSignature = () => {
    const canvas = canvasRef.current;
    if (canvas) {
      const ctx = canvas.getContext('2d');
      if (ctx) {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        // Reset drawing styles
        ctx.strokeStyle = '#000';
        ctx.lineWidth = 2;
        ctx.lineCap = 'round';
        ctx.lineJoin = 'round';
        // Fill with white background
        ctx.fillStyle = '#fff';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
      }
    }
    setSignature('');
  };

  // Check Added on Numbers (first step)
  const checkAddedOnNumbers = async () => {
    if (!selectedSubmission) {
      alert('No submission selected');
      return;
    }
    
    try {
      setLoading(true);
      
      // Update Firestore with checked status (use appropriate collection based on form type)
      const collectionName = selectedSubmission.formType === 'reimbursement' ? 'reimbursement-requests' : 'credit-card-receipts';
      const docRef = doc(db, collectionName, selectedSubmission.id);
      
      // 🔒 Firebase 데이터 sanitization 적용
      const updateData = sanitizeFirebaseDataClient({
        addedOnNumbersChecked: true,
        addedOnNumbersCheckedAt: new Date()
      });
      await setDoc(docRef, updateData, { merge: true });
      
      // Update local state
      setSelectedSubmission({
        ...selectedSubmission,
        addedOnNumbersChecked: true,
        addedOnNumbersCheckedAt: new Date()
      });
      
      // Update submissions list
      setSubmissions(prev => prev.map((sub: Submission) => 
        sub.id === selectedSubmission.id 
          ? { ...sub, addedOnNumbersChecked: true, addedOnNumbersCheckedAt: new Date() }
          : sub
      ));
      
      alert('✅ Added on Numbers checked successfully!');
      
    } catch (error) {
      alert('Error checking Added on Numbers: ' + (error instanceof Error ? error.message : 'Unknown error'));
    } finally {
      setLoading(false);
    }
  };

  // Manager Not Approve - Save signature with rejection status for reimbursement
  const managerNotApprove = async () => {
    if (!selectedSubmission) {
      alert('No submission selected');
      return;
    }
    
    if (selectedSubmission.formType !== 'reimbursement') {
      alert('Not Approved option is only available for reimbursement requests.');
      return;
    }
    
    if (!signature) {
      alert('Please provide a signature before processing.');
      return;
    }
    
    try {
      setLoading(true);
      
      // Convert signature data URL to blob
      if (!signature || typeof signature !== 'string') {
        throw new Error('Invalid signature data');
      }
      const base64Data = signature.split(',')[1];
      const byteCharacters = atob(base64Data);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray], { type: 'image/png' });
      
      // Create filename with timestamp
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const cardNumberPart = selectedSubmission.cardNumber && selectedSubmission.cardNumber.trim() !== '' 
        ? `${selectedSubmission.cardNumber}_` 
        : '';
      const fileName = `signatures/${selectedSubmission.employeeName}_${cardNumberPart}${timestamp}.png`;
      
      // Upload to Firebase Storage
      const storageRef = ref(storage, fileName);
      await uploadBytes(storageRef, blob);
      const downloadURL = await getDownloadURL(storageRef);
      
      // Save signature data
      setSavedSignature(signature);
      setIsSignatureSaved(true);
      
      // Store signature URL and rejection info in Firestore
      const collectionName = 'reimbursement-requests';
      const docRef = doc(db, collectionName, selectedSubmission.id);
      
      let retryCount = 0;
      const maxRetries = 3;
      let firestoreSuccess = false;
      
      while (retryCount < maxRetries && !firestoreSuccess) {
        try {
          const updateData: any = {
            signatureURL: downloadURL,
            signatureSavedAt: new Date(),
            approved: false, // Mark as not approved
            rejectionDate: new Date()
          };
          
          // 금액 및 사유 검증 및 sanitization
          if (amountAdjustedTo) {
            const sanitizedAmount = parseFloat(amountAdjustedTo);
            if (!isNaN(sanitizedAmount) && sanitizedAmount >= 0 && sanitizedAmount <= 1000000) {
              updateData.amountAdjustedTo = sanitizedAmount.toFixed(2);
            }
          }
          if (reasonForAdjustment) {
            updateData.reasonForAdjustment = reasonForAdjustment.substring(0, 1000).replace(/[<>]/g, '');
          }
          
          // 🔒 Firebase 데이터 sanitization 적용
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
      
      // Generate PDF with rejection status (saveToExcel handles not approved status via Firestore data)
      await saveToExcel();
      
      alert('❌ Reimbursement request marked as Not Approved. Signature saved and PDF generated.');
      
      // Close modal and refresh submissions
      setSelectedSubmission(null);
      setSignature('');
      setSavedSignature('');
      setIsSignatureSaved(false);
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

  // Manager Approve - Save signature and generate PDF in one action
  const managerApprove = async () => {
    if (!selectedSubmission) {
      alert('No submission selected');
      return;
    }
    
    if (!signature) {
      alert('Please provide a signature before approving.');
      return;
    }
    
    try {
      setLoading(true);
      
      // Convert signature data URL to blob (without fetch to avoid CORS)
      if (!signature || typeof signature !== 'string') {
        throw new Error('Invalid signature data');
      }
      const base64Data = signature.split(',')[1];
      const byteCharacters = atob(base64Data);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray], { type: 'image/png' });
      
      // Create filename with timestamp
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      // For reimbursement, cardNumber may be empty, so handle it differently
      const cardNumberPart = selectedSubmission.cardNumber && selectedSubmission.cardNumber.trim() !== '' 
        ? `${selectedSubmission.cardNumber}_` 
        : '';
      const fileName = `signatures/${selectedSubmission.employeeName}_${cardNumberPart}${timestamp}.png`;
      
      // Upload to Firebase Storage
      const storageRef = ref(storage, fileName);
      await uploadBytes(storageRef, blob);
      const downloadURL = await getDownloadURL(storageRef);
      
      // Save signature data
      setSavedSignature(signature);
      setIsSignatureSaved(true);
      
      // Store signature URL in Firestore for future reference (with retry logic)
      // Use appropriate collection based on form type
      const collectionName = selectedSubmission.formType === 'reimbursement' ? 'reimbursement-requests' : 'credit-card-receipts';
      const docRef = doc(db, collectionName, selectedSubmission.id);
      
      let retryCount = 0;
      const maxRetries = 3;
      let firestoreSuccess = false;
      
      while (retryCount < maxRetries && !firestoreSuccess) {
        try {
          const updateData: any = {
            signatureURL: downloadURL,
            signatureSavedAt: new Date()
          };
          
          // For reimbursement, mark as approved and optionally save adjustment info
          if (selectedSubmission.formType === 'reimbursement') {
            updateData.approved = true;
            if (amountAdjustedTo) {
              // 금액 검증 및 sanitization
              const sanitizedAmount = parseFloat(amountAdjustedTo);
              if (!isNaN(sanitizedAmount) && sanitizedAmount >= 0 && sanitizedAmount <= 1000000) {
                updateData.amountAdjustedTo = sanitizedAmount.toFixed(2);
              }
            }
            if (reasonForAdjustment) {
              // 텍스트 sanitization (길이 제한 및 특수문자 제거)
              updateData.reasonForAdjustment = reasonForAdjustment.substring(0, 1000).replace(/[<>]/g, '');
            }
          }
          
          // 🔒 Firebase 데이터 sanitization 적용
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
      
      // Generate PDF and save to Excel (without opening print dialog)
      await saveToExcel();
      
      alert('✅ Approved! Signature saved and PDF generated. Data saved to Excel file.');
      
      // Close modal and refresh submissions
      setSelectedSubmission(null);
      setSignature('');
      setSavedSignature('');
      setIsSignatureSaved(false);
      setReceiptFiles([]);
      setAmountAdjustedTo('');
      setReasonForAdjustment('');
      loadSubmissions();
      
    } catch (error) {
      alert('Error approving submission: ' + (error instanceof Error ? error.message : 'Unknown error'));
    } finally {
      setLoading(false);
    }
  };

  // Download CSV file
  const downloadExcelFile = async () => {
    try {
      setLoading(true);
      
      // Download the main CSV file (use the same filename as saveToExcel)
      const mainFileName = 'all-submissions.csv';
      const mainFileRef = ref(storage, `excel/${mainFileName}`);
      
      try {
        const downloadURL = await getDownloadURL(mainFileRef);
        
        // Download CSV file
        const link = document.createElement('a');
        link.href = downloadURL;
        link.download = mainFileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        alert('📊 CSV file downloaded successfully!');
      } catch (error) {
        alert('❌ CSV file not found. Please approve a submission first to create the CSV file.');
      }
    } catch (error) {
      alert('Error downloading CSV file: ' + (error instanceof Error ? error.message : 'Unknown error'));
    } finally {
      setLoading(false);
    }
  };

  // Save data to Excel file (client-side)
  const saveToExcel = async () => {
    if (!selectedSubmission) {
      return;
    }
    
    // Create local reference to avoid null checks
    const submission = selectedSubmission;
    
    // Check if submission has signature (either saved in Firestore or current session)
    const hasSignatureURL = submission.signatureURL && submission.signatureURL.trim() !== '';
    const hasCurrentSignature = (savedSignature && savedSignature.trim() !== '') || (signature && signature.trim() !== '');
    
    if (!hasSignatureURL && !hasCurrentSignature) {
      alert('❌ This submission has no signature. Please sign first before saving.');
      return;
    }
    
    try {
      const totalAmount = submission.purchases.reduce((sum: number, purchase: Purchase) => {
        return sum + (parseFloat(purchase.amount) || 0);
      }, 0);

      // Define CSV file reference (use combined file for both types)
      const mainFileName = 'all-submissions.csv';
      const mainFileRef = ref(storage, `excel/${mainFileName}`);
      
      let existingData = [
        ['Form Type', 'Employee Name', 'Office', 'Card Number', 'Purchase Date', 'Store/Website', 'Reason', 'Amount', 'Account Description', 'Total Amount', 'Submission Date', 'Status', 'PDF Link']
      ];
      
      try {
        // Get all submissions that have been processed (have signatureURL) from both collections
        const processedSubmissions: any[] = [];
        
        // Get credit card receipts
        const creditCardSnapshot = await getDocs(collection(db, 'credit-card-receipts'));
        creditCardSnapshot.forEach((doc) => {
          const data = doc.data();
          if (data.signatureURL) { // Only include signed/processed submissions
            processedSubmissions.push({
              id: doc.id,
              employeeName: data.name,
              office: data.office || 'N/A',
              cardNumber: data.cardNumber,
              date: data.date,
              purchases: data.data,
              signatureURL: data.signatureURL,
              signatureSavedAt: data.signatureSavedAt,
              pdfURL: data.pdfURL || '',
              formType: 'Credit Card Receipt'
            });
          }
        });
        
        // Get reimbursement requests
        const reimbursementSnapshot = await getDocs(collection(db, 'reimbursement-requests'));
        reimbursementSnapshot.forEach((doc) => {
          const data = doc.data();
          if (data.signatureURL) { // Only include signed/processed submissions
            processedSubmissions.push({
              id: doc.id,
              employeeName: data.name,
              office: data.office || 'N/A',
              cardNumber: data.cardNumber,
              date: data.date,
              purchases: data.data,
              signatureURL: data.signatureURL,
              signatureSavedAt: data.signatureSavedAt,
              pdfURL: data.pdfURL || '',
              formType: 'Reimbursement Request'
            });
          }
        });
        
        // Add all processed submissions to CSV data (excluding current submission)
        for (const submission of processedSubmissions) {
          // Skip current submission to avoid duplication
          // Note: using 'submission' constant defined at function start
          const currentSubmissionId = selectedSubmission?.id;
          if (submission.id === currentSubmissionId) {
            continue;
          }
          
          const totalAmount = submission.purchases.reduce((sum: number, purchase: Purchase) => {
            return sum + (parseFloat(purchase.amount) || 0);
          }, 0);
          
            // Add each purchase as a row with CSV Injection protection
            submission.purchases.forEach((purchase: Purchase, index: number) => {
              existingData.push([
                index === 0 ? sanitizeCSVCell(submission.formType || 'Credit Card Receipt') : '',
                index === 0 ? sanitizeCSVCell(submission.employeeName) : '',
                index === 0 ? sanitizeCSVCell(submission.office) : '',
                index === 0 ? sanitizeCSVCell(`****${submission.cardNumber}`) : '',
                sanitizeCSVCell(purchase.date),
                sanitizeCSVCell(purchase.vendor),
                sanitizeCSVCell(purchase.reason),
                sanitizeCSVCell(`$${parseFloat(purchase.amount).toFixed(2)}`),
                sanitizeCSVCell(purchase.description),
                index === 0 ? sanitizeCSVCell(`$${totalAmount.toFixed(2)}`) : '',
                sanitizeCSVCell(submission.signatureSavedAt ? submission.signatureSavedAt.toDate().toLocaleDateString() : new Date().toLocaleDateString()),
                sanitizeCSVCell('Approved & PDF Generated'),
                index === 0 ? sanitizeCSVCell(submission.pdfURL || '') : ''
              ]);
            });
        }
        
      } catch (error) {
      }

      // Use submission constant defined at function start
      const currentSubmission = selectedSubmission;
      if (!currentSubmission) return;
      
      const formTypeLabel = currentSubmission.formType === 'reimbursement' ? 'Reimbursement Request' : 'Credit Card Receipt';
      const newRows = currentSubmission.purchases.map((purchase: Purchase, index: number) => [
        index === 0 ? sanitizeCSVCell(formTypeLabel) : '',
        index === 0 ? sanitizeCSVCell(currentSubmission.employeeName) : '',
        index === 0 ? sanitizeCSVCell(currentSubmission.office) : '',
        index === 0 ? sanitizeCSVCell(`****${currentSubmission.cardNumber}`) : '',
        sanitizeCSVCell(purchase.date),
        sanitizeCSVCell(purchase.vendor),
        sanitizeCSVCell(purchase.reason),
        sanitizeCSVCell(`$${parseFloat(purchase.amount).toFixed(2)}`),
        sanitizeCSVCell(purchase.description),
        index === 0 ? sanitizeCSVCell(`$${totalAmount.toFixed(2)}`) : '',
        sanitizeCSVCell(new Date().toLocaleDateString()),
        sanitizeCSVCell('Approved & PDF Generated'),
        ''
      ]);

      // Combine existing and new data
      const updatedData = [...existingData, ...newRows];
      
      // Convert to CSV string with proper escaping (secure version)
      // sanitizeCSVCell already handles CSV injection protection
      // For extra safety, ensure proper CSV formatting with quotes when needed
      const csvString = updatedData.map((row: any[]) => 
        row.map((cell: string) => {
          // sanitizeCSVCell already processed the cell, but ensure proper CSV formatting
          // If cell contains comma, quote, or newline, wrap in quotes and escape internal quotes
          if (cell === null || cell === undefined) return '';
          const cellStr = String(cell);
          if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n') || cellStr.includes('\r')) {
            // Escape quotes by doubling them and wrap in quotes
            return `"${cellStr.replace(/"/g, '""')}"`;
          }
          return cellStr;
        }).join(',')
      ).join('\n');
      
      // Create blob and upload to Firebase Storage (like credit-card-receipts.tsx)
      try {
        const blob = new Blob([csvString], { type: 'text/csv' });
        await uploadBytes(mainFileRef, blob);
      } catch (error) {
        // Continue execution even if CSV upload fails
      }
      
      // CSV 저장 완료 후 Storage에서 서명 삭제
      if (selectedSubmission?.signatureURL && selectedSubmission.signatureURL.startsWith('https://')) {
        try {
          const url = new URL(selectedSubmission.signatureURL);
          const pathMatch = url.pathname.match(/\/o\/(.+)\?/);
          
          if (pathMatch && pathMatch[1]) {
            const filePath = decodeURIComponent(pathMatch[1]);
            const fileRef = ref(storage, filePath);
            await deleteObject(fileRef);
          }
        } catch (deleteError) {
          // 삭제 실패해도 계속 진행
        }
      }
      
    } catch (error) {
    }
  };

  // Delete completed submissions only (Reset function)
  const resetAllData = async () => {
    if (!confirm('⚠️ WARNING: This will delete ONLY completed submissions (signed and numbers checked). Incomplete submissions will remain. Are you sure you want to continue?')) {
      return;
    }
    
    try {
      setLoading(true);
      
      // Get all submissions from both collections
      const creditCardSnapshot = await getDocs(collection(db, 'credit-card-receipts'));
      const reimbursementSnapshot = await getDocs(collection(db, 'reimbursement-requests'));
      let completedCount = 0;
      let incompleteCount = 0;
      
      // First, collect completed submissions and delete their Storage files
      const completedSubmissions: Array<{id: string, employeeName: string, cardNumber: string, data: any, docRef: any, formType: 'credit-card' | 'reimbursement'}> = [];
      const deletePromises: Promise<any>[] = [];
      
      // Process credit card receipts
      creditCardSnapshot.forEach((doc) => {
        const data = doc.data();
        const isSigned = data.signatureURL && typeof data.signatureURL === 'string' && data.signatureURL.trim() !== '';
        const isNumbersChecked = data.addedOnNumbersChecked === true;
        
        if (isSigned && isNumbersChecked) {
          // Both signature and numbers check are complete - safe to delete
          completedSubmissions.push({
            id: doc.id,
            employeeName: data.name,
            cardNumber: data.cardNumber || '',
            data: data.data,
            docRef: doc.ref,
            formType: 'credit-card'
          });
          completedCount++;
        } else {
          // Incomplete submission - keep it
          incompleteCount++;
        }
      });
      
      // Process reimbursement requests
      reimbursementSnapshot.forEach((doc) => {
        const data = doc.data();
        const isSigned = data.signatureURL && typeof data.signatureURL === 'string' && data.signatureURL.trim() !== '';
        const isNumbersChecked = data.addedOnNumbersChecked === true;
        
        if (isSigned && isNumbersChecked) {
          // Both signature and numbers check are complete - safe to delete
          completedSubmissions.push({
            id: doc.id,
            employeeName: data.name,
            cardNumber: data.cardNumber || '',
            data: data.data,
            docRef: doc.ref,
            formType: 'reimbursement'
          });
          completedCount++;
        } else {
          // Incomplete submission - keep it
          incompleteCount++;
        }
      });
      
      if (completedCount === 0) {
        alert('No completed submissions found to reset. All submissions must be both signed and have numbers checked.');
        setLoading(false);
        return;
      }
      
      if (incompleteCount > 0) {
        alert(`Found ${incompleteCount} incomplete submissions that will be kept. Only ${completedCount} completed submissions will be deleted.`);
      }
      
      // Delete Storage files for completed submissions FIRST
      
      try {
        // Delete files for each completed submission
        for (const submission of completedSubmissions) {
          try {
            // Determine storage path based on form type
            const receiptStoragePath = submission.formType === 'reimbursement' ? 'reimbursement-receipts/' : 'receipts/';
            
            for (const purchase of submission.data) {
              if (purchase.receiptFiles) {
                let receiptFiles: string[] = [];
                
                // Handle different data types
                if (typeof purchase.receiptFiles === 'string') {
                  receiptFiles = purchase.receiptFiles.split(', ');
                } else if (Array.isArray(purchase.receiptFiles)) {
                  receiptFiles = purchase.receiptFiles;
                }
                
                for (const fileName of receiptFiles) {
                  if (fileName && typeof fileName === 'string' && fileName.trim()) {
                    const receiptRef = ref(storage, `${receiptStoragePath}${fileName.trim()}`);
                    try {
                      await deleteObject(receiptRef);
                    } catch (error) {
                    }
                  }
                }
              }
            }
            
            // Delete signature file for this submission
            try {
              // For reimbursement, cardNumber may be empty, so handle it differently
              const cardNumberPart = submission.cardNumber && submission.cardNumber.trim() !== '' 
                ? `${submission.cardNumber}_` 
                : '';
              const signatureSearchPattern = `${submission.employeeName}_${cardNumberPart}`;
              const signaturesRef = ref(storage, 'signatures/');
              const signaturesList = await listAll(signaturesRef);
              
              for (const item of signaturesList.items) {
                if (item.name.includes(signatureSearchPattern)) {
                  await deleteObject(item);
                }
              }
            } catch (error) {
            }
            
          } catch (submissionError) {
          }
        }
        
        // Delete Excel file (this is cumulative, so safe to delete)
        const excelRef = ref(storage, 'excel/');
        const excelList = await listAll(excelRef);
        for (const item of excelList.items) {
          await deleteObject(item);
        }
        
      } catch (storageError) {
        // Continue even if some Storage files fail to delete
      }
      
      // Now delete Firestore documents
      for (const submission of completedSubmissions) {
        deletePromises.push(deleteDoc(submission.docRef));
      }
      
      await Promise.all(deletePromises);
      
      // Reset state
      setSubmissions([]);
      setSelectedSubmission(null);
      setSignature('');
      setSavedSignature('');
      setIsSignatureSaved(false);
      setReceiptFiles([]);
      
      alert(`✅ Reset completed! Deleted ${completedCount} completed submissions. ${incompleteCount} incomplete submissions were kept.`);
      
    } catch (error) {
      alert('Error resetting data: ' + (error instanceof Error ? error.message : 'Unknown error'));
    } finally {
      setLoading(false);
    }
  };

  // Delete single submission data (but keep Excel file)
  const deleteAllSubmissionData = async () => {
    if (!selectedSubmission) {
      alert('No submission selected');
      return;
    }
    
    try {
      // Delete from Firestore
      await deleteDoc(doc(db, 'credit-card-receipts', selectedSubmission.id));
      
      // Delete receipt files from Storage
      for (const file of receiptFiles) {
        try {
          const fileRef = ref(storage, file.fullPath);
          await deleteObject(fileRef);
        } catch (error) {
        }
      }
      
      // Delete signature file from Storage if exists
      if (selectedSubmission.signatureURL) {
        try {
          // Extract file path from URL
          const url = new URL(selectedSubmission.signatureURL);
          const pathMatch = url.pathname.match(/\/o\/(.+)\?/);
          if (pathMatch) {
            const filePath = decodeURIComponent(pathMatch[1]);
            const signatureRef = ref(storage, filePath);
            await deleteObject(signatureRef);
          }
        } catch (error) {
          // Continue execution even if signature deletion fails
        }
      }
      
      alert('✅ PDF generated and data saved to Excel! Submission removed from review list.');
      setSelectedSubmission(null);
      setSignature('');
      setSavedSignature('');
      setIsSignatureSaved(false);
      setReceiptFiles([]);
      loadSubmissions(); // Reload submissions
      
    } catch (error) {
      alert('Error deleting data: ' + (error instanceof Error ? error.message : 'Unknown error'));
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
      backgroundColor: 'rgba(173, 216, 230, 0.95)', // 하늘색 배경
      padding: '20px',
      borderRadius: '15px',
      boxShadow: '0 8px 32px rgba(0, 0, 0, 0.1)',
      cursor: 'pointer',
      transition: 'all 0.3s ease',
      border: '2px solid #87CEEB' // 하늘색 테두리
    },
    submissionCardHover: {
      transform: 'translateY(-5px)',
      borderColor: '#4CAF50'
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
    signatureTitle: {
      fontSize: '18px',
      fontWeight: 'bold',
      marginBottom: '15px',
      color: '#2c3e50'
    },
    signatureCanvas: {
      border: '2px solid #dee2e6',
      borderRadius: '8px',
      cursor: 'crosshair',
      backgroundColor: 'white',
      touchAction: 'none' as const,
      maxWidth: '100%'
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
    buttonSecondary: {
      backgroundColor: '#6c757d',
      color: 'white',
      padding: '8px 16px',
      border: 'none',
      borderRadius: '6px',
      cursor: 'pointer',
      fontSize: '14px'
    },
    loading: {
      textAlign: 'center' as const,
      padding: '50px',
      fontSize: '18px',
      color: '#666'
    }
  };

  return (
    <div style={styles.body}>
      <div style={styles.container}>
        <header style={styles.header}>
          <h1 style={styles.title}>🏢 Credit Card Receipts</h1>
          
          {/* Filter Section */}
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
              <label style={{fontSize: '12px', fontWeight: 'bold', color: '#495057'}}>Date From:</label>
              <input
                type="date"
                value={filterDateFrom}
                onChange={(e) => setFilterDateFrom(e.target.value)}
                style={{
                  padding: '8px 12px',
                  border: '1px solid #ced4da',
                  borderRadius: '4px',
                  fontSize: '14px'
                }}
              />
            </div>
            
            <div style={{display: 'flex', flexDirection: 'column', gap: '5px'}}>
              <label style={{fontSize: '12px', fontWeight: 'bold', color: '#495057'}}>Date To:</label>
              <input
                type="date"
                value={filterDateTo}
                onChange={(e) => setFilterDateTo(e.target.value)}
                style={{
                  padding: '8px 12px',
                  border: '1px solid #ced4da',
                  borderRadius: '4px',
                  fontSize: '14px'
                }}
              />
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
                onClick={() => {
                  setFilterDateFrom('');
                  setFilterDateTo('');
                  setFilterOffice('');
                }}
              >
                Clear Filters
              </button>
            </div>
          </div>
          
          {/* Filter Status */}
          {(filterDateFrom || filterDateTo || filterOffice) && (
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
              {filterDateFrom && ` • From: ${filterDateFrom}`}
              {filterDateTo && ` • To: ${filterDateTo}`}
              {filterOffice && ` • Office: ${filterOffice}`}
            </div>
          )}
          
          <div style={{display: 'flex', gap: '10px', justifyContent: 'center', marginTop: '15px'}}>
            <button
              style={{
                ...styles.button,
                backgroundColor: '#28a745',
                fontSize: '14px',
                padding: '10px 20px'
              }}
              onClick={downloadExcelFile}
              disabled={loading}
            >
              📊 Download CSV File
            </button>
            <button
              style={{
                ...styles.button,
                backgroundColor: '#dc3545',
                fontSize: '14px',
                padding: '10px 20px'
              }}
              onClick={resetAllData}
              disabled={loading}
            >
              🔄 Reset All Data
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
              {(filterDateFrom || filterDateTo || filterOffice) ? 
                'No submissions match your current filters. Try adjusting your filter criteria.' : 
                'No submissions have been submitted yet.'
              }
            </p>
            {(filterDateFrom || filterDateTo || filterOffice) && (
              <button
                style={{
                  ...styles.button,
                  backgroundColor: '#6c757d',
                  fontSize: '14px',
                  padding: '10px 20px'
                }}
                onClick={() => {
                  setFilterDateFrom('');
                  setFilterDateTo('');
                  setFilterOffice('');
                }}
              >
                Clear Filters
              </button>
            )}
          </div>
        ) : (
          <div style={styles.submissionsGrid}>
            {getFilteredSubmissions().map((submission: Submission) => {
              // Different colors for reimbursement vs credit card
              // Reimbursement: white background (same for both signed and unsigned)
              const isReimbursement = submission.formType === 'reimbursement';
              const cardStyle = submission.signatureURL 
                ? (isReimbursement 
                  ? styles.submissionCardSigned // Use same sky blue when signed
                  : styles.submissionCardSigned)
                : (isReimbursement 
                  ? styles.submissionCard // Use white background when not signed
                  : styles.submissionCard);
              
              return (
              <div
                key={submission.id}
                style={cardStyle}
                onClick={() => {
                  setSelectedSubmission(submission);
                  loadReceiptFiles(submission);
                }}
              >
                <div style={styles.submissionHeader}>
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
                    {submission.signatureURL && (
                      <span style={{
                        marginLeft: '10px',
                        fontSize: '16px',
                        color: '#0066cc'
                      }}>
                        ✍️ Signed
                      </span>
                    )}
                  </div>
                  <div style={styles.submissionDate}>
                    Submitted: {submission.date || (() => {
                      // Fallback: Convert to California time zone if date string not available
                      const submissionDate = submission.submittedAt instanceof Date 
                        ? submission.submittedAt 
                        : new Date(submission.submittedAt);
                      
                      const formatter = new Intl.DateTimeFormat('en-US', {
                        timeZone: 'America/Los_Angeles',
                        year: 'numeric',
                        month: '2-digit',
                        day: '2-digit'
                      });
                      
                      const parts = formatter.formatToParts(submissionDate);
                      const year = parts.find(p => p.type === 'year')?.value || '';
                      const month = parts.find(p => p.type === 'month')?.value || '';
                      const day = parts.find(p => p.type === 'day')?.value || '';
                      
                      return `${year}-${month}-${day}`;
                    })()}
                  </div>
                </div>
                <div style={styles.submissionDetails}>
                  <p><strong>Office:</strong> {submission.office}</p>
                  {submission.formType !== 'reimbursement' && (
                    <p><strong>Card:</strong> ****{submission.cardNumber}</p>
                  )}
                  <p><strong>Purchases:</strong> {submission.purchases.length} items</p>
                  <p><strong>Total Amount:</strong> ${submission.totalAmount}</p>
                  
                  {/* Added on Numbers check only */}
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
                        📊 Added on Numbers:
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
                  
                  {submission.signatureURL && (
                    <p style={{color: '#0066cc', fontWeight: 'bold', marginTop: '8px'}}>
                      ✅ Signed this submission
                    </p>
                  )}
                </div>
              </div>
              );
            })}
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
                  <p><strong>Card Number:</strong> ****{selectedSubmission.cardNumber}</p>
                )}
                <p><strong>Submission Date:</strong> {selectedSubmission.date || (() => {
                  // Fallback: Convert to California time zone if date string not available
                  const submissionDate = selectedSubmission.submittedAt instanceof Date 
                    ? selectedSubmission.submittedAt 
                    : new Date(selectedSubmission.submittedAt);
                  
                  const formatter = new Intl.DateTimeFormat('en-US', {
                    timeZone: 'America/Los_Angeles',
                    year: 'numeric',
                    month: '2-digit',
                    day: '2-digit'
                  });
                  
                  const parts = formatter.formatToParts(submissionDate);
                  const year = parts.find(p => p.type === 'year')?.value || '';
                  const month = parts.find(p => p.type === 'month')?.value || '';
                  const day = parts.find(p => p.type === 'day')?.value || '';
                  
                  return `${year}-${month}-${day}`;
                })()}</p>

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
                      <th style={styles.tableHeader}>Store/Website</th>
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

                {/* Receipt Files Section */}
                <div style={{marginTop: '30px', padding: '20px', backgroundColor: '#f8f9fa', borderRadius: '10px', border: '2px solid #e9ecef'}}>
                  <h3 style={{marginBottom: '15px', color: '#2c3e50'}}>📎 Receipt Files</h3>
                  {loadingFiles ? (
                    <p>Loading files...</p>
                  ) : receiptFiles.length > 0 ? (
                    <div style={{display: 'flex', flexDirection: 'column', gap: '25px'}}>
                      {(() => {
                        // Group files by purchase number
                        const filesByPurchase: Record<number, ReceiptFile[]> = {};
                        receiptFiles.forEach((file: ReceiptFile) => {
                          // Extract purchase number from new filename pattern: Name_CardNumber_Timestamp_purchaseX_SequenceNumber_Filename
                          // e.g., "John_1234_2024-01-15T10-30-45-123Z_purchase1_0001759527445729_receipt.jpg" -> Purchase 1
                          const parts = file.name.split('_');
                          let purchaseNum = 0;
                          
                          // Find the purchase number by looking for "purchaseX" pattern
                          // Pattern: Name_CardNumber_SubmissionID_purchaseX_SequenceNumber_Filename
                          for (let i = 0; i < parts.length; i++) {
                            // Check if this part contains "purchase" followed by a number
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
                        
                        // Sort purchase numbers and render each group
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
                  <div style={styles.signatureTitle}>Digital Signature</div>
                  
                  {/* Reimbursement adjustment fields - only show for reimbursement and when not signed */}
                  {selectedSubmission.formType === 'reimbursement' && !selectedSubmission.signatureURL && (
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
                  
                  <p>Please sign below to approve this submission:</p>
                  
                  {/* Show saved signature if exists */}
                  {selectedSubmission.signatureURL && (
                    <div style={{marginBottom: '20px', padding: '15px', backgroundColor: '#e8f5e8', borderRadius: '8px', border: '2px solid #28a745'}}>
                      <h4 style={{margin: '0 0 10px 0', color: '#155724'}}>✅ Saved Signature</h4>
                      <canvas 
                        ref={savedSignatureCanvasRef}
                        width={800}
                        height={200}
                        style={{
                          maxWidth: '300px',
                          maxHeight: '100px',
                          width: '100%',
                          height: 'auto',
                          border: '1px solid #ddd',
                          backgroundColor: 'white',
                          borderRadius: '4px',
                          display: 'block'
                        }}
                      />
                      <p style={{margin: '10px 0 0 0', fontSize: '12px', color: '#666'}}>
                        Saved on: {selectedSubmission?.signatureSavedAt ? 
                          (selectedSubmission.signatureSavedAt && typeof (selectedSubmission.signatureSavedAt as any).toDate === 'function' ? 
                            new Date((selectedSubmission.signatureSavedAt as any).toDate()).toLocaleString() : 
                            new Date(selectedSubmission.signatureSavedAt as any).toLocaleString()) : 
                          'Unknown'}
                      </p>
                    </div>
                  )}
                  
                  {/* Only show canvas if not already signed */}
                  {!selectedSubmission?.signatureURL && (
                    <canvas
                      ref={canvasRef}
                      width={800}
                      height={200}
                      style={{...styles.signatureCanvas, width: '100%', height: 'auto', aspectRatio: '4/1'}}
                      onMouseDown={startDrawing}
                      onMouseMove={draw}
                      onMouseUp={stopDrawing}
                      onMouseLeave={stopDrawing}
                      onTouchStart={startDrawing}
                      onTouchMove={draw}
                      onTouchEnd={stopDrawing}
                    />
                  )}
                  
                  <div style={styles.signatureButtons}>
                    {selectedSubmission.signatureURL ? (
                      // Already signed - show approval status based on approved field (for reimbursement)
                      selectedSubmission.formType === 'reimbursement' && selectedSubmission.approved === false ? (
                        // Reimbursement: Not Approved
                        <div style={{
                          padding: '15px',
                          backgroundColor: '#f8d7da',
                          borderRadius: '8px',
                          border: '2px solid #dc3545',
                          textAlign: 'center'
                        }}>
                          <h4 style={{margin: '0 0 10px 0', color: '#721c24'}}>❌ Not Approved</h4>
                          <p style={{margin: '0', color: '#721c24', fontSize: '14px'}}>
                            This reimbursement request has been marked as not approved and PDF has been generated.
                          </p>
                        </div>
                      ) : (
                        // Approved (default for credit card or reimbursement with approved: true)
                        <div style={{
                          padding: '15px',
                          backgroundColor: '#d4edda',
                          borderRadius: '8px',
                          border: '2px solid #28a745',
                          textAlign: 'center'
                        }}>
                          <h4 style={{margin: '0 0 10px 0', color: '#155724'}}>✅ Approved</h4>
                          <p style={{margin: '0', color: '#155724', fontSize: '14px'}}>
                            This submission has been approved and PDF has been generated.
                          </p>
                        </div>
                      )
                    ) : (
                      // Not signed yet - show signature buttons
                      <>
                        <button
                          style={styles.buttonSecondary}
                          onClick={clearSignature}
                        >
                          Clear Signature
                        </button>
                        {selectedSubmission.formType === 'reimbursement' ? (
                          // Reimbursement: Show both Approve and Not Approved buttons
                          <>
                            <button
                              style={{...styles.button, backgroundColor: '#28a745'}}
                              onClick={managerApprove}
                              disabled={!signature || loading}
                            >
                              {loading ? 'Approving...' : '✅ Approve'}
                            </button>
                            <button
                              style={{...styles.button, backgroundColor: '#dc3545'}}
                              onClick={managerNotApprove}
                              disabled={!signature || loading}
                            >
                              {loading ? 'Processing...' : '❌ Not Approved'}
                            </button>
                          </>
                        ) : (
                          // Credit Card: Show only Approve button
                          <button
                            style={{...styles.button, backgroundColor: '#28a745'}}
                            onClick={managerApprove}
                            disabled={!signature || loading}
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