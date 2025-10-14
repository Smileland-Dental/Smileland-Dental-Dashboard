'use client';

import React, { useState, useEffect, useRef } from 'react';
import { collection, getDocs, deleteDoc, doc, setDoc } from 'firebase/firestore';
import { ref, listAll, getDownloadURL, uploadBytes, deleteObject } from 'firebase/storage';
import { db, storage } from '@/lib/firebase.config';

const AdminCreditCardReview = () => {
  const [submissions, setSubmissions] = useState([]);
  const [selectedSubmission, setSelectedSubmission] = useState(null);
  const [loading, setLoading] = useState(false);
  const [signature, setSignature] = useState('');
  const [isDrawing, setIsDrawing] = useState(false);
  const canvasRef = useRef(null);
  const [receiptFiles, setReceiptFiles] = useState([]);
  const [loadingFiles, setLoadingFiles] = useState(false);
  const [savedSignature, setSavedSignature] = useState('');
  const [isSignatureSaved, setIsSignatureSaved] = useState(false);
  const [filterDateFrom, setFilterDateFrom] = useState('');
  const [filterDateTo, setFilterDateTo] = useState('');
  const [filterOffice, setFilterOffice] = useState('');

  // Load all submissions from Firestore
  const loadSubmissions = async () => {
    try {
      setLoading(true);
      
      let retryCount = 0;
      const maxRetries = 3;
      let querySnapshot = null;
      
      while (retryCount < maxRetries && !querySnapshot) {
        try {
          querySnapshot = await getDocs(collection(db, 'credit-card-receipts'));
          console.log('✅ Firestore query successful');
        } catch (firestoreError) {
          retryCount++;
          console.error(`❌ Firestore query error (attempt ${retryCount}/${maxRetries}):`, firestoreError);
          
          if (retryCount < maxRetries) {
            // Wait before retry
            await new Promise(resolve => setTimeout(resolve, 1000 * retryCount));
          } else {
            throw new Error('Failed to load submissions after all retries');
          }
        }
      }
      const submissions = [];
      
      querySnapshot.forEach((doc) => {
        const data = doc.data();
        const totalAmount = data.data.reduce((sum, item) => sum + parseFloat(item.amount || 0), 0);
        
        submissions.push({
          id: doc.id,
          employeeName: data.name,
          cardNumber: data.cardNumber,
          date: data.date || data.data[0]?.date,
          office: data.office || 'N/A',
          submissionId: data.submissionId, // Include submission ID
          purchases: data.data.map(item => ({
            date: item.date,
            vendor: item.vendor,
            reason: item.reason,
            amount: item.amount,
            description: item.description,
            receiptFiles: item.receiptFiles ? item.receiptFiles.split(', ') : []
          })),
          totalAmount: totalAmount.toFixed(2),
          submittedAt: data.date ? new Date(data.date) : (data.createdAt?.toDate() || new Date()),
          lastUpdated: data.lastUpdated?.toDate() || new Date(),
          signatureURL: data.signatureURL,
          signatureSavedAt: data.signatureSavedAt?.toDate(),
          addedOnNumbersChecked: data.addedOnNumbersChecked || false,
          addedOnNumbersCheckedAt: data.addedOnNumbersCheckedAt?.toDate()
        });
      });
      
      // Sort by submission date (newest first)
      submissions.sort((a, b) => b.submittedAt - a.submittedAt);
      setSubmissions(submissions);
    } catch (error) {
      console.error('Error loading submissions:', error);
    } finally {
      setLoading(false);
    }
  };

  // Load receipt files for selected submission
  const loadReceiptFiles = async (submission) => {
    try {
      setLoadingFiles(true);
      const files = [];
      
      // Get all files from receipts folder
      const listRef = ref(storage, 'receipts/');
      const result = await listAll(listRef);
      
      // Filter files that belong to this specific submission
      // Use submission ID for precise matching
      let submissionPrefix = '';
      
      if (submission.submissionId) {
        // Use submission ID for exact matching
        submissionPrefix = `${submission.employeeName}_${submission.cardNumber}_${submission.submissionId}`;
        console.log('🔍 Using submission ID for matching:', submission.submissionId);
      } else {
        // Fallback to name + card number (less precise)
        submissionPrefix = `${submission.employeeName}_${submission.cardNumber}`;
        console.log('⚠️ No submission ID found, using fallback matching');
      }
      
      console.log('🔍 Filtering files for submission:', {
        employeeName: submission.employeeName,
        cardNumber: submission.cardNumber,
        submissionId: submission.submissionId,
        prefix: submissionPrefix
      });
      
      console.log('🔍 Total files in storage:', result.items.length);
      console.log('🔍 All files in storage:');
      result.items.forEach((item, index) => {
        console.log(`🔍   ${index + 1}. ${item.name}`);
      });
      
      for (const itemRef of result.items) {
        console.log('🔍 Checking file:', itemRef.name);
        console.log('🔍 Against prefix:', submissionPrefix);
        
        // Check if file starts with the submission prefix
        const matches = itemRef.name.startsWith(submissionPrefix);
        console.log(`🔍 File "${itemRef.name}" matches prefix "${submissionPrefix}": ${matches}`);
        
        if (matches) {
          try {
            const downloadURL = await getDownloadURL(itemRef);
            files.push({
              name: itemRef.name,
              url: downloadURL,
              fullPath: itemRef.fullPath
            });
            console.log('✅ Found matching file:', itemRef.name);
          } catch (urlError) {
            console.warn('Error getting download URL for file:', itemRef.name, urlError.message);
            // Continue with other files even if one fails
          }
        } else {
          console.log('❌ File does not match prefix:', itemRef.name);
        }
      }
      
      // Group files by purchase number first, then sort within each group
      const filesByPurchase = {};
      
      // Group files by purchase number
      // New filename pattern: Name_CardNumber_Timestamp_purchaseX_SequenceNumber_Filename
      files.forEach(file => {
        console.log('🔍 Processing file:', file.name);
        
        // Split filename and find purchase number
        const parts = file.name.split('_');
        let purchaseNum = 0;
        
        // Look for purchase number in the filename pattern
        // Pattern: Name_CardNumber_Timestamp_purchaseX_SequenceNumber_Filename
        // e.g., "John_1234_2024-01-15T10-30-45-123Z_purchase1_0001759527445729_receipt.jpg"
        // parts[0] = "John", parts[1] = "1234", parts[2] = "2024-01-15T10-30-45-123Z", parts[3] = "purchase1", parts[4] = "0001759527445729", parts[5] = "receipt.jpg"
        
        // More robust purchase number detection for new pattern
        // Pattern: Name_CardNumber_SubmissionID_purchaseX_SequenceNumber_Filename
        // e.g., "Stephanie Johnston_2831_2025-10-06T17-56-25-084Z_xqc54b_purchase2_001759773408730_11.png"
        // parts[0] = "Stephanie", parts[1] = "Johnston", parts[2] = "2831", parts[3] = "2025-10-06T17-56-25-084Z", parts[4] = "xqc54b", parts[5] = "purchase2", parts[6] = "001759773408730", parts[7] = "11.png"
        
        for (let i = 0; i < parts.length; i++) {
          console.log(`🔍 Checking part ${i}: "${parts[i]}"`);
          
          // Check if this part contains "purchase" followed by a number
          const purchaseMatch = parts[i].match(/purchase(\d+)/i);
          if (purchaseMatch) {
            const potentialPurchaseNum = parseInt(purchaseMatch[1], 10);
            console.log(`🔍 Found purchase match: "${purchaseMatch[0]}" -> number: ${potentialPurchaseNum}`);
            if (!isNaN(potentialPurchaseNum) && potentialPurchaseNum > 0) {
              purchaseNum = potentialPurchaseNum;
              console.log(`🔍 Set purchase number to: ${purchaseNum}`);
              break;
            }
          }
        }
        
        console.log('📁 File grouped under purchase:', purchaseNum, 'from parts:', parts);
        console.log('📁 Purchase number found:', purchaseNum, 'for file:', file.name);
        
        if (!filesByPurchase[purchaseNum]) {
          filesByPurchase[purchaseNum] = [];
        }
        filesByPurchase[purchaseNum].push(file);
      });
      
      console.log('📁 Files grouped by purchase:', filesByPurchase);
      console.log('📁 Total files found:', files.length);
      console.log('📁 Purchase groups:', Object.keys(filesByPurchase));
      
      // Log details about each purchase group
      Object.keys(filesByPurchase).forEach(purchaseNum => {
        console.log(`📁 Purchase ${purchaseNum}: ${filesByPurchase[purchaseNum].length} files`);
        filesByPurchase[purchaseNum].forEach(file => {
          console.log(`📁   - ${file.name}`);
        });
      });
      
      // Sort files within each purchase group by sequence number
      Object.keys(filesByPurchase).forEach(purchaseNum => {
        filesByPurchase[purchaseNum].sort((a, b) => {
          const getSequenceNumber = (filename) => {
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
      const sortedFiles = [];
      const purchaseNumbers = Object.keys(filesByPurchase).map(Number).sort((a, b) => a - b);
      
      purchaseNumbers.forEach(purchaseNum => {
        sortedFiles.push(...filesByPurchase[purchaseNum]);
      });
      
      // Replace the original files array with sorted files
      files.length = 0;
      files.push(...sortedFiles);
      
      setReceiptFiles(files);
      console.log('📁 Loaded receipt files:', files.length);
    } catch (error) {
      console.error('Error loading receipt files:', error);
      console.log('📁 CORS error or other issue - continuing without receipt files');
      setReceiptFiles([]);
    } finally {
      setLoadingFiles(false);
    }
  };

  // Load submissions on component mount
  useEffect(() => {
    loadSubmissions();
  }, []);

  // Filter submissions based on date range and office
  const getFilteredSubmissions = () => {
    let filtered = submissions;

    // Filter by date range (using submission date, not purchase date)
    if (filterDateFrom || filterDateTo) {
      filtered = filtered.filter(submission => {
        // Get the submission date in a more reliable way
        let submissionDateStr;
        
        if (submission.submittedAt instanceof Date) {
          // If it's already a Date object
          const year = submission.submittedAt.getFullYear();
          const month = String(submission.submittedAt.getMonth() + 1).padStart(2, '0');
          const day = String(submission.submittedAt.getDate()).padStart(2, '0');
          submissionDateStr = `${year}-${month}-${day}`;
        } else {
          // If it's a string or timestamp, convert to Date first
          const submissionDate = new Date(submission.submittedAt);
          const year = submissionDate.getFullYear();
          const month = String(submissionDate.getMonth() + 1).padStart(2, '0');
          const day = String(submissionDate.getDate()).padStart(2, '0');
          submissionDateStr = `${year}-${month}-${day}`;
        }
        
        console.log('Filtering submission:', {
          submittedAt: submission.submittedAt,
          submissionDateStr,
          filterDateFrom,
          filterDateTo,
          submission: submission.employeeName,
          match: filterDateFrom && filterDateTo ? 
            (submissionDateStr >= filterDateFrom && submissionDateStr <= filterDateTo) :
            (filterDateFrom ? submissionDateStr >= filterDateFrom : submissionDateStr <= filterDateTo)
        });
        
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
      filtered = filtered.filter(submission => 
        submission.office === filterOffice
      );
    }

    console.log('Filtered results:', filtered.length, 'of', submissions.length);
    return filtered;
  };

  // Get unique offices for filter dropdown
  const getUniqueOffices = () => {
    const offices = [...new Set(submissions.map(s => s.office))];
    return offices.filter(office => office && office !== 'N/A');
  };

  // Handle signature drawing
  const startDrawing = (e) => {
    setIsDrawing(true);
    const canvas = canvasRef.current;
    const ctx = canvas.getContext('2d');
    const rect = canvas.getBoundingClientRect();
    
    ctx.beginPath();
    ctx.moveTo(e.clientX - rect.left, e.clientY - rect.top);
  };

  const draw = (e) => {
    if (!isDrawing) return;
    
    const canvas = canvasRef.current;
    const ctx = canvas.getContext('2d');
    const rect = canvas.getBoundingClientRect();
    
    ctx.lineTo(e.clientX - rect.left, e.clientY - rect.top);
    ctx.stroke();
  };

  const stopDrawing = () => {
    setIsDrawing(false);
    const canvas = canvasRef.current;
    setSignature(canvas.toDataURL());
  };

  const clearSignature = () => {
    const canvas = canvasRef.current;
    if (canvas) {
      const ctx = canvas.getContext('2d');
      if (ctx) {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
      }
    }
    setSignature('');
  };

  // Check Added on Numbers (first step)
  const checkAddedOnNumbers = async () => {
    try {
      setLoading(true);
      
      // Update Firestore with checked status
      const docRef = doc(db, 'credit-card-receipts', selectedSubmission.id);
      await setDoc(docRef, {
        addedOnNumbersChecked: true,
        addedOnNumbersCheckedAt: new Date()
      }, { merge: true });
      
      // Update local state
      setSelectedSubmission({
        ...selectedSubmission,
        addedOnNumbersChecked: true,
        addedOnNumbersCheckedAt: new Date()
      });
      
      // Update submissions list
      setSubmissions(prev => prev.map(sub => 
        sub.id === selectedSubmission.id 
          ? { ...sub, addedOnNumbersChecked: true, addedOnNumbersCheckedAt: new Date() }
          : sub
      ));
      
      alert('✅ Added on Numbers checked successfully!');
      
    } catch (error) {
      console.error('Error checking Added on Numbers:', error);
      alert('Error checking Added on Numbers: ' + error.message);
    } finally {
      setLoading(false);
    }
  };

  // Manager Approve - Save signature and generate PDF in one action
  const managerApprove = async () => {
    if (!signature) {
      alert('Please provide a signature before approving.');
      return;
    }
    
    try {
      setLoading(true);
      
      // Convert signature data URL to blob (without fetch to avoid CORS)
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
      const fileName = `signatures/${selectedSubmission.employeeName}_${selectedSubmission.cardNumber}_${timestamp}.png`;
      
      // Upload to Firebase Storage
      const storageRef = ref(storage, fileName);
      await uploadBytes(storageRef, blob);
      const downloadURL = await getDownloadURL(storageRef);
      
      // Save signature data
      setSavedSignature(signature);
      setIsSignatureSaved(true);
      
      // Store signature URL in Firestore for future reference (with retry logic)
      const docRef = doc(db, 'credit-card-receipts', selectedSubmission.id);
      
      let retryCount = 0;
      const maxRetries = 3;
      let firestoreSuccess = false;
      
      while (retryCount < maxRetries && !firestoreSuccess) {
        try {
          await setDoc(docRef, {
            signatureURL: downloadURL,
            signatureSavedAt: new Date()
          }, { merge: true });
          firestoreSuccess = true;
          console.log('✅ Firestore update successful');
        } catch (firestoreError) {
          retryCount++;
          console.error(`❌ Firestore error (attempt ${retryCount}/${maxRetries}):`, firestoreError);
          
          if (retryCount < maxRetries) {
            // Wait before retry
            await new Promise(resolve => setTimeout(resolve, 1000 * retryCount));
          } else {
            console.warn('⚠️ Firestore update failed after all retries, continuing with PDF generation...');
            // Continue without Firestore update - signature is still saved in Storage
          }
        }
      }
      
      // Generate PDF and save to Excel (without opening print dialog)
      await saveToExcel();
      
      alert('✅ Dr. Oh approved! Signature saved and PDF generated. Data saved to Excel file.');
      
      // Close modal and refresh submissions
      setSelectedSubmission(null);
      setSignature('');
      setSavedSignature('');
      setIsSignatureSaved(false);
      setReceiptFiles([]);
      loadSubmissions();
      
    } catch (error) {
      console.error('Error approving submission:', error);
      alert('Error approving submission: ' + error.message);
    } finally {
      setLoading(false);
    }
  };

  // Generate PDF and save to Excel (without opening print dialog)
  const generatePDFAndSaveToExcel = async () => {
    try {
      // Save to Excel file (which will generate PDF and include PDF link)
      await saveToExcel();
      
      // DO NOT delete data automatically - only save to Excel
      console.log('✅ PDF generated and saved to Excel successfully');
    } catch (error) {
      console.error('Error generating PDF and saving to Excel:', error);
      throw error;
    }
  };

  // Generate PDF with saved signature and delete all data
  const generatePDF = async () => {
    console.log('🔵 [generatePDF] Function called');
    // Check if signature exists (either saved in current session or already saved in submission)
    const signatureToUse = savedSignature || selectedSubmission?.signatureURL;
    
    if (!signatureToUse) {
      alert('Please save a signature first before generating PDF.');
      return;
    }

    try {
      setLoading(true);
      
      // Get signature data - prefer current session signature, fallback to saved signature URL
      let signatureData = null;
      
      // Check if we have a current session signature
      if (signatureToUse && signatureToUse.trim() !== '') {
        signatureData = signatureToUse;
        console.log('📄 Using current session signature');
      }
      // If no current session signature, try to get from saved signature URL
      else if (selectedSubmission?.signatureURL && selectedSubmission.signatureURL.trim() !== '') {
        try {
          console.log('📄 Converting signature URL to base64 data...');
          const response = await fetch(selectedSubmission.signatureURL);
          const blob = await response.blob();
          const reader = new FileReader();
          signatureData = await new Promise((resolve) => {
            reader.onload = () => resolve(reader.result);
            reader.readAsDataURL(blob);
          });
          console.log('📄 Converted signature URL to base64 data');
        } catch (error) {
          console.error('📄 Error converting signature URL:', error);
          signatureData = selectedSubmission.signatureURL; // Fallback to URL
          console.log('📄 Using signature URL as fallback');
        }
      } else {
        console.log('📄 No signature data available');
      }
      
      console.log('📄 Final signature data type:', typeof signatureData);
      console.log('📄 Final signature data length:', signatureData ? signatureData.length : 0);
      console.log('📄 Final signature data start:', signatureData ? signatureData.substring(0, 50) + '...' : 'None');
      console.log('📄 Final signature data value:', signatureData);
      console.log('📄 About to send to API:', {
        signature: signatureData,
        signatureType: typeof signatureData,
        isNull: signatureData === null,
        isUndefined: signatureData === undefined,
        isString: typeof signatureData === 'string',
        isEmpty: signatureData === '',
        isMissing: signatureData === 'Missing'
      });
      
      // Generate PDF
      const pdfResponse = await fetch('/api/generate-credit-card-pdf', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          name: selectedSubmission.employeeName,
          cardNumber: selectedSubmission.cardNumber,
          date: selectedSubmission.date,
          office: selectedSubmission.office,
          purchases: selectedSubmission.purchases,
          filesData: receiptFiles, // Include receipt files with URLs
          signature: signatureData
        })
      });

      const pdfResult = await pdfResponse.json();
      
      if (pdfResult.success) {
        // PDF generated successfully (no print dialog)
        console.log('✅ PDF generated successfully');
        
        // Save to Excel file (optional)
        try {
          await saveToExcel();
        } catch (error) {
          console.log('Excel save failed, continuing with PDF generation');
        }
        
        // DO NOT delete data automatically - only generate PDF
        
      } else {
        throw new Error(pdfResult.error || 'PDF generation failed');
      }
    } catch (error) {
      console.error('Error generating PDF:', error);
      alert('Error generating PDF: ' + error.message);
    } finally {
      setLoading(false);
    }
  };

  // Download CSV file
  const downloadExcelFile = async () => {
    try {
      setLoading(true);
      
      // Download the main CSV file
      const mainFileName = 'credit-card-receipts.csv';
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
        alert('❌ CSV file not found. Please generate a PDF first to create the CSV file.');
      }
    } catch (error) {
      console.error('Error downloading CSV file:', error);
      alert('Error downloading CSV file: ' + error.message);
    } finally {
      setLoading(false);
    }
  };

  // Save data to Excel file (client-side)
  const saveToExcel = async () => {
    console.log('🟡 [saveToExcel] Function called');
    
    // Check if submission has signature (either saved in Firestore or current session)
    const hasSignatureURL = selectedSubmission?.signatureURL && selectedSubmission.signatureURL.trim() !== '';
    const hasCurrentSignature = (savedSignature && savedSignature.trim() !== '') || (signature && signature.trim() !== '');
    
    console.log('📊 Signature check:', {
      hasSignatureURL,
      hasCurrentSignature,
      signatureURL: selectedSubmission?.signatureURL,
      savedSignature: savedSignature ? 'Present' : 'Missing',
      currentSignature: signature ? 'Present' : 'Missing'
    });
    
    if (!hasSignatureURL && !hasCurrentSignature) {
      alert('❌ This submission has no signature. Please sign first before generating PDF.');
      return;
    }
    
    try {
      const totalAmount = selectedSubmission.purchases.reduce((sum, purchase) => {
        return sum + (parseFloat(purchase.amount) || 0);
      }, 0);

      // Generate PDF and save to Firebase Storage first
      // Build filesData from selectedSubmission data instead of receiptFiles state
      const filesData = [];
      for (const purchase of selectedSubmission.purchases) {
        if (purchase.receiptFiles) {
          let receiptFiles = [];
          
          // Handle different data types
          if (typeof purchase.receiptFiles === 'string') {
            receiptFiles = purchase.receiptFiles.split(', ');
          } else if (Array.isArray(purchase.receiptFiles)) {
            receiptFiles = purchase.receiptFiles;
          }
          
          for (const fileName of receiptFiles) {
            if (fileName && fileName.trim()) {
              try {
                const fileRef = ref(storage, `receipts/${fileName.trim()}`);
                const downloadURL = await getDownloadURL(fileRef);
                filesData.push({
                  name: fileName.trim(),
                  url: downloadURL,
                  fullPath: `receipts/${fileName.trim()}`
                });
              } catch (error) {
                console.warn(`⚠️ Could not get download URL for ${fileName}:`, error);
              }
            }
          }
        }
      }
      
      console.log('📄 Generating PDF with receipt files:', filesData.length);
      console.log('📄 Files data:', filesData);
      console.log('📄 Saved signature:', savedSignature);
      console.log('📄 Saved signature type:', typeof savedSignature);
      console.log('📄 Selected submission signature URL:', selectedSubmission?.signatureURL);
      console.log('📄 Selected submission data:', {
        hasSignatureURL: !!selectedSubmission?.signatureURL,
        signatureURL: selectedSubmission?.signatureURL,
        hasSignatureSavedAt: !!selectedSubmission?.signatureSavedAt,
        signatureSavedAt: selectedSubmission?.signatureSavedAt
      });
      
      // Get signature data - prefer current session signature, fallback to saved signature URL
      let signatureData = null;
      
      // Check if we have a current session signature (prefer signature over savedSignature)
      if (signature && signature.trim() !== '') {
        signatureData = signature;
        console.log('📄 Using current session signature (signature)');
      }
      else if (savedSignature && savedSignature.trim() !== '') {
        signatureData = savedSignature;
        console.log('📄 Using current session signature (savedSignature)');
      }
      // If no current session signature, try to get from saved signature URL
      else if (selectedSubmission?.signatureURL && selectedSubmission.signatureURL.trim() !== '') {
        try {
          console.log('📄 Converting signature URL to base64 data...');
          const response = await fetch(selectedSubmission.signatureURL);
          const blob = await response.blob();
          const reader = new FileReader();
          signatureData = await new Promise((resolve) => {
            reader.onload = () => resolve(reader.result);
            reader.readAsDataURL(blob);
          });
          console.log('📄 Converted signature URL to base64 data');
        } catch (error) {
          console.error('📄 Error converting signature URL:', error);
          signatureData = selectedSubmission.signatureURL; // Fallback to URL
          console.log('📄 Using signature URL as fallback');
        }
      } else {
        console.log('📄 No signature data available');
      }
      
      console.log('📄 Final signature data type:', typeof signatureData);
      console.log('📄 Final signature data length:', signatureData ? signatureData.length : 0);
      console.log('📄 Final signature data start:', signatureData ? signatureData.substring(0, 50) + '...' : 'None');
      console.log('📄 Final signature data value:', signatureData);
      console.log('📄 About to send to API:', {
        signature: signatureData,
        signatureType: typeof signatureData,
        isNull: signatureData === null,
        isUndefined: signatureData === undefined,
        isString: typeof signatureData === 'string',
        isEmpty: signatureData === '',
        isMissing: signatureData === 'Missing'
      });
      
      const pdfResponse = await fetch('/api/generate-credit-card-pdf', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          name: selectedSubmission.employeeName,
          cardNumber: selectedSubmission.cardNumber,
          date: selectedSubmission.date,
          office: selectedSubmission.office,
          purchases: selectedSubmission.purchases,
          filesData: filesData, // Built from selectedSubmission data
          signature: signatureData
        })
      });

      console.log('📄 PDF API response status:', pdfResponse.status);

      const pdfResult = await pdfResponse.json();
      let pdfDownloadURL = '';

      console.log('📄 PDF generation result:', pdfResult);
      console.log('📄 PDF success status:', pdfResult.success);
      console.log('📄 PDF error message:', pdfResult.error);

      if (pdfResult.success) {
        // Save PDF to Firebase Storage (with error handling)
        try {
          const pdfBlob = new Blob([pdfResult.html], { type: 'text/html' });
          const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
          const pdfFileName = `pdfs/${selectedSubmission.employeeName}_${selectedSubmission.cardNumber}_${timestamp}.html`;
          const pdfRef = ref(storage, pdfFileName);
          await uploadBytes(pdfRef, pdfBlob);
          pdfDownloadURL = await getDownloadURL(pdfRef);
          console.log('📄 PDF saved to Firebase Storage:', pdfDownloadURL);
          console.log('📄 PDF download URL set to:', pdfDownloadURL);
          
          // Save PDF URL to Firestore for future Excel generation
          try {
            const docRef = doc(db, 'credit-card-receipts', selectedSubmission.id);
            await setDoc(docRef, {
              pdfURL: pdfDownloadURL,
              pdfGeneratedAt: new Date()
            }, { merge: true });
            console.log('📄 PDF URL saved to Firestore');
          } catch (firestoreError) {
            console.warn('⚠️ Could not save PDF URL to Firestore:', firestoreError);
          }
        } catch (error) {
          console.error('Error saving PDF to Storage (non-critical):', error);
          // Continue without PDF link if storage fails
          pdfDownloadURL = '';
          console.log('📄 PDF download URL set to empty due to storage error');
        }
      } else {
        console.error('PDF generation failed:', pdfResult.error);
        // Use signature URL as fallback if PDF generation fails
        pdfDownloadURL = selectedSubmission?.signatureURL || '';
        console.log('📄 PDF download URL set to signature URL as fallback:', pdfDownloadURL);
      }

      // Build cumulative CSV from Firestore data (avoid CORS issues)
      console.log('📊 Building cumulative CSV from Firestore data...');
      console.log('📊 Current submission:', {
        id: selectedSubmission.id,
        employeeName: selectedSubmission.employeeName,
        office: selectedSubmission.office,
        cardNumber: selectedSubmission.cardNumber,
        purchases: selectedSubmission.purchases.length
      });
      
      // Define CSV file reference
      const mainFileName = 'credit-card-receipts.csv';
      const mainFileRef = ref(storage, `excel/${mainFileName}`);
      
      let existingData = [
        ['Employee Name', 'Office', 'Card Number', 'Purchase Date', 'Store/Website', 'Reason', 'Amount', 'Account Description', 'Total Amount', 'Submission Date', 'Status', 'PDF Link']
      ];
      
      try {
        // Get all submissions that have been processed (have signatureURL)
        const querySnapshot = await getDocs(collection(db, 'credit-card-receipts'));
        const processedSubmissions = [];
        
        querySnapshot.forEach((doc) => {
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
              pdfURL: data.pdfURL || ''
            });
          }
        });
        
        console.log('📊 Found', processedSubmissions.length, 'processed submissions in Firestore');
        
        // Add all processed submissions to CSV data (excluding current submission)
        for (const submission of processedSubmissions) {
          // Skip current submission to avoid duplication
          if (submission.id === selectedSubmission.id) {
            continue;
          }
          
          const totalAmount = submission.purchases.reduce((sum, purchase) => {
            return sum + (parseFloat(purchase.amount) || 0);
          }, 0);
          
            // Add each purchase as a row
            submission.purchases.forEach((purchase, index) => {
              existingData.push([
                index === 0 ? submission.employeeName : '',
                index === 0 ? submission.office : '',
                index === 0 ? `****${submission.cardNumber}` : '',
                purchase.date,
                purchase.vendor,
                purchase.reason,
                `$${parseFloat(purchase.amount).toFixed(2)}`,
                purchase.description,
                index === 0 ? `$${totalAmount.toFixed(2)}` : '',
                submission.signatureSavedAt ? submission.signatureSavedAt.toDate().toLocaleDateString() : new Date().toLocaleDateString(),
                'Approved & PDF Generated',
                index === 0 ? (submission.pdfURL || '') : ''
              ]);
            });
        }
        
        console.log('📊 Built cumulative CSV with', existingData.length, 'total rows');
        
      } catch (error) {
        console.error('Error building cumulative CSV:', error);
        console.log('📊 Using header-only CSV due to error');
      }

      // Add new data to existing data
      console.log('📊 Creating new rows with PDF download URL:', pdfDownloadURL);
      const newRows = selectedSubmission.purchases.map((purchase, index) => [
        index === 0 ? selectedSubmission.employeeName : '',
        index === 0 ? selectedSubmission.office : '',
        index === 0 ? `****${selectedSubmission.cardNumber}` : '',
        purchase.date,
        purchase.vendor,
        purchase.reason,
        `$${parseFloat(purchase.amount).toFixed(2)}`,
        purchase.description,
        index === 0 ? `$${totalAmount.toFixed(2)}` : '',
        new Date().toLocaleDateString(),
        'Approved & PDF Generated',
        index === 0 ? pdfDownloadURL : '' // PDF URL
      ]);

      // Combine existing and new data
      const updatedData = [...existingData, ...newRows];
      
      console.log('📊 Before combining:');
      console.log('📊 Existing data rows:', existingData.length);
      console.log('📊 New data rows:', newRows.length);
      console.log('📊 Total after combining:', updatedData.length);
      console.log('📊 New rows content:', newRows);
      console.log('📊 PDF URL in new rows:', newRows[0] ? newRows[0][11] : 'No rows');

      // Convert to CSV string (simple version)
      const csvString = updatedData.map(row => row.join(',')).join('\n');
      
      // Create blob and upload to Firebase Storage (like credit-card-receipts.tsx)
      try {
        const blob = new Blob([csvString], { type: 'text/csv' });
        console.log('📊 Uploading cumulative CSV file...');
        
        // Simple upload like credit-card-receipts.tsx
        await uploadBytes(mainFileRef, blob);
        
        console.log('✅ Cumulative CSV file updated successfully!');
        console.log('📊 Total rows in file:', updatedData.length);
        console.log('📊 Added', newRows.length, 'new rows for submission:', selectedSubmission.employeeName);
        console.log('📄 PDF link included:', pdfDownloadURL ? 'Yes' : 'No');
        console.log('📊 This CSV now contains ALL processed submissions');
      } catch (error) {
        console.error('Error uploading CSV file:', error);
        // Continue execution even if CSV upload fails
      }
      
    } catch (error) {
      console.error('Error saving CSV file:', error);
      // CSV 저장 실패해도 PDF 생성은 계속 진행
    }
  };

  // Delete completed submissions only (Reset function)
  const resetAllData = async () => {
    if (!confirm('⚠️ WARNING: This will delete ONLY completed submissions (signed and numbers checked). Incomplete submissions will remain. Are you sure you want to continue?')) {
      return;
    }
    
    try {
      setLoading(true);
      
      // Get all submissions
      const querySnapshot = await getDocs(collection(db, 'credit-card-receipts'));
      let completedCount = 0;
      let incompleteCount = 0;
      
      // First, collect completed submissions and delete their Storage files
      const completedSubmissions = [];
      const deletePromises = [];
      
      querySnapshot.forEach((doc) => {
        const data = doc.data();
        const isSigned = data.signatureURL && data.signatureURL.trim() !== '';
        const isNumbersChecked = data.addedOnNumbersChecked === true;
        
        if (isSigned && isNumbersChecked) {
          // Both signature and numbers check are complete - safe to delete
          completedSubmissions.push({
            id: doc.id,
            employeeName: data.name,
            cardNumber: data.cardNumber,
            data: data.data,
            docRef: doc.ref
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
      console.log('🗑️ Deleting Storage files for completed submissions only...');
      
      try {
        console.log(`🗑️ Found ${completedSubmissions.length} completed submissions to clean up files for`);
        
        // Delete files for each completed submission
        for (const submission of completedSubmissions) {
          try {
            // Delete receipt files for this submission (now safe to delete since PDF uses Base64)
            console.log(`🗑️ Processing submission: ${submission.employeeName} (${submission.cardNumber})`);
            console.log(`🗑️ Submission data:`, submission.data);
            
            for (const purchase of submission.data) {
              console.log(`🗑️ Processing purchase:`, purchase);
              
              if (purchase.receiptFiles) {
                let receiptFiles = [];
                
                // Handle different data types
                if (typeof purchase.receiptFiles === 'string') {
                  receiptFiles = purchase.receiptFiles.split(', ');
                } else if (Array.isArray(purchase.receiptFiles)) {
                  receiptFiles = purchase.receiptFiles;
                }
                
                console.log(`🗑️ Receipt files to delete:`, receiptFiles);
                
                for (const fileName of receiptFiles) {
                  if (fileName && fileName.trim()) {
                    const receiptRef = ref(storage, `receipts/${fileName.trim()}`);
                    try {
                      await deleteObject(receiptRef);
                      console.log(`🗑️ Deleted receipt: ${fileName}`);
                    } catch (error) {
                      console.warn(`⚠️ Could not delete receipt ${fileName}:`, error);
                    }
                  }
                }
              } else {
                console.log(`🗑️ No receipt files found for purchase:`, purchase);
              }
            }
            
            // Delete signature file for this submission
            try {
              console.log(`🗑️ Looking for signature files for ${submission.employeeName}_${submission.cardNumber}_`);
              const signaturesRef = ref(storage, 'signatures/');
              const signaturesList = await listAll(signaturesRef);
              console.log(`🗑️ Found ${signaturesList.items.length} signature files in storage`);
              
              for (const item of signaturesList.items) {
                console.log(`🗑️ Checking signature file: ${item.name}`);
                if (item.name.includes(`${submission.employeeName}_${submission.cardNumber}_`)) {
                  await deleteObject(item);
                  console.log(`🗑️ Deleted signature: ${item.name}`);
                }
              }
            } catch (error) {
              console.warn(`⚠️ Could not delete signature for ${submission.employeeName}:`, error);
            }
            
            // Keep PDF files for Excel links to work
            console.log(`📄 Keeping PDF files for ${submission.employeeName} - Excel links will remain functional`);
            
          } catch (submissionError) {
            console.warn(`⚠️ Error deleting files for submission ${submission.employeeName}:`, submissionError);
          }
        }
        
        // Delete Excel file (this is cumulative, so safe to delete)
        const excelRef = ref(storage, 'excel/');
        const excelList = await listAll(excelRef);
        for (const item of excelList.items) {
          await deleteObject(item);
          console.log(`🗑️ Deleted Excel: ${item.name}`);
        }
        
      } catch (storageError) {
        console.warn('⚠️ Error deleting Storage files:', storageError);
        // Continue even if some Storage files fail to delete
      }
      
      // Now delete Firestore documents
      for (const submission of completedSubmissions) {
        deletePromises.push(deleteDoc(submission.docRef));
      }
      
      await Promise.all(deletePromises);
      
      console.log(`Successfully deleted ${completedCount} completed submissions and associated Storage files`);
      
      // Reset state
      setSubmissions([]);
      setSelectedSubmission(null);
      setSignature('');
      setSavedSignature('');
      setIsSignatureSaved(false);
      setReceiptFiles([]);
      
      alert(`✅ Reset completed! Deleted ${completedCount} completed submissions. ${incompleteCount} incomplete submissions were kept.`);
      
    } catch (error) {
      console.error('Error resetting data:', error);
      alert('Error resetting data: ' + error.message);
    } finally {
      setLoading(false);
    }
  };

  // Delete single submission data (but keep Excel file)
  const deleteAllSubmissionData = async () => {
    try {
      // Delete from Firestore
      await deleteDoc(doc(db, 'credit-card-receipts', selectedSubmission.id));
      
      // Delete receipt files from Storage
      for (const file of receiptFiles) {
        try {
          const fileRef = ref(storage, file.fullPath);
          await deleteObject(fileRef);
        } catch (error) {
          console.error('Error deleting file:', file.name, error);
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
            console.log('Signature file deleted:', filePath);
          }
        } catch (error) {
          console.error('Error deleting signature file (non-critical):', error);
          // Continue execution even if signature deletion fails
        }
      }
      
      // Note: Excel file is NOT deleted - it accumulates all submissions
      console.log('✅ Submission data deleted, but Excel file preserved for record keeping');
      
      alert('✅ PDF generated and data saved to Excel! Submission removed from review list.');
      setSelectedSubmission(null);
      setSignature('');
      setSavedSignature('');
      setIsSignatureSaved(false);
      setReceiptFiles([]);
      loadSubmissions(); // Reload submissions
      
    } catch (error) {
      console.error('Error deleting submission data:', error);
      alert('Error deleting data: ' + error.message);
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
      backgroundColor: 'white'
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
          <h1 style={styles.title}>🏢 Company Credit Card Receipts</h1>
          
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
                {getUniqueOffices().map(office => (
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
            {getFilteredSubmissions().map((submission) => (
              <div
                key={submission.id}
                style={submission.signatureURL ? styles.submissionCardSigned : styles.submissionCard}
                onClick={() => {
                  setSelectedSubmission(submission);
                  loadReceiptFiles(submission);
                }}
              >
                <div style={styles.submissionHeader}>
                  <div style={styles.employeeName}>
                    {submission.employeeName}
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
                    Submitted: {submission.date}
                  </div>
                </div>
                <div style={styles.submissionDetails}>
                  <p><strong>Office:</strong> {submission.office}</p>
                  <p><strong>Card:</strong> ****{submission.cardNumber}</p>
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
                      ✅ Dr. Oh has signed this submission
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
                <p><strong>Card Number:</strong> ****{selectedSubmission.cardNumber}</p>
                <p><strong>Submission Date:</strong> {selectedSubmission.date}</p>

                <div style={{display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '20px'}}>
                  <h3 style={{margin: 0}}>Purchase Details</h3>
                  {!selectedSubmission.addedOnNumbersChecked && (
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
                      <th style={styles.tableHeader}>Account</th>
                    </tr>
                  </thead>
                  <tbody>
                    {selectedSubmission.purchases.map((purchase, index) => (
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
                        <td style={styles.tableCell}>{purchase.description}</td>
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
                        const filesByPurchase = {};
                        receiptFiles.forEach(file => {
                          // Extract purchase number from new filename pattern: Name_CardNumber_Timestamp_purchaseX_SequenceNumber_Filename
                          // e.g., "John_1234_2024-01-15T10-30-45-123Z_purchase1_0001759527445729_receipt.jpg" -> Purchase 1
                          const parts = file.name.split('_');
                          let purchaseNum = 0;
                          
                          // Find the purchase number by looking for "purchaseX" pattern
                          // Pattern: Name_CardNumber_SubmissionID_purchaseX_SequenceNumber_Filename
                          for (let i = 0; i < parts.length; i++) {
                            console.log(`🔍 UI Checking part ${i}: "${parts[i]}"`);
                            
                            // Check if this part contains "purchase" followed by a number
                            const purchaseMatch = parts[i].match(/purchase(\d+)/i);
                            if (purchaseMatch) {
                              const potentialPurchaseNum = parseInt(purchaseMatch[1], 10);
                              console.log(`🔍 UI Found purchase match: "${purchaseMatch[0]}" -> number: ${potentialPurchaseNum}`);
                              if (!isNaN(potentialPurchaseNum) && potentialPurchaseNum > 0) {
                                purchaseNum = potentialPurchaseNum;
                                console.log(`🔍 UI Set purchase number to: ${purchaseNum}`);
                                break;
                              }
                            }
                          }
                          
                          console.log('🔍 UI Processing file:', file.name, '-> Purchase:', purchaseNum);
                          
                          if (!filesByPurchase[purchaseNum]) {
                            filesByPurchase[purchaseNum] = [];
                          }
                          filesByPurchase[purchaseNum].push(file);
                        });
                        
                        console.log('🔍 UI Files grouped by purchase:', filesByPurchase);
                        
                        // Sort purchase numbers and render each group
                        const purchaseNumbers = Object.keys(filesByPurchase).map(Number).sort((a, b) => a - b);
                        
                        return purchaseNumbers.map(purchaseNum => (
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
                              {filesByPurchase[purchaseNum].map((file, index) => (
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
                                      onError={(e) => {
                                        e.target.style.display = 'none';
                                        e.target.nextSibling.style.display = 'block';
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
                  <p>Please sign below to approve this submission:</p>
                  
                  {/* Show saved signature if exists */}
                  {selectedSubmission.signatureURL && (
                    <div style={{marginBottom: '20px', padding: '15px', backgroundColor: '#e8f5e8', borderRadius: '8px', border: '2px solid #28a745'}}>
                      <h4 style={{margin: '0 0 10px 0', color: '#155724'}}>✅ Saved Signature</h4>
                      <img 
                        src={selectedSubmission.signatureURL} 
                        alt="Saved Signature" 
                        style={{
                          maxWidth: '300px',
                          maxHeight: '100px',
                          border: '1px solid #ddd',
                          backgroundColor: 'white',
                          borderRadius: '4px'
                        }}
                      />
                      <p style={{margin: '10px 0 0 0', fontSize: '12px', color: '#666'}}>
                        Saved on: {selectedSubmission.signatureSavedAt ? 
                          (selectedSubmission.signatureSavedAt.toDate ? 
                            new Date(selectedSubmission.signatureSavedAt.toDate()).toLocaleString() : 
                            new Date(selectedSubmission.signatureSavedAt).toLocaleString()) : 
                          'Unknown'}
                      </p>
                    </div>
                  )}
                  
                  {/* Only show canvas if not already signed */}
                  {!selectedSubmission.signatureURL && (
                    <canvas
                      ref={canvasRef}
                      width={600}
                      height={200}
                      style={styles.signatureCanvas}
                      onMouseDown={startDrawing}
                      onMouseMove={draw}
                      onMouseUp={stopDrawing}
                      onMouseLeave={stopDrawing}
                    />
                  )}
                  
                  <div style={styles.signatureButtons}>
                    {selectedSubmission.signatureURL ? (
                      // Already signed - show approved status
                      <div style={{
                        padding: '15px',
                        backgroundColor: '#d4edda',
                        borderRadius: '8px',
                        border: '2px solid #28a745',
                        textAlign: 'center'
                      }}>
                        <h4 style={{margin: '0 0 10px 0', color: '#155724'}}>✅ Approved by Dr. Oh</h4>
                        <p style={{margin: '0', color: '#155724', fontSize: '14px'}}>
                          This submission has been approved and PDF has been generated.
                        </p>
                      </div>
                    ) : (
                      // Not signed yet - show signature buttons
                      <>
                        <button
                          style={styles.buttonSecondary}
                          onClick={clearSignature}
                        >
                          Clear Signature
                        </button>
                        <button
                          style={{...styles.button, backgroundColor: '#28a745'}}
                          onClick={managerApprove}
                          disabled={!signature || loading}
                        >
                          {loading ? 'Approving...' : '✅ Dr. Oh Approve'}
                        </button>
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
