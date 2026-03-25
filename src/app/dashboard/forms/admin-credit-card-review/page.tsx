'use client';

import React, { useState, useEffect } from 'react';
import { Document, Image, Page, StyleSheet, Text, View, pdf } from '@react-pdf/renderer';
import { collection, getDocs, deleteDoc, doc, getDoc, setDoc } from 'firebase/firestore';
import { ref, listAll, getDownloadURL, uploadBytes, deleteObject } from 'firebase/storage';
import { onAuthStateChanged } from 'firebase/auth';
import { db, storage, auth } from '@/lib/firebase.config';
// Simple helper functions to replace security-client
const sanitizeCSVCell = (value: string | number | undefined | null): string => {
  if (value === null || value === undefined) return '';
  const str = String(value);
  // Basic CSV injection protection: remove leading =, +, -, @
  if (str.startsWith('=') || str.startsWith('+') || str.startsWith('-') || str.startsWith('@')) {
    return "'" + str;
  }
  // Escape quotes and wrap in quotes if contains comma, newline, or quote
  if (str.includes(',') || str.includes('\n') || str.includes('"')) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
};

const sanitizeFirebaseDataClient = (data: any): any => {
  // 🔒 보안: Firebase에 저장되는 데이터 sanitization
  if (!data || typeof data !== 'object') {
    return data;
  }
  
  const sanitized: any = {};
  for (const [key, value] of Object.entries(data)) {
    // 키 이름 검증 (알파벳, 숫자, 언더스코어만 허용)
    if (!/^[a-zA-Z0-9_]+$/.test(key)) {
      continue; // Skip invalid keys
    }
    
    // 값 타입별 sanitization
    if (value === null || value === undefined) {
      sanitized[key] = value;
    } else if (typeof value === 'string') {
      // 문자열: 길이 제한 및 특수문자 제거
      sanitized[key] = value.substring(0, 10000).replace(/[<>]/g, '');
    } else if (typeof value === 'number') {
      // 숫자: NaN 및 Infinity 체크
      sanitized[key] = isFinite(value) ? value : 0;
    } else if (value instanceof Date) {
      sanitized[key] = value;
    } else if (typeof value === 'boolean') {
      sanitized[key] = value;
    } else if (Array.isArray(value)) {
      // 배열: 최대 1000개 항목만 허용
      sanitized[key] = value.slice(0, 1000);
    } else if (typeof value === 'object') {
      // 객체: 재귀적으로 sanitize (최대 깊이 3)
      sanitized[key] = sanitizeFirebaseDataClient(value);
    }
  }
  
  return sanitized;
};

/** CSV "Submission Date" column: prefer createdAt, then date string, then pdfGeneratedAt */
const getSubmissionDateDisplayForCsv = (data: any): string => {
  if (data.createdAt?.toDate && typeof data.createdAt.toDate === 'function') {
    return data.createdAt.toDate().toLocaleDateString();
  }
  if (data.createdAt instanceof Date) {
    return data.createdAt.toLocaleDateString();
  }
  if (data.date && String(data.date).trim()) {
    return String(data.date).trim();
  }
  if (data.pdfGeneratedAt?.toDate && typeof data.pdfGeneratedAt.toDate === 'function') {
    return data.pdfGeneratedAt.toDate().toLocaleDateString();
  }
  return new Date().toLocaleDateString();
};

const ISO_CALENDAR_DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

/** YYYY-MM-DD in America/Los_Angeles */
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

/** Date range filter: ISO `date` when valid, else LA calendar date from `submittedAt`. */
function getSubmissionDateStringForFilter(submission: { date: string; submittedAt: Date }): string {
  if (submission.date && ISO_CALENDAR_DATE_RE.test(submission.date)) {
    return submission.date;
  }
  const at = submission.submittedAt instanceof Date ? submission.submittedAt : new Date(submission.submittedAt);
  return formatDatePacificLosAngeles(at);
}

/** List / modal: show stored `date` when non-empty, else LA YYYY-MM-DD from `submittedAt`. */
function getSubmissionDateDisplayLabel(submission: { date: string; submittedAt: Date }): string {
  if (submission.date?.trim()) {
    return submission.date.trim();
  }
  const at = submission.submittedAt instanceof Date ? submission.submittedAt : new Date(submission.submittedAt);
  return formatDatePacificLosAngeles(at);
}

/**
 * Firestore 문서가 워크플로우상 "완료"인지: PDF URL이 있고 Added on Numbers가 체크됨.
 * (CSV 다운로드 대상, Reset 시 삭제 대상 등)
 */
function isFirestoreDocPdfWorkflowComplete(data: {
  pdfURL?: unknown;
  addedOnNumbersChecked?: unknown;
}): boolean {
  const hasPDF = typeof data.pdfURL === 'string' && data.pdfURL.trim() !== '';
  const numbersChecked = data.addedOnNumbersChecked === true;
  return hasPDF && numbersChecked;
}

// Interfaces for type safety
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
  signed?: boolean; // Boolean flag to indicate if signature has been provided
  addedOnNumbersChecked?: boolean;
  addedOnNumbersCheckedAt?: Date;
  formType?: 'credit-card' | 'reimbursement'; // Form type to distinguish between credit card receipts and reimbursement requests
  amountAdjustedTo?: string;
  reasonForAdjustment?: string;
  approved?: boolean;
  pdfURL?: string; // URL of the generated PDF file
}

interface ReceiptFile {
  name: string;
  url: string;
}

// --- PDF generation (client-side): credit card + reimbursement ---
interface PdfFileData {
  name: string;
  url: string;
}

interface PdfSafePurchase {
  date: string;
  vendor: string;
  reason: string;
  amount: number;
  description: string;
}

const pdfEscapeHtml = (str: string): string => {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
};

const pdfSanitizeAmount = (amount: string | number): number => {
  const num = typeof amount === 'string' ? parseFloat(amount) : amount;
  return isNaN(num) || num < 0 ? 0 : Math.min(num, 1000000);
};

const pdfSanitizeString = (str: string | undefined | null, maxLength: number = 1000): string => {
  if (!str) return '';
  const sanitized = String(str)
    .replace(/[<>]/g, '')
    .substring(0, maxLength)
    .trim();
  return pdfEscapeHtml(sanitized);
};

function formatPdfStatusDateTimeNow(): string {
  return new Date().toLocaleString('en-US', { dateStyle: 'medium', timeStyle: 'medium' });
}

const PDF_A4_WIDTH = 595.28;
const PDF_A4_HEIGHT = 841.89;
const PDF_IMAGE_PAGE_PADDING = 24;

const pdfDocumentStyles = StyleSheet.create({
  page: {
    paddingTop: 24,
    paddingBottom: 24,
    paddingHorizontal: 24,
    fontSize: 10,
    color: '#000000',
  },
  header: {
    textAlign: 'center',
    marginBottom: 20,
  },
  headerTitle: {
    fontSize: 16,
    fontWeight: 'bold',
    textTransform: 'uppercase',
    letterSpacing: 1,
  },
  headerLine: {
    borderBottomWidth: 2,
    borderBottomColor: '#000000',
    marginTop: 8,
    marginBottom: 8,
  },
  infoSection: {
    marginBottom: 16,
  },
  infoRow: {
    flexDirection: 'row',
    marginBottom: 6,
    alignItems: 'flex-end',
  },
  infoLabel: {
    width: 120,
    fontWeight: 'bold',
  },
  infoValue: {
    flex: 1,
    borderBottomWidth: 1,
    borderBottomColor: '#000000',
    paddingBottom: 2,
  },
  sectionTitle: {
    fontSize: 12,
    fontWeight: 'bold',
    marginTop: 8,
    marginBottom: 8,
  },
  table: {
    borderWidth: 1,
    borderColor: '#000000',
    marginBottom: 12,
  },
  tableHeader: {
    flexDirection: 'row',
    backgroundColor: '#f5f5f5',
    borderBottomWidth: 1,
    borderBottomColor: '#000000',
  },
  tableRow: {
    flexDirection: 'row',
    borderBottomWidth: 1,
    borderBottomColor: '#000000',
  },
  lastTableRow: {
    borderBottomWidth: 0,
  },
  cell: {
    paddingVertical: 6,
    paddingHorizontal: 4,
    borderRightWidth: 1,
    borderRightColor: '#000000',
  },
  lastCell: {
    borderRightWidth: 0,
  },
  cellText: {
    fontSize: 9,
  },
  totalSection: {
    marginTop: 8,
    marginBottom: 16,
    textAlign: 'right',
    fontSize: 12,
    fontWeight: 'bold',
  },
  imagePage: {
    paddingTop: 24,
    paddingBottom: 24,
    paddingHorizontal: 24,
    justifyContent: 'center',
    alignItems: 'center',
  },
  imagePageImageWrap: {
    width: PDF_A4_WIDTH - PDF_IMAGE_PAGE_PADDING * 2,
    height: PDF_A4_HEIGHT - PDF_IMAGE_PAGE_PADDING * 2,
    borderWidth: 1,
    borderColor: '#cccccc',
    padding: 8,
    justifyContent: 'center',
    alignItems: 'center',
  },
  imagePageImage: {
    width: '100%',
    height: '100%',
    objectFit: 'contain',
  },
});

function buildCreditCardReceiptDocument(params: {
  safeOffice: string;
  safeName: string;
  safeCardNumber: string;
  safeDate: string;
  safePurchases: PdfSafePurchase[];
  totalAmount: number;
  base64FilesData: PdfFileData[];
  approved: boolean;
  safeStatusDecidedAt: string;
}) {
  const {
    safeOffice,
    safeName,
    safeCardNumber,
    safeDate,
    safePurchases,
    totalAmount,
    base64FilesData,
    approved,
    safeStatusDecidedAt,
  } = params;

  const statusLabel = approved ? 'Approved' : 'Not Approved';

  const widths = ['8%', '14%', '22%', '26%', '14%', '16%'];

  const summaryPage = React.createElement(
    Page,
    { size: 'A4', style: pdfDocumentStyles.page, wrap: true },
    React.createElement(
      View,
      { style: pdfDocumentStyles.header },
      React.createElement(Text, { style: pdfDocumentStyles.headerTitle }, `${safeOffice || 'Company'} Credit Card Receipt`),
      React.createElement(View, { style: pdfDocumentStyles.headerLine })
    ),
    React.createElement(
      View,
      { style: pdfDocumentStyles.infoSection },
      React.createElement(
        View,
        { style: pdfDocumentStyles.infoRow },
        React.createElement(Text, { style: pdfDocumentStyles.infoLabel }, 'Name:'),
        React.createElement(Text, { style: pdfDocumentStyles.infoValue }, safeName)
      ),
      React.createElement(
        View,
        { style: pdfDocumentStyles.infoRow },
        React.createElement(Text, { style: pdfDocumentStyles.infoLabel }, 'Card Number:'),
        React.createElement(Text, { style: pdfDocumentStyles.infoValue }, `${safeCardNumber}`)
      ),
      React.createElement(
        View,
        { style: pdfDocumentStyles.infoRow },
        React.createElement(Text, { style: pdfDocumentStyles.infoLabel }, 'Submission Date:'),
        React.createElement(Text, { style: pdfDocumentStyles.infoValue }, safeDate)
      ),
      React.createElement(
        View,
        { style: pdfDocumentStyles.infoRow },
        React.createElement(Text, { style: pdfDocumentStyles.infoLabel }, 'Status:'),
        React.createElement(Text, { style: pdfDocumentStyles.infoValue }, statusLabel)
      ),
      React.createElement(
        View,
        { style: pdfDocumentStyles.infoRow },
        React.createElement(Text, { style: pdfDocumentStyles.infoLabel }, 'Status date & time:'),
        React.createElement(Text, { style: pdfDocumentStyles.infoValue }, safeStatusDecidedAt)
      ),
    ),
    React.createElement(Text, { style: pdfDocumentStyles.sectionTitle }, 'Purchase Details'),
    React.createElement(
      View,
      { style: pdfDocumentStyles.table },
      React.createElement(
        View,
        { style: pdfDocumentStyles.tableHeader },
        ['#', 'Date', 'Store/Website', 'Reason', 'Amount', 'Account'].map((label, index) =>
          React.createElement(
            View,
            {
              key: `header-${label}`,
              style: [
                pdfDocumentStyles.cell,
                { width: widths[index] },
                ...(index === 5 ? [pdfDocumentStyles.lastCell] : []),
              ],
            },
            React.createElement(Text, { style: [pdfDocumentStyles.cellText, { fontWeight: 'bold' }] }, label)
          )
        )
      ),
      safePurchases.map((purchase, index) =>
        React.createElement(
          View,
          {
            key: `purchase-${index}`,
            style: [
              pdfDocumentStyles.tableRow,
              ...(index === safePurchases.length - 1 ? [pdfDocumentStyles.lastTableRow] : []),
            ],
          },
          React.createElement(
            View,
            { style: [pdfDocumentStyles.cell, { width: widths[0] }] },
            React.createElement(Text, { style: pdfDocumentStyles.cellText }, String(index + 1))
          ),
          React.createElement(
            View,
            { style: [pdfDocumentStyles.cell, { width: widths[1] }] },
            React.createElement(Text, { style: pdfDocumentStyles.cellText }, purchase.date)
          ),
          React.createElement(
            View,
            { style: [pdfDocumentStyles.cell, { width: widths[2] }] },
            React.createElement(Text, { style: pdfDocumentStyles.cellText }, purchase.vendor)
          ),
          React.createElement(
            View,
            { style: [pdfDocumentStyles.cell, { width: widths[3] }] },
            React.createElement(Text, { style: pdfDocumentStyles.cellText }, purchase.reason)
          ),
          React.createElement(
            View,
            { style: [pdfDocumentStyles.cell, { width: widths[4] }] },
            React.createElement(Text, { style: pdfDocumentStyles.cellText }, `$${purchase.amount.toFixed(2)}`)
          ),
          React.createElement(
            View,
            { style: [pdfDocumentStyles.cell, { width: widths[5] }, pdfDocumentStyles.lastCell] },
            React.createElement(Text, { style: pdfDocumentStyles.cellText }, purchase.description)
          )
        )
      )
    ),
    React.createElement(Text, { style: pdfDocumentStyles.totalSection }, `Total Amount: $${totalAmount.toFixed(2)}`)
  );

  const imagePages = base64FilesData.map((file, index) =>
    React.createElement(
      Page,
      { key: `image-page-${index}`, size: 'A4', style: pdfDocumentStyles.imagePage, wrap: false },
      React.createElement(
        View,
        { style: pdfDocumentStyles.imagePageImageWrap },
        React.createElement(Image, { src: file.url, style: pdfDocumentStyles.imagePageImage })
      )
    )
  );

  return React.createElement(Document, null, summaryPage, ...imagePages);
}

function buildReimbursementReceiptDocument(params: {
  safeOffice: string;
  safeName: string;
  safeDate: string;
  safePurchases: PdfSafePurchase[];
  totalAmount: number;
  base64FilesData: PdfFileData[];
  approved: boolean;
  safeAmountAdjustedTo: string;
  safeReasonForAdjustment: string;
  safeStatusDecidedAt: string;
}) {
  const {
    safeOffice,
    safeName,
    safeDate,
    safePurchases,
    totalAmount,
    base64FilesData,
    approved,
    safeAmountAdjustedTo,
    safeReasonForAdjustment,
    safeStatusDecidedAt,
  } = params;

  const widths = ['8%', '12%', '18%', '36%', '26%'];
  const headerLabels = ['#', 'Date', 'Store/Website', 'Reason', 'Amount'];

  const statusLabel = approved ? 'Approved' : 'Not Approved';

  const adjustmentBlock: React.ReactNode[] = [];
  if (safeAmountAdjustedTo) {
    adjustmentBlock.push(
      React.createElement(
        View,
        { key: 'adj-amt', style: pdfDocumentStyles.infoRow },
        React.createElement(Text, { style: pdfDocumentStyles.infoLabel }, 'Amount adjusted to:'),
        React.createElement(Text, { style: pdfDocumentStyles.infoValue }, `$${safeAmountAdjustedTo}`)
      )
    );
  }
  if (safeReasonForAdjustment) {
    adjustmentBlock.push(
      React.createElement(
        View,
        { key: 'adj-reason', style: { marginTop: 8 } },
        React.createElement(Text, { style: { fontWeight: 'bold', marginBottom: 4 } }, 'Reason for adjustment / notes:'),
        React.createElement(Text, { style: pdfDocumentStyles.cellText }, safeReasonForAdjustment)
      )
    );
  }

  const summaryPage = React.createElement(
    Page,
    { size: 'A4', style: pdfDocumentStyles.page, wrap: true },
    React.createElement(
      View,
      { style: pdfDocumentStyles.header },
      React.createElement(
        Text,
        { style: pdfDocumentStyles.headerTitle },
        `${safeOffice || 'Company'} Reimbursement Request`
      ),
      React.createElement(View, { style: pdfDocumentStyles.headerLine })
    ),
    React.createElement(
      View,
      { style: pdfDocumentStyles.infoSection },
      React.createElement(
        View,
        { style: pdfDocumentStyles.infoRow },
        React.createElement(Text, { style: pdfDocumentStyles.infoLabel }, 'Name:'),
        React.createElement(Text, { style: pdfDocumentStyles.infoValue }, safeName)
      ),
      React.createElement(
        View,
        { style: pdfDocumentStyles.infoRow },
        React.createElement(Text, { style: pdfDocumentStyles.infoLabel }, 'Submission Date:'),
        React.createElement(Text, { style: pdfDocumentStyles.infoValue }, safeDate)
      ),
      React.createElement(
        View,
        { style: pdfDocumentStyles.infoRow },
        React.createElement(Text, { style: pdfDocumentStyles.infoLabel }, 'Status:'),
        React.createElement(Text, { style: pdfDocumentStyles.infoValue }, statusLabel)
      ),
      React.createElement(
        View,
        { style: pdfDocumentStyles.infoRow },
        React.createElement(Text, { style: pdfDocumentStyles.infoLabel }, 'Status date & time:'),
        React.createElement(Text, { style: pdfDocumentStyles.infoValue }, safeStatusDecidedAt)
      ),
      ...adjustmentBlock
    ),
    React.createElement(Text, { style: pdfDocumentStyles.sectionTitle }, 'Expense Details'),
    React.createElement(
      View,
      { style: pdfDocumentStyles.table },
      React.createElement(
        View,
        { style: pdfDocumentStyles.tableHeader },
        headerLabels.map((label, index) =>
          React.createElement(
            View,
            {
              key: `h-${label}`,
              style: [
                pdfDocumentStyles.cell,
                { width: widths[index] },
                ...(index === headerLabels.length - 1 ? [pdfDocumentStyles.lastCell] : []),
              ],
            },
            React.createElement(Text, { style: [pdfDocumentStyles.cellText, { fontWeight: 'bold' }] }, label)
          )
        )
      ),
      safePurchases.map((purchase, index) =>
        React.createElement(
          View,
          {
            key: `p-${index}`,
            style: [
              pdfDocumentStyles.tableRow,
              ...(index === safePurchases.length - 1 ? [pdfDocumentStyles.lastTableRow] : []),
            ],
          },
          React.createElement(
            View,
            { style: [pdfDocumentStyles.cell, { width: widths[0] }] },
            React.createElement(Text, { style: pdfDocumentStyles.cellText }, String(index + 1))
          ),
          React.createElement(
            View,
            { style: [pdfDocumentStyles.cell, { width: widths[1] }] },
            React.createElement(Text, { style: pdfDocumentStyles.cellText }, purchase.date)
          ),
          React.createElement(
            View,
            { style: [pdfDocumentStyles.cell, { width: widths[2] }] },
            React.createElement(Text, { style: pdfDocumentStyles.cellText }, purchase.vendor)
          ),
          React.createElement(
            View,
            { style: [pdfDocumentStyles.cell, { width: widths[3] }] },
            React.createElement(Text, { style: pdfDocumentStyles.cellText }, purchase.reason)
          ),
          React.createElement(
            View,
            { style: [pdfDocumentStyles.cell, { width: widths[4] }, pdfDocumentStyles.lastCell] },
            React.createElement(Text, { style: pdfDocumentStyles.cellText }, `$${purchase.amount.toFixed(2)}`)
          )
        )
      )
    ),
    React.createElement(Text, { style: pdfDocumentStyles.totalSection }, `Total Amount: $${totalAmount.toFixed(2)}`)
  );

  const imagePages = base64FilesData.map((file, index) =>
    React.createElement(
      Page,
      { key: `img-${index}`, size: 'A4', style: pdfDocumentStyles.imagePage, wrap: false },
      React.createElement(
        View,
        { style: pdfDocumentStyles.imagePageImageWrap },
        React.createElement(Image, { src: file.url, style: pdfDocumentStyles.imagePageImage })
      )
    )
  );

  return React.createElement(Document, null, summaryPage, ...imagePages);
}

async function buildBase64FilesData(filesData: PdfFileData[]): Promise<PdfFileData[]> {
  const base64FilesData: PdfFileData[] = [];
  for (const file of filesData) {
    try {
      const base64Url = await convertReceiptImageUrlToBase64(file.url);
      base64FilesData.push({ ...file, url: base64Url });
    } catch {
      base64FilesData.push(file);
    }
  }
  return base64FilesData;
}

function buildSafePurchasesFromPurchases(purchases: Purchase[]): PdfSafePurchase[] {
  return purchases
    .filter((purchase: Purchase) => {
      return (
        purchase &&
        typeof purchase === 'object' &&
        purchase.date &&
        purchase.vendor &&
        purchase.amount
      );
    })
    .slice(0, 100)
    .map((purchase: Purchase) => ({
      date: pdfSanitizeString(purchase.date, 20),
      vendor: pdfSanitizeString(purchase.vendor, 200),
      reason: pdfSanitizeString(purchase.reason || '', 500),
      amount: pdfSanitizeAmount(purchase.amount),
      description: pdfSanitizeString(purchase.description || '', 200),
    }));
}

async function generateReimbursementPdfBlob(params: {
  name: string;
  date: string;
  office: string;
  purchases: Purchase[];
  filesData: PdfFileData[];
  approved: boolean;
  amountAdjustedTo?: string;
  reasonForAdjustment?: string;
}): Promise<Blob> {
  const { name, date, office, purchases, filesData, approved, amountAdjustedTo, reasonForAdjustment } = params;

  if (!name || !purchases) {
    throw new Error('Missing required fields');
  }
  if (!Array.isArray(purchases) || purchases.length === 0 || purchases.length > 100) {
    throw new Error('Invalid purchases data');
  }
  if (filesData && (!Array.isArray(filesData) || filesData.length > 50)) {
    throw new Error('Invalid files data');
  }

  const safeName = pdfSanitizeString(name, 100);
  const safeDate = pdfSanitizeString(date, 20);
  const safeOffice = pdfSanitizeString(office, 50);
  const safeAmountAdjustedTo = amountAdjustedTo?.trim()
    ? pdfSanitizeString(amountAdjustedTo.replace(/^\$+/, '').trim(), 20)
    : '';
  const safeReasonForAdjustment = reasonForAdjustment?.trim()
    ? pdfSanitizeString(reasonForAdjustment, 1000)
    : '';
  const safeStatusDecidedAt = pdfSanitizeString(formatPdfStatusDateTimeNow(), 120);

  const base64FilesData = filesData?.length ? await buildBase64FilesData(filesData) : [];

  const totalAmount = purchases.reduce((sum: number, purchase: Purchase) => {
    return sum + pdfSanitizeAmount(purchase.amount);
  }, 0);

  const safePurchases = buildSafePurchasesFromPurchases(purchases);

  const doc = buildReimbursementReceiptDocument({
    safeOffice,
    safeName,
    safeDate,
    safePurchases,
    totalAmount,
    base64FilesData,
    approved,
    safeAmountAdjustedTo,
    safeReasonForAdjustment,
    safeStatusDecidedAt,
  });

  return pdf(doc).toBlob();
}

function arrayBufferToBase64DataUrl(arrayBuffer: ArrayBuffer, mime: string): string {
  const bytes = new Uint8Array(arrayBuffer);
  let binary = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode.apply(null, chunk as unknown as number[]);
  }
  const base64 = btoa(binary);
  return `data:${mime};base64,${base64}`;
}

async function convertReceiptImageUrlToBase64(imageUrl: string): Promise<string> {
  try {
    try {
      const url = new URL(imageUrl);
      if (
        url.protocol !== 'https:' ||
        url.hostname !== 'firebasestorage.googleapis.com' ||
        !url.pathname.startsWith('/v0/b/')
      ) {
        return imageUrl;
      }
    } catch {
      return imageUrl;
    }

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 30000);

    const response = await fetch(imageUrl, { signal: controller.signal });
    clearTimeout(timeoutId);

    if (!response.ok) {
      throw new Error(`Failed to fetch image: ${response.status}`);
    }

    const arrayBuffer = await response.arrayBuffer();

    const MAX_FILE_SIZE = 10 * 1024 * 1024;
    if (arrayBuffer.byteLength > MAX_FILE_SIZE) {
      throw new Error('File size exceeds limit');
    }

    const contentTypeHeader = response.headers.get('content-type');
    const contentType = contentTypeHeader ? contentTypeHeader.split(';')[0].trim().toLowerCase() : '';
    if (contentType && !contentType.startsWith('image/')) {
      throw new Error('Invalid file type');
    }

    const allowedTypes = new Set(['image/png', 'image/jpeg', 'image/jpg']);
    const normalizedType = contentType === 'image/jpg' ? 'image/jpeg' : contentType;
    if (normalizedType && !allowedTypes.has(normalizedType)) {
      throw new Error('Unsupported image type');
    }

    return arrayBufferToBase64DataUrl(arrayBuffer, normalizedType || 'image/jpeg');
  } catch {
    return imageUrl;
  }
}

async function generateCreditCardPdfBlob(params: {
  name: string;
  cardNumber: string;
  date: string;
  office: string;
  purchases: Purchase[];
  filesData: PdfFileData[];
  /** Admin PDF flow: credit card is only approved in UI; default true */
  approved?: boolean;
}): Promise<Blob> {
  const { name, cardNumber, date, office, purchases, filesData } = params;
  const approved = params.approved !== false;

  if (!name || !cardNumber || !purchases) {
    throw new Error('Missing required fields');
  }
  if (!Array.isArray(purchases) || purchases.length === 0 || purchases.length > 100) {
    throw new Error('Invalid purchases data');
  }
  if (filesData && (!Array.isArray(filesData) || filesData.length > 50)) {
    throw new Error('Invalid files data');
  }

  const safeName = pdfSanitizeString(name, 100);
  const safeCardNumber = pdfSanitizeString(cardNumber, 4);
  const safeDate = pdfSanitizeString(date, 20);
  const safeOffice = pdfSanitizeString(office, 50);
  const safeStatusDecidedAt = pdfSanitizeString(formatPdfStatusDateTimeNow(), 120);

  const base64FilesData = filesData?.length ? await buildBase64FilesData(filesData) : [];

  const totalAmount = purchases.reduce((sum: number, purchase: Purchase) => {
    return sum + pdfSanitizeAmount(purchase.amount);
  }, 0);

  const safePurchases = buildSafePurchasesFromPurchases(purchases);

  const doc = buildCreditCardReceiptDocument({
    safeOffice,
    safeName,
    safeCardNumber,
    safeDate,
    safePurchases,
    totalAmount,
    base64FilesData,
    approved,
    safeStatusDecidedAt,
  });

  return pdf(doc).toBlob();
}

const AdminCreditCardReview = () => {
  const [submissions, setSubmissions] = useState<Submission[]>([]);
  const [selectedSubmission, setSelectedSubmission] = useState<Submission | null>(null);
  const [loading, setLoading] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const [receiptFiles, setReceiptFiles] = useState<ReceiptFile[]>([]);
  const [loadingFiles, setLoadingFiles] = useState(false);
  const [filterDateFrom, setFilterDateFrom] = useState('');
  const [filterDateTo, setFilterDateTo] = useState('');
  const [filterOffice, setFilterOffice] = useState('');
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
            signed: data.signed || false,
            addedOnNumbersChecked: data.addedOnNumbersChecked || false,
            addedOnNumbersCheckedAt: data.addedOnNumbersCheckedAt?.toDate(),
            formType: 'credit-card',
            pdfURL: data.pdfURL || undefined
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
            addedOnNumbersCheckedAt: data.addedOnNumbersCheckedAt?.toDate ? data.addedOnNumbersCheckedAt.toDate() : (data.addedOnNumbersCheckedAt ? new Date(data.addedOnNumbersCheckedAt) : undefined),
            formType: 'reimbursement',
            amountAdjustedTo: data.amountAdjustedTo || undefined,
            reasonForAdjustment: data.reasonForAdjustment || undefined,
            approved: data.approved !== undefined ? data.approved : undefined,
            pdfURL: data.pdfURL || undefined
          });
        });
      }
      
      // Sort: unchecked items first (by date), then checked items (by date)
      submissions.sort((a: Submission, b: Submission) => {
        const aChecked = a.addedOnNumbersChecked || false;
        const bChecked = b.addedOnNumbersChecked || false;
        
        // If one is checked and the other is not, unchecked comes first
        if (aChecked !== bChecked) {
          return aChecked ? 1 : -1; // unchecked (false) comes first
        }
        
        // If both have the same checked status, sort by date (newest first)
        return b.submittedAt.getTime() - a.submittedAt.getTime();
      });
      setSubmissions(submissions);
    } catch (error) {
      alert('Error loading submissions. Please try again.');
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
        // No submission ID - cannot reliably match files, so don't load any
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
              url: downloadURL
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

  // Role 통과 후에만 제출 목록 로드
  useEffect(() => {
    if (!pageReady) return;
    loadSubmissions();
  }, [pageReady]);

  // Filter submissions based on date range and office
  const getFilteredSubmissions = () => {
    let filtered = submissions;

    // Filter by date range (using submission date, not purchase date)
    if (filterDateFrom || filterDateTo) {
      filtered = filtered.filter((submission: Submission) => {
        const submissionDateStr = getSubmissionDateStringForFilter(submission);
        
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
      alert('Error updating status. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  // Manager Not Approve — generate PDF (no signature pad on this page)
  const managerNotApprove = async () => {
    if (!selectedSubmission) {
      alert('No submission selected');
      return;
    }
    
    if (selectedSubmission.formType !== 'reimbursement') {
      alert('Not Approved option is only available for reimbursement requests.');
      return;
    }
    
    try {
      setLoading(true);
      
      // Store rejection info in Firestore
      const collectionName = 'reimbursement-requests';
      const docRef = doc(db, collectionName, selectedSubmission.id);
      
      let retryCount = 0;
      const maxRetries = 3;
      let firestoreSuccess = false;
      
      while (retryCount < maxRetries && !firestoreSuccess) {
        try {
          const updateData: any = {
            approved: false, // Mark as not approved
            rejectionDate: new Date(),
            signed: true // Mark as signed
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
          } else {
          }
        }
      }
      
      // Generate PDF with rejection status
      await saveToExcel(true); // Pass notApproved flag
      
      alert('❌ Not approved!');
      
      // Close modal and refresh submissions
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

  // Manager Approve — generate PDF (no signature pad on this page)
  const managerApprove = async () => {
    if (!selectedSubmission) {
      alert('No submission selected');
      return;
    }
    
    try {
      setLoading(true);
      
      // Store approval status in Firestore
      // Use appropriate collection based on form type
      const collectionName = selectedSubmission.formType === 'reimbursement' ? 'reimbursement-requests' : 'credit-card-receipts';
      const docRef = doc(db, collectionName, selectedSubmission.id);
      
      let retryCount = 0;
      const maxRetries = 3;
      let firestoreSuccess = false;
      
      while (retryCount < maxRetries && !firestoreSuccess) {
        try {
          const updateData: any = {
            signed: true // Mark as signed
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
            // Wait before retry
            await new Promise(resolve => setTimeout(resolve, 1000 * retryCount));
          } else {
          }
        }
      }
      
      // Generate PDF (CSV will be generated when Download CSV File is clicked)
      await saveToExcel();
      
      alert('✅ Approved!');
      
      // Close modal and refresh submissions
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


  // Generate and Download CSV file
  const downloadExcelFile = async () => {
    try {
      setLoading(true);
      
      // Define CSV file reference
      const mainFileName = 'all-submissions.csv';
      const mainFileRef = ref(storage, `excel/${mainFileName}`);
      
      // Build CSV from Firestore data (완료 건만 — isFirestoreDocPdfWorkflowComplete)
      let csvData = [
        ['Form Type', 'Employee Name', 'Office', 'Card Number', 'Purchase Date', 'Store/Website', 'Reason', 'Amount', 'Account Description', 'Total Amount', 'Submission Date', 'Status', 'PDF Link']
      ];
      
      try {
        // Get all submissions from both collections
        const processedSubmissions: any[] = [];
        
        // Get credit card receipts
        const creditCardSnapshot = await getDocs(collection(db, 'credit-card-receipts'));
        creditCardSnapshot.forEach((doc) => {
          const data = doc.data();
          if (isFirestoreDocPdfWorkflowComplete(data)) {
            processedSubmissions.push({
              id: doc.id,
              employeeName: data.name,
              office: data.office || 'N/A',
              cardNumber: data.cardNumber,
              date: data.date,
              purchases: data.data,
              submissionDateDisplay: getSubmissionDateDisplayForCsv(data),
              pdfURL: data.pdfURL || '',
              formType: 'Credit Card Receipt'
            });
          }
        });
        
        // Get reimbursement requests
        const reimbursementSnapshot = await getDocs(collection(db, 'reimbursement-requests'));
        reimbursementSnapshot.forEach((doc) => {
          const data = doc.data();
          if (isFirestoreDocPdfWorkflowComplete(data)) {
            processedSubmissions.push({
              id: doc.id,
              employeeName: data.name,
              office: data.office || 'N/A',
              cardNumber: data.cardNumber,
              date: data.date,
              purchases: data.data,
              submissionDateDisplay: getSubmissionDateDisplayForCsv(data),
              pdfURL: data.pdfURL || '',
              formType: 'Reimbursement Request'
            });
          }
        });
        
        // Add all processed submissions to CSV data
        for (const submission of processedSubmissions) {
          // 🔒 보안: purchases 배열 검증
          if (!submission.purchases || !Array.isArray(submission.purchases) || submission.purchases.length === 0) {
            continue; // Skip invalid submissions
          }
          
          const totalAmount = submission.purchases.reduce((sum: number, purchase: Purchase) => {
            const amount = parseFloat(purchase.amount) || 0;
            // 🔒 보안: 금액 범위 검증
            return sum + (isFinite(amount) && amount >= 0 && amount <= 1000000 ? amount : 0);
          }, 0);
          
          // Add each purchase as a row with CSV Injection protection
          submission.purchases.forEach((purchase: Purchase, index: number) => {
            // 🔒 보안: purchase 객체 검증
            if (!purchase || typeof purchase !== 'object') {
              return; // Skip invalid purchases
            }
            
            csvData.push([
              index === 0 ? sanitizeCSVCell(submission.formType || 'Credit Card Receipt') : '',
              index === 0 ? sanitizeCSVCell(submission.employeeName) : '',
              index === 0 ? sanitizeCSVCell(submission.office) : '',
              index === 0 ? sanitizeCSVCell(submission.cardNumber) : '',
              sanitizeCSVCell(purchase.date),
              sanitizeCSVCell(purchase.vendor),
              sanitizeCSVCell(purchase.reason),
              sanitizeCSVCell(`$${parseFloat(purchase.amount).toFixed(2)}`),
              sanitizeCSVCell(purchase.description),
              index === 0 ? sanitizeCSVCell(`$${totalAmount.toFixed(2)}`) : '',
              sanitizeCSVCell(submission.submissionDateDisplay || new Date().toLocaleDateString()),
              sanitizeCSVCell('Approved & PDF Generated'),
              index === 0 ? sanitizeCSVCell(submission.pdfURL || '') : ''
            ]);
          });
        }
        
        // Convert to CSV string
        const csvString = csvData.map((row: any[]) => 
          row.map((cell: string) => {
            if (cell === null || cell === undefined) return '';
            const cellStr = String(cell);
            if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n') || cellStr.includes('\r')) {
              return `"${cellStr.replace(/"/g, '""')}"`;
            }
            return cellStr;
          }).join(',')
        ).join('\n');
        
        // Upload CSV to Firebase Storage
        const blob = new Blob([csvString], { type: 'text/csv' });
        await uploadBytes(mainFileRef, blob);
        
        // Download CSV file
        const downloadURL = await getDownloadURL(mainFileRef);
        const link = document.createElement('a');
        link.href = downloadURL;
        link.download = mainFileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        alert(`Downloaded successfully!`);
      } catch (error) {
        alert('❌ Error generating CSV file. Please try again.');
      }
    } catch (error) {
      alert('Error downloading CSV file. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  // Save data to Excel file (client-side)
  const saveToExcel = async (notApproved: boolean = false) => {
    if (!selectedSubmission) {
      return;
    }
    
    // Create local reference to avoid null checks
    const submission = selectedSubmission;
    
    try {
      // Generate PDF and save to Firebase Storage first
      // Build filesData from submission data instead of receiptFiles state
      const filesData: Array<{name: string, url: string, fullPath: string}> = [];
      for (const purchase of submission.purchases) {
        if (purchase.receiptFiles) {
          let receiptFiles: string[] = [];
          
          // Handle different data types
          if (typeof purchase.receiptFiles === 'string') {
            const filesString: string = purchase.receiptFiles;
            receiptFiles = filesString.split(', ');
          } else if (Array.isArray(purchase.receiptFiles)) {
            receiptFiles = purchase.receiptFiles;
          }
          
          for (const fileName of receiptFiles) {
            if (fileName && typeof fileName === 'string' && fileName.trim()) {
              try {
                // Use appropriate storage path based on form type
                const storagePath = submission.formType === 'reimbursement' ? 'reimbursement-receipts/' : 'receipts/';
                const fileRef = ref(storage, `${storagePath}${fileName.trim()}`);
                const downloadURL = await getDownloadURL(fileRef);
                filesData.push({
                  name: fileName.trim(),
                  url: downloadURL,
                  fullPath: `${storagePath}${fileName.trim()}`
                });
              } catch (error) {
              }
            }
          }
        }
      }
      
      let pdfBlob: Blob | null = null;

      try {
        if (submission.formType === 'reimbursement') {
          const approved = notApproved
            ? false
            : submission.approved !== undefined
              ? submission.approved
              : true;
          pdfBlob = await generateReimbursementPdfBlob({
            name: submission.employeeName,
            date: submission.date,
            office: submission.office,
            purchases: submission.purchases,
            filesData,
            approved,
            amountAdjustedTo: submission.amountAdjustedTo || amountAdjustedTo || '',
            reasonForAdjustment: submission.reasonForAdjustment || reasonForAdjustment || '',
          });
        } else {
          const approved = notApproved
            ? false
            : submission.approved !== undefined
              ? submission.approved
              : true;
          pdfBlob = await generateCreditCardPdfBlob({
            name: submission.employeeName,
            cardNumber: submission.cardNumber,
            date: submission.date,
            office: submission.office,
            purchases: submission.purchases,
            filesData,
            approved,
          });
        }
      } catch {
        pdfBlob = null;
      }

      if (pdfBlob) {
        try {
          const pdfArrayBuffer = await pdfBlob.arrayBuffer();
          const pdfBytes = new Uint8Array(pdfArrayBuffer);
          const pdfHeader = String.fromCharCode(pdfBytes[0], pdfBytes[1], pdfBytes[2], pdfBytes[3]);

          if (pdfHeader !== '%PDF') {
            return;
          }

          //const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
          //const timestamp = new Date();

          const now = new Date();
          const laTime = new Date(now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
          let hours = laTime.getHours();
          const minutes = laTime.getMinutes();
          const seconds = laTime.getSeconds();
          const ampm = hours >= 12 ? 'pm' : 'am';
          hours = hours % 12;
          hours = hours ? hours : 12;
          const timeStamp = `${hours}${minutes.toString().padStart(2, '0')}${seconds.toString().padStart(2, '0')}${ampm}`;

          const cardNumberPart =
            submission.cardNumber && submission.cardNumber.trim() !== ''
              ? `${submission.cardNumber}_`
              : '';
          const pdfFileName = `pdfs/${submission.employeeName}_${submission.formType}_${submission.date}_${timeStamp}.pdf`;
          const pdfRef = ref(storage, pdfFileName);
          await uploadBytes(pdfRef, pdfBlob);

          const pdfDownloadURL = await getDownloadURL(pdfRef);

          try {
            const collectionName =
              submission.formType === 'reimbursement' ? 'reimbursement-requests' : 'credit-card-receipts';
            const docRef = doc(db, collectionName, submission.id);
            await setDoc(
              docRef,
              {
                pdfURL: pdfDownloadURL,
                pdfGeneratedAt: new Date(),
              },
              { merge: true }
            );
          } catch (firestoreError) {
            // Continue even if Firestore save fails
          }
        } catch (error) {
          // Continue without PDF link if storage fails
        }
      }
      
    } catch (error) {
      // PDF 생성 실패 시 처리
    }
  };

  // Delete completed submissions only (Reset function)
  const resetAllData = async () => {
    if (!confirm('⚠️ This will delete ONLY completed submissions (PDF generated and numbers checked). Are you sure you want to continue?')) {
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
        if (isFirestoreDocPdfWorkflowComplete(data)) {
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
        if (isFirestoreDocPdfWorkflowComplete(data)) {
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
        alert('No completed submissions found to reset. All submissions must have PDF generated and numbers checked.');
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
            // Delete receipt files for this submission (now safe to delete since PDF uses Base64)
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
                      // Continue even if deletion fails
                    }
                  }
                }
              }
            }
          } catch (submissionError) {
            // Continue with next submission
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
      setReceiptFiles([]);
      
      alert(`✅ Reset completed! Deleted ${completedCount} completed submissions. ${incompleteCount} incomplete submissions were kept.`);
      
    } catch (error) {
      alert('Error resetting data. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  // Delete single submission completely
  const deleteSubmission = async () => {
    if (!selectedSubmission) {
      alert('No submission selected');
      return;
    }
    
    if (!confirm(`⚠️ Are you sure you want to delete "${selectedSubmission.employeeName}"'s submission?`)) {
      return;
    }
    
    try {
      setDeleting(true);
      
      const submission = selectedSubmission;
      const collectionName = submission.formType === 'reimbursement' ? 'reimbursement-requests' : 'credit-card-receipts';
      
      // 1. Delete from Firestore
      await deleteDoc(doc(db, collectionName, submission.id));
      
      // 2. Delete receipt files from Storage
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
                // Continue even if deletion fails
              }
            }
          }
        }
      }
      
      // 3. Delete PDF file from Storage
      let pdfDeleted = false;
      
      // Try to delete using pdfURL if available
      if (submission.pdfURL) {
        try {
          const url = new URL(submission.pdfURL);
          const pathMatch = url.pathname.match(/\/o\/(.+)\?/);
          if (pathMatch) {
            const filePath = decodeURIComponent(pathMatch[1]);
            const pdfRef = ref(storage, filePath);
            await deleteObject(pdfRef);
            pdfDeleted = true;
          }
        } catch (error) {
          // Continue to try alternative method
        }
      }
      
      // If pdfURL method failed or doesn't exist, try to find and delete from pdfs/ folder
      if (!pdfDeleted) {
        try {
          const pdfsRef = ref(storage, 'pdfs/');
          const pdfsList = await listAll(pdfsRef);
          
          // Build search pattern based on submission data
          const cardNumberPart = submission.cardNumber && submission.cardNumber.trim() !== '' 
            ? `${submission.cardNumber}_` 
            : '';
          const searchPrefix = `${submission.employeeName}_${cardNumberPart}`;
          
          // Find matching PDF files
          for (const item of pdfsList.items) {
            if (item.name.startsWith(searchPrefix)) {
              try {
                await deleteObject(item);
                pdfDeleted = true;
              } catch (error) {
                // Continue with next file
              }
            }
          }
        } catch (error) {
          // Continue even if PDF deletion fails
        }
      }
      
      // 4. Remove from Excel file
      try {
        const mainFileName = 'all-submissions.csv';
        const mainFileRef = ref(storage, `excel/${mainFileName}`);
        
        // Try to get existing CSV file
        try {
          const csvBlob = await fetch(await getDownloadURL(mainFileRef)).then(r => r.blob());
          const csvText = await csvBlob.text();
          const lines = csvText.split('\n');
          
          // Filter out lines that belong to this submission
          // CSV format: Form Type, Employee Name, Office, Card Number, Purchase Date, Store/Website, Reason, Amount, Account Description, Total Amount, Submission Date, Status, PDF Link
          const filteredLines = lines.filter((line, index) => {
            // Keep header
            if (index === 0) return true;
            // Skip empty lines
            if (!line.trim()) return true; // Keep empty lines to preserve structure
            
            // Parse CSV line (handle quoted values properly)
            const cells: string[] = [];
            let currentCell = '';
            let inQuotes = false;
            
            for (let i = 0; i < line.length; i++) {
              const char = line[i];
              const nextChar = i < line.length - 1 ? line[i + 1] : '';
              
              if (char === '"') {
                // Handle escaped quotes ("")
                if (inQuotes && nextChar === '"') {
                  currentCell += '"';
                  i++; // Skip next quote
                } else {
                  inQuotes = !inQuotes;
                }
              } else if (char === ',' && !inQuotes) {
                cells.push(currentCell.trim());
                currentCell = '';
              } else {
                currentCell += char;
              }
            }
            cells.push(currentCell.trim()); // Add last cell
            
            // Need at least 13 columns
            if (cells.length >= 13) {
              // Clean up cells (remove quotes and sanitize)
              const formType = cells[0].replace(/^["']|["']$/g, '').trim();
              const employeeName = cells[1].replace(/^["']|["']$/g, '').trim();
              const cardNumberCell = cells[3] ? cells[3].replace(/^["']|["']$/g, '').trim() : '';
              // Card number is stored as ****1234, so extract the actual number
              const cardNumber = cardNumberCell.replace(/^\*+/, '').trim();
              const pdfURL = cells[12] ? cells[12].replace(/^["']|["']$/g, '').trim() : '';
              
              // Expected form type label
              const expectedFormType = submission.formType === 'reimbursement' ? 'Reimbursement Request' : 'Credit Card Receipt';
              
              // Match by employee name, form type, and optionally card number or PDF URL
              const nameMatches = employeeName === submission.employeeName;
              const formTypeMatches = formType === expectedFormType;
              
              // For reimbursement, card number might be empty, so don't check it
              // For credit card, check card number match
              let cardMatches = true;
              if (submission.formType !== 'reimbursement') {
                if (cardNumber && submission.cardNumber) {
                  cardMatches = cardNumber === submission.cardNumber;
                } else if (!cardNumber && !submission.cardNumber) {
                  cardMatches = true; // Both empty
                } else {
                  cardMatches = false; // One has it, one doesn't
                }
              }
              
              // PDF URL matching (if available)
              let pdfMatches = true;
              if (submission.pdfURL && pdfURL) {
                pdfMatches = pdfURL === submission.pdfURL;
              }
              
              // Remove line if all conditions match
              if (nameMatches && formTypeMatches && cardMatches && pdfMatches) {
                return false; // Remove this line
              }
            }
            return true; // Keep this line
          });
          
          // Upload updated CSV
          const updatedCsv = filteredLines.join('\n');
          const blob = new Blob([updatedCsv], { type: 'text/csv' });
          await uploadBytes(mainFileRef, blob);
        } catch (error) {
          // Continue even if Excel update fails
        }
      } catch (error) {
        // Continue even if Excel update fails
      }
      
      alert('✅ Submission deleted successfully!');
      setSelectedSubmission(null);
      setReceiptFiles([]);
      loadSubmissions(); // Reload submissions
      
    } catch (error) {
      alert('Error deleting submission. Please try again.');
    } finally {
      setDeleting(false);
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
    // Highlight when submission.signed (approved); same for credit-card and reimbursement
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
              Download CSV File
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
              Reset All Data
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
            {getFilteredSubmissions().map((submission: Submission) => (
              <div
                key={submission.id}
                style={submission.signed ? styles.submissionCardSigned : styles.submissionCard}
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
                      // Already signed - show approval status based on approved field (for reimbursement)
                      (() => {
                        // Reimbursement: Check approved status
                        if (selectedSubmission.formType === 'reimbursement') {
                          // Strict check: approved must be explicitly false
                          const isNotApproved = selectedSubmission.approved === false;
                          const isApproved = selectedSubmission.approved === true;
                          
                          if (isNotApproved) {
                            // Not Approved
                            return (
                              <div style={{
                                padding: '15px',
                                backgroundColor: '#f8d7da',
                                borderRadius: '8px',
                                border: '2px solid #dc3545',
                                textAlign: 'center'
                              }}>
                                <h4 style={{margin: '0 0 10px 0', color: '#721c24'}}>❌ Not Approved</h4>
                              </div>
                            );
                          } else if (isApproved || selectedSubmission.approved === undefined) {
                            // Approved (true or undefined)
                            return (
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
                            );
                          } else {
                            // Fallback: show Approved
                            return (
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
                            );
                          }
                        } else {
                          // Credit Card: Always show Approved when signed
                          return (
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
                          );
                        }
                      })()
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
                  
                  {/* Delete Submission Button */}
                  <div style={{marginTop: '20px', paddingTop: '20px', borderTop: '2px solid #e9ecef'}}>
                    <button
                      style={{
                        ...styles.button,
                        backgroundColor: '#dc3545',
                        width: '100%'
                      }}
                      onClick={deleteSubmission}
                      disabled={loading || deleting}
                    >
                      {deleting ? 'Deleting...' : '🗑️ Delete Submission'}
                    </button>
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


