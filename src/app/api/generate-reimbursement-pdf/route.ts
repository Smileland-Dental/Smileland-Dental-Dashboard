import { NextRequest, NextResponse } from 'next/server';
import { escapeHtml, sanitizeAmount, sanitizeString, safeLog, logError } from '@/lib/security-server';

// Interfaces for type safety
interface Purchase {
  date: string;
  vendor: string;
  reason: string;
  amount: string;
  description: string;
}

interface FileData {
  name: string;
  url: string;
}

// Convert image URL to Base64
async function convertImageToBase64(imageUrl: string): Promise<string> {
  try {
    // 🔒 보안 1: Firebase Storage URL만 허용 (SSRF 방지)
    if (!imageUrl.includes('firebasestorage.googleapis.com')) {
      logError(`Invalid image URL: ${imageUrl}`, 'convertImageToBase64');
      return imageUrl; // Fallback to original URL
    }

    // 🔒 보안 2: 타임아웃 설정 (30초)
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 30000);

    const response = await fetch(imageUrl, { signal: controller.signal });
    clearTimeout(timeoutId);

    const arrayBuffer = await response.arrayBuffer();

    // 🔒 보안 3: 파일 크기 제한 10MB (DoS 방지)
    const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10MB
    if (arrayBuffer.byteLength > MAX_FILE_SIZE) {
      logError(`File too large: ${arrayBuffer.byteLength} bytes`, 'convertImageToBase64');
      return imageUrl; // Fallback to original URL
    }

    const base64 = Buffer.from(arrayBuffer).toString('base64');
    return `data:image/jpeg;base64,${base64}`;
  } catch (error) {
    logError(error, 'convertImageToBase64');
    return imageUrl; // Fallback to original URL
  }
}

// Generate PDF HTML for reimbursement requests
export async function POST(request: NextRequest) {
  try {
    const { name, cardNumber, date, office, purchases, filesData, signature, amountAdjustedTo, reasonForAdjustment, approved } = await request.json();

    // 🔒 보안: Production에서는 민감 정보 로그 출력 안 함
    safeLog('📄 PDF API received data:', { 
      hasName: !!name,
      hasCardNumber: !!cardNumber,
      hasDate: !!date,
      hasOffice: !!office,
      purchaseCount: purchases?.length,
      fileCount: filesData?.length,
      hasSignature: !!signature
    });

    // 🔒 입력 검증
    // Note: cardNumber is optional for reimbursement requests (can be empty string)
    if (!name || !purchases) {
      logError('Missing required fields', 'generate-reimbursement-pdf');
      return NextResponse.json({
        success: false,
        error: 'Missing required fields'
      });
    }

    // 🔒 XSS 방지: HTML Escape
    const safeName = sanitizeString(name, 100);
    // cardNumber can be empty for reimbursement requests
    const safeCardNumber = cardNumber ? sanitizeString(cardNumber, 4) : '';
    const safeDate = sanitizeString(date, 20);
    const safeOffice = sanitizeString(office, 50);
    // Sanitize adjustment fields
    const safeAmountAdjustedTo = amountAdjustedTo ? sanitizeString(amountAdjustedTo, 50) : '';
    const safeReasonForAdjustment = reasonForAdjustment ? sanitizeString(reasonForAdjustment, 500) : '';

    // Convert receipt images to Base64
    let base64FilesData: FileData[] = [];
    if (filesData && filesData.length > 0) {
      safeLog('📄 Converting receipt images to Base64...');
      for (const file of filesData as FileData[]) {
        try {
          const base64Url = await convertImageToBase64(file.url);
          base64FilesData.push({
            ...file,
            url: base64Url
          });
          safeLog(`📄 Converted file to Base64`);
        } catch (error) {
          logError(error, 'Base64 conversion');
          base64FilesData.push(file); // Use original URL as fallback
        }
      }
    }
    
    safeLog('📄 Base64 conversion complete:', base64FilesData.length);

    // Calculate total amount with sanitization
    const totalAmount = purchases.reduce((sum: number, purchase: Purchase) => {
      return sum + sanitizeAmount(purchase.amount);
    }, 0);

    // 🔒 XSS 방지: purchases 배열의 각 항목도 escape
    interface SafePurchase {
      date: string;
      vendor: string;
      reason: string;
      amount: number;
      description: string;
    }
    
    const safePurchases: SafePurchase[] = purchases.map((purchase: Purchase) => ({
      date: sanitizeString(purchase.date, 20),
      vendor: sanitizeString(purchase.vendor, 200),
      reason: sanitizeString(purchase.reason, 500),
      amount: sanitizeAmount(purchase.amount),
      description: sanitizeString(purchase.description, 200)
    }));

    // Generate HTML for PDF
    const html = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <title>Reimbursement Request - ${safeName}</title>
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 30px;
            background: white;
            color: #000;
            line-height: 1.4;
          }
          .header {
            text-align: center;
            margin-bottom: 30px;
          }
          .header h1 {
            color: #000;
            margin: 0;
            font-size: 20px;
            font-weight: bold;
            text-transform: uppercase;
            letter-spacing: 1px;
          }
          .header-line {
            border-bottom: 2px solid #000;
            margin: 10px 0 20px 0;
          }
          .info-section {
            margin-bottom: 25px;
          }
          .info-row {
            display: flex;
            margin-bottom: 8px;
            align-items: baseline;
          }
          .info-label {
            font-weight: bold;
            color: #000;
            min-width: 120px;
          }
          .info-value {
            color: #000;
            flex: 1;
            border-bottom: 1px solid #000;
            padding-bottom: 2px;
            margin-left: 10px;
          }
          .section-title {
            font-weight: bold;
            font-size: 14px;
            margin: 20px 0 10px 0;
            color: #000;
          }
          .purchases-table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
          }
          .purchases-table th {
            background: #f5f5f5;
            color: #000;
            padding: 8px 6px;
            text-align: left;
            font-weight: bold;
            border: 1px solid #000;
            font-size: 11px;
          }
          .purchases-table td {
            padding: 8px 6px;
            border: 1px solid #000;
            font-size: 11px;
          }
          .total-section {
            text-align: right;
            margin: 20px 0;
            font-weight: bold;
            font-size: 14px;
          }
          .receipt-files {
            margin-top: 20px;
          }
          .receipt-files h4 {
            margin: 0 0 10px 0;
            color: #000;
            font-size: 12px;
            font-weight: bold;
          }
          .signature-section {
            margin-top: 30px;
          }
          .signature-row {
            display: flex;
            margin-bottom: 15px;
            align-items: baseline;
          }
          .signature-label {
            font-weight: bold;
            color: #000;
            min-width: 200px;
          }
          .signature-line {
            flex: 1;
            border-bottom: 1px solid #000;
            margin-left: 10px;
            height: 20px;
          }
          .signature-image {
            max-width: 200px;
            max-height: 80px;
            margin: 10px 0;
          }
        </style>
      </head>
      <body>
        <div class="header">
          <h1>Reimbursement Request</h1>
          <div class="header-line"></div>
        </div>

        <div class="info-section">
          <div class="info-row">
            <span class="info-label">Name:</span>
            <span class="info-value">${safeName}</span>
          </div>
          ${safeCardNumber ? `
          <div class="info-row">
            <span class="info-label">Card Number:</span>
            <span class="info-value">XXXX-XXXX-XXXX-${safeCardNumber}</span>
          </div>
          ` : ''}
          <div class="info-row">
            <span class="info-label">Submission Date:</span>
            <span class="info-value">${safeDate}</span>
          </div>
        </div>

        <div class="section-title">Purchase Details</div>
        <table class="purchases-table">
          <thead>
            <tr>
              <th>#</th>
              <th>Date</th>
              <th>Item</th>
              <th>Reason</th>
              <th>Amount</th>
            </tr>
          </thead>
          <tbody>
            ${safePurchases.map((purchase, index: number) => `
              <tr>
                <td>${index + 1}</td>
                <td>${purchase.date}</td>
                <td>${purchase.vendor}</td>
                <td>${purchase.reason}</td>
                <td>$${purchase.amount.toFixed(2)}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <div class="total-section">Total Amount: $${totalAmount.toFixed(2)}</div>

        ${(safeAmountAdjustedTo || safeReasonForAdjustment) ? `
          <div class="info-section" style="margin-top: 25px; padding: 15px; background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 8px;">
            ${safeAmountAdjustedTo ? `
              <div class="info-row" style="margin-bottom: 12px;">
                <span class="info-label">Amount adjusted to:</span>
                <span class="info-value" style="font-weight: bold;">$${parseFloat(safeAmountAdjustedTo).toFixed(2)}</span>
              </div>
            ` : ''}
            ${safeReasonForAdjustment ? `
              <div class="info-row">
                <span class="info-label">Reason for Adjustment or Non-Approval:</span>
                <span class="info-value" style="padding-left: 10px; padding-top: 5px;">${safeReasonForAdjustment}</span>
              </div>
            ` : ''}
          </div>
        ` : ''}

        ${base64FilesData && base64FilesData.length > 0 ? `
          <div class="receipt-files">
            <h4>Receipt Images</h4>
            <div style="display: flex; flex-direction: column; gap: 15px; margin-top: 10px;">
              ${base64FilesData.map((file: FileData) => `
                <div style="padding: 10px; text-align: center; background: white;">
                  <img src="${file.url}" alt="Receipt" style="max-width: 100%; height: auto;" />
                </div>
              `).join('')}
            </div>
          </div>
        ` : ''}

        <div class="signature-section">
          ${approved !== undefined ? `
            <div class="info-row" style="margin-bottom: 20px; padding: 12px; background: #f8f9fa; border: 1px solid #000; border-radius: 8px;">
              <span class="info-label" style="font-weight: bold; font-size: 16px; color: #000;">
                Status: ${approved ? '✅ APPROVED' : '❌ NOT APPROVED'}
              </span>
            </div>
          ` : ''}
          <div class="signature-row">
            <span class="signature-label">Dr. Oh's Signature of Authorization:</span>
          </div>
          ${signature && signature !== 'Missing' && signature.trim() !== '' && (signature.startsWith('data:image') || signature.startsWith('http')) ? `
            <div style="text-align: center; margin: 10px 0;">
              <img src="${signature}" alt="Dr. Oh Signature" class="signature-image" style="max-width: 200px; max-height: 80px;" />
            </div>
          ` : ''}
          <div class="signature-row">
            <span class="signature-label">Date:</span>
            ${signature && signature !== 'Missing' && signature.trim() !== '' && (signature.startsWith('data:image') || signature.startsWith('http')) ? `
              <span class="info-value" style="margin-left: 10px;">${new Date().toLocaleDateString()}</span>
            ` : `
              <div class="signature-line"></div>
            `}
          </div>
        </div>
      </body>
      </html>
    `;

    safeLog('📄 PDF HTML generated successfully');
    
    return NextResponse.json({
      success: true,
      html: html,
      message: 'PDF generated successfully'
    });

  } catch (error) {
    logError(error, 'generate-reimbursement-pdf');
    
    // 🔒 보안: 에러 메시지에서 민감 정보 제외
    return NextResponse.json({
      success: false,
      error: 'Failed to generate PDF'
    }, { status: 500 });
  }
}

