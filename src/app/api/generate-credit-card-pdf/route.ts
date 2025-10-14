import { NextRequest, NextResponse } from 'next/server';

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
    const response = await fetch(imageUrl);
    const arrayBuffer = await response.arrayBuffer();
    const base64 = Buffer.from(arrayBuffer).toString('base64');
    return `data:image/jpeg;base64,${base64}`;
  } catch (error) {
    console.warn('⚠️ Could not convert image to Base64:', error);
    return imageUrl; // Fallback to original URL
  }
}

// Generate PDF HTML for credit card receipts
export async function POST(request: NextRequest) {
  try {
    const { name, cardNumber, date, office, purchases, filesData, signature } = await request.json();

    console.log('📄 PDF API received data:', { 
      name, 
      cardNumber, 
      date, 
      office, 
      purchases: purchases?.length, 
      filesData: filesData?.length,
      signature: signature ? 'Present' : 'Missing'
    });
    
    console.log('📄 Full request data:', {
      name,
      cardNumber,
      date,
      office,
      purchases,
      filesData,
      signature: signature ? 'Present' : 'Missing'
    });
    
    console.log('📄 Signature details:', {
      signatureType: typeof signature,
      signatureLength: signature ? signature.length : 0,
      signatureStart: signature ? signature.substring(0, 50) + '...' : 'None',
      hasSignature: !!signature,
      isBase64: signature ? signature.startsWith('data:image') : false,
      isURL: signature ? signature.startsWith('http') : false,
      isMissing: signature === 'Missing',
      signatureValue: signature
    });

    if (!name || !cardNumber || !purchases) {
      return NextResponse.json({
        success: false,
        error: 'Missing required fields'
      });
    }

    // Convert receipt images to Base64
    let base64FilesData: FileData[] = [];
    if (filesData && filesData.length > 0) {
      console.log('📄 Converting receipt images to Base64...');
      for (const file of filesData as FileData[]) {
        try {
          const base64Url = await convertImageToBase64(file.url);
          base64FilesData.push({
            ...file,
            url: base64Url
          });
          console.log(`📄 Converted ${file.name} to Base64`);
        } catch (error) {
          console.warn(`⚠️ Could not convert ${file.name} to Base64:`, error);
          base64FilesData.push(file); // Use original URL as fallback
        }
      }
    }
    
    console.log('📄 Base64 conversion complete:', base64FilesData.length, 'files converted');

    // Calculate total amount
    const totalAmount = purchases.reduce((sum: number, purchase: Purchase) => {
      return sum + (parseFloat(purchase.amount) || 0);
    }, 0);

    // Generate HTML for PDF
    const html = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <title>Credit Card Receipt - ${name}</title>
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
          <h1>${office || 'Company'} Credit Card Receipt</h1>
          <div class="header-line"></div>
        </div>

        <div class="info-section">
          <div class="info-row">
            <span class="info-label">Name:</span>
            <span class="info-value">${name}</span>
          </div>
          <div class="info-row">
            <span class="info-label">Card Number:</span>
            <span class="info-value">XXXX-XXXX-XXXX-${cardNumber}</span>
          </div>
          <div class="info-row">
            <span class="info-label">Submission Date:</span>
            <span class="info-value">${date}</span>
          </div>
        </div>

        <div class="section-title">Purchase Details</div>
        <table class="purchases-table">
          <thead>
            <tr>
              <th>#</th>
              <th>Date</th>
              <th>Store/Website</th>
              <th>Reason</th>
              <th>Amount</th>
              <th>Account</th>
            </tr>
          </thead>
          <tbody>
            ${purchases.map((purchase: Purchase, index: number) => `
              <tr>
                <td>${index + 1}</td>
                <td>${purchase.date}</td>
                <td>${purchase.vendor}</td>
                <td>${purchase.reason}</td>
                <td>$${parseFloat(purchase.amount).toFixed(2)}</td>
                <td>${purchase.description}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <div class="total-section">Total Amount: $${totalAmount.toFixed(2)}</div>

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

    console.log('📄 PDF HTML generated successfully, length:', html.length);
    console.log('📄 Signature section included:', signature && signature.trim() !== '' ? 'Yes' : 'No');
    console.log('📄 Final signature check:', {
      hasSignature: !!signature,
      signatureNotEmpty: signature && signature.trim() !== '',
      signatureLength: signature ? signature.length : 0
    });
    
    return NextResponse.json({
      success: true,
      html: html,
      message: 'PDF generated successfully'
    });

  } catch (error) {
    console.error('📄 Error generating PDF:', error);
    return NextResponse.json({
      success: false,
      error: error instanceof Error ? error.message : 'Unknown error'
    });
  }
}