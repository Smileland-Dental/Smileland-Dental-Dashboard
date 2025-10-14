import { NextRequest, NextResponse } from 'next/server';
import { initializeApp, getApps } from 'firebase-admin/app';
import { getFirestore } from 'firebase-admin/firestore';
import { credential } from 'firebase-admin';

// Initialize Firebase Admin SDK
if (!getApps().length) {
  initializeApp({
    credential: credential.applicationDefault(),
    projectId: process.env.FIREBASE_PROJECT_ID,
    storageBucket: process.env.FIREBASE_STORAGE_BUCKET,
  });
}

const db = getFirestore();

// Purchase interface for type safety
interface Purchase {
  date: string;
  vendor: string;
  reason: string;
  amount: string;
  description: string;
}

// Approve credit card submission and generate PDF
export async function POST(request: NextRequest) {
  try {
    const { submissionId, signature, employeeName, cardNumber, date, purchases } = await request.json();

    if (!submissionId || !signature || !employeeName || !cardNumber || !purchases) {
      return NextResponse.json({
        success: false,
        error: 'Missing required fields'
      });
    }

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
        <title>Credit Card Receipt - ${employeeName}</title>
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background: white;
          }
          .header {
            text-align: center;
            border-bottom: 3px solid #4CAF50;
            padding-bottom: 20px;
            margin-bottom: 30px;
          }
          .header h1 {
            color: #2c3e50;
            margin: 0;
            font-size: 28px;
          }
          .employee-info {
            background: #f8f9fa;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
          }
          .employee-info h3 {
            margin: 0 0 10px 0;
            color: #2c3e50;
          }
          .info-row {
            display: flex;
            justify-content: space-between;
            margin-bottom: 5px;
          }
          .info-label {
            font-weight: bold;
            color: #555;
          }
          .info-value {
            color: #333;
          }
          .purchases-table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 30px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
          }
          .purchases-table th {
            background: #4CAF50;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: bold;
          }
          .purchases-table td {
            padding: 12px;
            border-bottom: 1px solid #ddd;
          }
          .purchases-table tr:nth-child(even) {
            background: #f8f9fa;
          }
          .total-section {
            background: #e8f5e8;
            padding: 20px;
            border-radius: 8px;
            text-align: right;
            margin-bottom: 30px;
          }
          .total-amount {
            font-size: 24px;
            font-weight: bold;
            color: #2c3e50;
          }
          .signature-section {
            border-top: 2px solid #ddd;
            padding-top: 20px;
            margin-top: 30px;
          }
          .signature-box {
            border: 2px solid #ddd;
            height: 100px;
            margin: 10px 0;
            display: flex;
            align-items: center;
            justify-content: center;
            background: #f8f9fa;
          }
          .signature-image {
            max-width: 100%;
            max-height: 100%;
          }
          .approval-info {
            display: flex;
            justify-content: space-between;
            margin-top: 20px;
          }
          .approval-date {
            color: #666;
            font-size: 14px;
          }
        </style>
      </head>
      <body>
        <div class="header">
          <h1>🏢 Credit Card Receipt Approval</h1>
        </div>

        <div class="employee-info">
          <h3>Employee Information</h3>
          <div class="info-row">
            <span class="info-label">Employee Name:</span>
            <span class="info-value">${employeeName}</span>
          </div>
          <div class="info-row">
            <span class="info-label">Card Number:</span>
            <span class="info-value">****${cardNumber}</span>
          </div>
          <div class="info-row">
            <span class="info-label">Date of Purchase:</span>
            <span class="info-value">${date}</span>
          </div>
          <div class="info-row">
            <span class="info-label">Approval Date:</span>
            <span class="info-value">${new Date().toLocaleDateString()}</span>
          </div>
        </div>

        <table class="purchases-table">
          <thead>
            <tr>
              <th>Date</th>
              <th>Store/Website</th>
              <th>Reason</th>
              <th>Amount</th>
              <th>Account</th>
            </tr>
          </thead>
          <tbody>
            ${purchases.map((purchase: Purchase) => `
              <tr>
                <td>${purchase.date}</td>
                <td>${purchase.vendor}</td>
                <td>${purchase.reason}</td>
                <td>$${parseFloat(purchase.amount).toFixed(2)}</td>
                <td>${purchase.description}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>

        <div class="total-section">
          <div class="total-amount">Total Amount: $${totalAmount.toFixed(2)}</div>
        </div>

        <div class="signature-section">
          <h3>Manager Approval</h3>
          <p>This credit card submission has been reviewed and approved by management.</p>
          <div class="signature-box">
            <img src="${signature}" alt="Manager Signature" class="signature-image" />
          </div>
          <div class="approval-info">
            <span>Manager Signature</span>
            <span class="approval-date">Approved on: ${new Date().toLocaleDateString()}</span>
          </div>
        </div>
      </body>
      </html>
    `;

    // Delete the original submission from Firestore
    await db.collection('credit-card-receipts').doc(submissionId).delete();

    return NextResponse.json({
      success: true,
      html: html,
      message: 'Submission approved and PDF generated successfully'
    });

  } catch (error) {
    console.error('Error approving submission:', error);
    return NextResponse.json({
      success: false,
      error: 'Failed to approve submission'
    });
  }
}
