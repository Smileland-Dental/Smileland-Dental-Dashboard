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
  purchases: Purchase[];
  totalAmount: string;
  submittedAt: Date;
  lastUpdated: Date;
}

// Get all submitted credit card data
export async function GET(request: NextRequest) {
  try {
    // Get all documents from credit-card-receipts collection
    const snapshot = await db.collection('credit-card-receipts').get();
    
    const submissions: Submission[] = [];
    
    snapshot.forEach((doc) => {
      const data = doc.data();
      if (data.data && data.data.length > 0) {
        // Calculate total amount
        const totalAmount = data.data.reduce((sum: number, row: any[]) => {
          const amount = parseFloat(row[5]) || 0; // Amount is at index 5
          return sum + amount;
        }, 0);

        submissions.push({
          id: doc.id,
          employeeName: data.name,
          cardNumber: data.cardNumber,
          date: data.date || data.data[0][2], // Use stored date or fallback to first purchase date
          purchases: data.data.map((row: any[]) => ({
            date: row[2],
            vendor: row[3],
            reason: row[4],
            amount: row[5],
            description: row[6],
            receiptFiles: row[7] ? row[7].split(', ') : []
          })),
          totalAmount: totalAmount.toFixed(2),
          submittedAt: data.createdAt?.toDate() || new Date(),
          lastUpdated: data.lastUpdated?.toDate() || new Date()
        });
      }
    });

    // Sort by submission date (newest first)
    submissions.sort((a: Submission, b: Submission) => b.submittedAt.getTime() - a.submittedAt.getTime());

    return NextResponse.json({
      success: true,
      data: submissions
    });

  } catch (error) {
    console.error('Error loading submissions:', error);
    return NextResponse.json({
      success: false,
      error: 'Failed to load submissions'
    });
  }
}
