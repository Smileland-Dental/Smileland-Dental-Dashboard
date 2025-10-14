import { NextRequest, NextResponse } from 'next/server';
import { initializeApp, getApps, getApp } from 'firebase/app';
import { getStorage, ref, uploadBytes, getDownloadURL, listAll } from 'firebase/storage';
import * as XLSX from 'xlsx';

// Firebase 설정
const firebaseConfig = {
  apiKey: process.env.NEXT_PUBLIC_FIREBASE_API_KEY,
  authDomain: process.env.NEXT_PUBLIC_FIREBASE_AUTH_DOMAIN,
  projectId: process.env.NEXT_PUBLIC_FIREBASE_PROJECT_ID,
  storageBucket: process.env.NEXT_PUBLIC_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: process.env.NEXT_PUBLIC_FIREBASE_MESSAGING_SENDER_ID,
  appId: process.env.NEXT_PUBLIC_FIREBASE_APP_ID
};

// Firebase 초기화
const app = !getApps().length ? initializeApp(firebaseConfig) : getApp();
const storage = getStorage(app);

// Interface for type safety
interface Purchase {
  date: string;
  vendor: string;
  reason: string;
  amount: string;
  description: string;
}

// Excel 파일 생성 및 Firebase Storage에 저장
export async function POST(request: NextRequest) {
  try {
    const { 
      employeeName, 
      cardNumber, 
      date, 
      purchases, 
      totalAmount 
    } = await request.json();

    if (!employeeName || !cardNumber || !purchases) {
      return NextResponse.json({
        success: false,
        error: 'Missing required fields'
      });
    }

    // 기존 Excel 파일이 있는지 확인
    const excelFileName = 'credit-card-receipts.xlsx';
    const excelFileRef = ref(storage, `excel/${excelFileName}`);
    
    let existingData: any[] = [];
    let worksheet;
    
    try {
      // 기존 파일 다운로드 시도
      const existingFile = await getDownloadURL(excelFileRef);
      const response = await fetch(existingFile);
      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      worksheet = workbook.Sheets[workbook.SheetNames[0]];
      existingData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    } catch (error) {
      // 파일이 없으면 새로 생성
      console.log('No existing Excel file found, creating new one');
      existingData = [
        ['Employee Name', 'Card Number', 'Purchase Date', 'Store/Website', 'Reason', 'Amount', 'Account Description', 'Total Amount', 'Submission Date', 'Status']
      ];
    }

    // 새 데이터 추가
    const newRows = purchases.map((purchase: Purchase, index: number) => [
      index === 0 ? employeeName : '', // 첫 번째 행에만 직원명
      index === 0 ? `****${cardNumber}` : '', // 첫 번째 행에만 카드번호
      purchase.date,
      purchase.vendor,
      purchase.reason,
      `$${parseFloat(purchase.amount).toFixed(2)}`,
      purchase.description,
      index === 0 ? `$${totalAmount}` : '', // 첫 번째 행에만 총액
      new Date().toLocaleDateString(),
      'Approved & PDF Generated'
    ]);

    // 기존 데이터에 새 데이터 추가
    const updatedData = [...existingData, ...newRows];

    // Excel 파일 생성
    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.aoa_to_sheet(updatedData);
    
    // 스타일 적용 (헤더 행 굵게)
    const range = XLSX.utils.decode_range(newWorksheet['!ref'] || 'A1');
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (!newWorksheet[cellAddress]) continue;
      newWorksheet[cellAddress].s = {
        font: { bold: true },
        fill: { fgColor: { rgb: "E6E6FA" } }
      };
    }

    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Credit Card Receipts');

    // Excel 파일을 Buffer로 변환
    const excelBuffer = XLSX.write(newWorkbook, { type: 'buffer', bookType: 'xlsx' });

    // Firebase Storage에 업로드
    await uploadBytes(excelFileRef, excelBuffer, {
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });

    // 다운로드 URL 생성
    const downloadURL = await getDownloadURL(excelFileRef);

    return NextResponse.json({
      success: true,
      message: 'Excel file updated successfully',
      downloadURL: downloadURL,
      totalRows: updatedData.length
    });

  } catch (error) {
    console.error('Error creating/updating Excel file:', error);
    return NextResponse.json({
      success: false,
      error: 'Failed to create/update Excel file: ' + (error instanceof Error ? error.message : 'Unknown error')
    });
  }
}

// Excel 파일 다운로드 URL 가져오기
export async function GET() {
  try {
    const excelFileName = 'credit-card-receipts.xlsx';
    const excelFileRef = ref(storage, `excel/${excelFileName}`);
    
    const downloadURL = await getDownloadURL(excelFileRef);
    
    return NextResponse.json({
      success: true,
      downloadURL: downloadURL
    });
  } catch (error) {
    console.error('Error getting Excel file URL:', error);
    return NextResponse.json({
      success: false,
      error: 'Excel file not found'
    });
  }
}
