import { NextRequest, NextResponse } from 'next/server';

export async function POST(request: NextRequest) {
  try {
    // 이 API는 클라이언트에서 직접 Firebase Storage를 사용하도록 변경
    // 서버에서는 PDF 생성만 하고, 클라이언트에서 Storage에 저장
    return NextResponse.json({
      success: true,
      message: 'PDF should be saved from client side',
    });
  } catch (error) {
    console.error('Error in save-pdf API:', error);
    return NextResponse.json(
      {
        success: false,
        error: 'Failed: ' + (error as Error).message,
      },
      { status: 500 }
    );
  }
}