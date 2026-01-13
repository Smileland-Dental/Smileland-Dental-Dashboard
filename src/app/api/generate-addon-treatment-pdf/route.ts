import { NextRequest, NextResponse } from 'next/server';
import puppeteer from 'puppeteer';

// 🔒 보안: 요청 크기 제한 (5MB)
const MAX_REQUEST_SIZE = 5 * 1024 * 1024; // 5MB
// 🔒 보안: PDF 생성 데이터 크기 제한 (1MB)
const MAX_PDF_DATA_SIZE = 1024 * 1024; // 1MB
// 🔒 보안: 타임아웃 설정 (30초)
const PDF_GENERATION_TIMEOUT = 30000; // 30초

// 🔒 보안: HTML 이스케이프 함수 (XSS 방지)
function escapeHtml(text: string | undefined | null): string {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// 🔒 보안: 입력 검증 및 sanitization 함수
function sanitizeString(input: any, maxLength: number = 500): string {
  if (typeof input !== 'string') return '';
  // 길이 제한
  if (input.length > maxLength) {
    return input.substring(0, maxLength);
  }
  return input;
}

// 🔒 보안: PDF 파일명 sanitization
function sanitizePdfFilename(filename: string): string {
  if (!filename || typeof filename !== 'string') return 'document.pdf';
  // 경로 탐색 공격 방지
  let sanitized = filename.replace(/\.\./g, '_');
  // 허용된 문자만 유지
  sanitized = sanitized.replace(/[^a-zA-Z0-9._-]/g, '_');
  // 길이 제한 (255자)
  if (sanitized.length > 255) {
    sanitized = sanitized.substring(0, 255);
  }
  return sanitized;
}

// 🔒 보안: PDF 생성 데이터 크기 검증
function validatePdfGeneration(dataSize: number): boolean {
  return dataSize <= MAX_PDF_DATA_SIZE;
}

// 🔒 보안: 날짜 형식 검증
function validateDate(date: string): boolean {
  if (!date || typeof date !== 'string') return false;
  // YYYY-MM-DD 형식 확인
  const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
  return dateRegex.test(date);
}

// 🔒 보안: 환자 데이터 검증
function validatePatientRow(row: any): boolean {
  if (!row || typeof row !== 'object') return false;
  // 필수 필드 확인
  if (typeof row.rowNumber !== 'number' || row.rowNumber < 1 || row.rowNumber > 10000) return false;
  if (typeof row.patientName !== 'string' || row.patientName.length > 100) return false;
  if (typeof row.dob !== 'string' || row.dob.length > 20) return false;
  if (typeof row.time !== 'string' || row.time.length > 20) return false;
  return true;
}

// 🔒 보안: 안전한 Puppeteer 옵션
function getSecurePuppeteerOptions() {
  return {
    headless: true,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-accelerated-2d-canvas',
      '--no-first-run',
      '--no-zygote',
      '--disable-gpu'
    ]
  };
}

export async function POST(request: NextRequest) {
  let browser: Awaited<ReturnType<typeof puppeteer.launch>> | undefined;
  try {
    // 🔒 보안: 인증 확인 (예시 - 실제 인증 방식에 맞게 수정 필요)
    // const authHeader = request.headers.get('authorization');
    // if (!authHeader || !isValidToken(authHeader)) {
    //   return NextResponse.json({ success: false, error: 'Unauthorized' }, { status: 401 });
    // }

    // 🔒 보안: 요청 크기 확인
    const contentLength = request.headers.get('content-length');
    if (contentLength && parseInt(contentLength) > MAX_REQUEST_SIZE) {
      return NextResponse.json({ success: false, error: 'Request too large' }, { status: 413 });
    }

    // 🔒 보안: JSON 파싱 에러 처리
    let requestData;
    try {
      requestData = await request.json();
    } catch (parseError) {
      return NextResponse.json({ success: false, error: 'Invalid JSON format' }, { status: 400 });
    }

    const { dutyDate, selectedOffice, patientRows } = requestData;

    if (!dutyDate || !selectedOffice || !patientRows) {
      return NextResponse.json({ success: false, error: 'Duty date, office, and patient data are required' }, { status: 400 });
    }

    // 🔒 보안: 배열 데이터 검증
    if (!Array.isArray(patientRows)) {
      return NextResponse.json({ success: false, error: 'Invalid data format' }, { status: 400 });
    }

    if (patientRows.length > 1000) {
      return NextResponse.json({ success: false, error: 'Too many patient rows' }, { status: 400 });
    }

    // 🔒 보안: 각 환자 행 데이터 검증
    for (const row of patientRows) {
      if (!validatePatientRow(row)) {
        return NextResponse.json({ success: false, error: 'Invalid patient data format' }, { status: 400 });
      }
    }

    // 데이터 안전성 검증
    const safeDutyDate = sanitizeString(dutyDate, 50);
    const safeSelectedOffice = sanitizeString(selectedOffice, 50);
    
    // 🔒 보안: 날짜 형식 검증
    if (!validateDate(safeDutyDate)) {
      return NextResponse.json({ success: false, error: 'Invalid date format. Expected YYYY-MM-DD' }, { status: 400 });
    }
    
    // 🔒 보안: 오피스 값 검증
    const allowedOffices = ['Ming', 'Bernard', 'Delano', 'Tulare', 'Visalia', 'Fresno', 'California', 'Ortho'];
    if (!allowedOffices.includes(safeSelectedOffice)) {
      return NextResponse.json({ success: false, error: 'Invalid office value' }, { status: 400 });
    }
    
    // PDF 생성 데이터 크기 검증
    const dataSize = JSON.stringify({ dutyDate, selectedOffice, patientRows }).length;
    if (!validatePdfGeneration(dataSize)) {
      return NextResponse.json({ success: false, error: 'Data size too large for PDF generation' }, { status: 413 });
    }

    // 시간을 12시간 형식으로 변환하는 함수
    const convertTo12Hour = (timeStr: string): string => {
      if (!timeStr || timeStr === '-') return '-';
      
      // 🔒 보안: 이미 AM/PM이 포함된 경우 정규식으로 검증 후 반환 (중복 변환 방지)
      // 정규식으로 시간 형식 검증: "HH:MM AM" 또는 "HH:MM PM" 형식만 허용
      const timeWithAmPmRegex = /^\d{1,2}:\d{2}\s*(AM|PM)$/i;
      if (timeWithAmPmRegex.test(timeStr.trim())) {
        return timeStr.trim();
      }
      
      try {
        const [hours, minutes] = timeStr.split(':');
        const hour = parseInt(hours);
        const min = minutes || '00';
        if (hour === 0) return `12:${min} AM`;
        if (hour < 12) return `${hour}:${min} AM`;
        if (hour === 12) return `12:${min} PM`;
        return `${hour - 12}:${min} PM`;
      } catch {
        return timeStr;
      }
    };

    // HTML 템플릿 생성
    const html = `
      <!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>${escapeHtml(safeDutyDate)}_${escapeHtml(safeSelectedOffice)}_Add On Treatment</title>
        <style>
          @media print {
            body { margin: 0.3in; font-size: 8px; line-height: 1.1; }
            table { font-size: 7px; }
            .header { font-size: 10px; margin-bottom: 10px; }
            th, td { padding: 2px 4px; }
          }
          
          body {
            font-family: 'Segoe UI', 'Roboto', Arial, sans-serif;
            margin: 32px;
            color: #222;
            background: #fff;
          }
          
          .header {
            text-align: center;
            border-bottom: 2px solid #333;
            padding-bottom: 10px;
            margin-bottom: 20px;
          }
          
          .header h1 {
            margin: 0;
            font-size: 18px;
            font-weight: bold;
          }
          
          .info-section {
            display: flex;
            justify-content: space-between;
            margin-bottom: 15px;
            font-size: 10px;
          }
          
          table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 15px;
          }
          
          th, td {
            border: 1px solid #333;
            padding: 4px 6px;
            text-align: left;
            font-size: 8px;
            vertical-align: top;
          }
          
          th {
            background: #f0f0f0;
            font-weight: bold;
          }
        </style>
      </head>
      <body>
        <div class="header">
          <h1>Add-On Treatment</h1>
        </div>
        
        <div class="info-section">
          <div><strong>Date:</strong> ${escapeHtml(safeDutyDate)}</div>
          <div><strong>Location:</strong> ${escapeHtml(safeSelectedOffice)}</div>
          <div><strong>Generated:</strong> ${new Date().toLocaleDateString('en-US', {
            year: 'numeric',
            month: 'long',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit',
            hour12: true
          })}</div>
        </div>
        
        <table>
          <thead>
            <tr>
              <th style="width: 30px;">No.</th>
              <th>Name of Patient</th>
              <th>DOB</th>
              <th>Time</th>
            </tr>
          </thead>
          <tbody>
            ${patientRows.map((row: any, index: number) => {
              // 🔒 보안: 각 필드 검증 및 sanitization
              const safePatientName = sanitizeString(row.patientName || '', 100);
              const safeDob = sanitizeString(row.dob || '', 20);
              const safeTime = sanitizeString(row.time || '', 20);
              
              return `
              <tr>
                <td>${index + 1}</td>
                <td>${escapeHtml(safePatientName) || '-'}</td>
                <td>${escapeHtml(safeDob) || '-'}</td>
                <td>${escapeHtml(convertTo12Hour(safeTime)) || '-'}</td>
              </tr>
            `;
            }).join('')}
          </tbody>
        </table>
      </body>
      </html>
    `;

    // Puppeteer로 PDF 생성 (보안 강화)
    // 🔒 보안: 타임아웃 설정
    const pdfGenerationPromise = (async () => {
      browser = await puppeteer.launch(getSecurePuppeteerOptions());
      
      const page = await browser.newPage();
      
      // 🔒 보안: 페이지 네비게이션 타임아웃 설정
      page.setDefaultNavigationTimeout(10000); // 10초
      page.setDefaultTimeout(10000); // 10초
      
      await page.setContent(html, { waitUntil: 'networkidle0' });
      
      // 🔒 보안: PDF 생성 타임아웃
      const pdf = await Promise.race([
        page.pdf({
          format: 'A4',
          printBackground: true,
          margin: {
            top: '10mm',
            right: '10mm',
            bottom: '10mm',
            left: '10mm'
          },
          preferCSSPageSize: false,
          displayHeaderFooter: false
        }),
        new Promise<never>((_, reject) => 
          setTimeout(() => reject(new Error('PDF generation timeout')), PDF_GENERATION_TIMEOUT)
        )
      ]);
      
      await browser.close();
      browser = undefined;
      
      return pdf;
    })();

    const pdf = await pdfGenerationPromise;
    
    // 파일명 생성 (보안 검증)
    const filename = sanitizePdfFilename(`4) ${safeDutyDate}_${safeSelectedOffice}_Add On Treatment.pdf`);
    
    // 파일명 생성 (검증된 값 사용)
    // 🔒 보안: 파일명 추가 sanitization (경로 탐색 공격 방지)
    const safeFilename = sanitizePdfFilename(`4) ${safeDutyDate}_${safeSelectedOffice}_Add On Treatment.pdf`)
      .replace(/\.\./g, '_') // 경로 탐색 방지
      .substring(0, 255); // 파일명 길이 제한
    
    return new NextResponse(Buffer.from(pdf), {
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${safeFilename}"`,
        'X-Content-Type-Options': 'nosniff',
        'X-Frame-Options': 'DENY',
        'X-XSS-Protection': '1; mode=block',
        'Content-Security-Policy': "default-src 'none'",
        'Referrer-Policy': 'no-referrer',
        'Cache-Control': 'no-cache, no-store, must-revalidate',
        'Pragma': 'no-cache',
        'Expires': '0'
      }
    });

  } catch (error) {
    // 🔒 보안: 에러 로깅 (서버 로그에만 기록)
    console.error('Add-on treatment PDF generation error:', error instanceof Error ? error.message : 'Unknown error');
    // 🔒 보안: 에러 메시지에 민감한 정보 노출 방지
    const isTimeout = error instanceof Error && error.message.includes('timeout');
    return NextResponse.json({ 
      success: false, 
      error: isTimeout ? 'PDF generation timed out. Please try again.' : 'Failed to generate add-on treatment PDF' 
    }, { status: isTimeout ? 408 : 500 });
  } finally {
    // 🔒 보안: 브라우저 리소스 정리
    if (browser) {
      try {
        await browser.close();
      } catch (closeError) {
        console.error('Error closing browser:', closeError);
      }
    }
  }
}
