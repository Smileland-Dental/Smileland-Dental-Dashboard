import { NextRequest, NextResponse } from 'next/server';
import puppeteer from 'puppeteer';

// HTML 이스케이프 함수
function escapeHtml(text: string): string {
  const map: { [key: string]: string } = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  return text.replace(/[&<>"']/g, (m) => map[m]);
}

// 문자열 sanitize 함수
function sanitizeString(str: string, maxLength: number): string {
  if (!str) return '';
  const sanitized = String(str).trim().slice(0, maxLength);
  return sanitized.replace(/[<>]/g, '');
}

// 파일명 sanitize 함수 (파일명 인젝션 방지)
function sanitizeFilename(filename: string): string {
  return filename
    .replace(/[^a-zA-Z0-9._-]/g, '_')
    .replace(/\.\./g, '_')
    .slice(0, 255);
}

export async function POST(request: NextRequest) {
  try {
    // HTTPS 강제 (프로덕션 환경)
    const isProduction = process.env.NODE_ENV === 'production';
    if (isProduction) {
      const protocol = request.headers.get('x-forwarded-proto') || 
                       (request.nextUrl.protocol === 'https:' ? 'https' : 'http');
      
      if (protocol !== 'https') {
        const httpsUrl = request.nextUrl.clone();
        httpsUrl.protocol = 'https:';
        return NextResponse.redirect(httpsUrl, 301);
      }
    }

    const { inspectionDate, selectedOffice, selectedRestroom, restroomData } = await request.json();

    if (!inspectionDate || !selectedOffice || !selectedRestroom || !restroomData) {
      return NextResponse.json({ 
        success: false, 
        error: 'Inspection date, office, restroom, and data are required' 
      }, {
        status: 400
      });
    }

    // 입력 타입 검증
    if (typeof inspectionDate !== 'string' || 
        typeof selectedOffice !== 'string' || 
        typeof selectedRestroom !== 'string' ||
        typeof restroomData !== 'object' || 
        restroomData === null) {
      return NextResponse.json(
        { success: false, error: 'Invalid input format' },
        { status: 400 }
      );
    }

    // restroomData 크기 제한 (대략적인 크기 체크)
    const restroomDataSize = JSON.stringify(restroomData).length;
    if (restroomDataSize > 100000) { // 100KB 제한
      return NextResponse.json(
        { success: false, error: 'Data too large' },
        { status: 400 }
      );
    }

    // 날짜 형식 검증 (YYYY-MM-DD)
    const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
    if (!dateRegex.test(inspectionDate)) {
      return NextResponse.json(
        { success: false, error: 'Invalid date format' },
        { status: 400 }
      );
    }
    
    // Office 값 검증 (A, B, C만 허용)
    const validOffices = ['Bernard', 'California', 'Delano', 'Fresno', 'Ming', 'Ortho', 'Tulare', 'Visalia'];
    if (!validOffices.includes(selectedOffice)) {
      return NextResponse.json(
        { success: false, error: 'Invalid office selection' },
        { status: 400 }
      );
    }

    // Restroom 값 검증 (1, 2, 3만 허용)
    const validRestrooms = ['1', '2', '3'];
    if (!validRestrooms.includes(selectedRestroom)) {
      return NextResponse.json(
        { success: false, error: 'Invalid restroom selection' },
        { status: 400 }
      );
    }
    
    // 데이터 sanitize
    const safeInspectionDate = sanitizeString(inspectionDate, 20);
    const safeSelectedOffice = sanitizeString(selectedOffice, 100);
    const safeSelectedRestroom = sanitizeString(selectedRestroom, 50);

    // 컬럼 순서: Time, Check, Pick up Paper, Wipe Sinks and Mirrors, Wipe Toilets, Wipe Baby Table, Empty Trash, Toilet Paper, Soap, Toilet Seat Covers, Refresh Spray, Checked Time
    const COLUMN_NAMES = [
      'Time', 'Check', 'Pick up Paper', 'Wipe Sinks and Mirrors', 'Wipe Toilets',
      'Wipe Baby Table', 'Empty Trash', 'Toilet Paper', 'Soap',
      'Toilet Seat Covers', 'Refresh Spray', 'Checked Time'
    ];

    // 행 헤더 정의
    const ROW_HEADERS = [
      'Manager Inspection',
      '8 am', '9 am', '10 am', '11 am',
      'Manager Inspection',
      '12 pm', '1 pm', 'Sweep/Mop', '2 pm', '3 pm',
      'Manager Inspection',
      '4 pm', '5 pm', '6 pm', 'Sweep/Mop', '7 pm',
      'Deep Clean Manager Inspection'
    ];

    // HTML 템플릿 생성
    const html = `
      <!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>${escapeHtml(safeInspectionDate)}_${escapeHtml(safeSelectedOffice)}_Restroom ${escapeHtml(safeSelectedRestroom)} Inspection Log</title>
        <style>
          @media print {
            body { margin: 0.3in; font-size: 8px; line-height: 1.1; }
            table { font-size: 7px; }
            .header { font-size: 10px; margin-bottom: 10px; }
            th, td { padding: 2px 4px; }
          }
          
          body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background: white;
            color: #333;
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
          
          table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 15px;
          }
          
          th, td {
            border: 1px solid #333;
            padding: 4px 6px;
            text-align: center;
            font-size: 8px;
            vertical-align: middle;
          }
          
          th {
            background: #f0f0f0;
            font-weight: bold;
          }
          
          .spotless-header {
            background: #f3f3f3;
            font-weight: 700;
            color: #333;
          }
          
          .stocked-header {
            background: #f3f3f3;
            font-weight: 700;
            color: #333;
          }
          
          .spotless-subheader {
            background: #f0f0f0;
            font-size: 7px;
            font-weight: 400;
            color: #333;
          }
          
          .stocked-subheader {
            background: #f0f0f0;
            font-size: 7px;
            font-weight: 400;
            color: #333;
          }
          
          .footer {
            margin-top: 20px;
            text-align: center;
            font-size: 8px;
            color: #666;
            border-top: 1px solid #ddd;
            padding-top: 10px;
          }
        </style>
      </head>
      <body>
        <div class="header">
          <h1>Restroom Inspection Log</h1>
          <p>${escapeHtml(safeSelectedOffice)} - Restroom ${escapeHtml(safeSelectedRestroom)} (${escapeHtml(safeInspectionDate)})</p>
        </div>
        
        <table>
          <thead>
            <tr>
              <th colspan="2"></th>
              <th colspan="4" class="spotless-header">SPOTLESS</th>
              <th colspan="5" class="stocked-header">STOCKED</th>
              <th></th>
            </tr>
            <tr>
              <th>Time</th>
              <th>Check</th>
              <th>Pick up Paper</th>
              <th>Wipe Sinks and Mirrors</th>
              <th>Wipe Toilets</th>
              <th>Wipe Baby Table</th>
              <th>Empty Trash</th>
              <th>Toilet Paper</th>
              <th>Soap</th>
              <th>Toilet Seat Covers</th>
              <th>Refresh Spray</th>
              <th>Checked Time</th>
            </tr>
            <tr>
              <th colspan="2"></th>
              <th colspan="4" class="spotless-subheader">Perform each hour</th>
              <th colspan="5" class="stocked-subheader">Replenish as needed</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            ${ROW_HEADERS.map((header, rowIndex) => {
              const isMopRow = header.toLowerCase().includes('sweep/mop');
              const checkValue = sanitizeString(restroomData[`Row${rowIndex + 1}_Check`] || '', 50);
              const checkedTime = sanitizeString(restroomData[`Row${rowIndex + 1}_CheckedTime`] || '', 50);
              
              return `
                <tr>
                  <td>${escapeHtml(header)}</td>
                  <td>${escapeHtml(checkValue)}</td>
                  ${COLUMN_NAMES.slice(2, -1).map((_, colIndex) => {
                    if (isMopRow) {
                      return '<td></td>';
                    }
                    const cellValue = restroomData[`Row${rowIndex + 1}_Col${colIndex + 3}`];
                    return `<td>${cellValue === true ? '✔' : ''}</td>`;
                  }).join('')}
                  <td>${escapeHtml(checkedTime)}</td>
                </tr>
              `;
            }).join('')}
          </tbody>
        </table>
        
        <div class="footer">
          Generated: ${new Date().toLocaleDateString('en-US', {
            year: 'numeric',
            month: 'long',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit',
            hour12: true
          })}
        </div>
      </body>
      </html>
    `;

    // Puppeteer를 사용하여 PDF 생성
    let browser;
    try {
      browser = await puppeteer.launch({
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox']
      });
      
      const page = await browser.newPage();
      await page.setContent(html, { waitUntil: 'networkidle0' });
      
      // PDF 생성
      const pdf = await page.pdf({
        format: 'A4',
        landscape: true,
        margin: {
          top: '0.3in',
          right: '0.3in',
          bottom: '0.3in',
          left: '0.3in'
        },
        printBackground: true
      });
      
      await browser.close();
      
      // 파일명 sanitize (파일명 인젝션 방지)
      const safeFilename = sanitizeFilename(
        `${safeInspectionDate}_${safeSelectedOffice}_Restroom_${safeSelectedRestroom}_Inspection_Log.pdf`
      );
      
      // PDF 반환 (Buffer로 변환하여 타입 에러 해결)
      return new NextResponse(Buffer.from(pdf), {
        headers: {
          'Content-Type': 'application/pdf',
          'Content-Disposition': `attachment; filename="${safeFilename}"; filename*=UTF-8''${encodeURIComponent(safeFilename)}`,
          'Cache-Control': 'no-cache, no-store, must-revalidate',
          'Pragma': 'no-cache',
          'Expires': '0',
          'X-Content-Type-Options': 'nosniff',
          'X-Frame-Options': 'DENY',
          'X-XSS-Protection': '1; mode=block'
        }
      });
    } catch (pdfError: any) {
      if (browser) {
        await browser.close();
      }
      // 에러 메시지 노출 최소화 (보안)
      console.error('PDF generation error:', pdfError);
      throw new Error('PDF generation failed');
    }

  } catch (error: any) {
    // 에러 정보 노출 최소화
    console.error('Request error:', error);
    return NextResponse.json({ 
      success: false, 
      error: 'Failed to generate restroom inspection log' 
    }, {
      status: 500,
      headers: {
        'X-Content-Type-Options': 'nosniff',
        'X-Frame-Options': 'DENY',
        'X-XSS-Protection': '1; mode=block'
      }
    });
  }
}
