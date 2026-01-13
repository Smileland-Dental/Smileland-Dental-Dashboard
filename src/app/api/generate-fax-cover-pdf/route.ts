import { NextRequest, NextResponse } from 'next/server';
import puppeteer from 'puppeteer';

// 🔒 보안: 요청 크기 제한 (5MB)
const MAX_REQUEST_SIZE = 5 * 1024 * 1024; // 5MB
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

// 🔒 보안: 입력 검증 함수
function validateInput(input: any, maxLength: number = 500): boolean {
  if (typeof input !== 'string') return false;
  if (input.length > maxLength) return false;
  return true;
}

// 🔒 보안: 날짜 형식 검증
function validateDate(date: string): boolean {
  if (!date || typeof date !== 'string') return false;
  // MM/DD/YYYY 또는 YYYY-MM-DD 형식 확인
  const dateRegex1 = /^\d{2}\/\d{2}\/\d{4}$/;
  const dateRegex2 = /^\d{4}-\d{2}-\d{2}$/;
  return dateRegex1.test(date) || dateRegex2.test(date);
}

// 🔒 보안: 오피스 값 검증
function validateOffice(office: string | undefined): boolean {
  if (!office || typeof office !== 'string') return false;
  const allowedOffices = ['Ming', 'Bernard', 'Delano', 'Tulare', 'Visalia', 'Fresno', 'California', 'Ortho'];
  return allowedOffices.includes(office);
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

    const { faxDate, selectedOffice, formData, tableData, productionData, todayData, nextDayData, callLogData, supervisorData } = requestData;

    // 🔒 보안: 서버 측 검증
    if (!faxDate || !formData) {
      return NextResponse.json({ success: false, error: 'Date and form data are required' }, { status: 400 });
    }

    // 🔒 보안: 날짜 형식 검증
    if (!validateDate(faxDate) && (!formData.date || !validateDate(formData.date))) {
      return NextResponse.json({ success: false, error: 'Invalid date format' }, { status: 400 });
    }

    // 🔒 보안: 오피스 값 검증
    if (!validateOffice(selectedOffice)) {
      return NextResponse.json({ success: false, error: 'Invalid office value' }, { status: 400 });
    }

    // 🔒 보안: 배열 데이터 검증
    if (!Array.isArray(tableData) || !Array.isArray(productionData)) {
      return NextResponse.json({ success: false, error: 'Invalid data format' }, { status: 400 });
    }

    if (tableData.length > 100 || productionData.length > 100) {
      return NextResponse.json({ success: false, error: 'Data array too large' }, { status: 400 });
    }

    // 🔒 보안: 입력 검증 및 이스케이프 (XSS 방지)
    const safeFaxDate = escapeHtml(faxDate);
    const safeSelectedOffice = escapeHtml(selectedOffice || '');
    
    const safeFormData = {
      date: escapeHtml(formData.date || ''),
      officeTimeCheckIn: escapeHtml(formData.officeTimeCheckIn || ''),
      officeName: escapeHtml(formData.officeName || ''),
      timeCheckOut: escapeHtml(formData.timeCheckOut || ''),
      name: escapeHtml(formData.name || ''),
    };

    // 테이블 데이터 HTML 생성 (검증 및 이스케이프)
    const tableRows = tableData.map((row: any, index: number) => {
      const safeNameOfForm = validateInput(row.nameOfForm, 200) ? escapeHtml(row.nameOfForm) : '';
      const safeOtherText = validateInput(row.otherText, 200) ? escapeHtml(row.otherText) : '';
      const safeQty = validateInput(row.qty, 20) ? escapeHtml(row.qty) : '';
      
      const formName = safeNameOfForm === 'Other:' 
        ? `Other: ${safeOtherText}` 
        : safeNameOfForm;
      return `
        <tr>
          <td style="text-align: center; font-weight: bold;">${index + 1}</td>
          <td>${formName}</td>
          <td style="text-align: center;">${safeQty}</td>
        </tr>
      `;
    }).join('');

    // Total Pages 계산 (검증된 데이터 사용)
    const totalPages = tableData.reduce((sum: number, row: any) => {
      const qtyStr = validateInput(row.qty, 20) ? row.qty : '';
      const qty = parseFloat(qtyStr) || 0;
      // 합계가 음수이거나 너무 크면 0으로 제한
      return Math.max(0, Math.min(sum + qty, 999999));
    }, 0);

    // Production 데이터 HTML 생성 (검증 및 이스케이프)
    const productionRows = productionData.map((row: any, index: number) => {
      const safeDate = validateInput(row.date, 20) ? escapeHtml(row.date) : '';
      const safeNote = validateInput(row.note, 500) ? escapeHtml(row.note) : '';
      const safeStatus = validateInput(row.status, 50) ? escapeHtml(row.status) : '';
      return `
        <tr>
          <td>${safeDate}</td>
          <td>${safeNote}</td>
          <td style="text-align: center;">${safeStatus}</td>
        </tr>
      `;
    }).join('');

    // 🔒 보안: Today, Next Day, Call Log, Supervisor 데이터 검증 및 이스케이프
    const safeTodayData = {
      addOns: validateInput(todayData?.addOns, 500) ? escapeHtml(todayData.addOns) : '',
      noShows: validateInput(todayData?.noShows, 500) ? escapeHtml(todayData.noShows) : '',
      seen: validateInput(todayData?.seen, 500) ? escapeHtml(todayData.seen) : '',
    };

    const safeNextDayData = {
      opener: validateInput(nextDayData?.opener, 200) ? escapeHtml(nextDayData.opener) : '',
      closer: validateInput(nextDayData?.closer, 200) ? escapeHtml(nextDayData.closer) : '',
    };

    const safeCallLogData = {
      whoCalled: validateInput(callLogData?.whoCalled, 500) ? escapeHtml(callLogData.whoCalled) : '',
      appointmentsMade: validateInput(callLogData?.appointmentsMade, 500) ? escapeHtml(callLogData.appointmentsMade) : '',
    };

    const safeSupervisorData = {
      officeSupervisorManager: validateInput(supervisorData?.officeSupervisorManager, 200) ? escapeHtml(supervisorData.officeSupervisorManager) : '',
      spokeWith: validateInput(supervisorData?.spokeWith, 200) ? escapeHtml(supervisorData.spokeWith) : '',
      checkOutBy: validateInput(supervisorData?.checkOutBy, 200) ? escapeHtml(supervisorData.checkOutBy) : '',
    };

    // HTML 템플릿 생성
    const html = `
      <!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <meta http-equiv="Content-Security-Policy" content="default-src 'self'; script-src 'none'; object-src 'none'; base-uri 'self'; style-src 'unsafe-inline';">
        <title>${safeFaxDate}_${safeSelectedOffice}_End of Day Fax Cover</title>
        <style>
          @media print {
            body { margin: 0.2in; font-size: 12px; line-height: 1.3; }
            table { font-size: 9px; }
            .header { font-size: 16px; margin-bottom: 5px; }
            th, td { padding: 3px 4px; }
          }
          
          body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 10px;
            background: white;
            color: #333;
            font-size: 12px;
            line-height: 1.3;
          }
          
          .header {
            text-align: center;
            border-bottom: 1px solid #333;
            padding-bottom: 5px;
            margin-bottom: 8px;
          }
          
          .header h1 {
            margin: 0;
            font-size: 24px;
            font-weight: bold;
          }
          
          .header p {
            margin: 2px 0 0 0;
            font-size: 12px;
            color: #666;
          }
          
          .info-section {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            margin-bottom: 8px;
            font-size: 12px;
          }
          
          .info-item {
            display: flex;
            gap: 3px;
          }
          
          .info-label {
            font-weight: bold;
          }
          
          table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 5px;
            margin-bottom: 5px;
            font-size: 9px;
          }
          
          th, td {
            border: 1px solid #333;
            padding: 3px 4px;
            text-align: left;
            font-size: 9px;
            vertical-align: top;
            line-height: 1.2;
          }
          
          th {
            background: #f0f0f0;
            font-weight: bold;
          }
          
          .section {
            margin-top: 8px;
            margin-bottom: 5px;
          }
          
          .section-title {
            font-size: 15px;
            font-weight: bold;
            margin-bottom: 4px;
            border-bottom: 1px solid #ddd;
            padding-bottom: 2px;
          }
          
          .section-content {
            font-size: 12px;
            margin-bottom: 4px;
          }
          
          .section-row {
            display: flex;
            gap: 10px;
            margin-bottom: 3px;
          }
          
          .section-field {
            flex: 1;
          }
          
          .section-label {
            font-weight: bold;
            margin-right: 3px;
          }
        </style>
      </head>
      <body>
        <div class="header">
          <h1>End of Day Fax Cover</h1>
          <p>(Check out only when leaving the office)</p>
        </div>
        
        <div class="info-section">
          <div class="info-item">
            <span class="info-label">Date:</span>
            <span>${safeFormData.date}</span>
          </div>
          <div class="info-item">
            <span class="info-label">Office:</span>
            <span>${safeSelectedOffice}</span>
          </div>
        </div>

        <div class="info-section">
          <div class="info-item">
            <span class="info-label">Time Check In:</span>
            <span>${safeFormData.officeTimeCheckIn}</span>
          </div>
          <div class="info-item">
            <span class="info-label">Name:</span>
            <span>${safeFormData.officeName}</span>
          </div>
          <div class="info-item">
            <span class="info-label">Time Check Out:</span>
            <span>${safeFormData.timeCheckOut}</span>
          </div>
          <div class="info-item">
            <span class="info-label">Name:</span>
            <span>${safeFormData.name}</span>
          </div>
        </div>
        
        <table>
          <thead>
            <tr>
              <th style="width: 60px; text-align: center;">No.</th>
              <th>Name of Form</th>
              <th style="width: 100px; text-align: center;">Qty</th>
            </tr>
          </thead>
          <tbody>
            ${tableRows}
            <tr style="background-color: #f8f9fa; font-weight: bold;">
              <td colspan="2" style="text-align: center;">Total Pages</td>
              <td style="text-align: center; font-weight: bold;">${totalPages}</td>
            </tr>
          </tbody>
        </table>

        <div style="display: flex; gap: 10px; margin-top: 8px;">
          <div style="flex: 1;">
            <div class="section-title">Production</div>
            <table style="margin-bottom: 5px;">
              <thead>
                <tr>
                  <th style="width: 80px;">Date</th>
                  <th>Note</th>
                  <th style="width: 60px; text-align: center;">Status</th>
                </tr>
              </thead>
              <tbody>
                ${productionRows}
              </tbody>
            </table>
          </div>
          
          <div style="flex: 1;">
            <div class="section-title">Today</div>
            <div class="section-content">
              <div class="section-row">
                <span class="section-label">Add On's:</span>
                <span>${safeTodayData.addOns}</span>
              </div>
              <div class="section-row">
                <span class="section-label">No Shows:</span>
                <span>${safeTodayData.noShows}</span>
              </div>
              <div class="section-row">
                <span class="section-label">Seen:</span>
                <span>${safeTodayData.seen}</span>
              </div>
            </div>
            
            <div class="section-title" style="margin-top: 5px;">Next Day</div>
            <div class="section-content">
              <div class="section-row">
                <span class="section-label">Opener:</span>
                <span>${safeNextDayData.opener}</span>
              </div>
              <div class="section-row">
                <span class="section-label">Closer:</span>
                <span>${safeNextDayData.closer}</span>
              </div>
            </div>
          </div>
        </div>

        <div style="display: flex; gap: 10px; margin-top: 5px;">
          <div style="flex: 1;">
            <div class="section-title">Call Log</div>
            <div class="section-content">
              <div class="section-row">
                <span class="section-label">Who called:</span>
                <span>${safeCallLogData.whoCalled}</span>
              </div>
              <div class="section-row">
                <span class="section-label">Appointments made:</span>
                <span>${safeCallLogData.appointmentsMade}</span>
              </div>
            </div>
          </div>
          
          <div style="flex: 1;">
            <div class="section-title">Office Supervisor/Manager</div>
            <div class="section-content">
              <div class="section-row">
                <span class="section-label">Supervisor/Manager:</span>
                <span>${safeSupervisorData.officeSupervisorManager}</span>
              </div>
              <div class="section-row">
                <span class="section-label">Spoke with:</span>
                <span>${safeSupervisorData.spokeWith}</span>
              </div>
              <div class="section-row">
                <span class="section-label">Check out by:</span>
                <span>${safeSupervisorData.checkOutBy}</span>
              </div>
            </div>
          </div>
        </div>
      </body>
      </html>
    `;

    // Puppeteer로 PDF 생성 (보안 강화)
    // 🔒 보안: 타임아웃 설정
    const pdfGenerationPromise = (async () => {
      browser = await puppeteer.launch({
        headless: true,
        args: [
          // ⚠️ 주의: --no-sandbox는 보안상 위험하지만 일부 서버 환경에서 필요할 수 있음
          // 가능하면 Docker 컨테이너나 적절한 권한 설정으로 sandbox를 활성화하는 것을 권장
          '--no-sandbox',
          '--disable-setuid-sandbox',
          '--disable-dev-shm-usage',
          '--disable-accelerated-2d-canvas',
          '--no-first-run',
          '--no-zygote',
          '--disable-gpu',
          '--disable-web-security', // 필요시에만 사용
          '--disable-features=IsolateOrigins,site-per-process' // 필요시에만 사용
        ]
      });
      
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
            top: '5mm',
            right: '5mm',
            bottom: '5mm',
            left: '5mm'
          },
          preferCSSPageSize: false,
          displayHeaderFooter: false,
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
    
    // 파일명 생성 (검증된 값 사용)
    // 🔒 보안: 파일명 추가 sanitization (경로 탐색 공격 방지)
    const safeFilename = `1) ${safeFaxDate}_${safeSelectedOffice || 'Unknown'}_End of Day Fax Cover.pdf`
      .replace(/[^a-zA-Z0-9._-]/g, '_')
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
        'Referrer-Policy': 'no-referrer'
      }
    });

  } catch (error) {
    // 🔒 보안: 에러 로깅 (서버 로그에만 기록)
    console.error('Fax cover PDF generation error:', error instanceof Error ? error.message : 'Unknown error');
    // 🔒 보안: 에러 메시지에 민감한 정보 노출 방지
    // 클라이언트에는 일반적인 에러 메시지만 반환
    const isTimeout = error instanceof Error && error.message.includes('timeout');
    return NextResponse.json({ 
      success: false, 
      error: isTimeout ? 'PDF generation timed out. Please try again.' : 'Failed to generate fax cover PDF' 
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
