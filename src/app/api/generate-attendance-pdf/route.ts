import { NextRequest, NextResponse } from 'next/server';
import puppeteer from 'puppeteer';

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
  const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
  if (!dateRegex.test(date)) return false;
  const parsedDate = new Date(date);
  return !isNaN(parsedDate.getTime());
}

// 🔒 보안: 오피스 값 검증
function validateOffice(office: string | undefined): boolean {
  if (!office) return true; // 선택적 필드
  const allowedOffices = ['Ming', 'Bernard', 'Delano', 'Tulare', 'Visalia', 'Fresno', 'California', 'Ortho'];
  return allowedOffices.includes(office);
}

export async function POST(request: NextRequest) {
  let browser;
  try {
    const { date, office, filledBy, checkedBy, staffData, doctorData } = await request.json();

    // 🔒 보안: 서버 측 검증
    if (!date || !staffData || !doctorData) {
      return NextResponse.json({ success: false, error: 'Date and data are required' }, { status: 400 });
    }

    // 🔒 보안: 날짜 형식 검증
    if (!validateDate(date)) {
      return NextResponse.json({ success: false, error: 'Invalid date format' }, { status: 400 });
    }

    // 🔒 보안: 오피스 값 검증
    // 🔒 보안: 오피스는 필수 값
    if (!office) {
      return NextResponse.json({ success: false, error: 'Office is required' }, { status: 400 });
    }
    const selectedOffice = office;
    if (!validateOffice(selectedOffice)) {
      return NextResponse.json({ success: false, error: 'Invalid office value' }, { status: 400 });
    }

    // 🔒 보안: 배열 타입 및 길이 검증 (DoS 방지)
    if (!Array.isArray(staffData) || !Array.isArray(doctorData)) {
      return NextResponse.json({ success: false, error: 'Invalid data format' }, { status: 400 });
    }
    if (staffData.length > 1000 || doctorData.length > 100) {
      return NextResponse.json({ success: false, error: 'Data array too large' }, { status: 400 });
    }

    // 🔒 보안: 사용자 입력 검증 및 이스케이프
    const safeFilledBy = typeof filledBy === 'string' && validateInput(filledBy, 100) ? filledBy : '';
    const safeCheckedBy = typeof checkedBy === 'string' && validateInput(checkedBy, 100) ? checkedBy : '';

    // HTML 생성 (generateHTML 함수 내부에서 이스케이프 처리)
    const html = generateHTML(date, selectedOffice, safeFilledBy, safeCheckedBy, staffData, doctorData);

    // Puppeteer로 PDF 생성 (안정적인 설정)
    try {
      browser = await puppeteer.launch({
        headless: true,
        args: [
          '--no-sandbox',
          '--disable-setuid-sandbox',
          '--disable-dev-shm-usage', // 메모리 부족 문제 해결
          '--disable-accelerated-2d-canvas',
          '--no-first-run',
          '--no-zygote',
          '--disable-gpu'
        ],
        timeout: 30000, // 30초 타임아웃
      });
      
      const page = await browser.newPage();
      
      // 페이지 타임아웃 설정
      page.setDefaultTimeout(30000);
      
      // 🔒 보안: CSP는 유지하되, 스타일시트는 허용
      await page.setContent(html, {
        waitUntil: 'load', // load로 변경 (networkidle0은 외부 리소스 대기)
        timeout: 30000
      });
      
      // 추가 대기 (CSS 및 스타일 렌더링 완료)
      await new Promise(resolve => setTimeout(resolve, 2000));
      
      // 테이블이 제대로 렌더링되었는지 확인 (디버깅용)
      const tableExists = await page.evaluate(() => {
        return document.querySelector('table') !== null;
      });
      
      if (!tableExists) {
        console.warn('Warning: Table element not found in HTML');
      }
      
      // PDF 생성
      const pdfBuffer = await page.pdf({
        format: 'Letter',
        landscape: true,
        printBackground: false, // 흑백 출력
        margin: {
          top: '20px',
          right: '20px',
          bottom: '20px',
          left: '20px'
        },
        preferCSSPageSize: false,
        displayHeaderFooter: false,
        timeout: 30000,
      });
      
      // 🔒 보안: 파일명 이스케이프 (파일명 주입 공격 방지)
      const safeFilename = escapeHtml(`${date}_${selectedOffice || 'unknown'}_Attendance Tract.pdf`);
      const filename = `3) ${safeFilename}`;
      
      // Buffer를 Uint8Array로 변환하여 NextResponse에 전달
      return new NextResponse(new Uint8Array(pdfBuffer), {
        headers: {
          'Content-Type': 'application/pdf',
          'Content-Disposition': `attachment; filename="${filename.replace(/"/g, '\\"')}"`,
          // 🔒 보안: 추가 보안 헤더
          'X-Content-Type-Options': 'nosniff',
          'X-Frame-Options': 'DENY',
          'X-XSS-Protection': '1; mode=block'
        }
      });
      
    } catch (puppeteerError) {
      console.error('Puppeteer PDF generation error:', puppeteerError);
      // 🔒 보안: 폴백 제거 - PDF 생성 실패 시 에러만 반환 (HTML 반환하지 않음)
      return NextResponse.json({ 
        success: false, 
        error: 'PDF generation failed. Please try again later.' 
      }, { status: 500 });
    } finally {
      // 브라우저 정리
      if (browser) {
        await browser.close().catch(err => {
          console.error('Error closing browser:', err);
        });
      }
    }

  } catch (error) {
    console.error('Attendance PDF generation error:', error);
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    
    // 브라우저가 열려있으면 정리
    if (browser) {
      await browser.close().catch(err => {
        console.error('Error closing browser in catch block:', err);
      });
    }
    
    return NextResponse.json({ 
      success: false, 
      error: 'Failed to generate PDF: ' + errorMessage
    }, { status: 500 });
  }
}

function generateHTML(date: string, office: string, filledBy: string, checkedBy: string, staffData: any[], doctorData: any[]): string {
  // position별 그룹핑
  const staffByPosition: { [key: string]: any[] } = {};
  staffData.forEach(row => {
    let pos = row.position || '';
    if (pos === 'Dental Assistant') pos = 'DA';
    if (!staffByPosition[pos]) staffByPosition[pos] = [];
    staffByPosition[pos].push(row);
  });

  // Staff 테이블 HTML 생성
  let staffTableHTML = `
    <table style="width: 100%; border-collapse: collapse; font-size: 7pt; font-family: Arial;">
      <thead>
        <tr style="background-color: #e3f2fd; font-weight: bold;">
          <th style="border: 1px solid #000; padding: 2px; text-align: center;">Name</th>
          <th style="border: 1px solid #000; padding: 2px; text-align: center;">Present</th>
          <th style="border: 1px solid #000; padding: 2px; text-align: center;">Start Shift Tardy (Min)</th>
          <th style="border: 1px solid #000; padding: 2px; text-align: center;">Late from Lunch (Min)</th>
          <th style="border: 1px solid #000; padding: 2px; text-align: center;">Needs Clock Adj.</th>
          <th style="border: 1px solid #000; padding: 2px; text-align: center;">Overtime</th>
          <th style="border: 1px solid #000; padding: 2px; text-align: center;">OT Corp Authorized By</th>
          <th style="border: 1px solid #000; padding: 2px; text-align: center;">Sub. at Another Office</th>
          <th style="border: 1px solid #000; padding: 2px; text-align: center;">Incident Description</th>
          <th style="border: 1px solid #000; padding: 2px; text-align: center;">Notes</th>
        </tr>
      </thead>
      <tbody>
  `;

  Object.keys(staffByPosition).forEach(pos => {
    const rows = staffByPosition[pos];
    const presentCount = rows.filter(r => r.present === true || r.present === 'TRUE').length;
    
    // Position 구분 행
    staffTableHTML += `
      <tr style="background-color: #fce4ec; font-weight: bold;">
        <td colspan="4" style="border: 1px solid #000; padding: 2px;"></td>
        <td style="border: 1px solid #000; padding: 2px; text-align: center; font-size: 8pt;">${pos}</td>
        <td style="border: 1px solid #000; padding: 2px; text-align: center; font-size: 8pt;">${presentCount}</td>
        <td colspan="4" style="border: 1px solid #000; padding: 2px;"></td>
      </tr>
    `;
    
    // 직원 행
    rows.forEach(row => {
      // 🔒 보안: 입력 검증 및 이스케이프
      const safeName = validateInput(row.name) ? escapeHtml(row.name) : '';
      if (safeName && safeName.trim()) {
        staffTableHTML += `
          <tr>
            <td style="border: 1px solid #000; padding: 2px;">${safeName}</td>
            <td style="border: 1px solid #000; padding: 2px; text-align: center;">${(row.present === true || row.present === 'TRUE') ? '✔' : ''}</td>
            <td style="border: 1px solid #000; padding: 2px;">${validateInput(row.startTardy) ? escapeHtml(row.startTardy || '') : ''}</td>
            <td style="border: 1px solid #000; padding: 2px;">${validateInput(row.lateLunch) ? escapeHtml(row.lateLunch || '') : ''}</td>
            <td style="border: 1px solid #000; padding: 2px; text-align: center;">${(row.needsAdj === true || row.needsAdj === 'TRUE') ? '✔' : ''}</td>
            <td style="border: 1px solid #000; padding: 2px;">${validateInput(row.overtime) ? escapeHtml(row.overtime || '') : ''}</td>
            <td style="border: 1px solid #000; padding: 2px;">${validateInput(row.otCorp) ? escapeHtml(row.otCorp || '') : ''}</td>
            <td style="border: 1px solid #000; padding: 2px; text-align: center;">${(row.subAnother === true || row.subAnother === 'TRUE') ? '✔' : ''}</td>
            <td style="border: 1px solid #000; padding: 2px;">${validateInput(row.incident) ? escapeHtml(row.incident || '') : ''}</td>
            <td style="border: 1px solid #000; padding: 2px;">${validateInput(row.notes) ? escapeHtml(row.notes || '') : ''}</td>
          </tr>
        `;
      }
    });
  });

  staffTableHTML += '</tbody></table>';

  // Doctor 테이블 HTML 생성
  const doctorPresentCount = doctorData.filter(d => d.present === true || d.present === 'TRUE').length;
  let doctorTableHTML = `
    <table style="width: 100%; border-collapse: collapse; font-size: 7pt; font-family: Arial; margin-top: 20px;">
      <thead>
        <tr style="background-color: #e3f2fd; font-weight: bold;">
          <th style="border: 1px solid #000; padding: 2px; text-align: center;">Name</th>
          <th style="border: 1px solid #000; padding: 2px; text-align: center;">Present</th>
          <th style="border: 1px solid #000; padding: 2px; text-align: center;">Check In</th>
          <th style="border: 1px solid #000; padding: 2px; text-align: center;">Lunch Out</th>
          <th style="border: 1px solid #000; padding: 2px; text-align: center;">Lunch In</th>
          <th style="border: 1px solid #000; padding: 2px; text-align: center;">Check Out</th>
        </tr>
        <tr style="background-color: #fce4ec; font-weight: bold;">
          <td colspan="2" style="border: 1px solid #000; padding: 2px;"></td>
          <td style="border: 1px solid #000; padding: 2px; text-align: center; font-size: 8pt;">Doctor</td>
          <td style="border: 1px solid #000; padding: 2px; text-align: center; font-size: 8pt;">${doctorPresentCount}</td>
          <td colspan="2" style="border: 1px solid #000; padding: 2px;"></td>
        </tr>
      </thead>
      <tbody>
  `;

  doctorData.forEach(row => {
    // 🔒 보안: 입력 검증 및 이스케이프
    const safeName = validateInput(row.name) ? escapeHtml(row.name) : '';
    if (safeName && safeName.trim()) {
      const safeCheckIn = validateInput(row.checkIn) ? formatTimeToAMPM(row.checkIn || '') : '';
      const safeLunchOut = validateInput(row.lunchOut) ? formatTimeToAMPM(row.lunchOut || '') : '';
      const safeLunchIn = validateInput(row.lunchIn) ? formatTimeToAMPM(row.lunchIn || '') : '';
      const safeCheckOut = validateInput(row.checkOut) ? formatTimeToAMPM(row.checkOut || '') : '';
      
      doctorTableHTML += `
        <tr>
          <td style="border: 1px solid #000; padding: 2px;">${safeName}</td>
          <td style="border: 1px solid #000; padding: 2px; text-align: center;">${(row.present === true || row.present === 'TRUE') ? '✔' : ''}</td>
          <td style="border: 1px solid #000; padding: 2px;">${safeCheckIn}</td>
          <td style="border: 1px solid #000; padding: 2px;">${safeLunchOut}</td>
          <td style="border: 1px solid #000; padding: 2px;">${safeLunchIn}</td>
          <td style="border: 1px solid #000; padding: 2px;">${safeCheckOut}</td>
        </tr>
      `;
    }
  });

  doctorTableHTML += '</tbody></table>';

  // Daily Recap 계산
  let totalStaffPresent = 0;
  Object.keys(staffByPosition).forEach(pos => {
    totalStaffPresent += staffByPosition[pos].filter(r => r.present === true || r.present === 'TRUE').length;
  });
  const totalPresent = totalStaffPresent + doctorPresentCount;

  // 캘리포니아 시간
  const californiaTime = new Date().toLocaleString('en-US', {
    timeZone: 'America/Los_Angeles',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: true
  });

  return `
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <meta http-equiv="Content-Security-Policy" content="default-src 'self'; script-src 'none'; object-src 'none'; base-uri 'self'; style-src 'unsafe-inline';">
      <title>${office}_${date}_Attendance_Tract</title>
      <style>
        body {
          font-family: Arial, sans-serif;
          margin: 0;
          padding: 20px;
        }
        h1 {
          text-align: center;
          font-size: 14pt;
          font-weight: bold;
          margin-bottom: 8px;
        }
        .header-info {
          text-align: center;
          font-size: 8pt;
          margin-bottom: 8px;
        }
        .daily-recap {
          text-align: right;
          font-size: 10pt;
          font-weight: bold;
          margin-top: 4px;
          margin-bottom: 8px;
        }
        .footer {
          text-align: center;
          font-size: 8pt;
          font-style: italic;
          margin-top: 8px;
        }
      </style>
    </head>
    <body>
      <h1>${escapeHtml(office)} Attendance Tract</h1>
      <div class="header-info">
        Date: ${escapeHtml(date)} &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        Filled Out By: ${escapeHtml(filledBy)} &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        Management that Checked Times Today on Time Clock: ${escapeHtml(checkedBy)}
      </div>
      ${staffTableHTML}
      ${doctorTableHTML}
      <div class="daily-recap">Daily Recap: ${totalPresent}</div>
      <div class="footer">Generated on: ${californiaTime}</div>
    </body>
    </html>
  `;
}


// 🔒 보안: 시간 형식 검증 및 포맷팅
function formatTimeToAMPM(timeStr: string): string {
  if (!timeStr || typeof timeStr !== 'string' || timeStr.trim() === '') return '';
  // 🔒 보안: 입력 검증 (길이 제한)
  if (timeStr.length > 10) return '';
  const timeMatch = timeStr.match(/^(\d{1,2}):(\d{2})$/);
  if (!timeMatch) return escapeHtml(timeStr); // 검증 실패 시 이스케이프하여 반환
  let hour = parseInt(timeMatch[1], 10);
  const minute = timeMatch[2];
  // 🔒 보안: 시간 범위 검증
  if (isNaN(hour) || hour < 0 || hour > 23) return escapeHtml(timeStr);
  if (!/^\d{2}$/.test(minute) || parseInt(minute, 10) < 0 || parseInt(minute, 10) > 59) return escapeHtml(timeStr);
  const period = hour >= 12 ? 'PM' : 'AM';
  if (hour === 0) hour = 12;
  else if (hour > 12) hour = hour - 12;
  return `${hour}:${minute} ${period}`;
}

