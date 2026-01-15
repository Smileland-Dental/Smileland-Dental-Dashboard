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

    const { inspectionDate, selectedOffice, lobbyData } = await request.json();

    if (!inspectionDate || !selectedOffice || !lobbyData) {
      return NextResponse.json({ 
        success: false, 
        error: 'Inspection date, office, and data are required' 
      }, {
        status: 400
      });
    }
    
    // 데이터 sanitize
    const safeInspectionDate = sanitizeString(inspectionDate, 20);
    const safeSelectedOffice = sanitizeString(selectedOffice, 100);

    // 날짜를 요일로 변환
    const dateObj = new Date(inspectionDate + 'T00:00:00');
    const laDate = new Date(dateObj.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }));
    const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const dayName = days[laDate.getDay()];

    // 날짜 포맷팅 (MM/DD/YYYY)
    const month = String(laDate.getMonth() + 1).padStart(2, '0');
    const day = String(laDate.getDate()).padStart(2, '0');
    const year = laDate.getFullYear();
    const formattedDate = `${month}/${day}/${year}`;

    // 컬럼 정의
    const COLUMN_NAMES = [
      'Time', 'Check', 'Ipads/Games Working', 'Wipe Ipads/Games', 'Pick Up Litter/Sweep',
      'Entrance Area', 'Pass Out Water', 'Sweep/Vacuum', 'Wipe Ipads/Games',
      'Take Out Trash', 'Wipe Desk Tops', 'Wipe Seats', 'Wipe Windows/Door Handles', 'Checked Time'
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

    // HTML 테이블 생성
    const generateTableHTML = () => {
      let tableHTML = `
        <table style="width: 100%; border-collapse: collapse; font-size: 6px; margin: 0 auto;">
          <thead>
            <tr>
              <th rowspan="2" style="backgroundColor: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 40px; font-size: 6px;">Time</th>
              <th rowspan="2" style="backgroundColor: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 30px; font-size: 6px;">Check</th>
              <th colspan="5" style="backgroundColor: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; text-align: center; font-size: 6px;">Hourly Cleaning</th>
              <th colspan="6" style="backgroundColor: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; text-align: center; font-size: 6px;">End of Day Cleaning</th>
              <th rowspan="2" style="backgroundColor: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 35px; font-size: 6px;">Checked Time</th>
            </tr>
            <tr>
              <th style="backgroundColor: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 32px; font-size: 5px;">Ipads/<br>Games<br>Working</th>
              <th style="backgroundColor: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 32px; font-size: 5px;">Wipe<br>Ipads/<br>Games</th>
              <th style="backgroundColor: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 32px; font-size: 5px;">Pick Up<br>Litter/<br>Sweep</th>
              <th style="backgroundColor: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 34px; font-size: 5px;">Entrance<br>Area</th>
              <th style="backgroundColor: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 32px; font-size: 5px;">Pass<br>Out<br>Water</th>
              <th style="backgroundColor: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 32px; font-size: 5px;">Sweep/<br>Vacuum</th>
              <th style="backgroundColor: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 32px; font-size: 5px;">Wipe<br>Ipads/<br>Games</th>
              <th style="backgroundColor: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 32px; font-size: 5px;">Take<br>Out<br>Trash</th>
              <th style="backgroundColor: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 32px; font-size: 5px;">Wipe<br>Desk<br>Tops</th>
              <th style="backgroundColor: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 32px; font-size: 5px;">Wipe<br>Seats</th>
              <th style="backgroundColor: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 36px; font-size: 5px;">Wipe<br>Windows/<br>Door<br>Handles</th>
            </tr>
          </thead>
          <tbody>
      `;

      ROW_HEADERS.forEach((header, rowIndex) => {
        const isManagerInspection = header.includes('Manager Inspection');
        const isSweepMop = header.includes('Sweep/Mop');
        
        // Manager Inspection은 기본 높이, 나머지는 더 높게
        const rowHeight = isManagerInspection ? '18px' : '22px';
        tableHTML += `<tr style="min-height: ${rowHeight}; height: ${rowHeight};">`;
        
        // Time 컬럼
        tableHTML += `<td style="border: 1px solid #444; padding: 2px 2px; text-align: center; vertical-align: center; height: ${rowHeight}; font-size: 6px;">${escapeHtml(header)}</td>`;
        
        // Check 컬럼
        const checkValue = sanitizeString(lobbyData[`Row${rowIndex + 1}_Check`] || '', 50);
        tableHTML += `<td style="border: 1px solid #444; padding: 2px 2px; text-align: center; vertical-align: center; height: ${rowHeight}; font-size: 6px;">${escapeHtml(checkValue)}</td>`;
        
        // 나머지 컬럼들
        COLUMN_NAMES.slice(2, -1).forEach((columnName, colIndex) => {
          if (isSweepMop) {
            tableHTML += `<td style="border: 1px solid #444; padding: 2px 2px; text-align: center; vertical-align: center; height: ${rowHeight}; font-size: 6px;"></td>`;
          } else {
            const isChecked = lobbyData[`Row${rowIndex + 1}_Col${colIndex + 3}`] === true;
            tableHTML += `<td style="border: 1px solid #444; padding: 2px 2px; text-align: center; vertical-align: center; height: ${rowHeight}; font-size: 6px;">${isChecked ? '✔' : ''}</td>`;
          }
        });
        
        // Checked Time 컬럼
        const checkedTime = sanitizeString(lobbyData[`Row${rowIndex + 1}_CheckedTime`] || '', 50);
        tableHTML += `<td style="border: 1px solid #444; padding: 2px 2px; text-align: center; vertical-align: center; height: ${rowHeight}; font-size: 6px;">${escapeHtml(checkedTime)}</td>`;
        
        tableHTML += `</tr>`;
      });

      tableHTML += `
          </tbody>
        </table>
      `;

      return tableHTML;
    };

    // HTML 문서 생성
    const htmlContent = `
      <!DOCTYPE html>
      <html lang="en">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>${escapeHtml(safeInspectionDate)}_${escapeHtml(safeSelectedOffice)}_Lobby Inspection Log</title>
          <style>
            @media print {
              @page {
                size: A4 landscape;
                margin: 0.3in;
              }
              body { 
                margin: 0; 
                font-size: 6px; 
                line-height: 1.1;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
              }
            }
            
            body {
              font-family: Arial, sans-serif;
              margin: 10px 20px;
              font-size: 6px;
            }
            .header {
              text-align: center;
              margin-bottom: 15px;
            }
            .main-title {
              font-size: 14px;
              font-weight: bold;
              margin-bottom: 3px;
            }
            .sub-title {
              font-size: 10px;
              margin-bottom: 8px;
            }
            .footer {
              font-size: 6px;
              color: #888;
              margin-top: 15px;
              text-align: center;
            }
            table {
              width: 100%;
              border-collapse: collapse;
            }
            th, td {
              border: 1px solid #444;
              padding: 2px 2px;
              text-align: center;
              vertical-align: center;
              font-size: 6px;
            }
            td {
              min-height: 20px;
              height: 20px;
            }
            th {
              background-color: #f3f3f3;
              font-weight: bold;
            }
          </style>
        </head>
        <body>
          <div class="header">
            <div class="main-title">Lobby Inspection Log</div>
            <div class="sub-title">${escapeHtml(safeSelectedOffice)} (${escapeHtml(formattedDate)})</div>
            <hr>
          </div>
          
          ${generateTableHTML()}
          
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
      await page.setContent(htmlContent, { waitUntil: 'networkidle0' });
      
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
      
      // PDF 반환 (Buffer로 변환하여 타입 에러 해결)
      return new NextResponse(Buffer.from(pdf), {
        headers: {
          'Content-Type': 'application/pdf',
          'Content-Disposition': `attachment; filename="${safeInspectionDate}_${safeSelectedOffice}_Lobby Inspection Log.pdf"`,
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
      throw new Error(`PDF generation failed: ${pdfError.message}`);
    }

  } catch (error) {
    return NextResponse.json(
      { 
        success: false, 
        error: 'Failed to generate lobby inspection report' 
      },
      {
        status: 500,
        headers: {
          'X-Content-Type-Options': 'nosniff',
          'X-Frame-Options': 'DENY',
          'X-XSS-Protection': '1; mode=block'
        }
      }
    );
  }
}
