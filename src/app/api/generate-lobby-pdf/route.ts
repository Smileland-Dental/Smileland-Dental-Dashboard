import { NextRequest, NextResponse } from 'next/server';
import puppeteer from 'puppeteer';

export async function POST(request: NextRequest) {
  try {
    console.log('Lobby PDF generation started');
    const { inspectionDate, selectedOffice, lobbyData } = await request.json();
    console.log('Request data:', { inspectionDate, selectedOffice, lobbyDataKeys: Object.keys(lobbyData || {}) });

    if (!inspectionDate || !selectedOffice || !lobbyData) {
      console.error('Missing required data:', { inspectionDate, selectedOffice, hasLobbyData: !!lobbyData });
      return NextResponse.json({ error: 'Missing required data' }, { status: 400 });
    }

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
              <th rowspan="2" style="background: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 40px; font-size: 6px;">Time</th>
              <th rowspan="2" style="background: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 30px; font-size: 6px;">Check</th>
              <th colspan="5" style="background: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; text-align: center; font-size: 6px;">Hourly Cleaning</th>
              <th colspan="6" style="background: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; text-align: center; font-size: 6px;">End of Day Cleaning</th>
              <th rowspan="2" style="background: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 35px; font-size: 6px;">Checked Time</th>
            </tr>
            <tr>
              <th style="background: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 32px; font-size: 5px;">Ipads/<br>Games<br>Working</th>
              <th style="background: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 32px; font-size: 5px;">Wipe<br>Ipads/<br>Games</th>
              <th style="background: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 32px; font-size: 5px;">Pick Up<br>Litter/<br>Sweep</th>
              <th style="background: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 34px; font-size: 5px;">Entrance<br>Area</th>
              <th style="background: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 32px; font-size: 5px;">Pass<br>Out<br>Water</th>
              <th style="background: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 32px; font-size: 5px;">Sweep/<br>Vacuum</th>
              <th style="background: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 32px; font-size: 5px;">Wipe<br>Ipads/<br>Games</th>
              <th style="background: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 32px; font-size: 5px;">Take<br>Out<br>Trash</th>
              <th style="background: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 32px; font-size: 5px;">Wipe<br>Desk<br>Tops</th>
              <th style="background: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 32px; font-size: 5px;">Wipe<br>Seats</th>
              <th style="background: #f3f3f3; border: 1px solid #444; padding: 1px; font-weight: bold; width: 36px; font-size: 5px;">Wipe<br>Windows/<br>Door<br>Handles</th>
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
        tableHTML += `<td style="border: 1px solid #444; padding: 2px 2px; text-align: center; vertical-align: center; height: ${rowHeight}; font-size: 6px;">${header}</td>`;
        
        // Check 컬럼
        const checkValue = lobbyData[`Row${rowIndex + 1}_Check`] || '';
        tableHTML += `<td style="border: 1px solid #444; padding: 2px 2px; text-align: center; vertical-align: center; height: ${rowHeight}; font-size: 6px;">${checkValue}</td>`;
        
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
        const checkedTime = lobbyData[`Row${rowIndex + 1}_CheckedTime`] || '';
        tableHTML += `<td style="border: 1px solid #444; padding: 2px 2px; text-align: center; vertical-align: center; height: ${rowHeight}; font-size: 6px;">${checkedTime}</td>`;
        
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
      <html>
        <head>
          <meta charset="UTF-8">
          <title>Lobby Inspection Log</title>
          <style>
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
            <div class="main-title">Smileland Dental Lobby Inspection Log 🛋️</div>
            <div class="sub-title">${selectedOffice} (${formattedDate})</div>
            <hr>
          </div>
          
          ${generateTableHTML()}
          
          <div class="footer">
            Report generated on ${new Date().toLocaleString('en-US', { timeZone: 'America/Los_Angeles' })} PDT
          </div>
        </body>
      </html>
    `;

    // Puppeteer로 PDF 생성
    console.log('Launching Puppeteer browser...');
    let browser;
    try {
      browser = await puppeteer.launch({
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
    });

    console.log('Creating new page...');
    const page = await browser.newPage();
    console.log('Setting page content...');
    await page.setContent(htmlContent, { waitUntil: 'networkidle0' });
    
    // 페이지 로딩 대기
    console.log('Waiting for page to load...');
    await new Promise(resolve => setTimeout(resolve, 1000));

    console.log('Generating PDF...');
    const pdfBuffer = await page.pdf({
      format: 'A4',
      landscape: true,
      printBackground: true,
      margin: {
        top: '15mm',
        bottom: '15mm',
        left: '30mm',
        right: '30mm'
      },
      preferCSSPageSize: false,
      displayHeaderFooter: false
    });

    // 파일명 생성
    const filename = `5) ${inspectionDate}_${selectedOffice}_Lobby Inspection Log.pdf`;
    console.log('PDF generated successfully, filename:', filename);
    
      // 브라우저 안전하게 종료
      try {
        await browser.close();
        console.log('Browser closed successfully');
      } catch (closeError) {
        console.warn('Browser close warning:', closeError);
      }
      
      return new NextResponse(pdfBuffer, {
        headers: {
          'Content-Type': 'application/pdf',
          'Content-Disposition': `attachment; filename="${filename}"`
        }
      });
    } catch (innerError) {
      // 브라우저가 열려있다면 안전하게 종료
      try {
        if (browser) {
          await browser.close();
        }
      } catch (closeError) {
        console.warn('Browser close error in inner catch:', closeError);
      }
      throw innerError;
    }

  } catch (error) {
    console.error('PDF generation error:', error);
    return NextResponse.json(
      { success: false, error: 'Failed to generate PDF: ' + (error as Error).message },
      { status: 500 }
    );
  }
}
