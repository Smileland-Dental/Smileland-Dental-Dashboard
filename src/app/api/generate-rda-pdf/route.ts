import { NextRequest, NextResponse } from 'next/server';
import puppeteer from 'puppeteer';

export async function POST(request: NextRequest) {
  try {
    const { office, rdaName, date, treatmentData } = await request.json();

    // 데이터 검증
    if (!office || !rdaName || !date || !treatmentData) {
      return NextResponse.json({ error: 'Missing required data' }, { status: 400 });
    }

    // 입력값 검증 및 정리
    const safeOffice = String(office).trim().replace(/[^A-Za-z0-9]/g, '').substring(0, 10);
    const safeRdaName = String(rdaName).trim().replace(/[^A-Za-z0-9\s]/g, '').substring(0, 50);
    const safeDate = String(date).trim().replace(/[^0-9-]/g, '');
    
    // 날짜 형식 검증 (YYYY-MM-DD)
    if (!/^\d{4}-\d{2}-\d{2}$/.test(safeDate)) {
      return NextResponse.json({ error: 'Invalid date format' }, { status: 400 });
    }
    
    // 데이터 타입 검증
    if (!Array.isArray(treatmentData)) {
      return NextResponse.json({ error: 'Invalid treatment data format' }, { status: 400 });
    }
    
    // 배열 크기 제한 (DoS 공격 방지)
    if (treatmentData.length > 1000) {
      return NextResponse.json({ error: 'Treatment data array too large' }, { status: 400 });
    }

    // HTML 이스케이핑 함수 (길이 제한 포함)
    const escapeHtml = (text: string, maxLength: number = 1000): string => {
      if (typeof text !== 'string') return '';
      // 길이 제한 적용
      const limitedText = text.substring(0, maxLength);
      return limitedText
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
    };

    // 24시간제를 12시간제로 변환하는 함수
    const formatTime12Hour = (time24: string): string => {
      if (!time24 || typeof time24 !== 'string' || time24.trim() === '') return '';
      
      // 입력 길이 제한
      const limitedTime = time24.trim().substring(0, 10);
      
      // 시간 형식이 HH:MM 또는 H:MM인지 확인
      const timeMatch = limitedTime.match(/^(\d{1,2}):(\d{2})$/);
      if (!timeMatch) return ''; // 형식이 맞지 않으면 빈 문자열 반환 (안전)
      
      const hours = parseInt(timeMatch[1], 10);
      const minutes = timeMatch[2];
      
      // 시간 범위 검증
      if (hours < 0 || hours > 23 || parseInt(minutes, 10) < 0 || parseInt(minutes, 10) > 59) {
        return '';
      }
      
      if (hours === 0) {
        return `12:${minutes} AM`;
      } else if (hours < 12) {
        return `${hours}:${minutes} AM`;
      } else if (hours === 12) {
        return `12:${minutes} PM`;
      } else {
        return `${hours - 12}:${minutes} PM`;
      }
    };

    // 날짜 포맷팅
    const formatDate = (dateString: string) => {
      if (!dateString || typeof dateString !== 'string') return '';
      
      // 날짜 형식 검증
      if (!/^\d{4}-\d{2}-\d{2}$/.test(dateString)) return '';
      
      const [year, month, day] = dateString.split('-').map(Number);
      
      // 날짜 범위 검증
      if (isNaN(year) || isNaN(month) || isNaN(day) || 
          year < 1900 || year > 2100 || 
          month < 1 || month > 12 || 
          day < 1 || day > 31) {
        return '';
      }
      
      const date = new Date(year, month - 1, day);
      
      // 유효한 날짜인지 확인
      if (isNaN(date.getTime())) return '';
      
      return date.toLocaleDateString('en-US', {
        timeZone: 'America/Los_Angeles',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit'
      });
    };

    const formattedDate = formatDate(safeDate);

    // 치료 데이터 필터링 (빈 행 제외)
    const filteredData = treatmentData.filter(row => 
      row.patientName || row.startTime || row.roomNumber || 
      (row.services && row.services.length > 0) || row.explanation ||
      (row.sealantDetails && row.sealantDetails.length > 0)
    );


    // HTML 테이블 생성
    const generateTableHTML = () => {
      let tableHTML = `
        <table style="width: 100%; border-collapse: collapse; font-size: 8px; margin: 0 auto; border: 1px solid #000;">
          <thead>
            <tr>
              <th style="background: #f8f9fa; color: #000; border: 1px solid #000; padding: 6px; font-weight: bold; width: 5%; font-size: 8px; text-align: center;">#</th>
              <th style="background: #f8f9fa; color: #000; border: 1px solid #000; padding: 6px; font-weight: bold; width: 12%; font-size: 8px; text-align: left;">PT Name</th>
              <th style="background: #f8f9fa; color: #000; border: 1px solid #000; padding: 6px; font-weight: bold; width: 10%; font-size: 8px; text-align: left;">Start Time</th>
              <th style="background: #f8f9fa; color: #000; border: 1px solid #000; padding: 6px; font-weight: bold; width: 10%; font-size: 8px; text-align: left;">Room #</th>
              <th style="background: #f8f9fa; color: #000; border: 1px solid #000; padding: 6px; font-weight: bold; width: 28%; font-size: 8px; text-align: left;">Treatment or Services Performed</th>
              <th style="background: #f8f9fa; color: #000; border: 1px solid #000; padding: 6px; font-weight: bold; width: 35%; font-size: 8px; text-align: left;">Explanation of Treatment/Services or Amount Performed</th>
            </tr>
          </thead>
          <tbody>
      `;

      filteredData.forEach((row, index) => {
        const rowBgColor = index % 2 === 0 ? '#ffffff' : '#f8f9fa';
        tableHTML += `<tr style="min-height: 50px; background-color: ${rowBgColor};">`;
        
        // Row Number
        tableHTML += `<td style="border: 1px solid #000; padding: 6px; vertical-align: top; font-size: 8px; text-align: center; font-weight: bold;">${index + 1}</td>`;
        
        // Patient Name
        tableHTML += `<td style="border: 1px solid #000; padding: 6px; vertical-align: top; font-size: 8px;">${escapeHtml(String(row.patientName || ''), 100)}</td>`;
        
        // Start Time
        tableHTML += `<td style="border: 1px solid #000; padding: 6px; vertical-align: top; font-size: 8px;">${escapeHtml(formatTime12Hour(String(row.startTime || '')), 20)}</td>`;
        
        // Room Number
        tableHTML += `<td style="border: 1px solid #000; padding: 6px; vertical-align: top; font-size: 8px;">${escapeHtml(String(row.roomNumber || ''), 10)}</td>`;
        
        // Treatment or Services Performed
        tableHTML += `<td style="border: 1px solid #000; padding: 6px; vertical-align: top; font-size: 8px;">`;
        
        if (row.services && Array.isArray(row.services) && row.services.length > 0) {
          // services 배열 크기 제한 (DoS 방지)
          const limitedServices = row.services.slice(0, 50);
          tableHTML += `<ul style="margin: 0; padding-left: 15px; font-size: 8px;">`;
          limitedServices.forEach((service: any) => {
            tableHTML += `<li style="margin: 2px 0; font-size: 8px;">${escapeHtml(String(service || ''), 200)}</li>`;
          });
          tableHTML += `</ul>`;
        }
        
        // Sealant details는 별도 테이블에서 처리하므로 여기서는 제거
        
        tableHTML += `</td>`;
        
        // Explanation
        tableHTML += `<td style="border: 1px solid #000; padding: 6px; vertical-align: top; font-size: 8px;">${escapeHtml(String(row.explanation || ''), 500)}</td>`;
        
        tableHTML += `</tr>`;
      });

      tableHTML += `
          </tbody>
        </table>
      `;

      return tableHTML;
    };

    // Sealant Details 테이블 생성
    const generateSealantTableHTML = () => {
      // 모든 sealant details 수집
      const allSealantDetails: Array<{
        patientName: string;
        detail: any;
        rowIndex: number;
      }> = [];

      filteredData.forEach((row, rowIndex) => {
        if (row.sealantDetails && Array.isArray(row.sealantDetails) && row.sealantDetails.length > 0) {
          // sealantDetails 배열 크기 제한 (DoS 방지)
          const limitedDetails = row.sealantDetails.slice(0, 100);
          limitedDetails.forEach((detail: any) => {
            // detail이 객체인지 확인
            if (detail && typeof detail === 'object') {
              allSealantDetails.push({
                patientName: row.patientName || '',
                detail: detail,
                rowIndex: rowIndex
              });
            }
          });
        }
      });

      if (allSealantDetails.length === 0) {
        return '';
      }

      let sealantTableHTML = `
        <div style="margin-top: 30px;">
          <h3 style="font-size: 14px; font-weight: bold; margin-bottom: 15px; color: #000; border-bottom: 1px solid #000; padding-bottom: 5px;">
            Sealant Details
          </h3>
          <table style="width: 100%; border-collapse: collapse; font-size: 8px; margin: 0 auto; border: 1px solid #000;">
            <thead>
              <tr>
                <th style="background: #f8f9fa; color: #000; border: 1px solid #000; padding: 6px; font-weight: bold; width: 4%; font-size: 8px; text-align: center;">#</th>
                <th style="background: #f8f9fa; color: #000; border: 1px solid #000; padding: 6px; font-weight: bold; width: 10%; font-size: 8px; text-align: left;">PT Name</th>
                <th style="background: #f8f9fa; color: #000; border: 1px solid #000; padding: 6px; font-weight: bold; width: 8%; font-size: 8px; text-align: left;">Chart #</th>
                <th style="background: #f8f9fa; color: #000; border: 1px solid #000; padding: 6px; font-weight: bold; width: 8%; font-size: 8px; text-align: left;">DOB</th>
                <th style="background: #f8f9fa; color: #000; border: 1px solid #000; padding: 6px; font-weight: bold; width: 8%; font-size: 8px; text-align: left;">Tooth #</th>
                <th style="background: #f8f9fa; color: #000; border: 1px solid #000; padding: 6px; font-weight: bold; width: 6%; font-size: 8px; text-align: left;">Redo</th>
                <th style="background: #f8f9fa; color: #000; border: 1px solid #000; padding: 6px; font-weight: bold; width: 8%; font-size: 8px; text-align: left;">Acct Type</th>
                <th style="background: #f8f9fa; color: #000; border: 1px solid #000; padding: 6px; font-weight: bold; width: 8%; font-size: 8px; text-align: left;">Payable</th>
                <th style="background: #f8f9fa; color: #000; border: 1px solid #000; padding: 6px; font-weight: bold; width: 8%; font-size: 8px; text-align: left;">DX Dr.</th>
                <th style="background: #f8f9fa; color: #000; border: 1px solid #000; padding: 6px; font-weight: bold; width: 8%; font-size: 8px; text-align: left;">DR $</th>
                <th style="background: #f8f9fa; color: #000; border: 1px solid #000; padding: 6px; font-weight: bold; width: 8%; font-size: 8px; text-align: left;">RDA $</th>
              </tr>
            </thead>
            <tbody>
      `;

      allSealantDetails.forEach((item, index) => {
        const rowBgColor = index % 2 === 0 ? '#ffffff' : '#f8f9fa';
        sealantTableHTML += `
          <tr style="background-color: ${rowBgColor};">
            <td style="border: 1px solid #000; padding: 6px; vertical-align: top; font-size: 8px; text-align: center; font-weight: bold;">${index + 1}</td>
            <td style="border: 1px solid #000; padding: 6px; vertical-align: top; font-size: 8px;">${escapeHtml(String(item.patientName || ''), 100)}</td>
            <td style="border: 1px solid #000; padding: 6px; vertical-align: top; font-size: 8px;">${escapeHtml(String(item.detail.chartNumber || ''), 50)}</td>
            <td style="border: 1px solid #000; padding: 6px; vertical-align: top; font-size: 8px;">${escapeHtml(String(item.detail.dob || ''), 20)}</td>
            <td style="border: 1px solid #000; padding: 6px; vertical-align: top; font-size: 8px;">${escapeHtml(String(item.detail.toothNumber || ''), 10)}</td>
            <td style="border: 1px solid #000; padding: 6px; vertical-align: top; font-size: 8px;">${escapeHtml(String(item.detail.redo || ''), 10)}</td>
            <td style="border: 1px solid #000; padding: 6px; vertical-align: top; font-size: 8px;">${escapeHtml(String(item.detail.acctType || ''), 10)}</td>
            <td style="border: 1px solid #000; padding: 6px; vertical-align: top; font-size: 8px;">${escapeHtml(String(item.detail.payable || ''), 10)}</td>
            <td style="border: 1px solid #000; padding: 6px; vertical-align: top; font-size: 8px;">${escapeHtml(String(item.detail.dxDr || ''), 50)}</td>
            <td style="border: 1px solid #000; padding: 6px; vertical-align: top; font-size: 8px;">${escapeHtml(String(item.detail.drAmount || ''), 20)}</td>
            <td style="border: 1px solid #000; padding: 6px; vertical-align: top; font-size: 8px;">${escapeHtml(String(item.detail.rdaAmount || ''), 20)}</td>
          </tr>
        `;
      });

      sealantTableHTML += `
            </tbody>
          </table>
        </div>
      `;

      return sealantTableHTML;
    };

    // HTML 문서 생성
    const htmlContent = `
      <!DOCTYPE html>
      <html>
        <head>
          <meta charset="UTF-8">
          <title>RDA Treatment Sheet</title>
          <style>
            body {
              font-family: 'Arial', sans-serif;
              margin: 15px 20px;
              font-size: 10px;
              color: #000;
              line-height: 1.4;
            }
            .header {
              text-align: center;
              margin-bottom: 25px;
              border-bottom: 2px solid #000;
              padding-bottom: 15px;
            }
            .main-title {
              font-size: 20px;
              font-weight: bold;
              margin-bottom: 8px;
              color: #000;
              letter-spacing: 0.5px;
            }
            .info-section {
              display: flex;
              justify-content: space-between;
              margin-bottom: 25px;
              font-size: 11px;
              background-color: #f8f9fa;
              padding: 12px;
              border: 1px solid #dee2e6;
            }
            .info-item {
              display: flex;
              align-items: center;
            }
            .info-label {
              font-weight: bold;
              margin-right: 8px;
              color: #495057;
            }
            .info-value {
              color: #000;
            }
            .footer {
              font-size: 9px;
              color: #6c757d;
              margin-top: 25px;
              text-align: center;
              border-top: 1px solid #dee2e6;
              padding-top: 10px;
            }
            table {
              width: 100%;
              border-collapse: collapse;
              margin-top: 10px;
            }
            th, td {
              border: 1px solid #000;
              padding: 8px;
              text-align: left;
              vertical-align: top;
              font-size: 9px;
            }
            th {
              background-color: #f8f9fa;
              color: #000;
              font-weight: bold;
            }
            td {
              min-height: 50px;
            }
            .sealant-detail {
              margin: 4px 0;
              padding: 4px 6px;
              background-color: #f8f9fa;
              border: 1px solid #dee2e6;
              font-size: 7px;
            }
          </style>
        </head>
        <body>
          <div class="header">
            <div class="main-title">RDA/DA Treatment (Sealant)</div>
          </div>

          <div class="info-section">
            <div class="info-item">
              <span class="info-label">Date:</span>
              <span class="info-value">${escapeHtml(formattedDate)}</span>
            </div>
            <div class="info-item">
              <span class="info-label">Office:</span>
              <span class="info-value">${escapeHtml(safeOffice)}</span>
            </div>
            <div class="info-item">
              <span class="info-label">RDA/DA Name:</span>
              <span class="info-value">${escapeHtml(safeRdaName)}</span>
            </div>
          </div>
          
          ${generateTableHTML()}
          
          ${generateSealantTableHTML()}
          
          <div class="footer">
            Report generated on ${new Date().toLocaleString('en-US', { timeZone: 'America/Los_Angeles' })} PDT
          </div>
        </body>
      </html>
    `;

    // Puppeteer로 PDF 생성
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
          '--disable-gpu',
          '--disable-features=VizDisplayCompositor'
          // --disable-web-security 제거 (보안상 위험)
        ],
        timeout: 30000 // 30초 타임아웃
      });

      const page = await browser.newPage();
      
      // 페이지 설정
      await page.setViewport({ width: 1200, height: 800 });
      await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36');
      
      await page.setContent(htmlContent, { 
        waitUntil: 'networkidle0',
        timeout: 30000
      });
      
      // 페이지 로딩 대기
      await new Promise(resolve => setTimeout(resolve, 2000));

      const pdfBuffer = await page.pdf({
        format: 'A4',
        printBackground: true,
        margin: {
          top: '15mm',
          bottom: '15mm',
          left: '15mm',
          right: '15mm'
        },
        preferCSSPageSize: false,
        displayHeaderFooter: false,
        timeout: 30000
      });

      // 파일명 생성 (8) 포함) - 이미 검증된 값 사용
      const filename = `8) ${safeDate}_${safeOffice}_${safeRdaName}_RDA Treatment Sheet.pdf`;
      
      // 브라우저 안전하게 종료
      try {
        await browser.close();
      } catch (closeError) {
        // 브라우저 종료 오류는 무시
      }
      
      // 파일명 안전하게 인코딩 (RFC 5987)
      const encodedFilename = encodeURIComponent(filename);
      
      // Uint8Array를 Buffer로 변환 (NextResponse 호환성)
      const pdfBufferAsBuffer = Buffer.from(pdfBuffer);
      
      return new NextResponse(pdfBufferAsBuffer, {
        headers: {
          'Content-Type': 'application/pdf',
          'Content-Disposition': `attachment; filename="${filename}"; filename*=UTF-8''${encodedFilename}`
        }
      });
    } catch (innerError) {
      // 브라우저가 열려있다면 안전하게 종료
      try {
        if (browser) {
          await browser.close();
        }
      } catch (closeError) {
        // 브라우저 종료 오류는 무시
      }
      throw innerError;
    }

  } catch (error) {
    // 프로덕션에서는 상세한 에러 정보를 노출하지 않음
    
    // 에러 타입별 처리
    let errorMessage = 'Failed to generate PDF';
    let statusCode = 500;
    
    if (error instanceof Error) {
      if (error.message.includes('timeout')) {
        errorMessage = 'PDF generation timed out. Please try again.';
        statusCode = 408;
      } else if (error.message.includes('browser')) {
        errorMessage = 'Browser initialization failed. Please try again.';
        statusCode = 503;
      }
      // 프로덕션에서는 상세한 에러 메시지 제거
    }
    
    return NextResponse.json(
      { 
        success: false, 
        error: errorMessage,
        timestamp: new Date().toISOString()
      },
      { status: statusCode }
    );
  }
}
