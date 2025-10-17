import { NextRequest, NextResponse } from 'next/server';
import puppeteer from 'puppeteer';
import { 
  escapeHtml, 
  sanitizeString, 
  safeLog, 
  logError, 
  sanitizePdfFilename, 
  getSecurePuppeteerOptions, 
  validatePdfGeneration 
} from '@/lib/security-server';
import { Buffer } from 'buffer';

export async function POST(request: NextRequest) {
  try {
    const { dutyDate, patientRows } = await request.json();

    if (!dutyDate || !patientRows) {
      return NextResponse.json({ success: false, error: 'Duty date and patient data are required' });
    }

    // 데이터 안전성 검증
    const safeDutyDate = sanitizeString(dutyDate, 50);
    
    // PDF 생성 데이터 크기 검증
    const dataSize = JSON.stringify({ dutyDate, patientRows }).length;
    if (!validatePdfGeneration(dataSize)) {
      throw new Error('Data size too large for PDF generation');
    }

    safeLog('Add-on treatment PDF generation started', { dutyDate: safeDutyDate });

    // 시간을 12시간 형식으로 변환하는 함수
    const convertTo12Hour = (timeStr: string): string => {
      if (!timeStr || timeStr === '-') return '-';
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
        <title>${escapeHtml(safeDutyDate)}_Bernard_Add On Treatment</title>
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
          <h1>Smileland Dental Add-On Treatment</h1>
        </div>
        
        <div class="info-section">
          <div><strong>Date:</strong> ${escapeHtml(safeDutyDate)}</div>
          <div><strong>Location:</strong> Bernard</div>
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
              const safePatientName = sanitizeString(row.patientName, 100);
              const safeDob = sanitizeString(row.dob, 20);
              const safeTime = sanitizeString(row.time, 20);
              
              return `
              <tr>
                <td>${index + 1}</td>
                <td>${escapeHtml(safePatientName) || '-'}</td>
                <td>${escapeHtml(safeDob) || '-'}</td>
                <td>${escapeHtml(convertTo12Hour(safeTime))}</td>
              </tr>
            `;
            }).join('')}
          </tbody>
        </table>
        
        <div class="footer">
          <p>Smileland Dental</p>
          <p>Report generated on ${new Date().toLocaleDateString('en-US', {
            year: 'numeric',
            month: 'long',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit',
            hour12: true
          })}</p>
        </div>
      </body>
      </html>
    `;

    // Puppeteer로 PDF 생성 (보안 옵션 적용)
    const browser = await puppeteer.launch(getSecurePuppeteerOptions());
    
    const page = await browser.newPage();
    await page.setContent(html);
    
    // 페이지 로딩 대기
    await new Promise(resolve => setTimeout(resolve, 1000));
    
    const pdf = await page.pdf({
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
    });
    
    await browser.close();
    
    // 파일명 생성 (보안 검증)
    const filename = sanitizePdfFilename(`4) ${safeDutyDate}_Bernard_Add On Treatment.pdf`);
    
    return new NextResponse(Buffer.from(pdf), {
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${filename}"`,
        'X-Content-Type-Options': 'nosniff',
        'X-Frame-Options': 'DENY',
        'Cache-Control': 'no-cache, no-store, must-revalidate',
        'Pragma': 'no-cache',
        'Expires': '0'
      }
    });

  } catch (error) {
    logError(error, 'Add-on treatment PDF generation');
    return NextResponse.json({ 
      success: false, 
      error: 'Failed to generate add-on treatment PDF' 
    });
  }
}
