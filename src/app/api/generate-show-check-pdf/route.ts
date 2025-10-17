import { NextRequest, NextResponse } from 'next/server';
import puppeteer from 'puppeteer';
import { Buffer } from 'buffer';
import { escapeHtml, sanitizeString, safeLog, logError, sanitizePdfFilename, getSecurePuppeteerOptions, validatePdfGeneration } from '@/lib/security-server';

export async function POST(request: NextRequest) {
  try {
    const { showCheckData } = await request.json();

    if (!showCheckData) {
      return NextResponse.json({ success: false, error: 'Show check data is required' });
    }

    const { startDate, endDate, selectedOffice, appointments, generatedBy, timestamp } = showCheckData;
    
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
    
    // PDF 생성 데이터 크기 검증
    const dataSize = JSON.stringify(showCheckData).length;
    if (!validatePdfGeneration(dataSize)) {
      throw new Error('Data size too large for PDF generation');
    }
    
    // 데이터 안전성 검증 및 이스케이프
    const safeStartDate = sanitizeString(startDate, 20);
    const safeEndDate = sanitizeString(endDate, 20);
    const safeSelectedOffice = sanitizeString(selectedOffice, 100);
    const safeGeneratedBy = sanitizeString(generatedBy, 100);

    // 통계 계산
    const showCount = appointments.filter((apt: any) => apt.showStatus === 'show').length;
    const noShowCount = appointments.filter((apt: any) => apt.showStatus === 'no-show').length;
    const pendingCount = appointments.filter((apt: any) => apt.showStatus === 'pending').length;
    const showRate = appointments.length > 0 ? ((showCount / (showCount + noShowCount)) * 100).toFixed(1) : 0;

    // HTML 템플릿 생성
    const html = `
      <!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>${safeStartDate === safeEndDate ? escapeHtml(safeStartDate) : `${escapeHtml(safeStartDate)}_to_${escapeHtml(safeEndDate)}`}_${escapeHtml(safeSelectedOffice)}_Show_Check_Report</title>
        <style>
          body { 
            font-family: Arial, sans-serif; 
            margin: 20px; 
            background: white;
            color: #000;
            line-height: 1.4;
            font-size: 12px;
          }
          .header { 
            text-align: center; 
            margin-bottom: 20px; 
            border-bottom: 2px solid #000;
            padding-bottom: 10px;
          }
          .header h1 { 
            color: #000; 
            font-size: 20px; 
            margin: 0;
            font-weight: bold;
          }
          .header .subtitle {
            color: #666;
            font-size: 12px;
            margin: 5px 0 0 0;
          }
          .info-section { 
            display: flex; 
            justify-content: space-between; 
            margin-bottom: 20px;
            padding: 10px;
            border: 1px solid #ccc;
          }
          .info-column {
            flex: 1;
          }
          .info-item { 
            margin-bottom: 5px; 
            display: flex;
          }
          .info-label { 
            font-weight: bold; 
            min-width: 80px;
            margin-right: 10px;
          }
          .info-value { 
            color: #000;
          }
          .section-title {
            color: #000; 
            font-size: 16px;
            font-weight: bold;
            margin: 20px 0 10px 0;
            border-bottom: 1px solid #000;
            padding-bottom: 5px;
          }
          .stats {
            display: flex;
            justify-content: center;
            gap: 30px;
            margin-bottom: 15px;
            padding: 8px;
            border: 1px solid #ccc;
            background: #f9f9f9;
          }
          .stat-item {
            text-align: center;
          }
          .stat-value {
            font-size: 16px;
            font-weight: bold;
            margin-bottom: 2px;
          }
          .stat-label {
            font-weight: bold;
            font-size: 11px;
          }
          table { 
            width: 100%; 
            border-collapse: collapse; 
            margin: 10px 0;
          }
          th, td { 
            border: 1px solid #000; 
            padding: 6px 4px; 
            text-align: left;
            vertical-align: top;
            font-size: 10px;
          }
          th { 
            background: #f0f0f0; 
            color: #000; 
            font-weight: bold;
            text-align: center;
          }
          .number-cell {
            text-align: center;
            font-weight: bold;
          }
          .show { color: #333; font-weight: bold; }
          .no-show { color: #333; font-weight: bold; }
          .pending { color: #333; font-weight: bold; }
          .footer { 
            text-align: center; 
            margin-top: 20px; 
            padding-top: 10px;
            border-top: 1px solid #ccc;
            font-size: 10px;
          }
          .footer .clinic-name {
            font-size: 12px;
            font-weight: bold;
            margin-bottom: 3px;
          }
          .footer .date {
            font-size: 10px;
            color: #666;
          }
          @media print {
            @page {
              size: A4;
              margin: 0.3in;
            }
            body { 
              margin: 0; 
              font-size: 9px; 
              line-height: 1.1;
              -webkit-print-color-adjust: exact;
              print-color-adjust: exact;
            }
            .header { 
              margin: 0 0 10px 0; 
              padding: 8px 0; 
            }
            .header h1 { 
              font-size: 16px; 
              margin: 0;
            }
            .header .subtitle {
              font-size: 10px;
              margin: 0;
            }
            .info-section { 
              display: flex;
              justify-content: space-between;
              gap: 5px; 
              margin-bottom: 8px;
            }
            .info-item {
              margin-bottom: 3px;
              display: flex;
            }
            .info-label {
              font-size: 8px;
              min-width: 50px;
              margin-right: 5px;
            }
            .info-value {
              font-size: 8px;
            }
            .section-title {
              font-size: 11px;
              margin: 8px 0 5px 0;
              padding-bottom: 3px;
            }
            .stats {
              padding: 4px 8px;
              margin-bottom: 5px;
              font-size: 8px;
            }
            table {
              margin: 3px 0;
              font-size: 7px;
              width: 100%;
              table-layout: fixed;
            }
            th, td {
              padding: 2px;
              border: 0.5px solid #999;
              word-wrap: break-word;
              overflow: hidden;
              text-overflow: ellipsis;
            }
            th {
              font-size: 7px;
              padding: 4px 2px;
              height: 16px;
            }
            tr {
              height: 18px;
            }
            .number-cell {
              font-size: 7px;
              text-align: center;
            }
            .footer {
              margin-top: 10px;
              padding-top: 5px;
              font-size: 7px;
            }
            .footer .clinic-name {
              font-size: 8px;
            }
            .footer .date {
              font-size: 7px;
            }
          }
        </style>
      </head>
      <body>
        <div class="header">
          <h1>Show/No Show Check Report</h1>
        </div>
        
        <div class="info-section">
          <div><strong>Appointment Date Range:</strong> ${safeStartDate === safeEndDate ? escapeHtml(safeStartDate) : `${escapeHtml(safeStartDate)} to ${escapeHtml(safeEndDate)}`}</div>
          <div><strong>Appointment Office:</strong> ${escapeHtml(safeSelectedOffice)}</div>
          <div><strong>Checked by:</strong> ${escapeHtml(safeGeneratedBy)}</div>
          <div><strong>Total Appointments:</strong> ${appointments.length}</div>
        </div>
        
        <div class="stats">
          <div class="stat-item">
            <div class="stat-value">${showCount}</div>
            <div class="stat-label">Show</div>
          </div>
          <div class="stat-item">
            <div class="stat-value">${noShowCount}</div>
            <div class="stat-label">No Show</div>
          </div>
          <div class="stat-item">
            <div class="stat-value">${pendingCount}</div>
            <div class="stat-label">Pending</div>
          </div>
          <div class="stat-item">
            <div class="stat-value">${showRate}%</div>
            <div class="stat-label">Show Rate</div>
          </div>
        </div>
        
        ${appointments.length > 0 ? `
          <table>
            <thead>
              <tr>
                <th>No.</th>
                <th>Name</th>
                <th>Office</th>
                <th>Appt. Date</th>
                <th>Time</th>
                <th>Visit Type</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              ${appointments.map((apt: any, index: number) => {
                const safeName = sanitizeString(apt.name, 100);
                const safeOffice = sanitizeString(apt.office, 50);
                const safeApptDate = sanitizeString(apt.appt_date, 20);
                const safeTime = sanitizeString(apt.time, 20);
                const safeVisitType = sanitizeString(apt.visit_type, 50);
                const safeShowStatus = apt.showStatus || 'pending';
                
                return `
                <tr>
                  <td>${index + 1}</td>
                  <td>${escapeHtml(safeName) || '-'}</td>
                  <td>${escapeHtml(safeOffice) || '-'}</td>
                  <td>${escapeHtml(safeApptDate) || '-'}</td>
                  <td style="text-align: center;">${escapeHtml(convertTo12Hour(safeTime))}</td>
                  <td>${escapeHtml(safeVisitType) || '-'}</td>
                  <td class="${safeShowStatus}">${
                    safeShowStatus === 'show' ? 'Show' :
                    safeShowStatus === 'no-show' ? 'No Show' : 'Pending'
                  }</td>
                </tr>
              `;
              }).join('')}
            </tbody>
          </table>
        ` : `
          <div style="text-align: center; padding: 40px; color: #666;">
            No appointments found for the selected criteria.
          </div>
        `}
        
        <div class="footer">
          Generated: ${new Date(timestamp).toLocaleDateString('en-US', {
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
    
    // 파일명 생성 (보안 검증 적용)
    const filename = sanitizePdfFilename(`${safeStartDate === safeEndDate ? safeStartDate : `${safeStartDate}_to_${safeEndDate}`}_${safeSelectedOffice}_Show_Check_Report`);
    
    return new NextResponse(Buffer.from(pdf), {
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${filename}"`
      }
    });

  } catch (error) {
    logError(error, 'generate-show-check');
    return NextResponse.json({ 
      success: false, 
      error: 'Show check PDF 생성 중 오류가 발생했습니다.' 
    });
  }
}
