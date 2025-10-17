import { NextRequest, NextResponse } from 'next/server';
import puppeteer from 'puppeteer';
import { Buffer } from 'buffer';
import { escapeHtml, sanitizeString, safeLog, logError, sanitizePdfFilename, getSecurePuppeteerOptions, validatePdfGeneration, sanitizeArrayForPdf } from '@/lib/security-server';

export async function POST(request: NextRequest) {
  safeLog('✅ PDF API POST 요청 받음!');
  
  try {
    const { patientData } = await request.json();
    safeLog('📋 받은 데이터:', { hasData: !!patientData });
    
    // PDF 생성 데이터 크기 검증
    const dataSize = JSON.stringify(patientData).length;
    if (!validatePdfGeneration(dataSize)) {
      throw new Error('Data size too large for PDF generation');
    }
    
    // patientLogs 또는 patientRows 둘 다 처리하고 빈 행 필터링 (보안 강화)
    const allPatients = patientData.patientLogs || patientData.patientRows || [];
    const safePatients = sanitizeArrayForPdf(allPatients, 1000); // 최대 1000개 행으로 제한
    const patientList = safePatients.filter((row: any) => 
      row && (
        row.name || row.office || row.appt_date || row.apptDate || row.visit_type || row.visitType || 
        row.call_in || row.callIn || row.call_out || row.callOut || row.time || row.remark || 
        row.other_duty || row.otherDuty
      )
    );
    
    // 데이터 안전성 검증 및 이스케이프
    const safeDutyDate = sanitizeString(patientData.dutyDate, 50);
    const safeUserName = sanitizeString(patientData.userName, 100);
    const safeWorkOffice = sanitizeString(patientData.workOffice, 100);
    const safeWorkHoursFrom = sanitizeString(patientData.workHoursFrom, 20);
    const safeWorkHoursTo = sanitizeString(patientData.workHoursTo, 20);
    
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
    const safeDailyWorkReport = sanitizeString(patientData.dailyWorkReport, 2000);

    const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>${escapeHtml(safeDutyDate)}_${escapeHtml(safeUserName)}_${escapeHtml(safeWorkOffice)}_Patient Log</title>
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
          .count-summary {
            display: flex;
            justify-content: center;
            gap: 30px;
            margin-bottom: 15px;
            padding: 8px;
            border: 1px solid #ccc;
            background: #f9f9f9;
          }
          .count-item {
            text-align: center;
          }
          .count-label {
            font-weight: bold;
            font-size: 11px;
          }
          .count-number {
            font-size: 16px;
            font-weight: bold;
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
          .checkbox-cell {
            text-align: center;
            font-size: 12px;
          }
          .daily-report { 
            margin-top: 20px;
            padding: 10px;
            border: 1px solid #ccc;
          }
          .daily-report h3 { 
            color: #000; 
            margin-top: 0;
            font-size: 14px;
            font-weight: bold;
          }
          .daily-report-content {
            font-size: 11px;
            line-height: 1.4;
          }
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
            .info-grid { 
              display: flex;
              justify-content: space-between;
              gap: 5px; 
              margin-bottom: 8px;
            }
            .info-card {
              padding: 3px;
              margin-bottom: 0;
              flex: 1;
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
            .patient-summary {
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
            .checkbox-cell {
              font-size: 8px;
              text-align: center;
            }
            .daily-report {
              margin-top: 10px;
              padding: 5px;
              font-size: 8px;
            }
            .daily-report h3 {
              font-size: 9px;
              margin: 0 0 3px 0;
            }
            .daily-report-content {
              padding: 5px;
              font-size: 7px;
              line-height: 1.2;
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
          <h1>Patient Log Report</h1>
          <p class="subtitle">Smileland Dental</p>
        </div>

        <div class="info-section">
          <div class="info-column">
            <div class="info-item">
              <span class="info-label">Duty Date:</span>
            <span class="info-value">${escapeHtml(safeDutyDate) || '-'}</span>
          </div>
            <div class="info-item">
              <span class="info-label">Name:</span>
            <span class="info-value">${escapeHtml(safeUserName) || '-'}</span>
          </div>
          </div>
          <div class="info-column">
            <div class="info-item">
              <span class="info-label">Work Office:</span>
              <span class="info-value">${escapeHtml(safeWorkOffice) || '-'}</span>
            </div>
            <div class="info-item">
              <span class="info-label">Work Hours:</span>
              <span class="info-value">${escapeHtml(convertTo12Hour(safeWorkHoursFrom))} - ${escapeHtml(convertTo12Hour(safeWorkHoursTo))}</span>
            </div>
          </div>
        </div>

        <h2 class="section-title">Patient Log Summary</h2>
        
        <div class="count-summary">
          <div class="count-item">
            <div class="count-label">Total Appointments</div>
            <div class="count-number">
              ${(() => {
                const allPatients = patientData.patientLogs || patientData.patientRows || [];
                return allPatients.filter((row: any) => row.appt_date && row.name).length;
              })()}
            </div>
          </div>
          <div class="count-item">
            <div class="count-label">Incoming Calls</div>
            <div class="count-number">
              ${(() => {
                const allPatients = patientData.patientLogs || patientData.patientRows || [];
                return allPatients.filter((row: any) => row.call_in).length;
              })()}
            </div>
          </div>
          <div class="count-item">
            <div class="count-label">Outgoing Calls</div>
            <div class="count-number">
              ${(() => {
                const allPatients = patientData.patientLogs || patientData.patientRows || [];
                return allPatients.filter((row: any) => row.call_out).length;
              })()}
            </div>
          </div>
        </div>

        <h2 class="section-title">Patient Details</h2>
        
        ${patientList && patientList.length > 0 ? `
        <table>
          <thead>
            <tr>
                <th style="width: 4%;">No.</th>
                <th style="width: 15%;">Patient's Name</th>
                <th style="width: 8%;">Office</th>
                <th style="width: 10%;">Appt. Date</th>
                <th style="width: 12%;">Type of Visit</th>
                <th style="width: 6%;">Call In</th>
                <th style="width: 6%;">Call Out</th>
                <th style="width: 7%;">Time</th>
                <th style="width: 16%;">Remark</th>
                <th style="width: 16%;">Other Duty</th>
            </tr>
          </thead>
          <tbody>
              ${patientList.map((row: any, index: number) => {
                // 안전한 데이터 검증 및 정제
                const safeName = sanitizeString(row?.name, 50);
                const safeOffice = sanitizeString(row?.office, 50);
                const safeApptDate = sanitizeString(row?.appt_date || row?.apptDate, 20);
                const safeVisitType = sanitizeString(row?.visit_type || row?.visitType, 50);
                const safeTime = sanitizeString(row?.time, 20);
                const safeRemark = sanitizeString(row?.remark, 100);
                const safeOtherDuty = sanitizeString(row?.other_duty || row?.otherDuty, 100);
                
                // 체크박스 값 안전하게 처리
                const callIn = Boolean(row?.call_in || row?.callIn);
                const callOut = Boolean(row?.call_out || row?.callOut);
                
                return `
                <tr>
                  <td class="number-cell">${index + 1}</td>
                  <td>${escapeHtml(safeName).length > 12 ? escapeHtml(safeName).substring(0, 12) : escapeHtml(safeName) || '-'}</td>
                  <td>${escapeHtml(safeOffice) || '-'}</td>
                  <td style="text-align: center;">${escapeHtml(safeApptDate) || '-'}</td>
                  <td>${escapeHtml(safeVisitType).length > 8 ? escapeHtml(safeVisitType).substring(0, 8) : escapeHtml(safeVisitType) || '-'}</td>
                  <td class="checkbox-cell">${callIn ? '✓' : ''}</td>
                  <td class="checkbox-cell">${callOut ? '✓' : ''}</td>
                  <td style="text-align: center;">${escapeHtml(convertTo12Hour(safeTime))}</td>
                  <td>${escapeHtml(safeRemark).length > 15 ? escapeHtml(safeRemark).substring(0, 15) : escapeHtml(safeRemark) || '-'}</td>
                  <td>${escapeHtml(safeOtherDuty).length > 15 ? escapeHtml(safeOtherDuty).substring(0, 15) : escapeHtml(safeOtherDuty) || '-'}</td>
              </tr>
            `;
              }).join('')}
          </tbody>
        </table>
        ` : `
          <p style="text-align: center; padding: 20px; font-style: italic;">
            No patient data recorded.
          </p>
        `}

        ${safeDailyWorkReport ? `
          <div class="daily-report">
            <h3>Daily Work Report</h3>
            <div class="daily-report-content">
              ${escapeHtml(safeDailyWorkReport).replace(/\n/g, '<br>')}
            </div>
          </div>
        ` : ''}

        <div class="footer">
          <div class="clinic-name">Smileland Dental</div>
          <div class="date">Generated: ${new Date().toLocaleDateString('en-US', { 
            year: 'numeric', 
            month: 'long', 
            day: 'numeric',
            weekday: 'long',
            hour: '2-digit',
            minute: '2-digit',
            hour12: true
          })}</div>
        </div>
      </body>
      </html>
    `;

    const filename = sanitizePdfFilename(`${safeDutyDate}_${safeUserName}_${safeWorkOffice}_Patient_Log`);
    
    safeLog('🎯 PDF HTML 생성 완료');
    
    // Puppeteer로 PDF 생성 (보안 옵션 적용)
    const browser = await puppeteer.launch(getSecurePuppeteerOptions());
    
    const page = await browser.newPage();
    await page.setContent(htmlContent);
    
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
    
    // 브라우저 안전하게 종료
    await browser.close();
    
    // 파일명 생성 (이미 sanitizePdfFilename에서 .pdf 추가됨)
    const pdfFilename = filename;
    
    // PDF 응답 생성 (보안 헤더 추가)
    return new NextResponse(Buffer.from(pdf), {
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${pdfFilename}"`,
        'X-Content-Type-Options': 'nosniff',
        'X-Frame-Options': 'DENY',
        'Cache-Control': 'no-cache, no-store, must-revalidate',
        'Pragma': 'no-cache',
        'Expires': '0'
      }
    });

  } catch (error) {
    logError(error, 'generate-pdf');
    return NextResponse.json(
      { 
        success: false, 
        error: 'PDF 생성 중 오류가 발생했습니다.' 
      },
      { status: 500 }
    );
  }
}

// GET 요청 처리 (테스트용)
export async function GET() {
  return NextResponse.json({ 
    message: '🚀 PDF API가 정상 작동 중입니다!', 
    timestamp: new Date().toISOString(),
    status: 'healthy'
  });
}
