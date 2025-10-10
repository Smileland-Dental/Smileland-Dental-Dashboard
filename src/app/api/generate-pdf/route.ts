import { NextRequest, NextResponse } from 'next/server';
import puppeteer from 'puppeteer';

export async function POST(request: NextRequest) {
  console.log('✅ PDF API POST 요청 받음!');
  
  try {
    const { patientData } = await request.json();
    console.log('📋 받은 데이터:', patientData);
    
    // patientLogs 또는 patientRows 둘 다 처리하고 빈 행 필터링
    const allPatients = patientData.patientLogs || patientData.patientRows || [];
    const patientList = allPatients.filter(row => 
      row.name || row.office || row.appt_date || row.apptDate || row.visit_type || row.visitType || 
      row.call_in || row.callIn || row.call_out || row.callOut || row.time || row.remark || 
      row.other_duty || row.otherDuty
    );
    
    const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>${patientData.dutyDate}_${patientData.userName}_${patientData.workOffice}_Patient Log</title>
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
            <span class="info-value">${patientData.dutyDate || '-'}</span>
          </div>
            <div class="info-item">
              <span class="info-label">Name:</span>
            <span class="info-value">${patientData.userName || '-'}</span>
          </div>
          </div>
          <div class="info-column">
            <div class="info-item">
              <span class="info-label">Work Office:</span>
              <span class="info-value">${patientData.workOffice || '-'}</span>
            </div>
            <div class="info-item">
              <span class="info-label">Work Hours:</span>
              <span class="info-value">${patientData.workHoursFrom || '-'} - ${patientData.workHoursTo || '-'}</span>
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
                return allPatients.filter(row => row.appt_date && row.name).length;
              })()}
            </div>
          </div>
          <div class="count-item">
            <div class="count-label">Incoming Calls</div>
            <div class="count-number">
              ${(() => {
                const allPatients = patientData.patientLogs || patientData.patientRows || [];
                return allPatients.filter(row => row.call_in).length;
              })()}
            </div>
          </div>
          <div class="count-item">
            <div class="count-label">Outgoing Calls</div>
            <div class="count-number">
              ${(() => {
                const allPatients = patientData.patientLogs || patientData.patientRows || [];
                return allPatients.filter(row => row.call_out).length;
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
              ${patientList.map((row, index) => `
                <tr>
                  <td class="number-cell">${index + 1}</td>
                  <td>${(row.name || '-').length > 12 ? (row.name || '-').substring(0, 12) : (row.name || '-')}</td>
                  <td>${row.office || '-'}</td>
                  <td style="text-align: center;">${row.appt_date || row.apptDate || '-'}</td>
                  <td>${(row.visit_type || row.visitType || '-').length > 8 ? (row.visit_type || row.visitType || '-').substring(0, 8) : (row.visit_type || row.visitType || '-')}</td>
                  <td class="checkbox-cell">${row.call_in || row.callIn ? '✓' : ''}</td>
                  <td class="checkbox-cell">${row.call_out || row.callOut ? '✓' : ''}</td>
                  <td style="text-align: center;">${row.time || '-'}</td>
                  <td>${(row.remark || '-').length > 15 ? (row.remark || '-').substring(0, 15) : (row.remark || '-')}</td>
                  <td>${(row.other_duty || row.otherDuty || '-').length > 15 ? (row.other_duty || row.otherDuty || '-').substring(0, 15) : (row.other_duty || row.otherDuty || '-')}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
        ` : `
          <p style="text-align: center; padding: 20px; font-style: italic;">
            No patient data recorded.
          </p>
        `}

        ${patientData.dailyWorkReport ? `
          <div class="daily-report">
            <h3>Daily Work Report</h3>
            <div class="daily-report-content">
              ${patientData.dailyWorkReport.replace(/\n/g, '<br>')}
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
            minute: '2-digit'
          })}</div>
        </div>
      </body>
      </html>
    `;

    const filename = `${patientData.dutyDate}_${patientData.userName}_${patientData.workOffice}_Patient_Log`;
    
    console.log('🎯 PDF HTML 생성 완료');
    
    // Puppeteer로 PDF 생성
    const browser = await puppeteer.launch({
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    
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
    
    await browser.close();
    
    // 파일명 생성
    const pdfFilename = `${filename}.pdf`;
    
    return new NextResponse(pdf, {
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${pdfFilename}"`
      }
    });

  } catch (error) {
    console.error('❌ PDF 생성 오류:', error);
    return NextResponse.json(
      { 
        success: false, 
        error: `PDF 생성 실패: ${error.message}`,
        details: error.stack 
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
