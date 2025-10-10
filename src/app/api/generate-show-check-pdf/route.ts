import { NextRequest, NextResponse } from 'next/server';
import puppeteer from 'puppeteer';

export async function POST(request: NextRequest) {
  try {
    const { showCheckData } = await request.json();

    if (!showCheckData) {
      return NextResponse.json({ success: false, error: 'Show check data is required' });
    }

    const { startDate, endDate, selectedOffice, appointments, generatedBy, timestamp } = showCheckData;

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
        <title>${startDate === endDate ? startDate : `${startDate}_to_${endDate}`}_${selectedOffice}_Show_Check_Report</title>
        <style>
          @media print {
            body { margin: 0.3in; font-size: 8px; line-height: 1.1; }
            table { font-size: 7px; }
            .header { font-size: 10px; margin-bottom: 10px; }
            .stats { font-size: 8px; margin: 10px 0; }
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
          
          .info-section {
            display: flex;
            justify-content: space-between;
            margin-bottom: 15px;
            font-size: 10px;
          }
          
          .stats {
            display: flex;
            justify-content: center;
            gap: 20px;
            margin: 15px 0;
            padding: 10px;
            border: 1px solid #ddd;
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
            font-size: 10px;
            color: #666;
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
          }
          
          th {
            background: #f0f0f0;
            font-weight: bold;
          }
          
          .show { color: #333; font-weight: bold; }
          .no-show { color: #333; font-weight: bold; }
          .pending { color: #333; font-weight: bold; }
          
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
          <h1>Show/No Show Check Report</h1>
        </div>
        
        <div class="info-section">
          <div><strong>Appointment Date Range:</strong> ${startDate === endDate ? startDate : `${startDate} to ${endDate}`}</div>
          <div><strong>Appointment Office:</strong> ${selectedOffice}</div>
          <div><strong>Checked by:</strong> ${generatedBy}</div>
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
                <th>Visit Type</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              ${appointments.map((apt: any, index: number) => `
                <tr>
                  <td>${index + 1}</td>
                  <td>${apt.name || '-'}</td>
                  <td>${apt.office || '-'}</td>
                  <td>${apt.appt_date || '-'}</td>
                  <td>${apt.visit_type || '-'}</td>
                  <td class="${apt.showStatus || 'pending'}">${
                    apt.showStatus === 'show' ? 'Show' :
                    apt.showStatus === 'no-show' ? 'No Show' : 'Pending'
                  }</td>
                </tr>
              `).join('')}
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
            minute: '2-digit'
          })}
        </div>
      </body>
      </html>
    `;

    // Puppeteer로 PDF 생성
    const browser = await puppeteer.launch({
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    
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
    
    // 파일명 생성
    const filename = `${startDate === endDate ? startDate : `${startDate}_to_${endDate}`}_${selectedOffice}_Show_Check_Report.pdf`;
    
    return new NextResponse(pdf, {
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${filename}"`
      }
    });

  } catch (error) {
    console.error('Show check PDF generation error:', error);
    return NextResponse.json({ 
      success: false, 
      error: 'Failed to generate show check PDF' 
    });
  }
}
