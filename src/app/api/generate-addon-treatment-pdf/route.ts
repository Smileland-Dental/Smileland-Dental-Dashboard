import { NextRequest, NextResponse } from 'next/server';
import puppeteer from 'puppeteer';

export async function POST(request: NextRequest) {
  try {
    const { dutyDate, patientRows } = await request.json();

    if (!dutyDate || !patientRows) {
      return NextResponse.json({ success: false, error: 'Duty date and patient data are required' });
    }

    // HTML 템플릿 생성
    const html = `
      <!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>${dutyDate}_Bernard_Add On Treatment</title>
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
          <div><strong>Date:</strong> ${dutyDate}</div>
          <div><strong>Location:</strong> Bernard</div>
          <div><strong>Generated:</strong> ${new Date().toLocaleDateString('en-US', {
            year: 'numeric',
            month: 'long',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
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
            ${patientRows.map((row: any, index: number) => `
              <tr>
                <td>${index + 1}</td>
                <td>${row.patientName || ''}</td>
                <td>${row.dob || ''}</td>
                <td>${row.time || ''}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
        
        <div class="footer">
          <p>Smileland Dental</p>
          <p>Report generated on ${new Date().toLocaleDateString('en-US', {
            year: 'numeric',
            month: 'long',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
          })}</p>
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
    const filename = `4) ${dutyDate}_Bernard_Add On Treatment.pdf`;
    
    return new NextResponse(pdf, {
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${filename}"`
      }
    });

  } catch (error) {
    console.error('Add-on treatment PDF generation error:', error);
    return NextResponse.json({ 
      success: false, 
      error: 'Failed to generate add-on treatment PDF' 
    });
  }
}
