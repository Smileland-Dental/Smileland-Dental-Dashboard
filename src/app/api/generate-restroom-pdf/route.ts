import { NextRequest, NextResponse } from 'next/server';
import { 
  escapeHtml, 
  sanitizeString, 
  getSecurityHeaders
} from '@/lib/security-server';

export async function POST(request: NextRequest) {
  try {
    const { inspectionDate, selectedOffice, selectedRestroom, restroomData } = await request.json();

    if (!inspectionDate || !selectedOffice || !selectedRestroom || !restroomData) {
      return NextResponse.json({ 
        success: false, 
        error: 'Inspection date, office, restroom, and data are required' 
      }, {
        status: 400,
        headers: getSecurityHeaders()
      });
    }
    
    // 데이터 sanitize
    const safeInspectionDate = sanitizeString(inspectionDate, 20);
    const safeSelectedOffice = sanitizeString(selectedOffice, 100);
    const safeSelectedRestroom = sanitizeString(selectedRestroom, 50);

    // 컬럼 순서: Time, Check, Pick up Paper, Wipe Sinks and Mirrors, Wipe Toilets, Wipe Baby Table, Empty Trash, Toilet Paper, Soap, Toilet Seat Covers, Refresh Spray, Checked Time
    const COLUMN_NAMES = [
      'Time', 'Check', 'Pick up Paper', 'Wipe Sinks and Mirrors', 'Wipe Toilets',
      'Wipe Baby Table', 'Empty Trash', 'Toilet Paper', 'Soap',
      'Toilet Seat Covers', 'Refresh Spray', 'Checked Time'
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

    // HTML 템플릿 생성
    const html = `
      <!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>${escapeHtml(safeInspectionDate)}_${escapeHtml(safeSelectedOffice)}_Restroom ${escapeHtml(safeSelectedRestroom)} Inspection Log</title>
        <style>
          @media print {
            body { margin: 0.3in; font-size: 8px; line-height: 1.1; }
            table { font-size: 7px; }
            .header { font-size: 10px; margin-bottom: 10px; }
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
            vertical-align: middle;
          }
          
          th {
            background: #f0f0f0;
            font-weight: bold;
          }
          
          .spotless-header {
            background: #e3e8f0;
            font-weight: 700;
            color: #4a6fa1;
          }
          
          .stocked-header {
            background: #f3f7fa;
            font-weight: 700;
            color: #4a6fa1;
          }
          
          .spotless-subheader {
            background: #f7fafd;
            font-size: 7px;
            font-weight: 400;
            color: #4a6fa1;
          }
          
          .stocked-subheader {
            background: #fafdff;
            font-size: 7px;
            font-weight: 400;
            color: #4a6fa1;
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
          <h1>Restroom Inspection Log</h1>
          <p>${escapeHtml(safeSelectedOffice)} - Restroom ${escapeHtml(safeSelectedRestroom)} (${escapeHtml(safeInspectionDate)})</p>
        </div>
        
        <div class="info-section">
          <div><strong>Date:</strong> ${escapeHtml(safeInspectionDate)}</div>
          <div><strong>Office:</strong> ${escapeHtml(safeSelectedOffice)}</div>
          <div><strong>Restroom:</strong> ${escapeHtml(safeSelectedRestroom)}</div>
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
              <th colspan="2"></th>
              <th colspan="4" class="spotless-header">SPOTLESS</th>
              <th colspan="5" class="stocked-header">STOCKED</th>
              <th></th>
            </tr>
            <tr>
              <th>Time</th>
              <th>Check</th>
              <th>Pick up Paper</th>
              <th>Wipe Sinks and Mirrors</th>
              <th>Wipe Toilets</th>
              <th>Wipe Baby Table</th>
              <th>Empty Trash</th>
              <th>Toilet Paper</th>
              <th>Soap</th>
              <th>Toilet Seat Covers</th>
              <th>Refresh Spray</th>
              <th>Checked Time</th>
            </tr>
            <tr>
              <th colspan="2"></th>
              <th colspan="4" class="spotless-subheader">Perform each hour</th>
              <th colspan="5" class="stocked-subheader">Replenish as needed</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            ${ROW_HEADERS.map((header, rowIndex) => {
              const isMopRow = header.toLowerCase().includes('sweep/mop');
              const checkValue = sanitizeString(restroomData[`Row${rowIndex + 1}_Check`] || '', 50);
              const checkedTime = sanitizeString(restroomData[`Row${rowIndex + 1}_CheckedTime`] || '', 50);
              
              return `
                <tr>
                  <td>${escapeHtml(header)}</td>
                  <td>${escapeHtml(checkValue)}</td>
                  ${COLUMN_NAMES.slice(2, -1).map((columnName, colIndex) => {
                    if (isMopRow) {
                      return '<td></td>';
                    }
                    const cellValue = restroomData[`Row${rowIndex + 1}_Col${colIndex + 3}`];
                    return `<td>${cellValue === true ? '✔' : ''}</td>`;
                  }).join('')}
                  <td>${escapeHtml(checkedTime)}</td>
                </tr>
              `;
            }).join('')}
          </tbody>
        </table>
        
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
        
        <script>
          // 페이지 로드 후 자동으로 인쇄 대화상자 열기
          window.onload = function() {
            // 약간의 지연 후 인쇄 대화상자 열기 (스타일이 완전히 로드되도록)
            setTimeout(function() {
              window.print();
            }, 500);
          };
          
          // 인쇄 대화상자가 닫힌 후 부모 창에 메시지 전송
          window.addEventListener('afterprint', function() {
            if (window.opener) {
              window.opener.postMessage('print-completed', '*');
            }
          });
        </script>
      </body>
      </html>
    `;

    // HTML 반환 (클라이언트에서 인쇄/저장 가능)
    return new NextResponse(html, {
      headers: {
        'Content-Type': 'text/html; charset=utf-8',
        ...getSecurityHeaders(),
        'Cache-Control': 'no-cache, no-store, must-revalidate',
        'Pragma': 'no-cache',
        'Expires': '0'
      }
    });

  } catch (error) {
    return NextResponse.json({ 
      success: false, 
      error: 'Failed to generate restroom inspection report' 
    }, {
      status: 500,
      headers: getSecurityHeaders()
    });
  }
}
