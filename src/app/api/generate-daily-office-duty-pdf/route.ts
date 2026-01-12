import { NextRequest, NextResponse } from 'next/server';
import puppeteer from 'puppeteer';

// 🔒 보안: HTML 이스케이프 함수 (XSS 방지)
function escapeHtml(text: string | undefined | null): string {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// 🔒 보안: 입력 검증 함수
function validateInput(input: any, maxLength: number = 500): boolean {
  if (typeof input !== 'string') return false;
  if (input.length > maxLength) return false;
  return true;
}

// 🔒 보안: 날짜 형식 검증
function validateDate(date: string): boolean {
  const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
  if (!dateRegex.test(date)) return false;
  const parsedDate = new Date(date);
  return !isNaN(parsedDate.getTime());
}

// 🔒 보안: 오피스 값 검증
function validateOffice(office: string | undefined): boolean {
  if (!office) return true; // 선택적 필드
  const allowedOffices = ['Ming', 'Bernard', 'Delano', 'Tulare', 'Visalia', 'Fresno', 'California', 'Ortho'];
  return allowedOffices.includes(office);
}

export async function POST(request: NextRequest) {
  let browser;
  try {
    const { dutyDate, selectedOffice, dutyData, submitPassword } = await request.json();

    // 🔒 보안: 서버 측 검증
    if (!dutyDate || !dutyData) {
      return NextResponse.json({ success: false, error: 'Duty date and data are required' }, { status: 400 });
    }

    // 🔒 보안: 날짜 형식 검증
    if (!validateDate(dutyDate)) {
      return NextResponse.json({ success: false, error: 'Invalid date format' }, { status: 400 });
    }

    // 🔒 보안: 오피스 값 검증
    if (!validateOffice(selectedOffice)) {
      return NextResponse.json({ success: false, error: 'Invalid office value' }, { status: 400 });
    }

    // 🔒 보안: 비밀번호 검증 (서버 측에서도 확인)
    const expectedPassword = 'Halloween';
    if (!submitPassword || submitPassword !== expectedPassword) {
      return NextResponse.json({ success: false, error: 'Unauthorized: Invalid password' }, { status: 401 });
    }

    // 🔒 보안: 사용자 입력 검증 및 이스케이프 (XSS 방지)
    const safeDutyDate = escapeHtml(dutyDate);
    const safeSelectedOffice = escapeHtml(selectedOffice || '');
    
    // dutyData의 모든 값 검증 및 이스케이프
    const safeDutyData: { [key: string]: string } = {};
    for (const [key, value] of Object.entries(dutyData)) {
      if (validateInput(value, 500)) {
        // 줄바꿈은 <br>로 변환하되, 먼저 이스케이프
        safeDutyData[key] = escapeHtml(value as string).replace(/\n/g, '<br>');
      } else {
        // 검증 실패 시 빈 문자열
        safeDutyData[key] = '';
      }
    }

    // HTML 템플릿 생성 (프린트용)
    const html = `
      <!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <meta http-equiv="Content-Security-Policy" content="default-src 'self'; script-src 'none'; object-src 'none'; base-uri 'self'; style-src 'unsafe-inline';">
        <title>${safeDutyDate}_${safeSelectedOffice}_Daily_Office_Duty</title>
        <style>
          @media print {
            body { margin: 0.25in; font-size: 9px; line-height: 1.3; }
            table { font-size: 8px; }
            .header { font-size: 12px; margin-bottom: 8px; }
            .info-section { font-size: 9px; margin-bottom: 8px; }
            th, td { padding: 4px 5px; }
          }
          
          body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 15px;
            background: white;
            color: #000;
            line-height: 1.4;
            font-size: 9px;
          }
          
          .header {
            text-align: center;
            border-bottom: 1px solid #000;
            padding-bottom: 8px;
            margin-bottom: 10px;
          }
          
          .header h1 {
            margin: 0;
            font-size: 16px;
            font-weight: bold;
            color: #000;
          }
          
          .info-section {
            display: flex;
            justify-content: space-between;
            margin-bottom: 10px;
            font-size: 9px;
            color: #000;
          }
          
          table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 8px;
            background-color: white;
            font-size: 8px;
          }
          
          th, td {
            border: 1px solid #000;
            padding: 4px 5px;
            text-align: left;
            vertical-align: top;
            color: #000;
          }
          
          th {
            background-color: #f0f0f0;
            color: #000;
            font-weight: bold;
          }
          
          td {
            background-color: white;
            color: #000;
          }
          
          .duty-details {
            font-size: 7px;
            color: #000;
            margin-top: 2px;
          }
          
          .footer {
            margin-top: 15px;
            text-align: center;
            font-size: 8px;
            color: #000;
            border-top: 1px solid #000;
            padding-top: 8px;
          }
        </style>
      </head>
      <body>
        <div class="header">
          <h1>Daily Office Duty</h1>
        </div>
        
        <div class="info-section">
          <div><strong>Date:</strong> ${safeDutyDate}</div>
          <div><strong>Location:</strong> ${safeSelectedOffice}</div>
          <div><strong>Generated:</strong> ${new Date().toLocaleString('en-US', {
            timeZone: 'America/Los_Angeles',
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
              <th>Duty Description</th>
              <th>Details</th>
              <th style="width: 80px;">Done By</th>
              <th style="width: 80px;">Checked By</th>
              <th style="width: 60px;">Time</th>
            </tr>
          </thead>
          <tbody>
            <!-- Row 1 -->
            <tr>
              <td>1</td>
              <td>Turn Off Answering Service</td>
              <td>-</td>
              <td>${safeDutyData.Row1_Done || ''}</td>
              <td>${safeDutyData.Row1_Checked || ''}</td>
              <td>${safeDutyData.Row1_Time || ''}</td>
            </tr>
            <!-- Row 2 -->
            <tr>
              <td>2</td>
              <td>All charts filed back?</td>
              <td>${safeDutyData.Row2_YesNo || ''}</td>
              <td>${safeDutyData.Row2_Done || ''}</td>
              <td>${safeDutyData.Row2_Checked || ''}</td>
              <td>${safeDutyData.Row2_Time || ''}</td>
            </tr>
            <!-- Row 3 -->
            <tr>
              <td>3</td>
              <td>Charts pulled for next day</td>
              <td>-</td>
              <td>${safeDutyData.Row3_Done || ''}</td>
              <td>${safeDutyData.Row3_Checked || ''}</td>
              <td>${safeDutyData.Row3_Time || ''}</td>
            </tr>
            <!-- Row 4 -->
            <tr>
              <td>4</td>
              <td>Check eligibility <br><span class="duty-details">1st of every month come in early to check eligibility by 8:30 am</span></td>
              <td>-</td>
              <td>${safeDutyData.Row4_Done || ''}</td>
              <td>${safeDutyData.Row4_Checked || ''}</td>
              <td>${safeDutyData.Row4_Time || ''}</td>
            </tr>
            <!-- Row 5 -->
            <tr>
              <td>5</td>
              <td>If pt is not eligible call and inform</td>
              <td>How many pt's did you call: ${safeDutyData.Row5_CallNum || ''}</td>
              <td>${safeDutyData.Row5_Done || ''}</td>
              <td>${safeDutyData.Row5_Checked || ''}</td>
              <td>${safeDutyData.Row5_Time || ''}</td>
            </tr>
            <!-- Row 6 -->
            <tr>
              <td>6</td>
              <td>Insurance breakdown for next day's patients <br><span class="duty-details">Call and get ins. info if necessary</span></td>
              <td>-</td>
              <td>${safeDutyData.Row6_Done || ''}</td>
              <td>${safeDutyData.Row6_Checked || ''}</td>
              <td>${safeDutyData.Row6_Time || ''}</td>
            </tr>
            <!-- Row 7 -->
            <tr>
              <td>7</td>
              <td>Check ledger for any balance on the account <br><span class="duty-details">Fill out "Account with Balances Form" and fax </span><br><span class="duty-details">Called to inform patient of balance?</span></td>
              <td>${safeDutyData.Row7_YesNo || ''}</td>
              <td>${safeDutyData.Row7_Done || ''}</td>
              <td>${safeDutyData.Row7_Checked || ''}</td>
              <td>${safeDutyData.Row7_Time || ''}</td>
            </tr>
            <!-- Row 8 -->
            <tr>
              <td>8</td>
              <td>Morning confirmations <br><span class="duty-details">At least by noon</span></td>
              <td>-</td>
              <td>${safeDutyData.Row8_Done || ''}</td>
              <td>${safeDutyData.Row8_Checked || ''}</td>
              <td>${safeDutyData.Row8_Time || ''}</td>
            </tr>
            <!-- Row 9 -->
            <tr>
              <td>9</td>
              <td>No shows entered on ledger</td>
              <td>-</td>
              <td>${safeDutyData.Row9_Done || ''}</td>
              <td>${safeDutyData.Row9_Checked || ''}</td>
              <td>${safeDutyData.Row9_Time || ''}</td>
            </tr>
            <!-- Row 10 -->
            <tr>
              <td>10</td>
              <td>No shows stamped in patient charts</td>
              <td>-</td>
              <td>${safeDutyData.Row10_Done || ''}</td>
              <td>${safeDutyData.Row10_Checked || ''}</td>
              <td>${safeDutyData.Row10_Time || ''}</td>
            </tr>
            <!-- Row 11 -->
            <tr>
              <td>11</td>
              <td>Reconfirming completed? <br><span class="duty-details">Start at 4:00pm</span></td>
              <td>-</td>
              <td>${safeDutyData.Row11_Done || ''}</td>
              <td>${safeDutyData.Row11_Checked || ''}</td>
              <td>${safeDutyData.Row11_Time || ''}</td>
            </tr>
            <!-- Row 12 -->
            <tr>
              <td>12</td>
              <td>One week reminders completed?</td>
              <td>-</td>
              <td>${safeDutyData.Row12_Done || ''}</td>
              <td>${safeDutyData.Row12_Checked || ''}</td>
              <td>${safeDutyData.Row12_Time || ''}</td>
            </tr>
            <!-- Row 13 -->
            <tr>
              <td>13</td>
              <td>Call all treatment patients from today for post op</td>
              <td>-</td>
              <td>${safeDutyData.Row13_Done || ''}</td>
              <td>${safeDutyData.Row13_Checked || ''}</td>
              <td>${safeDutyData.Row13_Time || ''}</td>
            </tr>
            <!-- Row 14 -->
            <tr>
              <td>14</td>
              <td>Total lab case deposits/deliveries <br><span class="duty-details">Name/DOB</span></td>
              <td>${safeDutyData['Row14_Name/DOB'] || ''}</td>
              <td>${safeDutyData.Row14_Done || ''}</td>
              <td>${safeDutyData.Row14_Checked || ''}</td>
              <td>${safeDutyData.Row14_Time || ''}</td>
            </tr>
            <!-- Row 15 -->
            <tr>
              <td>15</td>
              <td>Check all undelivered lab cases and make appointments <br><span class="duty-details">Any Lab case that is more than 3 weeks old must be sent to corporate along with $20 deposit</span></td>
              <td>${safeDutyData.Row15_LabCases || ''}</td>
              <td>${safeDutyData.Row15_Done || ''}</td>
              <td>${safeDutyData.Row15_Checked || ''}</td>
              <td>${safeDutyData.Row15_Time || ''}</td>
            </tr>
            <!-- Row 16 -->
            <tr>
              <td>16</td>
              <td>Check all lab cases for next day <br><span class="duty-details">Call lab for next day pick up's</span></td>
              <td>-</td>
              <td>${safeDutyData.Row16_Done || ''}</td>
              <td>${safeDutyData.Row16_Checked || ''}</td>
              <td>${safeDutyData.Row16_Time || ''}</td>
            </tr>
            <!-- Row 17 -->
            <tr>
              <td>17</td>
              <td>N₂O/ Compressor Off</td>
              <td>-</td>
              <td>${safeDutyData.Row17_Done || ''}</td>
              <td>${safeDutyData.Row17_Checked || ''}</td>
              <td>${safeDutyData.Row17_Time || ''}</td>
            </tr>
            <!-- Row 18 -->
            <tr>
              <td>18</td>
              <td>Did you read the meter on the Oxygen/N₂O/Helium tank?</td>
              <td>${safeDutyData.Row18_YesNo || ''}</td>
              <td>${safeDutyData.Row18_Done || ''}</td>
              <td>${safeDutyData.Row18_Checked || ''}</td>
              <td>${safeDutyData.Row18_Time || ''}</td>
            </tr>
            <!-- Row 19 -->
            <tr>
              <td>19</td>
              <td>How many tanks are empty & need to be replaced?</td>
              <td>
                O₂: ${safeDutyData.Row19_O2 || '0'}<br>
                N₂O: ${safeDutyData.Row19_N2O || '0'}<br>
                He: ${safeDutyData.Row19_He || '0'}
              </td>
              <td>${safeDutyData.Row19_Done || ''}</td>
              <td>${safeDutyData.Row19_Checked || ''}</td>
              <td>${safeDutyData.Row19_Time || ''}</td>
            </tr>
            <!-- Row 20 -->
            <tr>
              <td>20</td>
              <td>Check restrooms initial logs hourly</td>
              <td>-</td>
              <td>${safeDutyData.Row20_Done || ''}</td>
              <td>${safeDutyData.Row20_Checked || ''}</td>
              <td>${safeDutyData.Row20_Time || ''}</td>
            </tr>
            <!-- Row 21 -->
            <tr>
              <td>21</td>
              <td>Swept/Mopped</td>
              <td>${safeDutyData.Row21_YesNo || ''}</td>
              <td>${safeDutyData.Row21_Done || ''}</td>
              <td>${safeDutyData.Row21_Checked || ''}</td>
              <td>${safeDutyData.Row21_Time || ''}</td>
            </tr>
            <!-- Row 22 -->
            <tr>
              <td>22</td>
              <td>Cleaned Breakroom</td>
              <td>${safeDutyData.Row22_YesNo || ''}</td>
              <td>${safeDutyData.Row22_Done || ''}</td>
              <td>${safeDutyData.Row22_Checked || ''}</td>
              <td>${safeDutyData.Row22_Time || ''}</td>
            </tr>
            <!-- Row 23 -->
            <tr>
              <td>23</td>
              <td>Sterilizers: Cycle Complete <br><span class="duty-details">(Do Not Push Stop)</span></td>
              <td>${safeDutyData.Row23_YesNo || ''}</td>
              <td>${safeDutyData.Row23_Done || ''}</td>
              <td>${safeDutyData.Row23_Checked || ''}</td>
              <td>${safeDutyData.Row23_Time || ''}</td>
            </tr>
            <!-- Row 24 -->
            <tr>
              <td>24</td>
              <td>Drained Ultrasonic</td>
              <td>${safeDutyData.Row24_YesNo || ''}</td>
              <td>${safeDutyData.Row24_Done || ''}</td>
              <td>${safeDutyData.Row24_Checked || ''}</td>
              <td>${safeDutyData.Row24_Time || ''}</td>
            </tr>
            <!-- Row 25 -->
            <tr>
              <td>25</td>
              <td>Spore Test <br><span class="duty-details">Every Monday</span></td>
              <td>-</td>
              <td>${safeDutyData.Row25_Done || ''}</td>
              <td>${safeDutyData.Row25_Checked || ''}</td>
              <td>${safeDutyData.Row25_Time || ''}</td>
            </tr>
            <!-- Row 26 -->
            <tr>
              <td>26</td>
              <td>Turn Off All TV's and Computers at the End of the Day</td>
              <td>-</td>
              <td>${safeDutyData.Row26_Done || ''}</td>
              <td>${safeDutyData.Row26_Checked || ''}</td>
              <td>${safeDutyData.Row26_Time || ''}</td>
            </tr>
            <!-- Row 27 -->
            <tr>
              <td>27</td>
              <td>Postcards Ready for Pick-up</td>
              <td>-</td>
              <td>${safeDutyData.Row27_Done || ''}</td>
              <td>${safeDutyData.Row27_Checked || ''}</td>
              <td>${safeDutyData.Row27_Time || ''}</td>
            </tr>
            <!-- Row 28 -->
            <tr>
              <td>28</td>
              <td>Clean traps everyday <br><span class="duty-details">(chair)</span></td>
              <td>-</td>
              <td>${safeDutyData.Row28_Done || ''}</td>
              <td>${safeDutyData.Row28_Checked || ''}</td>
              <td>${safeDutyData.Row28_Time || ''}</td>
            </tr>
            <!-- Row 29 -->
            <tr>
              <td>29</td>
              <td>Clean main trap 1st/15th <br><span class="duty-details">(by vacuum)</span></td>
              <td>-</td>
              <td>${safeDutyData.Row29_Done || ''}</td>
              <td>${safeDutyData.Row29_Checked || ''}</td>
              <td>${safeDutyData.Row29_Time || ''}</td>
            </tr>
            <!-- Row 30 -->
            <tr>
              <td>30</td>
              <td>Did you flush the lines with hot water?</td>
              <td>${safeDutyData.Row30_YesNo || ''}</td>
              <td>${safeDutyData.Row30_Done || ''}</td>
              <td>${safeDutyData.Row30_Checked || ''}</td>
              <td>${safeDutyData.Row30_Time || ''}</td>
            </tr>
            <!-- Row 31 -->
            <tr>
              <td>31</td>
              <td>Check all doors are locked</td>
              <td>-</td>
              <td>${safeDutyData.Row31_Done || ''}</td>
              <td>${safeDutyData.Row31_Checked || ''}</td>
              <td>${safeDutyData.Row31_Time || ''}</td>
            </tr>
            <!-- Row 32 -->
            <tr>
              <td>32</td>
              <td>Turn On Answering Service</td>
              <td>-</td>
              <td>${safeDutyData.Row32_Done || ''}</td>
              <td>${safeDutyData.Row32_Checked || ''}</td>
              <td>${safeDutyData.Row32_Time || ''}</td>
            </tr>
          </tbody>
        </table>
      </body>
      </html>
    `;
    
    // Puppeteer로 PDF 생성 (안정적인 설정)
    try {
      browser = await puppeteer.launch({
        headless: true,
        args: [
          '--no-sandbox',
          '--disable-setuid-sandbox',
          '--disable-dev-shm-usage', // 메모리 부족 문제 해결
          '--disable-accelerated-2d-canvas',
          '--no-first-run',
          '--no-zygote',
          '--disable-gpu'
        ],
        timeout: 30000, // 30초 타임아웃
      });
      
      const page = await browser.newPage();
      
      // 페이지 타임아웃 설정
      page.setDefaultTimeout(30000);
      
      // 🔒 보안: CSP는 유지하되, 스타일시트는 허용
      await page.setContent(html, {
        waitUntil: 'load', // load로 변경 (networkidle0은 외부 리소스 대기)
        timeout: 30000
      });
      
      // 추가 대기 (CSS 및 스타일 렌더링 완료)
      await new Promise(resolve => setTimeout(resolve, 2000));
      
      // 테이블이 제대로 렌더링되었는지 확인 (디버깅용)
      const tableExists = await page.evaluate(() => {
        return document.querySelector('table') !== null;
      });
      
      if (!tableExists) {
        console.warn('Warning: Table element not found in HTML');
      }
      
      // PDF 생성 (한 페이지에 맞추기)
      const pdfBuffer = await page.pdf({
        format: 'A4',
        printBackground: false, // 흑백 출력
        margin: {
          top: '8mm',
          right: '8mm',
          bottom: '8mm',
          left: '8mm'
        },
        preferCSSPageSize: false,
        displayHeaderFooter: false,
        timeout: 30000,
        // 한 페이지에 맞추기 (조금 더 크게)
        scale: 0.85
      });
      
      // 🔒 보안: 파일명 이스케이프 (파일명 주입 공격 방지)
      const safeFilename = escapeHtml(`${dutyDate}_${selectedOffice || 'unknown'}_Daily Office Duty.pdf`);
      const filename = `2) ${safeFilename}`;
      
      // Buffer를 Uint8Array로 변환하여 NextResponse에 전달
      return new NextResponse(new Uint8Array(pdfBuffer), {
        headers: {
          'Content-Type': 'application/pdf',
          'Content-Disposition': `attachment; filename="${filename.replace(/"/g, '\\"')}"`,
          // 🔒 보안: 추가 보안 헤더
          'X-Content-Type-Options': 'nosniff',
          'X-Frame-Options': 'DENY',
          'X-XSS-Protection': '1; mode=block'
        }
      });
      
    } catch (puppeteerError) {
      console.error('Puppeteer PDF generation error:', puppeteerError);
      // 🔒 보안: 폴백 제거 - PDF 생성 실패 시 에러만 반환 (HTML 반환하지 않음)
      return NextResponse.json({ 
        success: false, 
        error: 'PDF generation failed. Please try again later.' 
      }, { status: 500 });
    } finally {
      // 브라우저 정리
      if (browser) {
        await browser.close().catch(err => {
          console.error('Error closing browser:', err);
        });
      }
    }

  } catch (error) {
    console.error('Daily office duty PDF generation error:', error);
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    
    // 브라우저가 열려있으면 정리
    if (browser) {
      await browser.close().catch(err => {
        console.error('Error closing browser in catch block:', err);
      });
    }
    
    return NextResponse.json({ 
      success: false, 
      error: 'Failed to generate PDF: ' + errorMessage
    }, { status: 500 });
  }
}
