import { NextRequest, NextResponse } from 'next/server';
import puppeteer from 'puppeteer';

export async function POST(request: NextRequest) {
  try {
    const { dutyDate, selectedOffice, dutyData } = await request.json();

    if (!dutyDate || !dutyData) {
      return NextResponse.json({ success: false, error: 'Duty date and data are required' });
    }

    // HTML 템플릿 생성
    const html = `
      <!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>${dutyDate}_${selectedOffice || 'Ming'}_Daily_Office_Duty</title>
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
            text-align: left;
            font-size: 8px;
            vertical-align: top;
          }
          
          th {
            background: #f0f0f0;
            font-weight: bold;
          }
          
          .duty-details {
            font-size: 7px;
            color: #555;
            margin-top: 2px;
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
          <h1>Daily Office Duty Report</h1>
        </div>
        
        <div class="info-section">
          <div><strong>Date:</strong> ${dutyDate}</div>
          <div><strong>Location:</strong> ${selectedOffice || 'Ming'}</div>
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
              <td>${dutyData.Row1_Done || ''}</td>
              <td>${dutyData.Row1_Checked || ''}</td>
              <td>${dutyData.Row1_Time || ''}</td>
            </tr>
            <!-- Row 2 -->
            <tr>
              <td>2</td>
              <td>All charts filed back?</td>
              <td>${dutyData.Row2_YesNo || ''}</td>
              <td>${dutyData.Row2_Done || ''}</td>
              <td>${dutyData.Row2_Checked || ''}</td>
              <td>${dutyData.Row2_Time || ''}</td>
            </tr>
            <!-- Row 3 -->
            <tr>
              <td>3</td>
              <td>Charts pulled for next day</td>
              <td>-</td>
              <td>${dutyData.Row3_Done || ''}</td>
              <td>${dutyData.Row3_Checked || ''}</td>
              <td>${dutyData.Row3_Time || ''}</td>
            </tr>
            <!-- Row 4 -->
            <tr>
              <td>4</td>
              <td>Check eligibility <br><span class="duty-details">1st of every month come in early to check eligibility by 8:30 am</span></td>
              <td>-</td>
              <td>${dutyData.Row4_Done || ''}</td>
              <td>${dutyData.Row4_Checked || ''}</td>
              <td>${dutyData.Row4_Time || ''}</td>
            </tr>
            <!-- Row 5 -->
            <tr>
              <td>5</td>
              <td>If pt is not eligible call and inform</td>
              <td>How many pt's did you call: ${dutyData.Row5_CallNum || ''}</td>
              <td>${dutyData.Row5_Done || ''}</td>
              <td>${dutyData.Row5_Checked || ''}</td>
              <td>${dutyData.Row5_Time || ''}</td>
            </tr>
            <!-- Row 6 -->
            <tr>
              <td>6</td>
              <td>Insurance breakdown for next day's patients <br><span class="duty-details">Call and get ins. info if necessary</span></td>
              <td>-</td>
              <td>${dutyData.Row6_Done || ''}</td>
              <td>${dutyData.Row6_Checked || ''}</td>
              <td>${dutyData.Row6_Time || ''}</td>
            </tr>
            <!-- Row 7 -->
            <tr>
              <td>7</td>
              <td>Check ledger for any balance on the account <br><span class="duty-details">Fill out "Account with Balances Form" and fax to the AR Department at (661)328-1905</span><br><span class="duty-details">Called to inform patient of balance?</span></td>
              <td>${dutyData.Row7_YesNo || ''}</td>
              <td>${dutyData.Row7_Done || ''}</td>
              <td>${dutyData.Row7_Checked || ''}</td>
              <td>${dutyData.Row7_Time || ''}</td>
            </tr>
            <!-- Row 8 -->
            <tr>
              <td>8</td>
              <td>Morning confirmations <br><span class="duty-details">At least by noon</span></td>
              <td>-</td>
              <td>${dutyData.Row8_Done || ''}</td>
              <td>${dutyData.Row8_Checked || ''}</td>
              <td>${dutyData.Row8_Time || ''}</td>
            </tr>
            <!-- Row 9 -->
            <tr>
              <td>9</td>
              <td>No shows entered on ledger</td>
              <td>-</td>
              <td>${dutyData.Row9_Done || ''}</td>
              <td>${dutyData.Row9_Checked || ''}</td>
              <td>${dutyData.Row9_Time || ''}</td>
            </tr>
            <!-- Row 10 -->
            <tr>
              <td>10</td>
              <td>No shows stamped in patient charts</td>
              <td>-</td>
              <td>${dutyData.Row10_Done || ''}</td>
              <td>${dutyData.Row10_Checked || ''}</td>
              <td>${dutyData.Row10_Time || ''}</td>
            </tr>
            <!-- Row 11 -->
            <tr>
              <td>11</td>
              <td>Reconfirming completed? <br><span class="duty-details">Start at 4:00pm</span></td>
              <td>-</td>
              <td>${dutyData.Row11_Done || ''}</td>
              <td>${dutyData.Row11_Checked || ''}</td>
              <td>${dutyData.Row11_Time || ''}</td>
            </tr>
            <!-- Row 12 -->
            <tr>
              <td>12</td>
              <td>One week reminders completed?</td>
              <td>-</td>
              <td>${dutyData.Row12_Done || ''}</td>
              <td>${dutyData.Row12_Checked || ''}</td>
              <td>${dutyData.Row12_Time || ''}</td>
            </tr>
            <!-- Row 13 -->
            <tr>
              <td>13</td>
              <td>Call all treatment patients from today for post op</td>
              <td>-</td>
              <td>${dutyData.Row13_Done || ''}</td>
              <td>${dutyData.Row13_Checked || ''}</td>
              <td>${dutyData.Row13_Time || ''}</td>
            </tr>
            <!-- Row 14 -->
            <tr>
              <td>14</td>
              <td>Total lab case deposits/deliveries <br><span class="duty-details">Name/DOB</span></td>
              <td>${(dutyData['Row14_Name/DOB'] || '').replace(/\n/g, '<br>')}</td>
              <td>${dutyData.Row14_Done || ''}</td>
              <td>${dutyData.Row14_Checked || ''}</td>
              <td>${dutyData.Row14_Time || ''}</td>
            </tr>
            <!-- Row 15 -->
            <tr>
              <td>15</td>
              <td>Check all undelivered lab cases and make appointments <br><span class="duty-details">Any Lab case that is more than 3 weeks old must be sent to corporate along with $20 deposit</span></td>
              <td>${(dutyData.Row15_LabCases || '').replace(/\n/g, '<br>')}</td>
              <td>${dutyData.Row15_Done || ''}</td>
              <td>${dutyData.Row15_Checked || ''}</td>
              <td>${dutyData.Row15_Time || ''}</td>
            </tr>
            <!-- Row 16 -->
            <tr>
              <td>16</td>
              <td>Check all lab cases for next day <br><span class="duty-details">Call lab for next day pick up's</span></td>
              <td>-</td>
              <td>${dutyData.Row16_Done || ''}</td>
              <td>${dutyData.Row16_Checked || ''}</td>
              <td>${dutyData.Row16_Time || ''}</td>
            </tr>
            <!-- Row 17 -->
            <tr>
              <td>17</td>
              <td>N₂O/ Compressor Off</td>
              <td>-</td>
              <td>${dutyData.Row17_Done || ''}</td>
              <td>${dutyData.Row17_Checked || ''}</td>
              <td>${dutyData.Row17_Time || ''}</td>
            </tr>
            <!-- Row 18 -->
            <tr>
              <td>18</td>
              <td>Did you read the meter on the Oxygen/N₂O/Helium tank?</td>
              <td>${dutyData.Row18_YesNo || ''}</td>
              <td>${dutyData.Row18_Done || ''}</td>
              <td>${dutyData.Row18_Checked || ''}</td>
              <td>${dutyData.Row18_Time || ''}</td>
            </tr>
            <!-- Row 19 -->
            <tr>
              <td>19</td>
              <td>How many tanks are empty & need to be replaced?</td>
              <td>
                O₂: ${dutyData.Row19_O2 || '0'}<br>
                N₂O: ${dutyData.Row19_N2O || '0'}<br>
                He: ${dutyData.Row19_He || '0'}
              </td>
              <td>${dutyData.Row19_Done || ''}</td>
              <td>${dutyData.Row19_Checked || ''}</td>
              <td>${dutyData.Row19_Time || ''}</td>
            </tr>
            <!-- Row 20 -->
            <tr>
              <td>20</td>
              <td>Check restrooms initial logs hourly</td>
              <td>-</td>
              <td>${dutyData.Row20_Done || ''}</td>
              <td>${dutyData.Row20_Checked || ''}</td>
              <td>${dutyData.Row20_Time || ''}</td>
            </tr>
            <!-- Row 21 -->
            <tr>
              <td>21</td>
              <td>Swept/Mopped</td>
              <td>${dutyData.Row21_YesNo || ''}</td>
              <td>${dutyData.Row21_Done || ''}</td>
              <td>${dutyData.Row21_Checked || ''}</td>
              <td>${dutyData.Row21_Time || ''}</td>
            </tr>
            <!-- Row 22 -->
            <tr>
              <td>22</td>
              <td>Cleaned Breakroom</td>
              <td>${dutyData.Row22_YesNo || ''}</td>
              <td>${dutyData.Row22_Done || ''}</td>
              <td>${dutyData.Row22_Checked || ''}</td>
              <td>${dutyData.Row22_Time || ''}</td>
            </tr>
            <!-- Row 23 -->
            <tr>
              <td>23</td>
              <td>Sterilizers: Cycle Complete <br><span class="duty-details">(Do Not Push Stop)</span></td>
              <td>${dutyData.Row23_YesNo || ''}</td>
              <td>${dutyData.Row23_Done || ''}</td>
              <td>${dutyData.Row23_Checked || ''}</td>
              <td>${dutyData.Row23_Time || ''}</td>
            </tr>
            <!-- Row 24 -->
            <tr>
              <td>24</td>
              <td>Drained Ultrasonic</td>
              <td>${dutyData.Row24_YesNo || ''}</td>
              <td>${dutyData.Row24_Done || ''}</td>
              <td>${dutyData.Row24_Checked || ''}</td>
              <td>${dutyData.Row24_Time || ''}</td>
            </tr>
            <!-- Row 25 -->
            <tr>
              <td>25</td>
              <td>Spore Test <br><span class="duty-details">Every Monday</span></td>
              <td>-</td>
              <td>${dutyData.Row25_Done || ''}</td>
              <td>${dutyData.Row25_Checked || ''}</td>
              <td>${dutyData.Row25_Time || ''}</td>
            </tr>
            <!-- Row 26 -->
            <tr>
              <td>26</td>
              <td>Turn Off All TV's and Computers at the End of the Day</td>
              <td>-</td>
              <td>${dutyData.Row26_Done || ''}</td>
              <td>${dutyData.Row26_Checked || ''}</td>
              <td>${dutyData.Row26_Time || ''}</td>
            </tr>
            <!-- Row 27 -->
            <tr>
              <td>27</td>
              <td>Postcards Ready for Pick-up</td>
              <td>-</td>
              <td>${dutyData.Row27_Done || ''}</td>
              <td>${dutyData.Row27_Checked || ''}</td>
              <td>${dutyData.Row27_Time || ''}</td>
            </tr>
            <!-- Row 28 -->
            <tr>
              <td>28</td>
              <td>Clean traps everyday <br><span class="duty-details">(chair)</span></td>
              <td>-</td>
              <td>${dutyData.Row28_Done || ''}</td>
              <td>${dutyData.Row28_Checked || ''}</td>
              <td>${dutyData.Row28_Time || ''}</td>
            </tr>
            <!-- Row 29 -->
            <tr>
              <td>29</td>
              <td>Clean main trap 1st/15th <br><span class="duty-details">(by vacuum)</span></td>
              <td>-</td>
              <td>${dutyData.Row29_Done || ''}</td>
              <td>${dutyData.Row29_Checked || ''}</td>
              <td>${dutyData.Row29_Time || ''}</td>
            </tr>
            <!-- Row 30 -->
            <tr>
              <td>30</td>
              <td>Did you flush the lines with hot water?</td>
              <td>${dutyData.Row30_YesNo || ''}</td>
              <td>${dutyData.Row30_Done || ''}</td>
              <td>${dutyData.Row30_Checked || ''}</td>
              <td>${dutyData.Row30_Time || ''}</td>
            </tr>
            <!-- Row 31 -->
            <tr>
              <td>31</td>
              <td>Check all doors are locked</td>
              <td>-</td>
              <td>${dutyData.Row31_Done || ''}</td>
              <td>${dutyData.Row31_Checked || ''}</td>
              <td>${dutyData.Row31_Time || ''}</td>
            </tr>
            <!-- Row 32 -->
            <tr>
              <td>32</td>
              <td>Turn On Answering Service</td>
              <td>-</td>
              <td>${dutyData.Row32_Done || ''}</td>
              <td>${dutyData.Row32_Checked || ''}</td>
              <td>${dutyData.Row32_Time || ''}</td>
            </tr>
          </tbody>
        </table>
        
        <div class="footer">
          <p>Smileland Dental</p>
          <p>Corporate Office | Phone 661-328-0876 | Fax 661-327-4733</p>
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
    const filename = `${dutyDate}_${selectedOffice || 'Ming'}_Daily_Office_Duty.pdf`;
    
    return new NextResponse(pdf, {
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${filename}"`
      }
    });

  } catch (error) {
    console.error('Daily office duty PDF generation error:', error);
    return NextResponse.json({ 
      success: false, 
      error: 'Failed to generate daily office duty PDF' 
    });
  }
}
