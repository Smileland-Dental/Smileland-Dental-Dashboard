import { NextRequest, NextResponse } from 'next/server';
import { 
  escapeHtml, 
  sanitizeString, 
  safeLog, 
  logError, 
  validatePdfGeneration, 
  sanitizeArrayForPdf,
  validateRequest,
  getSecurityHeaders,
  validateAndParseRequestBody,
  verifyFirebaseAuth,
  verifyDataOwnership,
  enforceUserData
} from '@/lib/security-server';

export async function POST(request: NextRequest) {
  // Production에서는 로그를 표시하지 않음
  if (process.env.NODE_ENV !== 'production') {
    safeLog('✅ PDF API POST 요청 받음!');
  }
  
  try {
    // 1. Firebase Authentication 검증 (선택적 - 인프라 레벨에서 이미 제어됨)
    const authHeader = request.headers.get('authorization');
    let authenticatedUser: { uid: string; email?: string } | null = null;
    
    // 토큰이 있으면 검증 시도 (없어도 진행 가능 - 인프라 레벨에서 이미 제어됨)
    if (authHeader && authHeader.startsWith('Bearer ')) {
      authenticatedUser = await verifyFirebaseAuth(request);
      if (authenticatedUser) {
        if (process.env.NODE_ENV !== 'production') {
          safeLog('✅ 사용자 인증 성공:', { uid: authenticatedUser.uid, email: authenticatedUser.email });
        }
      } else {
        // 토큰 검증 실패해도 경고만 (인프라 레벨에서 이미 제어되므로)
        // Production에서는 로그를 표시하지 않음
        if (process.env.NODE_ENV !== 'production') {
          safeLog('ℹ️ 토큰 검증 실패 (인프라 레벨 제어에 의존)');
        }
        // 에러를 던지지 않고 계속 진행 (인프라 레벨에서 이미 접근 제어됨)
      }
    }
    
    // REQUIRE_AUTH 환경 변수가 true인 경우에만 강제 인증
    if (process.env.REQUIRE_AUTH === 'true' && !authenticatedUser) {
      logError(new Error('Authentication required but not provided'), 'generate-pdf');
      return NextResponse.json(
        { 
          success: false, 
          error: '인증이 필요합니다. 로그인 후 다시 시도해주세요.' 
        },
        { 
          status: 401,
          headers: getSecurityHeaders()
        }
      );
    }
    
    // 2. 종합 보안 검증
    const validation = await validateRequest(request, {
      requireTimestamp: true,
      requireSignature: false, // 선택적 (필요시 활성화)
      maxRequests: 50, // PDF 생성은 리소스 집약적이므로 더 낮은 제한
      windowMs: 60 * 1000 // 1분당 50회
    });
    
    if (!validation.valid || !validation.body) {
      logError(new Error(validation.error || 'Request validation failed'), 'generate-pdf');
      return NextResponse.json(
        { 
          success: false, 
          error: '보안 검증 실패: ' + (validation.error || 'Invalid request') 
        },
        { 
          status: 403,
          headers: getSecurityHeaders()
        }
      );
    }
    
    // validateRequest에서 이미 파싱한 본문 사용
    const body = validation.body;
    const { patientData } = body.data || body; // 보안 래퍼에서 데이터 추출
    
    // 인증된 사용자가 있으면 데이터 소유권 검증 및 강제 적용
    if (authenticatedUser && patientData) {
      // 데이터 소유권 검증
      if (!verifyDataOwnership(patientData, authenticatedUser.uid)) {
        logError(new Error('Data ownership verification failed'), 'generate-pdf');
        return NextResponse.json(
          { 
            success: false, 
            error: '데이터 소유권 검증 실패. 권한이 없습니다.' 
          },
          { 
            status: 403,
            headers: getSecurityHeaders()
          }
        );
      }
      
      // 사용자 정보 강제 적용 (보안 강화)
      patientData.userId = authenticatedUser.uid;
      if (authenticatedUser.email) {
        patientData.userEmail = authenticatedUser.email;
      }
    }
    
    if (process.env.NODE_ENV !== 'production') {
      safeLog('📋 받은 데이터:', { hasData: !!patientData });
    }
    
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
          .text-wrap-cell {
            word-wrap: break-word;
            word-break: break-word;
            white-space: normal;
            overflow-wrap: break-word;
            max-width: 100%;
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
              overflow: visible;
              white-space: normal;
            }
            th {
              font-size: 7px;
              padding: 4px 2px;
              height: 16px;
            }
            tr {
              height: auto;
              min-height: 18px;
            }
            .text-wrap-cell {
              word-wrap: break-word;
              word-break: break-word;
              white-space: normal;
              overflow-wrap: break-word;
              max-width: 100%;
              padding: 3px 2px;
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
          <h1>Patient Log</h1>
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
                  <td>${escapeHtml(safeName) || '-'}</td>
                  <td>${escapeHtml(safeOffice) || '-'}</td>
                  <td style="text-align: center;">${escapeHtml(safeApptDate) || '-'}</td>
                  <td class="text-wrap-cell">${escapeHtml(safeVisitType) || '-'}</td>
                  <td class="checkbox-cell">${callIn ? '✓' : ''}</td>
                  <td class="checkbox-cell">${callOut ? '✓' : ''}</td>
                  <td style="text-align: center;">${escapeHtml(convertTo12Hour(safeTime))}</td>
                  <td class="text-wrap-cell">${escapeHtml(safeRemark) || '-'}</td>
                  <td class="text-wrap-cell">${escapeHtml(safeOtherDuty) || '-'}</td>
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

    safeLog('🎯 PDF HTML 생성 완료');
    
    // HTML만 반환 (cc_route.ts와 동일한 방식)
    return NextResponse.json({
      success: true,
      html: htmlContent,
      message: 'PDF HTML generated successfully'
    });

  } catch (error) {
    logError(error, 'generate-pdf');
    return NextResponse.json(
      { 
        success: false, 
        error: 'PDF 생성 중 오류가 발생했습니다.' 
      },
      { 
        status: 500,
        headers: getSecurityHeaders()
      }
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
