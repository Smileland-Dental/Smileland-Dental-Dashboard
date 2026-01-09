import { NextRequest, NextResponse } from 'next/server';
import puppeteer from 'puppeteer';
import { Buffer } from 'buffer';
import { existsSync } from 'fs';
import { join } from 'path';
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
    
    // 디버깅을 위한 로깅 (개발 환경)
    if (process.env.NODE_ENV !== 'production') {
      safeLog('📦 Request body structure:', {
        hasBody: !!body,
        bodyKeys: body ? Object.keys(body) : [],
        hasData: !!(body?.data),
        dataKeys: body?.data ? Object.keys(body.data) : [],
        hasPatientData: !!(body?.patientData),
        hasDataPatientData: !!(body?.data?.patientData)
      });
    }
    
    // 보안 래퍼에서 데이터 추출: 
    // 1. body.data.patientData (보안 래퍼 안에 patientData가 있는 경우)
    // 2. body.patientData (직접 patientData가 있는 경우)
    // 3. body.data (보안 래퍼 자체가 patientData인 경우 - 이건 제외)
    let patientData: any = body.data?.patientData || body.patientData;
    
    // patientData를 찾지 못한 경우, body.data가 patientData 객체인지 확인
    if (!patientData && body.data && typeof body.data === 'object' && !Array.isArray(body.data)) {
      // body.data가 patientData 객체인지 확인 (dutyDate, userName 등의 필드가 있는지)
      if (body.data.dutyDate || body.data.userName || body.data.workOffice) {
        patientData = body.data;
      }
    }
    
    // 여전히 찾지 못한 경우, body 자체가 patientData인지 확인
    if (!patientData && body && typeof body === 'object' && !Array.isArray(body)) {
      if (body.dutyDate || body.userName || body.workOffice) {
        patientData = body;
      }
    }
    
    // patientData가 없으면 에러 반환
    if (!patientData) {
      logError(new Error('patientData is missing from request'), 'generate-pdf');
      if (process.env.NODE_ENV !== 'production') {
        safeLog('❌ Request body structure:', JSON.stringify(body, null, 2));
      }
      return NextResponse.json(
        { 
          success: false, 
          error: '요청 데이터가 올바르지 않습니다. patientData가 필요합니다.' 
        },
        { 
          status: 400,
          headers: getSecurityHeaders()
        }
      );
    }
    
    // patientData가 객체가 아닌 경우 처리
    if (typeof patientData !== 'object' || Array.isArray(patientData)) {
      logError(new Error('patientData is not an object'), 'generate-pdf');
      if (process.env.NODE_ENV !== 'production') {
        safeLog('❌ patientData type:', typeof patientData);
        safeLog('❌ patientData value:', JSON.stringify(patientData));
      }
      return NextResponse.json(
        { 
          success: false, 
          error: '요청 데이터 형식이 올바르지 않습니다.' 
        },
        { 
          status: 400,
          headers: getSecurityHeaders()
        }
      );
    }
    
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
      safeLog('📋 받은 데이터:', { 
        hasData: !!patientData,
        hasPatientLogs: !!(patientData.patientLogs),
        hasPatientRows: !!(patientData.patientRows),
        keys: Object.keys(patientData || {})
      });
    }
    
    // PDF 생성 데이터 크기 검증
    const dataSize = JSON.stringify(patientData).length;
    if (process.env.NODE_ENV !== 'production') {
      safeLog('📊 데이터 크기:', `${dataSize} bytes`);
    }
    if (!validatePdfGeneration(dataSize)) {
      throw new Error(`데이터 크기가 너무 큽니다: ${dataSize} bytes (최대 허용 크기 초과)`);
    }
    
    // patientLogs 또는 patientRows 둘 다 처리하고 빈 행 필터링 (보안 강화)
    const allPatients = patientData.patientLogs || patientData.patientRows || [];
    
    if (process.env.NODE_ENV !== 'production') {
      safeLog('👥 환자 데이터:', {
        totalPatients: allPatients.length,
        hasPatientLogs: !!(patientData.patientLogs),
        hasPatientRows: !!(patientData.patientRows)
      });
    }
    
    const safePatients = sanitizeArrayForPdf(allPatients, 1000); // 최대 1000개 행으로 제한
    const patientList = safePatients.filter((row: any) => 
      row && (
        row.name || row.office || row.appt_date || row.apptDate || row.visit_type || row.visitType || 
        row.call_in || row.callIn || row.call_out || row.callOut || row.time || row.remark || 
        row.other_duty || row.otherDuty
      )
    );
    
    if (process.env.NODE_ENV !== 'production') {
      safeLog('✅ 필터링된 환자 데이터:', `${patientList.length} 행`);
    }
    
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

    // 통계 계산 (HTML 템플릿 리터럴 안에서 실행하지 않고 미리 계산)
    const allPatientsForStats = patientData.patientLogs || patientData.patientRows || [];
    const totalAppointments = allPatientsForStats.filter((row: any) => row.appt_date && row.name).length;
    const incomingCalls = allPatientsForStats.filter((row: any) => row.call_in).length;
    const outgoingCalls = allPatientsForStats.filter((row: any) => row.call_out).length;
    
    // 날짜 포맷팅 미리 계산
    const generatedDate = new Date().toLocaleDateString('en-US', { 
      year: 'numeric', 
      month: 'long', 
      day: 'numeric',
      weekday: 'long',
      hour: '2-digit',
      minute: '2-digit',
      hour12: true
    });

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
            <div class="count-number">${totalAppointments}</div>
          </div>
          <div class="count-item">
            <div class="count-label">Incoming Calls</div>
            <div class="count-number">${incomingCalls}</div>
          </div>
          <div class="count-item">
            <div class="count-label">Outgoing Calls</div>
            <div class="count-number">${outgoingCalls}</div>
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
          <div class="date">Generated: ${generatedDate}</div>
        </div>
      </body>
      </html>
    `;

    safeLog('🎯 PDF HTML 생성 완료');
    
    // 파일명 생성 (안전하게)
    const filenameBase = `${safeDutyDate}_${safeUserName}_${safeWorkOffice}_Patient_Log`;
    const filename = filenameBase.replace(/[^a-zA-Z0-9_-]/g, '_') + '.pdf';
    
    // Puppeteer로 PDF 생성 (보안 옵션 강화)
    let browser;
    let pdf: Buffer | Uint8Array;
    
    // 브라우저 실행 파일 경로 찾기 (Edge 우선, 없으면 Chromium)
    const findBrowserPath = (): string | undefined => {
      // 환경 변수에서 브라우저 경로 확인
      if (process.env.BROWSER_PATH && existsSync(process.env.BROWSER_PATH)) {
        return process.env.BROWSER_PATH;
      }
      
      // Windows에서 Microsoft Edge 경로 확인
      const edgePaths = [
        'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
        'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
        process.env.PROGRAMFILES ? join(process.env.PROGRAMFILES, 'Microsoft', 'Edge', 'Application', 'msedge.exe') : '',
        process.env['PROGRAMFILES(X86)'] ? join(process.env['PROGRAMFILES(X86)'], 'Microsoft', 'Edge', 'Application', 'msedge.exe') : ''
      ].filter(path => path && existsSync(path));
      
      if (edgePaths.length > 0) {
        if (process.env.NODE_ENV !== 'production') {
          safeLog('✅ Microsoft Edge 경로 발견:', edgePaths[0]);
        }
        return edgePaths[0];
      }
      
      // Edge를 찾지 못한 경우 Puppeteer의 Chromium 사용
      try {
        const chromiumPath = puppeteer.executablePath();
        if (process.env.NODE_ENV !== 'production') {
          safeLog('✅ Puppeteer Chromium 경로:', chromiumPath);
        }
        return chromiumPath;
      } catch (error) {
        if (process.env.NODE_ENV !== 'production') {
          safeLog('⚠️ Chromium 경로를 찾을 수 없습니다.');
        }
        return undefined;
      }
    };
    
    try {
      if (process.env.NODE_ENV !== 'production') {
        safeLog('🚀 Puppeteer 시작 중...');
      }
      
      const browserPath = findBrowserPath();
      const launchOptions: any = {
        headless: true,
        args: [
          // 보안: --no-sandbox는 보안상 위험하므로 환경 변수로 제어
          // Docker나 특정 서버 환경에서만 필요할 수 있음
          ...(process.env.ALLOW_NO_SANDBOX === 'true' ? ['--no-sandbox'] : []),
          '--disable-setuid-sandbox',
          '--disable-dev-shm-usage', // 메모리 문제 방지
          '--disable-gpu', // GPU 비활성화 (서버 환경)
          // 보안: 추가 보안 플래그
          '--disable-extensions', // 확장 프로그램 비활성화
          '--disable-plugins', // 플러그인 비활성화
          '--disable-background-networking', // 백그라운드 네트워크 비활성화
          '--no-first-run', // 첫 실행 설정 건너뛰기
          '--no-default-browser-check' // 기본 브라우저 확인 건너뛰기
        ],
        timeout: 30000
      };
      
      // 브라우저 경로가 있으면 사용
      if (browserPath) {
        launchOptions.executablePath = browserPath;
      }
      
      browser = await puppeteer.launch(launchOptions);
      
      if (process.env.NODE_ENV !== 'production') {
        safeLog('📄 새 페이지 생성 중...');
      }
      
      const page = await browser.newPage();
      
      // 외부 리소스 로드 차단 (보안)
      await page.setRequestInterception(true);
      page.on('request', (request) => {
        const url = request.url();
        // 로컬 리소스만 허용
        if (url.startsWith('data:') || url.startsWith('blob:') || url === 'about:blank') {
          request.continue();
        } else {
          request.abort();
        }
      });
      
      page.setDefaultTimeout(10000);
      
      // HTML 콘텐츠 설정
      await page.setContent(htmlContent, { 
        waitUntil: 'domcontentloaded',
        timeout: 10000
      });
      
      // 페이지 렌더링 대기
      await new Promise(resolve => setTimeout(resolve, 500));
      
      // PDF 생성
      pdf = await page.pdf({
        format: 'A4',
        printBackground: true,
        margin: {
          top: '10mm',
          right: '10mm',
          bottom: '10mm',
          left: '10mm'
        }
      });
      
    } catch (puppeteerError: any) {
      const errorMessage = puppeteerError?.message || 'Unknown error';
      
      // Chromium 경로를 찾지 못한 경우
      if (errorMessage.includes('Could not find Chrome') || 
          errorMessage.includes('did not perform an installation') ||
          errorMessage.includes('executable doesn\'t exist')) {
        if (process.env.NODE_ENV !== 'production') {
          safeLog('❌ Chromium이 설치되지 않았습니다. 다음 명령어를 실행하세요:');
          safeLog('   npx puppeteer browsers install chrome');
        }
        throw new Error(`PDF 생성 실패: Chromium이 설치되지 않았습니다. 'npx puppeteer browsers install chrome' 명령어를 실행해주세요.`);
      }
      
      if (process.env.NODE_ENV !== 'production') {
        safeLog('❌ Puppeteer 에러:', errorMessage);
      }
      throw new Error(`PDF 생성 실패: ${errorMessage}`);
    } finally {
      // 브라우저 종료
      if (browser) {
        try {
          await browser.close();
        } catch (closeError) {
          // 종료 오류는 무시
        }
      }
    }
    
    // 파일명 생성
    const pdfFilename = filename;
    
    // PDF 응답 생성 (보안 헤더 추가)
    // Buffer 또는 Uint8Array를 NextResponse와 호환되는 형식으로 변환
    // NextResponse는 Buffer와 Uint8Array 모두 받을 수 있음
    const pdfData = pdf instanceof Buffer ? pdf : Buffer.from(pdf);
    return new NextResponse(pdfData, {
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${pdfFilename}"`,
        ...getSecurityHeaders(),
        'Cache-Control': 'no-cache, no-store, must-revalidate',
        'Pragma': 'no-cache',
        'Expires': '0'
      }
    });

  } catch (error) {
    logError(error, 'generate-pdf');
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    const errorStack = error instanceof Error ? error.stack : undefined;
    
    // 개발 환경에서는 상세한 에러 정보 로깅
    if (process.env.NODE_ENV !== 'production') {
      safeLog('❌ PDF 생성 에러:', errorMessage);
      if (errorStack) {
        safeLog('❌ 에러 스택:', errorStack);
      }
    }
    
    // 사용자에게 표시할 에러 메시지 (보안상 민감한 정보는 제외)
    let userErrorMessage = 'PDF 생성 중 오류가 발생했습니다.';
    
    // 에러 메시지에서 민감하지 않은 정보만 포함
    if (errorMessage && !errorMessage.includes('password') && !errorMessage.includes('token') && !errorMessage.includes('secret')) {
      // 에러 메시지가 너무 길면 자르기
      const shortMessage = errorMessage.length > 200 ? errorMessage.substring(0, 200) + '...' : errorMessage;
      userErrorMessage = `PDF 생성 중 오류가 발생했습니다: ${shortMessage}`;
    }
    
    return NextResponse.json(
      { 
        success: false, 
        error: userErrorMessage
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
