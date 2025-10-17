// 🔒 서버 측 보안 유틸리티 함수 (API Routes용)

/**
 * HTML 특수 문자 이스케이프 (XSS 방지)
 * @param text - 이스케이프할 텍스트
 * @returns 안전하게 이스케이프된 텍스트
 */
export function escapeHtml(text: string | null | undefined): string {
  if (!text) return '';
  
  const map: Record<string, string> = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#x27;',
    '/': '&#x2F;',
  };
  
  return String(text).replace(/[&<>"'/]/g, (char) => map[char] || char);
}

/**
 * 숫자만 추출 (카드 번호 등)
 * @param text - 입력 텍스트
 * @returns 숫자만 포함된 문자열
 */
export function sanitizeCardNumber(text: string | null | undefined): string {
  if (!text) return '';
  return String(text).replace(/[^0-9]/g, '');
}

/**
 * 금액 검증 및 안전한 파싱
 * @param amount - 금액 문자열
 * @returns 안전하게 파싱된 숫자
 */
export function sanitizeAmount(amount: string | number | null | undefined): number {
  if (amount === null || amount === undefined) return 0;
  
  const parsed = parseFloat(String(amount));
  
  // NaN이거나 음수이거나 너무 큰 값은 0으로
  if (isNaN(parsed) || parsed < 0 || parsed > 1000000) {
    return 0;
  }
  
  return parsed;
}

/**
 * 날짜 검증 (YYYY-MM-DD 형식)
 * @param date - 날짜 문자열
 * @returns 유효한 날짜이면 true
 */
export function isValidDate(date: string | null | undefined): boolean {
  if (!date) return false;
  
  // YYYY-MM-DD 형식 체크
  const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
  if (!dateRegex.test(date)) return false;
  
  // 실제 날짜 유효성 체크
  const dateObj = new Date(date);
  return dateObj instanceof Date && !isNaN(dateObj.getTime());
}

/**
 * 안전한 문자열 추출 (길이 제한)
 * @param text - 입력 텍스트
 * @param maxLength - 최대 길이
 * @returns 길이 제한된 안전한 문자열
 */
export function sanitizeString(text: string | null | undefined, maxLength: number = 500): string {
  if (!text) return '';
  
  // HTML 이스케이프 + 길이 제한
  const escaped = escapeHtml(text);
  return escaped.substring(0, maxLength);
}

/**
 * Production 환경에서만 실행되는 안전한 로그
 * @param message - 로그 메시지 (민감 정보 제외)
 */
export function safeLog(message: string, data?: any): void {
  // Production에서는 로그 출력 안 함
  if (process.env.NODE_ENV === 'production') {
    return;
  }
  
  // Development에서만 출력
  console.log(message, data);
}

/**
 * 에러 로깅 (production에서도 중요한 에러는 기록)
 * @param error - 에러 객체
 * @param context - 에러 발생 컨텍스트
 */
export function logError(error: unknown, context: string): void {
  const errorMessage = error instanceof Error ? error.message : 'Unknown error';
  
  // Production에서는 간단한 에러 메시지만
  if (process.env.NODE_ENV === 'production') {
    console.error(`[${context}] Error:`, errorMessage);
  } else {
    // Development에서는 전체 에러 출력
    console.error(`[${context}] Error:`, error);
  }
}

/**
 * CSV Injection 방지 (CSV 파일용)
 * @param text - CSV 셀에 들어갈 텍스트
 * @returns CSV Injection이 방지된 안전한 텍스트
 */
export function sanitizeCSVCell(text: string | number | null | undefined): string {
  if (text === null || text === undefined) return '';
  
  const str = String(text);
  
  // 1. 빈 문자열
  if (str.trim() === '') return '';
  
  // 2. CSV Formula Injection 방지 (=, +, -, @, 탭, 캐리지 리턴으로 시작)
  if (str.match(/^[=+\-@\t\r]/)) {
    return `'${str.replace(/"/g, '""')}`;
  }
  
  // 3. 특수문자 escape (쉼표, 따옴표, 줄바꿈 포함 시 따옴표로 감싸기)
  if (str.includes(',') || str.includes('"') || str.includes('\n') || str.includes('\r')) {
    return `"${str.replace(/"/g, '""')}"`;
  }
  
  // 4. 일반 텍스트
  return str;
}

/**
 * Firebase Collection 이름 검증 (NoSQL Injection 방지)
 * @param collectionName - 검증할 컬렉션 이름
 * @returns 안전한 컬렉션 이름
 */
export function sanitizeFirebaseCollection(collectionName: string): string {
  // 허용된 컬렉션 이름만 사용
  const allowedCollections = [
    'patient-logs',
    'fee-reduction-requests',
    'credit-card-receipts',
    'email-notifications'
  ];
  
  if (!allowedCollections.includes(collectionName)) {
    throw new Error('Invalid collection name');
  }
  
  return collectionName;
}

/**
 * Firebase Document ID 검증
 * @param docId - 검증할 문서 ID
 * @returns 안전한 문서 ID
 */
export function sanitizeFirebaseDocId(docId: string): string {
  // Firebase 문서 ID는 영문자, 숫자, 하이픈, 언더스코어만 허용
  const sanitized = docId.replace(/[^a-zA-Z0-9_-]/g, '');
  
  // 길이 제한 (Firebase 문서 ID 최대 길이)
  if (sanitized.length > 1500) {
    throw new Error('Document ID too long');
  }
  
  return sanitized;
}

/**
 * Firebase 데이터 검증 및 정제
 * @param data - 검증할 데이터 객체
 * @returns 안전한 데이터 객체
 */
export function sanitizeFirebaseData(data: any): any {
  if (!data || typeof data !== 'object') {
    return {};
  }
  
  const sanitized: any = {};
  
  for (const [key, value] of Object.entries(data)) {
    // 키 검증 (영문자, 숫자, 언더스코어만 허용)
    const safeKey = String(key).replace(/[^a-zA-Z0-9_]/g, '');
    
    if (safeKey && safeKey.length <= 100) {
      if (typeof value === 'string') {
        sanitized[safeKey] = sanitizeString(value, 1000);
      } else if (typeof value === 'number') {
        // 숫자 범위 검증
        if (value >= -1e10 && value <= 1e10 && !isNaN(value)) {
          sanitized[safeKey] = value;
        }
      } else if (typeof value === 'boolean') {
        sanitized[safeKey] = value;
      } else if (value instanceof Date) {
        sanitized[safeKey] = value;
      } else if (Array.isArray(value)) {
        // 배열 검증 (최대 100개 요소)
        sanitized[safeKey] = value.slice(0, 100).map(item => 
          typeof item === 'string' ? sanitizeString(item, 500) : item
        );
      } else if (value && typeof value === 'object') {
        // 중첩 객체 검증 (최대 2단계)
        sanitized[safeKey] = sanitizeFirebaseData(value);
      }
    }
  }
  
  return sanitized;
}

/**
 * Firebase 쿼리 제한 (DoS 방지)
 * @param limit - 쿼리 제한 수
 * @returns 안전한 제한 수
 */
export function sanitizeFirebaseLimit(limit: number): number {
  const maxLimit = 1000; // 최대 1000개 문서
  const minLimit = 1;    // 최소 1개 문서
  
  if (typeof limit !== 'number' || isNaN(limit)) {
    return 50; // 기본값
  }
  
  return Math.max(minLimit, Math.min(maxLimit, Math.floor(limit)));
}

/**
 * PDF 파일명 검증 및 정제
 * @param filename - 검증할 파일명
 * @returns 안전한 파일명
 */
export function sanitizePdfFilename(filename: string): string {
  if (!filename || typeof filename !== 'string') {
    return 'document.pdf';
  }
  
  // 특수문자 제거 (영문자, 숫자, 하이픈, 언더스코어, 점만 허용)
  const sanitized = filename.replace(/[^a-zA-Z0-9._-]/g, '_');
  
  // 길이 제한 (최대 100자)
  const limited = sanitized.substring(0, 100);
  
  // .pdf 확장자 추가 (없는 경우)
  if (!limited.toLowerCase().endsWith('.pdf')) {
    return `${limited}.pdf`;
  }
  
  return limited;
}

/**
 * HTML 콘텐츠 보안 검증
 * @param html - 검증할 HTML 문자열
 * @returns 안전한 HTML 문자열
 */
export function sanitizeHtmlContent(html: string): string {
  if (!html || typeof html !== 'string') {
    return '';
  }
  
  // HTML 크기 제한 (최대 1MB)
  if (html.length > 1024 * 1024) {
    throw new Error('HTML content too large');
  }
  
  // 위험한 태그 및 속성 제거
  const dangerousTags = [
    'script', 'iframe', 'object', 'embed', 'applet', 'form', 'input', 'button',
    'link', 'meta', 'style', 'base', 'frame', 'frameset'
  ];
  
  let sanitized = html;
  
  // 위험한 태그 제거
  dangerousTags.forEach(tag => {
    const regex = new RegExp(`<${tag}[^>]*>.*?</${tag}>`, 'gi');
    sanitized = sanitized.replace(regex, '');
    
    const selfClosingRegex = new RegExp(`<${tag}[^>]*/?>`, 'gi');
    sanitized = sanitized.replace(selfClosingRegex, '');
  });
  
  // 위험한 속성 제거
  const dangerousAttributes = [
    'onload', 'onerror', 'onclick', 'onmouseover', 'onfocus', 'onblur',
    'onchange', 'onsubmit', 'onreset', 'onselect', 'onkeydown', 'onkeyup',
    'onkeypress', 'onmousedown', 'onmouseup', 'onmousemove', 'onmouseout',
    'javascript:', 'vbscript:', 'data:', 'file:'
  ];
  
  dangerousAttributes.forEach(attr => {
    const regex = new RegExp(`${attr}\\s*=\\s*["'][^"']*["']`, 'gi');
    sanitized = sanitized.replace(regex, '');
  });
  
  return sanitized;
}

/**
 * Puppeteer 보안 옵션 생성
 * @returns 안전한 Puppeteer 실행 옵션
 */
export function getSecurePuppeteerOptions(): any {
  return {
    headless: true,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-accelerated-2d-canvas',
      '--no-first-run',
      '--no-zygote',
      '--disable-gpu',
      '--disable-background-timer-throttling',
      '--disable-backgrounding-occluded-windows',
      '--disable-renderer-backgrounding',
      '--disable-features=TranslateUI',
      '--disable-ipc-flooding-protection',
      '--disable-extensions',
      '--disable-plugins',
      '--disable-images',
      '--disable-javascript',
      '--disable-web-security',
      '--disable-features=VizDisplayCompositor',
      '--memory-pressure-off',
      '--max_old_space_size=4096'
    ],
    timeout: 30000, // 30초 타임아웃
    protocolTimeout: 30000
  };
}

/**
 * PDF 생성 제한 (DoS 방지)
 * @param dataSize - 데이터 크기
 * @returns 허용 여부
 */
export function validatePdfGeneration(dataSize: number): boolean {
  const maxDataSize = 10 * 1024 * 1024; // 10MB 제한
  
  if (dataSize > maxDataSize) {
    return false;
  }
  
  return true;
}

/**
 * 안전한 배열 처리 (PDF 생성용)
 * @param array - 처리할 배열
 * @param maxLength - 최대 길이
 * @returns 안전한 배열
 */
export function sanitizeArrayForPdf(array: any[], maxLength: number = 1000): any[] {
  if (!Array.isArray(array)) {
    return [];
  }
  
  // 배열 길이 제한
  const limitedArray = array.slice(0, maxLength);
  
  // 각 요소 검증
  return limitedArray.map((item, index) => {
    if (!item || typeof item !== 'object') {
      return null;
    }
    
    // 기본적인 객체 검증
    const safeItem: any = {};
    
    // 허용된 필드만 추출
    const allowedFields = [
      'name', 'office', 'appt_date', 'apptDate', 'visit_type', 'visitType',
      'call_in', 'callIn', 'call_out', 'callOut', 'time', 'remark', 'other_duty', 'otherDuty'
    ];
    
    allowedFields.forEach(field => {
      if (item[field] !== undefined && item[field] !== null) {
        if (typeof item[field] === 'string') {
          safeItem[field] = sanitizeString(item[field], 500);
        } else if (typeof item[field] === 'boolean') {
          safeItem[field] = Boolean(item[field]);
        } else if (typeof item[field] === 'number') {
          safeItem[field] = isNaN(item[field]) ? 0 : item[field];
        }
      }
    });
    
    return safeItem;
  }).filter(item => item !== null);
}

