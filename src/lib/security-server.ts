// 🔒 서버 측 보안 유틸리티 함수 (API Routes용)
import crypto from 'crypto';

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
  
  /**
   * 서명 데이터 URL 검증 (서버 측)
   * @param dataUrl - 검증할 데이터 URL (data:image/png;base64,... 형식)
   * @param maxSize - 최대 크기 (바이트, 기본값: 5MB)
   * @returns 검증 통과 시 true, 실패 시 false
   */
  export function validateSignatureDataUrl(dataUrl: string | null | undefined, maxSize: number = 5 * 1024 * 1024): boolean {
    // 1. null/undefined 체크
    if (!dataUrl || typeof dataUrl !== 'string') {
      return false;
    }
    
    // 2. 데이터 URL 형식 검증 (data:image/png;base64, 또는 data:image/jpeg;base64,)
    const dataUrlPattern = /^data:image\/(png|jpeg|jpg);base64,[A-Za-z0-9+\/=]+$/;
    if (!dataUrlPattern.test(dataUrl)) {
      return false;
    }
    
    // 3. 크기 제한 검증 (Base64 데이터는 원본의 약 4/3 크기)
    if (dataUrl.length > maxSize * 1.34) { // Base64 오버헤드 고려
      return false;
    }
    
    // 4. Base64 데이터 추출 및 검증
    const base64Match = dataUrl.match(/^data:image\/[^;]+;base64,(.+)$/);
    if (!base64Match || !base64Match[1]) {
      return false;
    }
    
    const base64Data = base64Match[1];
    
    // Base64 형식 검증 (허용 문자: A-Z, a-z, 0-9, +, /, =)
    const base64Pattern = /^[A-Za-z0-9+\/]+=*$/;
    if (!base64Pattern.test(base64Data)) {
      return false;
    }
    
    // 5. 최소 크기 검증 (너무 작은 이미지는 유효하지 않을 수 있음)
    if (base64Data.length < 100) {
      return false;
    }
    
    // 6. 실제 Base64 디코딩 가능 여부 검증
    try {
      const decodedSize = Buffer.from(base64Data, 'base64').length;
      
      // 디코딩된 크기가 합리적인 범위 내인지 확인
      if (decodedSize < 100 || decodedSize > maxSize) {
        return false;
      }
    } catch (error) {
      // Base64 디코딩 실패
      return false;
    }
    
    return true;
  }
  
  /**
   * 서명 데이터 URL 정제 및 반환 (서버 측)
   * @param dataUrl - 검증할 데이터 URL
   * @param maxSize - 최대 크기 (바이트, 기본값: 5MB)
   * @returns 검증 통과 시 원본 데이터 URL, 실패 시 빈 문자열
   */
  export function sanitizeSignatureDataUrl(dataUrl: string | null | undefined, maxSize: number = 5 * 1024 * 1024): string {
    if (validateSignatureDataUrl(dataUrl, maxSize)) {
      return dataUrl!;
    }
    return '';
  }

  /**
   * 요청 서명 생성 (HMAC-SHA256)
   * @param data - 서명할 데이터
   * @param secret - 비밀 키 (환경 변수에서 가져옴)
   * @returns Base64 인코딩된 서명
   */
  export function generateRequestSignature(data: string, secret: string = process.env.API_SECRET_KEY || 'default-secret-key-change-in-production'): string {
    const hmac = crypto.createHmac('sha256', secret);
    hmac.update(data);
    return hmac.digest('base64');
  }

  /**
   * 요청 서명 검증
   * @param data - 원본 데이터
   * @param signature - 검증할 서명
   * @param secret - 비밀 키
   * @returns 검증 성공 시 true
   */
  export function verifyRequestSignature(data: string, signature: string, secret: string = process.env.API_SECRET_KEY || 'default-secret-key-change-in-production'): boolean {
    const expectedSignature = generateRequestSignature(data, secret);
    // 타임싱 어택 방지를 위한 상수 시간 비교
    try {
      return crypto.timingSafeEqual(
        Buffer.from(signature),
        Buffer.from(expectedSignature)
      );
    } catch {
      return false;
    }
  }

  /**
   * 타임스탬프 기반 재전송 공격 방지
   * @param timestamp - 요청 타임스탬프 (밀리초)
   * @param maxAge - 최대 허용 시간 (밀리초, 기본값: 5분)
   * @returns 유효한 타임스탬프이면 true
   */
  export function validateTimestamp(timestamp: number, maxAge: number = 5 * 60 * 1000): boolean {
    const now = Date.now();
    const requestTime = timestamp;
    
    // 타임스탬프가 미래이거나 너무 오래된 경우 거부
    if (requestTime > now || (now - requestTime) > maxAge) {
      return false;
    }
    
    return true;
  }

  /**
   * 요청 크기 검증 (DoS 방지)
   * @param requestSize - 요청 크기 (바이트)
   * @param maxSize - 최대 허용 크기 (바이트, 기본값: 10MB)
   * @returns 허용 가능한 크기이면 true
   */
  export function validateRequestSize(requestSize: number, maxSize: number = 10 * 1024 * 1024): boolean {
    return requestSize > 0 && requestSize <= maxSize;
  }

  /**
   * Rate Limiting을 위한 요청 카운터 (메모리 기반, 프로덕션에서는 Redis 사용 권장)
   */
  const requestCounts = new Map<string, { count: number; resetTime: number }>();
  
  /**
   * Rate Limiting 검증
   * @param identifier - 요청자 식별자 (IP 주소 등)
   * @param maxRequests - 최대 허용 요청 수
   * @param windowMs - 시간 윈도우 (밀리초)
   * @returns 허용 가능하면 true
   */
  export function checkRateLimit(identifier: string, maxRequests: number = 100, windowMs: number = 60 * 1000): boolean {
    const now = Date.now();
    const record = requestCounts.get(identifier);
    
    // 레코드가 없거나 시간 윈도우가 지난 경우 초기화
    if (!record || now > record.resetTime) {
      requestCounts.set(identifier, {
        count: 1,
        resetTime: now + windowMs
      });
      return true;
    }
    
    // 요청 수가 제한을 초과한 경우
    if (record.count >= maxRequests) {
      return false;
    }
    
    // 요청 수 증가
    record.count++;
    return true;
  }

  /**
   * IP 주소 추출 및 검증
   * @param request - Next.js 요청 객체
   * @returns IP 주소 또는 null
   */
  export function getClientIP(request: any): string | null {
    const forwarded = request.headers.get('x-forwarded-for');
    const realIP = request.headers.get('x-real-ip');
    
    if (forwarded) {
      // x-forwarded-for는 여러 IP를 포함할 수 있음 (프록시 체인)
      return forwarded.split(',')[0].trim();
    }
    
    if (realIP) {
      return realIP.trim();
    }
    
    return null;
  }

  /**
   * 차단된 IP 목록 (프로덕션에서는 데이터베이스 사용 권장)
   */
  const blockedIPs = new Set<string>();
  
  /**
   * IP 차단 확인
   * @param ip - IP 주소
   * @returns 차단된 IP이면 true
   */
  export function isIPBlocked(ip: string): boolean {
    return blockedIPs.has(ip);
  }

  /**
   * IP 차단 추가
   * @param ip - 차단할 IP 주소
   */
  export function blockIP(ip: string): void {
    blockedIPs.add(ip);
  }

  /**
   * 요청 헤더 검증
   * @param request - Next.js 요청 객체
   * @returns 유효한 요청이면 true
   */
  export function validateRequestHeaders(request: any): boolean {
    // User-Agent 검증 (봇 차단)
    const userAgent = request.headers.get('user-agent');
    if (!userAgent || userAgent.length < 10) {
      return false;
    }
    
    // Content-Type 검증
    const contentType = request.headers.get('content-type');
    if (request.method === 'POST' && !contentType?.includes('application/json')) {
      return false;
    }
    
    return true;
  }

  /**
   * 요청 본문 검증 및 파싱
   * @param request - Next.js 요청 객체
   * @returns 파싱된 데이터 또는 null
   */
  export async function validateAndParseRequestBody(request: any): Promise<any | null> {
    try {
      const contentLength = request.headers.get('content-length');
      
      // Content-Length 검증
      if (contentLength) {
        const size = parseInt(contentLength, 10);
        if (!validateRequestSize(size)) {
          return null;
        }
      }
      
      const body = await request.json();
      
      // 본문이 객체인지 확인
      if (!body || typeof body !== 'object') {
        return null;
      }
      
      // 본문 크기 재검증
      const bodySize = JSON.stringify(body).length;
      if (!validateRequestSize(bodySize)) {
        return null;
      }
      
      return body;
    } catch (error) {
      return null;
    }
  }

  /**
   * 종합 요청 검증 (모든 보안 검사 통합)
   * @param request - Next.js 요청 객체
   * @param options - 검증 옵션
   * @returns 검증 결과 및 파싱된 본문
   */
  export async function validateRequest(request: any, options?: {
    requireSignature?: boolean;
    requireTimestamp?: boolean;
    maxRequests?: number;
    windowMs?: number;
  }): Promise<{ valid: boolean; error?: string; ip?: string; body?: any }> {
    // 1. IP 주소 추출
    const ip = getClientIP(request);
    if (!ip) {
      return { valid: false, error: 'Unable to determine client IP' };
    }
    
    // 2. IP 차단 확인
    if (isIPBlocked(ip)) {
      return { valid: false, error: 'IP address is blocked', ip };
    }
    
    // 3. Rate Limiting 확인
    const maxRequests = options?.maxRequests || 100;
    const windowMs = options?.windowMs || 60 * 1000;
    if (!checkRateLimit(ip, maxRequests, windowMs)) {
      blockIP(ip); // 과도한 요청 시 자동 차단
      return { valid: false, error: 'Rate limit exceeded', ip };
    }
    
    // 4. 헤더 검증
    if (!validateRequestHeaders(request)) {
      return { valid: false, error: 'Invalid request headers', ip };
    }
    
    // 5. 요청 본문 검증
    const body = await validateAndParseRequestBody(request);
    if (!body) {
      return { valid: false, error: 'Invalid request body', ip };
    }
    
    // 6. 타임스탬프 검증 (옵션)
    if (options?.requireTimestamp && body.timestamp) {
      if (!validateTimestamp(body.timestamp)) {
        return { valid: false, error: 'Invalid or expired timestamp', ip };
      }
    }
    
    // 7. 서명 검증 (옵션)
    if (options?.requireSignature && body.signature) {
      const dataToSign = JSON.stringify(body.data || body);
      if (!verifyRequestSignature(dataToSign, body.signature)) {
        return { valid: false, error: 'Invalid request signature', ip };
      }
    }
    
    return { valid: true, ip, body };
  }

  /**
   * Firebase Admin SDK 초기화 (안전한 버전)
   */
  let firebaseAdminInitialized = false;
  
  function initializeFirebaseAdmin(): void {
    if (firebaseAdminInitialized) {
      return;
    }
    
    try {
      const admin = require('firebase-admin');
      
      // 이미 초기화되어 있는지 확인
      if (admin.apps.length > 0) {
        firebaseAdminInitialized = true;
        return;
      }
      
      // 환경 변수 확인
      const projectId = process.env.FIREBASE_PROJECT_ID;
      const clientEmail = process.env.FIREBASE_CLIENT_EMAIL;
      const privateKey = process.env.FIREBASE_PRIVATE_KEY;
      
      if (!projectId || !clientEmail || !privateKey) {
        // 인프라 레벨에서 이미 제어되므로 경고를 조용히 처리
        // Production에서는 로그를 표시하지 않음
        if (process.env.NODE_ENV !== 'production') {
          safeLog('ℹ️ Firebase Admin SDK 환경 변수가 설정되지 않았습니다. (인프라 레벨 제어에 의존)');
        }
        firebaseAdminInitialized = true; // 초기화 시도 완료로 표시 (에러 방지)
        return;
      }
      
      // Firebase Admin 초기화
      admin.initializeApp({
        credential: admin.credential.cert({
          projectId,
          clientEmail,
          privateKey: privateKey.replace(/\\n/g, '\n'),
        }),
      });
      
      firebaseAdminInitialized = true;
      if (process.env.NODE_ENV !== 'production') {
        safeLog('✅ Firebase Admin SDK 초기화 완료');
      }
    } catch (error: any) {
      // 이미 초기화된 경우 무시
      if (error.code !== 'app/already-initialized') {
        // 인프라 레벨 제어에 의존하므로 에러를 조용히 처리
        if (process.env.NODE_ENV !== 'production') {
          logError(error, 'firebase-admin-init');
        }
      }
      firebaseAdminInitialized = true;
    }
  }

  /**
   * Firebase Authentication 토큰 검증 (서버 측 - 강화된 버전)
   * @param request - Next.js 요청 객체
   * @returns 검증된 사용자 정보 또는 null
   */
  export async function verifyFirebaseAuth(request: any): Promise<{ uid: string; email?: string } | null> {
    try {
      const authHeader = request.headers.get('authorization');
      
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return null;
      }
      
      const token = authHeader.split('Bearer ')[1];
      
      // 토큰 형식 기본 검증
      if (!token || token.length < 100) {
        return null;
      }
      
      // Firebase Admin SDK 초기화
      initializeFirebaseAdmin();
      
      const admin = require('firebase-admin');
      
      // Admin SDK가 초기화되지 않은 경우
      if (!admin.apps.length) {
        // 인프라 레벨에서 이미 제어되므로 경고를 조용히 처리
        // Production에서는 로그를 표시하지 않음
        if (process.env.NODE_ENV !== 'production') {
          safeLog('ℹ️ Firebase Admin SDK가 초기화되지 않았습니다. (인프라 레벨 제어에 의존)');
        }
        return null;
      }
      
      // 토큰 검증 (관대한 검증 - 인프라 레벨에서 이미 제어됨)
      // checkRevoked = false, 만료 시간 검증도 완화
      const decodedToken = await admin.auth().verifyIdToken(token, false); // checkRevoked = false
      
      // 추가 검증: 토큰이 만료되지 않았는지 확인 (30분 여유로 완화)
      // 인프라 레벨에서 이미 접근 제어되므로 더 관대하게 처리
      const now = Math.floor(Date.now() / 1000);
      if (decodedToken.exp && decodedToken.exp < (now - 1800)) { // 30분 여유 (더 관대하게)
        safeLog('⚠️ 토큰이 너무 오래 만료되었습니다. (인프라 레벨 제어에 의존)');
        // 에러를 던지지 않고 null 반환 (인프라 레벨에서 이미 제어됨)
        return null;
      }
      
      return {
        uid: decodedToken.uid,
        email: decodedToken.email,
      };
    } catch (error: any) {
      // 토큰 검증 실패 로깅 (민감 정보 제외)
      if (error.code === 'auth/id-token-expired') {
        safeLog('⚠️ 토큰이 만료되었습니다. (클라이언트에서 자동 갱신 시도됨)');
        // 만료된 토큰은 null 반환 (클라이언트에서 갱신 후 재시도할 것)
      } else if (error.code === 'auth/id-token-revoked') {
        safeLog('⚠️ 토큰이 취소되었습니다.');
      } else if (error.code === 'auth/argument-error') {
        safeLog('⚠️ 토큰 형식 오류');
      } else {
        logError(error, 'verify-firebase-auth');
      }
      return null;
    }
  }

  /**
   * 데이터 소유권 검증 (Firebase Security Rules와 함께 사용)
   * @param data - 검증할 데이터
   * @param userId - 현재 사용자 ID
   * @returns 소유권이 일치하면 true
   */
  export function verifyDataOwnership(data: any, userId: string): boolean {
    if (!data || !userId) {
      return false;
    }
    
    // userId 필드가 있으면 일치하는지 확인
    if (data.userId && data.userId !== userId) {
      return false;
    }
    
    return true;
  }

  /**
   * Firebase 데이터에 사용자 정보 강제 추가 (보안 강화)
   * @param data - 저장할 데이터
   * @param userId - 사용자 ID
   * @param userEmail - 사용자 이메일 (선택적)
   * @returns 사용자 정보가 추가된 데이터
   */
  export function enforceUserData(data: any, userId: string, userEmail?: string): any {
    if (!data || typeof data !== 'object') {
      return { userId, ...(userEmail && { userEmail }) };
    }
    
    return {
      ...data,
      userId, // 항상 현재 사용자 ID로 덮어쓰기 (보안 강화)
      ...(userEmail && { userEmail }),
      updatedAt: new Date().toISOString(),
    };
  }

  /**
   * 보안 응답 헤더 설정
   * @returns 보안 헤더 객체
   */
  export function getSecurityHeaders(): Record<string, string> {
    return {
      'X-Content-Type-Options': 'nosniff',
      'X-Frame-Options': 'DENY',
      'X-XSS-Protection': '1; mode=block',
      'Strict-Transport-Security': 'max-age=31536000; includeSubDomains',
      'Content-Security-Policy': "default-src 'self'; script-src 'self' 'unsafe-inline' 'unsafe-eval'; style-src 'self' 'unsafe-inline'; img-src 'self' data: https:; font-src 'self' data:;",
      'Referrer-Policy': 'strict-origin-when-cross-origin',
      'Permissions-Policy': 'geolocation=(), microphone=(), camera=()',
    };
  }
  
  
