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

