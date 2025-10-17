// 🔒 클라이언트 측 보안 유틸리티 함수들 (브라우저 환경용)

/**
 * Production 모드에서 모든 console 메서드를 무효화
 */
export const disableConsoleInProduction = () => {
  if (process.env.NODE_ENV === 'production') {
    const noop = () => {};
    console.log = noop;
    console.warn = noop;
    console.error = noop;
    console.info = noop;
    console.debug = noop;
    console.trace = noop;
  }
};

/**
 * 개발자 도구 열림 감지
 */
export const detectDevTools = () => {
  if (process.env.NODE_ENV === 'production') {
    // 개발자 도구 감지 (크기 기반)
    const threshold = 160;
    const widthThreshold = window.outerWidth - window.innerWidth > threshold;
    const heightThreshold = window.outerHeight - window.innerHeight > threshold;
    
    if (widthThreshold || heightThreshold) {
      return true;
    }
    
    // debugger 감지
    const start = new Date().getTime();
    debugger; // DevTools가 열려있으면 여기서 멈춤
    const end = new Date().getTime();
    
    return end - start > 100;
  }
  return false;
};

/**
 * 우클릭 방지
 */
export const disableRightClick = () => {
  if (process.env.NODE_ENV === 'production') {
    document.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      return false;
    });
  }
};

/**
 * 특정 키보드 단축키 방지 (F12, Ctrl+Shift+I, Ctrl+Shift+J, Ctrl+U)
 */
export const disableDevToolsShortcuts = () => {
  if (process.env.NODE_ENV === 'production') {
    document.addEventListener('keydown', (e) => {
      // F12
      if (e.key === 'F12') {
        e.preventDefault();
        return false;
      }
      
      // Ctrl+Shift+I (개발자 도구)
      if (e.ctrlKey && e.shiftKey && e.key === 'I') {
        e.preventDefault();
        return false;
      }
      
      // Ctrl+Shift+J (콘솔)
      if (e.ctrlKey && e.shiftKey && e.key === 'J') {
        e.preventDefault();
        return false;
      }
      
      // Ctrl+Shift+C (요소 선택)
      if (e.ctrlKey && e.shiftKey && e.key === 'C') {
        e.preventDefault();
        return false;
      }
      
      // Ctrl+U (소스 보기)
      if (e.ctrlKey && e.key === 'u') {
        e.preventDefault();
        return false;
      }
    });
  }
};

/**
 * 복사 방지
 */
export const disableCopy = () => {
  if (process.env.NODE_ENV === 'production') {
    document.addEventListener('copy', (e) => {
      e.preventDefault();
      return false;
    });
    
    document.addEventListener('cut', (e) => {
      e.preventDefault();
      return false;
    });
  }
};

/**
 * 텍스트 선택 방지
 */
export const disableTextSelection = () => {
  if (process.env.NODE_ENV === 'production') {
    document.body.style.userSelect = 'none';
    document.body.style.webkitUserSelect = 'none';
    (document.body.style as any).msUserSelect = 'none';
  }
};

/**
 * 개발자 도구 모니터링 (주기적으로 체크)
 */
export const monitorDevTools = (onDetected?: () => void) => {
  if (process.env.NODE_ENV === 'production') {
    setInterval(() => {
      if (detectDevTools()) {
        if (onDetected) {
          onDetected();
        } else {
          // 기본 동작: 페이지 리다이렉트 또는 경고
          document.body.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;height:100vh;font-size:24px;color:#dc2626;">⚠️ Unauthorized Access Detected</div>';
        }
      }
    }, 1000);
  }
};

/**
 * localStorage/sessionStorage 데이터 암호화
 */
export const encryptData = (data: string): string => {
  // 간단한 Base64 인코딩 (실제로는 더 강력한 암호화 필요)
  return btoa(encodeURIComponent(data));
};

export const decryptData = (encrypted: string): string => {
  try {
    return decodeURIComponent(atob(encrypted));
  } catch {
    return '';
  }
};

/**
 * Firebase 데이터 검증 (클라이언트 측)
 * @param data - 검증할 데이터
 * @returns 안전한 데이터
 */
export const sanitizeFirebaseDataClient = (data: any): any => {
  if (!data || typeof data !== 'object') {
    return {};
  }
  
  const sanitized: any = {};
  
  for (const [key, value] of Object.entries(data)) {
    // 키 검증
    const safeKey = String(key).replace(/[^a-zA-Z0-9_]/g, '');
    
    if (safeKey && safeKey.length <= 100) {
      if (typeof value === 'string') {
        // 문자열 길이 제한 및 특수문자 제거
        sanitized[safeKey] = value.substring(0, 1000).replace(/[<>]/g, '');
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
        sanitized[safeKey] = value.slice(0, 100);
      } else if (value && typeof value === 'object') {
        // 중첩 객체 검증
        sanitized[safeKey] = sanitizeFirebaseDataClient(value);
      }
    }
  }
  
  return sanitized;
};

/**
 * Firebase Document ID 검증 (클라이언트 측)
 * @param docId - 검증할 문서 ID
 * @returns 안전한 문서 ID
 */
export const sanitizeFirebaseDocIdClient = (docId: string): string => {
  // 영문자, 숫자, 하이픈, 언더스코어만 허용
  const sanitized = docId.replace(/[^a-zA-Z0-9_-]/g, '');
  
  // 길이 제한
  if (sanitized.length > 1500) {
    throw new Error('Document ID too long');
  }
  
  return sanitized;
};

/**
 * 모든 보안 조치 활성화
 */
export const enableAllSecurityMeasures = (options?: {
  disableConsole?: boolean;
  disableRightClick?: boolean;
  disableShortcuts?: boolean;
  disableCopy?: boolean;
  disableSelection?: boolean;
  monitorDevTools?: boolean;
}) => {
  const {
    disableConsole: shouldDisableConsole = true,
    disableRightClick: shouldDisableRightClick = true,
    disableShortcuts: shouldDisableShortcuts = true,
    disableCopy: shouldDisableCopy = false, // 사용자 경험을 위해 기본적으로 false
    disableSelection: shouldDisableSelection = false, // 사용자 경험을 위해 기본적으로 false
    monitorDevTools: shouldMonitor = false, // 너무 공격적일 수 있어 기본적으로 false
  } = options || {};
  
  if (shouldDisableConsole) disableConsoleInProduction();
  if (shouldDisableRightClick) disableRightClick();
  if (shouldDisableShortcuts) disableDevToolsShortcuts();
  if (shouldDisableCopy) disableCopy();
  if (shouldDisableSelection) disableTextSelection();
  if (shouldMonitor) monitorDevTools();
};

