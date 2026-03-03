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
   * 서명 데이터 URL 검증 (클라이언트 측)
   * @param dataUrl - 검증할 데이터 URL (data:image/png;base64,... 형식)
   * @param maxSize - 최대 크기 (바이트, 기본값: 5MB)
   * @returns 검증 통과 시 true, 실패 시 false
   */
  export const validateSignatureDataUrlClient = (dataUrl: string | null | undefined, maxSize: number = 5 * 1024 * 1024): boolean => {
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
      // 브라우저 환경에서 Base64 디코딩
      const binaryString = atob(base64Data);
      const decodedSize = binaryString.length;
      
      // 디코딩된 크기가 합리적인 범위 내인지 확인
      if (decodedSize < 100 || decodedSize > maxSize) {
        return false;
      }
    } catch (error) {
      // Base64 디코딩 실패
      return false;
    }
    
    return true;
  };
  
  /**
   * 서명 데이터 URL 정제 및 반환 (클라이언트 측)
   * @param dataUrl - 검증할 데이터 URL
   * @param maxSize - 최대 크기 (바이트, 기본값: 5MB)
   * @returns 검증 통과 시 원본 데이터 URL, 실패 시 빈 문자열
   */
  export const sanitizeSignatureDataUrlClient = (dataUrl: string | null | undefined, maxSize: number = 5 * 1024 * 1024): string => {
    if (validateSignatureDataUrlClient(dataUrl, maxSize)) {
      return dataUrl!;
    }
    return '';
  };
  
  /**
   * 요청 서명 생성 (클라이언트 측)
   * @param data - 서명할 데이터
   * @param secret - 비밀 키 (환경 변수에서 가져옴, 클라이언트에서는 제한적)
   * @returns Base64 인코딩된 서명
   */
  export const generateRequestSignature = async (data: string): Promise<string> => {
    // 클라이언트에서는 간단한 해시 사용 (실제로는 서버에서 검증 필요)
    const encoder = new TextEncoder();
    const dataBuffer = encoder.encode(data);
    const hashBuffer = await crypto.subtle.digest('SHA-256', dataBuffer);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return btoa(String.fromCharCode(...hashArray));
  };

  /**
   * 타임스탬프 추가
   * @returns 현재 타임스탬프 (밀리초)
   */
  export const getTimestamp = (): number => {
    return Date.now();
  };

  /**
   * 보안 요청 생성 (서명 및 타임스탬프 포함)
   * @param data - 전송할 데이터
   * @returns 보안이 강화된 요청 데이터
   */
  export const createSecureRequest = async (data: any): Promise<any> => {
    const timestamp = getTimestamp();
    const dataString = JSON.stringify(data);
    
    // 서명 생성 (실제로는 서버에서 재검증)
    const signature = await generateRequestSignature(dataString + timestamp);
    
    return {
      data,
      timestamp,
      signature,
      // 추가 보안: 요청 ID (재전송 공격 방지)
      requestId: `${timestamp}-${Math.random().toString(36).substring(2, 15)}`
    };
  };

  /**
   * Firebase Authentication 토큰 가져오기 (강화된 버전)
   * @param forceRefresh - 토큰 강제 갱신 여부
   * @returns ID 토큰 또는 null
   */
  export const getAuthToken = async (forceRefresh: boolean = false): Promise<string | null> => {
    try {
      // Firebase Auth가 있는지 확인
      const { getAuth } = await import('firebase/auth');
      const auth = getAuth();
      
      if (!auth.currentUser) {
        return null;
      }
      
      // 토큰 가져오기 (만료 시 자동 갱신)
      const token = await auth.currentUser.getIdToken(forceRefresh);
      
      // 토큰 유효성 검증 (기본적인 형식 체크)
      if (!token || token.length < 100) {
        return null;
      }
      
      return token;
    } catch (error) {
      // Firebase Auth가 없거나 사용할 수 없는 경우
      return null;
    }
  };

  /**
   * Firebase Authentication 사용자 확인
   * @returns 사용자 정보 또는 null
   */
  export const getCurrentUser = async (): Promise<{ uid: string; email?: string } | null> => {
    try {
      const { getAuth } = await import('firebase/auth');
      const auth = getAuth();
      
      if (!auth.currentUser) {
        return null;
      }
      
      return {
        uid: auth.currentUser.uid,
        email: auth.currentUser.email || undefined,
      };
    } catch (error) {
      return null;
    }
  };

  /**
   * Firebase Authentication 초기화 대기
   * @param timeout - 최대 대기 시간 (밀리초, 기본값: 3초)
   * @returns 인증이 준비되면 true, 타임아웃 시 false
   */
  export const waitForAuthInit = async (timeout: number = 3000): Promise<boolean> => {
    try {
      const { getAuth, onAuthStateChanged } = await import('firebase/auth');
      const auth = getAuth();
      
      // 이미 초기화되어 있으면 즉시 반환
      if (auth.currentUser !== null) {
        return true;
      }
      
      // onAuthStateChanged를 사용하여 초기화 대기
      // 주의: onAuthStateChanged는 초기화 시 즉시 호출되므로, 사용자가 없어도 호출됨
      return new Promise((resolve) => {
        let resolved = false;
        let timeoutId: NodeJS.Timeout | null = null;
        
        const unsubscribe = onAuthStateChanged(auth, (user) => {
          if (!resolved) {
            resolved = true;
            if (timeoutId) clearTimeout(timeoutId);
            unsubscribe();
            // 사용자가 있으면 true, 없으면 false (로그인하지 않은 경우)
            resolve(user !== null);
          }
        });
        
        // 타임아웃 설정 (초기화가 너무 오래 걸리면 false 반환)
        timeoutId = setTimeout(() => {
          if (!resolved) {
            resolved = true;
            unsubscribe();
            // 타임아웃 시에도 한 번 더 확인
            resolve(auth.currentUser !== null);
          }
        }, timeout);
      });
    } catch (error) {
      return false;
    }
  };

  /**
   * Firebase Authentication 상태 확인 (강화된 버전 - 초기화 대기 포함)
   * @param waitForInit - 초기화 대기 여부 (기본값: true)
   * @returns 인증 상태
   */
  export const checkAuthState = async (waitForInit: boolean = true): Promise<boolean> => {
    try {
      const { getAuth } = await import('firebase/auth');
      const auth = getAuth();
      
      // 초기화 대기 (새로고침 시 필요)
      if (waitForInit && auth.currentUser === null) {
        const initialized = await waitForAuthInit(3000); // 3초 대기
        if (!initialized) {
          // 초기화 대기 후에도 사용자가 없으면 false
          return false;
        }
      }
      
      if (!auth.currentUser) {
        return false;
      }
      
      // 토큰도 확인 (토큰이 없거나 유효하지 않으면 false)
      try {
        const token = await auth.currentUser.getIdToken(false);
        if (!token || token.length < 100) {
          return false;
        }
        return true;
      } catch (tokenError) {
        // 토큰 가져오기 실패 시 false 반환
        return false;
      }
    } catch (error) {
      return false;
    }
  };

  /**
   * 보안 fetch 래퍼 (자동 서명 및 검증) - 강화된 버전
   * @param url - 요청 URL
   * @param options - fetch 옵션
   * @param requireAuth - 인증 필수 여부 (기본값: false)
   * @returns fetch 응답
   */
  export const secureFetch = async (
    url: string, 
    options: RequestInit = {},
    requireAuth: boolean = false
  ): Promise<Response> => {
    // 인증이 필요한 경우 사용자 확인
    if (requireAuth) {
      const isAuthenticated = await checkAuthState();
      if (!isAuthenticated) {
        throw new Error('Authentication required. Please log in first.');
      }
    }
    
    // Firebase Authentication 토큰 가져오기 (인증이 필요한 경우 강제 갱신)
    let authToken = await getAuthToken(requireAuth); // requireAuth이면 강제 갱신
    
    // 토큰이 없고 인증이 필요한 경우 재시도 (한 번 더 갱신 시도)
    if (requireAuth && !authToken) {
      // 한 번 더 강제 갱신 시도
      authToken = await getAuthToken(true);
      if (!authToken) {
        throw new Error('Authentication token not available. Please log in again.');
      }
    }
    
    // 요청 본문이 있으면 보안 강화
    if (options.body) {
      let bodyData: any;
      
      try {
        bodyData = JSON.parse(options.body as string);
      } catch {
        bodyData = options.body;
      }
      
      // 보안 요청 생성
      const secureData = await createSecureRequest(bodyData);
      
      // 보안 헤더 추가
      const secureHeaders = new Headers(options.headers);
      secureHeaders.set('Content-Type', 'application/json');
      secureHeaders.set('X-Requested-With', 'XMLHttpRequest');
      secureHeaders.set('X-Client-Time', String(getTimestamp()));
      
      // Firebase Authentication 토큰 추가 (있는 경우)
      if (authToken) {
        secureHeaders.set('Authorization', `Bearer ${authToken}`);
      }
      
      // 보안 요청 생성
      const secureOptions: RequestInit = {
        ...options,
        headers: secureHeaders,
        body: JSON.stringify(secureData),
        credentials: 'same-origin', // CSRF 보호
      };
      
      let response = await fetch(url, secureOptions);
      
      // 401 에러 시 토큰 갱신 후 재시도 (requireAuth인 경우에도 재시도)
      if (response.status === 401 && authToken) {
        const refreshedToken = await getAuthToken(true);
        if (refreshedToken && refreshedToken !== authToken) {
          secureHeaders.set('Authorization', `Bearer ${refreshedToken}`);
          response = await fetch(url, {
            ...secureOptions,
            headers: secureHeaders,
          });
          // 재시도 후에도 401이면 그대로 반환
        } else if (!refreshedToken) {
          // 토큰 갱신 실패 - 사용자가 로그아웃되었을 수 있음
          // 에러를 던지지 않고 401 응답을 그대로 반환하여 클라이언트에서 처리하도록 함
        }
      }
      
      return response;
    }
    
    // 본문이 없는 경우 기본 fetch
    const secureHeaders = new Headers(options.headers);
    secureHeaders.set('X-Requested-With', 'XMLHttpRequest');
    
    // Firebase Authentication 토큰 추가 (있는 경우)
    if (authToken) {
      secureHeaders.set('Authorization', `Bearer ${authToken}`);
    }
    
    let response = await fetch(url, {
      ...options,
      headers: secureHeaders,
      credentials: 'same-origin',
    });
    
    // 401 에러 시 토큰 갱신 후 재시도 (requireAuth인 경우에도 재시도)
    if (response.status === 401 && authToken) {
      const refreshedToken = await getAuthToken(true);
      if (refreshedToken && refreshedToken !== authToken) {
        secureHeaders.set('Authorization', `Bearer ${refreshedToken}`);
        response = await fetch(url, {
          ...options,
          headers: secureHeaders,
          credentials: 'same-origin',
        });
        // 재시도 후에도 401이면 그대로 반환
      } else if (!refreshedToken) {
        // 토큰 갱신 실패 - 사용자가 로그아웃되었을 수 있음
        // 에러를 던지지 않고 401 응답을 그대로 반환하여 클라이언트에서 처리하도록 함
      }
    }
    
    return response;
  };

  /**
   * 요청 재시도 방지 (같은 요청 ID로 중복 요청 차단)
   */
  const sentRequestIds = new Set<string>();
  const REQUEST_ID_TTL = 5 * 60 * 1000; // 5분

  /**
   * 요청 ID 검증 (중복 요청 방지)
   * @param requestId - 요청 ID
   * @returns 중복이 아니면 true
   */
  export const validateRequestId = (requestId: string): boolean => {
    // 오래된 요청 ID 정리
    if (sentRequestIds.size > 1000) {
      sentRequestIds.clear();
    }
    
    if (sentRequestIds.has(requestId)) {
      return false; // 중복 요청
    }
    
    sentRequestIds.add(requestId);
    
    // TTL 후 자동 제거
    setTimeout(() => {
      sentRequestIds.delete(requestId);
    }, REQUEST_ID_TTL);
    
    return true;
  };

  /**
   * 클라이언트 핑거프린트 생성 (기기 식별)
   * @returns 고유한 핑거프린트
   */
  export const generateClientFingerprint = async (): Promise<string> => {
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    if (ctx) {
      ctx.textBaseline = 'top';
      ctx.font = '14px Arial';
      ctx.fillText('Fingerprint', 2, 2);
    }
    
    const fingerprint = [
      navigator.userAgent,
      navigator.language,
      screen.width + 'x' + screen.height,
      new Date().getTimezoneOffset(),
      canvas.toDataURL(),
    ].join('|');
    
    const encoder = new TextEncoder();
    const data = encoder.encode(fingerprint);
    const hashBuffer = await crypto.subtle.digest('SHA-256', data);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return btoa(String.fromCharCode(...hashArray)).substring(0, 32);
  };

  /**
   * CSV Injection 방지 (CSV 파일용) - 클라이언트 측
   * @param text - CSV 셀에 들어갈 텍스트
   * @returns CSV Injection이 방지된 안전한 텍스트
   */
  export const sanitizeCSVCell = (text: string | number | null | undefined): string => {
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
  };

  /**
   * PDF 파일명 검증 및 정제 (클라이언트 측)
   * @param filename - 검증할 파일명
   * @returns 안전한 파일명
   */
  export const sanitizePdfFilename = (filename: string): string => {
    if (!filename || typeof filename !== 'string') {
      return 'document.html';
    }
    
    // 특수문자 제거 (영문자, 숫자, 하이픈, 언더스코어, 점만 허용)
    const sanitized = filename.replace(/[^a-zA-Z0-9._-]/g, '_');
    
    // 길이 제한 (최대 200자)
    const limited = sanitized.substring(0, 200);
    
    // 경로 traversal 공격 방지 (../ 제거)
    const safePath = limited.replace(/\.\./g, '');
    
    return safePath;
  };

  /**
   * 날짜 형식 검증 (YYYY-MM-DD)
   * @param date - 검증할 날짜 문자열
   * @returns 유효한 날짜이면 true
   */
  export const isValidDate = (date: string): boolean => {
    if (!date || typeof date !== 'string') return false;
    
    // YYYY-MM-DD 형식 체크
    const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
    if (!dateRegex.test(date)) return false;
    
    // 실제 날짜 유효성 체크
    const dateObj = new Date(date);
    return dateObj instanceof Date && !isNaN(dateObj.getTime());
  };

  /**
   * 파일 크기 검증 (DoS 방지)
   * @param blob - 검증할 Blob 객체
   * @param maxSize - 최대 크기 (바이트, 기본값: 10MB)
   * @returns 허용 가능한 크기이면 true
   */
  export const validateFileSize = (blob: Blob, maxSize: number = 10 * 1024 * 1024): boolean => {
    return blob.size > 0 && blob.size <= maxSize;
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
  
