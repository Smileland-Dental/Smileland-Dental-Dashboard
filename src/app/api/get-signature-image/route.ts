import { NextRequest, NextResponse } from 'next/server';

// 🔒 보안 설정
const MAX_FILE_SIZE = 5 * 1024 * 1024; // 5MB
const ALLOWED_DOMAINS = [
  'firebasestorage.googleapis.com',
  'storage.googleapis.com'
];
const ALLOWED_CONTENT_TYPES = [
  'image/png',
  'image/jpeg',
  'image/jpg',
  'image/webp'
];

// Rate limiting (간단한 in-memory 구현)
const requestCounts = new Map<string, { count: number; resetTime: number }>();
const RATE_LIMIT = 50; // 분당 최대 요청 수
const RATE_WINDOW = 60 * 1000; // 1분

function checkRateLimit(ip: string): boolean {
  const now = Date.now();
  const record = requestCounts.get(ip);

  if (!record || now > record.resetTime) {
    requestCounts.set(ip, { count: 1, resetTime: now + RATE_WINDOW });
    return true;
  }

  if (record.count >= RATE_LIMIT) {
    return false;
  }

  record.count++;
  return true;
}

// URL 검증 함수 (SSRF 방지)
function isValidFirebaseUrl(url: string): boolean {
  try {
    const parsedUrl = new URL(url);
    
    // 1. HTTPS만 허용
    if (parsedUrl.protocol !== 'https:') {
      return false;
    }
    
    // 2. Firebase Storage 도메인만 허용
    if (!ALLOWED_DOMAINS.some(domain => parsedUrl.hostname === domain || parsedUrl.hostname.endsWith('.' + domain))) {
      return false;
    }
    
    // 3. 내부 IP 주소 차단 (추가 보안)
    const hostname = parsedUrl.hostname;
    if (
      hostname === 'localhost' ||
      hostname === '127.0.0.1' ||
      hostname.startsWith('192.168.') ||
      hostname.startsWith('10.') ||
      hostname.startsWith('172.16.') ||
      hostname.startsWith('169.254.') // AWS 메타데이터
    ) {
      return false;
    }
    
    // 4. 경로에 "Employee Fee Reduction Signature" 또는 "signatures" 포함 확인
    const decodedPath = decodeURIComponent(parsedUrl.pathname);
    if (!decodedPath.includes('Employee Fee Reduction Signature') && 
        !decodedPath.includes('signatures/')) {
      return false;
    }
    
    return true;
  } catch (error) {
    return false;
  }
}

export async function GET(request: NextRequest) {
  try {
    // 1. Rate Limiting 체크
    const ip = request.headers.get('x-forwarded-for') || 
               request.headers.get('x-real-ip') || 
               'unknown';
    
    if (!checkRateLimit(ip)) {
      return NextResponse.json(
        { error: 'Too many requests. Please try again later.' },
        { status: 429 }
      );
    }

    // 2. URL 파라미터 가져오기
    const searchParams = request.nextUrl.searchParams;
    const imageUrl = searchParams.get('url');

    if (!imageUrl) {
      return NextResponse.json(
        { error: 'URL parameter is required' },
        { status: 400 }
      );
    }

    // 3. URL 검증 (SSRF 방지)
    if (!isValidFirebaseUrl(imageUrl)) {
      return NextResponse.json(
        { error: 'Invalid or unauthorized URL' },
        { status: 403 }
      );
    }

    // 4. Firebase Storage URL 가져오기 (timeout 설정)
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 10000); // 10초 timeout

    const response = await fetch(imageUrl, {
      signal: controller.signal,
      headers: {
        'User-Agent': 'Employee-Fee-Reduction-System/1.0'
      }
    });
    clearTimeout(timeoutId);
    
    if (!response.ok) {
      return NextResponse.json(
        { error: 'Failed to fetch image' },
        { status: response.status }
      );
    }

    // 5. Content-Type 검증
    const contentType = response.headers.get('content-type');
    if (!contentType || !ALLOWED_CONTENT_TYPES.some(type => contentType.includes(type))) {
      return NextResponse.json(
        { error: 'Invalid content type. Only images are allowed.' },
        { status: 400 }
      );
    }

    // 6. 파일 크기 제한 체크
    const contentLength = response.headers.get('content-length');
    if (contentLength && parseInt(contentLength) > MAX_FILE_SIZE) {
      return NextResponse.json(
        { error: 'File size exceeds limit (5MB)' },
        { status: 413 }
      );
    }

    // 7. 이미지 데이터를 ArrayBuffer로 가져오기 (크기 체크하면서)
    const arrayBuffer = await response.arrayBuffer();
    
    if (arrayBuffer.byteLength > MAX_FILE_SIZE) {
      return NextResponse.json(
        { error: 'File size exceeds limit (5MB)' },
        { status: 413 }
      );
    }
    
    const buffer = Buffer.from(arrayBuffer);

    // 8. 이미지를 Base64로 인코딩
    const base64 = buffer.toString('base64');
    const dataUrl = `data:${contentType};base64,${base64}`;

    return NextResponse.json({ dataUrl });
  } catch (error) {
    if (error instanceof Error) {
      if (error.name === 'AbortError') {
        return NextResponse.json(
          { error: 'Request timeout' },
          { status: 504 }
        );
      }
    }
    
    return NextResponse.json(
      { error: 'Internal server error' },
      { status: 500 }
    );
  }
}

