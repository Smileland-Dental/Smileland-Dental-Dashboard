import { NextResponse } from 'next/server';

export function middleware(request) {
  const authToken = request.cookies.get('firebaseAuthToken')?.value;

  const { pathname } = request.nextUrl;

  // Allow access to the login page
  if (pathname === '/') {
    return NextResponse.next();
  }

  // If there's no auth token and the user is not on the login page, redirect them
  if (!authToken) {
    return NextResponse.redirect(new URL('/', request.url));
  }

  return NextResponse.next();
}

export const config = {
  matcher: ['/((?!api|_next/static|_next/image|icon.ico).*)'],
};