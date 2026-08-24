import type { NextRequest } from "next/server";
import { NextResponse } from "next/server";
import { SCHEDULE_SESSION_COOKIE_NAME, verifyScheduleSessionToken } from "@/lib/auth";

export async function GET(request: NextRequest) {
  const token = request.cookies.get(SCHEDULE_SESSION_COOKIE_NAME)?.value ?? "";
  const session = token ? verifyScheduleSessionToken(token) : null;
  if (!session) {
    return NextResponse.json({ authenticated: false }, { status: 401, headers: { "Cache-Control": "no-store" } });
  }
  return NextResponse.json({ authenticated: true, user: session }, { headers: { "Cache-Control": "no-store" } });
}
