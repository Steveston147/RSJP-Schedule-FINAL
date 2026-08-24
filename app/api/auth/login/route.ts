import type { NextRequest } from "next/server";
import { NextResponse } from "next/server";
import {
  createScheduleSessionToken,
  getPublicScheduleAuthConfig,
  SCHEDULE_SESSION_COOKIE_NAME,
  SCHEDULE_SESSION_TTL_SECONDS,
  verifyScheduleCredentials,
} from "@/lib/auth";
import { checkScheduleLoginRateLimit } from "@/lib/rateLimit";

function isSameOrigin(request: NextRequest) {
  const origin = request.headers.get("origin");
  return !origin || origin === request.nextUrl.origin;
}

function clientKey(request: NextRequest) {
  return (request.headers.get("x-forwarded-for")?.split(",")[0] || request.headers.get("x-real-ip") || "unknown").trim();
}

export async function POST(request: NextRequest) {
  if (!isSameOrigin(request)) {
    return NextResponse.json({ ok: false, error: "Invalid origin." }, { status: 403, headers: { "Cache-Control": "no-store" } });
  }

  const auth = getPublicScheduleAuthConfig();
  if (!auth.configured) {
    return NextResponse.json(
      { ok: false, error: "Production用の認証設定が未完了です。管理者に確認してください。" },
      { status: 503, headers: { "Cache-Control": "no-store" } },
    );
  }

  const body = (await request.json().catch(() => null)) as { email?: unknown; password?: unknown } | null;
  const email = typeof body?.email === "string" ? body.email : "";
  const password = typeof body?.password === "string" ? body.password : "";

  const limit = checkScheduleLoginRateLimit(clientKey(request), email);
  if (!limit.allowed) {
    return NextResponse.json(
      { ok: false, error: "ログイン試行が多すぎます。しばらくしてから再度お試しください。" },
      { status: 429, headers: { "Cache-Control": "no-store", "Retry-After": String(limit.retryAfter) } },
    );
  }

  const user = verifyScheduleCredentials(email, password);
  if (!user) {
    return NextResponse.json(
      { ok: false, error: "メールアドレスまたはパスワードを確認してください。" },
      { status: 401, headers: { "Cache-Control": "no-store" } },
    );
  }

  try {
    const token = createScheduleSessionToken(user);
    const response = NextResponse.json({ ok: true, user }, { headers: { "Cache-Control": "no-store" } });
    response.cookies.set(SCHEDULE_SESSION_COOKIE_NAME, token, {
      httpOnly: true,
      secure: process.env.NODE_ENV === "production",
      sameSite: "strict",
      path: "/",
      maxAge: SCHEDULE_SESSION_TTL_SECONDS,
    });
    return response;
  } catch {
    return NextResponse.json(
      { ok: false, error: "認証セッションを作成できませんでした。管理者に確認してください。" },
      { status: 503, headers: { "Cache-Control": "no-store" } },
    );
  }
}
