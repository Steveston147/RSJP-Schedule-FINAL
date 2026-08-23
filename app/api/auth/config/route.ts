import { NextResponse } from "next/server";
import { getPublicScheduleAuthConfig } from "@/lib/auth";

export async function GET() {
  return NextResponse.json(getPublicScheduleAuthConfig(), { headers: { "Cache-Control": "no-store" } });
}
