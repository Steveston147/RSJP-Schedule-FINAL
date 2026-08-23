'use client';

import type { ReactNode } from "react";
import { useCallback, useEffect, useRef, useState } from "react";

type SessionResponse = {
  authenticated?: boolean;
  user?: { expiresAt?: number };
};

const REVALIDATE_INTERVAL_MS = 5 * 60 * 1000;

export default function SessionBoundary({ children }: { children: ReactNode }) {
  const [verified, setVerified] = useState(false);
  const expiryTimer = useRef<ReturnType<typeof setTimeout> | null>(null);

  const clearExpiryTimer = useCallback(() => {
    if (expiryTimer.current) clearTimeout(expiryTimer.current);
    expiryTimer.current = null;
  }, []);

  const redirectToLogin = useCallback(() => {
    clearExpiryTimer();
    setVerified(false);
    window.location.replace("/login");
  }, [clearExpiryTimer]);

  const verifySession = useCallback(async () => {
    try {
      const response = await fetch("/api/auth/session", {
        method: "GET",
        cache: "no-store",
        credentials: "same-origin",
      });
      if (!response.ok) {
        redirectToLogin();
        return;
      }

      const data = (await response.json()) as SessionResponse;
      const expiresAt = data.user?.expiresAt;
      if (!data.authenticated || typeof expiresAt !== "number") {
        redirectToLogin();
        return;
      }

      const remainingMs = expiresAt * 1000 - Date.now();
      if (remainingMs <= 0) {
        redirectToLogin();
        return;
      }

      clearExpiryTimer();
      expiryTimer.current = setTimeout(redirectToLogin, remainingMs + 250);
      setVerified(true);
    } catch {
      redirectToLogin();
    }
  }, [clearExpiryTimer, redirectToLogin]);

  useEffect(() => {
    void verifySession();
    const interval = window.setInterval(() => void verifySession(), REVALIDATE_INTERVAL_MS);
    const onFocus = () => void verifySession();
    const onVisibility = () => {
      if (document.visibilityState === "visible") void verifySession();
    };

    window.addEventListener("focus", onFocus);
    document.addEventListener("visibilitychange", onVisibility);
    return () => {
      window.clearInterval(interval);
      window.removeEventListener("focus", onFocus);
      document.removeEventListener("visibilitychange", onVisibility);
      clearExpiryTimer();
    };
  }, [clearExpiryTimer, verifySession]);

  if (!verified) {
    return (
      <main className="min-h-screen bg-slate-50 px-6 py-20 text-center text-sm font-semibold text-slate-600">
        認証状態を確認しています…
      </main>
    );
  }

  return <>{children}</>;
}
