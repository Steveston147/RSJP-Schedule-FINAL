'use client';

import { FormEvent, useEffect, useState } from "react";

type AuthConfig = { mode: "demo" | "restricted"; configured: boolean; demoHint: string | null };

export default function LoginPage() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [config, setConfig] = useState<AuthConfig | null>(null);

  useEffect(() => {
    fetch("/api/auth/config", { cache: "no-store" })
      .then((response) => response.json())
      .then((value) => setConfig(value))
      .catch(() => setConfig(null));
  }, []);

  async function submit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setBusy(true);
    setError("");
    try {
      const response = await fetch("/api/auth/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email, password }),
      });
      const body = await response.json().catch(() => null) as { error?: string } | null;
      if (!response.ok) {
        setError(body?.error || "ログインできませんでした。");
        return;
      }
      window.location.replace("/");
    } catch {
      setError("通信に失敗しました。再度お試しください。");
    } finally {
      setBusy(false);
    }
  }

  return (
    <main className="flex min-h-screen items-center justify-center bg-slate-100 px-4 py-10 text-slate-900">
      <section className="w-full max-w-md rounded-2xl border border-slate-200 bg-white p-7 shadow-lg">
        <p className="text-xs font-black tracking-[0.18em] text-rose-800">ONESTOP AI PLATFORM</p>
        <h1 className="mt-2 text-3xl font-black">RSJP Schedule</h1>
        <p className="mt-2 text-sm leading-6 text-slate-600">短期受入プログラムのスケジュール作成・学内業務用</p>

        <form className="mt-7 space-y-5" onSubmit={submit}>
          <label className="block">
            <span className="mb-1.5 block text-sm font-bold">メールアドレス</span>
            <input
              type="email"
              autoComplete="email"
              required
              value={email}
              onChange={(event) => setEmail(event.target.value)}
              className="w-full rounded-xl border border-slate-300 px-4 py-3 outline-none ring-rose-700 transition focus:ring-2"
            />
          </label>

          <label className="block">
            <span className="mb-1.5 block text-sm font-bold">RSJP Schedule専用パスワード</span>
            <input
              type="password"
              autoComplete="current-password"
              required
              value={password}
              onChange={(event) => setPassword(event.target.value)}
              className="w-full rounded-xl border border-slate-300 px-4 py-3 outline-none ring-rose-700 transition focus:ring-2"
            />
          </label>

          {error && <p className="rounded-xl border border-rose-200 bg-rose-50 px-4 py-3 text-sm font-semibold text-rose-800">{error}</p>}

          <button
            type="submit"
            disabled={busy || config?.configured === false}
            className="w-full rounded-xl bg-rose-800 px-4 py-3 font-black text-white transition hover:bg-rose-900 disabled:cursor-not-allowed disabled:opacity-50"
          >
            {busy ? "確認中..." : "ログイン"}
          </button>
        </form>

        <div className="mt-5 rounded-xl border border-slate-200 bg-slate-50 p-4 text-xs leading-5 text-slate-600">
          <p className="font-bold text-slate-800">社内PoC認証</p>
          <p>Microsoft 365やGoogle等の既存パスワードは入力しないでください。RSJP Schedule専用のPoCパスワードのみ使用します。</p>
          {config?.demoHint && <p className="mt-2 font-semibold text-amber-800">{config.demoHint}</p>}
          {config && !config.configured && <p className="mt-2 font-semibold text-rose-800">Production用の認証設定が未完了です。</p>}
        </div>
      </section>
    </main>
  );
}
