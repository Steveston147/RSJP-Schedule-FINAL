'use client';

const STORAGE_KEY = "rsjp_schedule_mvp_state_v2";

async function endSession() {
  await fetch("/api/auth/logout", { method: "POST" });
  window.location.replace("/login");
}

export default function AuthControls() {
  async function logout() {
    await endSession();
  }

  async function clearAndLogout() {
    const ok = window.confirm(
      "この端末に保存されているRSJP Scheduleの作業データを削除してログアウトします。必要なJSONバックアップを保存済みか確認してください。続行しますか？",
    );
    if (!ok) return;
    localStorage.removeItem(STORAGE_KEY);
    await endSession();
  }

  return (
    <div className="border-b border-slate-200 bg-white px-4 py-3 shadow-sm">
      <div className="mx-auto flex max-w-7xl flex-wrap items-center justify-between gap-3">
        <div>
          <p className="text-xs font-black tracking-[0.16em] text-rose-800">RSJP SCHEDULE / CONTROLLED PoC</p>
          <p className="text-xs text-slate-500">作業データはこのブラウザのlocalStorageに保存されます。</p>
        </div>
        <div className="flex flex-wrap gap-2">
          <button
            type="button"
            onClick={logout}
            className="rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm font-bold text-slate-700 hover:bg-slate-50"
          >
            ログアウト
          </button>
          <button
            type="button"
            onClick={clearAndLogout}
            className="rounded-lg border border-rose-300 bg-rose-50 px-3 py-2 text-sm font-bold text-rose-800 hover:bg-rose-100"
          >
            端末データを削除してログアウト
          </button>
        </div>
      </div>
    </div>
  );
}
