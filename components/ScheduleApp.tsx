'use client';

import React, { useEffect, useMemo, useRef, useState } from "react";

type ProgramType = "RSJP" | "Custom";

type Category =
  | "JapaneseClass"
  | "Orientation"
  | "Escort"
  | "CampusTour"
  | "Cultural"
  | "CompanyVisit"
  | "BuddyLunch"
  | "Ceremony"
  | "Other";

type TransportMode = "None" | "Bus" | "Walk" | "OnCampus";

type BusTripType = "OneWay" | "RoundTrip";

type YesNo = "Yes" | "No";

type ExportLang = "ja" | "en";

type Program = {
  id: string;
  name: string;
  type: ProgramType;
  studentsCount: number;
  startDate: string; // YYYY-MM-DD
  endDate: string;   // YYYY-MM-DD

  japanese: {
    enabled: YesNo;
    startTime: string;       // HH:MM
    lessonMinutes: number;   // e.g., 50
    breakMinutes: number;    // e.g., 10
    periods: number;         // 1..3
    classCount: number;      // 1..20
    classNames?: string[];   // optional class display names
    defaultTeacherRoom?: string;
  };

  // per-day override for Japanese class schedule
  japaneseOverrides: Record<string, JapaneseDayOverride>;

  lastUpdated: number;
};

type JapaneseDayOverride = {
  enabled: boolean;
  startTime: string;
  lessonMinutes: number;
  breakMinutes: number;
  periods: number;
  classCount: number;
  classrooms: string[];   // length can be < classCount, fallback handled
  teacherRooms: string[]; // optional multiple teachers
};

type ScheduleItem = {
  id: string;
  programId: string;
  date: string; // YYYY-MM-DD
  startTime: string; // HH:MM
  endTime: string; // HH:MM
  category: Category;
  title: string;
  titleEn?: string; // 追加：英語タイトル（任意）
  location: string;
  locationEn?: string; // 追加：英語場所（任意）
  roomNeeded: YesNo;
  studentsCount: number;
  buddyCount: number;
  kvhRequired: YesNo;
  kvhCount: number;
  transportMode: TransportMode;
  busCompany: string;
  busCount: number;
  busTripType: BusTripType;
  busPickup: string;
  busDropoff: string;
  arrangementsNeeded: YesNo;
  notes: string;
  notesEn?: string; // 追加：英語備考（任意）

  // 日本語講座用（CSVで見分けやすくする）
  classIndex?: number; // 1..classCount
  periodIndex?: number; // 1..periods
  generated?: boolean;
  generatedKind?: "Auto";
};

type AppState = {
  programs: Program[];
  items: ScheduleItem[];
  selectedProgramId: string | null;
  lastUpdated: number;
};

const STORAGE_KEY = "rsjp_schedule_mvp_state_v2";

function isValidImportedState(value: unknown): value is AppState {
  if (!value || typeof value !== "object") return false;
  const candidate = value as Partial<AppState>;
  if (!Array.isArray(candidate.programs) || !Array.isArray(candidate.items)) return false;

  const programIds = new Set<string>();
  for (const program of candidate.programs) {
    if (!program || typeof program.id !== "string" || !program.id.trim() || typeof program.name !== "string") return false;
    if (programIds.has(program.id)) return false;
    programIds.add(program.id);
  }

  for (const item of candidate.items) {
    if (!item || typeof item.id !== "string" || !item.id.trim() || typeof item.programId !== "string" || !programIds.has(item.programId)) return false;
    if (typeof item.date !== "string" || typeof item.startTime !== "string" || typeof item.endTime !== "string") return false;
  }

  return true;
}

function uid() {
  return Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
}

function pad2(n: number) {
  return String(n).padStart(2, "0");
}

function toDateObj(dateISO: string) {
  const [y, m, d] = dateISO.split("-").map((s) => Number(s));
  return new Date(Date.UTC(y, m - 1, d));
}

function addDaysISO(dateISO: string, days: number) {
  const dt = toDateObj(dateISO);
  dt.setUTCDate(dt.getUTCDate() + days);
  return `${dt.getUTCFullYear()}-${pad2(dt.getUTCMonth() + 1)}-${pad2(dt.getUTCDate())}`;
}

function isWeekendISO(dateISO: string) {
  const dt = toDateObj(dateISO);
  const dow = dt.getUTCDay(); // 0 Sun .. 6 Sat
  return dow === 0 || dow === 6;
}

function dayOfWeekJP(dateISO: string) {
  const dt = toDateObj(dateISO);
  const names = ["日", "月", "火", "水", "木", "金", "土"];
  return names[dt.getUTCDay()] ?? "";
}

function dayOfWeekEN(dateISO: string) {
  const d = toDateObj(dateISO);
  const names = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  return names[d.getUTCDay()] ?? "";
}

function clamp(n: number, min: number, max: number) {
  return Math.max(min, Math.min(max, n));
}

function escapeHTML(s: string) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function csvEscape(s: string) {
  const v = String(s ?? "");
  if (/[",\n\r]/.test(v)) return `"${v.replaceAll('"', '""')}"`;
  return v;
}

function downloadTextFile(filename: string, text: string, mime = "text/plain;charset=utf-8") {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function programTypeLabel(type: ProgramType) {
  return type === "RSJP" ? "RSJP（レギュラー）" : "カスタム";
}

function categoryLabelJP(cat: Category) {
  switch (cat) {
    case "JapaneseClass": return "日本語講座";
    case "Orientation": return "オリエン";
    case "Escort": return "引率";
    case "CampusTour": return "キャンパスツアー";
    case "Cultural": return "文化体験";
    case "CompanyVisit": return "企業訪問";
    case "BuddyLunch": return "バディランチ";
    case "Ceremony": return "修了式";
    default: return "その他";
  }
}

function categoryLabelEN(cat: Category) {
  switch (cat) {
    case "JapaneseClass": return "Japanese Class";
    case "Orientation": return "Orientation";
    case "Escort": return "Escort";
    case "CampusTour": return "Campus Tour";
    case "Cultural": return "Cultural Experience";
    case "CompanyVisit": return "Company Visit";
    case "BuddyLunch": return "Buddy Lunch";
    case "Ceremony": return "Completion Ceremony";
    default: return "Other";
  }
}

function categoryLabelByLang(cat: Category, lang: ExportLang) {
  return lang === "en" ? categoryLabelEN(cat) : categoryLabelJP(cat);
}

function defaultClassNameFor(i: number) {
  return `クラス${i}`;
}

function makeNewProgram(): Program {
  const id = uid();
  const today = new Date();
  const y = today.getFullYear();
  const m = today.getMonth() + 1;
  const d = today.getDate();
  const start = `${y}-${pad2(m)}-${pad2(d)}`;
  const end = start;
  return {
    id,
    name: "新規プログラム",
    type: "RSJP",
    studentsCount: 20,
    startDate: start,
    endDate: end,
    japanese: {
      enabled: "Yes",
      startTime: "09:00",
      lessonMinutes: 50,
      breakMinutes: 10,
      periods: 2,
      classCount: 4,
      classNames: ["A", "B", "C", "D"].map((x) => `クラス${x}`),
      defaultTeacherRoom: "",
    },
    japaneseOverrides: {},
    lastUpdated: Date.now(),
  };
}

function sortItems(items: ScheduleItem[]) {
  return [...items].sort((a, b) => {
    const ka = `${a.date} ${a.startTime} ${a.endTime} ${a.category} ${a.title}`;
    const kb = `${b.date} ${b.startTime} ${b.endTime} ${b.category} ${b.title}`;
    return ka < kb ? -1 : ka > kb ? 1 : 0;
  });
}

function makeItemBase(program: Program, date: string): Omit<ScheduleItem, "id"> {
  return {
    programId: program.id,
    date,
    startTime: "09:00",
    endTime: "10:00",
    category: "Other",
    title: "",
    titleEn: "",
    location: "",
    locationEn: "",
    roomNeeded: "No",
    studentsCount: program.studentsCount,
    buddyCount: 0,
    kvhRequired: "No",
    kvhCount: 0,
    transportMode: "None",
    busCompany: "",
    busCount: 1,
    busTripType: "OneWay",
    busPickup: "",
    busDropoff: "",
    arrangementsNeeded: "No",
    notes: "",
    notesEn: "",
    generated: false,
  };
}

function removeGeneratedAuto(items: ScheduleItem[], programId: string) {
  return items.filter((it) => !(it.programId === programId && it.generated && it.generatedKind === "Auto"));
}

function dedupeItemsForProgram(items: ScheduleItem[], programId: string) {
  // 同一プログラム内で「日付・時間・カテゴリ・タイトル・場所・メモ」が完全一致する重複を削除
  const seen = new Set<string>();
  const out: ScheduleItem[] = [];
  for (const it of items) {
    if (it.programId !== programId) {
      out.push(it);
      continue;
    }
    const key = [
      it.programId,
      it.date,
      it.startTime,
      it.endTime,
      it.category,
      it.title,
      it.titleEn ?? "",
      it.location,
      it.locationEn ?? "",
      it.notes,
      it.notesEn ?? "",
    ].join("|");
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(it);
  }
  return out;
}

function buildAutoItems(program: Program): ScheduleItem[] {
  const items: ScheduleItem[] = [];

  // 1) Orientation on first day
  const firstDay = program.startDate;
  items.push({
    id: uid(),
    ...makeItemBase(program, firstDay),
    category: "Escort",
    title: "引率",
    startTime: "09:00",
    endTime: "10:00",
    generated: true,
    generatedKind: "Auto",
  });

  items.push({
    id: uid(),
    ...makeItemBase(program, firstDay),
    category: "Orientation",
    title: "Orientation",
    startTime: "10:30",
    endTime: "11:30",
    generated: true,
    generatedKind: "Auto",
  });

  items.push({
    id: uid(),
    ...makeItemBase(program, firstDay),
    category: "CampusTour",
    title: "Campus Tour",
    startTime: "11:30",
    endTime: "12:20",
    generated: true,
    generatedKind: "Auto",
  });

  items.push({
    id: uid(),
    ...makeItemBase(program, program.endDate),
    category: "Ceremony",
    title: "Completion Ceremony",
    startTime: "13:10",
    endTime: "14:10",
    generated: true,
    generatedKind: "Auto",
  });

  // 2) Japanese class for weekdays between start/end if enabled
  const isEnabled = program.japanese.enabled === "Yes";
  if (isEnabled) {
    let d = program.startDate;
    while (d <= program.endDate) {
      const ov = program.japaneseOverrides[d];
      const enabled = ov ? ov.enabled : true;

      const isWeekend = isWeekendISO(d);
      if (!isWeekend && enabled) {
        const startTime = ov ? ov.startTime : program.japanese.startTime;
        const lesson = ov ? ov.lessonMinutes : program.japanese.lessonMinutes;
        const brk = ov ? ov.breakMinutes : program.japanese.breakMinutes;
        const periods = ov ? ov.periods : program.japanese.periods;
        const classCount = ov ? ov.classCount : program.japanese.classCount;
        const classrooms = ov ? ov.classrooms : [];
        const teacherRooms = ov ? ov.teacherRooms : [];

        // build periods
        const [h0, m0] = startTime.split(":").map((s) => Number(s));
        let t = h0 * 60 + m0;

        for (let p = 1; p <= periods; p++) {
          const pStart = `${pad2(Math.floor(t / 60))}:${pad2(t % 60)}`;
          t += lesson;
          const pEnd = `${pad2(Math.floor(t / 60))}:${pad2(t % 60)}`;
          t += brk;

          for (let c = 1; c <= classCount; c++) {
            const room = classrooms[c - 1] ?? "";
            const teacher = teacherRooms[0] ?? "";
            const className = (program.japanese.classNames?.[c - 1] ?? defaultClassNameFor(c)).trim() || defaultClassNameFor(c);

            const title = `日本語講座 ${className}（${p}限）`;
            const notes = teacher ? `講師控室: ${teacher}` : "";

            items.push({
              id: uid(),
              ...makeItemBase(program, d),
              category: "JapaneseClass",
              title,
              location: room,
              notes,
              startTime: pStart,
              endTime: pEnd,
              classIndex: c,
              periodIndex: p,
              generated: true,
              generatedKind: "Auto",
            });
          }
        }
      }

      d = addDaysISO(d, 1);
    }
  }

  return items;
}

function ynLabel(v: any, lang: ExportLang) {
  const s = String(v ?? "");
  if (lang === "en") return s; // already "Yes"/"No"
  if (s === "Yes") return "あり";
  if (s === "No") return "なし";
  return s;
}

function simpleExportTranslate(text: string, lang: ExportLang) {
  const s = String(text ?? "");
  if (lang !== "en") return s;
  // very small, safe replacements (fallback keeps original JP)
  const repl: Array<[RegExp, string]> = [
    [/日本語講座/g, "Japanese Class"],
    [/文化体験/g, "Cultural Experience"],
    [/企業訪問/g, "Company Visit"],
    [/キャンパスツアー/g, "Campus Tour"],
    [/オリエン(テーション)?/g, "Orientation"],
    [/引率/g, "Escort"],
    [/修了式/g, "Completion Ceremony"],
    [/バディランチ/g, "Buddy Lunch"],
    [/バディ/g, "Buddy"],
    [/留学生/g, "Students"],
    [/備考/g, "Notes"],

    // cultural / places
    [/マンガミュージアム/g, "Manga Museum"],
    [/漢字ミュージアム/g, "Kanji Museum"],
    [/京友禅/g, "Yuzen Dyeing"],
    [/二条城/g, "Nijo Castle"],
    [/金閣寺/g, "Golden Pavilion"],
    [/茶道/g, "Tea Ceremony"],
    [/書道/g, "Calligraphy"],
    [/和太鼓/g, "Wadaiko Drumming"],
    [/和食/g, "Japanese Cuisine"],

    // common words
    [/ホテル/g, "Hotel"],
    [/寮/g, "Dorm"],
    [/大学/g, "University"],
    [/集合場所を入力/g, "enter meeting point"],
    [/集合/g, "Meeting point"],
    [/教室を入力/g, "enter classroom"],
    [/教室/g, "Classroom"],
    [/講師控室を入力/g, "enter teacher room"],
    [/講師控室/g, "Teacher room"],

    // punctuation
    [/（/g, "("],
    [/）/g, ")"],
    [/：/g, ":"],
    [/→/g, " to "],
  ];
  let out = s;
  for (const [rx, rep] of repl) {
    out = out.replace(rx, rep);
  }
  return out;
}

function pickExportText(
  it: any,
  field: "title" | "location" | "notes",
  lang: ExportLang
) {
  const base = String(it?.[field] ?? "");
  if (lang !== "en") return base;

  const enKey =
    field === "title" ? "titleEn" : field === "location" ? "locationEn" : "notesEn";

  const en = String(it?.[enKey] ?? "").trim();
  if (en) return en;

  return simpleExportTranslate(base, lang);
}

function normalizeHHMM(v: string) {
  const s = String(v ?? "").trim();
  const m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return "09:00";
  const hh = clamp(Number(m[1]), 0, 23);
  const mm = clamp(Number(m[2]), 0, 59);
  return `${pad2(hh)}:${pad2(mm)}`;
}

function parseISODateInput(v: string) {
  const s = String(v ?? "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return null;
  return s;
}

function icsEscape(s: string) {
  return String(s ?? "")
    .replaceAll("\\", "\\\\")
    .replaceAll(";", "\\;")
    .replaceAll(",", "\\,")
    .replaceAll("\n", "\\n");
}

function toICSDateTime(dateISO: string, hhmm: string) {
  const [y, m, d] = dateISO.split("-").map((x) => Number(x));
  const [hh, mm] = hhmm.split(":").map((x) => Number(x));
  // keep as local time for Asia/Tokyo in DTSTART;TZID
  const dt = new Date(Date.UTC(y, m - 1, d, hh, mm, 0));
  return `${dt.getUTCFullYear()}${pad2(dt.getUTCMonth() + 1)}${pad2(dt.getUTCDate())}T${pad2(dt.getUTCHours())}${pad2(dt.getUTCMinutes())}00`;
}

function buildICSAsiaTokyo(program: any, items: any[], lang: ExportLang = "ja") {
  const lines: string[] = [];
  lines.push("BEGIN:VCALENDAR");
  lines.push("VERSION:2.0");
  lines.push("PRODID:-//RSJP Scheduler//JP//EN");
  lines.push("CALSCALE:GREGORIAN");
  lines.push("METHOD:PUBLISH");

  const programItems = items.filter((x) => x.programId === program.id);

  for (const it of sortItems(programItems)) {
    const uidLine = `${program.id}-${it.id}@rsjp`;
    const dtstamp = toICSDateTime(program.startDate, "00:00");
    const dtStart = toICSDateTime(it.date, it.startTime);
    const dtEnd = toICSDateTime(it.date, it.endTime);

    const titleFor = pickExportText(it, "title", lang);
    const summary = lang === "en" ? `${program.name} | ${titleFor}` : `${program.name}｜${String(it.title ?? "")}`;
    const location = pickExportText(it, "location", lang);

    const descParts: string[] = [];
    descParts.push(`${lang === "en" ? "Category" : "カテゴリ"}: ${categoryLabelByLang(it.category as Category, lang)}`);
    descParts.push(`${lang === "en" ? "Time" : "時間"}: ${it.startTime}-${it.endTime}`);
    if (it.transportMode && it.transportMode !== "None") descParts.push(`${lang === "en" ? "Transport" : "移動"}: ${it.transportMode}`);
    if (it.busCompany) descParts.push(`${lang === "en" ? "Bus company" : "バス会社"}: ${it.busCompany}`);
    if (it.busCount) descParts.push(`${lang === "en" ? "Buses" : "台数"}: ${it.busCount}`);
    if (it.busPickup) descParts.push(`${lang === "en" ? "Pickup" : "乗車"}: ${simpleExportTranslate(String(it.busPickup ?? ""), lang)}`);
    if (it.busDropoff) descParts.push(`${lang === "en" ? "Dropoff" : "降車"}: ${simpleExportTranslate(String(it.busDropoff ?? ""), lang)}`);
    descParts.push(`${lang === "en" ? "Room needed" : "教室"}: ${ynLabel(it.roomNeeded, lang)}`);
    descParts.push(`${lang === "en" ? "Students" : "留学生"}: ${it.studentsCount ?? program.studentsCount}`);
    descParts.push(`${lang === "en" ? "Buddy" : "バディ"}: ${it.buddyCount ?? 0}`);
    descParts.push(`${lang === "en" ? "KVH" : "KVH"}: ${ynLabel(it.kvhRequired, lang)} (${it.kvhCount ?? 0})`);
    descParts.push(`${lang === "en" ? "Arrangements" : "手配"}: ${ynLabel(it.arrangementsNeeded, lang)}`);

    const notesOut = pickExportText(it, "notes", lang);
    if (notesOut) descParts.push(`${lang === "en" ? "Notes" : "備考"}: ${notesOut}`);

    lines.push("BEGIN:VEVENT");
    lines.push(`UID:${icsEscape(uidLine)}`);
    lines.push(`DTSTAMP:${dtstamp}`);
    lines.push(`DTSTART;TZID=Asia/Tokyo:${dtStart}`);
    lines.push(`DTEND;TZID=Asia/Tokyo:${dtEnd}`);
    lines.push(`SUMMARY:${icsEscape(summary)}`);
    if (location) lines.push(`LOCATION:${icsEscape(location)}`);
    lines.push(`DESCRIPTION:${icsEscape(descParts.join("\n"))}`);
    lines.push("END:VEVENT");
  }

  lines.push("END:VCALENDAR");
  return lines.join("\r\n");
}

function buildCSV(program: Program, items: ScheduleItem[], lang: ExportLang) {
  const headers = [
    "ProgramName",
    "ProgramType",
    "Date",
    "DayOfWeek",
    "StartTime",
    "EndTime",
    "Category",
    "Title",
    "Location",
    "RoomNeeded",
    "StudentsCount",
    "BuddyCount",
    "KVHRequired",
    "KVHCount",
    "TransportMode",
    "BusCompany",
    "BusCount",
    "BusTripType",
    "BusPickup",
    "BusDropoff",
    "ArrangementsNeeded",
    "Notes",
    "Generated",
  ];

  const rows = [...items]
    .filter((it) => it.programId === program.id)
    .sort((a, b) => {
      const ka = `${a.date} ${a.startTime} ${a.endTime} ${a.category} ${a.title}`;
      const kb = `${b.date} ${b.startTime} ${b.endTime} ${b.category} ${b.title}`;
      return ka < kb ? -1 : ka > kb ? 1 : 0;
    })
    .map((it) => {
      const row: Record<string, string> = {
        ProgramName: program.name,
        ProgramType: program.type,
        Date: it.date,
        DayOfWeek: lang === "en" ? dayOfWeekEN(it.date) : dayOfWeekJP(it.date),
        StartTime: it.startTime,
        EndTime: it.endTime,
        Category: it.category,
        Title: pickExportText(it, "title", lang),
        Location: pickExportText(it, "location", lang),
        RoomNeeded: it.roomNeeded,
        StudentsCount: String(it.studentsCount ?? program.studentsCount),
        BuddyCount: String(it.buddyCount ?? 0),
        KVHRequired: it.kvhRequired,
        KVHCount: String(it.kvhCount ?? 0),
        TransportMode: it.transportMode,
        BusCompany: it.busCompany,
        BusCount: String(it.busCount ?? 0),
        BusTripType: it.busTripType,
        BusPickup: simpleExportTranslate(String(it.busPickup ?? ""), lang),
        BusDropoff: simpleExportTranslate(String(it.busDropoff ?? ""), lang),
        ArrangementsNeeded: it.arrangementsNeeded,
        Notes: pickExportText(it, "notes", lang),
        Generated: it.generated ? "Yes" : "No",
      };
      return headers.map((h) => csvEscape(row[h] ?? "")).join(",");
    });

  return [headers.join(","), ...rows].join("\n");
}

function buildCalendarHTMLMultiMonth(program: Program, items: any[], lang: ExportLang = "ja") {
  // map date -> items
  const map = new Map<string, any[]>();
  for (const it of items.filter((x) => x.programId === program.id)) {
    if (!map.has(it.date)) map.set(it.date, []);
    map.get(it.date)!.push(it);
  }
  for (const [k, v] of map.entries()) {
    v.sort((a, b) => {
      const ka = `${a.startTime} ${a.endTime} ${a.category} ${a.title}`;
      const kb = `${b.startTime} ${b.endTime} ${b.category} ${b.title}`;
      return ka < kb ? -1 : ka > kb ? 1 : 0;
    });
  }

  const start = toDateObj(program.startDate);
  const end = toDateObj(program.endDate);

  const startY = start.getUTCFullYear();
  const startM = start.getUTCMonth(); // 0-based
  const endY = end.getUTCFullYear();
  const endM = end.getUTCMonth();

  const monthSections: string[] = [];

  let y = startY;
  let m = startM;

  const dowHeaders = lang === "en"
    ? `<tr>
      <th>Sun</th><th>Mon</th><th>Tue</th><th>Wed</th><th>Thu</th><th>Fri</th><th>Sat</th>
    </tr>`
    : `<tr>
      <th>日</th><th>月</th><th>火</th><th>水</th><th>木</th><th>金</th><th>土</th>
    </tr>`;

  // --- helpers for compact day rendering ---
  const isJapaneseLessonItem = (it: any) =>
    String(it?.category ?? "") === "JapaneseClass" || String(it?.title ?? "").includes("日本語講座");

  const parseClassNo = (it: any): number | null => {
    if (typeof it?.classIndex === "number" && Number.isFinite(it.classIndex)) return it.classIndex;
    const t = String(it?.title ?? "");
    const m = t.match(/クラス\s*([0-9０-９]+)/);
    if (m?.[1]) {
      const n = Number(m[1].replace(/[０-９]/g, (c: string) => String("０１２３４５６７８９".indexOf(c))));
      if (Number.isFinite(n) && n > 0) return n;
    }
    return null;
  };

  const pickTeacherRoomLabel = (jpItems: any[]): string => {
    const set = new Set<string>();

    // 1) notes: "講師控室: YY310"
    for (const it of jpItems) {
      const note = String(it?.notes ?? "");
      const m = note.match(/講師控室\s*:\s*(.+)$/);
      if (m?.[1]) {
        const v = String(m[1]).trim();
        if (v) set.add(v);
      }
    }

    // 2) program default
    const def = String(program?.japanese?.defaultTeacherRoom ?? "").trim();
    if (def) set.add(def);

    const list = Array.from(set).filter(Boolean);
    if (list.length === 0) return "";
    if (list.length === 1) return lang === "en" ? `Teacher (${escapeHTML(list[0])})` : `講師（${escapeHTML(list[0])}）`;
    return lang === "en" ? `Teacher (${escapeHTML(list[0])} +${list.length - 1})` : `講師（${escapeHTML(list[0])} 他${list.length - 1}）`;
  };

  const renderDayItemsCompact = (dISO: string, list: any[]): string => {
    const jp = list.filter(isJapaneseLessonItem);
    const other = list.filter((it) => !isJapaneseLessonItem(it));

    // other: keep short lines
    const otherLines = other.slice(0, 10).map((it) => {
      const cat = categoryLabelByLang(it.category as Category, lang);
      const time = `${it.startTime}-${it.endTime}`;
      const title = escapeHTML(pickExportText(it, "title", lang));
      const loc = pickExportText(it, "location", lang).trim();
      const locationLabelFor = (catRaw: any, locRaw: string) => {
        const cat = String(catRaw ?? "");
        const loc = String(locRaw ?? "").trim();
        if (!loc) return "";
        // 必須系はラベルを付けて見やすく
        if (cat === "Escort" || cat === "CampusTour" || cat === "Cultural" || cat === "CompanyVisit") return lang === "en" ? `Meet-up: ${loc}` : `集合: ${loc}`;
        if (cat === "Orientation" || cat === "JapaneseClass" || cat === "Ceremony") return lang === "en" ? `Room: ${loc}` : `教室: ${loc}`;
        return lang === "en" ? `Place: ${loc}` : `場所: ${loc}`;
      };

      // 旧：locLabel（互換のため残すが、表示はlocLineに移行）
      const locLabel = loc ? `（${escapeHTML(loc)}）` : "";
      const locLine = loc ? `<div class="sub">${escapeHTML(locationLabelFor(it.category, loc))}</div>` : "";
      const noteText = pickExportText(it, "notes", lang).trim();
      const noteLine = noteText ? `<div class="sub muted">${escapeHTML(noteText)}</div>` : "";
      return `<div class="it"><div><span class="t">${escapeHTML(cat)}</span> <span class="tm">${escapeHTML(time)}</span></div><div>${title}</div>${locLine}${noteLine}</div>`;
    });

    const more = other.length > 10 ? `<div class="more">…他 ${other.length - 10} 件</div>` : "";

    // jp: 1 block + class lines + teacher room
    let jpBlock = "";
    if (jp.length > 0) {
      // min start / max end
      let minStart = jp[0].startTime;
      let maxEnd = jp[0].endTime;
      for (const it of jp) {
        if (it.startTime && it.startTime < minStart) minStart = it.startTime;
        if (it.endTime && it.endTime > maxEnd) maxEnd = it.endTime;
      }

      const classRoom = new Map<number, string>();
      for (const it of jp) {
        const cno = parseClassNo(it);
        if (!cno) continue;
        const room = String(it.location ?? "").trim();
        if (!classRoom.has(cno)) classRoom.set(cno, room);
      }

      const classNos = Array.from(classRoom.keys()).sort((a, b) => a - b);
      const classLines = classNos.slice(0, 20).map((cno) => {
        const room = classRoom.get(cno) ?? "";
        const roomLabel = room ? `（${escapeHTML(room)}）` : "";
        const name = (program.japanese.classNames?.[cno - 1] ?? defaultClassNameFor(cno)).trim() || defaultClassNameFor(cno);
        return `<div class="sub">${escapeHTML(name)}${roomLabel}</div>`;
      });

      const teacher = pickTeacherRoomLabel(jp);
      const teacherLine = teacher ? `<div class="sub">${teacher}</div>` : "";

      jpBlock = `
          <div class="jp">
            <div class="it jphead"><span class="t">${escapeHTML(lang === "en" ? "Japanese Class" : "日本語講座")}</span> <span class="tm">${escapeHTML(minStart)}-${escapeHTML(maxEnd)}</span></div>
            ${classLines.join("")}
            ${teacherLine}
          </div>
        `;
    }

    return `${jpBlock}${otherLines.join("")}${more}`;
  };

  const cell = (dISO: string | null) => {
    if (!dISO) return `<td class="empty"></td>`;
    const list = map.get(dISO) ?? [];
    const inRange = dISO >= program.startDate && dISO <= program.endDate;

    const body = renderDayItemsCompact(dISO, list);

    return `
        <td class="${inRange ? "inrange" : "outrange"}">
          <div class="d">${Number(dISO.slice(-2))}</div>
          <div class="list">${body}</div>
        </td>`;
  };

  while (y < endY || (y === endY && m <= endM)) {
    const first = new Date(Date.UTC(y, m, 1));
    const last = new Date(Date.UTC(y, m + 1, 0));
    const firstDow = first.getUTCDay();
    const daysInMonth = last.getUTCDate();

    const monthLabel = `${y}-${pad2(m + 1)}`;

    const weeks: string[] = [];
    let day = 1;

    for (let r = 0; r < 6; r++) {
      const tds: string[] = [];
      for (let c = 0; c < 7; c++) {
        const idx = r * 7 + c;
        if (idx < firstDow || day > daysInMonth) {
          tds.push(cell(null));
        } else {
          const dISO = `${y}-${pad2(m + 1)}-${pad2(day)}`;
          tds.push(cell(dISO));
          day++;
        }
      }
      weeks.push(`<tr>${tds.join("")}</tr>`);
      if (day > daysInMonth) break;
    }

    monthSections.push(`
        <section class="month">
          <h2>${escapeHTML(program.name)} ${monthLabel}</h2>
          <div class="meta">${lang === "en" ? `Period: ${program.startDate} → ${program.endDate} / Students: ${program.studentsCount}` : `期間: ${program.startDate} → ${program.endDate} / 留学生: ${program.studentsCount}`}</div>
          <table>
            <thead>${dowHeaders}</thead>
            <tbody>${weeks.join("\n")}</tbody>
          </table>
        </section>
      `);

    // next month
    m++;
    if (m >= 12) {
      m = 0;
      y++;
    }
  }

  return `<!doctype html>
<html lang="${lang === "en" ? "en" : "ja"}">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>${escapeHTML(program.name)} ${lang === "en" ? "Calendar" : "カレンダー"}</title>
<style>
  body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; padding: 16px; }
  h1 { font-size: 18px; margin: 0 0 8px; }
  h2 { font-size: 16px; margin: 0 0 6px; }
  .topmeta { font-size: 12px; opacity: .75; margin-bottom: 14px; }
  .meta { font-size: 12px; opacity: .75; margin: 0 0 10px; }
  table { width: 100%; border-collapse: collapse; table-layout: fixed; }
  th, td { border: 1px solid #ddd; vertical-align: top; padding: 6px; }
  th { background: #f7f7f7; font-size: 12px; }
  td { height: 120px; }
  td.empty { background: #fafafa; }
  td.outrange { background: #fcfcfc; opacity: .55; }
  .d { font-weight: 800; font-size: 12px; margin-bottom: 6px; }
  .list { font-size: 11px; line-height: 1.25; }
  .it { margin-bottom: 4px; }
  .t { font-weight: 800; }
  .tm { font-weight: 800; margin-left: 2px; }
  .more { font-size: 11px; opacity: .7; }

  .jp { margin-bottom: 6px; padding-bottom: 4px; border-bottom: 1px dotted #ddd; }
  .jphead { margin-bottom: 4px; }
  .sub { margin-left: 10px; margin-bottom: 3px; }
  .muted { opacity: .7; }

  .month { margin-bottom: 18px; }
  @media print {
    body { padding: 0; }
    td { height: 110px; }
    .month { page-break-after: always; }
  }
</style>
</head>
<body>
  <h1>${escapeHTML(program.name)} (${escapeHTML(programTypeLabel(program.type))}) ${lang === "en" ? "Calendar" : "カレンダー"}</h1>
  <div class="topmeta">${lang === "en" ? `Period: ${program.startDate} → ${program.endDate} / Students: ${program.studentsCount}` : `期間: ${program.startDate} → ${program.endDate} / 留学生: ${program.studentsCount}`}</div>
  ${monthSections.join("\n")}
</body>
</html>`;
}

export default function ScheduleApp() {
  const [state, setState] = useState<AppState>(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) {
        const p = makeNewProgram();
        return { programs: [p], items: [], selectedProgramId: p.id, lastUpdated: Date.now() };
      }
      const parsed = JSON.parse(raw) as AppState;
      if (!parsed || !Array.isArray(parsed.programs) || !Array.isArray(parsed.items)) {
        const p = makeNewProgram();
        return { programs: [p], items: [], selectedProgramId: p.id, lastUpdated: Date.now() };
      }
      // ensure at least 1 program
      if (parsed.programs.length === 0) {
        const p = makeNewProgram();
        return { programs: [p], items: [], selectedProgramId: p.id, lastUpdated: Date.now() };
      }
      return parsed;
    } catch {
      const p = makeNewProgram();
      return { programs: [p], items: [], selectedProgramId: p.id, lastUpdated: Date.now() };
    }
  });

  const [activeTab, setActiveTab] = useState<"program" | "calendar" | "share">("program");
  const [exportLang, setExportLang] = useState<ExportLang>("ja");
  const [showEnglishFields, setShowEnglishFields] = useState(false);
  const [programQuery, setProgramQuery] = useState("");
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  // calendar add form
  const [newDate, setNewDate] = useState<string>("");
  const [newStartTime, setNewStartTime] = useState("09:00");
  const [newEndTime, setNewEndTime] = useState("10:00");
  const [newCategory, setNewCategory] = useState<Category>("Other");
  const [newTitle, setNewTitle] = useState("");
  const [newLocation, setNewLocation] = useState("");
  const [roomNeeded, setRoomNeeded] = useState<YesNo>("No");
  const [studentsCount, setStudentsCount] = useState<number>(20);
  const [buddyCount, setBuddyCount] = useState<number>(0);
  const [kvhRequired, setKvhRequired] = useState<YesNo>("No");
  const [kvhCount, setKvhCount] = useState<number>(0);
  const [transportMode, setTransportMode] = useState<TransportMode>("None");
  const [busCompany, setBusCompany] = useState("");
  const [busCount, setBusCount] = useState<number>(1);
  const [busTripType, setBusTripType] = useState<BusTripType>("OneWay");
  const [busPickup, setBusPickup] = useState("");
  const [busDropoff, setBusDropoff] = useState("");
  const [arrangementsNeeded, setArrangementsNeeded] = useState<YesNo>("No");
  const [notes, setNotes] = useState("");

  // override UI
  const [overrideDate, setOverrideDate] = useState<string>("");
  const [ovEnabled, setOvEnabled] = useState<boolean>(true);
  const [ovStartTime, setOvStartTime] = useState<string>("09:00");
  const [ovLesson, setOvLesson] = useState<number>(50);
  const [ovBreak, setOvBreak] = useState<number>(10);
  const [ovPeriods, setOvPeriods] = useState<number>(2);
  const [ovClassCount, setOvClassCount] = useState<number>(4);
  const [ovClassrooms, setOvClassrooms] = useState<string>("");
  const [ovTeacherRooms, setOvTeacherRooms] = useState<string>("");

  // persist
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } catch {
      // ignore
    }
  }, [state]);

  const selectedProgram = useMemo(() => {
    const id = state.selectedProgramId ?? state.programs[0]?.id ?? null;
    return state.programs.find((p) => p.id === id) ?? null;
  }, [state.programs, state.selectedProgramId]);

  const selectedProgramItems = useMemo(() => {
    if (!selectedProgram) return [];
    return sortItems(state.items.filter((it) => it.programId === selectedProgram.id));
  }, [state.items, selectedProgram?.id]);

  // when selected program changes, sync studentsCount default in add-form
  useEffect(() => {
    if (!selectedProgram) return;
    setStudentsCount(selectedProgram.studentsCount);
    setNewDate(selectedProgram.startDate);
  }, [selectedProgram?.id]);

  // set override default date
  useEffect(() => {
    if (!selectedProgram) return;
    setOverrideDate(selectedProgram.startDate);
    loadOverrideFromProgram(selectedProgram.startDate);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedProgram?.id]);

  function updateProgram(partial: Partial<Program>) {
    if (!selectedProgram) return;
    const now = Date.now();
    const updated: Program = { ...selectedProgram, ...partial, lastUpdated: now };
    setState((prev) => ({
      ...prev,
      programs: prev.programs.map((p) => (p.id === updated.id ? updated : p)),
      lastUpdated: now,
    }));
  }

  function selectProgram(id: string) {
    setState((prev) => ({ ...prev, selectedProgramId: id }));
  }

  function addProgram() {
    const p = makeNewProgram();
    setState((prev) => ({
      ...prev,
      programs: [p, ...prev.programs],
      selectedProgramId: p.id,
      lastUpdated: Date.now(),
    }));
    setActiveTab("program");
  }

  function deleteProgram(programId: string) {
    const ok = window.confirm("このプログラムを削除します。予定もすべて削除されます。よろしいですか？");
    if (!ok) return;
    setState((prev) => {
      const nextPrograms = prev.programs.filter((p) => p.id !== programId);
      const nextItems = prev.items.filter((it) => it.programId !== programId);
      const nextSelected =
        prev.selectedProgramId === programId ? nextPrograms[0]?.id ?? null : prev.selectedProgramId;
      return {
        ...prev,
        programs: nextPrograms.length ? nextPrograms : [makeNewProgram()],
        items: nextItems,
        selectedProgramId: nextSelected,
        lastUpdated: Date.now(),
      };
    });
  }

  function addItemFromForm() {
    if (!selectedProgram) return;
    if (!newDate) {
      alert("日付を入力してください。");
      return;
    }
    const date = parseISODateInput(newDate);
    if (!date) {
      alert("日付は YYYY-MM-DD 形式で入力してください。");
      return;
    }

    const it: ScheduleItem = {
      id: uid(),
      ...makeItemBase(selectedProgram, date),
      startTime: normalizeHHMM(newStartTime),
      endTime: normalizeHHMM(newEndTime),
      category: newCategory,
      title: newTitle,
      location: newLocation,
      roomNeeded,
      studentsCount,
      buddyCount,
      kvhRequired,
      kvhCount,
      transportMode,
      busCompany,
      busCount,
      busTripType,
      busPickup,
      busDropoff,
      arrangementsNeeded,
      notes,
      generated: false,
    };

    setState((prev) => ({
      ...prev,
      items: [...prev.items, it],
      lastUpdated: Date.now(),
    }));

    // keep values; just clear title/location/notes
    setNewTitle("");
    setNewLocation("");
    setNotes("");
  }

  function updateItem(itemId: string, partial: Partial<ScheduleItem>) {
    setState((prev) => ({
      ...prev,
      items: prev.items.map((it) => (it.id === itemId ? { ...it, ...partial } : it)),
      lastUpdated: Date.now(),
    }));
  }

  function deleteItem(itemId: string) {
    setState((prev) => ({
      ...prev,
      items: prev.items.filter((it) => it.id !== itemId),
      lastUpdated: Date.now(),
    }));
  }

  function regenAutoItems() {
    if (!selectedProgram) return;

    const now = Date.now();
    setState((prev) => {
      // replace program updated timestamp
      const nextPrograms = prev.programs.map((p) =>
        p.id === selectedProgram.id ? { ...p, lastUpdated: now } : p
      );

      const kept = removeGeneratedAuto(prev.items, selectedProgram.id);
      const auto = buildAutoItems(selectedProgram);

      return {
        ...prev,
        programs: nextPrograms,
        items: [...kept, ...auto],
        lastUpdated: now,
      };
    });

    alert("自動生成（再生成）を実行しました。");
  }

  function exportCSV() {
    if (!selectedProgram) return;
    const csv = "\uFEFF" + buildCSV(selectedProgram, state.items, exportLang);
    const safeName = selectedProgram.name.replace(/[\\/:*?"<>|]/g, "_");
    const langTag = exportLang === "en" ? "EN" : "JA";
    downloadTextFile(`${safeName}_schedule_${langTag}.csv`, csv, "text/csv;charset=utf-8");
  }

  function exportICS() {
    if (!selectedProgram) return;
    const ics = buildICSAsiaTokyo(selectedProgram, state.items, exportLang);
    const safeName = selectedProgram.name.replace(/[\\/:*?"<>|]/g, "_");
    const langTag = exportLang === "en" ? "EN" : "JA";
    downloadTextFile(`${safeName}_schedule_${langTag}.ics`, ics, "text/calendar;charset=utf-8");
  }

  function exportHTML() {
    if (!selectedProgram) return;
    const html = buildCalendarHTMLMultiMonth(selectedProgram, state.items, exportLang);
    const safeName = selectedProgram.name.replace(/[\\/:*?"<>|]/g, "_");
    const langTag = exportLang === "en" ? "EN" : "JA";
    downloadTextFile(`${safeName}_calendar_${langTag}.html`, html, "text/html;charset=utf-8");
  }

  function exportJSON() {
    const payload: AppState = {
      ...state,
      lastUpdated: Date.now(),
    };
    const safe = `ScheduleData_${new Date(payload.lastUpdated).toISOString().slice(0, 10)}.json`;
    downloadTextFile(safe, JSON.stringify(payload, null, 2), "application/json");
  }

  function importJSONFile(file: File) {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const txt = String(reader.result ?? "");
        const parsed = JSON.parse(txt) as unknown;
        if (!isValidImportedState(parsed)) {
          alert("JSONの形式または参照関係が正しくありません。Program IDの重複や孤立した予定がないか確認してください。");
          return;
        }

        const incomingUpdated = parsed.lastUpdated ?? 0;
        const currentUpdated = state.lastUpdated ?? 0;

        if (incomingUpdated && currentUpdated && incomingUpdated < currentUpdated) {
          const ok = window.confirm(
            `読み込むJSONは現在のデータより古い可能性があります。\n\n現在: ${new Date(currentUpdated).toLocaleString()}\n読み込み: ${new Date(incomingUpdated).toLocaleString()}\n\nそれでも読み込みますか？`
          );
          if (!ok) return;
        }

        const now = Date.now();
        setState({
          programs: parsed.programs,
          items: parsed.items,
          selectedProgramId: parsed.selectedProgramId ?? (parsed.programs[0]?.id ?? null),
          lastUpdated: now,
        });
        setActiveTab("program");
      } catch {
        alert("JSONの読み込みに失敗しました。");
      }
    };
    reader.readAsText(file);
  }

  function openImportDialog() {
    fileInputRef.current?.click();
  }

  function loadOverrideFromProgram(date: string) {
    if (!selectedProgram) return;
    const ov = selectedProgram.japaneseOverrides[date];
    if (!ov) {
      setOvEnabled(true);
      setOvStartTime(selectedProgram.japanese.startTime);
      setOvLesson(selectedProgram.japanese.lessonMinutes);
      setOvBreak(selectedProgram.japanese.breakMinutes);
      setOvPeriods(selectedProgram.japanese.periods);
      setOvClassCount(selectedProgram.japanese.classCount);
      setOvClassrooms("");
      setOvTeacherRooms("");
      return;
    }
    setOvEnabled(ov.enabled);
    setOvStartTime(ov.startTime);
    setOvLesson(ov.lessonMinutes);
    setOvBreak(ov.breakMinutes);
    setOvPeriods(ov.periods);
    setOvClassCount(ov.classCount);
    setOvClassrooms(ov.classrooms.join("\n"));
    setOvTeacherRooms(ov.teacherRooms.join("\n"));
  }

  function applyJapaneseOverride() {
    if (!selectedProgram) return;
    if (!overrideDate) {
      alert("上書き対象の日付を選んでください。");
      return;
    }
    const ov: JapaneseDayOverride = {
      enabled: ovEnabled,
      startTime: normalizeHHMM(ovStartTime),
      lessonMinutes: clamp(ovLesson, 30, 120),
      breakMinutes: clamp(ovBreak, 0, 60),
      periods: clamp(ovPeriods, 1, 3),
      classCount: clamp(ovClassCount, 1, 20),
      classrooms: ovClassrooms.split("\n").map((s) => s.trim()).filter(Boolean),
      teacherRooms: ovTeacherRooms.split("\n").map((s) => s.trim()).filter(Boolean),
    };

    const nextOverrides = { ...selectedProgram.japaneseOverrides, [overrideDate]: ov };
    updateProgram({ japaneseOverrides: nextOverrides });
  }

  function deleteOverride(date: string) {
    if (!selectedProgram) return;
    const next = { ...selectedProgram.japaneseOverrides };
    delete next[date];
    updateProgram({ japaneseOverrides: next });
  }

  // ★追加：初日に日本語講座を入れる（上書き作成→自動生成まで自動実行）
  function quickEnableJapaneseOnFirstDay() {
    if (!selectedProgram) return;

    const firstDate = selectedProgram.startDate;
    const defaultTime = selectedProgram.japanese.startTime || "09:00";

    const input = window.prompt(
      `初日（${firstDate}）の日本語講座 開始時刻を入力してください（例: 13:00）。\n空欄なら ${defaultTime} を使います。`,
      defaultTime
    );
    if (input === null) return; // キャンセル

    const startTime = normalizeHHMM((input || defaultTime).trim() || defaultTime);

    const ov: JapaneseDayOverride = {
      enabled: true,
      startTime,
      lessonMinutes: selectedProgram.japanese.lessonMinutes,
      breakMinutes: selectedProgram.japanese.breakMinutes,
      periods: selectedProgram.japanese.periods,
      classCount: selectedProgram.japanese.classCount,
      classrooms: [],
      teacherRooms: [],
    };

    const now = Date.now();
    const nextOverrides = { ...selectedProgram.japaneseOverrides, [firstDate]: ov };
    const updatedProgram: Program = { ...selectedProgram, japaneseOverrides: nextOverrides, lastUpdated: now };

    // 1回の setState で「上書き反映 + 自動生成」を同時にやる（手入力イベントは残す）
    setState((prev) => {
      const nextPrograms = prev.programs.map((p) => (p.id === updatedProgram.id ? updatedProgram : p));
      const kept = removeGeneratedAuto(prev.items, updatedProgram.id);
      const auto = buildAutoItems(updatedProgram);
      return {
        ...prev,
        programs: nextPrograms,
        items: [...kept, ...auto],
        lastUpdated: now,
      };
    });

    // 上書き編集UIも初日に合わせる
    setOverrideDate(firstDate);
    setOvEnabled(true);
    setOvStartTime(startTime);
    setOvLesson(ov.lessonMinutes);
    setOvBreak(ov.breakMinutes);
    setOvPeriods(ov.periods);
    setOvClassCount(ov.classCount);
    setOvClassrooms("");
    setOvTeacherRooms("");

    alert("初日の日本語講座（上書き）を作成し、自動生成（再生成）まで実行しました。");
  }

  function resetAllData() {
    const ok = window.confirm("全データを初期化します（localStorage も消えます）。よろしいですか？");
    if (!ok) return;
    localStorage.removeItem(STORAGE_KEY);
    const p = makeNewProgram();
    setState({ programs: [p], items: [], selectedProgramId: p.id, lastUpdated: Date.now() });
    setActiveTab("program");
  }

  function removeDuplicatesNow() {
    if (!selectedProgram) return;
    const next = dedupeItemsForProgram(state.items, selectedProgram.id);
    setState((prev) => ({ ...prev, items: next, lastUpdated: Date.now() }));
    alert("重複を削除しました。");
  }

  // calendar preview HTML
  const calendarHTML = useMemo(() => {
    if (!selectedProgram) return "";
    return buildCalendarHTMLMultiMonth(selectedProgram, state.items, exportLang);
  }, [selectedProgram?.id, state.items, state.lastUpdated, exportLang]);

  const showBusFields = transportMode === "Bus";

  return (
    <div className="rsjpApp" style={{
      fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif",
      padding: 12,
      maxWidth: 1300,
      margin: "0 auto",
    }}>
      <style>{`
        /* ===== RSJP Pro UI (safe) ===== */
        .rsjpApp, .rsjpApp * { box-sizing: border-box; }
        .rsjpApp { 
          color: #0f172a;
          background:
            radial-gradient(1200px 600px at 18% 12%, rgba(161,0,0,0.10), transparent 55%),
            radial-gradient(900px 500px at 82% 6%, rgba(161,0,0,0.07), transparent 60%),
            linear-gradient(180deg, #ffffff, #f8fafc);
          min-height: 100vh;
        }
        .rsjpApp input,
        .rsjpApp select,
        .rsjpApp textarea {
          max-width: 100%;
          min-width: 0;
          height: 40px;
          line-height: 40px;
          border-radius: 12px;
          border: 1px solid rgba(15,23,42,0.12);
          padding: 0 12px;
          background: rgba(255,255,255,0.92);
          outline: none;
        }
        .rsjpApp textarea { height: auto; line-height: 1.4; padding: 10px 12px; }
        .rsjpApp button {
          height: 40px;
          border-radius: 999px;
          border: 1px solid rgba(15,23,42,0.12);
          background: rgba(255,255,255,0.88);
          padding: 0 14px;
          font-weight: 800;
          cursor: pointer;
          transition: transform .06s ease, box-shadow .12s ease, background .12s ease;
        }
        .rsjpApp button:hover { box-shadow: 0 8px 20px rgba(2,6,23,0.10); transform: translateY(-1px); }
        .rsjpApp button:active { transform: translateY(0px); box-shadow: 0 4px 12px rgba(2,6,23,0.08); }
        .rsjpApp button.rsjpPrimary {
          background: #A10000;
          color: white;
          border: 1px solid rgba(0,0,0,0.05);
          box-shadow: 0 10px 22px rgba(161,0,0,0.25);
        }
        .rsjpLayout{
          display:grid;
          grid-template-columns: 320px 1fr;
          gap: 14px;
          margin-top: 12px;
          align-items: start;
          min-width: 0;
        }
        @media (max-width: 980px){
          .rsjpLayout{ grid-template-columns: 1fr; }
        }
      `}</style>

      {/* Header */}
      <div style={{
        display: "flex",
        gap: 10,
        alignItems: "center",
        justifyContent: "space-between",
        padding: "12px 12px",
        borderBottom: "1px solid rgba(15,23,42,0.10)",
        position: "sticky",
        top: 0,
        background: "rgba(255,255,255,0.85)",
        backdropFilter: "blur(10px)",
        zIndex: 5,
        borderRadius: 14,
      }}>
        <div style={{ fontWeight: 900, fontSize: 16 }}>
          RSJP / カスタム　スケジュール作成（MVP）
          <span style={{ fontWeight: 600, fontSize: 12, marginLeft: 10, opacity: 0.75 }}>
            自動保存 + JSON共有 + CSV出力 + カレンダープレビュー
          </span>
        </div>

        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <button className="rsjpPrimary" onClick={addProgram}>＋ 新規プログラム</button>
          <button onClick={exportJSON}>⬇ JSON書き出し</button>
          <button onClick={openImportDialog}>⬆ JSON読み込み</button>
          <input
            ref={fileInputRef}
            type="file"
            accept="application/json"
            style={{ display: "none" }}
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) importJSONFile(f);
              e.currentTarget.value = "";
            }}
          />
        </div>
      </div>

      <div className="rsjpLayout">
        {/* Left: Program list */}
        <div style={{
          border: "1px solid rgba(15,23,42,0.10)",
          borderRadius: 16,
          padding: 12,
          background: "rgba(255,255,255,0.78)",
          boxShadow: "0 10px 24px rgba(2,6,23,0.08)",
          backdropFilter: "blur(10px)",
          minWidth: 0,
        }}>
          <div style={{ fontWeight: 900, marginBottom: 8 }}>プログラム一覧</div>
          <input
            value={programQuery}
            onChange={(e) => setProgramQuery(e.target.value)}
            placeholder="検索（プログラム名）"
            style={{ width: "100%", marginBottom: 10 }}
          />

          <div style={{ display: "grid", gap: 10 }}>
            {state.programs
              .filter((p) => p.name.includes(programQuery))
              .map((p) => (
                <div key={p.id} style={{
                  border: p.id === selectedProgram?.id ? "2px solid rgba(161,0,0,0.55)" : "1px solid rgba(15,23,42,0.12)",
                  borderRadius: 14,
                  padding: 10,
                  background: "rgba(255,255,255,0.90)",
                  cursor: "pointer",
                }}
                  onClick={() => selectProgram(p.id)}
                  title="クリックで選択"
                >
                  <div style={{ fontWeight: 900 }}>{p.name}</div>
                  <div style={{ fontSize: 12, opacity: 0.75 }}>
                    {programTypeLabel(p.type)}
                  </div>
                  <div style={{ fontSize: 12, opacity: 0.75 }}>
                    {p.startDate} → {p.endDate}
                  </div>
                  <div style={{ fontSize: 12, opacity: 0.75 }}>
                    留学生: {p.studentsCount}
                  </div>

                  <div style={{ display: "flex", justifyContent: "space-between", gap: 8, marginTop: 8 }}>
                    <button
                      onClick={(e) => { e.stopPropagation(); deleteProgram(p.id); }}
                      style={{ height: 34 }}
                    >
                      削除
                    </button>
                  </div>
                </div>
              ))}
          </div>

          <div style={{ fontSize: 12, opacity: 0.75, marginTop: 10, lineHeight: 1.6 }}>
            ・クリックで選択。右上のJSON書き出しでバックアップできます。<br />
            ・自動生成：初日テンプレ / 平日日本語講座 / 最終日修了式<br />
            ・文化体験/企業訪問/バディランチは「日程・予定」タブで追加<br />
            ・カレンダーの見た目は「カレンダープレビュー」でスクショ同等表示
          </div>
        </div>

        {/* Right: Main */}
        <div style={{ minWidth: 0 }}>
          {!selectedProgram ? (
            <div style={{ padding: 12 }}>プログラムを選択してください。</div>
          ) : (
            <div style={{
              border: "1px solid rgba(15,23,42,0.10)",
              borderRadius: 16,
              padding: 12,
              background: "rgba(255,255,255,0.78)",
              boxShadow: "0 10px 24px rgba(2,6,23,0.08)",
              backdropFilter: "blur(10px)",
              minWidth: 0,
            }}>
              {/* Tabs */}
              <div className="rsjpTabRow" style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
                <button className={activeTab === "program" ? "rsjpActiveTab" : undefined} onClick={() => setActiveTab("program")}>プログラム設定</button>
                <button className={activeTab === "calendar" ? "rsjpActiveTab" : undefined} onClick={() => setActiveTab("calendar")}>日程・予定</button>
                <button className={activeTab === "share" ? "rsjpActiveTab" : undefined} onClick={() => setActiveTab("share")}>共有・出力</button>

                <div style={{ flex: 1 }} />

                {activeTab !== "share" && (
                  <button className="rsjpPrimary" onClick={regenAutoItems}>
                    ⚡ 自動生成（再生成）
                  </button>
                )}
                {activeTab === "share" && (
                  <>
                    <button onClick={exportCSV}>CSV出力</button>
                    <button onClick={exportICS}>ICS出力</button>
                    <button onClick={exportHTML}>HTML出力</button>
                  </>
                )}
              </div>

              {activeTab === "program" && (
                <div style={{ display: "grid", gap: 14 }}>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 220px 220px", gap: 10 }}>
                    <input
                      value={selectedProgram.name}
                      onChange={(e) => updateProgram({ name: e.target.value })}
                      placeholder="プログラム名"
                      style={{ width: "100%" }}
                    />
                    <select
                      value={selectedProgram.type}
                      onChange={(e) => updateProgram({ type: e.target.value as ProgramType })}
                    >
                      <option value="RSJP">RSJP（レギュラー）</option>
                      <option value="Custom">カスタム</option>
                    </select>
                    <input
                      type="number"
                      value={selectedProgram.studentsCount}
                      onChange={(e) => updateProgram({ studentsCount: clamp(Number(e.target.value), 1, 999) })}
                      placeholder="留学生数"
                    />
                  </div>

                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
                    <div>
                      <div style={{ fontSize: 12, opacity: 0.75, marginBottom: 4 }}>開始日</div>
                      <input
                        value={selectedProgram.startDate}
                        onChange={(e) => {
                          const v = parseISODateInput(e.target.value);
                          if (!v) return;
                          updateProgram({ startDate: v });
                        }}
                        placeholder="YYYY-MM-DD"
                      />
                    </div>
                    <div>
                      <div style={{ fontSize: 12, opacity: 0.75, marginBottom: 4 }}>終了日</div>
                      <input
                        value={selectedProgram.endDate}
                        onChange={(e) => {
                          const v = parseISODateInput(e.target.value);
                          if (!v) return;
                          updateProgram({ endDate: v });
                        }}
                        placeholder="YYYY-MM-DD"
                      />
                    </div>
                    <div style={{ display: "flex", alignItems: "end", gap: 8 }}>
                      <button onClick={quickEnableJapaneseOnFirstDay} title="初日に日本語講座（上書き）を作り、自動生成まで実行します">
                        初日に日本語講座を入れる
                      </button>
                    </div>
                  </div>

                  <div style={{ borderTop: "1px solid rgba(15,23,42,0.10)", paddingTop: 10 }}>
                    <div style={{ fontWeight: 900, marginBottom: 8 }}>日本語講座（自動生成）</div>

                    <div style={{ display: "grid", gridTemplateColumns: "140px 1fr 1fr 1fr 1fr", gap: 10, alignItems: "center" }}>
                      <div style={{ fontSize: 12, opacity: 0.75 }}>有効</div>
                      <select value={selectedProgram.japanese.enabled} onChange={(e) => updateProgram({ japanese: { ...selectedProgram.japanese, enabled: e.target.value as YesNo } })}>
                        <option value="Yes">あり</option>
                        <option value="No">なし</option>
                      </select>

                      <div style={{ fontSize: 12, opacity: 0.75 }}>開始時刻</div>
                      <input value={selectedProgram.japanese.startTime} onChange={(e) => updateProgram({ japanese: { ...selectedProgram.japanese, startTime: normalizeHHMM(e.target.value) } })} />
                      <div />
                    </div>

                    <div style={{ display: "grid", gridTemplateColumns: "140px 1fr 140px 1fr 140px 1fr", gap: 10, alignItems: "center", marginTop: 8 }}>
                      <div style={{ fontSize: 12, opacity: 0.75 }}>授業(分)</div>
                      <input type="number" value={selectedProgram.japanese.lessonMinutes} onChange={(e) => updateProgram({ japanese: { ...selectedProgram.japanese, lessonMinutes: clamp(Number(e.target.value), 30, 120) } })} />
                      <div style={{ fontSize: 12, opacity: 0.75 }}>休憩(分)</div>
                      <input type="number" value={selectedProgram.japanese.breakMinutes} onChange={(e) => updateProgram({ japanese: { ...selectedProgram.japanese, breakMinutes: clamp(Number(e.target.value), 0, 60) } })} />
                      <div style={{ fontSize: 12, opacity: 0.75 }}>コマ数</div>
                      <input type="number" value={selectedProgram.japanese.periods} onChange={(e) => updateProgram({ japanese: { ...selectedProgram.japanese, periods: clamp(Number(e.target.value), 1, 3) } })} />
                    </div>

                    <div style={{ display: "grid", gridTemplateColumns: "140px 1fr 1fr", gap: 10, alignItems: "center", marginTop: 8 }}>
                      <div style={{ fontSize: 12, opacity: 0.75 }}>クラス数</div>
                      <input type="number" value={selectedProgram.japanese.classCount} onChange={(e) => updateProgram({ japanese: { ...selectedProgram.japanese, classCount: clamp(Number(e.target.value), 1, 20) } })} />
                      <input
                        value={selectedProgram.japanese.defaultTeacherRoom ?? ""}
                        onChange={(e) => updateProgram({ japanese: { ...selectedProgram.japanese, defaultTeacherRoom: e.target.value } })}
                        placeholder="講師控室（デフォルト）"
                      />
                    </div>

                    <div style={{ fontSize: 12, opacity: 0.7, marginTop: 10, lineHeight: 1.6 }}>
                      ・日本語講座は「平日」に自動生成されます（週末は除外）。<br />
                      ・日別の時刻やクラス数などを変えたい場合は、下の「日別上書き」を使ってください。<br />
                      ・上書きを作った後は「自動生成（再生成）」で反映されます（初日ボタンは自動で再生成まで実行します）。
                    </div>

                    {/* Overrides */}
                    <div style={{ border: "1px solid rgba(15,23,42,0.10)", borderRadius: 14, padding: 10, marginTop: 12, background: "rgba(255,255,255,0.9)" }}>
                      <div style={{ fontWeight: 900, marginBottom: 8 }}>日別上書き（日本語講座）</div>

                      <div style={{ display: "grid", gridTemplateColumns: "220px 1fr 1fr", gap: 10, alignItems: "center" }}>
                        <div>
                          <div style={{ fontSize: 12, opacity: 0.75, marginBottom: 4 }}>対象日</div>
                          <input
                            value={overrideDate}
                            onChange={(e) => {
                              const v = parseISODateInput(e.target.value);
                              if (!v) return;
                              setOverrideDate(v);
                              loadOverrideFromProgram(v);
                            }}
                            placeholder="YYYY-MM-DD"
                          />
                        </div>

                        <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, opacity: 0.85 }}>
                          <input type="checkbox" checked={ovEnabled} onChange={(e) => setOvEnabled(e.target.checked)} />
                          この日に日本語講座を「実施する」
                        </label>

                        <button onClick={() => applyJapaneseOverride()} className="rsjpPrimary">
                          この内容で上書き保存
                        </button>
                      </div>

                      <div style={{ display: "grid", gridTemplateColumns: "140px 1fr 140px 1fr 140px 1fr", gap: 10, alignItems: "center", marginTop: 10 }}>
                        <div style={{ fontSize: 12, opacity: 0.75 }}>開始時刻</div>
                        <input value={ovStartTime} onChange={(e) => setOvStartTime(normalizeHHMM(e.target.value))} />
                        <div style={{ fontSize: 12, opacity: 0.75 }}>授業(分)</div>
                        <input type="number" value={ovLesson} onChange={(e) => setOvLesson(clamp(Number(e.target.value), 30, 120))} />
                        <div style={{ fontSize: 12, opacity: 0.75 }}>休憩(分)</div>
                        <input type="number" value={ovBreak} onChange={(e) => setOvBreak(clamp(Number(e.target.value), 0, 60))} />
                      </div>

                      <div style={{ display: "grid", gridTemplateColumns: "140px 1fr 140px 1fr", gap: 10, alignItems: "center", marginTop: 8 }}>
                        <div style={{ fontSize: 12, opacity: 0.75 }}>コマ数</div>
                        <input type="number" value={ovPeriods} onChange={(e) => setOvPeriods(clamp(Number(e.target.value), 1, 3))} />
                        <div style={{ fontSize: 12, opacity: 0.75 }}>クラス数</div>
                        <input type="number" value={ovClassCount} onChange={(e) => setOvClassCount(clamp(Number(e.target.value), 1, 20))} />
                      </div>

                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 10 }}>
                        <div>
                          <div style={{ fontSize: 12, opacity: 0.75, marginBottom: 4 }}>教室（1行=1クラス）</div>
                          <textarea value={ovClassrooms} onChange={(e) => setOvClassrooms(e.target.value)} rows={5} placeholder="例）YY310\nYY311\nYY312" />
                        </div>
                        <div>
                          <div style={{ fontSize: 12, opacity: 0.75, marginBottom: 4 }}>講師控室（1行=1名）</div>
                          <textarea value={ovTeacherRooms} onChange={(e) => setOvTeacherRooms(e.target.value)} rows={5} placeholder="例）YY210\nYY211" />
                        </div>
                      </div>

                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, marginTop: 10 }}>
                        <div style={{ fontSize: 12, opacity: 0.75 }}>
                          ※上書きを削除したい場合は「削除」を押してください。
                        </div>
                        <button onClick={() => deleteOverride(overrideDate)}>この日の上書きを削除</button>
                      </div>
                    </div>
                  </div>

                  <div style={{ borderTop: "1px solid rgba(15,23,42,0.10)", paddingTop: 12 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
                      <div style={{ fontWeight: 900 }}>トラブル時</div>
                      <button onClick={resetAllData}>全データ初期化（localStorage削除）</button>
                    </div>
                  </div>
                </div>
              )}

              {activeTab === "calendar" && (
                <div>
                  {/* Add form */}
                  <div style={{ border: "1px solid rgba(15,23,42,0.10)", borderRadius: 14, padding: 10, background: "rgba(255,255,255,0.9)" }}>
                    <div style={{ fontWeight: 900, marginBottom: 8 }}>予定を追加</div>

                    <div style={{ display: "grid", gridTemplateColumns: "160px 140px 140px 1fr 1fr", gap: 8, alignItems: "center" }}>
                      <input value={newDate} onChange={(e) => setNewDate(e.target.value)} placeholder="YYYY-MM-DD" />
                      <input value={newStartTime} onChange={(e) => setNewStartTime(e.target.value)} placeholder="開始 HH:MM" />
                      <input value={newEndTime} onChange={(e) => setNewEndTime(e.target.value)} placeholder="終了 HH:MM" />

                      <select value={newCategory} onChange={(e) => setNewCategory(e.target.value as Category)}>
                        <option value="JapaneseClass">日本語講座</option>
                        <option value="Orientation">オリエン</option>
                        <option value="Escort">引率</option>
                        <option value="CampusTour">キャンパスツアー</option>
                        <option value="Cultural">文化体験</option>
                        <option value="CompanyVisit">企業訪問</option>
                        <option value="BuddyLunch">バディランチ</option>
                        <option value="Ceremony">修了式</option>
                        <option value="Other">その他</option>
                      </select>

                      <input value={newTitle} onChange={(e) => setNewTitle(e.target.value)} placeholder="タイトル" />
                    </div>

                    <div style={{ display: "grid", gridTemplateColumns: "1fr 220px 220px 1fr", gap: 8, marginTop: 8, alignItems: "center" }}>
                      <input value={newLocation} onChange={(e) => setNewLocation(e.target.value)} placeholder="場所/集合" />
                      <select value={roomNeeded} onChange={(e) => setRoomNeeded(e.target.value as YesNo)}>
                        <option value="No">教室: なし</option>
                        <option value="Yes">教室: あり</option>
                      </select>
                      <input type="number" value={studentsCount} onChange={(e) => setStudentsCount(clamp(Number(e.target.value), 0, 999))} placeholder="留学生数" />
                      <input value={notes} onChange={(e) => setNotes(e.target.value)} placeholder="備考" />
                    </div>

                    <div style={{ display: "grid", gridTemplateColumns: "220px 220px 220px 1fr", gap: 8, marginTop: 8, alignItems: "center" }}>
                      <input type="number" value={buddyCount} onChange={(e) => setBuddyCount(clamp(Number(e.target.value), 0, 999))} placeholder="バディ数" />
                      <select value={kvhRequired} onChange={(e) => setKvhRequired(e.target.value as YesNo)}>
                        <option value="No">KVH: なし</option>
                        <option value="Yes">KVH: あり</option>
                      </select>
                      <input type="number" value={kvhCount} onChange={(e) => setKvhCount(clamp(Number(e.target.value), 0, 999))} placeholder="KVH人数" />
                      <select value={transportMode} onChange={(e) => setTransportMode(e.target.value as TransportMode)}>
                        <option value="None">移動: なし</option>
                        <option value="Bus">移動: バス</option>
                        <option value="Walk">移動: 徒歩</option>
                        <option value="OnCampus">移動: 学内</option>
                      </select>
                    </div>

                    {showBusFields && (
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 160px 160px 1fr 1fr", gap: 8, marginTop: 8 }}>
                        <input value={busCompany} onChange={(e) => setBusCompany(e.target.value)} placeholder="バス会社" />
                        <input type="number" value={busCount} onChange={(e) => setBusCount(clamp(Number(e.target.value), 1, 99))} placeholder="台数" />
                        <select value={busTripType} onChange={(e) => setBusTripType(e.target.value as BusTripType)}>
                          <option value="OneWay">片道</option>
                          <option value="RoundTrip">往復</option>
                        </select>
                        <input value={busPickup} onChange={(e) => setBusPickup(e.target.value)} placeholder="乗車場所" />
                        <input value={busDropoff} onChange={(e) => setBusDropoff(e.target.value)} placeholder="降車場所" />
                      </div>
                    )}

                    <div style={{ display: "grid", gridTemplateColumns: "220px 1fr 1fr", gap: 8, marginTop: 8, alignItems: "center" }}>
                      <select value={arrangementsNeeded} onChange={(e) => setArrangementsNeeded(e.target.value as YesNo)}>
                        <option value="No">手配: なし</option>
                        <option value="Yes">手配: あり</option>
                      </select>

                      <button onClick={addItemFromForm} className="rsjpPrimary" style={{ fontWeight: 700 }}>
                        ＋ この内容で追加
                      </button>
                      <div style={{ fontSize: 12, opacity: 0.75 }}>
                        ※追加後、下の「予定一覧」や「カレンダープレビュー」に反映されます
                      </div>
                    </div>
                  </div>

                  {/* カレンダープレビュー（スクショ同等表示） */}
                  <div style={{ display: "flex", gap: 10, marginTop: 14, alignItems: "center", flexWrap: "wrap" }}>
                    <div style={{ fontWeight: 900 }}>カレンダープレビュー</div>
                    <div style={{ fontSize: 12, opacity: 0.75 }}>※印刷や配布用は「共有・出力」タブの HTML出力 を推奨</div>
                  </div>

                  {calendarHTML && (
                    <div style={{ border: "1px solid rgba(15,23,42,0.10)", borderRadius: 14, overflow: "hidden", marginTop: 8 }}>
                      <iframe
                        srcDoc={calendarHTML}
                        style={{ width: "100%", height: 520, border: "0" }}
                        sandbox="allow-same-origin"
                      />
                    </div>
                  )}

                  {/* 予定一覧（編集） */}
                  <div style={{ marginTop: 14 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, marginBottom: 8, flexWrap: "wrap" }}>
  <div style={{ fontWeight: 800 }}>予定一覧（編集）</div>
  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
    <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, opacity: 0.85 }}>
      <input type="checkbox" checked={showEnglishFields} onChange={(e) => setShowEnglishFields(e.target.checked)} />
      英語欄を表示（EN手入力）
    </label>
    <button onClick={removeDuplicatesNow}>重複を削除（今のプログラム）</button>
  </div>
</div>

                    <div style={{ fontSize: 12, opacity: 0.7, marginBottom: 10, lineHeight: 1.5 }}>
                      ・英語出力にしたい予定は、EN欄（Title/Location/Notes）を入れると確実です（辞書は補助）。<br />
                      ・日本語講座（自動生成）は「備考」に講師控室などが入り、カレンダー側でまとめ表示されます。
                    </div>

                    <div style={{ display: "grid", gap: 10 }}>
                      {selectedProgramItems.map((it) => (
                        <div key={it.id} style={{ border: "1px solid rgba(15,23,42,0.10)", borderRadius: 14, padding: 10, background: "rgba(255,255,255,0.90)" }}>
                          <div style={{ display: "flex", gap: 8, alignItems: "center", justifyContent: "space-between", flexWrap: "wrap" }}>
                            <div style={{ fontWeight: 900 }}>
                              {it.date}（{dayOfWeekJP(it.date)}） {it.startTime}-{it.endTime}
                              {it.generated ? <span style={{ fontSize: 12, marginLeft: 8, opacity: 0.7 }}>Auto</span> : null}
                            </div>

                            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                              <button
                                onClick={() => {
                                  const ok = window.confirm("この予定を削除します。よろしいですか？");
                                  if (ok) deleteItem(it.id);
                                }}
                              >
                                削除
                              </button>
                            </div>
                          </div>

                          <div style={{ display: "grid", gridTemplateColumns: "160px 140px 140px 220px 1fr", gap: 8, marginTop: 8, alignItems: "center" }}>
                            <input
                              value={it.date}
                              onChange={(e) => updateItem(it.id, { date: e.target.value })}
                              style={{ width: "100%" }}
                              title="日付"
                            />
                            <input
                              value={it.startTime}
                              onChange={(e) => updateItem(it.id, { startTime: normalizeHHMM(e.target.value) })}
                              style={{ width: "100%" }}
                              title="開始"
                            />
                            <input
                              value={it.endTime}
                              onChange={(e) => updateItem(it.id, { endTime: normalizeHHMM(e.target.value) })}
                              style={{ width: "100%" }}
                              title="終了"
                            />

                            <select
                              value={it.category}
                              onChange={(e) => updateItem(it.id, { category: e.target.value as Category })}
                              style={{ width: "100%" }}
                              title="カテゴリ"
                            >
                              <option value="JapaneseClass">日本語講座</option>
                              <option value="Orientation">オリエン</option>
                              <option value="Escort">引率</option>
                              <option value="CampusTour">キャンパスツアー</option>
                              <option value="Cultural">文化体験</option>
                              <option value="CompanyVisit">企業訪問</option>
                              <option value="BuddyLunch">バディランチ</option>
                              <option value="Ceremony">修了式</option>
                              <option value="Other">その他</option>
                            </select>

                            <input
                                        value={it.title}
                                        onChange={(e) => updateItem(it.id, { title: e.target.value })}
                                        style={{ width: "100%" }}
                                        title="タイトル"
                                      />
                                      {showEnglishFields && (
                                        <input
                                          value={it.titleEn ?? ""}
                                          onChange={(e) => updateItem(it.id, { titleEn: e.target.value })}
                                          style={{ width: "100%" }}
                                          placeholder="Title (EN) 例）Tea Ceremony"
                                          title="Title (EN)"
                                        />
                                      )}
                          </div>

                          <div style={{ display: "grid", gridTemplateColumns: "1fr 220px 220px 1fr", gap: 8, marginTop: 8, alignItems: "center" }}>
                            <input
                                        value={it.location}
                                        onChange={(e) => updateItem(it.id, { location: e.target.value })}
                                        style={{ width: "100%" }}
                                        placeholder="場所/集合"
                                      />
                                      {showEnglishFields && (
                                        <input
                                          value={it.locationEn ?? ""}
                                          onChange={(e) => updateItem(it.id, { locationEn: e.target.value })}
                                          style={{ width: "100%" }}
                                          placeholder="Location (EN) 例）Nijo Castle / Meeting point"
                                        />
                                      )}
                            <select value={it.roomNeeded} onChange={(e) => updateItem(it.id, { roomNeeded: e.target.value as YesNo })}>
                              <option value="No">教室: なし</option>
                              <option value="Yes">教室: あり</option>
                            </select>
                            <input
                              type="number"
                              value={it.studentsCount}
                              onChange={(e) => updateItem(it.id, { studentsCount: clamp(Number(e.target.value), 0, 999) })}
                              style={{ width: "100%" }}
                              placeholder="留学生数"
                            />
                            <select value={it.transportMode} onChange={(e) => updateItem(it.id, { transportMode: e.target.value as TransportMode })}>
                              <option value="Bus">移動: バス</option>
                              <option value="Walk">移動: 徒歩</option>
                              <option value="OnCampus">移動: 学内</option>
                              <option value="None">移動: なし</option>
                            </select>
                            <input
                                        value={it.notes}
                                        onChange={(e) => updateItem(it.id, { notes: e.target.value })}
                                        style={{ width: "100%" }}
                                        placeholder="備考"
                                      />
                                      {showEnglishFields && (
                                        <input
                                          value={it.notesEn ?? ""}
                                          onChange={(e) => updateItem(it.id, { notesEn: e.target.value })}
                                          style={{ width: "100%" }}
                                          placeholder="Notes (EN) optional"
                                        />
                                      )}
                          </div>

                          <div style={{ fontSize: 12, opacity: 0.75, marginTop: 8, lineHeight: 1.6 }}>
                            ・CSV/ICS/HTMLの英語出力は「共有・出力」タブで切り替えます。<br />
                            ・辞書にない固有名詞（新しい文化体験名・企業名など）は EN欄へ入力してください。
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              )}

              {activeTab === "share" && (
                <div style={{ display: "grid", gap: 12 }}>
                  <div style={{ border: "1px solid rgba(15,23,42,0.10)", borderRadius: 14, padding: 12, background: "rgba(255,255,255,0.90)" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                      <div style={{ fontWeight: 900 }}>共有・出力</div>

                      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                        <div style={{ fontSize: 12, opacity: 0.75 }}>出力言語</div>
                        <select value={exportLang} onChange={(e) => setExportLang(e.target.value as ExportLang)}>
                          <option value="ja">日本語</option>
                          <option value="en">英語</option>
                        </select>
                      </div>
                    </div>

                    <div style={{ marginTop: 10, display: "grid", gap: 8 }}>
                      <button onClick={exportCSV} className="rsjpPrimary">CSV出力（UTF-8 / Excel向け）</button>
                      <button onClick={exportICS}>ICS出力（日本時間 / カレンダー取込）</button>
                      <button onClick={exportHTML}>カレンダーHTML出力（印刷・配布用）</button>
                      <button onClick={exportJSON}>⬇ JSON書き出し（バックアップ・共有）</button>
                    </div>

                    <div style={{ fontSize: 12, opacity: 0.7, marginTop: 10, lineHeight: 1.6 }}>
                      ・画面表示は日本語固定。CSV/ICS/HTMLのみ英語化できます。<br />
                      ・英語出力で日本語が残る場合：その予定の EN欄（Title/Location/Notes）を埋めると確実です。<br />
                      ・相手がこのアプリに取り込むなら JSON を渡してください（読み込み後に「自動生成（再生成）」推奨）。
                    </div>
                  </div>

                  <div style={{ border: "1px solid rgba(15,23,42,0.10)", borderRadius: 14, padding: 12, background: "rgba(255,255,255,0.90)" }}>
                    <div style={{ fontWeight: 900, marginBottom: 8 }}>読み込み</div>
                    <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                      <button onClick={openImportDialog}>⬆ JSON読み込み</button>
                      <button onClick={regenAutoItems} className="rsjpPrimary">自動生成（再生成）</button>
                    </div>
                    <div style={{ fontSize: 12, opacity: 0.7, marginTop: 10, lineHeight: 1.6 }}>
                      ・JSON読み込み後に、必要なら「自動生成（再生成）」で Auto 部分を作り直してください。
                    </div>
                  </div>

                  <div style={{ border: "1px solid rgba(15,23,42,0.10)", borderRadius: 14, padding: 12, background: "rgba(255,255,255,0.90)" }}>
                    <div style={{ fontWeight: 900, marginBottom: 8 }}>トラブル時</div>
                    <button onClick={resetAllData}>全データ初期化（localStorage削除）</button>
                    <div style={{ fontSize: 12, opacity: 0.7, marginTop: 10, lineHeight: 1.6 }}>
                      ・データが壊れた場合は初期化 → JSON読み込みで復旧できます。
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
