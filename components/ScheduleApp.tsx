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

type JapaneseDayOverride = {
  enabled: boolean; // その日に日本語講座を入れる/入れない
  startTime: string; // HH:MM
  lessonMinutes: number; // 50 or 60
  breakMinutes: number; // 10
  periods: number; // 1-3
  classCount: number; // 1+
  // 教室は「クラス数ぶん」まで入力可。足りない分は default を使う。
  classrooms: string[];
  teacherRooms: string[];
};

type JapaneseDefaults = {
  enabled: boolean;
  startTime: string; // HH:MM
  lessonMinutes: number; // 50/60
  breakMinutes: number; // 10
  periods: number; // 1-3
  classCount: number; // 1+
  defaultClassroom: string;
  defaultTeacherRoom: string;
  classNames: string[]; // クラス名（例：嵐山、宇治…）
  classRooms: string[]; // クラス別 教室（例：YY301...）
};

type CeremonyDefaults = {
  startTime: string; // HH:MM
  durationMinutes: number; // e.g., 60/90
  location: string;
};

type Program = {
  id: string;
  name: string;
  type: ProgramType;
  startDate: string; // YYYY-MM-DD
  endDate: string; // YYYY-MM-DD
  studentsCount: number; // 基本人数
  defaultBuddyCount: number; // 初期値=5
  japanese: JapaneseDefaults;
  japaneseOverrides: Record<string, JapaneseDayOverride>; // date -> override
  ceremony: CeremonyDefaults;
  lastUpdated: number;
};

type EventMaster = {
  id: string;
  title: string;
  category: "Cultural" | "CompanyVisit" | "BuddyLunch" | "Other";
  defaultStartTime?: string; // HH:MM
  defaultDurationMinutes?: number; // e.g. 95
  defaultKVHRequired?: boolean;
  defaultKVHCount?: number;
  defaultTransportMode?: TransportMode;
  defaultArrangementsNeeded?: boolean;
  defaultNotes?: string;
};

type ScheduleItem = {
  id: string;
  programId: string;
  date: string; // YYYY-MM-DD
  startTime: string; // HH:MM
  endTime: string; // HH:MM
  category: Category;
  title: string;
  location: string;
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

  // 日本語講座用（CSVで見分けやすくする）
  classIndex?: number; // 1..classCount
  periodIndex?: number; // 1..periods
  generated: boolean;
  generatedKind?: "Auto";
};

type AppState = {
  programs: Program[];
  items: ScheduleItem[];
  selectedProgramId: string | null;
  lastUpdated: number;
};

const STORAGE_KEY = "rsjp_schedule_builder_v1";

const DEFAULT_JP_CLASS_NAMES = ["嵐山", "宇治", "祇園", "西陣", "東山"];

function defaultClassNameFor(index1: number) {
  if (index1 >= 1 && index1 <= DEFAULT_JP_CLASS_NAMES.length) return DEFAULT_JP_CLASS_NAMES[index1 - 1];
  return `クラス${index1}`;
}

function normalizeClassNames(input: any, classCount: number): string[] {
  const arr = Array.isArray(input) ? input : [];
  const out: string[] = [];
  for (let i = 1; i <= classCount; i++) {
    const v = String(arr[i - 1] ?? "").trim();
    out.push(v || defaultClassNameFor(i));
  }
  return out;
}

const DEFAULT_JP_CLASS_ROOMS = ["YY301", "YY302", "YY303", "YY304", "YY305"];

function defaultClassRoomFor(index1: number) {
  // 1 -> YY301, 2 -> YY302 ...
  if (index1 >= 1 && index1 <= DEFAULT_JP_CLASS_ROOMS.length) return DEFAULT_JP_CLASS_ROOMS[index1 - 1];
  return `YY${300 + index1}`;
}

function normalizeClassRooms(input: any, classCount: number): string[] {
  const arr = Array.isArray(input) ? input : [];
  const out: string[] = [];
  for (let i = 1; i <= classCount; i++) {
    const v = String(arr[i - 1] ?? "").trim();
    out.push(v || defaultClassRoomFor(i));
  }
  return out;
}


function uid(prefix = "id") {
  return `${prefix}_${Math.random().toString(16).slice(2)}_${Date.now().toString(16)}`;
}

function clamp(n: number, min: number, max: number) {
  return Math.max(min, Math.min(max, n));
}

function pad2(n: number) {
  return n < 10 ? `0${n}` : `${n}`;
}

function parseTimeToMinutes(hhmm: string): number | null {
  const m = /^(\d{1,2}):(\d{2})$/.exec(hhmm.trim());
  if (!m) return null;
  const hh = Number(m[1]);
  const mm = Number(m[2]);
  if (!Number.isFinite(hh) || !Number.isFinite(mm)) return null;
  if (hh < 0 || hh > 23 || mm < 0 || mm > 59) return null;
  return hh * 60 + mm;
}

function minutesToTime(mins: number) {
  const m = ((mins % (24 * 60)) + 24 * 60) % (24 * 60);
  const hh = Math.floor(m / 60);
  const mm = m % 60;
  return `${pad2(hh)}:${pad2(mm)}`;
}

function addMinutes(hhmm: string, add: number) {
  const base = parseTimeToMinutes(hhmm);
  if (base === null) return hhmm;
  return minutesToTime(base + add);
}

function isISODate(s: string) {
  return /^\d{4}-\d{2}-\d{2}$/.test(s);
}

function toDateObj(iso: string): Date {
  // 日付だけ扱うので UTC で統一（ずれにくい）
  const [y, m, d] = iso.split("-").map(Number);
  return new Date(Date.UTC(y, m - 1, d));
}

function fromDateObjUTC(d: Date): string {
  const y = d.getUTCFullYear();
  const m = d.getUTCMonth() + 1;
  const day = d.getUTCDate();
  return `${y}-${pad2(m)}-${pad2(day)}`;
}

function eachDateInclusive(startISO: string, endISO: string): string[] {
  if (!isISODate(startISO) || !isISODate(endISO)) return [];
  const start = toDateObj(startISO);
  const end = toDateObj(endISO);
  if (start.getTime() > end.getTime()) return [];
  const out: string[] = [];
  let cur = start;
  while (cur.getTime() <= end.getTime()) {
    out.push(fromDateObjUTC(cur));
    cur = new Date(cur.getTime());
    cur.setUTCDate(cur.getUTCDate() + 1);
  }
  return out;
}

function dayOfWeekJP(iso: string): string {
  const d = toDateObj(iso);
  const fmt = new Intl.DateTimeFormat("ja-JP", { weekday: "short", timeZone: "UTC" });
  return fmt.format(d);
}

function isWeekday(iso: string): boolean {
  const d = toDateObj(iso);
  const dow = d.getUTCDay(); // 0 Sun .. 6 Sat
  return dow >= 1 && dow <= 5;
}

function csvEscape(v: string) {
  const s = (v ?? "").toString();
  if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function downloadTextFile(filename: string, text: string, mime = "text/plain") {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function makeDefaultEventMasters(): EventMaster[] {
  const cultural: EventMaster[] = [
    { id: "tea", title: "茶道体験", category: "Cultural", defaultStartTime: "13:10", defaultDurationMinutes: 170, defaultKVHRequired: true, defaultKVHCount: 1, defaultTransportMode: "Bus", defaultArrangementsNeeded: true },
    { id: "calligraphy", title: "書道体験", category: "Cultural", defaultStartTime: "13:10", defaultDurationMinutes: 95, defaultKVHRequired: true, defaultKVHCount: 1, defaultTransportMode: "Bus", defaultArrangementsNeeded: true },
    { id: "maiko", title: "舞妓体験（鑑賞）", category: "Cultural", defaultStartTime: "13:10", defaultDurationMinutes: 95, defaultKVHRequired: true, defaultKVHCount: 1, defaultTransportMode: "Bus", defaultArrangementsNeeded: true },
    { id: "yuzen", title: "友禅染体験", category: "Cultural", defaultStartTime: "13:10", defaultDurationMinutes: 95, defaultKVHRequired: true, defaultKVHCount: 1, defaultTransportMode: "Bus", defaultArrangementsNeeded: true },
    { id: "kitanotenmangu", title: "寺社参拝（例：北野天満宮）", category: "Cultural", defaultStartTime: "13:10", defaultDurationMinutes: 95, defaultKVHRequired: true, defaultKVHCount: 1, defaultTransportMode: "Bus", defaultArrangementsNeeded: false },
    { id: "wagashi", title: "和菓子作り体験", category: "Cultural", defaultStartTime: "13:10", defaultDurationMinutes: 95, defaultKVHRequired: true, defaultKVHCount: 1, defaultTransportMode: "Bus", defaultArrangementsNeeded: true },
    { id: "ikebana", title: "華道体験", category: "Cultural", defaultStartTime: "13:10", defaultDurationMinutes: 95, defaultKVHRequired: true, defaultKVHCount: 1, defaultTransportMode: "Bus", defaultArrangementsNeeded: true },
    { id: "zen", title: "禅体験（坐禅）", category: "Cultural", defaultStartTime: "13:10", defaultDurationMinutes: 95, defaultKVHRequired: true, defaultKVHCount: 1, defaultTransportMode: "Bus", defaultArrangementsNeeded: false },
    { id: "taiko", title: "和太鼓体験", category: "Cultural", defaultStartTime: "13:10", defaultDurationMinutes: 95, defaultKVHRequired: true, defaultKVHCount: 1, defaultTransportMode: "Bus", defaultArrangementsNeeded: true },
    { id: "cook", title: "日本料理体験", category: "Cultural", defaultStartTime: "13:10", defaultDurationMinutes: 95, defaultKVHRequired: true, defaultKVHCount: 1, defaultTransportMode: "Bus", defaultArrangementsNeeded: true },
    { id: "goldleaf", title: "金箔体験", category: "Cultural", defaultStartTime: "13:10", defaultDurationMinutes: 95, defaultKVHRequired: true, defaultKVHCount: 1, defaultTransportMode: "Bus", defaultArrangementsNeeded: true },
    { id: "craft", title: "伝統工芸見学", category: "Cultural", defaultStartTime: "13:10", defaultDurationMinutes: 95, defaultKVHRequired: true, defaultKVHCount: 1, defaultTransportMode: "Bus", defaultArrangementsNeeded: false },
    { id: "lecture", title: "京都文化講義（学内）", category: "Cultural", defaultStartTime: "13:10", defaultDurationMinutes: 95, defaultKVHRequired: true, defaultKVHCount: 1, defaultTransportMode: "OnCampus", defaultArrangementsNeeded: true },
    { id: "walk", title: "市内散策（ガイド付き）", category: "Cultural", defaultStartTime: "13:10", defaultDurationMinutes: 95, defaultKVHRequired: true, defaultKVHCount: 1, defaultTransportMode: "Walk", defaultArrangementsNeeded: false },
    { id: "museum", title: "博物館・資料館見学", category: "Cultural", defaultStartTime: "13:10", defaultDurationMinutes: 95, defaultKVHRequired: true, defaultKVHCount: 1, defaultTransportMode: "Bus", defaultArrangementsNeeded: false },
  ];

  const company: EventMaster[] = [
    { id: "company_visit", title: "企業訪問", category: "CompanyVisit", defaultStartTime: "13:10", defaultDurationMinutes: 95, defaultKVHRequired: true, defaultKVHCount: 1, defaultTransportMode: "Bus", defaultArrangementsNeeded: true },
  ];

  const buddy: EventMaster[] = [
    { id: "buddy_lunch", title: "バディランチ", category: "BuddyLunch", defaultStartTime: "12:20", defaultDurationMinutes: 50, defaultKVHRequired: false, defaultKVHCount: 0, defaultTransportMode: "OnCampus", defaultArrangementsNeeded: false },
  ];

  return [...cultural, ...company, ...buddy];
}

function makeNewProgram(): Program {
  const todayUTC = new Date();
  const todayISO = fromDateObjUTC(new Date(Date.UTC(todayUTC.getUTCFullYear(), todayUTC.getUTCMonth(), todayUTC.getUTCDate())));
  const endISO = todayISO;
  const now = Date.now();

  return {
    id: uid("prog"),
    name: "新規プログラム",
    type: "RSJP",
    startDate: todayISO,
    endDate: endISO,
    studentsCount: 20,
    defaultBuddyCount: 5,
    japanese: {
      enabled: true,
      startTime: "09:00",
      lessonMinutes: 50,
      breakMinutes: 10,
      periods: 3,
      classCount: 1,
      defaultClassroom: "",
      defaultTeacherRoom: "YY305",
      classNames: normalizeClassNames(undefined, 1),
      classRooms: normalizeClassRooms(undefined, 1),
    },
    japaneseOverrides: {},
    ceremony: {
      startTime: "13:10",
      durationMinutes: 60,
      location: "",
    },
    lastUpdated: now,
  };
}

function loadState(): AppState {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      const p = makeNewProgram();
      return { programs: [p], items: [], selectedProgramId: p.id, lastUpdated: Date.now() };
    }
    const parsed = JSON.parse(raw) as AppState;
    if (!parsed || !Array.isArray(parsed.programs) || !Array.isArray(parsed.items)) throw new Error("bad state");
    const migratedPrograms = parsed.programs.map((p: any) => {
      const cc = clamp(Number(p?.japanese?.classCount ?? 1), 1, 50);
      const existing = p?.japanese?.classNames;
      const fixed = {
        ...p,
        japanese: {
          ...p.japanese,
          defaultTeacherRoom: String(p?.japanese?.defaultTeacherRoom ?? "").trim() || "YY305",
          classCount: cc,
          classNames: normalizeClassNames(existing, cc),
          classRooms: normalizeClassRooms(p?.japanese?.classRooms ?? (p?.japanese?.defaultClassroom ? Array(cc).fill(String(p.japanese.defaultClassroom)) : undefined), cc),
        },
      };
      return fixed;
    });

    return {
      programs: migratedPrograms,
      items: parsed.items,
      selectedProgramId: parsed.selectedProgramId ?? (parsed.programs[0]?.id ?? null),
      lastUpdated: parsed.lastUpdated ?? Date.now(),
    };
  } catch {
    const p = makeNewProgram();
    return { programs: [p], items: [], selectedProgramId: p.id, lastUpdated: Date.now() };
  }
}

function saveState(state: AppState) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function programTypeLabel(t: ProgramType) {
  return t === "RSJP" ? "RSJP（レギュラー）" : "カスタム";
}

function categoryLabel(c: Category) {
  switch (c) {
    case "JapaneseClass":
      return "日本語講座";
    case "Orientation":
      return "オリエン";
    case "Escort":
      return "引率";
    case "CampusTour":
      return "キャンパスツアー";
    case "Cultural":
      return "文化体験";
    case "CompanyVisit":
      return "企業訪問";
    case "BuddyLunch":
      return "バディランチ";
    case "Ceremony":
      return "修了式";
    case "Other":
      return "その他";
  }
}

function categoryLabelByLang(c: Category, lang: ExportLang) {
  if (lang === "en") {
    switch (c) {
      case "JapaneseClass":
        return "Japanese Class";
      case "Orientation":
        return "Orientation";
      case "Escort":
        return "Escort";
      case "CampusTour":
        return "Campus Tour";
      case "Cultural":
        return "Cultural Experience";
      case "CompanyVisit":
        return "Company Visit";
      case "BuddyLunch":
        return "Buddy Lunch";
      case "Ceremony":
        return "Completion Ceremony";
      case "Other":
        return "Other";
    }
  }
  return categoryLabel(c);
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
  ];
  let out = s;
  for (const [r, v] of repl) out = out.replace(r, v);
  return out;
}


function normalizeHHMM(v: string) {
  const mins = parseTimeToMinutes(v);
  return mins === null ? v : minutesToTime(mins);
}

function makeItemBase(program: Program, date: string): Omit<ScheduleItem, "id"> {
  return {
    programId: program.id,
    date,
    startTime: "09:00",
    endTime: "10:00",
    category: "Other",
    title: "",
    location: "",
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
      it.location,
      it.notes,
    ].join("|");
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(it);
  }
  return out;
}


function normalizeItemsForProgram(items: ScheduleItem[], programId: string) {
  // 1) 完全一致の重複を削除
  let out = dedupeItemsForProgram(items, programId);

  // 2) 文化体験は「同じ日には1つだけ」にする（時間やタイトルが違っても2つ目以降は削除）
  //    ※残すのは「開始時刻が早いもの」優先
  const byDate = new Map<string, ScheduleItem[]>();
  for (const it of out) {
    if (it.programId !== programId) continue;
    if (it.category !== "Cultural") continue;
    const list = byDate.get(it.date) ?? [];
    list.push(it);
    byDate.set(it.date, list);
  }

  const removeIds = new Set<string>();
  for (const [date, list] of byDate.entries()) {
    if (list.length <= 1) continue;
    const sorted = list
      .slice()
      .sort((a, b) => {
        const ta = parseTimeToMinutes(a.startTime) ?? 99999;
        const tb = parseTimeToMinutes(b.startTime) ?? 99999;
        if (ta !== tb) return ta - tb;
        return String(a.id).localeCompare(String(b.id));
      });
    for (const it of sorted.slice(1)) removeIds.add(it.id);
  }

  if (removeIds.size > 0) {
    out = out.filter((it) => !removeIds.has(it.id));
  }

  return out;
}

function buildAutoItems(program: Program): ScheduleItem[] {
  const dates = eachDateInclusive(program.startDate, program.endDate);
  if (dates.length === 0) return [];
  const firstDate = dates[0];
  const lastDate = dates[dates.length - 1];

  const out: ScheduleItem[] = [];

  // 初日テンプレ
  out.push({
    id: uid("item"),
    ...makeItemBase(program, firstDate),
    startTime: "09:00",
    endTime: "10:00",
    category: "Escort",
    title: "ホテル/寮→大学 引率",
    location: "（集合場所を入力）",
    roomNeeded: "Yes",
    buddyCount: 0,
    kvhRequired: "No",
    kvhCount: 0,
    transportMode: "None",
    arrangementsNeeded: "No",
    generated: true,
    generatedKind: "Auto",
  });

  out.push({
    id: uid("item"),
    ...makeItemBase(program, firstDate),
    startTime: "10:30",
    endTime: "11:30",
    category: "Orientation",
    title: "オリエンテーション（1時間）",
    location: "（教室を入力）",
    roomNeeded: "Yes",
    buddyCount: 0,
    kvhRequired: "No",
    kvhCount: 0,
    transportMode: "OnCampus",
    arrangementsNeeded: "Yes",
    generated: true,
    generatedKind: "Auto",
  });

  out.push({
    id: uid("item"),
    ...makeItemBase(program, firstDate),
    startTime: "11:30",
    endTime: "12:20",
    category: "CampusTour",
    title: "バディによるキャンパスツアー",
    location: "（集合場所を入力）",
    roomNeeded: "Yes",
    buddyCount: program.defaultBuddyCount,
    kvhRequired: "No",
    kvhCount: 0,
    transportMode: "OnCampus",
    arrangementsNeeded: "No",
    generated: true,
    generatedKind: "Auto",
  });

  // 日本語講座（平日 + 初日は「上書きがあれば入る」）
  for (const date of dates) {
    const override = program.japaneseOverrides[date];
    const baseEnabled = program.japanese.enabled;

    const enabled = override ? override.enabled : baseEnabled && isWeekday(date) && date !== firstDate; // 初日は自動では入れない
    if (!enabled) continue;

    const startTime = normalizeHHMM(override?.startTime ?? program.japanese.startTime);
    const lesson = override?.lessonMinutes ?? program.japanese.lessonMinutes;
    const brk = override?.breakMinutes ?? program.japanese.breakMinutes;
    const periods = clamp(override?.periods ?? program.japanese.periods, 1, 3);
    const classCount = clamp(override?.classCount ?? program.japanese.classCount, 1, 20);

    const classRoomsBase = normalizeClassRooms(program.japanese.classRooms ?? (program.japanese.defaultClassroom ? Array(clamp(program.japanese.classCount, 1, 50)).fill(program.japanese.defaultClassroom) : undefined), clamp(program.japanese.classCount, 1, 50));
    const defaultTeacherRoom = program.japanese.defaultTeacherRoom.trim();

    const classrooms = (override?.classrooms ?? []).map((s) => s.trim()).filter(Boolean);
    const teacherRooms = (override?.teacherRooms ?? []).map((s) => s.trim()).filter(Boolean);

    let curStart = startTime;
    for (let p = 1; p <= periods; p++) {
      const curEnd = addMinutes(curStart, lesson);

      for (let c = 1; c <= classCount; c++) {
        const room = classrooms[c - 1] ?? classRoomsBase[c - 1] ?? "";
        const tRoom = teacherRooms[c - 1] ?? defaultTeacherRoom ?? "";
        const className = (program.japanese.classNames?.[c - 1] ?? defaultClassNameFor(c)).trim() || defaultClassNameFor(c);

        out.push({
          id: uid("item"),
          ...makeItemBase(program, date),
          startTime: curStart,
          endTime: curEnd,
          category: "JapaneseClass",
          title: `日本語講座（${p}コマ目） ${className}`,
          location: room,
          roomNeeded: "Yes",
          buddyCount: 0,
          kvhRequired: "No",
          kvhCount: 0,
          transportMode: "OnCampus",
          arrangementsNeeded: "Yes",
          notes: tRoom ? `講師控室: ${tRoom}` : "",
          classIndex: c,
          periodIndex: p,
          generated: true,
          generatedKind: "Auto",
        });
      }

      curStart = addMinutes(curEnd, brk);
    }
  }

  // 最終日：修了式
  out.push({
    id: uid("item"),
    ...makeItemBase(program, lastDate),
    startTime: normalizeHHMM(program.ceremony.startTime),
    endTime: addMinutes(normalizeHHMM(program.ceremony.startTime), program.ceremony.durationMinutes),
    category: "Ceremony",
    title: "修了式",
    location: program.ceremony.location ?? "",
    roomNeeded: "Yes",
    buddyCount: 0,
    kvhRequired: "No",
    kvhCount: 0,
    transportMode: "OnCampus",
    arrangementsNeeded: "Yes",
    notes: "",
    generated: true,
    generatedKind: "Auto",
  });

  return out;
}

function sortItems(items: ScheduleItem[]) {
  const key = (it: ScheduleItem) => `${it.date} ${it.startTime} ${it.endTime} ${it.category} ${it.title}`;
  // ★FIX: key(a) < key(a) になっていたので修正
  return [...items].sort((a, b) => (key(a) < key(b) ? -1 : key(a) > key(b) ? 1 : 0));
}

function buildCSV(program: Program, items: ScheduleItem[]) {
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
        DayOfWeek: dayOfWeekJP(it.date),
        StartTime: it.startTime,
        EndTime: it.endTime,
        Category: it.category,
        Title: it.title,
        Location: it.location,
        RoomNeeded: it.roomNeeded,
        StudentsCount: String(it.studentsCount ?? program.studentsCount),
        BuddyCount: String(it.buddyCount ?? 0),
        KVHRequired: it.kvhRequired,
        KVHCount: String(it.kvhCount ?? 0),
        TransportMode: it.transportMode,
        BusCompany: it.transportMode === "Bus" ? it.busCompany : "",
        BusCount: it.transportMode === "Bus" ? String(it.busCount ?? 1) : "",
        BusTripType: it.transportMode === "Bus" ? it.busTripType : "",
        BusPickup: it.transportMode === "Bus" ? it.busPickup : "",
        BusDropoff: it.transportMode === "Bus" ? it.busDropoff : "",
        ArrangementsNeeded: it.arrangementsNeeded,
        Notes: it.notes ?? "",
      };

      return headers.map((h) => csvEscape(row[h] ?? "")).join(",");
    });

  return [headers.join(","), ...rows].join("\n");
}

export default function ScheduleApp() {
  const [state, setState] = useState<AppState>(() => loadState());
  const masters = useMemo(() => makeDefaultEventMasters(), []);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const selectedProgram = useMemo(() => {
    return state.programs.find((p) => p.id === state.selectedProgramId) ?? null;
  }, [state.programs, state.selectedProgramId]);

  const selectedProgramItems = useMemo(() => {
    if (!selectedProgram) return [];
    const list = state.items.filter((it) => it.programId === selectedProgram.id);
    return list.sort((a, b) => {
      const ka = `${a.date} ${a.startTime} ${a.endTime} ${a.category} ${a.title}`;
      const kb = `${b.date} ${b.startTime} ${b.endTime} ${b.category} ${b.title}`;
      return ka < kb ? -1 : ka > kb ? 1 : 0;
    });
  }, [state.items, selectedProgram]);

  const datesInProgram = useMemo(() => {
    if (!selectedProgram) return [];
    return eachDateInclusive(selectedProgram.startDate, selectedProgram.endDate);
  }, [selectedProgram]);

  const [activeTab, setActiveTab] = useState<"program" | "calendar" | "share">("program");
  const [exportLang, setExportLang] = useState<ExportLang>("ja");
  const [programQuery, setProgramQuery] = useState("");
  const filteredPrograms = useMemo(() => {
    const q = programQuery.trim().toLowerCase();
    if (!q) return state.programs;
    return state.programs.filter((p) => p.name.toLowerCase().includes(q));
  }, [state.programs, programQuery]);


  // 日付からイベント追加
  const [selectedDate, setSelectedDate] = useState<string>("");
  const [addCategory, setAddCategory] = useState<Category>("Cultural");
  const [useMaster, setUseMaster] = useState<boolean>(true);
  const [masterId, setMasterId] = useState<string>(masters.find((m) => m.category === "Cultural")?.id ?? masters[0]?.id ?? "");
  const [freeTitle, setFreeTitle] = useState<string>("");
  const [location, setLocation] = useState<string>("");
  const [roomNeeded, setRoomNeeded] = useState<YesNo>("Yes");
  const [studentsCount, setStudentsCount] = useState<number>(selectedProgram?.studentsCount ?? 0);
  const [buddyCount, setBuddyCount] = useState<number>(selectedProgram?.defaultBuddyCount ?? 5);
  const [kvhRequired, setKVHRequired] = useState<YesNo>("Yes");
  const [kvhCount, setKVHCount] = useState<number>(1);
  const [transportMode, setTransportMode] = useState<TransportMode>("Bus");
  const [busCompany, setBusCompany] = useState<string>("");
  const [busCount, setBusCount] = useState<number>(1);
  const [busTripType, setBusTripType] = useState<BusTripType>("OneWay");
  const [busPickup, setBusPickup] = useState<string>("");
  const [busDropoff, setBusDropoff] = useState<string>("");
  const [arrangementsNeeded, setArrangementsNeeded] = useState<YesNo>("Yes");
  const [notes, setNotes] = useState<string>("");

  const [eventStartTime, setEventStartTime] = useState<string>("13:10");
  const [eventEndTime, setEventEndTime] = useState<string>("16:00");

  // 日本語講座：日付ごとの上書き編集
  const [overrideDate, setOverrideDate] = useState<string>("");
  const [ovEnabled, setOvEnabled] = useState<boolean>(true);
  const [ovStartTime, setOvStartTime] = useState<string>("09:00");
  const [ovLesson, setOvLesson] = useState<number>(50);
  const [ovBreak, setOvBreak] = useState<number>(10);
  const [ovPeriods, setOvPeriods] = useState<number>(3);
  const [ovClassCount, setOvClassCount] = useState<number>(1);
  const [ovClassrooms, setOvClassrooms] = useState<string>("");
  const [ovTeacherRooms, setOvTeacherRooms] = useState<string>("");

  // calendar tab: preview srcDoc
  const [calendarPreviewEnabled, setCalendarPreviewEnabled] = useState<boolean>(true);

  // localStorage 自動保存
  useEffect(() => {
    saveState(state);
  }, [state]);

  useEffect(() => {
    if (!selectedProgram) return;
    setStudentsCount(selectedProgram.studentsCount);
    setBuddyCount(selectedProgram.defaultBuddyCount);
  }, [selectedProgram?.id]);

  useEffect(() => {
    if (!selectedProgram || !selectedDate) return;

    const m = masters.find((x) => x.id === masterId);
    const isCultural = addCategory === "Cultural";

    if (isCultural) {
      setEventStartTime("13:10");
      setEventEndTime("16:00");
      setRoomNeeded("Yes");
      setKVHRequired("Yes");
      setKVHCount(1);
      setTransportMode("Bus");
      setArrangementsNeeded("Yes");
      setFreeTitle("");
      setNotes(m?.defaultNotes ?? "");
      if (useMaster && m) {
        // nothing else
      }
    } else {
      setEventStartTime("13:10");
      setEventEndTime("14:45");
      setRoomNeeded("Yes");
      setKVHRequired("No");
      setKVHCount(0);
      setTransportMode("OnCampus");
      setArrangementsNeeded("No");
      setFreeTitle("");
      setNotes("");
    }
  }, [selectedDate, addCategory, masterId, useMaster, masters, selectedProgram]);

  function updateProgram(patch: Partial<Program>) {
    if (!selectedProgram) return;
    const now = Date.now();
    setState((prev) => ({
      ...prev,
      programs: prev.programs.map((p) => (p.id === selectedProgram.id ? { ...p, ...patch, lastUpdated: now } : p)),
      lastUpdated: now,
    }));
  }

  function createProgram() {
    const p = makeNewProgram();
    const now = Date.now();
    setState((prev) => ({
      ...prev,
      programs: [p, ...prev.programs],
      selectedProgramId: p.id,
      lastUpdated: now,
    }));
    setActiveTab("program");
  }

  function deleteProgram(programId: string) {
    const ok = window.confirm("このプログラムを削除します。関連する予定データも削除されます。よろしいですか？");
    if (!ok) return;
    const now = Date.now();
    setState((prev) => {
      const nextPrograms = prev.programs.filter((p) => p.id !== programId);
      const nextItems = prev.items.filter((it) => it.programId !== programId);
      const nextSelected = prev.selectedProgramId === programId ? (nextPrograms[0]?.id ?? null) : prev.selectedProgramId;
      return { programs: nextPrograms, items: nextItems, selectedProgramId: nextSelected, lastUpdated: now };
    });
  }

  function generateAuto() {
    if (!selectedProgram) return;
    const auto = buildAutoItems(selectedProgram);
    const now = Date.now();
    setState((prev) => ({
      ...prev,
      items: normalizeItemsForProgram([...removeGeneratedAuto(prev.items, selectedProgram.id), ...auto], selectedProgram.id),
      lastUpdated: now,
    }));
  }

  function removeDuplicatesNow() {
    if (!selectedProgram) return;
    setState((prev) => ({
      ...prev,
      items: normalizeItemsForProgram(prev.items, selectedProgram.id),
      lastUpdated: Date.now(),
    }));
    alert("重複している予定を整理しました。\n・完全一致の重複を削除\n・文化体験は同じ日に1つだけ残す\n※文化体験（茶道）が複数入った場合の修正に使えます。");
  }

  function addManualEvent() {
    if (!selectedProgram) return;
    if (!selectedDate) {
      alert("日付を選んでください。");
      return;
    }

    const st = normalizeHHMM(eventStartTime);
    const et = normalizeHHMM(eventEndTime);

    const master = useMaster ? masters.find((m) => m.id === masterId) : undefined;
    const resolvedTitle =
      useMaster && master ? master.title : freeTitle.trim() ? freeTitle.trim() : "イベント";

    const base = makeItemBase(selectedProgram, selectedDate);

    const item: ScheduleItem = {
      id: uid("item"),
      ...base,
      startTime: st,
      endTime: et,
      category: addCategory,
      title: resolvedTitle,
      location,
      roomNeeded,
      studentsCount: Number.isFinite(studentsCount) ? studentsCount : selectedProgram.studentsCount,
      buddyCount: Number.isFinite(buddyCount) ? buddyCount : selectedProgram.defaultBuddyCount,
      kvhRequired,
      kvhCount: kvhRequired === "Yes" ? clamp(kvhCount, 1, 50) : 0,
      transportMode,
      busCompany: transportMode === "Bus" ? busCompany : "",
      busCount: transportMode === "Bus" ? clamp(busCount, 1, 20) : 0,
      busTripType: transportMode === "Bus" ? busTripType : "OneWay",
      busPickup: transportMode === "Bus" ? busPickup : "",
      busDropoff: transportMode === "Bus" ? busDropoff : "",
      arrangementsNeeded,
      notes,
      generated: false,
    };

    const now = Date.now();
    setState((prev) => ({ ...prev, items: [...prev.items, item], lastUpdated: now }));
  }

  function updateItem(itemId: string, patch: Partial<ScheduleItem>) {
    const now = Date.now();
    setState((prev) => ({
      ...prev,
      items: selectedProgram ? normalizeItemsForProgram(prev.items.map((it) => (it.id === itemId ? { ...it, ...patch } : it)), selectedProgram.id) : prev.items.map((it) => (it.id === itemId ? { ...it, ...patch } : it)),
      lastUpdated: now,
    }));
  }

  function deleteItem(itemId: string) {
    const now = Date.now();
    setState((prev) => ({ ...prev, items: selectedProgram ? normalizeItemsForProgram(prev.items.filter((it) => it.id !== itemId), selectedProgram.id) : prev.items.filter((it) => it.id !== itemId), lastUpdated: now }));
  }

  function exportCSV() {
    if (!selectedProgram) return;
    const csv = "\uFEFF" + buildCSV(selectedProgram, state.items);
    const safeName = selectedProgram.name.replace(/[\\/:*?"<>|]/g, "_");
    downloadTextFile(`${safeName}_schedule.csv`, csv, "text/csv;charset=utf-8");
  }

  function escapeHTML(s: string) {
    return (s ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function exportICS() {
    if (!selectedProgram) return;
    const ics = buildICSAsiaTokyo(selectedProgram, state.items, exportLang);
    const safeName = selectedProgram.name.replace(/[\\/:*?"<>|]/g, "_");
    const langTag = exportLang === "en" ? "EN" : "JA";
    downloadTextFile(`${safeName}_schedule_${langTag}_JST.ics`, ics, "text/calendar;charset=utf-8");
  }

  function exportCalendarHTML() {
    if (!selectedProgram) return;
    const html = buildCalendarHTMLMultiMonth(selectedProgram, state.items, exportLang);
    const safeName = selectedProgram.name.replace(/[\\/:*?"<>|]/g, "_");
    const langTag = exportLang === "en" ? "EN" : "JA";
    downloadTextFile(`${safeName}_calendar_${langTag}.html`, html, "text/html;charset=utf-8");
  }

  // --- ICS (iCalendar) helpers (Asia/Tokyo fixed) ---
  function toICSLocalDateTime(dateISO: string, timeHHMM: string) {
    // YYYYMMDDTHHMMSS
    const ymd = dateISO.replaceAll("-", "");
    const hhmm = timeHHMM.replace(":", "");
    return `${ymd}T${hhmm}00`;
  }

  function icsEscape(s: string) {
    return (s ?? "")
      .replaceAll("\\", "\\\\")
      .replaceAll("\n", "\\n")
      .replaceAll(",", "\\,")
      .replaceAll(";", "\\;");
  }

  function buildICSAsiaTokyo(program: any, items: any[], lang: ExportLang = "ja") {
    const lines: string[] = [];
    lines.push("BEGIN:VCALENDAR");
    lines.push("VERSION:2.0");
    lines.push("PRODID:-//RSJP Scheduler//JP//EN");
    lines.push("CALSCALE:GREGORIAN");
    lines.push("METHOD:PUBLISH");

    // Asia/Tokyo (JST, no DST)
    lines.push("BEGIN:VTIMEZONE");
    lines.push("TZID:Asia/Tokyo");
    lines.push("X-LIC-LOCATION:Asia/Tokyo");
    lines.push("BEGIN:STANDARD");
    lines.push("TZOFFSETFROM:+0900");
    lines.push("TZOFFSETTO:+0900");
    lines.push("TZNAME:JST");
    lines.push("DTSTART:19700101T000000");
    lines.push("END:STANDARD");
    lines.push("END:VTIMEZONE");

    const list = items
      .filter((it) => it.programId === program.id)
      .slice()
      .sort((a, b) => {
        const ka = `${a.date} ${a.startTime} ${a.endTime} ${a.category} ${a.title}`;
        const kb = `${b.date} ${b.startTime} ${b.endTime} ${b.category} ${b.title}`;
        return ka < kb ? -1 : ka > kb ? 1 : 0;
      });

    // DTSTAMP must be UTC with Z
    const dtstamp = new Date().toISOString().replaceAll("-", "").replaceAll(":", "").replace(".000", ""); // YYYYMMDDTHHMMSSZ

    for (const it of list) {
      const uidLine = `${it.id}@rsjp-scheduler`;
      const titleFor = simpleExportTranslate(String(it.title ?? ""), lang);
      const summary = lang === "en" ? `${program.name} | ${titleFor}` : `${program.name}｜${String(it.title ?? "")}`;
      const location = simpleExportTranslate(String(it.location || ""), lang);

      const descParts: string[] = [];
      descParts.push(`${lang === "en" ? "Category" : "カテゴリ"}: ${categoryLabelByLang(it.category as Category, lang)}`);
      if (it.roomNeeded) descParts.push(`${lang === "en" ? "Room setup" : "教室手配"}: ${ynLabel(it.roomNeeded, lang)}`);
      if (typeof it.studentsCount === "number") descParts.push(`${lang === "en" ? "Students" : "留学生"}: ${it.studentsCount}`);
      if (typeof it.buddyCount === "number") descParts.push(`${lang === "en" ? "Buddies" : "バディ"}: ${it.buddyCount}`);
      if (it.kvhRequired) descParts.push(`KVH: ${ynLabel(it.kvhRequired, lang)}${String(it.kvhRequired) === "Yes" ? ` (${it.kvhCount})` : ""}`);
      if (it.transportMode) descParts.push(`${lang === "en" ? "Transport" : "移動"}: ${it.transportMode}`);
      if (it.notes) descParts.push(`${lang === "en" ? "Notes" : "備考"}: ${simpleExportTranslate(String(it.notes ?? ""), lang)}`);

      const dtStart = toICSLocalDateTime(it.date, it.startTime);
      const dtEnd = toICSLocalDateTime(it.date, it.endTime);

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

  // --- Monthly calendar HTML (multi-month) ---
  function buildCalendarHTMLMultiMonth(program: any, items: any[], lang: ExportLang = "ja") {
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
        const title = escapeHTML(simpleExportTranslate(String(it.title ?? ""), lang));
        const loc = simpleExportTranslate(String(it.location ?? "").trim(), lang);
        const locationLabelFor = (catRaw: any, locRaw: string) => {
          const cat = String(catRaw ?? "");
          const loc = String(locRaw ?? "").trim();
          if (!loc) return "";
          // 必須系はラベルを付けて見やすく
          if (cat === "Escort" || cat === "CampusTour" || cat === "Cultural" || cat === "CompanyVisit") return `集合: ${loc}`;
          if (cat === "Orientation" || cat === "JapaneseClass" || cat === "Ceremony") return `教室: ${loc}`;
          return `場所: ${loc}`;
        };

        // 旧：locLabel（互換のため残すが、表示はlocLineに移行）
        const locLabel = loc ? `（${escapeHTML(loc)}）` : "";
        const locLine = loc ? `<div class="sub">${escapeHTML(locationLabelFor(it.category, loc))}</div>` : "";
        const noteLine = String(it.notes ?? "").trim() ? `<div class="sub muted">${escapeHTML(String(it.notes ?? "").trim())}</div>` : "";
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
            <div class="it jphead"><span class="t">日本語講座</span> <span class="tm">${escapeHTML(minStart)}-${escapeHTML(maxEnd)}</span></div>
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
<html lang="ja">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>${escapeHTML(program.name)} カレンダー</title>
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
  <h1>${escapeHTML(program.name)}（${escapeHTML(programTypeLabel(program.type))}） カレンダー</h1>
  <div class="topmeta">期間: ${program.startDate} → ${program.endDate} / 留学生: ${program.studentsCount}</div>
  ${monthSections.join("\n")}
</body>
</html>`;
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
        const parsed = JSON.parse(txt) as AppState;
        if (!parsed || !Array.isArray(parsed.programs) || !Array.isArray(parsed.items)) {
          alert("JSONの形式が正しくありません。");
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

  const headerStyle: React.CSSProperties = {
    display: "flex",
    gap: 8,
    alignItems: "center",
    justifyContent: "space-between",
    padding: "12px 12px",
    borderBottom: "1px solid #ddd",
    position: "sticky",
    top: 0,
    background: "#fff",
    zIndex: 5,
  };

  const card: React.CSSProperties = {
    border: "1px solid #ddd",
    borderRadius: 10,
    padding: 12,
    background: "#fff",
  };

  // calendar preview HTML
  const calendarHTML = useMemo(() => {
    if (!selectedProgram) return "";
    return buildCalendarHTMLMultiMonth(selectedProgram, state.items);
  }, [selectedProgram?.id, state.items, state.lastUpdated]);

  // item groups by date (for calendar tab)
  const itemsByDate = useMemo(() => {
    const map = new Map<string, ScheduleItem[]>();
    for (const it of selectedProgramItems) {
      if (!map.has(it.date)) map.set(it.date, []);
      map.get(it.date)!.push(it);
    }
    for (const [k, v] of map.entries()) {
      map.set(k, sortItems(v));
    }
    return map;
  }, [selectedProgramItems]);

  const showBusFields = transportMode === "Bus";

  const ActionButton = ({
    label,
    hint,
    onClick,
    primary,
    disabled,
  }: {
    label: string;
    hint: string;
    onClick: () => void;
    primary?: boolean;
    disabled?: boolean;
  }) => (
    <div className="rsjpBtnStack">
      <button
        className={primary ? "rsjpPrimary" : undefined}
        onClick={onClick}
        disabled={disabled}
        title={hint}
        style={
          disabled
            ? { opacity: 0.55, cursor: "not-allowed" as const }
            : undefined
        }
      >
        {label}
      </button>
      <div className="rsjpHint">{hint}</div>
    </div>
  );


  return (
    <div className="rsjpApp" style={{
        fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif",
        padding: 12,
        maxWidth: 1300,
        margin: "0 auto",
      }}
    >
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
        /* cards (we hook by inline style fallback: apply to big wrappers via first-level children) */
        .rsjpApp .rsjpGlass {
          background: rgba(255,255,255,0.72);
          border: 1px solid rgba(15,23,42,0.10);
          border-radius: 16px;
          box-shadow: 0 10px 30px rgba(2, 6, 23, 0.12);
          backdrop-filter: blur(10px);
        }

        /* controls: do NOT force width, only prevent overflow */
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

        /* buttons */
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

        /* tab pills (only styling; layout untouched) */
        .rsjpApp .rsjpTabRow button {
          border-radius: 999px !important;
        }
        .rsjpApp .rsjpTabRow button.rsjpActiveTab {
          background: rgba(161,0,0,0.12);
          border-color: rgba(161,0,0,0.25);
        }

        /* Prevent flex children overflow causing overlaps */
        .rsjpApp [style*="display: flex"] { min-width: 0; }
        .rsjpApp [style*="display: flex"] > * { min-width: 0; }

/* --- Pro polish additions (v2) --- */
.rsjpHeader{ position: sticky; top: 0; z-index: 20; }

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

.rsjpSidebar{
  height: calc(100vh - 120px);
  overflow: auto;
  min-width: 0;
}
@media (max-width: 980px){
  .rsjpSidebar{ height: auto; max-height: 60vh; }
}

.rsjpProgramItem{
  padding: 10px;
  border-radius: 12px;
  border: 1px solid rgba(15,23,42,0.12);
  background: rgba(255,255,255,0.92);
  cursor: pointer;
  transition: transform .06s ease, box-shadow .12s ease, border-color .12s ease, background .12s ease;
}
.rsjpProgramItem:hover{
  transform: translateY(-1px);
  box-shadow: 0 12px 24px rgba(2,6,23,0.10);
}
.rsjpProgramItem[data-active="true"]{
  border-color: rgba(37,99,235,0.55);
  background: rgba(239,246,255,0.92);
  box-shadow: 0 14px 28px rgba(2,6,23,0.12);
}

.rsjpTabRow{
  display:flex;
  gap: 8px;
  flex-wrap: wrap;
  padding: 6px;
  border-radius: 999px;
  background: rgba(15,23,42,0.04);
  border: 1px solid rgba(15,23,42,0.08);
}
.rsjpTabBtn{
  height: 36px;
  padding: 0 14px;
  border-radius: 999px;
  border: 1px solid transparent;
  background: transparent;
  font-weight: 900;
}
.rsjpTabBtnActive{
  background: rgba(255,255,255,0.94);
  border-color: rgba(15,23,42,0.12);
  box-shadow: 0 10px 22px rgba(2,6,23,0.10);
}

.rsjpDanger{
  background: rgba(220,38,38,0.10);
  border-color: rgba(220,38,38,0.25);
  color: #b91c1c;
}
.rsjpDanger:hover{ background: rgba(220,38,38,0.14); }

.rsjpTopActions{ display:flex; gap: 14px; align-items:flex-start; flex-wrap:wrap; }
.rsjpBtnStack{ display:flex; flex-direction:column; align-items:center; gap:3px; }
.rsjpHint{ font-size: 11px; color:#475569; opacity:0.85; line-height:1.25; text-align:center; max-width: 160px; }
.rsjpSidebarSearch{ width: 100%; padding: 9px 10px; border-radius: 12px; border: 1px solid rgba(15,23,42,0.18); background: rgba(255,255,255,0.85); outline: none; }
.rsjpSidebarSearch:focus{ border-color: rgba(15,23,42,0.45); box-shadow: 0 0 0 4px rgba(2,6,23,0.06); }

.rsjpProgramItem{ border:1px solid rgba(15,23,42,0.12); border-radius:14px; padding:12px; cursor:pointer; background: rgba(255,255,255,0.82); transition: transform .08s ease, box-shadow .08s ease, border-color .08s ease; }
.rsjpProgramItem:hover{ transform: translateY(-1px); box-shadow: 0 18px 34px rgba(2,6,23,0.08); border-color: rgba(15,23,42,0.20); }
.rsjpProgramItemActive{ border: 2px solid rgba(15,23,42,0.55); background: rgba(255,255,255,0.95); }
.rsjpProgramMeta{ font-size: 12px; opacity: 0.78; }
.rsjpTabHelp{ font-size: 11px; color:#475569; opacity:0.82; margin-top:6px; }
`}</style>

      <div className="rsjpGlass rsjpHeader" style={headerStyle}>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          <div style={{ fontWeight: 800, fontSize: 18 }}>RSJP / カスタム　スケジュール作成（MVP）</div>
          <div style={{ fontSize: 12, opacity: 0.7 }}>自動保存 + JSON共有 + CSV出力 + カレンダープレビュー</div>
        </div>
        <div className="rsjpTopActions">
          <ActionButton
            label="➕ 新規プログラム"
            hint="新しいプログラムを作成"
            onClick={createProgram}
            primary
          />
          <ActionButton
            label="⬇️ JSON書き出し"
            hint="データを保存・共有"
            onClick={exportJSON}
          />
          <ActionButton
            label="⬆️ JSON読み込み"
            hint="受け取ったJSONで復元"
            onClick={openImportDialog}
          />
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
        {/* 左：プログラム一覧 */}
        <div className="rsjpSidebar" style={{ ...card }}>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>プログラム一覧</div>
            <input
              className="rsjpSidebarSearch"
              placeholder="検索（プログラム名）"
              value={programQuery}
              onChange={(e) => setProgramQuery(e.target.value)}
            />
            <div className="rsjpTabHelp" style={{ marginTop: 6 }}>
              クリックで選択。右上の「JSON書き出し」でバックアップできます。
            </div>

          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {filteredPrograms.map((p) => (
              <div
                key={p.id}
                className={`rsjpProgramItem ${p.id === state.selectedProgramId ? "rsjpProgramItemActive" : ""}`}
                onClick={() => setState((prev) => ({ ...prev, selectedProgramId: p.id }))}
              >
                <div style={{ fontWeight: 700, marginBottom: 2 }}>{p.name}</div>
                <div style={{ fontSize: 12, opacity: 0.75 }}>{programTypeLabel(p.type)}</div>
                <div style={{ fontSize: 12, opacity: 0.75 }}>
                  {p.startDate} → {p.endDate}
                </div>
                <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      if (window.confirm("このプログラムを削除します。よろしいですか？")) deleteProgram(p.id);
                    }}
                    style={{ fontSize: 12 }}
                  >
                    削除
                  </button>
                </div>
              </div>
            ))}
            {filteredPrograms.length === 0 && (
              <div className="rsjpHint" style={{ textAlign: "left", maxWidth: "none", padding: "8px 2px" }}>
                該当するプログラムがありません。
              </div>
            )}
          </div>

          <div style={{ marginTop: 14, fontSize: 12, opacity: 0.7, lineHeight: 1.4 }}>
            <div>・自動生成：初日テンプレ / 平日日本語講座 / 最終日修了式</div>
            <div>・文化体験/企業訪問/バディランチは「日程・予定」タブで追加</div>
            <div>・カレンダーの見た目は「カレンダープレビュー」でスクショと同等表示</div>
          </div>
        </div>

        {/* 右：メイン */}
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          {!selectedProgram ? (
            <div style={card}>プログラムを選択してください。</div>
          ) : (
            <>
              {/* タブ */}
              <div style={{ ...card }} className="rsjpTabRow">
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                  <button
                    className={`rsjpTabBtn ${activeTab === "program" ? "rsjpTabBtnActive" : ""}`}
                    onClick={() => setActiveTab("program")}
                  >
                    プログラム設定
                  </button>
                  <button
                    className={`rsjpTabBtn ${activeTab === "calendar" ? "rsjpTabBtnActive" : ""}`}
                    onClick={() => setActiveTab("calendar")}
                  >
                    日程・予定
                  </button>
                  <button
                    className={`rsjpTabBtn ${activeTab === "share" ? "rsjpTabBtnActive" : ""}`}
                    onClick={() => setActiveTab("share")}
                  >
                    共有・出力
                  </button>
                </div>

                <div style={{ marginLeft: "auto", display: "flex", gap: 14, flexWrap: "wrap", alignItems: "flex-start" }}>
                  <ActionButton label="⚡ 自動生成（再生成）" hint="初日/講座/最終日を作成" onClick={generateAuto} primary />
                  <ActionButton label="CSV出力" hint="予定を表で保存" onClick={exportCSV} />
                  <ActionButton label="ICS出力" hint="Google/Outlookに追加" onClick={exportICS} />
                  <ActionButton label="HTML出力" hint="印刷・共有用カレンダー" onClick={exportCalendarHTML} />
                </div>
              </div>

{activeTab === "program" && (
                <div style={{ ...card }}>
                  <div style={{ fontWeight: 800, marginBottom: 10 }}>プログラム設定</div>

                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                    <div>
                      <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>プログラム名</label>
                      <input value={selectedProgram.name} onChange={(e) => updateProgram({ name: e.target.value })} style={{ width: "100%" }} />
                    </div>

                    <div>
                      <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>種別</label>
                      <select value={selectedProgram.type} onChange={(e) => updateProgram({ type: e.target.value as ProgramType })} style={{ width: "100%" }}>
                        <option value="RSJP">RSJP（レギュラー）</option>
                        <option value="Custom">カスタム</option>
                      </select>
                    </div>

                    <div>
                      <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>開始日</label>
                      <input type="date" value={selectedProgram.startDate} onChange={(e) => updateProgram({ startDate: e.target.value })} style={{ width: "100%" }} />
                    </div>

                    <div>
                      <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>終了日</label>
                      <input type="date" value={selectedProgram.endDate} onChange={(e) => updateProgram({ endDate: e.target.value })} style={{ width: "100%" }} />
                    </div>

                    <div>
                      <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>参加留学生人数（基本）</label>
                      <input
                        type="number"
                        value={selectedProgram.studentsCount}
                        onChange={(e) => updateProgram({ studentsCount: clamp(Number(e.target.value), 0, 9999) })}
                        style={{ width: "100%" }}
                      />
                    </div>

                    <div>
                      <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>バディ人数（イベント初期値）</label>
                      <input
                        type="number"
                        value={selectedProgram.defaultBuddyCount}
                        onChange={(e) => updateProgram({ defaultBuddyCount: clamp(Number(e.target.value), 0, 9999) })}
                        style={{ width: "100%" }}
                      />
                      <div style={{ fontSize: 12, opacity: 0.7, marginTop: 4 }}>※イベント追加時の初期値（デフォルト5）</div>
                    </div>
                  </div>

                  <hr style={{ margin: "14px 0" }} />

                  <div style={{ fontWeight: 800, marginBottom: 10 }}>日本語講座（デフォルト）</div>

                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 10 }}>
                    <div>
                      <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>有無</label>
                      <select
                        value={selectedProgram.japanese.enabled ? "Yes" : "No"}
                        onChange={(e) => updateProgram({ japanese: { ...selectedProgram.japanese, enabled: e.target.value === "Yes" } })}
                        style={{ width: "100%" }}
                      >
                        <option value="Yes">あり</option>
                        <option value="No">なし</option>
                      </select>
                    </div>

                    <div>
                      <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>開始時刻（通常）</label>
                      <input
                        value={selectedProgram.japanese.startTime}
                        onChange={(e) => updateProgram({ japanese: { ...selectedProgram.japanese, startTime: normalizeHHMM(e.target.value) } })}
                        placeholder="09:00"
                        style={{ width: "100%" }}
                      />
                    </div>

                    <div>
                      <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>授業分</label>
                      <select
                        value={String(selectedProgram.japanese.lessonMinutes)}
                        onChange={(e) => updateProgram({ japanese: { ...selectedProgram.japanese, lessonMinutes: Number(e.target.value) } })}
                        style={{ width: "100%" }}
                      >
                        <option value="50">50</option>
                        <option value="60">60</option>
                      </select>
                    </div>

                    <div>
                      <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>休憩分</label>
                      <input
                        type="number"
                        value={selectedProgram.japanese.breakMinutes}
                        onChange={(e) => updateProgram({ japanese: { ...selectedProgram.japanese, breakMinutes: clamp(Number(e.target.value), 0, 60) } })}
                        style={{ width: "100%" }}
                      />
                    </div>

                    <div>
                      <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>コマ数（1〜3）</label>
                      <input
                        type="number"
                        value={selectedProgram.japanese.periods}
                        onChange={(e) => updateProgram({ japanese: { ...selectedProgram.japanese, periods: clamp(Number(e.target.value), 1, 3) } })}
                        style={{ width: "100%" }}
                      />
                    </div>

                    <div>
                      <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>クラス数</label>
                      <input
                        type="number"
                        value={selectedProgram.japanese.classCount}
                        onChange={(e) => {
                          const nextCount = clamp(Number(e.target.value), 1, 50);
                          const nextNames = normalizeClassNames(selectedProgram.japanese.classNames, nextCount);
                          const nextRooms = normalizeClassRooms(selectedProgram.japanese.classRooms, nextCount);
                          updateProgram({ japanese: { ...selectedProgram.japanese, classCount: nextCount, classNames: nextNames, classRooms: nextRooms } });
                        }}
                        style={{ width: "100%" }}
                      />
                    </div>

                    <div style={{ gridColumn: "1 / -1", marginTop: 6, padding: 10, borderRadius: 10, border: "1px solid #eee", background: "#fafafa" }}>
                      <div style={{ fontWeight: 800, marginBottom: 6 }}>クラス名（デフォルト）</div>
                      <div style={{ fontSize: 12, opacity: 0.75, marginBottom: 8, lineHeight: 1.5 }}>
                        ・デフォルト順：嵐山 → 宇治 → 祇園 → 西陣 → 東山（5クラス）<br />
                        ・クラス数が5未満の場合は、嵐山から順に自動採用します
                      </div>

                      <div style={{ display: "grid", gridTemplateColumns: "repeat(5, minmax(0, 1fr))", gap: 8 }}>
                        {Array.from({ length: clamp(selectedProgram.japanese.classCount, 1, 50) }).slice(0, 10).map((_, i) => {
                          const idx1 = i + 1;
                          const cur = selectedProgram.japanese.classNames?.[i] ?? defaultClassNameFor(idx1);
                          return (
                            <div key={idx1}>
                              <div style={{ fontSize: 12, opacity: 0.75, marginBottom: 4 }}>クラス{idx1}</div>
                              <input
                                value={cur}
                                onChange={(e) => {
                                  const next = [...normalizeClassNames(selectedProgram.japanese.classNames, selectedProgram.japanese.classCount)];
                                  next[i] = e.target.value;
                                  updateProgram({ japanese: { ...selectedProgram.japanese, classNames: normalizeClassNames(next, selectedProgram.japanese.classCount) } });
                                }}
                                style={{ width: "100%" }}
                              />
                            </div>
                          );
                        })}
                      </div>

                      {selectedProgram.japanese.classCount > 10 && (
                        <div style={{ fontSize: 12, opacity: 0.7, marginTop: 8 }}>
                          ※この画面では最大10クラスまで編集表示しています（11クラス目以降は自動で「クラス11」等になります）
                        </div>
                      )}

                      <div style={{ display: "flex", gap: 8, marginTop: 10, flexWrap: "wrap" }}>
                        <button
                          onClick={() => {
                            const next = normalizeClassNames(undefined, selectedProgram.japanese.classCount);
                            updateProgram({ japanese: { ...selectedProgram.japanese, classNames: next } });
                          }}
                        >
                          デフォルト（嵐山〜）を適用
                        </button>
                      </div>
                    </div>


                    <div>
                      <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>デフォルト教室</label>
                      <input
                        value={selectedProgram.japanese.defaultClassroom}
                        onChange={(e) => updateProgram({ japanese: { ...selectedProgram.japanese, defaultClassroom: e.target.value } })}
                        style={{ width: "100%" }}
                      />
                    </div>

                    <div style={{ gridColumn: "1 / -1", marginTop: 10, padding: 10, borderRadius: 10, border: "1px solid #eee", background: "#fafafa" }}>
                      <div style={{ fontWeight: 800, marginBottom: 6 }}>クラス別教室（デフォルト）</div>
                      <div style={{ fontSize: 12, opacity: 0.75, marginBottom: 8, lineHeight: 1.5 }}>
                        ・例：嵐山=YY301、宇治=YY302、祇園=YY303… のようにクラスごとに教室を持てます<br />
                        ・特異日（例：教室変更）は「日本語講座：日付ごとの上書き」でその日だけ上書きしてください
                      </div>

                      <div style={{ display: "grid", gridTemplateColumns: "repeat(5, minmax(0, 1fr))", gap: 8 }}>
                        {Array.from({ length: clamp(selectedProgram.japanese.classCount, 1, 50) }).slice(0, 10).map((_, i) => {
                          const idx1 = i + 1;
                          const cur = selectedProgram.japanese.classRooms?.[i] ?? defaultClassRoomFor(idx1);
                          const cname = selectedProgram.japanese.classNames?.[i] ?? defaultClassNameFor(idx1);
                          return (
                            <div key={idx1}>
                              <div style={{ fontSize: 12, opacity: 0.75, marginBottom: 4 }}>
                                {cname}（クラス{idx1}）
                              </div>
                              <input
                                value={cur}
                                onChange={(e) => {
                                  const next = [...normalizeClassRooms(selectedProgram.japanese.classRooms, selectedProgram.japanese.classCount)];
                                  next[i] = e.target.value;
                                  updateProgram({ japanese: { ...selectedProgram.japanese, classRooms: normalizeClassRooms(next, selectedProgram.japanese.classCount) } });
                                }}
                                style={{ width: "100%" }}
                                placeholder={defaultClassRoomFor(idx1)}
                              />
                            </div>
                          );
                        })}
                      </div>

                      {selectedProgram.japanese.classCount > 10 && (
  <div style={{ fontSize: 12, opacity: 0.7, marginTop: 8 }}>
    ※この画面では最大10クラスまで編集表示しています（11クラス目以降は自動で YY300, YY301... を採用します）
  </div>
)}

                      <div style={{ display: "flex", gap: 8, marginTop: 10, flexWrap: "wrap" }}>
                        <button
                          onClick={() => {
                            const next = normalizeClassRooms(undefined, selectedProgram.japanese.classCount);
                            updateProgram({ japanese: { ...selectedProgram.japanese, classRooms: next } });
                          }}
                        >
                          デフォルト（YY301〜）を適用
                        </button>
                      </div>
                    </div>


                    <div>
                      <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>デフォルト講師控室</label>
                      <input
                        value={selectedProgram.japanese.defaultTeacherRoom}
                        onChange={(e) => updateProgram({ japanese: { ...selectedProgram.japanese, defaultTeacherRoom: e.target.value } })}
                        style={{ width: "100%" }}
                      />
                    </div>
                  </div>

                  <div style={{ fontSize: 12, opacity: 0.7, marginTop: 8, lineHeight: 1.5 }}>
                    ・自動生成では<strong>初日には日本語講座を自動で入れません</strong>（必要な場合は「初日に日本語講座を入れる」または「日付上書き」）
                    <br />
                    ・デフォルトは「平日（月〜金）」に入ります（上書きで例外OK）
                  </div>

                  <hr style={{ margin: "14px 0" }} />

                  <div style={{ fontWeight: 800, marginBottom: 10 }}>日本語講座：日付ごとの上書き（初日午後など）</div>

                  {/* ★追加：ワンタッチ */}
                  <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 10, flexWrap: "wrap" }}>
                    <button onClick={quickEnableJapaneseOnFirstDay} style={{ fontWeight: 700 }}>
                      初日に日本語講座を入れる（ワンタッチ）
                    </button>
                    <div style={{ fontSize: 12, opacity: 0.7 }}>
                      ※開始時刻だけ入力 → 他はデフォルト。押した瞬間に自動生成（再生成）まで実行します。
                    </div>
                  </div>

                  <div style={{ display: "grid", gridTemplateColumns: "260px 1fr", gap: 12 }}>
                    <div>
                      <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>日付</label>
                      <select
                        value={overrideDate}
                        onChange={(e) => {
                          const d = e.target.value;
                          setOverrideDate(d);
                          loadOverrideFromProgram(d);
                        }}
                        style={{ width: "100%" }}
                      >
                        <option value="">（選択）</option>
                        {datesInProgram.map((d) => (
                          <option key={d} value={d}>
                            {d}（{dayOfWeekJP(d)}）
                          </option>
                        ))}
                      </select>

                      <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                        <button onClick={applyJapaneseOverride} disabled={!overrideDate}>
                          この日を上書き保存
                        </button>
                        <button
                          onClick={() => {
                            if (!overrideDate) return;
                            deleteOverride(overrideDate);
                            loadOverrideFromProgram(overrideDate);
                          }}
                          disabled={!overrideDate}
                        >
                          上書き削除
                        </button>
                      </div>
                      <div style={{ fontSize: 12, opacity: 0.7, marginTop: 8 }}>※保存後は「自動生成（再生成）」で反映</div>
                    </div>

                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 10 }}>
                      <div>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>この日 日本語講座</label>
                        <select value={ovEnabled ? "Yes" : "No"} onChange={(e) => setOvEnabled(e.target.value === "Yes")} style={{ width: "100%" }}>
                          <option value="Yes">あり</option>
                          <option value="No">なし</option>
                        </select>
                      </div>
                      <div>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>開始時刻</label>
                        <input value={ovStartTime} onChange={(e) => setOvStartTime(normalizeHHMM(e.target.value))} style={{ width: "100%" }} />
                      </div>
                      <div>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>授業分</label>
                        <select value={String(ovLesson)} onChange={(e) => setOvLesson(Number(e.target.value))} style={{ width: "100%" }}>
                          <option value="50">50</option>
                          <option value="60">60</option>
                        </select>
                      </div>
                      <div>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>休憩分</label>
                        <input type="number" value={ovBreak} onChange={(e) => setOvBreak(clamp(Number(e.target.value), 0, 60))} style={{ width: "100%" }} />
                      </div>

                      <div>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>コマ数（1〜3）</label>
                        <input type="number" value={ovPeriods} onChange={(e) => setOvPeriods(clamp(Number(e.target.value), 1, 3))} style={{ width: "100%" }} />
                      </div>
                      <div>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>クラス数</label>
                        <input type="number" value={ovClassCount} onChange={(e) => setOvClassCount(clamp(Number(e.target.value), 1, 50))} style={{ width: "100%" }} />
                      </div>

                      <div style={{ gridColumn: "1 / span 2" }}>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>教室（クラス順に改行）</label>
                        <textarea value={ovClassrooms} onChange={(e) => setOvClassrooms(e.target.value)} rows={4} style={{ width: "100%" }} placeholder={"例）\n洋洋館305\n洋洋館306"} />
                      </div>
                      <div style={{ gridColumn: "3 / span 2" }}>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>講師控室（クラス順に改行）</label>
                        <textarea value={ovTeacherRooms} onChange={(e) => setOvTeacherRooms(e.target.value)} rows={4} style={{ width: "100%" }} placeholder={"例）\n洋洋館301\n洋洋館302"} />
                      </div>
                    </div>
                  </div>

                  <hr style={{ margin: "14px 0" }} />

                  <div style={{ fontWeight: 800, marginBottom: 10 }}>最終日：修了式（自動生成で入る）</div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
                    <div>
                      <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>開始時刻</label>
                      <input
                        value={selectedProgram.ceremony.startTime}
                        onChange={(e) => updateProgram({ ceremony: { ...selectedProgram.ceremony, startTime: normalizeHHMM(e.target.value) } })}
                        style={{ width: "100%" }}
                      />
                    </div>
                    <div>
                      <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>所要分</label>
                      <input
                        type="number"
                        value={selectedProgram.ceremony.durationMinutes}
                        onChange={(e) => updateProgram({ ceremony: { ...selectedProgram.ceremony, durationMinutes: clamp(Number(e.target.value), 10, 240) } })}
                        style={{ width: "100%" }}
                      />
                    </div>
                    <div>
                      <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>会場</label>
                      <input value={selectedProgram.ceremony.location} onChange={(e) => updateProgram({ ceremony: { ...selectedProgram.ceremony, location: e.target.value } })} style={{ width: "100%" }} />
                    </div>
                  </div>
                </div>
              )}

              {activeTab === "calendar" && (
                <div style={{ ...card }}>
                  <div style={{ fontWeight: 800, marginBottom: 10 }}>日程・予定</div>

                  {/* イベント追加 */}
                  <div style={{ ...card, borderRadius: 12, background: "#fafafa", marginBottom: 12 }}>
                    <div style={{ fontWeight: 800, marginBottom: 8 }}>イベント追加（手入力）</div>

                    <div style={{ display: "grid", gridTemplateColumns: "220px 200px 1fr", gap: 10, alignItems: "end" }}>
                      <div>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>日付</label>
                        <select value={selectedDate} onChange={(e) => setSelectedDate(e.target.value)} style={{ width: "100%" }}>
                          <option value="">（選択）</option>
                          {datesInProgram.map((d) => (
                            <option key={d} value={d}>
                              {d}（{dayOfWeekJP(d)}）
                            </option>
                          ))}
                        </select>
                      </div>

                      <div>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>カテゴリ</label>
                        <select value={addCategory} onChange={(e) => setAddCategory(e.target.value as Category)} style={{ width: "100%" }}>
                          <option value="Cultural">文化体験</option>
                          <option value="CompanyVisit">企業訪問</option>
                          <option value="BuddyLunch">バディランチ</option>
                          <option value="Other">その他</option>
                        </select>
                      </div>

                      <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                        <label style={{ display: "flex", gap: 6, alignItems: "center", fontSize: 12 }}>
                          <input type="checkbox" checked={useMaster} onChange={(e) => setUseMaster(e.target.checked)} />
                          マスタから選ぶ
                        </label>

                        {useMaster ? (
                          <select
                            value={masterId}
                            onChange={(e) => setMasterId(e.target.value)}
                            style={{ minWidth: 260, flex: "1 1 260px" }}
                          >
                            {masters
                              .filter((m) => {
                                if (addCategory === "Cultural") return m.category === "Cultural";
                                if (addCategory === "CompanyVisit") return m.category === "CompanyVisit";
                                if (addCategory === "BuddyLunch") return m.category === "BuddyLunch";
                                return true;
                              })
                              .map((m) => (
                                <option key={m.id} value={m.id}>
                                  {m.title}
                                </option>
                              ))}
                          </select>
                        ) : (
                          <input
                            value={freeTitle}
                            onChange={(e) => setFreeTitle(e.target.value)}
                            placeholder="イベント名（自由入力）"
                            style={{ minWidth: 260, flex: "1 1 260px" }}
                          />
                        )}
                      </div>
                    </div>

                    <div style={{ display: "grid", gridTemplateColumns: "160px 160px 1fr 1fr", gap: 10, marginTop: 10 }}>
                      <div>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>開始</label>
                        <input value={eventStartTime} onChange={(e) => setEventStartTime(normalizeHHMM(e.target.value))} style={{ width: "100%" }} placeholder="13:10" />
                      </div>
                      <div>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>終了</label>
                        <input value={eventEndTime} onChange={(e) => setEventEndTime(normalizeHHMM(e.target.value))} style={{ width: "100%" }} placeholder="16:00" />
                      </div>
                      <div>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>場所 / 集合</label>
                        <input value={location} onChange={(e) => setLocation(e.target.value)} style={{ width: "100%" }} placeholder="例）洋洋館305 / ○○前集合" />
                      </div>
                      <div>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>教室手配</label>
                        <select value={roomNeeded} onChange={(e) => setRoomNeeded(e.target.value as YesNo)} style={{ width: "100%" }}>
                          <option value="Yes">要</option>
                          <option value="No">不要</option>
                        </select>
                      </div>
                    </div>

                    <div style={{ display: "grid", gridTemplateColumns: "200px 200px 200px 1fr", gap: 10, marginTop: 10 }}>
                      <div>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>留学生人数</label>
                        <input type="number" value={studentsCount} onChange={(e) => setStudentsCount(clamp(Number(e.target.value), 0, 9999))} style={{ width: "100%" }} />
                      </div>
                      <div>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>バディ人数</label>
                        <input type="number" value={buddyCount} onChange={(e) => setBuddyCount(clamp(Number(e.target.value), 0, 9999))} style={{ width: "100%" }} />
                      </div>
                      <div>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>KVH</label>
                        <select value={kvhRequired} onChange={(e) => setKVHRequired(e.target.value as YesNo)} style={{ width: "100%" }}>
                          <option value="Yes">要</option>
                          <option value="No">不要</option>
                        </select>
                      </div>
                      <div>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>KVH人数</label>
                        <input
                          type="number"
                          value={kvhCount}
                          onChange={(e) => setKVHCount(clamp(Number(e.target.value), 0, 50))}
                          style={{ width: "100%" }}
                          disabled={kvhRequired === "No"}
                        />
                      </div>
                    </div>

                    <div style={{ display: "grid", gridTemplateColumns: "200px 1fr 1fr", gap: 10, marginTop: 10 }}>
                      <div>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>移動</label>
                        <select value={transportMode} onChange={(e) => setTransportMode(e.target.value as TransportMode)} style={{ width: "100%" }}>
                          <option value="Bus">バス</option>
                          <option value="Walk">徒歩</option>
                          <option value="OnCampus">学内</option>
                          <option value="None">なし</option>
                        </select>
                      </div>
                      <div>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>手配要否</label>
                        <select value={arrangementsNeeded} onChange={(e) => setArrangementsNeeded(e.target.value as YesNo)} style={{ width: "100%" }}>
                          <option value="Yes">要</option>
                          <option value="No">不要</option>
                        </select>
                      </div>
                      <div>
                        <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>備考</label>
                        <input value={notes} onChange={(e) => setNotes(e.target.value)} style={{ width: "100%" }} placeholder="例）先方連絡済み / 〇〇支払 など" />
                      </div>
                    </div>

                    {showBusFields && (
                      <div style={{ marginTop: 10, padding: 10, border: "1px solid #ddd", borderRadius: 10, background: "#fff" }}>
                        <div style={{ fontWeight: 700, marginBottom: 8 }}>バス詳細</div>
                        <div style={{ display: "grid", gridTemplateColumns: "1fr 160px 160px", gap: 10 }}>
                          <div>
                            <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>バス会社</label>
                            <input value={busCompany} onChange={(e) => setBusCompany(e.target.value)} style={{ width: "100%" }} />
                          </div>
                          <div>
                            <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>台数</label>
                            <input type="number" value={busCount} onChange={(e) => setBusCount(clamp(Number(e.target.value), 1, 20))} style={{ width: "100%" }} />
                          </div>
                          <div>
                            <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>片道/往復</label>
                            <select value={busTripType} onChange={(e) => setBusTripType(e.target.value as BusTripType)} style={{ width: "100%" }}>
                              <option value="OneWay">片道</option>
                              <option value="RoundTrip">往復</option>
                            </select>
                          </div>
                        </div>

                        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 10 }}>
                          <div>
                            <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>乗車（Pickup）</label>
                            <input value={busPickup} onChange={(e) => setBusPickup(e.target.value)} style={{ width: "100%" }} />
                          </div>
                          <div>
                            <label style={{ display: "block", fontSize: 12, opacity: 0.75 }}>降車（Dropoff）</label>
                            <input value={busDropoff} onChange={(e) => setBusDropoff(e.target.value)} style={{ width: "100%" }} />
                          </div>
                        </div>
                      </div>
                    )}

                    <div style={{ display: "flex", gap: 8, marginTop: 10, alignItems: "center", flexWrap: "wrap" }}>
                      <button onClick={addManualEvent} style={{ fontWeight: 700 }}>
                        ＋ この内容で追加
                      </button>
                      <div style={{ fontSize: 12, opacity: 0.75 }}>
                        ※追加後、下の「予定一覧」や「カレンダープレビュー」に反映されます
                      </div>
                    </div>
                  </div>

                  {/* カレンダープレビュー（スクショ同等表示） */}
                  <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", marginBottom: 10 }}>
                    <div style={{ fontWeight: 800 }}>カレンダープレビュー（スクショ同等）</div>
                    <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, opacity: 0.8 }}>
                      <input type="checkbox" checked={calendarPreviewEnabled} onChange={(e) => setCalendarPreviewEnabled(e.target.checked)} />
                      プレビュー表示
                    </label>
                    <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
                      <button onClick={exportCalendarHTML}>HTMLをダウンロード</button>
                    </div>
                  </div>

                  {calendarPreviewEnabled && (
                    <div style={{ border: "1px solid #ddd", borderRadius: 10, overflow: "hidden", height: 520 }}>
                      <iframe
                        key={state.lastUpdated}
                        title="calendar-preview"
                        srcDoc={calendarHTML}
                        style={{ width: "100%", height: "100%", border: "none" }}
                        sandbox="allow-same-origin"
                      />
                    </div>
                  )}

                  {/* 予定一覧（編集） */}
                  <div style={{ marginTop: 14 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, marginBottom: 8 }}>
  <div style={{ fontWeight: 800 }}>予定一覧（編集）</div>
  <button onClick={removeDuplicatesNow}>重複を削除（今のプログラム）</button>
</div>

                    <div style={{ fontSize: 12, opacity: 0.7, marginBottom: 10, lineHeight: 1.5 }}>
                      ・日本語講座/初日テンプレ/修了式は「自動生成（再生成）」で上書き（Autoだけ置換）<br />
                      ・手入力イベントは消えません（削除しない限り残ります）
                    </div>

                    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                      {datesInProgram.map((d) => {
                        const list = itemsByDate.get(d) ?? [];
                        return (
                          <details key={d} open={list.length > 0}>
                            <summary style={{ cursor: "pointer", fontWeight: 700 }}>
                              {d}（{dayOfWeekJP(d)}）{" "}
                              <span style={{ fontWeight: 400, opacity: 0.7 }}>— {list.length} 件</span>
                            </summary>

                            {list.length === 0 ? (
                              <div style={{ fontSize: 12, opacity: 0.7, padding: "6px 0" }}>予定なし</div>
                            ) : (
                              <div style={{ display: "flex", flexDirection: "column", gap: 8, padding: "10px 0" }}>
                                {list.map((it) => (
                                  <div
                                    key={it.id}
                                    style={{
                                      border: "1px solid #ddd",
                                      borderRadius: 10,
                                      padding: 10,
                                      background: it.generated ? "#fafafa" : "#fff",
                                    }}
                                  >
                                    <div style={{ display: "grid", gridTemplateColumns: "110px 110px 160px 1fr 180px", gap: 8, alignItems: "center" }}>
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
                                      >
                                        <option value="JapaneseClass">日本語講座</option>
                                        <option value="Escort">引率</option>
                                        <option value="Orientation">オリエン</option>
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
                                      <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                                        <button
                                          onClick={() => {
                                            const ok = window.confirm("この予定を削除します。よろしいですか？");
                                            if (!ok) return;
                                            deleteItem(it.id);
                                          }}
                                          style={{ fontSize: 12 }}
                                        >
                                          削除
                                        </button>
                                      </div>
                                    </div>

                                    <div style={{ display: "grid", gridTemplateColumns: "1fr 220px 220px 1fr", gap: 8, marginTop: 8, alignItems: "center" }}>
                                      <input
                                        value={it.location}
                                        onChange={(e) => updateItem(it.id, { location: e.target.value })}
                                        style={{ width: "100%" }}
                                        placeholder="場所/集合"
                                      />
                                      <select value={it.roomNeeded} onChange={(e) => updateItem(it.id, { roomNeeded: e.target.value as YesNo })} style={{ width: "100%" }}>
                                        <option value="Yes">教室手配: 要</option>
                                        <option value="No">教室手配: 不要</option>
                                      </select>
                                      <select
                                        value={it.transportMode}
                                        onChange={(e) => updateItem(it.id, { transportMode: e.target.value as TransportMode })}
                                        style={{ width: "100%" }}
                                      >
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
                                    </div>

                                    <div style={{ fontSize: 12, opacity: 0.75, marginTop: 8 }}>
                                      {it.generated ? "Auto（自動生成）" : "手入力"} / 留学生:{it.studentsCount} / バディ:{it.buddyCount} / KVH:{it.kvhRequired === "Yes" ? `${it.kvhCount}` : "なし"}
                                    </div>
                                  </div>
                                ))}
                              </div>
                            )}
                          </details>
                        );
                      })}
                    </div>
                  </div>
                </div>
              )}

              {activeTab === "share" && (
                <div style={{ ...card }}>
                  <div style={{ fontWeight: 800, marginBottom: 10 }}>共有・出力</div>

                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                    <div style={{ ...card }}>
                      <div style={{ fontWeight: 800, marginBottom: 8 }}>出力</div>
                      <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", marginBottom: 10 }}>
                        <div style={{ fontSize: 12, opacity: 0.75, fontWeight: 700 }}>出力言語</div>
                        <select value={exportLang} onChange={(e) => setExportLang(e.target.value as ExportLang)}>
                          <option value="ja">日本語</option>
                          <option value="en">英語</option>
                        </select>
                        <div style={{ fontSize: 12, opacity: 0.7 }}>※画面表示は日本語のまま。ICS/HTMLの文字のみ切替。</div>
                      </div>
                      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                        <button onClick={exportCSV}>CSV出力（UTF-8 / Excel向け）</button>
                        <button onClick={exportICS}>ICS出力（日本時間 / カレンダー取込）</button>
                        <button onClick={exportCalendarHTML}>カレンダーHTML出力（印刷・配布用）</button>
                        <button onClick={exportJSON}>⬇️ JSON書き出し（バックアップ・共有）</button>
                      </div>
                      <div style={{ fontSize: 12, opacity: 0.75, marginTop: 10, lineHeight: 1.5 }}>
                        ・スクショと同じ見た目で配布する場合は <b>カレンダーHTML</b> を推奨（そのまま開けます）<br />
                        ・相手がこのアプリに取り込むなら <b>JSON</b> を渡してください
                      </div>
                    </div>

                    <div style={{ ...card }}>
                      <div style={{ fontWeight: 800, marginBottom: 8 }}>読み込み</div>
                      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                        <button onClick={openImportDialog}>⬆️ JSON読み込み</button>
                        <button onClick={() => { generateAuto(); alert("自動生成（再生成）しました。"); }}>自動生成（再生成）</button>
                      </div>
                      <div style={{ fontSize: 12, opacity: 0.75, marginTop: 10, lineHeight: 1.5 }}>
                        ・⬆️ JSON読み込み後に、必要なら「自動生成（再生成）」で Auto 部分を作り直してください
                      </div>
                    </div>
                  </div>

                  <div style={{ ...card, marginTop: 12 }}>
                    <div style={{ fontWeight: 800, marginBottom: 8 }}>トラブル時</div>
                    <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                      <button onClick={resetAllData} style={{ fontWeight: 700 }}>
                        全データ初期化（localStorage削除）
                      </button>
                      <button onClick={removeDuplicatesNow}>
                        重複を削除（同じ予定を1つに）
                      </button>
                      <div style={{ fontSize: 12, opacity: 0.75 }}>
                        保存先キー: <code>{STORAGE_KEY}</code> / もし「挙動が前に戻った」場合は、古い localStorage が原因のことがあります
                      </div>
                    </div>
                  </div>

                  <div style={{ ...card, marginTop: 12 }}>
                    <div style={{ fontWeight: 800, marginBottom: 8 }}>現在のプログラム（メモ）</div>
                    <div style={{ fontSize: 12, opacity: 0.8, lineHeight: 1.6 }}>
                      名前: <b>{selectedProgram.name}</b> / 種別: <b>{programTypeLabel(selectedProgram.type)}</b>
                      <br />
                      期間: <b>{selectedProgram.startDate} → {selectedProgram.endDate}</b> / 留学生: <b>{selectedProgram.studentsCount}</b>
                      <br />
                      予定件数: <b>{selectedProgramItems.length}</b>
                    </div>
                  </div>
                </div>
              )}
            </>
          )}
        </div>
      </div>

      <div style={{ marginTop: 10, fontSize: 12, opacity: 0.65 }}>
        保存先：localStorage（キー: {STORAGE_KEY}） / 共有：JSON Export/Import / CSV：UTF-8
      </div>
    </div>
  );
}
