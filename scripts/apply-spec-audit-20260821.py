from pathlib import Path

app = Path('components/ScheduleApp.tsx')
text = app.read_text(encoding='utf-8')

anchor = 'const STORAGE_KEY = "rsjp_schedule_mvp_state_v2";\n'
validation = '''const STORAGE_KEY = "rsjp_schedule_mvp_state_v2";\n\nfunction isValidImportedState(value: unknown): value is AppState {\n  if (!value || typeof value !== "object") return false;\n  const candidate = value as Partial<AppState>;\n  if (!Array.isArray(candidate.programs) || !Array.isArray(candidate.items)) return false;\n\n  const programIds = new Set<string>();\n  for (const program of candidate.programs) {\n    if (!program || typeof program.id !== "string" || !program.id.trim() || typeof program.name !== "string") return false;\n    if (programIds.has(program.id)) return false;\n    programIds.add(program.id);\n  }\n\n  for (const item of candidate.items) {\n    if (!item || typeof item.id !== "string" || !item.id.trim() || typeof item.programId !== "string" || !programIds.has(item.programId)) return false;\n    if (typeof item.date !== "string" || typeof item.startTime !== "string" || typeof item.endTime !== "string") return false;\n  }\n\n  return true;\n}\n'''
if anchor not in text:
    raise SystemExit('storage anchor not found')
text = text.replace(anchor, validation, 1)

old = '''        const parsed = JSON.parse(txt) as AppState;\n        if (!parsed || !Array.isArray(parsed.programs) || !Array.isArray(parsed.items)) {\n          alert("JSONの形式が正しくありません。");\n          return;\n        }'''
new = '''        const parsed = JSON.parse(txt) as unknown;\n        if (!isValidImportedState(parsed)) {\n          alert("JSONの形式または参照関係が正しくありません。Program IDの重複や孤立した予定がないか確認してください。");\n          return;\n        }'''
if old not in text:
    raise SystemExit('import validation target not found')
text = text.replace(old, new, 1)
app.write_text(text, encoding='utf-8')

layout = Path('app/layout.tsx')
text = layout.read_text(encoding='utf-8')
old = '  description: "RSJP and custom programme schedule planning tool",\n};'
new = '  description: "RSJP and custom programme schedule planning tool",\n  robots: { index: false, follow: false },\n};'
if old not in text:
    raise SystemExit('layout metadata target not found')
layout.write_text(text.replace(old, new, 1), encoding='utf-8')
