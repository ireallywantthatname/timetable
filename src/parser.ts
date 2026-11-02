import * as XLSX from "xlsx";
import type {
  CalendarEvent,
  EventCategory,
  ModuleInfo,
  ParsedTimetable,
  SyncConfig,
} from "./types.ts";

const DAY_COLS = ["D", "E", "F", "G", "H", "I", "J"] as const;

interface MergeSpan {
  startCol: string;
  startRow: number;
  endCol: string;
  endRow: number;
}

export function excelSerialToDate(serial: number): string {
  const utcMs = Date.UTC(1899, 11, 30) + Math.round(serial) * 86_400_000;
  return new Date(utcMs).toISOString().slice(0, 10);
}

export function excelFractionToTime(fraction: number): string {
  const totalMinutes = Math.round(fraction * 24 * 60);
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
}

export function normalizeModuleCode(text: string): string {
  return text.replace(/\bPUSL\s+(\d{4})\b/gi, "PUSL$1");
}

export function matchesTargetDegree(text: string, keywords: string[]): boolean {
  const lower = text.toLowerCase();
  return keywords.some((kw) => {
    const escaped = kw.toLowerCase().replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const regex = new RegExp(`(?:^|[^a-z0-9])${escaped}(?:$|[^a-z0-9])`, "i");
    return regex.test(lower);
  });
}

export function parseInlineTimeOverride(
  segment: string,
  defaultStart: string,
  defaultEnd: string
): { startTime: string; endTime: string } {
  const match = segment.match(
    /\b(\d{1,2})(?::(\d{2}))?\s*-\s*(\d{1,2})(?::(\d{2}))?\s*(am|pm)\b/i
  );
  if (!match) {
    return { startTime: defaultStart, endTime: defaultEnd };
  }

  let startHour = Number.parseInt(match[1]!, 10);
  const startMin = Number.parseInt(match[2] ?? "0", 10);
  let endHour = Number.parseInt(match[3]!, 10);
  const endMin = Number.parseInt(match[4] ?? "0", 10);
  const meridiem = match[5]!.toLowerCase();

  if (meridiem === "pm") {
    if (endHour < 12) endHour += 12;
    if (startHour < 12 && startHour < Number.parseInt(match[3]!, 10)) {
      startHour += 12;
    }
  } else if (meridiem === "am") {
    if (startHour === 12) startHour = 0;
    if (endHour === 12) endHour = 12;
  }

  return {
    startTime: `${String(startHour).padStart(2, "0")}:${String(startMin).padStart(2, "0")}`,
    endTime: `${String(endHour).padStart(2, "0")}:${String(endMin).padStart(2, "0")}`,
  };
}

export function splitCompoundCellSegments(rawText: string): string[] {
  const normalized = normalizeModuleCode(rawText.trim());
  const lines = normalized.split(/\r?\n/);
  const blocks: string[] = [];
  let currentBlock: string[] = [];

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    if (/^emerging\s+technologies/i.test(trimmed) && currentBlock.length > 0) {
      blocks.push(currentBlock.join("\n"));
      currentBlock = [trimmed];
    } else {
      currentBlock.push(trimmed);
    }
  }
  if (currentBlock.length > 0) {
    blocks.push(currentBlock.join("\n"));
  }

  const segments: string[] = [];
  for (const block of blocks) {
    const slashParts = block.split(/\s*\/\s*(?=PUSL\d{4}\b)/i);
    for (const part of slashParts) {
      if (part.trim()) {
        segments.push(part.trim());
      }
    }
  }

  return segments;
}

function buildEventUid(
  category: EventCategory,
  identifier: string,
  startDate: string,
  startTime?: string
): string {
  const slug = identifier
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
  const timePart = startTime ? `-${startTime.replace(":", "")}` : "";
  return `nsbm-b13-y3s1-${category}-${slug}-${startDate}${timePart}@timetable.nsbm`;
}

export function parseTimetableWorkbook(
  buffer: ArrayBuffer,
  config: SyncConfig
): ParsedTimetable {
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error("Workbook does not contain any sheets");
  }

  const sheet = workbook.Sheets[sheetName]!;

  const getCellText = (col: string, row: number): string => {
    const cell = sheet[`${col}${row}`] as XLSX.CellObject | undefined;
    if (!cell || cell.v === undefined || cell.v === null) return "";
    return String(cell.v).trim();
  };

  const getCellNumber = (col: string, row: number): number | undefined => {
    const cell = sheet[`${col}${row}`] as XLSX.CellObject | undefined;
    if (!cell || cell.v === undefined || cell.v === null) return undefined;
    const num = typeof cell.v === "number" ? cell.v : Number.parseFloat(String(cell.v));
    return Number.isFinite(num) ? num : undefined;
  };

  const mergeMap = new Map<string, MergeSpan>();
  for (const merge of sheet["!merges"] ?? []) {
    const startCol = XLSX.utils.encode_col(merge.s.c);
    const startRow = merge.s.r + 1;
    const endCol = XLSX.utils.encode_col(merge.e.c);
    const endRow = merge.e.r + 1;
    mergeMap.set(`${startCol}${startRow}`, { startCol, startRow, endCol, endRow });
  }

  const title = getCellText("B", 1) || sheetName;
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:J300");
  const maxRow = range.e.r + 1;

  const weekRows: number[] = [];
  for (let r = 1; r <= maxRow; r++) {
    const bVal = getCellText("B", r);
    if (/^week\s+\d+$/i.test(bVal)) {
      weekRows.push(r);
    }
  }

  const firstWeekRow = weekRows[0] ?? 15;

  const modules: Record<string, ModuleInfo> = {};
  for (let r = 1; r < firstWeekRow; r++) {
    const rawCode = normalizeModuleCode(getCellText("B", r));
    if (/^PUSL\d{4}$/i.test(rawCode)) {
      const code = rawCode.toUpperCase();
      const name = getCellText("D", r).replace(/\s+/g, " ").trim();
      const lecturer = getCellText("G", r).replace(/\s+/g, " ").trim();
      const degrees = getCellText("I", r).replace(/\s+/g, " ").trim();
      modules[code] = { code, name, lecturer, degrees };
    }
  }

  const allowedSet = new Set(config.allowedModuleCodes.map((c) => c.toUpperCase()));
  const targetModules: Record<string, ModuleInfo> = {};
  for (const [code, info] of Object.entries(modules)) {
    if (
      allowedSet.has(code) ||
      matchesTargetDegree(info.degrees, config.targetDegreeKeywords)
    ) {
      if (allowedSet.size === 0 || allowedSet.has(code)) {
        targetModules[code] = info;
      }
    }
  }

  const events: CalendarEvent[] = [];

  for (let wIdx = 0; wIdx < weekRows.length; wIdx++) {
    const weekRow = weekRows[wIdx]!;
    const nextWeekRow = weekRows[wIdx + 1] ?? maxRow + 1;
    const weekLabel = getCellText("B", weekRow);

    const colDates: Record<string, string> = {};
    for (const col of DAY_COLS) {
      const serial = getCellNumber(col, weekRow);
      if (serial !== undefined) {
        colDates[col] = excelSerialToDate(serial);
      }
    }

    for (let r = weekRow + 1; r < nextWeekRow; r++) {
      for (const col of DAY_COLS) {
        const cellRef = `${col}${r}`;
        const rawCell = getCellText(col, r);
        if (!rawCell) continue;

        const span = mergeMap.get(cellRef) ?? {
          startCol: col,
          startRow: r,
          endCol: col,
          endRow: r,
        };

        const startDate = colDates[span.startCol];
        const endDate = colDates[span.endCol] ?? startDate;
        if (!startDate || !endDate) continue;

        const startFrac = getCellNumber("B", span.startRow);
        const endFrac = getCellNumber("C", span.endRow);
        const defaultStartTime =
          startFrac !== undefined ? excelFractionToTime(startFrac) : "09:00";
        const defaultEndTime =
          endFrac !== undefined ? excelFractionToTime(endFrac) : "17:00";

        const normalizedCell = normalizeModuleCode(rawCell);

        if (/^holiday$/i.test(normalizedCell)) {
          if (config.includeHolidays) {
            events.push({
              uid: buildEventUid("holiday", "holiday", startDate),
              title: "Holiday",
              description: `Academic Holiday (${weekLabel})\nSource Cell: ${cellRef}`,
              category: "holiday",
              weekLabel,
              startDate,
              endDate,
              allDay: true,
              rawCell,
              cellRef,
            });
          }
          continue;
        }

        if (/^study\s+leave$/i.test(normalizedCell)) {
          if (config.includeStudyLeave) {
            events.push({
              uid: buildEventUid("study_leave", "study-leave", startDate),
              title: "Study Leave",
              description: `Study Leave Period (${weekLabel})\nSource Cell: ${cellRef}`,
              category: "study_leave",
              weekLabel,
              startDate,
              endDate,
              allDay: true,
              rawCell,
              cellRef,
            });
          }
          continue;
        }

        if (/coursework\s*submission/i.test(normalizedCell)) {
          const matchedCodes = Array.from(
            new Set(
              (normalizedCell.match(/\bPUSL\d{4}\b/gi) ?? []).map((c) => c.toUpperCase())
            )
          );
          for (const code of matchedCodes) {
            const mod = targetModules[code];
            if (!mod) continue;
            events.push({
              uid: buildEventUid("submission", code, startDate, defaultStartTime),
              title: `[Submission] ${mod.code}: ${mod.name}`,
              description: [
                `Type: Coursework Submission`,
                `Module: ${mod.code} — ${mod.name}`,
                `Lecturer: ${mod.lecturer}`,
                `Week: ${weekLabel}`,
                `Raw Timetable Entry: ${rawCell.replace(/\s+/g, " ")}`,
              ].join("\n"),
              category: "submission",
              moduleCode: mod.code,
              moduleName: mod.name,
              lecturer: mod.lecturer,
              weekLabel,
              startDate,
              endDate,
              startTime: defaultStartTime,
              endTime: defaultEndTime,
              allDay: false,
              rawCell,
              cellRef,
            });
          }
          continue;
        }

        if (/\bexam\b/i.test(normalizedCell)) {
          const matchedCodes = Array.from(
            new Set(
              (normalizedCell.match(/\bPUSL\d{4}\b/gi) ?? []).map((c) => c.toUpperCase())
            )
          );
          for (const code of matchedCodes) {
            const mod = targetModules[code];
            if (!mod) continue;
            events.push({
              uid: buildEventUid("exam", code, startDate, defaultStartTime),
              title: `[Exam] ${mod.code}: ${mod.name}`,
              description: [
                `Type: Examination`,
                `Module: ${mod.code} — ${mod.name}`,
                `Lecturer: ${mod.lecturer}`,
                `Week: ${weekLabel}`,
                `Raw Timetable Entry: ${rawCell.replace(/\s+/g, " ")}`,
              ].join("\n"),
              category: "exam",
              moduleCode: mod.code,
              moduleName: mod.name,
              lecturer: mod.lecturer,
              weekLabel,
              startDate,
              endDate,
              startTime: defaultStartTime,
              endTime: defaultEndTime,
              allDay: false,
              rawCell,
              cellRef,
            });
          }
          continue;
        }

        if (
          /industry\s+preparation\s+training/i.test(normalizedCell) ||
          /plymouth\s+induction/i.test(normalizedCell)
        ) {
          if (config.includeSpecialSessions) {
            const cleanTitle = normalizedCell.replace(/\s+/g, " ").trim();
            const startColIdx = DAY_COLS.indexOf(span.startCol as (typeof DAY_COLS)[number]);
            const endColIdx = DAY_COLS.indexOf(span.endCol as (typeof DAY_COLS)[number]);
            const colIndices =
              startColIdx >= 0 && endColIdx >= startColIdx
                ? Array.from({ length: endColIdx - startColIdx + 1 }, (_, i) => startColIdx + i)
                : [0];

            for (const cIdx of colIndices) {
              const dayCol = DAY_COLS[cIdx] ?? span.startCol;
              const dayDate = colDates[dayCol] ?? startDate;
              events.push({
                uid: buildEventUid("special", cleanTitle, dayDate, defaultStartTime),
                title: cleanTitle,
                description: [
                  `Type: Special Programme / Induction`,
                  `Week: ${weekLabel}`,
                  `Source Cell: ${cellRef}`,
                ].join("\n"),
                category: "special",
                weekLabel,
                startDate: dayDate,
                endDate: dayDate,
                startTime: defaultStartTime,
                endTime: defaultEndTime,
                allDay: false,
                rawCell,
                cellRef,
              });
            }
          }
          continue;
        }

        const segments = splitCompoundCellSegments(normalizedCell);
        for (const segment of segments) {
          if (/^emerging\s+technologies/i.test(segment)) {
            if (
              config.includeSpecialSessions &&
              matchesTargetDegree(segment, config.targetDegreeKeywords)
            ) {
              const { startTime, endTime } = parseInlineTimeOverride(
                segment,
                defaultStartTime,
                defaultEndTime
              );
              events.push({
                uid: buildEventUid(
                  "special",
                  "emerging-technologies",
                  startDate,
                  startTime
                ),
                title: "Emerging Technologies Session",
                description: [
                  `Type: Special Session`,
                  `Details: ${segment.replace(/\s+/g, " ")}`,
                  `Week: ${weekLabel}`,
                  `Source Cell: ${cellRef}`,
                ].join("\n"),
                category: "special",
                weekLabel,
                startDate,
                endDate,
                startTime,
                endTime,
                allDay: false,
                rawCell,
                cellRef,
              });
            }
            continue;
          }

          const codeMatch = segment.match(/\b(PUSL\d{4})\b/i);
          if (!codeMatch) continue;

          const code = codeMatch[1]!.toUpperCase();
          const mod = targetModules[code];
          if (!mod) continue;

          const remainderAfterCode = segment
            .slice(codeMatch.index! + code.length)
            .trim();
          if (remainderAfterCode.length > 0) {
            const hasDegreeRestriction =
              /\b(SE|CS|Comp\s*Science|CN|DS|TM|AI|Com\.?\s*Sec(?:urity|u)?|C\s*Sec)\b/i.test(
                remainderAfterCode
              );
            if (
              hasDegreeRestriction &&
              !matchesTargetDegree(remainderAfterCode, config.targetDegreeKeywords)
            ) {
              continue;
            }
          }

          const isWorkshop = /\bworkshop\b/i.test(segment);
          const category: EventCategory = isWorkshop ? "workshop" : "lecture";
          const { startTime, endTime } = parseInlineTimeOverride(
            segment,
            defaultStartTime,
            defaultEndTime
          );

          const prefix = isWorkshop ? "[Workshop] " : "";
          const title = `${prefix}${mod.code}: ${mod.name} (${mod.lecturer})`;

          events.push({
            uid: buildEventUid(category, mod.code, startDate, startTime),
            title,
            description: [
              `Type: ${isWorkshop ? "Workshop" : "Lecture"}`,
              `Module: ${mod.code} — ${mod.name}`,
              `Lecturer: ${mod.lecturer}`,
              `Week: ${weekLabel}`,
              `Raw Timetable Entry: ${segment.replace(/\s+/g, " ")}`,
            ].join("\n"),
            category,
            moduleCode: mod.code,
            moduleName: mod.name,
            lecturer: mod.lecturer,
            weekLabel,
            startDate,
            endDate,
            startTime,
            endTime,
            allDay: false,
            rawCell,
            cellRef,
          });
        }
      }
    }
  }

  events.sort((a, b) => {
    if (a.startDate !== b.startDate) return a.startDate.localeCompare(b.startDate);
    const tA = a.startTime ?? "00:00";
    const tB = b.startTime ?? "00:00";
    if (tA !== tB) return tA.localeCompare(tB);
    return a.title.localeCompare(b.title);
  });

  return {
    sheetName,
    title,
    modules,
    targetModules,
    events,
  };
}
