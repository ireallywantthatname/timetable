import { describe, expect, test } from "bun:test";
import * as XLSX from "xlsx";
import { DEFAULT_CONFIG } from "../src/config.ts";
import { buildSharePointDownloadUrl } from "../src/fetcher.ts";
import { addDaysToIsoDate, foldIcsLine, generateIcsContent } from "../src/ics.ts";
import {
  excelFractionToTime,
  excelSerialToDate,
  matchesTargetDegree,
  normalizeModuleCode,
  parseInlineTimeOverride,
  parseTimetableWorkbook,
  splitCompoundCellSegments,
} from "../src/parser.ts";

function buildMockNsbmWorkbookBuffer(): ArrayBuffer {
  const ws: XLSX.WorkSheet = {};

  const setCell = (ref: string, val: string | number) => {
    ws[ref] =
      typeof val === "number" ? { t: "n", v: val } : { t: "s", v: val };
  };

  setCell("B1", "Stage 3 (24.1/24.2/24.3/25.1) Semester 1 Timetable");

  const moduleRows: Array<[number, string, string, string, string]> = [
    [5, "PUSL3190", "Computing Project", "Mr. Diluka Wijesinghe", "Comp Science/SE/\nCom.Security/CN/DS/TM/AI"],
    [6, "PUSL3120", "Full Stack Development", "Dr. Mark Dixon", "Comp Science/SE/TM/AI"],
    [7, "PUSL3123", "AI and Machine Learning ", "Dr. Vivek Singh", "Comp Science/SE/DS/AI"],
    [8, "PUSL3131", "Security Operation and Incident Management", "Dr. Hai Van Dang", "CN / Com.Security"],
    [9, "PUSL3132", "Ethical Hacking", "Prof. Nathan Clarke", "CN / Com.Security"],
    [10, "PUSL3133", "Digital Forensics and Malware Analysis", "Mr. Isuru Sri Bandara", "Com.Security"],
    [11, "PUSL3159", "Entreprise, Innovations and Creativity", "Ms. Aushadari Kaushalya", "TM"],
    [12, "PUSL3205", "Large Language Models and Generative AI", "Ms.Thilini Bakmeedeniya", "DS/AI"],
  ];

  for (const [r, code, name, lecturer, deg] of moduleRows) {
    setCell(`B${r}`, code);
    setCell(`D${r}`, name);
    setCell(`G${r}`, lecturer);
    setCell(`I${r}`, deg);
  }

  setCell("B15", "Week 1");
  const cols = ["D", "E", "F", "G", "H", "I", "J"];
  cols.forEach((c, idx) => setCell(`${c}15`, 46230 + idx));

  for (let i = 0; i < 8; i++) {
    const r = 16 + i;
    setCell(`B${r}`, (9 + i) / 24);
    setCell(`C${r}`, (10 + i) / 24);
  }

  setCell("I16", "PUSL3123 - Comp Science/SE/Com.Security/CN/DS/TM/AI");
  setCell("H16", "PUSL3190 -SE");
  setCell("H18", "PUSL3190  CS/Com.Sec/DS/CN/TM/AI");
  setCell("J16", "PUSL3159 -  (TM) /PUSL3133 - Com.Secu");
  setCell(
    "D16",
    "PUSL3132(CN/C Sec) \n( Prof. Nathan Clarke) 9-12 am    \nEmerging Technologies(Generative and Agentic AI )- SE/CS/DS/TM/AI (11-1 pm)"
  );
  setCell("E16", "PUSL 3133 / PUSL3205\nCOURSEWORK SUBMISSION");
  setCell("F16", "EXAM \nPUSL 3131");
  setCell("G16", "HOLIDAY");

  ws["!ref"] = "A1:J25";
  ws["!merges"] = [
    { s: { c: 8, r: 15 }, e: { c: 8, r: 17 } },
    { s: { c: 7, r: 15 }, e: { c: 7, r: 16 } },
    { s: { c: 7, r: 17 }, e: { c: 7, r: 18 } },
    { s: { c: 9, r: 15 }, e: { c: 9, r: 17 } },
    { s: { c: 3, r: 15 }, e: { c: 3, r: 19 } },
    { s: { c: 4, r: 15 }, e: { c: 4, r: 22 } },
    { s: { c: 5, r: 15 }, e: { c: 5, r: 22 } },
    { s: { c: 6, r: 15 }, e: { c: 6, r: 22 } },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Semester 01 PU Batch 13 Stage 3");
  return XLSX.write(wb, { type: "array", bookType: "xlsx" }) as ArrayBuffer;
}

describe("Timetable Parser & Helpers", () => {
  test("converts Excel serial dates and time fractions accurately", () => {
    expect(excelSerialToDate(46230)).toBe("2026-07-27");
    expect(excelSerialToDate(46398)).toBe("2027-01-11");
    expect(excelFractionToTime(0.375)).toBe("09:00");
    expect(excelFractionToTime(0.5)).toBe("12:00");
    expect(excelFractionToTime(0.5416666666666666)).toBe("13:00");
    expect(excelFractionToTime(0.7083333333333334)).toBe("17:00");
  });

  test("normalizes spaced module codes", () => {
    expect(normalizeModuleCode("EXAM \nPUSL 3133")).toBe("EXAM \nPUSL3133");
    expect(normalizeModuleCode("PUSL 3120 / PUSL 3131")).toBe("PUSL3120 / PUSL3131");
  });

  test("parses inline time overrides like 9-12 am and 1-4 pm", () => {
    expect(
      parseInlineTimeOverride(
        "PUSL3132(CN/C Sec) ( Prof. Nathan Clarke) 9-12 am",
        "09:00",
        "13:00"
      )
    ).toEqual({ startTime: "09:00", endTime: "12:00" });

    expect(
      parseInlineTimeOverride(
        "PUSL3132(CN/C Sec) ( Prof. Nathan Clarke 1-4 pm )",
        "13:00",
        "17:00"
      )
    ).toEqual({ startTime: "13:00", endTime: "16:00" });
  });

  test("splits compound cells cleanly", () => {
    const segments1 = splitCompoundCellSegments(
      "PUSL3120 (CS/TM/AI)  \nGroup B\nDr. Mark Dixon / PUSL3131 (CN/Com.Security)-Dr. Hai Van"
    );
    expect(segments1).toHaveLength(2);
    expect(segments1[1]).toContain("PUSL3131");

    const segments2 = splitCompoundCellSegments(
      "PUSL3132(CN/C Sec) \n( Prof. Nathan Clarke) 9-12 am    \nEmerging Technologies(Generative and Agentic AI )- SE/CS/DS/TM/AI (11-1 pm)"
    );
    expect(segments2).toHaveLength(2);
    expect(segments2[0]).toContain("PUSL3132");
    expect(segments2[1]).toStartWith("Emerging Technologies");
  });

  test("matches target degree tokens without false positives", () => {
    expect(
      matchesTargetDegree("CN / Com.Security", DEFAULT_CONFIG.targetDegreeKeywords)
    ).toBe(true);
    expect(
      matchesTargetDegree("SE/CS/DS/TM/AI", DEFAULT_CONFIG.targetDegreeKeywords)
    ).toBe(false);
    expect(
      matchesTargetDegree("-SE", DEFAULT_CONFIG.targetDegreeKeywords)
    ).toBe(false);
  });

  test("parses mock NSBM workbook and filters strictly for Computer Security", () => {
    const buffer = buildMockNsbmWorkbookBuffer();
    const parsed = parseTimetableWorkbook(buffer, DEFAULT_CONFIG);

    expect(Object.keys(parsed.targetModules).sort()).toEqual([
      "PUSL3131",
      "PUSL3132",
      "PUSL3133",
      "PUSL3190",
    ]);

    expect(parsed.events.some((e) => e.moduleCode === "PUSL3123")).toBe(false);
    expect(parsed.events.some((e) => e.moduleCode === "PUSL3159")).toBe(false);
    expect(parsed.events.some((e) => e.moduleCode === "PUSL3205")).toBe(false);
    expect(parsed.events.some((e) => e.rawCell === "PUSL3190 -SE")).toBe(false);

    const pusl3132Event = parsed.events.find((e) => e.moduleCode === "PUSL3132");
    expect(pusl3132Event).toBeDefined();
    expect(pusl3132Event?.startTime).toBe("09:00");
    expect(pusl3132Event?.endTime).toBe("12:00");
    expect(pusl3132Event?.title).toBe(
      "PUSL3132: Ethical Hacking (Prof. Nathan Clarke)"
    );

    const submissions = parsed.events.filter((e) => e.category === "submission");
    expect(submissions).toHaveLength(1);
    expect(submissions[0]?.moduleCode).toBe("PUSL3133");
    expect(submissions[0]?.title).toBe(
      "[Submission] PUSL3133: Digital Forensics and Malware Analysis"
    );

    const exams = parsed.events.filter((e) => e.category === "exam");
    expect(exams).toHaveLength(1);
    expect(exams[0]?.moduleCode).toBe("PUSL3131");

    const holidays = parsed.events.filter((e) => e.category === "holiday");
    expect(holidays).toHaveLength(1);
    expect(holidays[0]?.allDay).toBe(true);
  });

  test("generates valid RFC 5545 .ics output with folded lines and exclusive all-day DTEND", () => {
    const buffer = buildMockNsbmWorkbookBuffer();
    const parsed = parseTimetableWorkbook(buffer, DEFAULT_CONFIG);
    const ics = generateIcsContent(
      parsed.events,
      "NSBM Com.Sec Timetable",
      "Test description",
      "Asia/Colombo",
      "20260727T000000Z"
    );

    expect(ics).toContain("BEGIN:VCALENDAR\r\n");
    expect(ics).toContain("TZID:Asia/Colombo\r\n");
    expect(ics).toContain("DTSTART;TZID=Asia/Colombo:20260727T090000\r\n");
    expect(ics).toContain("DTEND;TZID=Asia/Colombo:20260727T120000\r\n");
    expect(ics).toContain("DTSTART;VALUE=DATE:20260730\r\n");
    expect(ics).toContain("DTEND;VALUE=DATE:20260731\r\n");

    for (const line of ics.split("\r\n")) {
      expect(new TextEncoder().encode(line).length).toBeLessThanOrEqual(75);
    }
  });

  test("builds SharePoint download URL and date math helpers", () => {
    expect(
      buildSharePointDownloadUrl(
        "https://nsbm365.sharepoint.com/:x:/g/SOC/IQB?rtime=cvNnYsA430g"
      )
    ).toBe("https://nsbm365.sharepoint.com/:x:/g/SOC/IQB?download=1");
    expect(addDaysToIsoDate("2026-12-27", 1)).toBe("2026-12-28");
    expect(foldIcsLine("A".repeat(100))).toContain("\r\n ");
  });
});
