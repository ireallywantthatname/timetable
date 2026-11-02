import type { CalendarEvent } from "./types.ts";

export function escapeIcsText(value: string): string {
  return value
    .replace(/\\/g, "\\\\")
    .replace(/;/g, "\\;")
    .replace(/,/g, "\\,")
    .replace(/\r?\n/g, "\\n");
}

export function foldIcsLine(line: string): string {
  const encoder = new TextEncoder();
  if (encoder.encode(line).length <= 75) {
    return line;
  }

  const parts: string[] = [];
  let current = "";
  let currentBytes = 0;
  let maxBytes = 75;

  for (const char of line) {
    const charBytes = encoder.encode(char).length;
    if (currentBytes + charBytes > maxBytes) {
      parts.push(current);
      current = ` ${char}`;
      currentBytes = 1 + charBytes;
      maxBytes = 75;
    } else {
      current += char;
      currentBytes += charBytes;
    }
  }
  if (current) {
    parts.push(current);
  }

  return parts.join("\r\n");
}

export function addDaysToIsoDate(isoDate: string, days: number): string {
  const [y, m, d] = isoDate.split("-").map((n) => Number.parseInt(n, 10));
  const utc = Date.UTC(y!, m! - 1, d!) + days * 86_400_000;
  return new Date(utc).toISOString().slice(0, 10);
}

function formatIcsDate(isoDate: string): string {
  return isoDate.replace(/-/g, "");
}

function formatIcsDateTime(isoDate: string, time: string): string {
  return `${formatIcsDate(isoDate)}T${time.replace(":", "")}00`;
}

export function generateIcsContent(
  events: CalendarEvent[],
  calendarName: string,
  calendarDescription: string,
  timezone = "Asia/Colombo",
  dtStamp = "20260727T000000Z"
): string {
  const lines: string[] = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//NSBM Computer Security Timetable//TypeScript Bun Sync//EN",
    "CALSCALE:GREGORIAN",
    "METHOD:PUBLISH",
    `X-WR-CALNAME:${escapeIcsText(calendarName)}`,
    `X-WR-CALDESC:${escapeIcsText(calendarDescription)}`,
    `X-WR-TIMEZONE:${timezone}`,
    "REFRESH-INTERVAL;VALUE=DURATION:PT6H",
    "X-PUBLISHED-TTL:PT6H",
    "BEGIN:VTIMEZONE",
    `TZID:${timezone}`,
    "BEGIN:STANDARD",
    "DTSTART:19700101T000000",
    "TZOFFSETFROM:+0530",
    "TZOFFSETTO:+0530",
    "TZNAME:+0530",
    "END:STANDARD",
    "END:VTIMEZONE",
  ];

  for (const ev of events) {
    lines.push("BEGIN:VEVENT");
    lines.push(`UID:${ev.uid}`);
    lines.push(`DTSTAMP:${dtStamp}`);

    if (ev.allDay) {
      const dtStart = formatIcsDate(ev.startDate);
      const dtEndExclusive = formatIcsDate(addDaysToIsoDate(ev.endDate, 1));
      lines.push(`DTSTART;VALUE=DATE:${dtStart}`);
      lines.push(`DTEND;VALUE=DATE:${dtEndExclusive}`);
    } else {
      const startTime = ev.startTime ?? "09:00";
      const endTime = ev.endTime ?? "17:00";
      lines.push(
        `DTSTART;TZID=${timezone}:${formatIcsDateTime(ev.startDate, startTime)}`
      );
      lines.push(
        `DTEND;TZID=${timezone}:${formatIcsDateTime(ev.endDate, endTime)}`
      );
    }

    lines.push(`SUMMARY:${escapeIcsText(ev.title)}`);
    lines.push(`DESCRIPTION:${escapeIcsText(ev.description)}`);
    lines.push(`CATEGORIES:${ev.category.toUpperCase()}`);
    lines.push(
      `TRANSP:${ev.category === "holiday" || ev.category === "study_leave" ? "TRANSPARENT" : "OPAQUE"}`
    );
    lines.push("STATUS:CONFIRMED");
    lines.push("END:VEVENT");
  }

  lines.push("END:VCALENDAR");

  return `${lines.map(foldIcsLine).join("\r\n")}\r\n`;
}
