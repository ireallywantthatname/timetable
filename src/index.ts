import { mkdir } from "node:fs/promises";
import { join } from "node:path";
import { DEFAULT_CONFIG } from "./config.ts";
import { downloadTimetableXlsx } from "./fetcher.ts";
import { generateIcsContent } from "./ics.ts";
import { parseTimetableWorkbook } from "./parser.ts";

async function main() {
  const args = process.argv.slice(2);
  let inputPath: string | undefined;

  for (let i = 0; i < args.length; i++) {
    if (args[i] === "--input" && args[i + 1]) {
      inputPath = args[i + 1];
      i++;
    }
  }

  let buffer: ArrayBuffer;
  if (inputPath) {
    buffer = await Bun.file(inputPath).arrayBuffer();
  } else {
    buffer = await downloadTimetableXlsx(DEFAULT_CONFIG.sharepointUrl);
  }

  const parsed = parseTimetableWorkbook(buffer, DEFAULT_CONFIG);
  const outDir = DEFAULT_CONFIG.outputDir;
  await mkdir(outDir, { recursive: true });

  const nowIso = new Date().toISOString();
  const dtStamp = nowIso.replace(/[-:]/g, "").replace(/\.\d+Z$/, "Z");

  const icsContent = generateIcsContent(
    parsed.events,
    "NSBM Com.Sec Timetable",
    "NSBM BSc (Hons) Computer Security Timetable",
    DEFAULT_CONFIG.timezone,
    dtStamp
  );

  const filePath = join(outDir, "timetable.ics");
  await Bun.write(filePath, icsContent);
  console.log(`Wrote ${filePath} (${parsed.events.length} events)`);
}

if (import.meta.main) {
  main().catch((err) => {
    console.error("Error syncing timetable:", err);
    process.exit(1);
  });
}
