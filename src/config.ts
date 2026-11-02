import type { SyncConfig } from "./types.ts";

export const DEFAULT_SHAREPOINT_URL =
  "https://nsbm365.sharepoint.com/:x:/g/SOC/IQBCDncIiBTZSq63v-2wpBDWAaMcsQMSiTTGK-mRgTwuLGc?rtime=cvNnYsA430g";

export const DEFAULT_CONFIG: SyncConfig = {
  sharepointUrl: process.env.TIMETABLE_URL?.trim() || DEFAULT_SHAREPOINT_URL,
  targetDegreeKeywords: ["com.security", "com.sec", "com.secu", "c sec", "computer security", "sec"],
  allowedModuleCodes: ["PUSL3190", "PUSL3131", "PUSL3132", "PUSL3133"],
  includeSpecialSessions: true,
  includeHolidays: true,
  includeStudyLeave: true,
  timezone: "Asia/Colombo",
  outputDir: "public",
};
