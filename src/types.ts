export type EventCategory =
  | "lecture"
  | "workshop"
  | "submission"
  | "exam"
  | "special"
  | "holiday"
  | "study_leave";

export interface ModuleInfo {
  code: string;
  name: string;
  lecturer: string;
  degrees: string;
}

export interface CalendarEvent {
  uid: string;
  title: string;
  description: string;
  category: EventCategory;
  moduleCode?: string;
  moduleName?: string;
  lecturer?: string;
  weekLabel: string;
  startDate: string;
  endDate: string;
  startTime?: string;
  endTime?: string;
  allDay: boolean;
  rawCell: string;
  cellRef: string;
}

export interface SyncConfig {
  sharepointUrl: string;
  targetDegreeKeywords: string[];
  allowedModuleCodes: string[];
  includeSpecialSessions: boolean;
  includeHolidays: boolean;
  includeStudyLeave: boolean;
  timezone: string;
  outputDir: string;
}

export interface ParsedTimetable {
  sheetName: string;
  title: string;
  modules: Record<string, ModuleInfo>;
  targetModules: Record<string, ModuleInfo>;
  events: CalendarEvent[];
}
