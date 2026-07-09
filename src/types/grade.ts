export type CellValue = string | number | boolean | Date | null | undefined;

export interface ColumnMapping {
  courseName?: string;
  credit?: string;
  score?: string;
}

export interface RawSheetRow {
  rowNumber: number;
  values: Record<string, CellValue>;
  cells: CellValue[];
}

export interface ParsedSheet {
  fileName: string;
  sheetName: string;
  headers: string[];
  rows: RawSheetRow[];
  originalDataRowCount: number;
  inferredMapping: ColumnMapping;
  needsManualMapping: boolean;
}

export interface ManualCourse {
  id: string;
  name: string;
  credit: number;
  score: number;
}

export interface CourseRecord {
  id: string;
  name: string;
  credit: number;
  score: number;
  gpa: number;
  scoreCredit: number;
  creditGpa: number;
  sourceRow: string;
  sourceType: 'excel' | 'manual';
  originalIndex: number;
  rawValues?: Record<string, CellValue>;
}

export interface IgnoredRow {
  id: string;
  rowNumber: string;
  reason: string;
  rawValues: Record<string, CellValue>;
}

export interface DuplicateGroup {
  courseName: string;
  scores: number[];
  credits: number[];
  sourceRows: string[];
  keptScore: number;
  keptCredit: number;
  keptSourceRow: string;
  replaced: boolean;
}

export interface GradeStats {
  originalRows: number;
  ignoredRows: number;
  duplicateCourseCount: number;
  dedupedCourseCount: number;
  totalCredits: number;
  scoreSum: number;
  scoreCreditSum: number;
  creditGpaSum: number;
}

export interface GradeMetrics {
  arithmeticAverage: number | null;
  weightedAverage: number | null;
  averageGpa: number | null;
}

export interface CalculationResult {
  courses: CourseRecord[];
  ignoredRows: IgnoredRow[];
  duplicateGroups: DuplicateGroup[];
  stats: GradeStats;
  metrics: GradeMetrics;
}

export type SortKey = 'name' | 'credit' | 'score';
export type SortDirection = 'asc' | 'desc';
