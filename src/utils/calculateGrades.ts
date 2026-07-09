import type {
  CalculationResult,
  CellValue,
  ColumnMapping,
  CourseRecord,
  DuplicateGroup,
  IgnoredRow,
  ManualCourse,
  RawSheetRow,
} from '../types/grade';
import { scoreToGpa, toPercentScore } from './gpaRules';

export function calculateGrades(
  rows: RawSheetRow[],
  mapping: ColumnMapping,
  manualCourses: ManualCourse[],
  originalDataRowCount: number,
): CalculationResult {
  const ignoredRows: IgnoredRow[] = [];
  const validCourses: CourseRecord[] = [];
  const hasExcelMapping = Boolean(mapping.courseName && mapping.credit && mapping.score);

  if (hasExcelMapping) {
    const completeMapping = mapping as Required<ColumnMapping>;
    rows.forEach((row, index) => {
      const parsed = parseRow(row, completeMapping, index);
      if (parsed.course) {
        validCourses.push(parsed.course);
      } else if (parsed.ignored) {
        ignoredRows.push(parsed.ignored);
      }
    });
  }

  manualCourses.forEach((course, index) => {
    const gpa = scoreToGpa(course.score);
    validCourses.push({
      id: course.id,
      name: course.name.trim(),
      credit: course.credit,
      score: course.score,
      gpa,
      scoreCredit: course.score * course.credit,
      creditGpa: course.credit * gpa,
      sourceRow: '模拟课程',
      sourceType: 'manual',
      originalIndex: rows.length + index + 1,
    });
  });

  const { courses, duplicateGroups } = dedupeCourses(validCourses);
  const scoreSum = sum(courses.map((course) => course.score));
  const totalCredits = sum(courses.map((course) => course.credit));
  const scoreCreditSum = sum(courses.map((course) => course.scoreCredit));
  const creditGpaSum = sum(courses.map((course) => course.creditGpa));

  return {
    courses,
    ignoredRows,
    duplicateGroups,
    stats: {
      originalRows: originalDataRowCount,
      ignoredRows: ignoredRows.length,
      duplicateCourseCount: duplicateGroups.length,
      dedupedCourseCount: courses.length,
      totalCredits,
      scoreSum,
      scoreCreditSum,
      creditGpaSum,
    },
    metrics: {
      arithmeticAverage: courses.length > 0 ? scoreSum / courses.length : null,
      weightedAverage: totalCredits > 0 ? scoreCreditSum / totalCredits : null,
      averageGpa: totalCredits > 0 ? creditGpaSum / totalCredits : null,
    },
  };
}

function parseRow(
  row: RawSheetRow,
  mapping: Required<ColumnMapping>,
  originalIndex: number,
): { course?: CourseRecord; ignored?: IgnoredRow } {
  const rawName = row.values[mapping.courseName];
  const rawCredit = row.values[mapping.credit];
  const rawScore = row.values[mapping.score];
  const name = typeof rawName === 'string' ? rawName.trim() : String(rawName ?? '').trim();
  const credit = toNumber(rawCredit);
  const score = toPercentScore(rawScore);

  if (!name) {
    return ignoredRow(row, '课程名称缺失');
  }

  if (credit === null || credit < 0) {
    return ignoredRow(row, '学分缺失或不是有效数字');
  }

  if (score === null) {
    return ignoredRow(row, '成绩缺失或无法转换为百分制成绩');
  }

  if (score < 0 || score > 100) {
    return ignoredRow(row, '成绩超出 0-100 范围');
  }

  const gpa = scoreToGpa(score);
  return {
    course: {
      id: `excel-${row.rowNumber}-${originalIndex}`,
      name,
      credit,
      score,
      gpa,
      scoreCredit: score * credit,
      creditGpa: credit * gpa,
      sourceRow: `第 ${row.rowNumber} 行`,
      sourceType: 'excel',
      originalIndex,
      rawValues: row.values,
    },
  };
}

function ignoredRow(row: RawSheetRow, reason: string): { ignored: IgnoredRow } {
  return {
    ignored: {
      id: `ignored-${row.rowNumber}`,
      rowNumber: `第 ${row.rowNumber} 行`,
      reason,
      rawValues: row.values,
    },
  };
}

function dedupeCourses(courses: CourseRecord[]): {
  courses: CourseRecord[];
  duplicateGroups: DuplicateGroup[];
} {
  const groups = new Map<string, CourseRecord[]>();

  courses.forEach((course) => {
    const key = normalizeCourseName(course.name);
    const group = groups.get(key) ?? [];
    group.push(course);
    groups.set(key, group);
  });

  const keptCourses: CourseRecord[] = [];
  const duplicateGroups: DuplicateGroup[] = [];

  groups.forEach((group) => {
    const kept = group.reduce((best, current) => (isBetterCourse(current, best) ? current : best));
    keptCourses.push(kept);

    if (group.length > 1) {
      duplicateGroups.push({
        courseName: kept.name,
        scores: group.map((course) => course.score),
        credits: group.map((course) => course.credit),
        sourceRows: group.map((course) => course.sourceRow),
        keptScore: kept.score,
        keptCredit: kept.credit,
        keptSourceRow: kept.sourceRow,
        replaced: kept.id !== group[0].id,
      });
    }
  });

  return {
    courses: keptCourses.sort((a, b) => a.originalIndex - b.originalIndex),
    duplicateGroups: duplicateGroups.sort((a, b) => a.courseName.localeCompare(b.courseName, 'zh-CN')),
  };
}

function isBetterCourse(current: CourseRecord, best: CourseRecord): boolean {
  if (current.score !== best.score) {
    return current.score > best.score;
  }
  if (current.credit !== best.credit) {
    return current.credit > best.credit;
  }
  return false;
}

function normalizeCourseName(name: string): string {
  return name.trim().toLocaleLowerCase('zh-CN');
}

function toNumber(value: CellValue): number | null {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }

  if (typeof value !== 'string') {
    return null;
  }

  const numericText = value.trim().replace(/学分$/, '');
  if (!numericText) {
    return null;
  }

  const parsed = Number(numericText);
  return Number.isFinite(parsed) ? parsed : null;
}

function sum(values: number[]): number {
  return values.reduce((total, value) => total + value, 0);
}
