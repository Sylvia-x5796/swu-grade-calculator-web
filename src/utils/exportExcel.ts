import * as XLSX from 'xlsx';
import type { CalculationResult, CellValue } from '../types/grade';

export function exportCalculationResult(result: CalculationResult): void {
  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(buildSummary(result)), '汇总结果');
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.json_to_sheet(
      result.courses.map((course) => ({
        课程名称: course.name,
        学分: round(course.credit),
        成绩: round(course.score),
        自动换算绩点: round(course.gpa),
        '成绩 × 学分': round(course.scoreCredit),
        '学分 × 绩点': round(course.creditGpa),
        数据来源行号: course.sourceRow,
      })),
    ),
    '去重后课程表',
  );
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.json_to_sheet(
      result.duplicateGroups.map((group) => ({
        课程名称: group.courseName,
        所有出现过的成绩: group.scores.map((score) => round(score)).join(' / '),
        所有对应学分: group.credits.map((credit) => round(credit)).join(' / '),
        来源行号: group.sourceRows.join(' / '),
        最终保留成绩: round(group.keptScore),
        最终保留学分: round(group.keptCredit),
        最终保留来源: group.keptSourceRow,
        是否发生替换: group.replaced ? '是' : '否',
      })),
    ),
    '重复课程对比表',
  );
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.json_to_sheet(
      result.ignoredRows.map((row) => ({
        行号: row.rowNumber,
        忽略原因: row.reason,
        原始数据: stringifyRawValues(row.rawValues),
      })),
    ),
    '被忽略数据',
  );

  XLSX.writeFile(workbook, 'grade-calculation-result.xlsx');
}

function buildSummary(result: CalculationResult): Array<Record<string, string | number>> {
  return [
    { 项目: '算数平均分', 数值: formatNullable(result.metrics.arithmeticAverage) },
    { 项目: '加权平均分', 数值: formatNullable(result.metrics.weightedAverage) },
    { 项目: '平均学分绩点 GPA / 5.00', 数值: formatNullable(result.metrics.averageGpa) },
    { 项目: '原始课程行数', 数值: result.stats.originalRows },
    { 项目: '被忽略行数', 数值: result.stats.ignoredRows },
    { 项目: '重复课程数量', 数值: result.stats.duplicateCourseCount },
    { 项目: '重复取最高后课程数', 数值: result.stats.dedupedCourseCount },
    { 项目: '总学分', 数值: round(result.stats.totalCredits) },
    { 项目: '成绩总和 Σ成绩', 数值: round(result.stats.scoreSum) },
    { 项目: 'Σ(成绩 × 学分)', 数值: round(result.stats.scoreCreditSum) },
    { 项目: 'Σ(学分 × 绩点)', 数值: round(result.stats.creditGpaSum) },
  ];
}

function formatNullable(value: number | null): string {
  return value === null ? '-' : round(value).toFixed(2);
}

function round(value: number): number {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function stringifyRawValues(values: Record<string, CellValue>): string {
  return Object.entries(values)
    .filter(([, value]) => value !== null && value !== undefined && String(value).trim() !== '')
    .map(([key, value]) => `${key}: ${stringifyCell(value)}`)
    .join('; ');
}

function stringifyCell(value: CellValue): string {
  if (value instanceof Date) {
    return value.toLocaleDateString('zh-CN');
  }
  return value === null || value === undefined ? '' : String(value);
}
