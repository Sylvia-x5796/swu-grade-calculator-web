import * as XLSX from 'xlsx';
import type { CellValue, ColumnMapping, ParsedSheet, RawSheetRow } from '../types/grade';

const maxHeaderScanRows = 10;

export async function parseExcelFile(file: File): Promise<ParsedSheet> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });
  const sheetName = workbook.SheetNames[0];

  if (!sheetName) {
    return {
      fileName: file.name,
      sheetName: '',
      headers: [],
      rows: [],
      originalDataRowCount: 0,
      inferredMapping: {},
      needsManualMapping: true,
    };
  }

  const sheet = workbook.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json<CellValue[]>(sheet, {
    header: 1,
    defval: '',
    raw: true,
  });

  const headerRowIndex = detectHeaderRow(matrix);
  const headers = makeUniqueHeaders(matrix[headerRowIndex] ?? []);
  const dataRows = matrix.slice(headerRowIndex + 1);
  const rows: RawSheetRow[] = dataRows
    .map((cells, index) => ({
      rowNumber: headerRowIndex + index + 2,
      values: rowToValues(headers, cells),
      cells,
    }))
    .filter((row) => !isEmptyRow(row.cells));

  const inferredMapping = inferColumnMapping(headers);
  const needsManualMapping = !(
    inferredMapping.courseName &&
    inferredMapping.credit &&
    inferredMapping.score
  );

  return {
    fileName: file.name,
    sheetName,
    headers,
    rows,
    originalDataRowCount: rows.length,
    inferredMapping,
    needsManualMapping,
  };
}

export function inferColumnMapping(headers: string[]): ColumnMapping {
  return {
    courseName: findBestHeader(headers, 'courseName'),
    credit: findBestHeader(headers, 'credit'),
    score: findBestHeader(headers, 'score'),
  };
}

function detectHeaderRow(matrix: CellValue[][]): number {
  const scanLength = Math.min(matrix.length, maxHeaderScanRows);
  let bestIndex = 0;
  let bestScore = -1;

  for (let index = 0; index < scanLength; index += 1) {
    const row = matrix[index] ?? [];
    if (isEmptyRow(row)) {
      continue;
    }

    const headers = makeUniqueHeaders(row);
    const mapping = inferColumnMapping(headers);
    const matchScore =
      (mapping.courseName ? 4 : 0) + (mapping.credit ? 3 : 0) + (mapping.score ? 3 : 0);
    const filledCells = row.filter((cell) => !isEmptyCell(cell)).length;
    const score = matchScore * 10 + Math.min(filledCells, 8);

    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  }

  return bestIndex;
}

function makeUniqueHeaders(row: CellValue[]): string[] {
  const counts = new Map<string, number>();

  return row.map((cell, index) => {
    const base = stringifyCell(cell).trim() || `未命名列${index + 1}`;
    const count = counts.get(base) ?? 0;
    counts.set(base, count + 1);
    return count === 0 ? base : `${base}(${count + 1})`;
  });
}

function rowToValues(headers: string[], cells: CellValue[]): Record<string, CellValue> {
  return headers.reduce<Record<string, CellValue>>((values, header, index) => {
    values[header] = cells[index] ?? '';
    return values;
  }, {});
}

function findBestHeader(
  headers: string[],
  target: 'courseName' | 'credit' | 'score',
): string | undefined {
  let bestHeader: string | undefined;
  let bestScore = 0;

  for (const header of headers) {
    const score = scoreHeader(header, target);
    if (score > bestScore) {
      bestHeader = header;
      bestScore = score;
    }
  }

  return bestHeader;
}

function scoreHeader(header: string, target: 'courseName' | 'credit' | 'score'): number {
  const normalized = normalizeHeader(header);
  if (!normalized) {
    return 0;
  }

  if (target === 'courseName') {
    if (['课程名称', '课程名', 'coursename'].includes(normalized)) return 100;
    if (['课程', 'course', '科目', 'subject'].includes(normalized)) return 80;
    if (normalized.includes('课程') && normalized.includes('名称')) return 90;
    if (normalized.includes('course') && normalized.includes('name')) return 90;
    if (normalized.includes('课程')) return 45;
    if (normalized.includes('科目')) return 45;
    return 0;
  }

  if (target === 'credit') {
    if (['学分', 'credit', 'credits'].includes(normalized)) return 100;
    if (normalized.includes('学分')) return 90;
    if (normalized.includes('credit')) return 90;
    return 0;
  }

  if (normalized.includes('绩点') || normalized.includes('gpa') || normalized.includes('point')) {
    return 0;
  }

  if (['成绩', '分数', 'score', 'grade'].includes(normalized)) return 100;
  if (normalized.includes('成绩')) return 90;
  if (normalized.includes('分数')) return 90;
  if (normalized.includes('score')) return 90;
  if (normalized.includes('grade')) return 60;
  return 0;
}

function normalizeHeader(header: string): string {
  return header
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[：:()（）_\-./]/g, '');
}

function isEmptyRow(row: CellValue[]): boolean {
  return row.every(isEmptyCell);
}

function isEmptyCell(value: CellValue): boolean {
  return value === null || value === undefined || String(value).trim() === '';
}

function stringifyCell(value: CellValue): string {
  if (value instanceof Date) {
    return value.toLocaleDateString('zh-CN');
  }
  return value === null || value === undefined ? '' : String(value);
}
