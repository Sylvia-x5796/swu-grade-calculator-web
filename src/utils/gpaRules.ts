import type { CellValue } from '../types/grade';

const levelScoreMap: Record<string, number> = {
  优秀: 95,
  A: 95,
  良好: 85,
  B: 85,
  中等: 75,
  C: 75,
  及格: 65,
  D: 65,
  不及格: 55,
  E: 55,
};

export function toPercentScore(value: CellValue): number | null {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }

  if (typeof value !== 'string') {
    return null;
  }

  const trimmed = value.trim();
  if (!trimmed) {
    return null;
  }

  const levelKey = trimmed.replace(/\s+/g, '').toUpperCase();
  if (Object.prototype.hasOwnProperty.call(levelScoreMap, levelKey)) {
    return levelScoreMap[levelKey];
  }

  const numericText = trimmed.replace(/分$/, '').replace(/%$/, '');
  const score = Number(numericText);
  return Number.isFinite(score) ? score : null;
}

export function scoreToGpa(score: number): number {
  if (score >= 99) return 5.0;
  if (score >= 96) return 4.8;
  if (score >= 93) return 4.6;
  if (score >= 90) return 4.3;
  if (score >= 87) return 4.0;
  if (score >= 84) return 3.6;
  if (score >= 81) return 3.3;
  if (score >= 78) return 3.0;
  if (score >= 75) return 2.6;
  if (score >= 72) return 2.3;
  if (score >= 69) return 2.0;
  if (score >= 67) return 1.8;
  if (score >= 65) return 1.6;
  if (score >= 63) return 1.4;
  if (score >= 61) return 1.2;
  if (score >= 60) return 1.0;
  return 0;
}
