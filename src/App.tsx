import { useMemo, useState } from 'react';
import type { ChangeEvent, FormEvent } from 'react';
import type {
  CalculationResult,
  CellValue,
  ColumnMapping,
  ManualCourse,
  SortDirection,
  SortKey,
} from './types/grade';
import { calculateGrades } from './utils/calculateGrades';
import { exportCalculationResult } from './utils/exportExcel';
import { toPercentScore } from './utils/gpaRules';
import { parseExcelFile } from './utils/parseExcel';
import type { ParsedSheet } from './types/grade';

const defaultSort: { key: SortKey; direction: SortDirection } = {
  key: 'score',
  direction: 'desc',
};

function App() {
  const [parsedSheet, setParsedSheet] = useState<ParsedSheet | null>(null);
  const [columnMapping, setColumnMapping] = useState<ColumnMapping>({});
  const [manualCourses, setManualCourses] = useState<ManualCourse[]>([]);
  const [manualDraft, setManualDraft] = useState({ name: '', credit: '', score: '' });
  const [manualError, setManualError] = useState('');
  const [targetArithmetic, setTargetArithmetic] = useState('');
  const [targetWeighted, setTargetWeighted] = useState('');
  const [sort, setSort] = useState(defaultSort);
  const [isParsing, setIsParsing] = useState(false);
  const [fileError, setFileError] = useState('');

  const isMappingComplete = Boolean(
    columnMapping.courseName && columnMapping.credit && columnMapping.score,
  );

  const result = useMemo(
    () =>
      calculateGrades(
        parsedSheet?.rows ?? [],
        isMappingComplete ? columnMapping : {},
        manualCourses,
        parsedSheet?.originalDataRowCount ?? 0,
      ),
    [columnMapping, isMappingComplete, manualCourses, parsedSheet],
  );

  const sortedCourses = useMemo(() => {
    return [...result.courses].sort((a, b) => {
      const direction = sort.direction === 'asc' ? 1 : -1;
      if (sort.key === 'name') {
        return a.name.localeCompare(b.name, 'zh-CN') * direction;
      }
      return (a[sort.key] - b[sort.key]) * direction;
    });
  }, [result.courses, sort]);

  const arithmeticTargetMessage = useMemo(
    () => buildArithmeticTargetMessage(targetArithmetic, result),
    [targetArithmetic, result],
  );

  const weightedTargetMessage = useMemo(
    () => buildWeightedTargetMessage(targetWeighted, result),
    [targetWeighted, result],
  );

  async function handleFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    if (!file.name.toLowerCase().endsWith('.xlsx')) {
      setFileError('请上传 .xlsx 格式的成绩单文件。');
      return;
    }

    setIsParsing(true);
    setFileError('');

    try {
      const parsed = await parseExcelFile(file);
      setParsedSheet(parsed);
      setColumnMapping(parsed.inferredMapping);
      setManualCourses([]);
      setSort(defaultSort);
    } catch (error) {
      setParsedSheet(null);
      setColumnMapping({});
      setFileError(error instanceof Error ? error.message : '文件解析失败，请检查成绩单格式。');
    } finally {
      setIsParsing(false);
      event.target.value = '';
    }
  }

  function handleMappingChange(key: keyof ColumnMapping, value: string) {
    setColumnMapping((current) => ({
      ...current,
      [key]: value || undefined,
    }));
  }

  function handleAddManualCourse(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const name = manualDraft.name.trim();
    const credit = Number(manualDraft.credit.trim());
    const score = toPercentScore(manualDraft.score);

    if (!name) {
      setManualError('请输入课程名称。');
      return;
    }

    if (!Number.isFinite(credit) || credit < 0) {
      setManualError('请输入有效学分。');
      return;
    }

    if (score === null || score < 0 || score > 100) {
      setManualError('请输入 0-100 的成绩，或优秀、良好、中等、及格、不及格。');
      return;
    }

    setManualCourses((current) => [
      ...current,
      {
        id: createId(),
        name,
        credit,
        score,
      },
    ]);
    setManualDraft({ name: '', credit: '', score: '' });
    setManualError('');
  }

  function handleSort(nextKey: SortKey) {
    setSort((current) => {
      if (current.key === nextKey) {
        return {
          key: nextKey,
          direction: current.direction === 'asc' ? 'desc' : 'asc',
        };
      }
      return {
        key: nextKey,
        direction: nextKey === 'name' ? 'asc' : 'desc',
      };
    });
  }

  const canShowCalculatedTables = Boolean(parsedSheet && isMappingComplete);
  const hasAnyResult = result.courses.length > 0 || result.ignoredRows.length > 0;

  return (
    <main className="app-shell">
      <header className="page-header">
        <div className="header-copy">
          <p className="eyebrow">西南大学 5 分制绩点</p>
          <h1>成绩单均分与绩点计算器</h1>
          <p className="subtitle">本地解析成绩单，自动清洗、去重、换算 GPA，并支持模拟成绩。</p>
        </div>

        <label className="upload-box">
          <input
            className="file-input"
            type="file"
            accept=".xlsx"
            onChange={handleFileChange}
            disabled={isParsing}
          />
          <span className="upload-title">{isParsing ? '正在解析...' : '上传 .xlsx 成绩单'}</span>
          <span className="upload-note">
            {parsedSheet
              ? `${parsedSheet.fileName} · ${parsedSheet.sheetName}`
              : '读取第一个工作表，文件只在浏览器本地处理'}
          </span>
        </label>
        {fileError && <p className="error-text">{fileError}</p>}
      </header>

      {parsedSheet && (
        <section className="content-section">
          <div className="section-heading">
            <div>
              <p className="eyebrow">列识别</p>
              <h2>确认成绩单列</h2>
            </div>
            <span className={isMappingComplete ? 'status-pill success' : 'status-pill warning'}>
              {isMappingComplete ? '已就绪' : '需要选择列'}
            </span>
          </div>

          <div className="mapping-grid">
            <FieldSelect
              label="课程名称列"
              value={columnMapping.courseName ?? ''}
              headers={parsedSheet.headers}
              onChange={(value) => handleMappingChange('courseName', value)}
            />
            <FieldSelect
              label="学分列"
              value={columnMapping.credit ?? ''}
              headers={parsedSheet.headers}
              onChange={(value) => handleMappingChange('credit', value)}
            />
            <FieldSelect
              label="成绩列"
              value={columnMapping.score ?? ''}
              headers={parsedSheet.headers}
              onChange={(value) => handleMappingChange('score', value)}
            />
          </div>

          {!isMappingComplete && (
            <p className="notice">
              未能完整识别课程名称、学分、成绩三列，请手动选择后继续计算。
            </p>
          )}
        </section>
      )}

      {canShowCalculatedTables && (
        <>
          <section className="content-section">
            <div className="section-heading">
              <div>
                <p className="eyebrow">结果</p>
                <h2>主要指标</h2>
              </div>
              <button
                className="primary-button"
                type="button"
                disabled={result.courses.length === 0}
                onClick={() => exportCalculationResult(result)}
              >
                导出结果
              </button>
            </div>

            <div className="result-grid">
              <MetricCard
                label="算数平均分"
                value={formatMetric(result.metrics.arithmeticAverage)}
              />
              <MetricCard
                label="加权平均分"
                value={formatMetric(result.metrics.weightedAverage)}
              />
              <MetricCard
                label="平均学分绩点 GPA / 5.00"
                value={formatMetric(result.metrics.averageGpa)}
              />
            </div>
          </section>

          <section className="content-section">
            <div className="section-heading">
              <div>
                <p className="eyebrow">统计</p>
                <h2>详细统计</h2>
              </div>
            </div>
            <div className="stats-grid">
              <StatItem label="原始课程行数" value={result.stats.originalRows} />
              <StatItem label="被忽略行数" value={result.stats.ignoredRows} />
              <StatItem label="重复课程数量" value={result.stats.duplicateCourseCount} />
              <StatItem label="重复取最高后课程数" value={result.stats.dedupedCourseCount} />
              <StatItem label="总学分" value={formatNumber(result.stats.totalCredits)} />
              <StatItem label="成绩总和 Σ成绩" value={formatNumber(result.stats.scoreSum)} />
              <StatItem
                label="Σ(成绩 × 学分)"
                value={formatNumber(result.stats.scoreCreditSum)}
              />
              <StatItem
                label="Σ(学分 × 绩点)"
                value={formatNumber(result.stats.creditGpaSum)}
              />
            </div>
          </section>

          <section className="content-section">
            <div className="section-heading">
              <div>
                <p className="eyebrow">模拟</p>
                <h2>手动新增课程 / 模拟成绩</h2>
              </div>
            </div>
            <form className="manual-form" onSubmit={handleAddManualCourse}>
              <label>
                <span>课程名称</span>
                <input
                  value={manualDraft.name}
                  onChange={(event) =>
                    setManualDraft((current) => ({ ...current, name: event.target.value }))
                  }
                  placeholder="例如：高等数学"
                />
              </label>
              <label>
                <span>学分</span>
                <input
                  value={manualDraft.credit}
                  onChange={(event) =>
                    setManualDraft((current) => ({ ...current, credit: event.target.value }))
                  }
                  inputMode="decimal"
                  placeholder="2"
                />
              </label>
              <label>
                <span>成绩</span>
                <input
                  value={manualDraft.score}
                  onChange={(event) =>
                    setManualDraft((current) => ({ ...current, score: event.target.value }))
                  }
                  placeholder="99 或 优秀"
                />
              </label>
              <button className="secondary-button" type="submit">
                添加
              </button>
            </form>
            {manualError && <p className="error-text">{manualError}</p>}
            {manualCourses.length > 0 && (
              <div className="manual-list">
                {manualCourses.map((course) => (
                  <div className="manual-row" key={course.id}>
                    <span>{course.name}</span>
                    <span>{formatNumber(course.credit)} 学分</span>
                    <span>{formatNumber(course.score)} 分</span>
                    <button
                      className="ghost-button"
                      type="button"
                      onClick={() =>
                        setManualCourses((current) => current.filter((item) => item.id !== course.id))
                      }
                    >
                      删除
                    </button>
                  </div>
                ))}
              </div>
            )}
          </section>

          <section className="content-section">
            <div className="section-heading">
              <div>
                <p className="eyebrow">目标</p>
                <h2>目标分析</h2>
              </div>
            </div>
            <div className="target-grid">
              <label>
                <span>目标算数平均分</span>
                <input
                  value={targetArithmetic}
                  onChange={(event) => setTargetArithmetic(event.target.value)}
                  inputMode="decimal"
                  placeholder="88"
                />
                <small>{arithmeticTargetMessage}</small>
              </label>
              <label>
                <span>目标加权平均分</span>
                <input
                  value={targetWeighted}
                  onChange={(event) => setTargetWeighted(event.target.value)}
                  inputMode="decimal"
                  placeholder="87"
                />
                <small>{weightedTargetMessage}</small>
              </label>
            </div>
          </section>

          <section className="content-section">
            <div className="section-heading">
              <div>
                <p className="eyebrow">课程</p>
                <h2>去重后的课程表</h2>
              </div>
            </div>
            {sortedCourses.length > 0 ? (
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>
                        <SortButton label="课程名称" active={sort.key === 'name'} onClick={() => handleSort('name')} direction={sort.direction} />
                      </th>
                      <th>
                        <SortButton label="学分" active={sort.key === 'credit'} onClick={() => handleSort('credit')} direction={sort.direction} />
                      </th>
                      <th>
                        <SortButton label="成绩" active={sort.key === 'score'} onClick={() => handleSort('score')} direction={sort.direction} />
                      </th>
                      <th>自动换算绩点</th>
                      <th>成绩 × 学分</th>
                      <th>学分 × 绩点</th>
                      <th>数据来源行号</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sortedCourses.map((course) => (
                      <tr key={course.id}>
                        <td>{course.name}</td>
                        <td>{formatNumber(course.credit)}</td>
                        <td>{formatNumber(course.score)}</td>
                        <td>{formatNumber(course.gpa)}</td>
                        <td>{formatNumber(course.scoreCredit)}</td>
                        <td>{formatNumber(course.creditGpa)}</td>
                        <td>{course.sourceRow}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <EmptyState text={hasAnyResult ? '没有可用于计算的课程。' : '上传成绩单或添加模拟课程后显示。'} />
            )}
          </section>

          <section className="content-section">
            <div className="section-heading">
              <div>
                <p className="eyebrow">重复</p>
                <h2>重复课程对比表</h2>
              </div>
            </div>
            {result.duplicateGroups.length > 0 ? (
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>课程名称</th>
                      <th>所有出现过的成绩</th>
                      <th>所有对应学分</th>
                      <th>最终保留成绩</th>
                      <th>是否发生替换</th>
                    </tr>
                  </thead>
                  <tbody>
                    {result.duplicateGroups.map((group) => (
                      <tr className={group.replaced ? 'duplicate-replaced' : undefined} key={group.courseName}>
                        <td>{group.courseName}</td>
                        <td>{group.scores.map((score) => formatNumber(score)).join(' / ')}</td>
                        <td>{group.credits.map((credit) => formatNumber(credit)).join(' / ')}</td>
                        <td>
                          {formatNumber(group.keptScore)}
                          <span className="muted-inline"> · {group.keptSourceRow}</span>
                        </td>
                        <td>
                          <span className={group.replaced ? 'tag changed' : 'tag'}>
                            {group.replaced ? '是' : '否'}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <EmptyState text="暂无重复课程。" />
            )}
          </section>

          <section className="content-section">
            <div className="section-heading">
              <div>
                <p className="eyebrow">清洗</p>
                <h2>被忽略数据</h2>
              </div>
            </div>
            {result.ignoredRows.length > 0 ? (
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>行号</th>
                      <th>忽略原因</th>
                      <th>原始数据</th>
                    </tr>
                  </thead>
                  <tbody>
                    {result.ignoredRows.map((row) => (
                      <tr key={row.id}>
                        <td>{row.rowNumber}</td>
                        <td>{row.reason}</td>
                        <td>{stringifyRawValues(row.rawValues)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <EmptyState text="暂无被忽略数据。" />
            )}
          </section>
        </>
      )}
    </main>
  );
}

interface FieldSelectProps {
  label: string;
  value: string;
  headers: string[];
  onChange: (value: string) => void;
}

function FieldSelect({ label, value, headers, onChange }: FieldSelectProps) {
  return (
    <label className="field-select">
      <span>{label}</span>
      <select value={value} onChange={(event) => onChange(event.target.value)}>
        <option value="">请选择</option>
        {headers.map((header) => (
          <option value={header} key={header}>
            {header}
          </option>
        ))}
      </select>
    </label>
  );
}

function MetricCard({ label, value }: { label: string; value: string }) {
  return (
    <article className="result-card">
      <span>{label}</span>
      <strong>{value}</strong>
    </article>
  );
}

function StatItem({ label, value }: { label: string; value: string | number }) {
  return (
    <div className="stat-item">
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}

function SortButton({
  label,
  active,
  direction,
  onClick,
}: {
  label: string;
  active: boolean;
  direction: SortDirection;
  onClick: () => void;
}) {
  return (
    <button className="sort-button" type="button" onClick={onClick}>
      {label}
      <span aria-hidden="true">{active ? (direction === 'asc' ? '↑' : '↓') : '↕'}</span>
    </button>
  );
}

function EmptyState({ text }: { text: string }) {
  return <div className="empty-state">{text}</div>;
}

function buildArithmeticTargetMessage(targetText: string, result: CalculationResult): string {
  const target = parseTarget(targetText);
  if (targetText.trim() === '') {
    return '输入目标后显示还差多少总成绩分。';
  }
  if (target === null) {
    return '请输入 0-100 的目标分数。';
  }
  if (result.courses.length === 0) {
    return '当前还没有可计算课程。';
  }

  const requiredScoreSum = target * result.courses.length;
  const deficit = requiredScoreSum - result.stats.scoreSum;
  const current = formatMetric(result.metrics.arithmeticAverage);

  if (deficit <= 0) {
    return `当前算数平均分 ${current}，已经达到目标 ${formatNumber(target)}。`;
  }

  return `当前算数平均分 ${current}，按 ${result.courses.length} 门课程计算，目标需要 Σ成绩 ${formatNumber(
    requiredScoreSum,
  )}，还差 ${formatNumber(deficit)} 个总成绩分。`;
}

function buildWeightedTargetMessage(targetText: string, result: CalculationResult): string {
  const target = parseTarget(targetText);
  if (targetText.trim() === '') {
    return '输入目标后显示还差多少加权分。';
  }
  if (target === null) {
    return '请输入 0-100 的目标分数。';
  }
  if (result.courses.length === 0 || result.stats.totalCredits <= 0) {
    return '当前还没有可计算学分。';
  }

  const requiredWeightedPoints = target * result.stats.totalCredits;
  const deficit = requiredWeightedPoints - result.stats.scoreCreditSum;
  const current = formatMetric(result.metrics.weightedAverage);

  if (deficit <= 0) {
    return `当前加权平均分 ${current}，已经达到目标 ${formatNumber(target)}。`;
  }

  const twoCreditExample = deficit / 2;
  return `当前加权平均分 ${current}，目标需要 Σ(成绩 × 学分) ${formatNumber(
    requiredWeightedPoints,
  )}，还差 ${formatNumber(deficit)} 个“学分 × 提高分数”。相当于 2 学分课程合计提高 ${formatNumber(
    twoCreditExample,
  )} 分。`;
}

function parseTarget(value: string): number | null {
  const parsed = Number(value.trim());
  if (!Number.isFinite(parsed) || parsed < 0 || parsed > 100) {
    return null;
  }
  return parsed;
}

function formatMetric(value: number | null): string {
  return value === null ? '--' : value.toFixed(2);
}

function formatNumber(value: number): string {
  const rounded = Math.round((value + Number.EPSILON) * 100) / 100;
  return Number.isInteger(rounded) ? String(rounded) : rounded.toFixed(2);
}

function stringifyRawValues(values: Record<string, CellValue>): string {
  const text = Object.entries(values)
    .filter(([, value]) => value !== null && value !== undefined && String(value).trim() !== '')
    .map(([key, value]) => `${key}: ${stringifyCell(value)}`)
    .join('；');
  return text || '空';
}

function stringifyCell(value: CellValue): string {
  if (value instanceof Date) {
    return value.toLocaleDateString('zh-CN');
  }
  return value === null || value === undefined ? '' : String(value);
}

function createId(): string {
  if (typeof crypto !== 'undefined' && 'randomUUID' in crypto) {
    return crypto.randomUUID();
  }
  return `manual-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

export default App;
