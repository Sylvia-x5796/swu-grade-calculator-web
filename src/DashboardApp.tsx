import { useMemo, useState } from 'react';
import type { ChangeEvent, FormEvent, ReactNode } from 'react';
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

type TargetStatus = 'idle' | 'success' | 'warm';

interface TargetMessage {
  title: string;
  text: string;
  status: TargetStatus;
}

function DashboardApp() {
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

  const duplicateCourseNames = useMemo(
    () => new Set(result.duplicateGroups.map((group) => normalizeCourseName(group.courseName))),
    [result.duplicateGroups],
  );

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
    <main className={`app-shell ${parsedSheet ? 'app-shell--dashboard' : ''}`}>
      <header className={`page-header ${parsedSheet ? 'page-header--compact' : 'page-header--hero'}`}>
        <div className="header-copy">
          <p className="eyebrow">{parsedSheet ? '上传成功' : '本地计算 · 不上传服务器'}</p>
          <h1>成绩单均分与绩点计算器</h1>
          {parsedSheet ? (
            <div className="file-status-row">
              <span className="file-chip">{parsedSheet.fileName}</span>
              <span className={isMappingComplete ? 'status-pill success' : 'status-pill warning'}>
                {isMappingComplete ? '已就绪' : '需要确认'}
              </span>
            </div>
          ) : (
            <>
              <p className="subtitle">
                上传西南大学成绩单，浏览器本地完成清洗、去重、均分和 GPA 换算。
              </p>
              <div className="feature-row" aria-label="功能亮点">
                <span>自动识别列</span>
                <span>重复课程取最高</span>
                <span>模拟成绩变化</span>
              </div>
            </>
          )}
        </div>

        <label className={`upload-box ${parsedSheet ? 'upload-box--compact' : ''}`}>
          <input
            className="file-input"
            type="file"
            accept=".xlsx"
            onChange={handleFileChange}
            disabled={isParsing}
          />
          <span className="upload-title">
            {isParsing ? '正在解析...' : parsedSheet ? '重新上传' : '上传 .xlsx 成绩单'}
          </span>
          <span className="upload-note">
            {parsedSheet
              ? '成绩单已读取完成，可以确认列信息后查看结果。'
              : '读取第一个工作表，文件只在浏览器本地处理'}
          </span>
        </label>
        {fileError && <p className="error-text header-error">{fileError}</p>}
      </header>

      {parsedSheet && !isMappingComplete && (
        <ColumnMappingCard
          headers={parsedSheet.headers}
          columnMapping={columnMapping}
          isMappingComplete={isMappingComplete}
          onMappingChange={handleMappingChange}
          prominent
        />
      )}

      {canShowCalculatedTables && (
        <>
          <section className="content-section overview-section">
            <div className="section-heading">
              <div>
                <p className="eyebrow">你的当前结果</p>
                <h2>当前成绩概览</h2>
              </div>
              <button
                className="primary-button soft-action"
                type="button"
                disabled={result.courses.length === 0}
                onClick={() => exportCalculationResult(result)}
              >
                导出结果
              </button>
            </div>

            <div className="result-grid">
              <MetricCard
                tag="AVG"
                tone="sky"
                label="算数平均分"
                value={formatMetric(result.metrics.arithmeticAverage)}
                hint="所有课程成绩的直接平均"
              />
              <MetricCard
                tag="WGT"
                tone="mint"
                label="加权平均分"
                value={formatMetric(result.metrics.weightedAverage)}
                hint="按学分权重计算"
              />
              <MetricCard
                tag="GPA"
                tone="lavender"
                label="平均学分绩点 GPA / 5.00"
                value={formatMetric(result.metrics.averageGpa)}
                hint="按西南大学 5 分制换算"
              />
            </div>
          </section>

          <ColumnMappingCard
            headers={parsedSheet!.headers}
            columnMapping={columnMapping}
            isMappingComplete={isMappingComplete}
            onMappingChange={handleMappingChange}
          />

          <section className="content-section">
            <div className="section-heading">
              <div>
                <p className="eyebrow">计算明细</p>
                <h2>数据小结</h2>
              </div>
            </div>
            <div className="stats-grid">
              <StatItem label="原始课程行数" value={result.stats.originalRows} />
              <StatItem
                label="被忽略行数"
                value={result.stats.ignoredRows}
                hint={result.stats.ignoredRows === 0 ? '很干净' : '可在下方检查'}
                variant={result.stats.ignoredRows === 0 ? 'clean' : 'warm'}
              />
              <StatItem
                label="重复课程数量"
                value={result.stats.duplicateCourseCount}
                hint={result.stats.duplicateCourseCount > 0 ? '已按最高成绩处理' : '没有重复'}
                variant={result.stats.duplicateCourseCount > 0 ? 'pink' : 'default'}
              />
              <StatItem label="重复取最高后课程数" value={result.stats.dedupedCourseCount} />
              <StatItem label="总学分" value={formatNumber(result.stats.totalCredits)} />
              <StatItem label="成绩总和" value={formatNumber(result.stats.scoreSum)} />
              <StatItem label="Σ(成绩 × 学分)" value={formatNumber(result.stats.scoreCreditSum)} />
              <StatItem label="Σ(学分 × 绩点)" value={formatNumber(result.stats.creditGpaSum)} />
            </div>
          </section>

          <section className="content-section action-section">
            <div className="action-grid">
              <article className="tool-panel manual-panel">
                <div className="panel-heading">
                  <div>
                    <p className="eyebrow">模拟成绩</p>
                    <h2>试试还没出的成绩会带来什么变化</h2>
                  </div>
                  <span className="sticker">不会修改原始文件</span>
                </div>
                <p className="panel-copy">添加一门课程后会自动重新计算，重复课程也会继续按最高成绩保留。</p>
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
                        <span className="manual-name">{course.name}</span>
                        <span>{formatNumber(course.credit)} 学分</span>
                        <span>{formatNumber(course.score)} 分</span>
                        <button
                          className="text-danger-button"
                          type="button"
                          onClick={() =>
                            setManualCourses((current) =>
                              current.filter((item) => item.id !== course.id),
                            )
                          }
                        >
                          删除
                        </button>
                      </div>
                    ))}
                  </div>
                )}
              </article>

              <article className="tool-panel target-panel">
                <div className="panel-heading">
                  <div>
                    <p className="eyebrow">目标分析</p>
                    <h2>看看离目标还差多少</h2>
                  </div>
                </div>
                <p className="panel-copy">输入目标均分，这里会把差距换成更容易理解的数字。</p>
                <div className="target-grid">
                  <label>
                    <span>目标算数平均分</span>
                    <input
                      value={targetArithmetic}
                      onChange={(event) => setTargetArithmetic(event.target.value)}
                      inputMode="decimal"
                      placeholder="88"
                    />
                    <small className={`target-message ${arithmeticTargetMessage.status}`}>
                      <span className="target-message-title">{arithmeticTargetMessage.title}</span>
                      <span>{arithmeticTargetMessage.text}</span>
                    </small>
                  </label>
                  <label>
                    <span>目标加权平均分</span>
                    <input
                      value={targetWeighted}
                      onChange={(event) => setTargetWeighted(event.target.value)}
                      inputMode="decimal"
                      placeholder="87"
                    />
                    <small className={`target-message ${weightedTargetMessage.status}`}>
                      <span className="target-message-title">{weightedTargetMessage.title}</span>
                      <span>{weightedTargetMessage.text}</span>
                    </small>
                  </label>
                </div>
              </article>
            </div>
          </section>

          <section className="content-section">
            <TableCard
              eyebrow="课程"
              title="去重后的课程表"
              description="所有最终计算都基于这里的课程列表。"
            >
              {sortedCourses.length > 0 ? (
                <div className="table-wrap">
                  <table>
                    <thead>
                      <tr>
                        <th>
                          <SortButton
                            label="课程名称"
                            active={sort.key === 'name'}
                            onClick={() => handleSort('name')}
                            direction={sort.direction}
                          />
                        </th>
                        <th>
                          <SortButton
                            label="学分"
                            active={sort.key === 'credit'}
                            onClick={() => handleSort('credit')}
                            direction={sort.direction}
                          />
                        </th>
                        <th>
                          <SortButton
                            label="成绩"
                            active={sort.key === 'score'}
                            onClick={() => handleSort('score')}
                            direction={sort.direction}
                          />
                        </th>
                        <th>自动换算绩点</th>
                        <th>成绩 × 学分</th>
                        <th>学分 × 绩点</th>
                        <th>数据来源行号</th>
                      </tr>
                    </thead>
                    <tbody>
                      {sortedCourses.map((course) => {
                        const isKeptDuplicate = duplicateCourseNames.has(normalizeCourseName(course.name));
                        return (
                          <tr key={course.id}>
                            <td>
                              <div className="course-name-cell">
                                <span>{course.name}</span>
                                <div className="tag-row">
                                  {course.sourceType === 'manual' && (
                                    <span className="tag tag-purple">模拟</span>
                                  )}
                                  {isKeptDuplicate && (
                                    <span className="tag tag-green">已保留最高</span>
                                  )}
                                </div>
                              </div>
                            </td>
                            <td>{formatNumber(course.credit)}</td>
                            <td>
                              <div className="score-cell">
                                <span>{formatNumber(course.score)}</span>
                                {course.score >= 90 && <span className="tag tag-score">高分</span>}
                              </div>
                            </td>
                            <td>{formatNumber(course.gpa)}</td>
                            <td>{formatNumber(course.scoreCredit)}</td>
                            <td>{formatNumber(course.creditGpa)}</td>
                            <td>{course.sourceRow}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              ) : (
                <EmptyState text={hasAnyResult ? '没有可用于计算的课程。' : '上传成绩单后显示课程明细。'} />
              )}
            </TableCard>
          </section>

          <section className="content-section">
            <TableCard
              eyebrow="重复"
              title="重复课程对比表"
              description="同名课程只保留成绩最高的一条，成绩相同时保留学分更高的一条。"
            >
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
                        <tr
                          className={group.replaced ? 'duplicate-replaced' : undefined}
                          key={group.courseName}
                        >
                          <td>{group.courseName}</td>
                          <td>{group.scores.map((score) => formatNumber(score)).join(' / ')}</td>
                          <td>{group.credits.map((credit) => formatNumber(credit)).join(' / ')}</td>
                          <td>
                            <span className="kept-score">{formatNumber(group.keptScore)}</span>
                            <span className="muted-inline"> · {group.keptSourceRow}</span>
                          </td>
                          <td>
                            <span className={group.replaced ? 'tag tag-warm' : 'tag'}>
                              {group.replaced ? '是' : '否'}
                            </span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : (
                <EmptyState text="暂无重复课程，这份成绩单很清爽。" />
              )}
            </TableCard>
          </section>

          <section className="content-section">
            <TableCard
              eyebrow="清洗"
              title="被忽略数据"
              description="这些行没有参与计算，可以检查是否需要手动处理。"
            >
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
                <EmptyState text="暂无被忽略数据，当前识别结果很干净。" />
              )}
            </TableCard>
          </section>

          <section className="export-panel">
            <div>
              <p className="eyebrow">导出</p>
              <h2>保存这次计算结果</h2>
              <p>导出的文件会包含汇总、去重课程、重复课程和被忽略数据四个工作表。</p>
            </div>
            <button
              className="primary-button soft-action"
              type="button"
              disabled={result.courses.length === 0}
              onClick={() => exportCalculationResult(result)}
            >
              导出 Excel
            </button>
          </section>
        </>
      )}
    </main>
  );
}

interface ColumnMappingCardProps {
  headers: string[];
  columnMapping: ColumnMapping;
  isMappingComplete: boolean;
  onMappingChange: (key: keyof ColumnMapping, value: string) => void;
  prominent?: boolean;
}

function ColumnMappingCard({
  headers,
  columnMapping,
  isMappingComplete,
  onMappingChange,
  prominent = false,
}: ColumnMappingCardProps) {
  return (
    <section className={`mapping-card ${prominent ? 'mapping-card--prominent' : ''}`}>
      <div className="mapping-copy">
        <div>
          <p className="eyebrow">确认识别到的列</p>
          <h2>系统已自动匹配课程名称、学分和成绩</h2>
        </div>
        <p>如有错误，可以在这里轻轻调整一下。</p>
      </div>
      <span className={isMappingComplete ? 'status-pill success' : 'status-pill warning'}>
        {isMappingComplete ? '已识别' : '需要确认'}
      </span>
      <div className="mapping-grid">
        <FieldSelect
          label="课程名称列"
          value={columnMapping.courseName ?? ''}
          headers={headers}
          onChange={(value) => onMappingChange('courseName', value)}
        />
        <FieldSelect
          label="学分列"
          value={columnMapping.credit ?? ''}
          headers={headers}
          onChange={(value) => onMappingChange('credit', value)}
        />
        <FieldSelect
          label="成绩列"
          value={columnMapping.score ?? ''}
          headers={headers}
          onChange={(value) => onMappingChange('score', value)}
        />
      </div>
    </section>
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

function MetricCard({
  tag,
  tone,
  label,
  value,
  hint,
}: {
  tag: string;
  tone: 'sky' | 'mint' | 'lavender';
  label: string;
  value: string;
  hint: string;
}) {
  return (
    <article className={`result-card result-card--${tone}`}>
      <span className="result-tag">{tag}</span>
      <span className="result-title">{label}</span>
      <strong className="result-value">{value}</strong>
      <p className="result-desc">{hint}</p>
    </article>
  );
}

function StatItem({
  label,
  value,
  hint,
  variant = 'default',
}: {
  label: string;
  value: string | number;
  hint?: string;
  variant?: 'default' | 'clean' | 'warm' | 'pink';
}) {
  return (
    <div className={`stat-card stat-card--${variant}`}>
      <span>{label}</span>
      <strong>{value}</strong>
      {hint && <small>{hint}</small>}
    </div>
  );
}

function TableCard({
  eyebrow,
  title,
  description,
  children,
}: {
  eyebrow: string;
  title: string;
  description: string;
  children: ReactNode;
}) {
  return (
    <article className="table-card">
      <div className="table-card-heading">
        <div>
          <p className="eyebrow">{eyebrow}</p>
          <h2>{title}</h2>
        </div>
        <p>{description}</p>
      </div>
      {children}
    </article>
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

function buildArithmeticTargetMessage(targetText: string, result: CalculationResult): TargetMessage {
  const target = parseTarget(targetText);
  if (targetText.trim() === '') {
    return { title: '等你输入目标', text: '输入目标后，这里会给你一个轻松的小估算。', status: 'idle' };
  }
  if (target === null) {
    return { title: '格式看一下', text: '填一个 0-100 之间的目标分数就好。', status: 'warm' };
  }
  if (result.courses.length === 0) {
    return { title: '暂无数据', text: '当前还没有可计算课程。', status: 'idle' };
  }

  const requiredScoreSum = target * result.courses.length;
  const deficit = requiredScoreSum - result.stats.scoreSum;
  const current = formatMetric(result.metrics.arithmeticAverage);

  if (deficit <= 0) {
    return {
      title: '已经达到目标',
      text: `当前算数平均分 ${current}，已经达到 ${formatNumber(target)}，可以安心一点啦。`,
      status: 'success',
    };
  }

  return {
    title: '还差一点点',
    text: `距离算数平均分 ${formatNumber(target)}，还差 ${formatNumber(deficit)} 分总成绩。`,
    status: 'warm',
  };
}

function buildWeightedTargetMessage(targetText: string, result: CalculationResult): TargetMessage {
  const target = parseTarget(targetText);
  if (targetText.trim() === '') {
    return { title: '等你输入目标', text: '输入目标后，这里会估算还差多少加权分。', status: 'idle' };
  }
  if (target === null) {
    return { title: '格式看一下', text: '填一个 0-100 之间的目标分数就好。', status: 'warm' };
  }
  if (result.courses.length === 0 || result.stats.totalCredits <= 0) {
    return { title: '暂无数据', text: '当前还没有可计算学分。', status: 'idle' };
  }

  const requiredWeightedPoints = target * result.stats.totalCredits;
  const deficit = requiredWeightedPoints - result.stats.scoreCreditSum;
  const current = formatMetric(result.metrics.weightedAverage);

  if (deficit <= 0) {
    return {
      title: '已经达到目标',
      text: `当前加权平均分 ${current}，已经达到 ${formatNumber(target)}，可以安心一点啦。`,
      status: 'success',
    };
  }

  return {
    title: '还差一点点',
    text: `距离加权平均分 ${formatNumber(target)}，还差 ${formatNumber(deficit)} 个加权分。`,
    status: 'warm',
  };
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

function normalizeCourseName(name: string): string {
  return name.trim().toLocaleLowerCase('zh-CN');
}

function createId(): string {
  if (typeof crypto !== 'undefined' && 'randomUUID' in crypto) {
    return crypto.randomUUID();
  }
  return `manual-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

export default DashboardApp;
