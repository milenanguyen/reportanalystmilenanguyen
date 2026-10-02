"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  Bar,
  CartesianGrid,
  ComposedChart,
  Legend,
  Line,
  LineChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import {
  buildDailyMetrics,
  buildDriverAnalysis,
  buildWeeklyMetrics,
  formatCurrency,
  formatMonthDay,
  formatNumber,
  formatPercent,
  formatRatio,
  getPercentChange,
  parseUploadedRows,
  type DailyMetric,
  type ImportedRow,
  type WeeklyMetric,
} from "./dashboard-utils";

function formatMetricValue(value: number | null, fallback = "N/A") {
  if (value === null || !Number.isFinite(value)) {
    return fallback;
  }

  return value.toFixed(2);
}

function formatTrend(delta: number | null, lowerIsBetter = false) {
  if (delta === null || !Number.isFinite(delta)) {
    return { label: "—", improvement: false };
  }

  const improvement = lowerIsBetter ? delta < 0 : delta > 0;
  return {
    label: `${delta >= 0 ? "+" : ""}${delta.toFixed(1)}%`,
    improvement,
  };
}

function renderMetricText(text: string) {
  const match = text.match(/^(.+?):\s*(.+)$/);
  if (
    match &&
    /^(Current ROAS|Previous ROAS|Current status|Primary driver|CPI|UV|CTR|CVR|ROAS)$/.test(match[1])
  ) {
    return (
      <>
        <span className="font-medium text-slate-700">{match[1]}:</span>{" "}
        {renderImportantNumbers(match[2])}
      </>
    );
  }

  return <>{text}</>;
}

function renderImportantNumbers(text: string) {
  return text.split(/([+-]?\$?\d[\d,]*(?:\.\d+)?%?)/g).map((part, index) =>
    /^[+-]?\$?\d[\d,]*(?:\.\d+)?%?$/.test(part) ? (
      <strong key={`${part}-${index}`} className="font-semibold text-slate-900">
        {part}
      </strong>
    ) : (
      <span key={`${part}-${index}`}>{part}</span>
    ),
  );
}

function splitRecommendationAction(action: string) {
  const match = action.match(/^\d+\.\s*(.*?)(?:\.\s*(.*))?$/);
  if (!match) {
    return { title: action, detail: "" };
  }

  return {
    title: match[1] ?? action,
    detail: match[2] ?? "",
  };
}

function renderParagraphText(text: string) {
  const parts = text.split(/(CPI|UV|CTR|CVR|ROAS)/g);
  return (
    <>
      {parts.map((part, index) =>
        /^(CPI|UV|CTR|CVR|ROAS)$/.test(part) ? (
          <strong key={`${part}-${index}`} className="text-slate-900">
            {part}
          </strong>
        ) : (
          <span key={`${part}-${index}`}>{renderImportantNumbers(part)}</span>
        ),
      )}
    </>
  );
}

export default function Home() {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [rawRows, setRawRows] = useState<ImportedRow[]>([]);
  const [weeklyMetrics, setWeeklyMetrics] = useState<WeeklyMetric[]>([]);
  const [dailyMetrics, setDailyMetrics] = useState<DailyMetric[]>([]);
  const [selectedWeekKey, setSelectedWeekKey] = useState<string>("");
  const [errorMessage, setErrorMessage] = useState<string>("");
  const [fileName, setFileName] = useState<string>("");

  const processFile = async (file: File) => {
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

      if (!firstSheet) {
        throw new Error("The uploaded file does not contain a valid worksheet.");
      }

      const rows = XLSX.utils.sheet_to_json(firstSheet, {
        defval: "",
        raw: false,
      }) as Array<Record<string, unknown>>;

      const parsedRows = parseUploadedRows(rows);

      if (parsedRows.length === 0) {
        throw new Error("No usable rows were found in the uploaded Excel file.");
      }

      const weekly = buildWeeklyMetrics(parsedRows);
      const daily = buildDailyMetrics(parsedRows);

      setRawRows(parsedRows);
      setWeeklyMetrics(weekly);
      setDailyMetrics(daily);
      setSelectedWeekKey(weekly[0]?.weekKey ?? "");
      setErrorMessage("");
      setFileName(file.name);
    } catch (error) {
      setRawRows([]);
      setWeeklyMetrics([]);
      setDailyMetrics([]);
      setSelectedWeekKey("");
      setErrorMessage(
        error instanceof Error
          ? error.message
          : "The Excel file could not be processed. Please upload a valid campaign file.",
      );
      setFileName("");
    }
  };

  const handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0];
    if (!selectedFile) {
      return;
    }

    await processFile(selectedFile);

    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  const sortedWeeklyMetrics = useMemo(
    () => [...weeklyMetrics].sort((a, b) => b.weekStart.getTime() - a.weekStart.getTime()),
    [weeklyMetrics],
  );

  useEffect(() => {
    if (sortedWeeklyMetrics.length > 0 && !sortedWeeklyMetrics.some((week) => week.weekKey === selectedWeekKey)) {
      setSelectedWeekKey(sortedWeeklyMetrics[0].weekKey);
    }
  }, [selectedWeekKey, sortedWeeklyMetrics]);

  const selectedWeek = sortedWeeklyMetrics.find((week) => week.weekKey === selectedWeekKey) ?? null;
  const selectedWeekIndex = selectedWeek
    ? sortedWeeklyMetrics.findIndex((week) => week.weekKey === selectedWeek.weekKey)
    : -1;
  const previousWeek = selectedWeekIndex > 0 ? sortedWeeklyMetrics[selectedWeekIndex + 1] ?? null : null;
  const driverAnalysis = selectedWeek && previousWeek ? buildDriverAnalysis(selectedWeek, previousWeek) : null;

  const conclusionTitle = driverAnalysis?.title ?? "No driver identified";
  const conclusionSummary = driverAnalysis?.summary ?? "No driver summary is available yet.";

  const conclusionTone =
    conclusionTitle === "Good Performance"
      ? "border-violet-200 bg-violet-50 text-violet-800"
      : conclusionTitle === "Profitable but Below Previous Week"
        ? "border-amber-200 bg-amber-50 text-amber-800"
        : conclusionTitle === "Good Signal — Not Yet Profitable"
          ? "border-purple-200 bg-purple-50 text-purple-800"
          : "border-rose-200 bg-rose-50 text-rose-800";

  const weeklyComparisonData =
    selectedWeek && previousWeek
      ? [
          {
            name: "Previous Week",
            roas: previousWeek.roas,
            cpi: previousWeek.cpi,
            uv: previousWeek.uv,
          },
          {
            name: "Selected Week",
            roas: selectedWeek.roas,
            cpi: selectedWeek.cpi,
            uv: selectedWeek.uv,
          },
        ]
      : [];

  return (
    <div className="dashboard-canvas min-h-screen px-3 py-4 text-slate-900 sm:px-6 sm:py-8 lg:px-8">
      <div className="dashboard-shell mx-auto max-w-7xl">
        <header className="dashboard-header mb-7 border-b border-slate-200 pb-5">
          <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
            <div>
              <p className="text-[11px] font-semibold uppercase tracking-[0.16em] text-violet-700">Fitness App</p>
              <h1 className="mt-1 text-2xl font-semibold tracking-tight text-slate-900 sm:text-3xl">
                User Acquisition Performance
              </h1>
              <p className="mt-1 text-sm text-slate-500">Weekly and daily acquisition performance analysis</p>
            </div>

            <div className="flex items-center gap-3">
              <input
                ref={fileInputRef}
                id="excel-upload"
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={handleFileChange}
                className="hidden"
              />
              <label
                htmlFor="excel-upload"
                className="inline-flex cursor-pointer items-center justify-center rounded-lg border border-violet-200 bg-white px-3.5 py-2 text-sm font-semibold text-violet-800 transition hover:border-violet-300 hover:bg-violet-50"
              >
                Upload Excel
              </label>
            </div>
          </div>
        </header>

        {errorMessage ? (
          <div className="mb-6 rounded-xl border border-red-200 bg-red-50 p-4 text-sm text-red-700">
            {errorMessage}
          </div>
        ) : null}

        {fileName ? (
          <div className="mb-6 flex items-center gap-2 text-sm text-slate-600">
            <span className="inline-flex rounded-full bg-slate-200 px-2.5 py-1 font-medium text-slate-700">
              File loaded
            </span>
            <span>{fileName}</span>
          </div>
        ) : null}

        {!errorMessage && sortedWeeklyMetrics.length === 0 ? (
          <div className="mb-6 rounded-2xl border border-dashed border-slate-300 bg-white p-10 text-center text-slate-600">
            Upload an Excel file to calculate weekly performance and driver analysis.
          </div>
        ) : null}

        {sortedWeeklyMetrics.length > 0 ? (
          <>
            <section className="dashboard-section mb-7">
              <div className="mb-4 flex items-center justify-between gap-3">
                <div>
                  <p className="text-[11px] font-semibold uppercase tracking-[0.16em] text-slate-500">Overview</p>
                  <h2 className="mt-1 text-xl font-semibold text-slate-900">Weekly Performance Overview</h2>
                </div>
              </div>

              <div className="weekly-table-scroll max-h-[350px] overflow-y-auto overflow-x-auto rounded-xl border border-slate-200 bg-white">
                <table className="min-w-full border-separate border-spacing-0 text-left text-sm">
                  <thead className="bg-violet-50/90">
                    <tr>
                      <th className="px-3 py-2.5 text-left text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-600">Week</th>
                      <th className="px-3 py-2.5 text-right text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-600">Cost</th>
                      <th className="px-3 py-2.5 text-right text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-600">CPI</th>
                      <th className="roas-heading px-3 py-2.5 text-right text-[11px] font-semibold uppercase tracking-[0.12em]">ROAS</th>
                      <th className="px-3 py-2.5 text-right text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-600">UV</th>
                      <th className="px-3 py-2.5 text-right text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-600">CTR</th>
                      <th className="px-3 py-2.5 text-right text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-600">CVR</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sortedWeeklyMetrics.map((week, index) => {
                      const previous = sortedWeeklyMetrics[index + 1] ?? null;

                      const renderTrendCell = (
                        current: number | null,
                        previousValue: number | null,
                        lowerIsBetter = false,
                        format: (value: number | null) => string,
                      ) => {
                        const delta = getPercentChange(current, previousValue);
                        const trend = formatTrend(delta, lowerIsBetter);
                        const value = format(current);

                        return (
                          <div className="space-y-1">
                            <div className="text-right font-semibold text-slate-900">{value}</div>
                            <div
                              className={`ml-auto inline-flex rounded-full px-1.5 py-0.5 text-[10px] font-semibold ${
                                delta === null || !Number.isFinite(delta)
                                  ? "bg-slate-100 text-slate-500"
                                  : trend.improvement
                                    ? "bg-emerald-100 text-emerald-700"
                                    : "bg-rose-100 text-rose-700"
                              }`}
                            >
                              {trend.label}
                            </div>
                          </div>
                        );
                      };

                      return (
                        <tr key={week.weekKey} className="weekly-table-row border-t border-slate-100 bg-white text-slate-700 hover:bg-violet-50/40">
                          <td className="px-3 py-3 text-left font-medium text-slate-900">{week.weekLabel}</td>
                          <td className="px-3 py-3 text-right">{renderTrendCell(week.cost, previous?.cost ?? null, true, (value) => (value === null ? "N/A" : formatCurrency(value)))}</td>
                          <td className="px-3 py-3 text-right">{renderTrendCell(week.cpi, previous?.cpi ?? null, true, (value) => (value === null ? "N/A" : formatMetricValue(value)))}</td>
                          <td className="roas-cell px-3 py-3 text-right">{renderTrendCell(week.roas, previous?.roas ?? null, false, (value) => (value === null ? "N/A" : formatMetricValue(value)))}</td>
                          <td className="px-3 py-3 text-right">{renderTrendCell(week.uv, previous?.uv ?? null, false, (value) => (value === null ? "N/A" : formatMetricValue(value)))}</td>
                          <td className="px-3 py-3 text-right">{renderTrendCell(week.ctr, previous?.ctr ?? null, false, (value) => (value === null ? "N/A" : formatMetricValue(value)))}</td>
                          <td className="px-3 py-3 text-right">{renderTrendCell(week.cvr, previous?.cvr ?? null, false, (value) => (value === null ? "N/A" : formatMetricValue(value)))}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </section>

            <section className="dashboard-section mb-7">
              <div className="mb-4">
                <p className="text-[11px] font-semibold uppercase tracking-[0.16em] text-slate-500">Performance</p>
                <h2 className="mt-1 text-xl font-semibold text-slate-900">Daily Performance</h2>
                <p className="mt-1 text-sm text-slate-600">ROAS, CPI and UV with daily installs</p>
              </div>

              <div className="chart-surface h-80 w-full rounded-xl border border-slate-200 bg-white p-3">
                <ResponsiveContainer width="100%" height="100%">
                  <ComposedChart data={dailyMetrics}>
                    <CartesianGrid stroke="#eee9f4" strokeDasharray="4 4" />
                    <XAxis dataKey="dateLabel" tick={{ fontSize: 11, fill: "#777487" }} tickLine={false} axisLine={{ stroke: "#e8e3ef" }} />
                    <YAxis
                      yAxisId="metrics"
                      tick={{ fontSize: 11, fill: "#777487" }}
                      tickLine={false}
                      axisLine={false}
                      domain={[(dataMin: number) => Math.min(dataMin, 0), (dataMax: number) => dataMax * 1.2]}
                    />
                    <YAxis
                      yAxisId="installs"
                      orientation="right"
                      tick={{ fontSize: 11, fill: "#777487" }}
                      tickLine={false}
                      axisLine={false}
                      domain={[(dataMin: number) => Math.max(0, dataMin * 0.8), (dataMax: number) => dataMax * 1.2]}
                    />
                    <Tooltip
                      labelFormatter={(label) => `Date: ${label}`}
                      formatter={(value: number | string | ReadonlyArray<number | string> | undefined, name: string | number | undefined) => {
                        const normalizedValue = Array.isArray(value)
                          ? Number(value[0] ?? 0)
                          : Number(value ?? 0);

                        if (String(name) === "Installs") {
                          return [new Intl.NumberFormat("en-US").format(Math.round(normalizedValue)), "Installs"];
                        }

                        return [normalizedValue.toFixed(2), String(name ?? "")];
                      }}
                    />
                    <Legend wrapperStyle={{ fontSize: 12, paddingTop: 8 }} />
                    <Line yAxisId="metrics" type="monotone" dataKey="roas" name="ROAS" stroke="#7c4dff" strokeWidth={2.75} strokeLinecap="round" dot={false} activeDot={{ r: 4 }} />
                    <Line yAxisId="metrics" type="monotone" dataKey="cpi" name="CPI" stroke="#c84dad" strokeWidth={2.5} strokeLinecap="round" dot={false} activeDot={{ r: 4 }} />
                    <Line yAxisId="metrics" type="monotone" dataKey="uv" name="UV" stroke="#5367c8" strokeWidth={2.5} strokeLinecap="round" dot={false} activeDot={{ r: 4 }} />
                    <Bar yAxisId="installs" dataKey="installs" name="Installs" fill="#cfc0ff" radius={[6, 6, 0, 0]} maxBarSize={18} />
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
            </section>

            <section className="dashboard-section driver-analysis-section">
              <div className="driver-analysis-header">
                <p className="text-[11px] font-semibold uppercase tracking-[0.16em] text-slate-500">Diagnostics</p>
                <h2 className="driver-analysis-title text-slate-900">Driver Analysis</h2>
                <p className="driver-analysis-description">Compare acquisition performance, identify weekly drivers, and review the selected week.</p>
              </div>

              <div className="dashboard-inset week-control-panel rounded-xl border border-slate-200 bg-slate-50">
                <div className="week-control-row flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                  <label htmlFor="week-select" className="font-medium text-slate-700">
                    Analyze Week
                  </label>
                  <select
                    id="week-select"
                    value={selectedWeekKey}
                    onChange={(event) => setSelectedWeekKey(event.target.value)}
                    className="min-w-[220px] rounded-lg border border-violet-200 bg-white px-3 py-2 text-sm text-slate-800 shadow-sm focus:border-violet-400 focus:outline-none"
                  >
                    {sortedWeeklyMetrics.map((week) => (
                      <option key={week.weekKey} value={week.weekKey}>
                        {week.weekLabel}
                      </option>
                    ))}
                  </select>
                </div>

                <div className="week-comparison-copy text-slate-600">
                  Compared with: {previousWeek ? previousWeek.weekLabel : "No previous week available"}
                </div>
              </div>

              {selectedWeek ? (
                <div className="selected-week-block">
                  <h3 className="selected-week-heading text-slate-800">Selected Week Performance</h3>
                  <section aria-label="Selected week key metrics" className="selected-kpi-grid grid grid-cols-2 lg:grid-cols-4">
                    <div className="metric-chip metric-chip-roas rounded-2xl border p-4">
                      <div className="text-xs font-medium text-slate-600">ROAS</div>
                      <div className="mt-2 text-2xl font-semibold text-slate-900">{formatMetricValue(selectedWeek.roas)}</div>
                      <div className="mt-1 text-xs text-slate-500">{selectedWeek.weekLabel}</div>
                    </div>
                    <div className="metric-chip metric-chip-cpi rounded-2xl border p-4">
                      <div className="text-xs font-medium text-slate-600">CPI</div>
                      <div className="mt-2 text-2xl font-semibold text-slate-900">
                        {selectedWeek.cpi === null ? "N/A" : formatCurrency(selectedWeek.cpi)}
                      </div>
                      <div className="mt-1 text-xs text-slate-500">{selectedWeek.weekLabel}</div>
                    </div>
                    <div className="metric-chip metric-chip-uv rounded-2xl border p-4">
                      <div className="text-xs font-medium text-slate-600">UV</div>
                      <div className="mt-2 text-2xl font-semibold text-slate-900">{formatMetricValue(selectedWeek.uv)}</div>
                      <div className="mt-1 text-xs text-slate-500">{selectedWeek.weekLabel}</div>
                    </div>
                    <div className="metric-chip metric-chip-cost rounded-2xl border p-4">
                      <div className="text-xs font-medium text-slate-600">Cost</div>
                      <div className="mt-2 text-2xl font-semibold text-slate-900">
                        {selectedWeek.cost === null ? "N/A" : formatCurrency(selectedWeek.cost)}
                      </div>
                      <div className="mt-1 text-xs text-slate-500">{selectedWeek.weekLabel}</div>
                    </div>
                  </section>
                </div>
              ) : null}

              {selectedWeek && previousWeek ? (
                <div className="driver-analysis-content">
                  <div className="dashboard-inset comparison-panel rounded-xl border border-slate-200 bg-slate-50">
                    <div className="comparison-heading">
                      <div className="comparison-title text-slate-800">
                        Weekly Performance Comparison
                      </div>
                      <div className="comparison-subtitle text-slate-600">Selected week vs. previous week</div>
                    </div>
                    <div className="chart-surface comparison-chart w-full rounded-xl border p-3">
                      <ResponsiveContainer width="100%" height="100%">
                        <LineChart data={weeklyComparisonData}>
                          <CartesianGrid stroke="#eee9f4" strokeDasharray="4 4" />
                          <XAxis dataKey="name" tick={{ fontSize: 12, fill: "#777487" }} tickLine={false} axisLine={{ stroke: "#e8e3ef" }} />
                          <YAxis tick={{ fontSize: 12, fill: "#777487" }} tickLine={false} axisLine={false} domain={[(dataMin: number) => Math.min(dataMin, 0), (dataMax: number) => dataMax * 1.2]} />
                          <Tooltip
                            labelFormatter={(label) => `Week: ${label}`}
                            formatter={(value: number | string | ReadonlyArray<number | string> | undefined, name: string | number | undefined) => {
                              const normalizedValue = Array.isArray(value)
                                ? Number(value[0] ?? 0)
                                : Number(value ?? 0);
                              return [normalizedValue.toFixed(2), String(name ?? "")];
                            }}
                          />
                          <Legend wrapperStyle={{ fontSize: 13, paddingTop: 10 }} />
                          <Line type="monotone" dataKey="roas" name="ROAS" stroke="#7c4dff" strokeWidth={2.75} strokeLinecap="round" dot={{ r: 3, fill: "#7c4dff", strokeWidth: 0 }} activeDot={{ r: 5 }} />
                          <Line type="monotone" dataKey="cpi" name="CPI" stroke="#c84dad" strokeWidth={2.5} strokeLinecap="round" dot={{ r: 3, fill: "#c84dad", strokeWidth: 0 }} activeDot={{ r: 5 }} />
                          <Line type="monotone" dataKey="uv" name="UV" stroke="#5367c8" strokeWidth={2.5} strokeLinecap="round" dot={{ r: 3, fill: "#5367c8", strokeWidth: 0 }} activeDot={{ r: 5 }} />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>
                  </div>

                  <div className="dashboard-inset conclusion-panel rounded-xl border border-slate-200 bg-slate-50">
                    <h3 className="conclusion-title text-slate-800">Conclusion</h3>
                    <div className={`conclusion-badge inline-flex items-center rounded-lg border font-semibold ${conclusionTone}`}>
                      {conclusionTitle}
                    </div>
                    <p className="conclusion-copy text-slate-700">
                      {renderImportantNumbers(conclusionSummary)}
                    </p>
                  </div>

                  <div className="analysis-panel rounded-xl border border-slate-200 bg-white">
                    <h3 className="analysis-title text-slate-800">Analysis</h3>
                    <div className="analysis-subsections">
                      {driverAnalysis?.analysis?.map((section) => (
                        <div key={section.heading} className="analysis-subsection rounded-xl border border-slate-200 bg-slate-50">
                          <h4 className="analysis-subsection-title text-violet-800">
                            {section.heading}
                          </h4>
                          <ul className="analysis-bullets pl-5 text-slate-700">
                            {section.bullets.map((bullet) => (
                              <li key={bullet} className="list-disc">
                                {renderMetricText(bullet)}
                              </li>
                            ))}
                          </ul>
                          {section.paragraphs.map((paragraph) => (
                            <p key={paragraph} className="analysis-paragraph text-slate-700">
                              {paragraph.includes("CPI") || paragraph.includes("UV") || paragraph.includes("CTR") || paragraph.includes("CVR") || paragraph.includes("ROAS")
                                ? renderParagraphText(paragraph)
                                : paragraph}
                            </p>
                          ))}
                        </div>
                      ))}
                    </div>
                  </div>

                </div>
              ) : (
                <div className="rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-800">
                  A previous week is required for comparison. Please select a week that has an immediately previous week available.
                </div>
              )}
            </section>

            <section className="recommendation-section mt-7 rounded-2xl border border-violet-200 bg-violet-50/70 shadow-sm">
              <div className="recommendation-header">
                <p className="text-xs font-semibold uppercase tracking-[0.16em] text-violet-700">Actions</p>
                <h2 className="recommendation-title text-slate-900">Recommendation</h2>
              </div>

              {selectedWeek && previousWeek && driverAnalysis ? (
                <div className="recommendation-content">
                  <div className="recommendation-card rounded-xl border border-violet-100 bg-white/80">
                    <div className="recommendation-subtitle text-violet-800">Executive Summary</div>
                    <p className="recommendation-copy text-slate-700">{driverAnalysis.recommendationSummary}</p>
                  </div>

                  <div className="recommendation-card rounded-xl border border-violet-100 bg-white/80">
                    <div className="recommendation-subtitle text-violet-800">Recommended Actions</div>
                    <ol className="recommendation-list pl-5 text-slate-700">
                      {driverAnalysis.recommendationActions.map((action, index) => {
                        const parsed = splitRecommendationAction(action);
                        return (
                          <li key={`${index}-${parsed.title}`} className="list-decimal pl-1">
                            <div>
                              <strong className="text-slate-900">{parsed.title}</strong>
                              {parsed.detail ? <span className="text-slate-700">. {parsed.detail}</span> : null}
                            </div>
                          </li>
                        );
                      })}
                    </ol>
                  </div>
                </div>
              ) : (
                <div className="rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-800">
                  Recommendation is unavailable because there is no previous week for comparison.
                </div>
              )}
            </section>
          </>
        ) : null}
      </div>
    </div>
  );
}
