import * as XLSX from "xlsx";

export const REQUIRED_COLUMNS = [
  "Date",
  "Campaign name",
  "Cost",
  "Impressions",
  "Clicks",
  "Installs",
  "Revenue",
  "Note",
] as const;

export type RequiredColumn = (typeof REQUIRED_COLUMNS)[number];

export type ImportedRow = {
  Date: unknown;
  "Campaign name": unknown;
  Cost: unknown;
  Impressions: unknown;
  Clicks: unknown;
  Installs: unknown;
  Revenue: unknown;
  Note: unknown;
};

export type WeeklyMetric = {
  weekKey: string;
  weekLabel: string;
  weekStart: Date;
  weekEnd: Date;
  cost: number;
  impressions: number;
  clicks: number;
  installs: number;
  revenue: number;
  roas: number | null;
  cpi: number | null;
  uv: number | null;
  ctr: number | null;
  cvr: number | null;
  noteLimited: boolean;
};

export type DailyMetric = {
  dateKey: string;
  date: string;
  dateLabel: string;
  cost: number;
  impressions: number;
  clicks: number;
  installs: number;
  revenue: number;
  roas: number | null;
  cpi: number | null;
  uv: number | null;
};

export function normalizeHeader(value: string): string {
  return value.trim().toLowerCase().replace(/\s+/g, " ");
}

export function sanitizeNumber(value: unknown): number {
  if (value === null || value === undefined || value === "") {
    return 0;
  }

  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  if (typeof value === "string") {
    const trimmedString = value.trim();
    if (!trimmedString) {
      return 0;
    }

    const sanitized = trimmedString
      .replace(/[$,%\s]/g, "")
      .replace("(", "-")
      .replace(")", "");

    const parsed = Number(sanitized);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  if (typeof value === "boolean") {
    return value ? 1 : 0;
  }

  return 0;
}

export function parseDateValue(value: unknown): Date | null {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : value;
  }

  if (typeof value === "number") {
    const date = XLSXDateFromNumber(value);
    return date;
  }

  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) {
      return null;
    }

    const directDate = new Date(trimmed);
    if (!Number.isNaN(directDate.getTime())) {
      return directDate;
    }

    const match = trimmed.match(/^\d{4}-\d{2}-\d{2}$/);
    if (match) {
      const [year, month, day] = trimmed.split("-").map(Number);
      const rebuilt = new Date(year, month - 1, day);
      if (!Number.isNaN(rebuilt.getTime())) {
        return rebuilt;
      }
    }
  }

  return null;
}

function XLSXDateFromNumber(value: number): Date | null {
  const excelDate = value as number;
  if (!Number.isFinite(excelDate)) {
    return null;
  }

  const hasDatePacking = excelDate > 10000 && excelDate < 100000000;
  if (hasDatePacking) {
    try {
      const parsed = (globalThis as typeof globalThis & { XLSX?: { SSF?: { parse_date_code?: (code: number) => { y: number; m: number; d: number } } } }).XLSX?.SSF?.parse_date_code;
      if (typeof parsed === "function") {
        const dateParts = parsed(excelDate);
        if (dateParts) {
          return new Date(dateParts.y, dateParts.m - 1, dateParts.d);
        }
      }
    } catch {
      // Ignore parse failures and fall back to JavaScript Date.
    }
  }

  const parsed = new Date(excelDate);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

export function parseUploadedRows(records: Array<Record<string, unknown>>): ImportedRow[] {
  if (!records || records.length === 0) {
    return [];
  }

  const normalizedRecords = records.map((row) => {
    const normalized: Record<string, unknown> = {};
    Object.entries(row).forEach(([key, value]) => {
      normalized[normalizeHeader(key)] = value;
    });
    return normalized;
  });

  const requiredHeaders = REQUIRED_COLUMNS.map((column) => normalizeHeader(column));
  const firstRecordKeys = Object.keys(normalizedRecords[0] ?? {});
  const missingColumns = requiredHeaders.filter(
    (header) => !firstRecordKeys.includes(header),
  );

  if (missingColumns.length > 0) {
    throw new Error(
      `Missing required columns: ${missingColumns.join(", ")}. Required columns are: ${REQUIRED_COLUMNS.join(", ")}.`,
    );
  }

  return normalizedRecords
    .filter((row) =>
      Object.values(row).some(
        (value) => value !== null && value !== undefined && String(value).trim() !== "",
      ),
    )
    .map((row) => {
      const mapped: Record<string, unknown> = {};
      REQUIRED_COLUMNS.forEach((column) => {
        mapped[column] = row[normalizeHeader(column)];
      });
      return mapped as ImportedRow;
    });
}

export function getWeekStart(date: Date): Date {
  const result = new Date(date);
  const day = result.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  result.setDate(result.getDate() + diff);
  result.setHours(0, 0, 0, 0);
  return result;
}

export function addDays(date: Date, amount: number): Date {
  const result = new Date(date);
  result.setDate(result.getDate() + amount);
  return result;
}

export function formatMonthDay(date: Date): string {
  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "numeric",
  }).format(date);
}

export function formatDateKey(date: Date): string {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(
    date.getDate(),
  ).padStart(2, "0")}`;
}

export function safeDivide(numerator: number, denominator: number): number | null {
  if (!Number.isFinite(numerator) || !Number.isFinite(denominator) || denominator === 0) {
    return null;
  }

  return numerator / denominator;
}

export function getPercentChange(current: number | null, previous: number | null): number | null {
  if (current === null || previous === null || previous === 0) {
    return null;
  }

  return ((current - previous) / previous) * 100;
}

export function formatRatio(value: number | null, digits = 2): string {
  if (value === null || !Number.isFinite(value)) {
    return "N/A";
  }

  return value.toFixed(digits);
}

export function formatPercent(value: number | null): string {
  if (value === null || !Number.isFinite(value)) {
    return "N/A";
  }

  return `${value >= 0 ? "+" : ""}${value.toFixed(1)}%`;
}

export function formatCurrency(value: number): string {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 2,
  }).format(value);
}

export function formatNumber(value: number | null): string {
  if (value === null || !Number.isFinite(value)) {
    return "N/A";
  }

  return new Intl.NumberFormat("en-US", {
    maximumFractionDigits: 2,
  }).format(value);
}

export function buildDailyMetrics(rows: ImportedRow[]): DailyMetric[] {
  const grouped = new Map<string, { date: Date; cost: number; impressions: number; clicks: number; installs: number; revenue: number }>();

  rows.forEach((row) => {
    const date = parseDateValue(row.Date);
    if (!date) {
      return;
    }

    const dateKey = formatDateKey(date);
    const current = grouped.get(dateKey) ?? {
      date,
      cost: 0,
      impressions: 0,
      clicks: 0,
      installs: 0,
      revenue: 0,
    };

    current.cost += sanitizeNumber(row.Cost);
    current.impressions += sanitizeNumber(row.Impressions);
    current.clicks += sanitizeNumber(row.Clicks);
    current.installs += sanitizeNumber(row.Installs);
    current.revenue += sanitizeNumber(row.Revenue);
    grouped.set(dateKey, current);
  });

  return [...grouped.values()]
    .sort((a, b) => a.date.getTime() - b.date.getTime())
    .map((entry) => ({
      dateKey: formatDateKey(entry.date),
      date: entry.date.toISOString(),
      dateLabel: new Intl.DateTimeFormat("en-US", {
        month: "short",
        day: "numeric",
      }).format(entry.date),
      cost: entry.cost,
      impressions: entry.impressions,
      clicks: entry.clicks,
      installs: entry.installs,
      revenue: entry.revenue,
      roas: safeDivide(entry.revenue, entry.cost),
      cpi: safeDivide(entry.cost, entry.installs),
      uv: safeDivide(entry.revenue, entry.installs),
    }));
}

export function buildWeeklyMetrics(rows: ImportedRow[]): WeeklyMetric[] {
  const grouped = new Map<string, { startDate: Date; rows: ImportedRow[] }>();

  rows.forEach((row) => {
    const date = parseDateValue(row.Date);
    if (!date) {
      return;
    }

    const weekStart = getWeekStart(date);
    const key = formatDateKey(weekStart);
    const current = grouped.get(key) ?? { startDate: weekStart, rows: [] };
    current.rows.push(row);
    grouped.set(key, current);
  });

  return [...grouped.values()]
    .map(({ startDate, rows }) => {
      const cost = rows.reduce((sum, row) => sum + sanitizeNumber(row.Cost), 0);
      const impressions = rows.reduce(
        (sum, row) => sum + sanitizeNumber(row.Impressions),
        0,
      );
      const clicks = rows.reduce((sum, row) => sum + sanitizeNumber(row.Clicks), 0);
      const installs = rows.reduce((sum, row) => sum + sanitizeNumber(row.Installs), 0);
      const revenue = rows.reduce((sum, row) => sum + sanitizeNumber(row.Revenue), 0);
      const endDate = addDays(startDate, 6);
      const totalRows = rows.length;

      return {
        weekKey: formatDateKey(startDate),
        weekLabel: `${formatMonthDay(startDate)}–${formatMonthDay(endDate)}`,
        weekStart: startDate,
        weekEnd: endDate,
        cost,
        impressions,
        clicks,
        installs,
        revenue,
        roas: safeDivide(revenue, cost),
        cpi: safeDivide(cost, installs),
        uv: safeDivide(revenue, installs),
        ctr: safeDivide(clicks, impressions),
        cvr: safeDivide(installs, clicks),
        noteLimited: rows.some((row) => {
          const value = String(row.Note ?? "").toLowerCase();
          return value.includes("limited by target");
        }),
        rowCount: totalRows,
      } as WeeklyMetric & { rowCount: number };
    })
    .sort((a, b) => b.weekStart.getTime() - a.weekStart.getTime())
    .map(({ rowCount, ...week }) => week);
}

export function buildDriverAnalysis(selectedWeek: WeeklyMetric, previousWeek: WeeklyMetric) {
  const roasChange = getPercentChange(selectedWeek.roas, previousWeek.roas);
  const cpiChange = getPercentChange(selectedWeek.cpi, previousWeek.cpi);
  const uvChange = getPercentChange(selectedWeek.uv, previousWeek.uv);
  const costChange = getPercentChange(selectedWeek.cost, previousWeek.cost);
  const impressionsChange = getPercentChange(selectedWeek.impressions, previousWeek.impressions);
  const clicksChange = getPercentChange(selectedWeek.clicks, previousWeek.clicks);
  const installsChange = getPercentChange(selectedWeek.installs, previousWeek.installs);
  const revenueChange = getPercentChange(selectedWeek.revenue, previousWeek.revenue);
  const ctrChange = getPercentChange(selectedWeek.ctr, previousWeek.ctr);
  const cvrChange = getPercentChange(selectedWeek.cvr, previousWeek.cvr);

  const delta =
    selectedWeek.roas !== null && previousWeek.roas !== null
      ? selectedWeek.roas - previousWeek.roas
      : null;
  const roasChangeText =
    roasChange === null ? "N/A" : `${roasChange >= 0 ? "+" : ""}${roasChange.toFixed(1)}%`;

  const cpiIncreased = selectedWeek.cpi !== null && previousWeek.cpi !== null && selectedWeek.cpi > previousWeek.cpi;
  const cpiDecreased = selectedWeek.cpi !== null && previousWeek.cpi !== null && selectedWeek.cpi < previousWeek.cpi;
  const spendIncreased = selectedWeek.cost > previousWeek.cost;
  const ctrDecreased = selectedWeek.ctr !== null && previousWeek.ctr !== null && selectedWeek.ctr < previousWeek.ctr;
  const cvrDecreased = selectedWeek.cvr !== null && previousWeek.cvr !== null && selectedWeek.cvr < previousWeek.cvr;
  const uvDecreased = selectedWeek.uv !== null && previousWeek.uv !== null && selectedWeek.uv < previousWeek.uv;
  const uvIncreased = selectedWeek.uv !== null && previousWeek.uv !== null && selectedWeek.uv > previousWeek.uv;
  const roasIncreased = selectedWeek.roas !== null && previousWeek.roas !== null && selectedWeek.roas > previousWeek.roas;
  const roasDecreased = selectedWeek.roas !== null && previousWeek.roas !== null && selectedWeek.roas < previousWeek.roas;

  const currentROASAboveOne = selectedWeek.roas !== null && selectedWeek.roas >= 1;
  const currentROASBelowOne = selectedWeek.roas !== null && selectedWeek.roas < 1;

  const buildDriverDetails = () => {
    let driverTitle = "Mixed driver signal";
    let driverSummary = `ROAS decreased from ${formatRatio(previousWeek.roas)} to ${formatRatio(selectedWeek.roas)} (${roasChangeText}), but the available indicators do not clearly match a single predefined driver.`;
    let driverRecommendation = "The available signals are mixed and further investigation is needed.";

    if (cpiIncreased && spendIncreased && ctrDecreased && cvrDecreased) {
      driverTitle = "Higher CPI combined with weaker traffic efficiency / scaling pressure";
      driverSummary = `CPI increased from ${formatRatio(previousWeek.cpi)} to ${formatRatio(selectedWeek.cpi)} (${cpiChange === null ? "N/A" : `${cpiChange >= 0 ? "+" : ""}${cpiChange.toFixed(1)}%`}), spend increased from ${formatCurrency(previousWeek.cost)} to ${formatCurrency(selectedWeek.cost)} (${costChange === null ? "N/A" : `${costChange >= 0 ? "+" : ""}${costChange.toFixed(1)}%`}), CTR declined from ${formatRatio(previousWeek.ctr)} to ${formatRatio(selectedWeek.ctr)} (${ctrChange === null ? "N/A" : `${ctrChange >= 0 ? "+" : ""}${ctrChange.toFixed(1)}%`}), and CVR declined from ${formatRatio(previousWeek.cvr)} to ${formatRatio(selectedWeek.cvr)} (${cvrChange === null ? "N/A" : `${cvrChange >= 0 ? "+" : ""}${cvrChange.toFixed(1)}%`}).`;
      driverRecommendation = "Adjust the budget OR launch a new campaign/ad group to refresh learning.";
    } else if (cpiIncreased && !(spendIncreased && ctrDecreased && cvrDecreased)) {
      driverTitle = "Possible creative fatigue";
      driverSummary = `CPI increased from ${formatRatio(previousWeek.cpi)} to ${formatRatio(selectedWeek.cpi)} (${cpiChange === null ? "N/A" : `${cpiChange >= 0 ? "+" : ""}${cpiChange.toFixed(1)}%`}), while other indicators were relatively stable. This points to possible creative fatigue.`;
      driverRecommendation = "Refresh or change creatives.";
    } else if (uvDecreased) {
      if (!selectedWeek.noteLimited) {
        driverTitle = "Lower user value";
        driverSummary = `UV decreased from ${formatRatio(previousWeek.uv)} to ${formatRatio(selectedWeek.uv)} (${uvChange === null ? "N/A" : `${uvChange >= 0 ? "+" : ""}${uvChange.toFixed(1)}%`}), which suggests the campaign is attracting lower-value users.`;
        driverRecommendation = "Consider turning off underperforming ad groups or replacing creatives.";
      } else {
        driverTitle = "Lower UV under Limited by target conditions";
        driverSummary = `UV decreased from ${formatRatio(previousWeek.uv)} to ${formatRatio(selectedWeek.uv)} (${uvChange === null ? "N/A" : `${uvChange >= 0 ? "+" : ""}${uvChange.toFixed(1)}%`}), and the selected week is marked as 'Limited by target'.`;
        driverRecommendation = "Increase the tROAS target.";
      }
    }

    return { driverTitle, driverSummary, driverRecommendation };
  };

  if (selectedWeek.roas !== null && previousWeek.roas !== null) {
    if (currentROASAboveOne && roasIncreased) {
      const positiveDriverText =
        cpiDecreased && uvIncreased
          ? "The improvement is supported by both lower CPI and higher UV."
          : cpiDecreased
            ? "The improvement is supported by lower CPI."
            : uvIncreased
              ? "The improvement is supported by higher UV."
              : "The underlying efficiency and value metrics supported the ROAS improvement.";

      return {
        title: "Good Performance",
        summary: `ROAS increased from ${formatRatio(previousWeek.roas)} to ${formatRatio(selectedWeek.roas)} (${roasChangeText}) and remains above the 1.00 profitability benchmark. This indicates both profitable performance and positive week-over-week momentum. ${positiveDriverText}`,
        conclusion: `Good Performance: ROAS increased from ${formatRatio(previousWeek.roas)} to ${formatRatio(selectedWeek.roas)} (${roasChangeText}) and remains above the 1.00 profitability benchmark. This indicates both profitable performance and positive week-over-week momentum. ${positiveDriverText}`,
        recommendation: "Maintain the current approach while monitoring CPI, UV and ROAS to confirm the improvement is sustainable.",
        analysis: [
          {
            heading: "ROAS PERFORMANCE",
            bullets: [
              `Current ROAS: ${formatRatio(selectedWeek.roas)}`,
              `Previous ROAS: ${formatRatio(previousWeek.roas)}`,
              `Absolute change: ${delta === null ? "N/A" : `${delta >= 0 ? "+" : ""}${delta.toFixed(2)}`}`,
              `WoW change: ${roasChangeText}`,
              `Current status: Profitable and improving`,
            ],
            paragraphs: [
              `ROAS increased from ${formatRatio(previousWeek.roas)} to ${formatRatio(selectedWeek.roas)} and remains above 1.00, which indicates the campaign is currently profitable and improving week over week. The driver analysis should focus on whether that improvement is being supported by better acquisition efficiency, higher user value, or both.`,
            ],
          },
        ],
        recommendationSummary: `ROAS increased and remains above 1.00, which is a strong positive signal. The campaign is profitable and moving in the right direction, so the recommended approach is to maintain the winning setup while watching CPI and UV for sustainability.`,
        recommendationActions: [
          `1. Maintain the current strategy and keep monitoring ROAS, CPI and UV to confirm that the positive momentum remains sustainable.`,
          `2. Build on the winning traffic pattern, but continue checking whether the improvement is being driven by lower CPI, higher UV, or both.`,
        ],
      };
    }

    if (currentROASAboveOne && roasDecreased) {
      const driverDetails = buildDriverDetails();
      return {
        title: "Profitable but Below Previous Week",
        summary: `ROAS decreased from ${formatRatio(previousWeek.roas)} to ${formatRatio(selectedWeek.roas)} (${roasChangeText}), although current ROAS remains above the 1.00 profitability benchmark. The campaign is still profitable, but performance is weaker than the previous week. ${driverDetails.driverSummary}`,
        conclusion: `Profitable but Below Previous Week: ROAS decreased from ${formatRatio(previousWeek.roas)} to ${formatRatio(selectedWeek.roas)} (${roasChangeText}), although current ROAS remains above the 1.00 profitability benchmark. The campaign is still profitable, but performance is weaker than the previous week. ${driverDetails.driverSummary}`,
        recommendation: driverDetails.driverRecommendation,
        analysis: [
          {
            heading: "ROAS PERFORMANCE",
            bullets: [
              `Current ROAS: ${formatRatio(selectedWeek.roas)}`,
              `Previous ROAS: ${formatRatio(previousWeek.roas)}`,
              `Absolute change: ${delta === null ? "N/A" : `${delta >= 0 ? "+" : ""}${delta.toFixed(2)}`}`,
              `WoW change: ${roasChangeText}`,
              `Current status: Profitable but below previous week`,
            ],
            paragraphs: [
              `ROAS decreased from ${formatRatio(previousWeek.roas)} to ${formatRatio(selectedWeek.roas)} even though the current level remains above 1.00. This means the campaign is still generating profit, but it is operating below the prior week’s efficiency and the driver analysis should focus on the reason for the weaker return.`,
            ],
          },
          {
            heading: "DRIVER DIAGNOSIS",
            bullets: [
              `ROAS: ${formatRatio(previousWeek.roas)} → ${formatRatio(selectedWeek.roas)} (${roasChangeText})`,
              `CPI: ${formatRatio(previousWeek.cpi)} → ${formatRatio(selectedWeek.cpi)} (${cpiChange === null ? "N/A" : `${cpiChange >= 0 ? "+" : ""}${cpiChange.toFixed(1)}%`})`,
              `UV: ${formatRatio(previousWeek.uv)} → ${formatRatio(selectedWeek.uv)} (${uvChange === null ? "N/A" : `${uvChange >= 0 ? "+" : ""}${uvChange.toFixed(1)}%`})`,
              `Campaign status: ${selectedWeek.noteLimited ? "Limited by target" : "Not marked as Limited by target"}`,
            ],
            paragraphs: [
              `${driverDetails.driverSummary} This week remains profitable because ROAS is still above 1.00, but the revenue efficiency has weakened relative to the previous week. The fact that the campaign remains above 1.00 means the issue is not a profitability crisis, but a deterioration in efficiency or user value compared with the prior week.`,
            ],
          },
        ],
        recommendationSummary: `The campaign remains profitable because current ROAS is still above 1.00, but the selected week underperformed the previous week. The change should be evaluated through the existing CPI and UV driver rules before making a corrective move.`,
        recommendationActions: [
          `1. ${driverDetails.driverRecommendation} This response follows the defined driver rules for the selected week’s ROAS decline.`,
          `2. Continue monitoring ROAS, CPI and UV to confirm whether the downshift is temporary or structural.`,
        ],
      };
    }

    if (currentROASBelowOne && roasDecreased) {
      const driverDetails = buildDriverDetails();
      return {
        title: "Low Performance",
        summary: `ROAS decreased from ${formatRatio(previousWeek.roas)} to ${formatRatio(selectedWeek.roas)} (${roasChangeText}) and remains below the 1.00 profitability benchmark. The campaign is therefore in Low Performance, with the decline driven by ${driverDetails.driverTitle.toLowerCase()}. ${driverDetails.driverSummary}`,
        conclusion: `Low Performance: ROAS decreased from ${formatRatio(previousWeek.roas)} to ${formatRatio(selectedWeek.roas)} (${roasChangeText}) and remains below the 1.00 profitability benchmark. The campaign is therefore in Low Performance, with the decline driven by ${driverDetails.driverTitle.toLowerCase()}. ${driverDetails.driverSummary}`,
        recommendation: driverDetails.driverRecommendation,
        analysis: [
          {
            heading: "ROAS PERFORMANCE",
            bullets: [
              `Current ROAS: ${formatRatio(selectedWeek.roas)}`,
              `Previous ROAS: ${formatRatio(previousWeek.roas)}`,
              `Absolute change: ${delta === null ? "N/A" : `${delta >= 0 ? "+" : ""}${delta.toFixed(2)}`}`,
              `WoW change: ${roasChangeText}`,
              `Current status: Low Performance (<1.00)`,
            ],
            paragraphs: [
              `ROAS decreased from ${formatRatio(previousWeek.roas)} to ${formatRatio(selectedWeek.roas)} and remains below the 1.00 profitability benchmark. The selected week therefore satisfies the true negative case and requires a driver-based response.`,
            ],
          },
          {
            heading: "DRIVER DIAGNOSIS",
            bullets: [
              `Primary driver: ${driverDetails.driverTitle}`,
              `CPI: ${formatRatio(previousWeek.cpi)} → ${formatRatio(selectedWeek.cpi)} (${cpiChange === null ? "N/A" : `${cpiChange >= 0 ? "+" : ""}${cpiChange.toFixed(1)}%`})`,
              `UV: ${formatRatio(previousWeek.uv)} → ${formatRatio(selectedWeek.uv)} (${uvChange === null ? "N/A" : `${uvChange >= 0 ? "+" : ""}${uvChange.toFixed(1)}%`})`,
              `CTR: ${formatRatio(previousWeek.ctr)} → ${formatRatio(selectedWeek.ctr)} (${ctrChange === null ? "N/A" : `${ctrChange >= 0 ? "+" : ""}${ctrChange.toFixed(1)}%`})`,
              `CVR: ${formatRatio(previousWeek.cvr)} → ${formatRatio(selectedWeek.cvr)} (${cvrChange === null ? "N/A" : `${cvrChange >= 0 ? "+" : ""}${cvrChange.toFixed(1)}%`})`,
            ],
            paragraphs: [
              `${driverDetails.driverSummary} This negative ROAS trend is consistent with the predefined low-performance rule, and the selected week should be interpreted as a material efficiency problem rather than a temporary fluctuation.`,
            ],
          },
        ],
        recommendationSummary: `${driverDetails.driverSummary} ${driverDetails.driverRecommendation}`,
        recommendationActions: [
          `1. ${driverDetails.driverRecommendation}`,
          `2. Continue monitoring ROAS, CPI, UV, CTR and CVR to confirm whether the negative trajectory is reversed after the corrective action.`,
        ],
      };
    }

    if (currentROASBelowOne && roasIncreased) {
      const positiveDriverText =
        cpiDecreased && uvIncreased
          ? "The improvement is supported by both lower CPI and higher UV."
          : cpiDecreased
            ? "The improvement is supported by lower CPI."
            : uvIncreased
              ? "The improvement is supported by higher UV."
              : "The underlying performance mix improved enough to lift ROAS week over week.";

      return {
        title: "Good Signal — Not Yet Profitable",
        summary: `ROAS increased from ${formatRatio(previousWeek.roas)} to ${formatRatio(selectedWeek.roas)} (${roasChangeText}), indicating a positive week-over-week improvement. However, current ROAS remains below the 1.00 profitability benchmark, so the campaign is not yet profitable. ${positiveDriverText}`,
        conclusion: `Good Signal — Not Yet Profitable: ROAS increased from ${formatRatio(previousWeek.roas)} to ${formatRatio(selectedWeek.roas)} (${roasChangeText}), indicating a positive week-over-week improvement. However, current ROAS remains below the 1.00 profitability benchmark, so the campaign is not yet profitable. ${positiveDriverText}`,
        recommendation: "Scale Winning Creatives",
        analysis: [
          {
            heading: "ROAS PERFORMANCE",
            bullets: [
              `Current ROAS: ${formatRatio(selectedWeek.roas)}`,
              `Previous ROAS: ${formatRatio(previousWeek.roas)}`,
              `Absolute change: ${delta === null ? "N/A" : `${delta >= 0 ? "+" : ""}${delta.toFixed(2)}`}`,
              `WoW change: ${roasChangeText}`,
              `Current status: Positive movement but still below profitability`,
            ],
            paragraphs: [
              `ROAS increased from ${formatRatio(previousWeek.roas)} to ${formatRatio(selectedWeek.roas)}, which is a positive directional signal. Even so, the campaign remains below the 1.00 benchmark, so it is not yet profitable. The improvement indicates that the underlying setup is moving in the right direction, but it still requires efficient scaling.`,
            ],
          },
          {
            heading: "DRIVER DIAGNOSIS",
            bullets: [
              `ROAS: ${formatRatio(previousWeek.roas)} → ${formatRatio(selectedWeek.roas)} (${roasChangeText})`,
              `CPI: ${formatRatio(previousWeek.cpi)} → ${formatRatio(selectedWeek.cpi)} (${cpiChange === null ? "N/A" : `${cpiChange >= 0 ? "+" : ""}${cpiChange.toFixed(1)}%`})`,
              `UV: ${formatRatio(previousWeek.uv)} → ${formatRatio(selectedWeek.uv)} (${uvChange === null ? "N/A" : `${uvChange >= 0 ? "+" : ""}${uvChange.toFixed(1)}%`})`,
            ],
            paragraphs: [
              `${positiveDriverText} This is not a negative ROAS case because the direction is positive; it is a sub-profitability growth case in which the selected week shows encouraging improvement but still requires additional efficiency to reach and sustain the 1.00 benchmark.`,
            ],
          },
        ],
        recommendationSummary: `ROAS improved week over week, but the campaign is still below the 1.00 profitability benchmark. The right action is to scale the winning creative set carefully rather than scaling the entire campaign indiscriminately.`,
        recommendationActions: [
          `1. Identify the creatives or ad groups contributing to the ROAS improvement and base the scale decision on those winners.`,
          `2. Gradually increase spend on the winning creatives rather than scaling the entire campaign indiscriminately.`,
          `3. Monitor CPI closely to ensure additional spend does not cause acquisition costs to deteriorate.`,
          `4. Monitor UV to confirm that higher volume is not reducing user quality.`,
          `5. Continue tracking ROAS until it reaches and sustains the 1.00 profitability benchmark.`,
        ],
      };
    }
  }

  return {
    title: selectedWeek.roas !== null && selectedWeek.roas >= 1 ? "Good Performance" : "Low Performance",
    summary: selectedWeek.roas !== null
      ? `Current ROAS is ${formatRatio(selectedWeek.roas)}. No previous week was available for comparison, so WoW direction could not be assessed.`
      : "Current ROAS is unavailable.",
    conclusion: selectedWeek.roas !== null && selectedWeek.roas >= 1
      ? `Good Performance: Current ROAS is ${formatRatio(selectedWeek.roas)} and remains above the 1.00 profitability benchmark. There is no previous week available for comparison, so week-over-week movement cannot be assessed.`
      : `Low Performance: Current ROAS is ${formatRatio(selectedWeek.roas)} and remains below the 1.00 profitability benchmark. There is no previous week available for comparison, so week-over-week movement cannot be assessed.`,
    recommendation: selectedWeek.roas !== null && selectedWeek.roas >= 1
      ? "Maintain the current approach while monitoring performance for the next period."
      : "Monitor the campaign and review the current traffic and creative mix while no previous week is available for comparison.",
    analysis: [],
    recommendationSummary: selectedWeek.roas !== null && selectedWeek.roas >= 1
      ? "Current ROAS is above 1.00, but there is no previous week available for comparison, so weekly direction cannot be assessed."
      : "Current ROAS is below 1.00, but there is no previous week available for comparison, so weekly direction cannot be assessed.",
    recommendationActions: [
      "No previous week is available for comparison, so driver analysis based on week-over-week movement is unavailable.",
    ],
  };
}
