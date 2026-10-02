import * as XLSX from "xlsx";

export const REQUIRED_COLUMNS = [
  "Date",
  "Campaign Name",
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
  "Campaign Name": unknown;
  Cost: unknown;
  Impressions: unknown;
  Clicks: unknown;
  Installs: unknown;
  Revenue: unknown;
  Note: unknown;
};

export type WeeklyMetric = {
  campaignName: string;
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
  campaignName: string;
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
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[_-]+/g, " ")
    .replace(/[^a-z0-9 ]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

export function parseBudgetValue(value: unknown): number | null {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  const text = String(value).trim();
  if (!text) {
    return null;
  }

  const normalized = text
    .replace(/[$,\s]/g, "")
    .replace(/%/g, "")
    .replace(/\(\s*$/, "-")
    .replace(/\)$/, "");

  const parsed = Number(normalized);
  if (!Number.isFinite(parsed)) {
    return null;
  }

  return parsed;
}

export function readTotalBudgetFromSheet(sheet: XLSX.WorkSheet): number | null {
  if (!sheet) {
    return null;
  }

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false }) as unknown[][];
  if (!rows || rows.length === 0) {
    return null;
  }

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex];
    if (!Array.isArray(row)) {
      continue;
    }

    for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
      const normalized = normalizeHeader(String(row[columnIndex] ?? ""));
      if (normalized !== "total budget" && normalized !== "budget") {
        continue;
      }

      const rightValue = parseBudgetValue(row[columnIndex + 1]);
      if (rightValue !== null && rightValue > 0) {
        return rightValue;
      }

      const belowRow = rows[rowIndex + 1];
      const belowValue = Array.isArray(belowRow) ? parseBudgetValue(belowRow[columnIndex]) : null;
      if (belowValue !== null && belowValue > 0) {
        return belowValue;
      }
    }
  }

  return null;
}

export function findRawDataSheet(workbook: XLSX.WorkBook): { sheetName: string; rows: Array<Record<string, unknown>> } | null {
  const requiredHeaders = REQUIRED_COLUMNS.map((column) => normalizeHeader(column));
  const sheetNames = workbook.SheetNames ?? [];

  if (sheetNames.length !== 1 || sheetNames[0].trim() !== "Raw Data") {
    return null;
  }

  const worksheet = workbook.Sheets?.[sheetNames[0]];
  if (!worksheet) {
    return null;
  }

  const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", raw: false }) as unknown[][];
  const headerRowIndex = rawRows.findIndex((row) => {
    if (!Array.isArray(row)) {
      return false;
    }
    const normalizedHeaders = row.map((cell) => normalizeHeader(String(cell ?? "")));
    return requiredHeaders.every((header) => normalizedHeaders.includes(header));
  });

  if (headerRowIndex === -1) {
    return null;
  }

  const headerRow = rawRows[headerRowIndex] ?? [];
  const rows = rawRows
    .slice(headerRowIndex + 1)
    .filter((dataRow) => Array.isArray(dataRow) && dataRow.some((cell) => String(cell ?? "").trim() !== ""))
    .map((dataRow) => {
      const mapped: Record<string, unknown> = {};
      headerRow.forEach((headerCell, columnIndex) => {
        const key = normalizeHeader(String(headerCell ?? ""));
        if (!key) {
          return;
        }
        mapped[key] = Array.isArray(dataRow) ? dataRow[columnIndex] : "";
      });
      return mapped;
    })
    .filter((row) => {
      const campaignName = String(row["campaign name"] ?? "").trim();
      const hasBudgetMarker = Object.values(row).some((value) => {
        const text = String(value ?? "").trim();
        return normalizeHeader(text).includes("total budget");
      });
      const hasAnyData = Object.values(row).some((value) => String(value ?? "").trim() !== "");
      return hasAnyData && !hasBudgetMarker && campaignName !== "";
    });

  return { sheetName: sheetNames[0], rows };
}

export function sanitizeNumber(value: unknown): number {
  if (value === null || value === undefined || value === "") {
    return 0;
  }

  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) {
      return 0;
    }

    const sanitized = trimmed
      .replace(/[$,%\s]/g, "")
      .replace(/\(/g, "-")
      .replace(/\)/g, "");

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
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) {
      return null;
    }

    const direct = new Date(trimmed);
    if (!Number.isNaN(direct.getTime())) {
      return direct;
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
  const keys = new Set(normalizedRecords.flatMap((record) => Object.keys(record)));
  const missingColumns = requiredHeaders.filter((header) => !keys.has(header));

  if (missingColumns.length > 0) {
    throw new Error(
      `Missing required columns: ${missingColumns.join(", ")}. Required columns are: ${REQUIRED_COLUMNS.join(", ")}.`,
    );
  }

  const parsed: ImportedRow[] = [];

  normalizedRecords.forEach((row, index) => {
    const mapped: Record<string, unknown> = {};
    REQUIRED_COLUMNS.forEach((column) => {
      mapped[column] = row[normalizeHeader(column)];
    });

    const hasData = Object.values(mapped).some(
      (value) => value !== null && value !== undefined && String(value).trim() !== "",
    );

    if (!hasData) {
      return;
    }

    if (String(mapped["Campaign Name"] ?? "").trim() === "") {
      throw new Error(`Campaign Name is missing in row ${index + 2}.`);
    }

    if (parseDateValue(mapped.Date) === null) {
      throw new Error(`Date is missing or invalid in row ${index + 2}.`);
    }

    const numericFields = [
      ["Cost", mapped.Cost],
      ["Impressions", mapped.Impressions],
      ["Clicks", mapped.Clicks],
      ["Installs", mapped.Installs],
      ["Revenue", mapped.Revenue],
    ] as const;

    for (const [fieldName, value] of numericFields) {
      const text = value === null || value === undefined ? "" : String(value).trim();
      if (!text) {
        throw new Error(`${fieldName} is missing in row ${index + 2}.`);
      }

      const parsedValue = parseBudgetValue(value);
      if (parsedValue === null) {
        throw new Error(`${fieldName} is invalid in row ${index + 2}.`);
      }
    }

    parsed.push(mapped as ImportedRow);
  });

  return parsed;
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
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
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

export function formatCurrency(value: number | null): string {
  if (value === null || !Number.isFinite(value)) {
    return "N/A";
  }

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

export function extractTotalBudget(records: Array<Record<string, unknown>>): number | null {
  if (!records || records.length === 0) {
    return null;
  }

  for (const row of records) {
    for (const [key, value] of Object.entries(row)) {
      const normalizedKey = normalizeHeader(key);
      if (normalizedKey === "total budget" || normalizedKey === "budget") {
        const parsed = parseTotalBudgetValue(value);
        if (parsed !== null) {
          return parsed;
        }
      }
    }
  }

  return null;
}

export function extractTotalBudgetFromWorkbook(workbook: XLSX.WorkBook): number | null {
  if (!workbook || !workbook.Sheets) {
    return null;
  }

  const sheetName = workbook.SheetNames?.[0];
  if (!sheetName || sheetName.trim() !== "Raw Data") {
    return null;
  }

  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    return null;
  }

  return readTotalBudgetFromSheet(worksheet);
}

function parseTotalBudgetValue(value: unknown): number | null {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  const text = String(value).trim();
  if (!text) {
    return null;
  }

  const match = text.match(/\$?\s*([0-9][0-9,]*(?:\.\d+)?)/);
  const candidate = match ? match[1] : text;
  const parsed = Number(candidate.replace(/[$,%\s]/g, ""));

  if (!Number.isFinite(parsed) || parsed <= 0) {
    return null;
  }

  return parsed;
}

export function getCampaignName(row: ImportedRow): string {
  return String(row["Campaign Name"] ?? "").trim();
}

export function buildDailyMetrics(rows: ImportedRow[]): DailyMetric[] {
  const byCampaign = new Map<string, Map<string, { date: Date; cost: number; impressions: number; clicks: number; installs: number; revenue: number }>>();

  rows.forEach((row) => {
    const campaignName = getCampaignName(row);
    if (!campaignName) {
      return;
    }

    const date = parseDateValue(row.Date);
    if (!date) {
      return;
    }

    const dateKey = formatDateKey(date);
    const campaignMap = byCampaign.get(campaignName) ?? new Map();
    const current = campaignMap.get(dateKey) ?? { date, cost: 0, impressions: 0, clicks: 0, installs: 0, revenue: 0 };

    current.cost += sanitizeNumber(row.Cost);
    current.impressions += sanitizeNumber(row.Impressions);
    current.clicks += sanitizeNumber(row.Clicks);
    current.installs += sanitizeNumber(row.Installs);
    current.revenue += sanitizeNumber(row.Revenue);
    campaignMap.set(dateKey, current);
    byCampaign.set(campaignName, campaignMap);
  });

  const output: DailyMetric[] = [];
  byCampaign.forEach((campaignMap, campaignName) => {
    [...campaignMap.values()]
      .sort((a, b) => a.date.getTime() - b.date.getTime())
      .forEach((entry) => {
        output.push({
          campaignName,
          dateKey: formatDateKey(entry.date),
          date: entry.date.toISOString(),
          dateLabel: new Intl.DateTimeFormat("en-US", { month: "short", day: "numeric" }).format(entry.date),
          cost: entry.cost,
          impressions: entry.impressions,
          clicks: entry.clicks,
          installs: entry.installs,
          revenue: entry.revenue,
          roas: safeDivide(entry.revenue, entry.cost),
          cpi: safeDivide(entry.cost, entry.installs),
          uv: safeDivide(entry.revenue, entry.installs),
        });
      });
  });

  return output;
}

export function buildWeeklyMetrics(rows: ImportedRow[]): WeeklyMetric[] {
  const byCampaign = new Map<string, Map<string, ImportedRow[]>>();

  rows.forEach((row) => {
    const campaignName = getCampaignName(row);
    if (!campaignName) {
      return;
    }

    const date = parseDateValue(row.Date);
    if (!date) {
      return;
    }

    const weekStart = getWeekStart(date);
    const weekKey = formatDateKey(weekStart);
    const campaignMap = byCampaign.get(campaignName) ?? new Map();
    const current = campaignMap.get(weekKey) ?? [];
    current.push(row);
    campaignMap.set(weekKey, current);
    byCampaign.set(campaignName, campaignMap);
  });

  const output: WeeklyMetric[] = [];
  byCampaign.forEach((campaignMap, campaignName) => {
    campaignMap.forEach((campaignRows, weekKey) => {
      const firstDate = parseDateValue(campaignRows[0]?.Date);
      if (!firstDate) {
        return;
      }

      const weekStart = getWeekStart(firstDate);
      const weekEnd = addDays(weekStart, 6);
      const cost = campaignRows.reduce((sum, row) => sum + sanitizeNumber(row.Cost), 0);
      const impressions = campaignRows.reduce((sum, row) => sum + sanitizeNumber(row.Impressions), 0);
      const clicks = campaignRows.reduce((sum, row) => sum + sanitizeNumber(row.Clicks), 0);
      const installs = campaignRows.reduce((sum, row) => sum + sanitizeNumber(row.Installs), 0);
      const revenue = campaignRows.reduce((sum, row) => sum + sanitizeNumber(row.Revenue), 0);

      output.push({
        campaignName,
        weekKey,
        weekLabel: `${formatMonthDay(weekStart)}–${formatMonthDay(weekEnd)}`,
        weekStart,
        weekEnd,
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
        noteLimited: campaignRows.some((row) => String(row.Note ?? "").toLowerCase().includes("limited by target")),
      });
    });
  });

  return output.sort((a, b) => b.weekStart.getTime() - a.weekStart.getTime());
}

export function buildPortfolioStats(rows: ImportedRow[], totalBudget: number | null) {
  const spendByCampaign = new Map<string, number>();
  rows.forEach((row) => {
    const name = getCampaignName(row);
    if (!name) {
      return;
    }

    spendByCampaign.set(name, (spendByCampaign.get(name) ?? 0) + sanitizeNumber(row.Cost));
  });

  const campaigns = [...spendByCampaign.entries()].sort((a, b) => b[1] - a[1]);
  const totalSpend = campaigns.reduce((sum, [, value]) => sum + value, 0);
  const remainingBudget = totalBudget !== null ? Math.max(totalBudget - totalSpend, 0) : 0;
  const spendRate = totalBudget && totalBudget > 0 ? totalSpend / totalBudget : null;

  return { campaigns, totalSpend, remainingBudget, spendRate };
}

export function getPortofolioCampaigns(rows: ImportedRow[]) {
  return [...new Set(rows.map((row) => getCampaignName(row)).filter(Boolean))].sort();
}
