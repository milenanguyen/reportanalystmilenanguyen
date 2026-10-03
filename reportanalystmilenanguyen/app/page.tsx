"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { Bar, BarChart, CartesianGrid, ComposedChart, Legend, Line, LineChart, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import {
  buildDailyMetrics,
  buildWeeklyMetrics,
  extractTotalBudget,
  extractTotalBudgetFromWorkbook,
  findRawDataSheet,
  formatCurrency,
  formatRatio,
  getPercentChange,
  parseUploadedRows,
  readTotalBudgetFromSheet,
  sanitizeNumber,
  type DailyMetric,
  type ImportedRow,
  type WeeklyMetric,
} from "./dashboard-utils";

function formatMetricValue(value: number | null, fallback = "N/A") {
  return value === null || !Number.isFinite(value) ? fallback : value.toFixed(2);
}

function formatNumber(value: number | null, digits = 0) {
  if (value === null || !Number.isFinite(value)) {
    return "N/A";
  }

  return new Intl.NumberFormat("en-US", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  }).format(value);
}

function safePercentChange(current: number | null, previous: number | null) {
  if (current === null || previous === null || !Number.isFinite(current) || !Number.isFinite(previous) || previous === 0) {
    return null;
  }

  return ((current - previous) / previous) * 100;
}

function formatSignedPercent(value: number | null) {
  if (value === null || !Number.isFinite(value)) {
    return "N/A";
  }

  return `${value >= 0 ? "+" : ""}${value.toFixed(1)}%`;
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

  return { title: match[1] ?? action, detail: match[2] ?? "" };
}

function renderMetricText(text: string) {
  const match = text.match(/^(.+?):\s*(.+)$/);
  if (match && /^(Current ROAS|Previous ROAS|Current status|Primary driver|CPI|UV|CTR|CVR|ROAS)$/.test(match[1])) {
    return (
      <>
        <span className="font-medium text-slate-700">{match[1]}:</span> {renderImportantNumbers(match[2])}
      </>
    );
  }

  return <>{text}</>;
}

function renderParagraphText(text: string) {
  const parts = text.split(/(CPI|UV|CTR|CVR|ROAS)/g);
  return (
    <>
      {parts.map((part, index) =>
        /^(CPI|UV|CTR|CVR|ROAS)$/.test(part) ? (
          <strong key={`${part}-${index}`} className="text-slate-900">{part}</strong>
        ) : (
          <span key={`${part}-${index}`}>{renderImportantNumbers(part)}</span>
        ),
      )}
    </>
  );
}

function getCampaignDisplayName(name: string) {
  return name?.trim() || "Unknown Campaign";
}

function compareNumericChange(current: number | null, previous: number | null) {
  if (current === null || previous === null || !Number.isFinite(current) || !Number.isFinite(previous) || previous === 0) {
    return {
      direction: "no comparison",
      absolute: null,
      percent: null,
    };
  }

  const absolute = current - previous;
  const percent = (absolute / previous) * 100;

  if (Math.abs(absolute) < 0.0000001) {
    return {
      direction: "stable",
      absolute: 0,
      percent: 0,
    };
  }

  return {
    direction: current > previous ? "increased" : "decreased",
    absolute,
    percent,
  };
}

function getDirectionLabel(value: number | null, previous: number | null) {
  const comparison = compareNumericChange(value, previous);

  if (comparison.direction === "no comparison") {
    return "no change data";
  }

  if (comparison.direction === "stable") {
    return "remained stable";
  }

  return comparison.direction;
}

function buildCampaignRecommendation(selectedWeek: WeeklyMetric | null, previousWeek: WeeklyMetric | null) {
  if (!selectedWeek) {
    return {
      title: "Monitor campaign performance.",
      actions: [
        "Upload a valid Excel file with campaign data.",
        "Select a valid campaign and week to generate a recommendation.",
      ],
    };
  }

  const currentRoas = selectedWeek.roas;
  const prevRoas = previousWeek?.roas ?? null;
  const currentCpi = selectedWeek.cpi;
  const prevCpi = previousWeek?.cpi ?? null;
  const currentUv = selectedWeek.uv;
  const prevUv = previousWeek?.uv ?? null;
  const currentCtr = selectedWeek.ctr;
  const prevCtr = previousWeek?.ctr ?? null;
  const currentCvr = selectedWeek.cvr;
  const prevCvr = previousWeek?.cvr ?? null;
  const cpiUp = currentCpi !== null && prevCpi !== null && currentCpi > prevCpi;
  const cpiDown = currentCpi !== null && prevCpi !== null && currentCpi < prevCpi;
  const uvDown = currentUv !== null && prevUv !== null && currentUv < prevUv;
  const spendUp = selectedWeek.cost > (previousWeek?.cost ?? 0);
  const ctrDown = currentCtr !== null && prevCtr !== null && currentCtr < prevCtr;
  const cvrDown = currentCvr !== null && prevCvr !== null && currentCvr < prevCvr;

  if (currentRoas !== null && prevRoas !== null && currentRoas < 1 && currentRoas > prevRoas) {
    return {
      title: "Scale Winning Creatives.",
      actions: [
        "Continue developing the creatives and ad groups contributing to the ROAS improvement.",
        "Monitor whether ROAS can reach >= 1 while CPI and UV remain stable.",
        "Avoid aggressive budget expansion while the campaign remains below break-even.",
      ],
    };
  }

  if (currentRoas !== null && prevRoas !== null && currentRoas >= 1 && currentRoas > prevRoas) {
    return {
      title: "Scale Winning Creatives / Consider Controlled Budget Increase.",
      actions: [
        "Gradually increase exposure to the creatives and ad groups contributing to the improvement.",
        "Monitor CPI and UV closely as spend increases.",
        "Maintain controlled scaling if ROAS remains above break-even.",
      ],
    };
  }

  if (cpiUp && spendUp && ctrDown && cvrDown) {
    return {
      title: "Adjust Budget or Launch a New Campaign / Ad Group to Refresh Learning.",
      actions: [
        "Reduce or adjust the current scaling pace and monitor CPI, CTR, and CVR.",
        "If efficiency does not recover, launch a new campaign or ad group to refresh learning.",
        "Avoid aggressive budget expansion while acquisition efficiency remains deteriorated.",
      ],
    };
  }

  if (cpiUp && !ctrDown && !cvrDown) {
    return {
      title: "Refresh / Replace Creatives.",
      actions: [
        "Review creative asset performance for signs of creative fatigue.",
        "Test new creative variations and monitor CPI after the refresh.",
        "Avoid aggressive scaling until acquisition efficiency stabilizes.",
      ],
    };
  }

  if (uvDown && !selectedWeek.noteLimited) {
    return {
      title: "Turn Off Underperforming Ad Groups or Replace Creatives.",
      actions: [
        "Identify ad groups or creatives associated with lower user value.",
        "Reduce exposure to underperforming segments and test replacement creatives.",
        "Monitor UV and ROAS after the adjustments.",
      ],
    };
  }

  if (uvDown && selectedWeek.noteLimited) {
    return {
      title: "Increase tROAS.",
      actions: [
        "Reassess user quality and UV after the target adjustment.",
        "Avoid aggressive budget expansion while the campaign remains below break-even.",
      ],
    };
  }

  if (cpiUp && uvDown) {
    const primaryDriver = Math.abs((currentCpi ?? 0) - (prevCpi ?? 0)) >= Math.abs((currentUv ?? 0) - (prevUv ?? 0)) ? "Adjust Budget or Launch a New Campaign / Ad Group to Refresh Learning." : "Turn Off Underperforming Ad Groups or Replace Creatives.";
    return {
      title: primaryDriver,
      actions: [
        "Address the primary driver using the corresponding action above.",
        "Monitor the secondary driver after the change.",
        "Avoid aggressive budget expansion while overall efficiency is deteriorating.",
      ],
    };
  }

  if (currentRoas !== null && currentRoas >= 1) {
    return {
      title: "Scale Winning Creatives / Consider Controlled Budget Increase.",
      actions: [
        "Keep scaling the creatives and ad groups that are driving sustained efficiency.",
        "Monitor CPI and UV as spend rises.",
        "Hold a controlled pace if performance becomes less stable.",
      ],
    };
  }

  return {
    title: "Adjust Budget or Launch a New Campaign / Ad Group to Refresh Learning.",
    actions: [
      "Reduce or reallocate spend toward more efficient traffic.",
      "Test a fresh campaign or ad group to recover learning.",
      "Avoid continued acceleration while acquisition efficiency remains weak.",
    ],
  };
}

function buildDriverNarrative(selectedWeek: WeeklyMetric | null, previousWeek: WeeklyMetric | null) {
  if (!selectedWeek) {
    return {
      title: "Low Performance",
      summary: "No selected week is available.",
      recommendation: "Upload a valid Excel file with campaign data.",
    };
  }

  const currentRoas = selectedWeek.roas;
  const prevRoas = previousWeek?.roas ?? null;
  const currentCpi = selectedWeek.cpi;
  const prevCpi = previousWeek?.cpi ?? null;
  const currentUv = selectedWeek.uv;
  const prevUv = previousWeek?.uv ?? null;

  if (!previousWeek) {
    const roasText = currentRoas === null ? "N/A" : formatRatio(currentRoas);
    const status = currentRoas !== null && currentRoas >= 1 ? "Good Performance" : "Low Performance";
    return {
      title: status,
      summary: `Current ROAS is ${roasText}. There is no previous-week comparison available, so the performance should be interpreted in isolation.`,
      recommendation: currentRoas !== null && currentRoas >= 1 ? "Maintain the current approach and monitor the next comparable week." : "Monitor spend efficiency and creative quality until a prior week becomes available for comparison.",
    };
  }

  const roasDelta = getPercentChange(currentRoas, prevRoas);
  const deltaText = roasDelta === null || !Number.isFinite(roasDelta) ? "N/A" : `${roasDelta >= 0 ? "+" : ""}${roasDelta.toFixed(1)}%`;
  const roasText = currentRoas === null ? "N/A" : formatRatio(currentRoas);
  const prevText = prevRoas === null ? "N/A" : formatRatio(prevRoas);
  const cpiBetter = currentCpi !== null && prevCpi !== null && currentCpi < prevCpi;
  const cpiWorse = currentCpi !== null && prevCpi !== null && currentCpi > prevCpi;
  const uvBetter = currentUv !== null && prevUv !== null && currentUv > prevUv;
  const uvWorse = currentUv !== null && prevUv !== null && currentUv < prevUv;

  if (currentRoas !== null && prevRoas !== null && currentRoas >= 1 && currentRoas > prevRoas) {
    const driverMessage =
      cpiBetter && uvBetter
        ? `CPI improved from ${formatRatio(prevCpi)} to ${formatRatio(currentCpi)} and UV improved from ${formatRatio(prevUv)} to ${formatRatio(currentUv)}. The ROAS increase is therefore explained by stronger acquisition efficiency and better user value.`
        : cpiBetter
          ? `CPI improved from ${formatRatio(prevCpi)} to ${formatRatio(currentCpi)}, which lowered acquisition cost and supported the stronger ROAS result.`
          : uvBetter
            ? `UV increased from ${formatRatio(prevUv)} to ${formatRatio(currentUv)}, which lifted revenue per install and improved ROAS.`
            : `ROAS improved even though the driver mix was mixed, with the strongest measurable uplift coming from the current week values.`;

    return {
      title: "Good Performance",
      summary: `Current ROAS is ${roasText}, above the 1.00 break-even threshold and up from ${prevText} (${deltaText}). ${driverMessage}`,
      recommendation: "Maintain the current pacing and scale the strongest performing traffic while monitoring for saturation.",
    };
  }

  if (currentRoas !== null && prevRoas !== null && currentRoas >= 1 && currentRoas < prevRoas) {
    const driverMessage =
      cpiWorse && uvWorse
        ? `CPI increased from ${formatRatio(prevCpi)} to ${formatRatio(currentCpi)} while UV declined from ${formatRatio(prevUv)} to ${formatRatio(currentUv)}. Both acquisition efficiency and user value worsened, so both metrics pressured ROAS.`
        : cpiWorse
          ? `CPI increased from ${formatRatio(prevCpi)} to ${formatRatio(currentCpi)}. The spend increase raised acquisition cost enough to reduce ROAS despite the campaign staying above break-even.`
          : uvWorse
            ? `UV decreased from ${formatRatio(prevUv)} to ${formatRatio(currentUv)}. Lower value per install reduced ROAS even though the campaign remains profitable.`
            : `The campaign remains profitable but the performance signal is weaker than the previous week. The move is mixed rather than driven by a single metric.`;

    return {
      title: "Profitable but Below Previous Week",
      summary: `Current ROAS is ${roasText}, below the previous ${prevText} (${deltaText}). The campaign remains above break-even, but performance slipped. ${driverMessage}`,
      recommendation: cpiWorse ? "Adjust budget or launch a new campaign/ad group to refresh learning." : uvWorse ? "Turn off underperforming ad groups or replace creatives." : "Maintain / Control Budget",
    };
  }

  if (currentRoas !== null && prevRoas !== null && currentRoas < 1 && currentRoas < prevRoas) {
    const driverMessage =
      cpiWorse && uvWorse
        ? `ROAS declined from ${prevText} to ${roasText} (${deltaText}). CPI increased from ${formatRatio(prevCpi)} to ${formatRatio(currentCpi)} and UV fell from ${formatRatio(prevUv)} to ${formatRatio(currentUv)}. Both acquisition cost and user value weakened, with the larger movement appearing in the efficiency component.`
        : cpiWorse
          ? `ROAS dropped from ${prevText} to ${roasText} (${deltaText}) because CPI increased from ${formatRatio(prevCpi)} to ${formatRatio(currentCpi)}. Signals are consistent with possible scaling dilution or rising acquisition costs.`
          : uvWorse
            ? `ROAS dropped from ${prevText} to ${roasText} (${deltaText}) because UV declined from ${formatRatio(prevUv)} to ${formatRatio(currentUv)}. Lower-value users reduced the monetization yield.`
            : `ROAS declined from ${prevText} to ${roasText} (${deltaText}), and the change appears mixed rather than dominated by a single driver.`;

    const recommendation =
      cpiWorse && selectedWeek.cost > (previousWeek.cost ?? 0) && selectedWeek.ctr !== null && previousWeek.ctr !== null && selectedWeek.ctr < previousWeek.ctr && selectedWeek.cvr !== null && previousWeek.cvr !== null && selectedWeek.cvr < previousWeek.cvr
        ? "Adjust budget or launch a new campaign/ad group to refresh learning."
        : cpiWorse
          ? "Change / refresh creatives."
          : uvWorse && selectedWeek.noteLimited
            ? "Increase tROAS."
            : uvWorse
              ? "Turn off underperforming ad groups or replace creatives."
              : "Adjust budget or launch a new campaign/ad group to refresh learning.";

    return {
      title: "Low Performance",
      summary: `Current ROAS is ${roasText}, below the 1.00 break-even threshold and down from ${prevText} (${deltaText}). ${driverMessage}`,
      recommendation,
    };
  }

  if (currentRoas !== null && prevRoas !== null && currentRoas < 1 && currentRoas > prevRoas) {
    const driverMessage =
      cpiBetter && uvBetter
        ? `ROAS improved from ${prevText} to ${roasText} (${deltaText}). CPI fell from ${formatRatio(prevCpi)} to ${formatRatio(currentCpi)} and UV increased from ${formatRatio(prevUv)} to ${formatRatio(currentUv)}, so the improvement is being driven by better acquisition efficiency and higher user value.`
        : cpiBetter
          ? `ROAS improved from ${prevText} to ${roasText} (${deltaText}) as CPI fell from ${formatRatio(prevCpi)} to ${formatRatio(currentCpi)}, reducing acquisition cost.`
          : uvBetter
            ? `ROAS improved from ${prevText} to ${roasText} (${deltaText}) as UV increased from ${formatRatio(prevUv)} to ${formatRatio(currentUv)}, improving revenue per install.`
            : `ROAS improved from ${prevText} to ${roasText} (${deltaText}), but the lift is mixed and should be interpreted cautiously.`;

    return {
      title: "Good Signal — Not Yet Profitable",
      summary: `Current ROAS is ${roasText}, which is still below 1.00, but it improved from ${prevText} (${deltaText}). ${driverMessage}`,
      recommendation: "Scale Winning Creatives",
    };
  }

  return {
    title: currentRoas !== null && currentRoas >= 1 ? "Good Performance" : "Low Performance",
    summary: `Current ROAS is ${roasText}. The current week cannot be fully benchmarked against the previous week because the comparison is incomplete or unavailable.`,
    recommendation: "Monitor the campaign and review traffic quality before increasing spend.",
  };
}

export default function Home() {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [rawRows, setRawRows] = useState<ImportedRow[]>([]);
  const [errorMessage, setErrorMessage] = useState<string>("");
  const [fileName, setFileName] = useState<string>("");
  const [selectedCampaign, setSelectedCampaign] = useState<string>("");
  const [selectedWeekKey, setSelectedWeekKey] = useState<string>("");
  const [totalBudget, setTotalBudget] = useState<number | null>(null);

  const allWeeklyMetrics = useMemo(() => buildWeeklyMetrics(rawRows), [rawRows]);
  const allDailyMetrics = useMemo(() => buildDailyMetrics(rawRows), [rawRows]);

  const campaignNames = useMemo(() => {
    const names = new Set<string>();
    allWeeklyMetrics.forEach((metric) => names.add(metric.campaignName));

    return [...names].sort((a, b) => {
      const latestA = allWeeklyMetrics.filter((week) => week.campaignName === a).sort((x, y) => y.weekStart.getTime() - x.weekStart.getTime())[0];
      const latestB = allWeeklyMetrics.filter((week) => week.campaignName === b).sort((x, y) => y.weekStart.getTime() - x.weekStart.getTime())[0];
      return (latestB?.weekStart.getTime() ?? 0) - (latestA?.weekStart.getTime() ?? 0);
    });
  }, [allWeeklyMetrics]);

  useEffect(() => {
    if (campaignNames.length === 0) {
      setSelectedCampaign("");
      return;
    }

    if (!selectedCampaign || !campaignNames.includes(selectedCampaign)) {
      setSelectedCampaign(campaignNames[0]);
    }
  }, [campaignNames, selectedCampaign]);

  const selectedCampaignWeekly = useMemo(
    () => [...allWeeklyMetrics].filter((metric) => metric.campaignName === selectedCampaign).sort((a, b) => b.weekStart.getTime() - a.weekStart.getTime()),
    [allWeeklyMetrics, selectedCampaign],
  );

  const selectedCampaignDaily = useMemo(
    () => [...allDailyMetrics].filter((metric) => metric.campaignName === selectedCampaign).sort((a, b) => a.dateKey.localeCompare(b.dateKey)),
    [allDailyMetrics, selectedCampaign],
  );

  useEffect(() => {
    if (selectedCampaignWeekly.length === 0) {
      setSelectedWeekKey("");
      return;
    }

    if (!selectedWeekKey || !selectedCampaignWeekly.some((week) => week.weekKey === selectedWeekKey)) {
      setSelectedWeekKey(selectedCampaignWeekly[0].weekKey);
    }
  }, [selectedCampaignWeekly, selectedWeekKey]);

  const selectedWeek = selectedCampaignWeekly.find((week) => week.weekKey === selectedWeekKey) ?? selectedCampaignWeekly[0] ?? null;
  const selectedWeekIndex = selectedWeek ? selectedCampaignWeekly.findIndex((week) => week.weekKey === selectedWeek.weekKey) : -1;
  const previousWeek = selectedWeekIndex >= 0 ? selectedCampaignWeekly[selectedWeekIndex + 1] ?? null : null;
  const driverConclusion = buildDriverNarrative(selectedWeek, previousWeek);
  const campaignRecommendation = buildCampaignRecommendation(selectedWeek, previousWeek);

  const spendByCampaign = useMemo(() => {
    const map = new Map<string, number>();
    rawRows.forEach((row) => {
      const campaignName = String(row["Campaign Name"] ?? "").trim();
      if (!campaignName) {
        return;
      }
      map.set(campaignName, (map.get(campaignName) ?? 0) + sanitizeNumber(row.Cost));
    });
    return [...map.entries()].sort((a, b) => b[1] - a[1]);
  }, [rawRows]);

  const totalSpend = spendByCampaign.reduce((sum, [, value]) => sum + value, 0);
  const remainingBudget = totalBudget !== null ? totalBudget - totalSpend : 0;
  const spendRate = totalBudget && totalBudget > 0 ? totalSpend / totalBudget : null;

  const budgetChartData = useMemo(() => {
    const chartValues: Record<string, string | number> = { name: "Portfolio" };
    spendByCampaign.forEach(([campaignName, value]) => {
      chartValues[campaignName] = value;
    });
    if (totalBudget !== null) {
      chartValues.Remaining = Math.max(totalBudget - totalSpend, 0);
    }
    return [chartValues];
  }, [spendByCampaign, totalBudget, totalSpend]);

  const campaignRecommendationRows = useMemo(() => {
    return campaignNames.map((campaignName) => {
      const weeks = [...allWeeklyMetrics].filter((metric) => metric.campaignName === campaignName).sort((a, b) => b.weekStart.getTime() - a.weekStart.getTime());
      const current = weeks[0] ?? null;
      const previous = weeks[1] ?? null;
      const currentRoas = current?.roas ?? null;
      const previousRoas = previous?.roas ?? null;
      const wowRoas = getPercentChange(currentRoas, previousRoas);
      const currentCpi = current?.cpi ?? null;
      const currentUv = current?.uv ?? null;
      const currentSpend = current?.cost ?? 0;

      let action = "Maintain / Control Budget";
      let reason = "No WoW comparison available";

      if (previous && currentRoas !== null && previousRoas !== null) {
        if (currentRoas >= 1 && currentRoas > previousRoas) {
          action = "Increase / Add Budget";
          reason = `ROAS improved from ${formatRatio(previousRoas)} to ${formatRatio(currentRoas)} and is above break-even. CPI is ${currentCpi === null ? "N/A" : formatRatio(currentCpi)}, and UV is ${currentUv === null ? "N/A" : formatRatio(currentUv)}.`;
        } else if (currentRoas < 1 && currentRoas < previousRoas) {
          action = "Pause / Refresh Campaign";
          reason = `ROAS declined from ${formatRatio(previousRoas)} to ${formatRatio(currentRoas)} and remains below break-even. Current spend is ${formatCurrency(currentSpend)}.`;
        } else if (currentRoas >= 1 && currentRoas <= previousRoas) {
          action = "Maintain / Control Budget";
          reason = `ROAS is ${formatRatio(currentRoas)} and remains above break-even, but the trend is flat or lower. Current spend is ${formatCurrency(currentSpend)}.`;
        } else if (currentRoas < 1 && currentRoas > previousRoas) {
          action = "Maintain / Control Budget";
          reason = `ROAS improved from ${formatRatio(previousRoas)} to ${formatRatio(currentRoas)}, but remains below break-even. Additional expansion should be controlled.`;
        }
      } else if (currentRoas !== null) {
        action = currentRoas >= 1 ? "Increase / Add Budget" : "Maintain / Control Budget";
        reason = currentRoas >= 1 ? `Current ROAS is ${formatRatio(currentRoas)} and above break-even, with no previous week available for comparison.` : `Current ROAS is ${formatRatio(currentRoas)} and below break-even. No previous week comparison is available.`;
      }

      return {
        campaignName,
        currentRoas,
        previousRoas,
        wowRoas,
        currentCpi,
        currentUv,
        currentSpend,
        action,
        reason,
      };
    });
  }, [campaignNames, allWeeklyMetrics]);

  const weeklyComparisonData =
    selectedWeek && previousWeek
      ? [
          { name: "Previous Week", roas: previousWeek.roas, cpi: previousWeek.cpi, uv: previousWeek.uv },
          { name: "Selected Week", roas: selectedWeek.roas, cpi: selectedWeek.cpi, uv: selectedWeek.uv },
        ]
      : [];

  const processFile = async (file: File) => {
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });

      if (workbook.SheetNames.length !== 1 || workbook.SheetNames[0].trim() !== "Raw Data") {
        throw new Error("Please upload an Excel file containing exactly one worksheet named Raw Data.");
      }

      const dataSheet = findRawDataSheet(workbook);
      if (!dataSheet) {
        throw new Error("Please upload an Excel file containing exactly one worksheet named Raw Data.");
      }

      const parsedRows = parseUploadedRows(dataSheet.rows);
      const budgetValue = readTotalBudgetFromSheet(workbook.Sheets["Raw Data"]) ?? extractTotalBudgetFromWorkbook(workbook) ?? extractTotalBudget(dataSheet.rows);

      if (budgetValue === null || !Number.isFinite(budgetValue)) {
        throw new Error("Missing Total Budget. Please add Total Budget to the Raw Data sheet.");
      }

      if (parsedRows.length === 0) {
        throw new Error("No usable rows were found in the uploaded Excel file.");
      }

      if (parsedRows.some((row) => String(row["Campaign Name"] ?? "").trim() === "")) {
        throw new Error("Campaign Name is missing from one or more rows. Please complete the required campaign field.");
      }

      setRawRows(parsedRows);
      setTotalBudget(budgetValue);
      setErrorMessage("");
      setFileName(file.name);
    } catch (error) {
      setRawRows([]);
      setTotalBudget(null);
      setSelectedCampaign("");
      setSelectedWeekKey("");
      setErrorMessage(error instanceof Error ? error.message : "The Excel file could not be processed.");
      setFileName("");
    }
  };

  const handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0];
    if (!selectedFile) return;
    await processFile(selectedFile);
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  const conclusionTone =
    driverConclusion.title === "Good Performance"
      ? "border-violet-200 bg-violet-50 text-violet-800"
      : driverConclusion.title === "Profitable but Below Previous Week"
        ? "border-amber-200 bg-amber-50 text-amber-800"
        : driverConclusion.title === "Good Signal — Not Yet Profitable"
          ? "border-purple-200 bg-purple-50 text-purple-800"
          : "border-rose-200 bg-rose-50 text-rose-800";

  return (
    <div className="dashboard-canvas min-h-screen px-3 py-4 text-slate-900 sm:px-6 sm:py-8 lg:px-8">
      <div className="dashboard-shell mx-auto max-w-7xl">
        <header className="dashboard-header mb-7 border-b border-slate-200 pb-5">
          <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
            <div>
              <p className="text-[11px] font-semibold uppercase tracking-[0.16em] text-violet-700">Milena Nguyen - Mobile App</p>
              <h1 className="mt-1 text-2xl font-semibold tracking-tight text-slate-900 sm:text-3xl">Campaign Performance</h1>
              <p className="mt-1 text-sm text-slate-500">Weekly and daily acquisition performance analysis</p>
            </div>
            <div className="flex items-center gap-3">
              <input ref={fileInputRef} id="excel-upload" type="file" accept=".xlsx,.xls,.csv" onChange={handleFileChange} className="hidden" />
              <label htmlFor="excel-upload" className="inline-flex cursor-pointer items-center justify-center rounded-lg border border-violet-200 bg-white px-3.5 py-2 text-sm font-semibold text-violet-800 transition hover:border-violet-300 hover:bg-violet-50">Upload Excel</label>
            </div>
          </div>
        </header>

        {errorMessage ? (
          <div className="mb-6 rounded-xl border border-red-200 bg-red-50 p-4 text-sm text-red-700">{errorMessage}</div>
        ) : null}

        {fileName ? (
          <div className="mb-6 flex items-center gap-2 text-sm text-slate-600">
            <span className="inline-flex rounded-full bg-slate-200 px-2.5 py-1 font-medium text-slate-700">File loaded</span>
            <span>{fileName}</span>
          </div>
        ) : null}

        {!errorMessage && rawRows.length === 0 ? (
          <div className="mb-6 rounded-2xl border border-dashed border-slate-300 bg-white p-10 text-center text-slate-600">Upload an Excel file to calculate campaign pacing, weekly performance, and driver analysis.</div>
        ) : null}

        {rawRows.length > 0 ? (
          <>
            <section className="dashboard-section mb-7">
              <div className="mb-4">
                <p className="text-[11px] font-semibold uppercase tracking-[0.16em] text-slate-500">1. Campaign Spend Pacing</p>
                <h2 className="mt-1 text-2xl font-semibold text-slate-900">Campaign Spend Pacing</h2>
              </div>

              <div className="mb-5 grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
                <div className="rounded-2xl border border-violet-200 bg-violet-50 p-4">
                  <p className="text-xs font-semibold uppercase tracking-[0.12em] text-violet-700">Total Budget</p>
                  <div className="mt-2 text-2xl font-semibold text-slate-900">{totalBudget === null ? "N/A" : formatCurrency(totalBudget)}</div>
                </div>
                <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
                  <p className="text-xs font-semibold uppercase tracking-[0.12em] text-slate-600">Total Spend</p>
                  <div className="mt-2 text-2xl font-semibold text-slate-900">{formatCurrency(totalSpend)}</div>
                </div>
                <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
                  <p className="text-xs font-semibold uppercase tracking-[0.12em] text-slate-600">Remaining Budget</p>
                  <div className="mt-2 text-2xl font-semibold text-slate-900">{formatCurrency(remainingBudget)}</div>
                </div>
                <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
                  <p className="text-xs font-semibold uppercase tracking-[0.12em] text-slate-600">Spend Rate</p>
                  <div className="mt-2 text-2xl font-semibold text-slate-900">{spendRate === null ? "N/A" : `${(spendRate * 100).toFixed(1)}%`}</div>
                </div>
              </div>

              {spendByCampaign.length > 0 && totalBudget !== null ? (
                <>
                  <div className="chart-surface h-72 w-full rounded-xl border border-slate-200 bg-white p-3">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={budgetChartData} layout="vertical" margin={{ top: 10, right: 20, left: 20, bottom: 10 }}>
                        <CartesianGrid stroke="#eee9f4" strokeDasharray="4 4" />
                        <XAxis type="number" domain={[0, totalBudget || 1]} tickFormatter={(value) => `$${Number(value).toLocaleString()}`} tick={{ fontSize: 11, fill: "#777487" }} axisLine={false} />
                        <YAxis type="category" dataKey="name" tick={{ fontSize: 12, fill: "#777487" }} axisLine={false} width={90} />
                        <Tooltip
                          formatter={(value: number | string | ReadonlyArray<number | string> | undefined) => {
                            const numeric = Number(Array.isArray(value) ? value[0] : value ?? 0);
                            return formatCurrency(numeric);
                          }}
                          labelFormatter={() => "Portfolio"}
                        />
                        <Legend />
                        {spendByCampaign.map(([campaignName, value], index) => (
                          <Bar key={campaignName} dataKey={campaignName} stackId="budget" fill={["#7c4dff", "#c84dad", "#5367c8", "#a855f7", "#f59e0b", "#ec4899"][index % 6]} radius={index === 0 ? [0, 6, 6, 0] : [0, 0, 0, 0]} />
                        ))}
                        {totalBudget !== null ? <Bar dataKey="Remaining" stackId="budget" fill="#e9e3f4" radius={[0, 6, 6, 0]} /> : null}
                      </BarChart>
                    </ResponsiveContainer>
                  </div>

                  <p className="mt-4 text-sm leading-6 text-slate-700">
                    Total spend is <strong>{formatCurrency(totalSpend)}</strong>, representing <strong>{spendRate === null ? "N/A" : `${(spendRate * 100).toFixed(1)}%`}</strong> of the <strong>{formatCurrency(totalBudget)}</strong> total budget. <strong>{formatCurrency(remainingBudget)}</strong> remains available. {spendByCampaign[0] ? (<span>Campaign <strong>{spendByCampaign[0][0]}</strong> accounts for <strong>{((spendByCampaign[0][1] / totalBudget) * 100).toFixed(1)}%</strong> of the budget consumption.</span>) : null}
                  </p>
                </>
              ) : (
                <div className="rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-800">A valid Total Budget is required to display campaign pacing.</div>
              )}
            </section>

            <section className="dashboard-section mb-7">
              <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
                <div>
                  <p className="text-[11px] font-semibold uppercase tracking-[0.16em] text-slate-500">2. Weekly Performance Overview</p>
                  <h2 className="mt-1 text-2xl font-semibold text-slate-900">Weekly Performance Overview</h2>
                </div>
                <div className="w-full max-w-sm">
                  <label className="mb-1 block text-sm font-medium text-slate-700">Analyze Campaign</label>
                  <select
                    value={selectedCampaign}
                    onChange={(event) => setSelectedCampaign(event.target.value)}
                    className="w-full rounded-lg border border-violet-200 bg-white px-3 py-2 text-sm text-slate-800 shadow-sm focus:border-violet-400 focus:outline-none"
                  >
                    {campaignNames.map((campaignName) => (
                      <option key={campaignName} value={campaignName}>{getCampaignDisplayName(campaignName)}</option>
                    ))}
                  </select>
                </div>
              </div>

              {selectedCampaignWeekly.length > 0 ? (
                <div className="weekly-table-scroll max-h-[360px] overflow-y-auto overflow-x-auto rounded-xl border border-slate-200 bg-white">
                  <table className="min-w-full border-separate border-spacing-0 text-left text-sm">
                    <thead className="bg-violet-50/90">
                      <tr>
                        <th className="px-3 py-2.5 text-left text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-600">Week</th>
                        <th className="px-3 py-2.5 text-right text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-600">Cost</th>
                        <th className="px-3 py-2.5 text-right text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-600">CPI</th>
                        <th className="px-3 py-2.5 text-right text-[11px] font-semibold uppercase tracking-[0.12em] text-violet-700">ROAS</th>
                        <th className="px-3 py-2.5 text-right text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-600">UV</th>
                        <th className="px-3 py-2.5 text-right text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-600">CTR</th>
                        <th className="px-3 py-2.5 text-right text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-600">CVR</th>
                      </tr>
                    </thead>
                    <tbody>
                      {selectedCampaignWeekly.map((week, index) => {
                        const previousMetric = selectedCampaignWeekly[index + 1] ?? null;
                        const renderTrendCell = (value: number | null, prevValue: number | null, lowerIsBetter = false, formatter: (input: number | null) => string) => {
                          const delta = getPercentChange(value, prevValue);
                          const trend = formatTrend(delta, lowerIsBetter);
                          return (
                            <div className="space-y-1">
                              <div className="text-right font-semibold text-slate-900">{formatter(value)}</div>
                              <div className={`ml-auto inline-flex rounded-full px-1.5 py-0.5 text-[10px] font-semibold ${delta === null || !Number.isFinite(delta) ? "bg-slate-100 text-slate-500" : trend.improvement ? "bg-emerald-100 text-emerald-700" : "bg-rose-100 text-rose-700"}`}>
                                {trend.label}
                              </div>
                            </div>
                          );
                        };

                        return (
                          <tr key={week.weekKey} className="border-t border-slate-100 hover:bg-violet-50/40">
                            <td className="px-3 py-3 text-left font-medium text-slate-900">{week.weekLabel}</td>
                            <td className="px-3 py-3 text-right">{renderTrendCell(week.cost, previousMetric?.cost ?? null, true, (input) => input === null ? "N/A" : formatCurrency(input))}</td>
                            <td className="px-3 py-3 text-right">{renderTrendCell(week.cpi, previousMetric?.cpi ?? null, true, (input) => input === null ? "N/A" : formatMetricValue(input))}</td>
                            <td className="px-3 py-3 text-right"><span className="font-bold text-violet-700">{renderTrendCell(week.roas, previousMetric?.roas ?? null, false, (input) => input === null ? "N/A" : formatMetricValue(input))}</span></td>
                            <td className="px-3 py-3 text-right">{renderTrendCell(week.uv, previousMetric?.uv ?? null, false, (input) => input === null ? "N/A" : formatMetricValue(input))}</td>
                            <td className="px-3 py-3 text-right">{renderTrendCell(week.ctr, previousMetric?.ctr ?? null, false, (input) => input === null ? "N/A" : formatMetricValue(input))}</td>
                            <td className="px-3 py-3 text-right">{renderTrendCell(week.cvr, previousMetric?.cvr ?? null, false, (input) => input === null ? "N/A" : formatMetricValue(input))}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              ) : (
                <div className="rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-800">This campaign has no weekly data available.</div>
              )}
            </section>

            <section className="dashboard-section mb-7">
              <div className="mb-4">
                <p className="text-[11px] font-semibold uppercase tracking-[0.16em] text-slate-500">3. Daily Performance</p>
                <h2 className="mt-1 text-2xl font-semibold text-slate-900">Daily Performance</h2>
                <p className="mt-1 text-sm text-slate-600">ROAS, CPI and UV with daily installs</p>
              </div>

              {selectedCampaignDaily.length > 0 ? (
                <div className="chart-surface h-[340px] w-full rounded-xl border border-slate-200 bg-white p-3">
                  <ResponsiveContainer width="100%" height="100%">
                    <ComposedChart data={selectedCampaignDaily}>
                      <CartesianGrid stroke="#eee9f4" strokeDasharray="4 4" />
                      <XAxis dataKey="dateLabel" tick={{ fontSize: 11, fill: "#777487" }} tickLine={false} axisLine={{ stroke: "#e8e3ef" }} />
                      <YAxis yAxisId="metrics" tick={{ fontSize: 11, fill: "#777487" }} tickLine={false} axisLine={false} domain={[(dataMin: number) => Math.min(0, dataMin), (dataMax: number) => Math.max(1, dataMax * 1.3)]} />
                      <YAxis yAxisId="installs" orientation="right" tick={{ fontSize: 11, fill: "#777487" }} tickLine={false} axisLine={false} domain={[(dataMin: number) => Math.max(0, dataMin * 0.8), (dataMax: number) => Math.max(1, dataMax * 1.2)]} />
                      <Tooltip
                        formatter={(value: number | string | ReadonlyArray<number | string> | undefined, name: string | number | undefined) => {
                          const numeric = Number(Array.isArray(value) ? value[0] : value ?? 0);
                          return [name === "Installs" ? new Intl.NumberFormat("en-US").format(Math.round(numeric)) : numeric.toFixed(2), String(name ?? "")] as [string, string];
                        }}
                        labelFormatter={(label) => `Date: ${label}`}
                      />
                      <Legend wrapperStyle={{ fontSize: 12, paddingTop: 8 }} />
                      <Line yAxisId="metrics" dataKey="roas" name="ROAS" type="monotone" stroke="#7c4dff" strokeWidth={2.8} dot={false} activeDot={{ r: 4 }} />
                      <Line yAxisId="metrics" dataKey="cpi" name="CPI" type="monotone" stroke="#c84dad" strokeWidth={2.5} dot={false} activeDot={{ r: 4 }} />
                      <Line yAxisId="metrics" dataKey="uv" name="UV" type="monotone" stroke="#5367c8" strokeWidth={2.5} dot={false} activeDot={{ r: 4 }} />
                      <Bar yAxisId="installs" dataKey="installs" name="Installs" fill="#cfc0ff" radius={[8, 8, 0, 0]} maxBarSize={24} />
                    </ComposedChart>
                  </ResponsiveContainer>
                </div>
              ) : (
                <div className="rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-800">No daily data is available for the selected campaign.</div>
              )}
            </section>

            <section className="dashboard-section mb-7">
              <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
                <div>
                  <p className="text-[11px] font-semibold uppercase tracking-[0.16em] text-slate-500">4. Driver Analysis</p>
                  <h2 className="mt-1 text-2xl font-semibold text-slate-900">Driver Analysis</h2>
                </div>
                <div className="flex w-full max-w-xl flex-col gap-3 sm:flex-row">
                  <div className="flex-1">
                    <label className="mb-1 block text-sm font-medium text-slate-700">Analyze Campaign</label>
                    <select value={selectedCampaign} onChange={(event) => setSelectedCampaign(event.target.value)} className="w-full rounded-lg border border-violet-200 bg-white px-3 py-2 text-sm text-slate-800 shadow-sm focus:border-violet-400 focus:outline-none">
                      {campaignNames.map((campaignName) => (
                        <option key={campaignName} value={campaignName}>{getCampaignDisplayName(campaignName)}</option>
                      ))}
                    </select>
                  </div>
                  <div className="flex-1">
                    <label className="mb-1 block text-sm font-medium text-slate-700">Analyze Week</label>
                    <select value={selectedWeekKey} onChange={(event) => setSelectedWeekKey(event.target.value)} className="w-full rounded-lg border border-violet-200 bg-white px-3 py-2 text-sm text-slate-800 shadow-sm focus:border-violet-400 focus:outline-none">
                      {selectedCampaignWeekly.map((week) => (
                        <option key={week.weekKey} value={week.weekKey}>{week.weekLabel}</option>
                      ))}
                    </select>
                  </div>
                </div>
              </div>

              {selectedWeek ? (
                <>
                  <div className="mb-5 grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
                    <div className="metric-chip metric-chip-roas rounded-2xl border p-4">
                      <div>ROAS</div>
                      <div>{selectedWeek.roas === null ? "N/A" : formatMetricValue(selectedWeek.roas)}</div>
                      <div>{selectedWeek.weekLabel}</div>
                    </div>
                    <div className="metric-chip metric-chip-cpi rounded-2xl border p-4">
                      <div>CPI</div>
                      <div>{selectedWeek.cpi === null ? "N/A" : formatCurrency(selectedWeek.cpi)}</div>
                      <div>{selectedWeek.weekLabel}</div>
                    </div>
                    <div className="metric-chip metric-chip-uv rounded-2xl border p-4">
                      <div>UV</div>
                      <div>{selectedWeek.uv === null ? "N/A" : formatMetricValue(selectedWeek.uv)}</div>
                      <div>{selectedWeek.weekLabel}</div>
                    </div>
                    <div className="metric-chip metric-chip-cost rounded-2xl border p-4">
                      <div>Cost</div>
                      <div>{selectedWeek.cost === null ? "N/A" : formatCurrency(selectedWeek.cost)}</div>
                      <div>{selectedWeek.weekLabel}</div>
                    </div>
                  </div>

                  {previousWeek ? (
                    <div className="mb-5 rounded-xl border border-slate-200 bg-slate-50 p-4">
                      <div className="mb-3 text-lg font-semibold text-slate-900">Weekly Performance Comparison</div>
                      <div className="chart-surface h-[300px] w-full rounded-xl border border-slate-200 bg-white p-3">
                        <ResponsiveContainer width="100%" height="100%">
                          <LineChart data={weeklyComparisonData}>
                            <CartesianGrid stroke="#eee9f4" strokeDasharray="4 4" />
                            <XAxis dataKey="name" tick={{ fontSize: 12, fill: "#777487" }} tickLine={false} axisLine={{ stroke: "#e8e3ef" }} />
                            <YAxis tick={{ fontSize: 12, fill: "#777487" }} tickLine={false} axisLine={false} domain={[(dataMin: number) => Math.min(0, dataMin), (dataMax: number) => dataMax === 0 ? 1 : dataMax * 1.3]} />
                            <Tooltip
                              formatter={(value: number | string | ReadonlyArray<number | string> | undefined, name: string | number | undefined) => {
                                const numeric = Number(Array.isArray(value) ? value[0] : value ?? 0);
                                return [numeric.toFixed(2), String(name ?? "")] as [string, string];
                              }}
                            />
                            <Legend />
                            <Line type="monotone" dataKey="roas" name="ROAS" stroke="#7c4dff" strokeWidth={3} dot={{ r: 4 }} activeDot={{ r: 5 }} />
                            <Line type="monotone" dataKey="cpi" name="CPI" stroke="#c84dad" strokeWidth={3} dot={{ r: 4 }} activeDot={{ r: 5 }} />
                            <Line type="monotone" dataKey="uv" name="UV" stroke="#5367c8" strokeWidth={3} dot={{ r: 4 }} activeDot={{ r: 5 }} />
                          </LineChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                  ) : (
                    <div className="mb-5 rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-800">No previous week available. The weekly comparison is intentionally not fabricated.</div>
                  )}

                  <div className="mb-5 rounded-xl border border-slate-200 bg-slate-50 p-4">
                    <h3 className="mb-2 text-lg font-semibold text-slate-900">Conclusion</h3>
                    <div className={`inline-flex items-center rounded-lg border px-3 py-1 text-sm font-semibold ${conclusionTone}`}>{driverConclusion.title}</div>
                    <p className="mt-3 text-sm leading-6 text-slate-700">{renderImportantNumbers(driverConclusion.summary)}</p>
                  </div>

                  <div className="rounded-xl border border-slate-200 bg-white p-4">
                    <h3 className="mb-3 text-lg font-semibold text-slate-900">Analysis</h3>
                    <div className="space-y-4">
                      <div className="rounded-xl border border-slate-200 bg-slate-50 p-4">
                        <h4 className="mb-2 text-lg font-semibold text-violet-800">ROAS Performance</h4>
                        {previousWeek ? (
                          <ul className="list-disc space-y-2 pl-5 text-[15px] leading-7 text-slate-700">
                            <li>Current ROAS was <strong>{selectedWeek.roas === null ? "N/A" : formatRatio(selectedWeek.roas)}</strong>, versus <strong>{previousWeek.roas === null ? "N/A" : formatRatio(previousWeek.roas)}</strong> in the previous week.</li>
                            <li>Absolute change was <strong>{selectedWeek.roas !== null && previousWeek.roas !== null ? `${(selectedWeek.roas - previousWeek.roas).toFixed(2)}` : "N/A"}</strong>, or <strong>{selectedWeek.roas !== null && previousWeek.roas !== null && previousWeek.roas !== 0 ? `${(((selectedWeek.roas - previousWeek.roas) / previousWeek.roas) * 100).toFixed(1)}%` : "N/A"}</strong> WoW.</li>
                            <li>{selectedWeek.roas !== null && previousWeek.roas !== null && selectedWeek.roas > previousWeek.roas ? `ROAS increased from ${formatRatio(previousWeek.roas)} to ${formatRatio(selectedWeek.roas)} (+${(((selectedWeek.roas - previousWeek.roas) / previousWeek.roas) * 100).toFixed(1)}%), indicating a positive WoW movement.` : selectedWeek.roas !== null && previousWeek.roas !== null && selectedWeek.roas < previousWeek.roas ? `ROAS decreased from ${formatRatio(previousWeek.roas)} to ${formatRatio(selectedWeek.roas)} (${(((selectedWeek.roas - previousWeek.roas) / previousWeek.roas) * 100).toFixed(1)}%), indicating a negative WoW movement.` : "ROAS was effectively stable versus the previous week."}</li>
                            <li>{selectedWeek.roas !== null && selectedWeek.roas >= 1 ? "The campaign is currently above the 1.00 break-even threshold and meaningfully profitable." : "The campaign remains below the 1.00 break-even threshold and needs stronger monetization efficiency to become sustainable."}</li>
                          </ul>
                        ) : (
                          <p className="text-[15px] leading-7 text-slate-700">No previous week is available for ROAS comparison, so the current performance is assessed in isolation.</p>
                        )}
                      </div>

                      <div className="rounded-xl border border-slate-200 bg-slate-50 p-4">
                        <h4 className="mb-2 text-lg font-semibold text-violet-800">Spend & Scale</h4>
                        {previousWeek ? (
                          <ul className="list-disc space-y-2 pl-5 text-[15px] leading-7 text-slate-700">
                            <li>Cost {compareNumericChange(selectedWeek.cost, previousWeek.cost).direction === "increased" ? `increased from ${formatCurrency(previousWeek.cost)} to ${formatCurrency(selectedWeek.cost)} (+${Math.abs(((selectedWeek.cost - previousWeek.cost) / previousWeek.cost) * 100).toFixed(1)}%)` : compareNumericChange(selectedWeek.cost, previousWeek.cost).direction === "decreased" ? `decreased from ${formatCurrency(previousWeek.cost)} to ${formatCurrency(selectedWeek.cost)} (${(((selectedWeek.cost - previousWeek.cost) / previousWeek.cost) * 100).toFixed(1)}%)` : `remained essentially flat at ${formatCurrency(selectedWeek.cost)}.`}</li>
                            <li>Impressions {compareNumericChange(selectedWeek.impressions, previousWeek.impressions).direction === "increased" ? `rose from ${formatNumber(previousWeek.impressions)} to ${formatNumber(selectedWeek.impressions)} (+${Math.abs(((selectedWeek.impressions - previousWeek.impressions) / previousWeek.impressions) * 100).toFixed(1)}%)` : compareNumericChange(selectedWeek.impressions, previousWeek.impressions).direction === "decreased" ? `fell from ${formatNumber(previousWeek.impressions)} to ${formatNumber(selectedWeek.impressions)} (${(((selectedWeek.impressions - previousWeek.impressions) / previousWeek.impressions) * 100).toFixed(1)}%)` : `were effectively stable at ${formatNumber(selectedWeek.impressions)}.`}</li>
                            <li>Clicks {compareNumericChange(selectedWeek.clicks, previousWeek.clicks).direction === "increased" ? `increased from ${formatNumber(previousWeek.clicks)} to ${formatNumber(selectedWeek.clicks)} (+${Math.abs(((selectedWeek.clicks - previousWeek.clicks) / previousWeek.clicks) * 100).toFixed(1)}%)` : compareNumericChange(selectedWeek.clicks, previousWeek.clicks).direction === "decreased" ? `declined from ${formatNumber(previousWeek.clicks)} to ${formatNumber(selectedWeek.clicks)} (${(((selectedWeek.clicks - previousWeek.clicks) / previousWeek.clicks) * 100).toFixed(1)}%)` : `were effectively stable at ${formatNumber(selectedWeek.clicks)}.`}</li>
                            <li>Installs {compareNumericChange(selectedWeek.installs, previousWeek.installs).direction === "increased" ? `increased from ${formatNumber(previousWeek.installs)} to ${formatNumber(selectedWeek.installs)} (+${Math.abs(((selectedWeek.installs - previousWeek.installs) / previousWeek.installs) * 100).toFixed(1)}%)` : compareNumericChange(selectedWeek.installs, previousWeek.installs).direction === "decreased" ? `fell from ${formatNumber(previousWeek.installs)} to ${formatNumber(selectedWeek.installs)} (${(((selectedWeek.installs - previousWeek.installs) / previousWeek.installs) * 100).toFixed(1)}%)` : `were effectively stable at ${formatNumber(selectedWeek.installs)}.`}</li>
                          </ul>
                        ) : (
                          <p className="text-[15px] leading-7 text-slate-700">No previous week is available, so the spend-scaling pattern cannot be benchmarked against a historical baseline.</p>
                        )}
                      </div>

                      <div className="rounded-xl border border-slate-200 bg-slate-50 p-4">
                        <h4 className="mb-2 text-lg font-semibold text-violet-800">Acquisition Efficiency</h4>
                        {previousWeek ? (
                          <ul className="list-disc space-y-2 pl-5 text-[15px] leading-7 text-slate-700">
                            <li>CPI {selectedWeek.cpi !== null && previousWeek.cpi !== null ? `${compareNumericChange(selectedWeek.cpi, previousWeek.cpi).direction === "increased" ? "increased" : compareNumericChange(selectedWeek.cpi, previousWeek.cpi).direction === "decreased" ? "decreased" : "remained effectively flat"} from ${formatRatio(previousWeek.cpi)} to ${formatRatio(selectedWeek.cpi)} (${selectedWeek.cpi !== previousWeek.cpi ? `${(((selectedWeek.cpi - previousWeek.cpi) / previousWeek.cpi) * 100).toFixed(1)}%` : "0.0%"})` : "is not available"}.</li>
                            <li>CTR {selectedWeek.ctr !== null && previousWeek.ctr !== null ? `${compareNumericChange(selectedWeek.ctr, previousWeek.ctr).direction === "increased" ? "improved" : compareNumericChange(selectedWeek.ctr, previousWeek.ctr).direction === "decreased" ? "weakened" : "held steady"} from ${formatRatio(previousWeek.ctr)} to ${formatRatio(selectedWeek.ctr)} (${selectedWeek.ctr !== previousWeek.ctr ? `${(((selectedWeek.ctr - previousWeek.ctr) / previousWeek.ctr) * 100).toFixed(1)}%` : "0.0%"})` : "is not available"}.</li>
                            <li>CVR {selectedWeek.cvr !== null && previousWeek.cvr !== null ? `${compareNumericChange(selectedWeek.cvr, previousWeek.cvr).direction === "increased" ? "improved" : compareNumericChange(selectedWeek.cvr, previousWeek.cvr).direction === "decreased" ? "softened" : "held steady"} from ${formatRatio(previousWeek.cvr)} to ${formatRatio(selectedWeek.cvr)} (${selectedWeek.cvr !== previousWeek.cvr ? `${(((selectedWeek.cvr - previousWeek.cvr) / previousWeek.cvr) * 100).toFixed(1)}%` : "0.0%"})` : "is not available"}.</li>
                          </ul>
                        ) : (
                          <p className="text-[15px] leading-7 text-slate-700">No previous week is available, so acquisition efficiency cannot be compared with a historical baseline.</p>
                        )}
                      </div>

                      <div className="rounded-xl border border-slate-200 bg-slate-50 p-4">
                        <h4 className="mb-2 text-lg font-semibold text-violet-800">User Value</h4>
                        {previousWeek ? (
                          <ul className="list-disc space-y-2 pl-5 text-[15px] leading-7 text-slate-700">
                            <li>UV {selectedWeek.uv !== null && previousWeek.uv !== null ? `${compareNumericChange(selectedWeek.uv, previousWeek.uv).direction === "increased" ? "increased" : compareNumericChange(selectedWeek.uv, previousWeek.uv).direction === "decreased" ? "declined" : "held steady"} from ${formatRatio(previousWeek.uv)} to ${formatRatio(selectedWeek.uv)} (${selectedWeek.uv !== previousWeek.uv ? `${(((selectedWeek.uv - previousWeek.uv) / previousWeek.uv) * 100).toFixed(1)}%` : "0.0%"})` : "is not available"}.</li>
                            <li>Revenue moved from <strong>{formatCurrency(previousWeek.revenue)}</strong> to <strong>{formatCurrency(selectedWeek.revenue)}</strong>, which is a {selectedWeek.revenue >= previousWeek.revenue ? "positive" : "negative"} change of <strong>{previousWeek.revenue === 0 ? "N/A" : `${(((selectedWeek.revenue - previousWeek.revenue) / previousWeek.revenue) * 100).toFixed(1)}%`}</strong>.</li>
                            <li>Revenue per install (UV) is the key user-value signal here: {selectedWeek.uv !== null && previousWeek.uv !== null ? `it ${compareNumericChange(selectedWeek.uv, previousWeek.uv).direction === "increased" ? "supported" : compareNumericChange(selectedWeek.uv, previousWeek.uv).direction === "decreased" ? "compressed" : "stayed stable in"} the ROAS outcome.` : "not enough data is available to infer the effect."}</li>
                          </ul>
                        ) : (
                          <p className="text-[15px] leading-7 text-slate-700">No previous week is available, so user-value comparison is not yet benchmarked against a baseline.</p>
                        )}
                      </div>

                      <div className="rounded-xl border border-slate-200 bg-slate-50 p-4">
                        <h4 className="mb-2 text-lg font-semibold text-violet-800">Driver Diagnosis</h4>
                        {previousWeek ? (
                          <div className="space-y-2 text-[15px] leading-7 text-slate-700">
                            <p>{selectedWeek.roas !== null && previousWeek.roas !== null ? `ROAS ${selectedWeek.roas > previousWeek.roas ? "increased" : selectedWeek.roas < previousWeek.roas ? "decreased" : "was stable"} from ${formatRatio(previousWeek.roas)} to ${formatRatio(selectedWeek.roas)} (${selectedWeek.roas !== previousWeek.roas ? `${(((selectedWeek.roas - previousWeek.roas) / previousWeek.roas) * 100).toFixed(1)}%` : "0.0%"}), which is the primary driver of the performance shift.` : "ROAS comparison is not available for this period."}</p>
                            <p>{selectedWeek.cpi !== null && previousWeek.cpi !== null && selectedWeek.uv !== null && previousWeek.uv !== null ? `The underlying mix shows CPI ${selectedWeek.cpi > previousWeek.cpi ? "increasing" : selectedWeek.cpi < previousWeek.cpi ? "declining" : "stable"} from ${formatRatio(previousWeek.cpi)} to ${formatRatio(selectedWeek.cpi)}, while UV ${selectedWeek.uv > previousWeek.uv ? "improved" : selectedWeek.uv < previousWeek.uv ? "softened" : "held steady"} from ${formatRatio(previousWeek.uv)} to ${formatRatio(selectedWeek.uv)}. ${selectedWeek.roas !== null && previousWeek.roas !== null && selectedWeek.roas > previousWeek.roas ? "This combination supported the ROAS lift." : "This combination explains the ROAS movement."}` : "CPI and UV comparison is not available for this period."}</p>
                            <p>{selectedWeek.noteLimited && selectedWeek.uv !== null && previousWeek.uv !== null && selectedWeek.uv < previousWeek.uv ? 'The note "Limited by target" was present while UV declined, which is consistent with the target-constrained ROAS pattern.' : selectedWeek.uv !== null && previousWeek.uv !== null && selectedWeek.uv < previousWeek.uv ? "UV declined without a target constraint, which indicates lower-value user quality is compressing monetization efficiency." : "The observed ROAS change is not being driven by a material UV decline, so the efficiency mix remains the most important factor to monitor."}</p>
                          </div>
                        ) : (
                          <p className="text-[15px] leading-7 text-slate-700">There is no previous-week comparison, so driver diagnosis is limited to the current week only.</p>
                        )}
                      </div>
                    </div>

                    <div className="mt-5 rounded-xl border border-violet-200 bg-violet-50 p-4">
                      <h4 className="mb-3 text-xl font-semibold text-violet-800">Campaign Recommendation</h4>
                      <p className="mb-3 text-base font-semibold text-violet-900">{campaignRecommendation.title}</p>
                      <ol className="list-decimal space-y-2 pl-5 text-[15px] leading-6 text-slate-700">
                        {campaignRecommendation.actions.map((action, index) => (
                          <li key={`${action}-${index}`}>{action}</li>
                        ))}
                      </ol>
                    </div>
                  </div>
                </>
              ) : (
                <div className="rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-800">No data is available for the selected campaign and week.</div>
              )}
            </section>

            <section className="dashboard-section mb-7">
              <div className="mb-4">
                <p className="text-[11px] font-semibold uppercase tracking-[0.16em] text-slate-500">5. Campaign Level's Recommendation</p>
                <h2 className="mt-1 text-2xl font-semibold text-slate-900">Campaign Level's Recommendation</h2>
              </div>

              <div className="mb-5 grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
                <div className="rounded-2xl border border-violet-200 bg-violet-50 p-4">
                  <p className="text-xs font-semibold uppercase tracking-[0.12em] text-violet-700">Total Budget</p>
                  <div className="mt-2 text-2xl font-semibold text-slate-900">{totalBudget === null ? "N/A" : formatCurrency(totalBudget)}</div>
                </div>
                <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
                  <p className="text-xs font-semibold uppercase tracking-[0.12em] text-slate-600">Total Spend</p>
                  <div className="mt-2 text-2xl font-semibold text-slate-900">{formatCurrency(totalSpend)}</div>
                </div>
                <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
                  <p className="text-xs font-semibold uppercase tracking-[0.12em] text-slate-600">Remaining Budget</p>
                  <div className="mt-2 text-2xl font-semibold text-slate-900">{formatCurrency(remainingBudget)}</div>
                </div>
                <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
                  <p className="text-xs font-semibold uppercase tracking-[0.12em] text-slate-600">Spend Rate</p>
                  <div className="mt-2 text-2xl font-semibold text-slate-900">{spendRate === null ? "N/A" : `${(spendRate * 100).toFixed(1)}%`}</div>
                </div>
              </div>

              <div className="overflow-x-auto rounded-xl border border-slate-200 bg-white">
                <table className="min-w-full text-left text-sm">
                  <thead className="bg-violet-50/90">
                    <tr>
                      <th className="px-3 py-3 text-left text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-600">Campaign</th>
                      <th className="px-3 py-3 text-right text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-600">Current ROAS</th>
                      <th className="px-3 py-3 text-right text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-600">WoW ROAS</th>
                      <th className="px-3 py-3 text-right text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-600">Current CPI</th>
                      <th className="px-3 py-3 text-right text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-600">Current UV</th>
                      <th className="px-3 py-3 text-right text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-600">Current Spend</th>
                      <th className="px-3 py-3 text-right text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-600">Budget Action</th>
                      <th className="px-3 py-3 text-left text-[11px] font-semibold uppercase tracking-[0.12em] text-slate-600">Reason</th>
                    </tr>
                  </thead>
                  <tbody>
                    {campaignRecommendationRows.map((row) => (
                      <tr key={row.campaignName} className="border-t border-slate-100 hover:bg-violet-50/40">
                        <td className="px-3 py-3 font-medium text-slate-900">{row.campaignName}</td>
                        <td className="px-3 py-3 text-right font-semibold text-slate-900">{row.currentRoas === null ? "N/A" : formatRatio(row.currentRoas)}</td>
                        <td className="px-3 py-3 text-right">{row.previousRoas === null ? "No WoW comparison available" : row.wowRoas === null ? "N/A" : `${row.wowRoas >= 0 ? "+" : ""}${row.wowRoas.toFixed(1)}%`}</td>
                        <td className="px-3 py-3 text-right">{row.currentCpi === null ? "N/A" : formatRatio(row.currentCpi)}</td>
                        <td className="px-3 py-3 text-right">{row.currentUv === null ? "N/A" : formatRatio(row.currentUv)}</td>
                        <td className="px-3 py-3 text-right">{formatCurrency(row.currentSpend)}</td>
                        <td className="px-3 py-3 text-right font-semibold text-violet-700">{row.action}</td>
                        <td className="px-3 py-3 text-left text-slate-700">{row.reason}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {totalBudget !== null && totalSpend > totalBudget ? (
                <div className="mt-4 rounded-xl border border-rose-200 bg-rose-50 p-4 text-sm text-rose-700">Total spend exceeds the total budget. The current portfolio has already surpassed the available budget.</div>
              ) : (
                <div className="mt-4 rounded-xl border border-violet-100 bg-violet-50 p-4 text-sm text-violet-700">{remainingBudget > 0 ? `There is ${formatCurrency(remainingBudget)} remaining in the total budget, so improving campaigns can be considered for additional spend when supported by ROAS and UV performance.` : `Remaining budget is limited, so broad portfolio expansion is not recommended without clear performance evidence.`}</div>
              )}
            </section>
          </>
        ) : null}
      </div>
    </div>
  );
}
