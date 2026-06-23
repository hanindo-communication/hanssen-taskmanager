import {
  parseTiktokAdsProductCampaignMatrix,
  type TiktokAdsCampaignAggregate,
  type TiktokAdsReportSummary,
} from './tiktok-ads-campaign-xlsx';

export type TiktokAdsWeekReport = {
  weekIndex: number;
  label: string;
  startDate: string;
  endDate: string;
  sourceFiles: string[];
  summary: TiktokAdsReportSummary;
};

export type TiktokAdsWeeklyBundle = {
  weeks: TiktokAdsWeekReport[];
  warnings: string[];
};

export type WowMetricKey = 'cost' | 'grossRevenue' | 'roi';

export type WowCampaignRow = {
  campaignName: string;
  weeks: Array<{
    cost: number | null;
    grossRevenue: number | null;
    roi: number | null;
  }>;
};

export type WowTotalsRow = {
  cost: number;
  grossRevenue: number;
  roi: number | null;
};

export type ParsedFileInput = {
  fileName: string;
  matrix: string[][];
  uploadOrder: number;
};

const FILENAME_RANGE_RE = /(\d{4}-\d{2}-\d{2})\s*-\s*(\d{4}-\d{2}-\d{2})/i;

function parseIsoDate(s: string): Date | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s.trim());
  if (!m) return null;
  const d = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
  return Number.isNaN(d.getTime()) ? null : d;
}

export function parseDateRangeFromFilename(fileName: string): { start: string; end: string } | null {
  const m = FILENAME_RANGE_RE.exec(fileName);
  if (!m) return null;
  const start = parseIsoDate(m[1]);
  const end = parseIsoDate(m[2]);
  if (!start || !end || start > end) return null;
  return { start: m[1], end: m[2] };
}

function formatDateLabel(iso: string): string {
  const d = parseIsoDate(iso);
  if (!d) return iso;
  return d.toLocaleDateString('id-ID', {
    day: 'numeric',
    month: 'short',
    year: 'numeric',
    timeZone: 'UTC',
  });
}

function stripExtension(fileName: string): string {
  return fileName.replace(/\.xlsx$/i, '');
}

export function formatReportLabel(index: number, fileName: string): string {
  const range = parseDateRangeFromFilename(fileName);
  if (range) {
    return `Laporan ${index} · ${formatDateLabel(range.start)} – ${formatDateLabel(range.end)}`;
  }
  return `Laporan ${index} · ${stripExtension(fileName)}`;
}

function findCampaign(
  campaigns: TiktokAdsCampaignAggregate[],
  name: string
): TiktokAdsCampaignAggregate | undefined {
  return campaigns.find((c) => c.campaignName === name);
}

export function buildWowRows(weeks: TiktokAdsWeekReport[]): {
  campaigns: WowCampaignRow[];
  totals: WowTotalsRow[];
} {
  const nameSet = new Set<string>();
  for (const w of weeks) {
    for (const c of w.summary.campaigns) {
      if (c.cost > 0) nameSet.add(c.campaignName);
    }
  }

  const campaignNames = [...nameSet].sort((a, b) => a.localeCompare(b, 'id-ID'));

  const campaigns: WowCampaignRow[] = campaignNames.map((campaignName) => ({
    campaignName,
    weeks: weeks.map((w) => {
      const c = findCampaign(w.summary.campaigns, campaignName);
      if (!c || c.cost <= 0) {
        return { cost: null, grossRevenue: null, roi: null };
      }
      return { cost: c.cost, grossRevenue: c.grossRevenue, roi: c.roi };
    }),
  }));

  const totals: WowTotalsRow[] = weeks.map((w) => ({
    cost: w.summary.totalCost,
    grossRevenue: w.summary.totalGrossRevenue,
    roi: w.summary.blendedRoi,
  }));

  return { campaigns, totals };
}

export function calcDeltaPct(current: number | null, previous: number | null): number | null {
  if (current === null || previous === null || previous === 0) return null;
  return ((current - previous) / previous) * 100;
}

type DraftReport = {
  fileName: string;
  uploadOrder: number;
  summary: TiktokAdsReportSummary;
  startDate: string;
  endDate: string;
  hasDateRange: boolean;
};

function sortDrafts(drafts: DraftReport[]): DraftReport[] {
  return [...drafts].sort((a, b) => {
    if (a.hasDateRange && b.hasDateRange) return a.startDate.localeCompare(b.startDate);
    if (a.hasDateRange && !b.hasDateRange) return -1;
    if (!a.hasDateRange && b.hasDateRange) return 1;
    return a.uploadOrder - b.uploadOrder;
  });
}

function draftsToWeeks(drafts: DraftReport[]): TiktokAdsWeekReport[] {
  const sorted = sortDrafts(drafts);
  return sorted.map((d, i) => {
    const index = i + 1;
    return {
      weekIndex: index,
      label: formatReportLabel(index, d.fileName),
      startDate: d.hasDateRange ? d.startDate : `laporan-${index}`,
      endDate: d.hasDateRange ? d.endDate : `laporan-${index}`,
      sourceFiles: [d.fileName],
      summary: d.summary,
    };
  });
}

function parseFilesToDrafts(files: ParsedFileInput[], warnings: string[]): DraftReport[] {
  const drafts: DraftReport[] = [];

  for (const f of files) {
    const { summary, error } = parseTiktokAdsProductCampaignMatrix(f.matrix);
    if (error || !summary) {
      warnings.push(`"${f.fileName}": ${error ?? 'Gagal memproses.'}`);
      continue;
    }

    const range = parseDateRangeFromFilename(f.fileName);
    drafts.push({
      fileName: f.fileName,
      uploadOrder: f.uploadOrder,
      summary,
      startDate: range?.start ?? '',
      endDate: range?.end ?? '',
      hasDateRange: range !== null,
    });
  }

  return drafts;
}

/** Satu file XLSX = satu laporan. Sort kronologis by tanggal di nama file. */
export function buildWeeklyBundle(files: ParsedFileInput[]): TiktokAdsWeeklyBundle {
  const warnings: string[] = [];
  if (files.length === 0) {
    return { weeks: [], warnings: ['Tidak ada file dipilih.'] };
  }

  const drafts = parseFilesToDrafts(files, warnings);
  return { weeks: draftsToWeeks(drafts), warnings };
}

/** Append file baru ke session existing; skip duplikat nama file. */
export function appendToWeeklyBundle(
  existing: ParsedFileInput[],
  incoming: ParsedFileInput[]
): { bundle: TiktokAdsWeeklyBundle; parsedFiles: ParsedFileInput[]; addedFileNames: string[] } {
  const warnings: string[] = [];
  const existingNames = new Set(existing.map((f) => f.fileName.toLowerCase()));
  const added: ParsedFileInput[] = [];
  const addedFileNames: string[] = [];

  for (const f of incoming) {
    const key = f.fileName.toLowerCase();
    if (existingNames.has(key)) {
      warnings.push(`"${f.fileName}": duplikat — file sudah ada, dilewati.`);
      continue;
    }
    existingNames.add(key);
    added.push(f);
    addedFileNames.push(f.fileName);
  }

  const merged = [...existing, ...added];
  if (merged.length === 0) {
    return {
      bundle: { weeks: [], warnings: [...warnings, 'Tidak ada file dipilih.'] },
      parsedFiles: [],
      addedFileNames: [],
    };
  }

  const drafts = parseFilesToDrafts(merged, warnings);
  return {
    bundle: { weeks: draftsToWeeks(drafts), warnings },
    parsedFiles: merged,
    addedFileNames,
  };
}
