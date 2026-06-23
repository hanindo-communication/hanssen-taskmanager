/** TikTok Ads "Product campaign data" XLSX export — agregasi per Campaign name, dengan dukungan kolom Date. */

export type TiktokAdsCampaignAggregate = {
  campaignName: string;
  cost: number;
  netCost: number;
  grossRevenue: number;
  skuOrders: number;
  /** Daily budget cap TikTok (MAX per nama kampanye saat merge baris). */
  currentBudget: number;
  roi: number | null;
  costPerOrder: number | null;
};

export type TiktokAdsReportSummary = {
  rowCount: number;
  campaignCount: number;
  totalCost: number;
  totalNetCost: number;
  totalGrossRevenue: number;
  totalSkuOrders: number;
  /** Jumlah MAX daily budget per kampanye (bukan spending). */
  sumOfDailyBudgetCaps: number;
  blendedRoi: number | null;
  blendedCostPerOrder: number | null;
  campaigns: TiktokAdsCampaignAggregate[];
};

export type TiktokAdsDailyRow = {
  date: string;
  sourceFile: string;
  campaignName: string;
  cost: number;
  netCost: number;
  grossRevenue: number;
  skuOrders: number;
  currentBudget: number;
};

const DATE_HEADER_CANDIDATES = ['date', 'day', 'reporting date', 'report date'];

function normalizeHeader(h: string): string {
  return h.trim().toLowerCase();
}

function parseNumber(raw: string | number | undefined | null): number {
  if (raw === undefined || raw === null || raw === '') return 0;
  if (typeof raw === 'number') return Number.isFinite(raw) ? raw : 0;
  const s = String(raw).trim().replace(/,/g, '');
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : 0;
}

function getField(row: Record<string, string>, ...candidates: string[]): string {
  const keys = Object.keys(row);
  for (const c of candidates) {
    const found = keys.find((k) => normalizeHeader(k) === normalizeHeader(c));
    if (found !== undefined) return String(row[found] ?? '').trim();
  }
  return '';
}

function matrixToObjects(matrix: string[][]): Record<string, string>[] {
  if (matrix.length < 2) return [];
  const headerRow = matrix[0] ?? [];
  const keys = headerRow.map((h, i) => {
    const t = String(h ?? '').trim();
    return t || `__col_${i}`;
  });
  const out: Record<string, string>[] = [];
  for (let r = 1; r < matrix.length; r++) {
    const row = matrix[r] ?? [];
    const obj: Record<string, string> = {};
    let any = false;
    for (let c = 0; c < keys.length; c++) {
      const v = row[c];
      const s = v === undefined || v === null ? '' : String(v).trim();
      obj[keys[c]] = s;
      if (s) any = true;
    }
    if (any) out.push(obj);
  }
  return out;
}

function headersFromMatrix(matrix: string[][]): string[] {
  return (matrix[0] ?? []).map((h) => normalizeHeader(String(h ?? '')));
}

function findDateHeaderKey(headers: string[]): string | null {
  for (const candidate of DATE_HEADER_CANDIDATES) {
    const found = headers.find((h) => h === candidate);
    if (found) return found;
  }
  return null;
}

function parseDateValue(raw: string): string | null {
  const s = raw.trim();
  if (!s) return null;

  const iso = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;

  const dmy = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/.exec(s);
  if (dmy) {
    const dd = dmy[1].padStart(2, '0');
    const mm = dmy[2].padStart(2, '0');
    return `${dmy[3]}-${mm}-${dd}`;
  }

  const parsed = new Date(s);
  if (!Number.isNaN(parsed.getTime())) {
    return parsed.toISOString().slice(0, 10);
  }

  return null;
}

export function matrixHasDateColumn(matrix: string[][]): boolean {
  const headers = headersFromMatrix(matrix);
  return findDateHeaderKey(headers) !== null;
}

type Agg = {
  cost: number;
  netCost: number;
  grossRevenue: number;
  skuOrders: number;
  budgetMax: number;
};

function mergeIntoAgg(map: Map<string, Agg>, name: string, row: Agg): void {
  const prev = map.get(name);
  if (!prev) {
    map.set(name, { ...row });
  } else {
    prev.cost += row.cost;
    prev.netCost += row.netCost;
    prev.grossRevenue += row.grossRevenue;
    prev.skuOrders += row.skuOrders;
    prev.budgetMax = Math.max(prev.budgetMax, row.budgetMax);
  }
}

function aggMapToSummary(byName: Map<string, Agg>, rowCount: number): TiktokAdsReportSummary | null {
  if (byName.size === 0) return null;

  const campaigns: TiktokAdsCampaignAggregate[] = [];
  let totalCost = 0;
  let totalNetCost = 0;
  let totalGrossRevenue = 0;
  let totalSkuOrders = 0;
  let sumBudgetCaps = 0;

  for (const [campaignName, a] of byName) {
    totalCost += a.cost;
    totalNetCost += a.netCost;
    totalGrossRevenue += a.grossRevenue;
    totalSkuOrders += a.skuOrders;
    sumBudgetCaps += a.budgetMax;

    const roi = a.cost > 0 ? a.grossRevenue / a.cost : null;
    const costPerOrder = a.skuOrders > 0 ? a.cost / a.skuOrders : null;

    campaigns.push({
      campaignName,
      cost: a.cost,
      netCost: a.netCost,
      grossRevenue: a.grossRevenue,
      skuOrders: a.skuOrders,
      currentBudget: a.budgetMax,
      roi,
      costPerOrder,
    });
  }

  campaigns.sort((x, y) => y.cost - x.cost);

  const blendedRoi = totalCost > 0 ? totalGrossRevenue / totalCost : null;
  const blendedCostPerOrder = totalSkuOrders > 0 ? totalCost / totalSkuOrders : null;

  return {
    rowCount,
    campaignCount: campaigns.length,
    totalCost,
    totalNetCost,
    totalGrossRevenue,
    totalSkuOrders,
    sumOfDailyBudgetCaps: sumBudgetCaps,
    blendedRoi,
    blendedCostPerOrder,
    campaigns,
  };
}

export function aggregateCampaignRows(
  rows: Array<{
    campaignName: string;
    cost: number;
    netCost: number;
    grossRevenue: number;
    skuOrders: number;
    currentBudget: number;
  }>
): TiktokAdsReportSummary {
  const byName = new Map<string, Agg>();
  for (const r of rows) {
    if (!r.campaignName) continue;
    mergeIntoAgg(byName, r.campaignName, {
      cost: r.cost,
      netCost: r.netCost,
      grossRevenue: r.grossRevenue,
      skuOrders: r.skuOrders,
      budgetMax: r.currentBudget,
    });
  }
  return aggMapToSummary(byName, rows.length) ?? {
    rowCount: 0,
    campaignCount: 0,
    totalCost: 0,
    totalNetCost: 0,
    totalGrossRevenue: 0,
    totalSkuOrders: 0,
    sumOfDailyBudgetCaps: 0,
    blendedRoi: null,
    blendedCostPerOrder: null,
    campaigns: [],
  };
}

export function parseTiktokAdsDailyRows(matrix: string[][], sourceFile = ''): TiktokAdsDailyRow[] {
  if (!matrix.length || !(matrix[0]?.length)) return [];

  const headers = headersFromMatrix(matrix);
  const dateHeader = findDateHeaderKey(headers);
  if (!dateHeader) return [];

  const objects = matrixToObjects(matrix);
  const out: TiktokAdsDailyRow[] = [];

  for (const row of objects) {
    const dateRaw = getField(row, 'Date', 'Day', 'Reporting date', 'Report date');
    const date = parseDateValue(dateRaw);
    const name = getField(row, 'Campaign name');
    if (!date || !name) continue;

    out.push({
      date,
      sourceFile,
      campaignName: name,
      cost: parseNumber(getField(row, 'Cost')),
      netCost: parseNumber(getField(row, 'Net Cost')),
      grossRevenue: parseNumber(getField(row, 'Gross revenue')),
      skuOrders: parseNumber(getField(row, 'SKU orders')),
      currentBudget: parseNumber(getField(row, 'Current budget')),
    });
  }

  return out;
}

export function parseTiktokAdsProductCampaignMatrix(matrix: string[][]): {
  summary: TiktokAdsReportSummary | null;
  error?: string;
} {
  if (!matrix.length || !(matrix[0]?.length)) {
    return { summary: null, error: 'Lembar kosong.' };
  }

  const headers = headersFromMatrix(matrix);
  const needName = headers.some((h) => h === 'campaign name');
  const hasCost = headers.some((h) => h === 'cost');
  const hasNetCost = headers.some((h) => h === 'net cost');
  const hasGmv = headers.some((h) => h === 'gross revenue');

  if (!needName) {
    return {
      summary: null,
      error:
        'Format tidak dikenali. Pastikan ekspor TikTok berisi kolom Campaign name (Product campaign data).',
    };
  }
  if (!hasCost && !hasNetCost) {
    return {
      summary: null,
      error: 'Kolom Cost atau Net Cost tidak ditemukan.',
    };
  }
  if (!hasGmv) {
    return {
      summary: null,
      error: 'Kolom Gross revenue tidak ditemukan.',
    };
  }

  const objects = matrixToObjects(matrix);
  const byName = new Map<string, Agg>();

  for (const row of objects) {
    const name = getField(row, 'Campaign name');
    if (!name) continue;

    mergeIntoAgg(byName, name, {
      cost: parseNumber(getField(row, 'Cost')),
      netCost: parseNumber(getField(row, 'Net Cost')),
      grossRevenue: parseNumber(getField(row, 'Gross revenue')),
      skuOrders: parseNumber(getField(row, 'SKU orders')),
      budgetMax: parseNumber(getField(row, 'Current budget')),
    });
  }

  if (byName.size === 0) {
    return {
      summary: null,
      error: 'Tidak ada baris kampanye dengan Campaign name yang terisi.',
    };
  }

  const summary = aggMapToSummary(byName, objects.length);
  return { summary };
}
