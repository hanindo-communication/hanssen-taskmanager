'use client';

import { useEffect, useMemo, useRef, type RefObject } from 'react';
import boardStyles from '@/components/board/board-client.module.css';
import {
  buildCombinedSelectionSummary,
  type CombinedCampaignAggregate,
  type TiktokAdsWeekReport,
} from '@/lib/report-generator/tiktok-ads-weekly';
import styles from './ReportGeneratorPanel.module.css';

type ThemeVars = {
  paper: string;
  text: string;
  grid: string;
  colors: string[];
  fontFamily: string;
};

type PlotlyModule = typeof import('plotly.js-basic-dist');

const plotlyConfig = { displayModeBar: false, responsive: true } as const;

function formatIdr(n: number): string {
  return new Intl.NumberFormat('id-ID', { maximumFractionDigits: 0 }).format(Math.round(n));
}

function formatRoi(n: number | null): string {
  if (n === null || !Number.isFinite(n)) return '—';
  return n.toLocaleString('id-ID', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatPct(n: number | null): string {
  if (n === null || !Number.isFinite(n)) return '—';
  const sign = n > 0 ? '+' : '';
  return `${sign}${n.toFixed(1)}%`;
}

function shortReportLabel(label: string): string {
  return label.replace(/^Laporan \d+ · /, '');
}

function readThemeFromEl(el: HTMLElement | null): ThemeVars {
  if (!el) {
    return {
      paper: 'rgb(240 242 240 / 0.94)',
      text: '#1e3a5a',
      grid: '#7ec1d9',
      colors: ['#246e7f', '#4a9eb4', '#b49d86', '#936d62', '#1e3a5a'],
      fontFamily: 'Instrument Sans, sans-serif',
    };
  }
  const s = getComputedStyle(el);
  const primary = s.getPropertyValue('--brand-strong').trim() || '#246e7f';
  const brand = s.getPropertyValue('--brand').trim() || '#4a9eb4';
  const accent = s.getPropertyValue('--danger').trim() || '#b49d86';
  const surface = s.getPropertyValue('--panel-rgb').trim()
    ? `rgb(var(--panel-rgb) / 0.94)`
    : 'rgb(240 242 240 / 0.94)';
  const text = s.getPropertyValue('--text-primary').trim() || '#1e3a5a';
  const grid = s.getPropertyValue('--brand-rgb').trim()
    ? 'rgb(var(--brand-rgb) / 0.22)'
    : '#7ec1d9';
  const body =
    s.getPropertyValue('--font-body').trim() ||
    s.getPropertyValue('font-family').trim() ||
    'Instrument Sans, sans-serif';
  return {
    paper: surface,
    text,
    grid,
    colors: [primary, brand, accent, '#3a6ea5', '#9e7b5f'],
    fontFamily: body,
  };
}

function compareNullableNumber(a: number | null, b: number | null): number {
  const left = a === null || !Number.isFinite(a) ? -Infinity : a;
  const right = b === null || !Number.isFinite(b) ? -Infinity : b;
  return right - left;
}

function overallRoi(totalCost: number, totalGrossRevenue: number): number | null {
  return totalCost > 0 ? totalGrossRevenue / totalCost : null;
}

function metricValueForCampaign(
  campaign: CombinedCampaignAggregate,
  metric: 'cost' | 'grossRevenue' | 'roi'
): number | null {
  if (metric === 'cost') return campaign.totalCost;
  if (metric === 'grossRevenue') return campaign.totalGrossRevenue;
  return campaign.blendedRoi;
}

function topCampaignRows(
  campaigns: CombinedCampaignAggregate[],
  metric: 'cost' | 'grossRevenue' | 'roi'
): CombinedCampaignAggregate[] {
  const list = [...campaigns].filter((c) => metric !== 'roi' || c.blendedRoi !== null);
  list.sort((a, b) => {
    const diff = compareNullableNumber(metricValueForCampaign(a, metric), metricValueForCampaign(b, metric));
    if (diff !== 0) return diff;
    return b.totalCost - a.totalCost;
  });
  return list.slice(0, 5);
}

function campaignRanks(campaigns: CombinedCampaignAggregate[]) {
  const byRoi = [...campaigns]
    .filter((c) => c.blendedRoi !== null)
    .sort((a, b) => {
      const diff = compareNullableNumber(a.blendedRoi, b.blendedRoi);
      if (diff !== 0) return diff;
      return b.totalCost - a.totalCost;
    });
  const bySpend = [...campaigns].sort((a, b) => b.totalCost - a.totalCost);

  const roiRankMap = new Map<string, number>();
  byRoi.forEach((c, i) => roiRankMap.set(c.campaignName, i + 1));
  const spendRankMap = new Map<string, number>();
  bySpend.forEach((c, i) => spendRankMap.set(c.campaignName, i + 1));

  return { byRoi, bySpend, roiRankMap, spendRankMap };
}

function recommendationTrendText(delta: number | null): string {
  if (delta === null || !Number.isFinite(delta)) return 'trend belum cukup data';
  return `tren ROI ${formatPct(delta)}`;
}

type Props = {
  weeks: TiktokAdsWeekReport[];
  onEditSelection: () => void;
  onBackToSingle: () => void;
};

export function TiktokAdsCombinedView({ weeks, onEditSelection, onBackToSingle }: Props) {
  const summaryRootRef = useRef<HTMLDivElement>(null);
  const overallCostRef = useRef<HTMLDivElement>(null);
  const overallGmvRef = useRef<HTMLDivElement>(null);
  const overallRoiRef = useRef<HTMLDivElement>(null);
  const campaignCostRef = useRef<HTMLDivElement>(null);
  const campaignGmvRef = useRef<HTMLDivElement>(null);
  const campaignRoiRef = useRef<HTMLDivElement>(null);

  const data = useMemo(() => buildCombinedSelectionSummary(weeks), [weeks]);
  const aggregate = useMemo(() => {
    const totalCost = data.totals.reduce((sum, t) => sum + t.cost, 0);
    const totalGrossRevenue = data.totals.reduce((sum, t) => sum + t.grossRevenue, 0);
    const combinedRoi = overallRoi(totalCost, totalGrossRevenue);
    const selectedLabels = data.reports.map((w) => w.label);
    const { byRoi, bySpend, roiRankMap, spendRankMap } = campaignRanks(data.campaigns);

    const scaleCandidates = byRoi.slice(0, 3).map((campaign) => {
      const roiRank = roiRankMap.get(campaign.campaignName) ?? 0;
      const spendRank = spendRankMap.get(campaign.campaignName) ?? 0;
      return {
        campaign,
        roiRank,
        spendRank,
        recommendation: `Naikkan budget`,
        reason: `ROI rank #${roiRank}, spend rank #${spendRank}${campaign.roiTrend !== null ? `, ${recommendationTrendText(campaign.roiTrend)}` : ''}.`,
      };
    });

    const reduceCandidates = bySpend
      .filter((campaign) => campaign.totalCost > 0)
      .filter((campaign) => {
        if (combinedRoi === null) return true;
        return campaign.blendedRoi === null || campaign.blendedRoi < combinedRoi;
      })
      .sort((a, b) => {
        const roiDiff = compareNullableNumber(a.blendedRoi, b.blendedRoi);
        if (roiDiff !== 0) return roiDiff;
        return b.totalCost - a.totalCost;
      })
      .slice(0, 3)
      .map((campaign) => {
        const roiRank = roiRankMap.get(campaign.campaignName) ?? 0;
        const spendRank = spendRankMap.get(campaign.campaignName) ?? 0;
        return {
          campaign,
          roiRank,
          spendRank,
          recommendation: campaign.blendedRoi === null ? 'Pantau / turunkan budget' : 'Pantau / turunkan budget',
          reason: `Spending rank #${spendRank}, ROI rank #${roiRank}${campaign.roiTrend !== null ? `, ${recommendationTrendText(campaign.roiTrend)}` : ''}.`,
        };
      });

    const bestRoiReport = [...data.totals]
      .filter((t) => t.roi !== null)
      .sort((a, b) => (b.roi ?? 0) - (a.roi ?? 0))[0];
    const topSpendReport = [...data.totals].sort((a, b) => b.cost - a.cost)[0];
    const topGmvReport = [...data.totals].sort((a, b) => b.grossRevenue - a.grossRevenue)[0];

    return {
      totalCost,
      totalGrossRevenue,
      combinedRoi,
      selectedLabels,
      scaleCandidates,
      reduceCandidates,
      bestRoiReport,
      topSpendReport,
      topGmvReport,
      byRoi,
      bySpend,
    };
  }, [data]);

  useEffect(() => {
    let cancelled = false;
    let Plotly: PlotlyModule['default'] | null = null;

    const render = async () => {
      if (!Plotly) {
        const mod = await import('plotly.js-basic-dist');
        Plotly = mod.default ?? mod;
      }
      if (cancelled || !Plotly) return;
      const plotly = Plotly;

      const theme = chartTheme(summaryRootRef.current);
      const reportLabels = data.totals.map((t) => compactReportLabel(t.reportLabel));

      const renderLine = async (
        ref: RefObject<HTMLDivElement>,
        values: Array<number | null>,
        color: string,
        title: string,
        metricLabel: string,
        decimals = 0
      ) => {
        if (!ref.current) return;
        await plotly.react(
          ref.current,
          [
            {
              type: 'scatter',
              mode: 'lines+markers',
              x: reportLabels,
              y: values,
              line: { color, width: 3 },
              marker: { color, size: 9 },
              connectgaps: false,
              hovertemplate: `%{x}<br>${metricLabel}: %{y:${decimals > 0 ? `,.${decimals}f` : ',.0f'}}<extra></extra>`,
            },
          ],
          {
            paper_bgcolor: theme.paper,
            plot_bgcolor: theme.paper,
            margin: { l: 45, r: 10, t: 10, b: 45 },
            title: { text: title, font: { color: theme.text, family: theme.fontFamily, size: 12 } },
            xaxis: { color: theme.text, tickangle: -15 },
            yaxis: { gridcolor: theme.grid, color: theme.text, separatethousands: true },
            font: { color: theme.text, family: theme.fontFamily },
          },
          plotlyConfig
        );
      };

      const renderCampaignChart = async (
        ref: RefObject<HTMLDivElement>,
        rows: CombinedCampaignAggregate[],
        metric: 'cost' | 'grossRevenue' | 'roi',
        title: string,
        metricLabel: string,
        decimals = 0
      ) => {
        if (!ref.current) return;
        const traces = data.reports.map((report, reportIndex) => ({
          type: 'bar' as const,
          orientation: 'h' as const,
            name: compactReportLabel(report.label),
          x: rows.map((row) => {
            const reportRow = row.reports[reportIndex];
            return metric === 'cost'
              ? reportRow.cost ?? 0
              : metric === 'grossRevenue'
                ? reportRow.grossRevenue ?? 0
                : reportRow.roi ?? 0;
          }),
          y: rows.map((row) => row.campaignName),
          marker: { color: theme.colors[reportIndex % theme.colors.length] },
          hovertemplate: `%{y}<br>${report.label}<br>${metricLabel}: %{x:${decimals > 0 ? `,.${decimals}f` : ',.0f'}}<extra></extra>`,
        }));

        await plotly.react(
          ref.current,
          traces,
          {
            barmode: 'group',
            paper_bgcolor: theme.paper,
            plot_bgcolor: theme.paper,
            margin: { l: 150, r: 10, t: 20, b: 35 },
            title: { text: title, font: { color: theme.text, family: theme.fontFamily, size: 12 } },
            xaxis: {
              gridcolor: theme.grid,
              color: theme.text,
              separatethousands: true,
              tickformat: decimals > 0 ? ',.2f' : undefined,
            },
            yaxis: {
              color: theme.text,
              categoryorder: 'array',
              categoryarray: [...rows].map((row) => row.campaignName).reverse(),
              automargin: true,
            },
            legend: { orientation: 'h', y: 1.12, x: 0 },
            font: { color: theme.text, family: theme.fontFamily },
          },
          plotlyConfig
        );
      };

      await renderLine(
        overallCostRef,
        data.totals.map((t) => t.cost),
        theme.colors[0],
        'Spending antar laporan',
        'Spending'
      );
      await renderLine(
        overallGmvRef,
        data.totals.map((t) => t.grossRevenue),
        theme.colors[1],
        'GMV antar laporan',
        'GMV'
      );
      await renderLine(
        overallRoiRef,
        data.totals.map((t) => t.roi),
        theme.colors[2],
        'ROI antar laporan',
        'ROI',
        2
      );

      await renderCampaignChart(
        campaignCostRef,
        topCampaignRows(data.campaigns, 'cost'),
        'cost',
        'Campaign spending comparison',
        'Spending'
      );
      await renderCampaignChart(
        campaignGmvRef,
        topCampaignRows(data.campaigns, 'grossRevenue'),
        'grossRevenue',
        'Campaign GMV comparison',
        'GMV'
      );
      await renderCampaignChart(
        campaignRoiRef,
        topCampaignRows(data.campaigns, 'roi'),
        'roi',
        'Campaign ROI comparison',
        'ROI',
        2
      );
    };

    void render();

    return () => {
      cancelled = true;
    };
  }, [data]);

  useEffect(() => {
    const onResize = () => {
      void import('plotly.js-basic-dist').then((mod) => {
        const Plotly = mod.default ?? mod;
        [overallCostRef, overallGmvRef, overallRoiRef, campaignCostRef, campaignGmvRef, campaignRoiRef].forEach(
          (ref) => {
            if (ref.current) {
              try {
                Plotly.Plots.resize(ref.current);
              } catch {
                /* ignore */
              }
            }
          }
        );
      });
    };

    window.addEventListener('resize', onResize);
    return () => window.removeEventListener('resize', onResize);
  }, []);

  useEffect(() => {
    return () => {
      void import('plotly.js-basic-dist').then((mod) => {
        const Plotly = mod.default ?? mod;
        [overallCostRef, overallGmvRef, overallRoiRef, campaignCostRef, campaignGmvRef, campaignRoiRef].forEach(
          (ref) => {
            if (ref.current) Plotly.purge(ref.current);
          }
        );
      });
    };
  }, []);

  return (
    <div ref={summaryRootRef} className={styles.combinedStack}>
      <section className={`${styles.section} ${styles.combinedHeroSection}`}>
        <div className={styles.combinedHeroHead}>
          <div>
            <p className={boardStyles.heroEyebrow}>Laporan gabungan</p>
            <h3 className={styles.sectionTitle}>View gabungan terpilih</h3>
            <p className={styles.sectionHint}>
              {weeks.length} laporan dipilih · chart di bawah membandingkan Spending, GMV, ROI, lalu
              rekomendasi aksi campaign.
            </p>
          </div>
          <div className={styles.bulkActions}>
            <button
              type="button"
              className={`${styles.printBtn} ${styles.printBtnSecondary}`}
              onClick={onEditSelection}
            >
              Ubah pilihan laporan
            </button>
            <button
              type="button"
              className={`${styles.printBtn} ${styles.printBtnSecondary}`}
              onClick={onBackToSingle}
            >
              Lihat per laporan
            </button>
          </div>
        </div>

        <div className={styles.combinedChips}>
          {data.reports.map((report) => (
            <span key={report.label} className={styles.combinedChip} title={report.label}>
              {report.label}
            </span>
          ))}
        </div>

        <div className={styles.weekHeroMetrics}>
          <div className={styles.infoTile}>
            <p className={styles.infoTileLabel}>Total Cost gabungan</p>
            <p className={styles.infoTileValue}>Rp {formatIdr(aggregate.totalCost)}</p>
          </div>
          <div className={styles.infoTile}>
            <p className={styles.infoTileLabel}>Total GMV gabungan</p>
            <p className={styles.infoTileValue}>Rp {formatIdr(aggregate.totalGrossRevenue)}</p>
          </div>
          <div className={styles.infoTile}>
            <p className={styles.infoTileLabel}>ROI gabungan</p>
            <p className={styles.infoTileValue}>{formatRoi(aggregate.combinedRoi)}</p>
          </div>
        </div>

        <div className={styles.combinedMiniSummary}>
          <div className={styles.combinedMiniSummaryItem}>
            <span className={styles.infoTileLabel}>ROI terbaik</span>
            <strong>{aggregate.bestRoiReport ? aggregate.bestRoiReport.reportLabel : '—'}</strong>
          </div>
          <div className={styles.combinedMiniSummaryItem}>
            <span className={styles.infoTileLabel}>Spend terbesar</span>
            <strong>{aggregate.topSpendReport ? aggregate.topSpendReport.reportLabel : '—'}</strong>
          </div>
          <div className={styles.combinedMiniSummaryItem}>
            <span className={styles.infoTileLabel}>GMV terbesar</span>
            <strong>{aggregate.topGmvReport ? aggregate.topGmvReport.reportLabel : '—'}</strong>
          </div>
        </div>
      </section>

      <section className={styles.section}>
        <h3 className={styles.sectionTitle}>Overall metrics</h3>
        <p className={styles.sectionHint}>
          Tiga line chart ini nunjukin pergerakan Spending, GMV, dan ROI antar laporan terpilih.
        </p>
        <div className={styles.combinedChartGrid}>
          <article className={styles.chartCard}>
            <div ref={overallCostRef} className={styles.chartPlot} />
          </article>
          <article className={styles.chartCard}>
            <div ref={overallGmvRef} className={styles.chartPlot} />
          </article>
          <article className={styles.chartCard}>
            <div ref={overallRoiRef} className={styles.chartPlot} />
          </article>
        </div>
      </section>

      <section className={styles.section}>
        <h3 className={styles.sectionTitle}>Campaign-level comparison</h3>
        <p className={styles.sectionHint}>
          Diagram ini ambil campaign teratas per metric dan bandingin per laporan terpilih.
        </p>
        <div className={styles.combinedChartStack}>
          <article className={styles.chartCard}>
            <div ref={campaignCostRef} className={styles.chartPlotTall} />
          </article>
          <article className={styles.chartCard}>
            <div ref={campaignGmvRef} className={styles.chartPlotTall} />
          </article>
          <article className={styles.chartCard}>
            <div ref={campaignRoiRef} className={styles.chartPlotTall} />
          </article>
        </div>
      </section>

      <section className={styles.section}>
        <h3 className={styles.sectionTitle}>Summary rekomendasi</h3>
        <p className={styles.sectionHint}>
          Ranking-based: lihat ROI, spend, dan tren antar laporan untuk tentukan campaign yang perlu
          dinaikkan atau diturunkan budget-nya.
        </p>
        <div className={styles.recommendGrid}>
          <article className={styles.recommendCard}>
            <h4 className={styles.rankInfographicTitle}>Naikkan budget</h4>
            <ul className={styles.recommendList}>
              {aggregate.scaleCandidates.length > 0 ? (
                aggregate.scaleCandidates.map(({ campaign, roiRank, spendRank, reason }) => (
                  <li key={campaign.campaignName} className={styles.recommendItemGood}>
                    <strong>{campaign.campaignName}</strong>
                    <span>
                      ROI {formatRoi(campaign.blendedRoi)} · spend Rp {formatIdr(campaign.totalCost)} ·
                      rank ROI #{roiRank} / spend #{spendRank}
                    </span>
                    <small>{reason}</small>
                  </li>
                ))
              ) : (
                <li className={styles.empty}>Belum ada campaign yang cukup kuat buat dinaikkan.</li>
              )}
            </ul>
          </article>
          <article className={styles.recommendCard}>
            <h4 className={styles.rankInfographicTitle}>Pantau / turunkan budget</h4>
            <ul className={styles.recommendList}>
              {aggregate.reduceCandidates.length > 0 ? (
                aggregate.reduceCandidates.map(({ campaign, roiRank, spendRank, reason }) => (
                  <li key={campaign.campaignName} className={styles.recommendItemBad}>
                    <strong>{campaign.campaignName}</strong>
                    <span>
                      ROI {formatRoi(campaign.blendedRoi)} · spend Rp {formatIdr(campaign.totalCost)} ·
                      rank ROI #{roiRank} / spend #{spendRank}
                    </span>
                    <small>{reason}</small>
                  </li>
                ))
              ) : (
                <li className={styles.empty}>Tidak ada campaign yang perlu diturunkan berdasarkan ranking ini.</li>
              )}
            </ul>
          </article>
        </div>
      </section>
    </div>
  );
}

function compactReportLabel(label: string): string {
  return label.replace(/^Laporan \d+ · /, '');
}

function chartTheme(el: HTMLElement | null): ThemeVars {
  const s = el ? getComputedStyle(el) : null;
  return {
    paper: 'rgba(255, 255, 255, 0.02)',
    text: '#243b57',
    grid: 'rgba(84, 120, 155, 0.18)',
    colors: ['#174f73', '#2c86b3', '#a86b44', '#5d7ea6', '#7f9bb4'],
    fontFamily:
      s?.getPropertyValue('--font-body').trim() ||
      s?.getPropertyValue('font-family').trim() ||
      'Instrument Sans, sans-serif',
  };
}
