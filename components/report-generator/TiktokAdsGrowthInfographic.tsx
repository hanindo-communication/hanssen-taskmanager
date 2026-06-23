'use client';

import { useMemo } from 'react';
import {
  buildWowRows,
  calcDeltaPct,
  type TiktokAdsWeekReport,
  type WowMetricKey,
} from '@/lib/report-generator/tiktok-ads-weekly';
import styles from './ReportGeneratorPanel.module.css';

function formatIdr(n: number): string {
  return new Intl.NumberFormat('id-ID', { maximumFractionDigits: 0 }).format(Math.round(n));
}

function formatRoi(n: number | null): string {
  if (n === null || !Number.isFinite(n)) return '—';
  return n.toLocaleString('id-ID', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatDelta(n: number | null): string {
  if (n === null || !Number.isFinite(n)) return '—';
  const sign = n > 0 ? '+' : '';
  return `${sign}${n.toLocaleString('id-ID', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%`;
}

function deltaClass(metric: WowMetricKey, delta: number | null): string {
  if (delta === null || !Number.isFinite(delta) || delta === 0) return styles.deltaNeutral;
  const isGood = metric === 'cost' ? delta < 0 : delta > 0;
  return isGood ? styles.deltaPositive : styles.deltaNegative;
}

function deltaArrow(delta: number | null): string {
  if (delta === null || !Number.isFinite(delta) || delta === 0) return '→';
  return delta > 0 ? '▲' : '▼';
}

type MetricConfig = {
  key: WowMetricKey;
  title: string;
  hint: string;
  barClass: string;
  getValue: (total: { cost: number; grossRevenue: number; roi: number | null }) => number | null;
  formatValue: (v: number | null) => string;
};

const METRICS: MetricConfig[] = [
  {
    key: 'cost',
    title: 'Growth Cost',
    hint: 'Total spending antar periode',
    barClass: styles.barFillSpend,
    getValue: (t) => t.cost,
    formatValue: (v) => (v !== null ? `Rp ${formatIdr(v)}` : '—'),
  },
  {
    key: 'grossRevenue',
    title: 'Growth GMV',
    hint: 'Total Gross revenue antar periode',
    barClass: styles.barFillGmv,
    getValue: (t) => t.grossRevenue,
    formatValue: (v) => (v !== null ? `Rp ${formatIdr(v)}` : '—'),
  },
  {
    key: 'roi',
    title: 'Growth ROI',
    hint: 'ROI campuran antar periode',
    barClass: styles.barFillRoi,
    getValue: (t) => t.roi,
    formatValue: (v) => formatRoi(v),
  },
];

type Props = {
  weeks: TiktokAdsWeekReport[];
};

export function TiktokAdsGrowthInfographic({ weeks }: Props) {
  const { totals } = useMemo(() => buildWowRows(weeks), [weeks]);

  if (weeks.length < 2) return null;

  return (
    <div className={styles.threeColRankGrid}>
      {METRICS.map((metric) => (
        <GrowthMetricCard
          key={metric.key}
          metric={metric}
          weeks={weeks}
          totals={totals}
        />
      ))}
    </div>
  );
}

function GrowthMetricCard({
  metric,
  weeks,
  totals,
}: {
  metric: MetricConfig;
  weeks: TiktokAdsWeekReport[];
  totals: Array<{ cost: number; grossRevenue: number; roi: number | null }>;
}) {
  const values = totals.map((t) => metric.getValue(t));
  const max = Math.max(...values.map((v) => (v !== null && Number.isFinite(v) ? v : 0)), 1);

  const transitions = weeks.slice(1).map((w, i) => {
    const prev = values[i];
    const curr = values[i + 1];
    const delta = calcDeltaPct(curr, prev);
    return {
      fromLabel: `L${i + 1}`,
      toLabel: `L${i + 2}`,
      delta,
    };
  });

  const latestDelta = transitions[transitions.length - 1]?.delta ?? null;

  return (
    <div className={`${styles.rankInfographicCard} ${styles.growthCard}`}>
      <h4 className={styles.rankInfographicTitle}>{metric.title}</h4>
      <p className={styles.rankInfographicHint}>{metric.hint}</p>

      {weeks.map((w, i) => {
        const value = values[i];
        const pct =
          value !== null && Number.isFinite(value) && max > 0
            ? Math.min(100, Math.max(0, (value / max) * 100))
            : 0;
        return (
          <div key={w.weekIndex} className={styles.growthBarRow}>
            <div className={styles.barRankHeader}>
              <span className={styles.barRankBadge}>L{w.weekIndex}</span>
              <span className={styles.barCampaignName} title={w.label}>
                {w.label.replace(/^Laporan \d+ · /, '')}
              </span>
              <span className={styles.barValueCompact}>{metric.formatValue(value)}</span>
            </div>
            <div className={styles.barTrackTall}>
              <div className={metric.barClass} style={{ width: `${pct}%` }} />
            </div>
          </div>
        );
      })}

      {latestDelta !== null ? (
        <p className={`${styles.growthDeltaBadge} ${deltaClass(metric.key, latestDelta)}`}>
          {deltaArrow(latestDelta)} {formatDelta(latestDelta)}{' '}
          <span className={styles.growthDeltaSub}>
            (L{weeks.length - 1} → L{weeks.length})
          </span>
        </p>
      ) : null}

      {transitions.length > 1 ? (
        <ul className={styles.growthTransitionList}>
          {transitions.slice(0, -1).map((t) => (
            <li key={`${t.fromLabel}-${t.toLabel}`} className={deltaClass(metric.key, t.delta)}>
              {t.fromLabel} → {t.toLabel}: {formatDelta(t.delta)}
            </li>
          ))}
        </ul>
      ) : null}
    </div>
  );
}
