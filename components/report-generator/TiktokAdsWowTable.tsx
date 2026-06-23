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

function formatMetricValue(metric: WowMetricKey, value: number | null): string {
  if (value === null) return '—';
  if (metric === 'roi') return formatRoi(value);
  return `Rp ${formatIdr(value)}`;
}

const METRICS: Array<{ key: WowMetricKey; label: string }> = [
  { key: 'cost', label: 'Cost' },
  { key: 'grossRevenue', label: 'GMV' },
  { key: 'roi', label: 'ROI' },
];

type Props = {
  weeks: TiktokAdsWeekReport[];
};

export function TiktokAdsWowTable({ weeks }: Props) {
  const { campaigns, totals } = useMemo(() => buildWowRows(weeks), [weeks]);

  if (weeks.length < 2) return null;

  return (
    <div className={styles.tableWrap}>
      <table className={`${styles.table} ${styles.wowTable}`}>
        <thead>
          <tr>
            <th className={styles.wowStickyCol}>Campaign</th>
            {METRICS.flatMap((m) =>
              weeks.flatMap((w, wi) => {
                const cols = [
                  <th key={`${m.key}-w${w.weekIndex}`}>
                    L{w.weekIndex} {m.label}
                  </th>,
                ];
                if (wi > 0) {
                  cols.push(
                    <th key={`${m.key}-d${w.weekIndex}`} className={styles.wowDeltaCol}>
                      Δ%
                    </th>
                  );
                }
                return cols;
              })
            )}
          </tr>
        </thead>
        <tbody>
          {campaigns.map((row) => (
            <tr key={row.campaignName}>
              <td className={styles.wowStickyCol}>
                <span className={styles.adName}>{row.campaignName}</span>
              </td>
              {METRICS.flatMap((m) =>
                weeks.flatMap((w, wi) => {
                  const weekData = row.weeks[wi];
                  const value =
                    m.key === 'cost'
                      ? weekData?.cost ?? null
                      : m.key === 'grossRevenue'
                        ? weekData?.grossRevenue ?? null
                        : weekData?.roi ?? null;
                  const prevWeek = wi > 0 ? row.weeks[wi - 1] : null;
                  const prevValue =
                    prevWeek === null || prevWeek === undefined
                      ? null
                      : m.key === 'cost'
                        ? prevWeek.cost
                        : m.key === 'grossRevenue'
                          ? prevWeek.grossRevenue
                          : prevWeek.roi;
                  const delta = wi > 0 ? calcDeltaPct(value, prevValue) : null;

                  const cols = [
                    <td key={`${row.campaignName}-${m.key}-w${w.weekIndex}`} className={styles.numCell}>
                      {formatMetricValue(m.key, value)}
                    </td>,
                  ];
                  if (wi > 0) {
                    cols.push(
                      <td
                        key={`${row.campaignName}-${m.key}-d${w.weekIndex}`}
                        className={`${styles.numCell} ${styles.wowDeltaCol} ${deltaClass(m.key, delta)}`}
                      >
                        {formatDelta(delta)}
                      </td>
                    );
                  }
                  return cols;
                })
              )}
            </tr>
          ))}
          <tr className={styles.wowTotalsRow}>
            <td className={styles.wowStickyCol}>
              <strong>Total</strong>
            </td>
            {METRICS.flatMap((m) =>
              weeks.flatMap((w, wi) => {
                const total = totals[wi];
                const value =
                  m.key === 'cost'
                    ? total?.cost ?? null
                    : m.key === 'grossRevenue'
                      ? total?.grossRevenue ?? null
                      : total?.roi ?? null;
                const prevTotal = wi > 0 ? totals[wi - 1] : null;
                const prevValue =
                  prevTotal === null || prevTotal === undefined
                    ? null
                    : m.key === 'cost'
                      ? prevTotal.cost
                      : m.key === 'grossRevenue'
                        ? prevTotal.grossRevenue
                        : prevTotal.roi;
                const delta = wi > 0 ? calcDeltaPct(value, prevValue) : null;

                const cols = [
                  <td key={`total-${m.key}-w${w.weekIndex}`} className={styles.numCell}>
                    <strong>{formatMetricValue(m.key, value)}</strong>
                  </td>,
                ];
                if (wi > 0) {
                  cols.push(
                    <td
                      key={`total-${m.key}-d${w.weekIndex}`}
                      className={`${styles.numCell} ${styles.wowDeltaCol} ${deltaClass(m.key, delta)}`}
                    >
                      <strong>{formatDelta(delta)}</strong>
                    </td>
                  );
                }
                return cols;
              })
            )}
          </tr>
        </tbody>
      </table>
    </div>
  );
}
