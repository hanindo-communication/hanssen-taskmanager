'use client';

import { useCallback, useMemo, useRef, useState } from 'react';
import boardStyles from '@/components/board/board-client.module.css';
import {
  type TiktokAdsCampaignAggregate,
  type TiktokAdsReportSummary,
} from '@/lib/report-generator/tiktok-ads-campaign-xlsx';
import { downloadElementAsPdf, downloadPdfZip } from '@/lib/report-generator/report-pdf';
import {
  appendToWeeklyBundle,
  buildWeeklyBundle,
  buildCombinedSelectionSummary,
  type ParsedFileInput,
  type TiktokAdsWeekReport,
  type TiktokAdsWeeklyBundle,
} from '@/lib/report-generator/tiktok-ads-weekly';
import { TiktokAdsCombinedView } from './TiktokAdsCombinedView';
import { TiktokAdsGrowthInfographic } from './TiktokAdsGrowthInfographic';
import { TiktokAdsWowTable } from './TiktokAdsWowTable';
import styles from './ReportGeneratorPanel.module.css';

function formatIdr(n: number): string {
  return new Intl.NumberFormat('id-ID', { maximumFractionDigits: 0 }).format(Math.round(n));
}

function formatRoi(n: number | null): string {
  if (n === null || !Number.isFinite(n)) return '—';
  return n.toLocaleString('id-ID', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function coerceMatrixFromXlsxSheet(rows: unknown[][]): string[][] {
  return (rows ?? []).map((row) =>
    (row ?? []).map((cell) => {
      if (cell === undefined || cell === null) return '';
      if (typeof cell === 'number') return String(cell);
      return String(cell).trim();
    })
  );
}

function normalizeSheetName(n: string): string {
  return n.trim().toLowerCase();
}

async function readXlsxMatrix(file: File): Promise<string[][]> {
  const buf = await file.arrayBuffer();
  const XLSX = await import('xlsx');
  const wb = XLSX.read(buf, { type: 'array' });
  const sheetName = wb.SheetNames.find((n) => normalizeSheetName(n) === 'data') ?? wb.SheetNames[0];
  if (!sheetName) throw new Error('no-sheet');
  const sheet = wb.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: '',
    raw: false,
  }) as unknown[][];
  return coerceMatrixFromXlsxSheet(raw);
}

async function filesToParsedInputs(
  files: File[],
  startOrder: number
): Promise<ParsedFileInput[]> {
  const parsed: ParsedFileInput[] = [];
  let order = startOrder;
  for (const file of files) {
    const matrix = await readXlsxMatrix(file);
    parsed.push({ fileName: file.name, matrix, uploadOrder: order++ });
  }
  return parsed;
}

function resetFileInput(input: HTMLInputElement | null): void {
  if (input) input.value = '';
}

function activeCampaigns(summary: TiktokAdsReportSummary): TiktokAdsCampaignAggregate[] {
  return summary.campaigns.filter((c) => c.cost > 0);
}

function topByRoi(campaigns: TiktokAdsCampaignAggregate[]): TiktokAdsCampaignAggregate[] {
  const list = [...campaigns].filter((c) => c.roi !== null && c.roi > 0);
  list.sort((a, b) => (b.roi ?? 0) - (a.roi ?? 0));
  return list.slice(0, 3);
}

function topByGmv(campaigns: TiktokAdsCampaignAggregate[]): TiktokAdsCampaignAggregate[] {
  const list = [...campaigns];
  list.sort((a, b) => b.grossRevenue - a.grossRevenue);
  return list.slice(0, 3);
}

function topBySpend(campaigns: TiktokAdsCampaignAggregate[]): TiktokAdsCampaignAggregate[] {
  const list = [...campaigns];
  list.sort((a, b) => b.cost - a.cost);
  return list.slice(0, 3);
}

function tableRowsFrom(campaigns: TiktokAdsCampaignAggregate[]): TiktokAdsCampaignAggregate[] {
  const list = [...campaigns];
  list.sort((a, b) => b.cost - a.cost);
  return list;
}

function findWeekIndexForFile(weeks: TiktokAdsWeekReport[], fileName: string): number {
  const idx = weeks.findIndex((w) => w.sourceFiles.includes(fileName));
  return idx >= 0 ? idx : weeks.length - 1;
}

function weekKey(week: TiktokAdsWeekReport): string {
  return week.sourceFiles[0] ?? `${week.weekIndex}`;
}

function formatFileLabel(label: string): string {
  return label.replace(/^Laporan \d+ · /, '');
}

export function TiktokAdsCampaignReportPanel() {
  const [error, setError] = useState<string | null>(null);
  const [bundle, setBundle] = useState<TiktokAdsWeeklyBundle | null>(null);
  const [parsedFiles, setParsedFiles] = useState<ParsedFileInput[]>([]);
  const [activeWeekIndex, setActiveWeekIndex] = useState(0);
  const [viewMode, setViewMode] = useState<'single' | 'combined'>('single');
  const [combinedModalOpen, setCombinedModalOpen] = useState(false);
  const [combinedSelection, setCombinedSelection] = useState<string[]>([]);
  const [combinedDraftSelection, setCombinedDraftSelection] = useState<string[]>([]);
  const [pdfBusy, setPdfBusy] = useState(false);
  const [uploadOrderCounter, setUploadOrderCounter] = useState(0);

  const initialInputRef = useRef<HTMLInputElement>(null);
  const addInputRef = useRef<HTMLInputElement>(null);
  const weekPdfRefs = useRef<Map<number, HTMLDivElement>>(new Map());
  const growthPdfRef = useRef<HTMLDivElement>(null);
  const wowPdfRef = useRef<HTMLDivElement>(null);

  const weeks = bundle?.weeks ?? [];
  const activeWeek = weeks[activeWeekIndex] ?? null;
  const fileNames = parsedFiles.map((f) => f.fileName);
  const availableWeekKeys = weeks.map(weekKey);
  const selectedCombinedWeeks = useMemo(() => {
    const keys = combinedSelection.length > 0 ? combinedSelection : availableWeekKeys;
    const keySet = new Set(keys);
    return weeks.filter((w) => keySet.has(weekKey(w)));
  }, [availableWeekKeys, combinedSelection, weeks]);
  const combinedSummary = useMemo(() => {
    if (selectedCombinedWeeks.length === 0) return null;
    return buildCombinedSelectionSummary(selectedCombinedWeeks);
  }, [selectedCombinedWeeks]);
  const combinedHero = useMemo(() => {
    if (!combinedSummary) return null;
    const totalCost = combinedSummary.totals.reduce((sum, t) => sum + t.cost, 0);
    const totalGrossRevenue = combinedSummary.totals.reduce((sum, t) => sum + t.grossRevenue, 0);
    const blendedRoi = totalCost > 0 ? totalGrossRevenue / totalCost : null;
    return {
      title: `${selectedCombinedWeeks.length} laporan dipilih`,
      totalCost,
      totalGrossRevenue,
      blendedRoi,
      campaignCount: combinedSummary.campaigns.length,
    };
  }, [combinedSummary, selectedCombinedWeeks.length]);

  const applyBundle = useCallback((next: TiktokAdsWeeklyBundle, files: ParsedFileInput[]) => {
    setBundle(next);
    setParsedFiles(files);
    if (next.weeks.length === 0) {
      setError(next.warnings[0] ?? 'Tidak ada laporan valid dari file yang diunggah.');
    }
  }, []);

  const onInitialFiles = useCallback(
    async (fileList: FileList | null) => {
      if (!fileList?.length) return;

      const files = [...fileList];
      const invalid = files.filter((f) => !f.name.toLowerCase().endsWith('.xlsx'));
      if (invalid.length > 0) {
        setError('Gunakan file .xlsx (Product campaign data dari TikTok Ads).');
        setBundle(null);
        setParsedFiles([]);
        resetFileInput(initialInputRef.current);
        return;
      }

      setError(null);
      weekPdfRefs.current.clear();

      try {
        const parsed = await filesToParsedInputs(files, 0);
        const next = buildWeeklyBundle(parsed);
        applyBundle(next, parsed);
        setUploadOrderCounter(parsed.length);
        setActiveWeekIndex(0);
        setViewMode('single');
        setCombinedSelection([]);
        setCombinedDraftSelection([]);
        setCombinedModalOpen(false);
        resetFileInput(initialInputRef.current);
      } catch {
        setError('Gagal membaca file Excel.');
        setBundle(null);
        setParsedFiles([]);
        resetFileInput(initialInputRef.current);
      }
    },
    [applyBundle]
  );

  const onAddFiles = useCallback(
    async (fileList: FileList | null) => {
      if (!fileList?.length || parsedFiles.length === 0) return;

      const files = [...fileList];
      const invalid = files.filter((f) => !f.name.toLowerCase().endsWith('.xlsx'));
      if (invalid.length > 0) {
        setError('Gunakan file .xlsx (Product campaign data dari TikTok Ads).');
        resetFileInput(addInputRef.current);
        return;
      }

      setError(null);

      try {
        const incoming = await filesToParsedInputs(files, uploadOrderCounter);
        const { bundle: next, parsedFiles: merged, addedFileNames } = appendToWeeklyBundle(
          parsedFiles,
          incoming
        );

        applyBundle(next, merged);
        setUploadOrderCounter((c) => c + incoming.length);

        if (addedFileNames.length > 0 && next.weeks.length > 0) {
          const lastAdded = addedFileNames[addedFileNames.length - 1];
          setActiveWeekIndex(findWeekIndexForFile(next.weeks, lastAdded));
        }

        resetFileInput(addInputRef.current);
      } catch {
        setError('Gagal membaca file Excel tambahan.');
        resetFileInput(addInputRef.current);
      }
    },
    [applyBundle, parsedFiles, uploadOrderCounter]
  );

  const handleDownloadActiveWeekPdf = useCallback(async () => {
    if (!activeWeek) return;
    const el = weekPdfRefs.current.get(activeWeek.weekIndex);
    if (!el) return;
    setPdfBusy(true);
    setError(null);
    try {
      await downloadElementAsPdf(
        el,
        `laporan-gmv-max-tiktok-laporan-${activeWeek.weekIndex}_${activeWeek.startDate}_${activeWeek.endDate}`
      );
    } catch {
      setError('Gagal membuat PDF. Coba lagi.');
    } finally {
      setPdfBusy(false);
    }
  }, [activeWeek]);

  const handleDownloadAllPdf = useCallback(async () => {
    if (weeks.length === 0) return;
    setPdfBusy(true);
    setError(null);
    try {
      const entries: Array<{ element: HTMLElement; fileName: string }> = [];

      for (const week of weeks) {
        const el = weekPdfRefs.current.get(week.weekIndex);
        if (el) {
          entries.push({
            element: el,
            fileName: `laporan-gmv-max-tiktok-laporan-${week.weekIndex}_${week.startDate}_${week.endDate}.pdf`,
          });
        }
      }

      if (weeks.length >= 2) {
        const first = weeks[0];
        const last = weeks[weeks.length - 1];
        if (growthPdfRef.current) {
          entries.push({
            element: growthPdfRef.current,
            fileName: `laporan-gmv-max-tiktok-growth_${first.startDate}_${last.endDate}.pdf`,
          });
        }
        if (wowPdfRef.current) {
          entries.push({
            element: wowPdfRef.current,
            fileName: `laporan-gmv-max-tiktok-wow_${first.startDate}_${last.endDate}.pdf`,
          });
        }
      }

      const stamp = new Date().toISOString().slice(0, 10);
      await downloadPdfZip(entries, `laporan-gmv-max-tiktok-bulk-${stamp}`);
    } catch {
      setError('Gagal membuat ZIP PDF. Coba lagi.');
    } finally {
      setPdfBusy(false);
    }
  }, [weeks]);

  const activeCampaignList = useMemo(
    () => (activeWeek ? activeCampaigns(activeWeek.summary) : []),
    [activeWeek]
  );

  const openCombinedModal = useCallback(() => {
    const defaults = combinedSelection.length > 0 ? combinedSelection : availableWeekKeys;
    setCombinedDraftSelection(defaults);
    setCombinedModalOpen(true);
  }, [availableWeekKeys, combinedSelection]);

  const closeCombinedModal = useCallback(() => {
    setCombinedModalOpen(false);
  }, []);

  const confirmCombinedModal = useCallback(() => {
    const next = availableWeekKeys.filter((key) => combinedDraftSelection.includes(key));
    if (next.length === 0) return;
    setCombinedSelection(next);
    setViewMode('combined');
    setCombinedModalOpen(false);
  }, [availableWeekKeys, combinedDraftSelection]);

  const toggleDraftSelection = useCallback((key: string) => {
    setCombinedDraftSelection((current) =>
      current.includes(key) ? current.filter((item) => item !== key) : [...current, key]
    );
  }, []);

  return (
    <div className={weeks.length ? styles.printRoot : undefined} data-report-ui="tiktok-ads-gmv-max">
      <section
        className={`${boardStyles.overviewHero} ${styles.reportHero} ${styles.reportHeroTiktok}`}
      >
        <div className={styles.heroTextCol}>
          <p className={boardStyles.heroEyebrow}>Report Generator</p>
          <h2 className={boardStyles.heroTitle}>Laporan GMV MAX TikTok</h2>
          <p className={`${boardStyles.heroDescription} ${styles.heroDescriptionTight}`}>
            Unggah <strong>XLSX Product campaign data</strong> — range periode dibaca dari{' '}
            <strong>nama file</strong> (<code>YYYY-MM-DD - YYYY-MM-DD</code>). Mingguan, bulanan,
            atau quarter semua bisa. Setelah laporan pertama, pakai{' '}
            <strong>Pilih File tambahan</strong> untuk periode berikutnya dan lihat infografik
            growth antar laporan.
          </p>
          <div className={styles.uploadRow} data-no-pdf>
            <label className={styles.fileLabel}>
              <input
                ref={initialInputRef}
                type="file"
                multiple
                accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                className={styles.fileInput}
                onChange={(e) => void onInitialFiles(e.target.files)}
              />
              Pilih file XLSX (bulk)
            </label>
            {weeks.length > 0 ? (
              <label className={`${styles.fileLabel} ${styles.fileLabelSecondary}`}>
                <input
                  ref={addInputRef}
                  type="file"
                  multiple
                  accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                  className={styles.fileInput}
                  onChange={(e) => void onAddFiles(e.target.files)}
                />
                Pilih File tambahan
              </label>
            ) : null}
            {weeks.length > 0 ? (
              <button
                type="button"
                className={`${styles.printBtn} ${styles.printBtnSecondary}`}
                onClick={openCombinedModal}
              >
                Buat laporan gabungan
              </button>
            ) : null}
            {weeks.length > 0 ? (
              <div className={styles.bulkActions}>
                <button
                  type="button"
                  className={styles.printBtn}
                  disabled={pdfBusy || !activeWeek}
                  onClick={() => void handleDownloadActiveWeekPdf()}
                >
                  {pdfBusy ? 'Menyusun PDF…' : 'Unduh PDF laporan ini'}
                </button>
                <button
                  type="button"
                  className={`${styles.printBtn} ${styles.printBtnSecondary}`}
                  disabled={pdfBusy}
                  onClick={() => void handleDownloadAllPdf()}
                >
                  {pdfBusy ? 'Menyusun ZIP…' : 'Unduh semua PDF (ZIP)'}
                </button>
              </div>
            ) : null}
          </div>
          {fileNames.length > 0 ? (
            <ul className={styles.uploadedFileList} data-no-pdf>
              {fileNames.map((name) => (
                <li key={name} className={styles.uploadedFileChip} title={name}>
                  {name}
                </li>
              ))}
            </ul>
          ) : null}
          {error ? <p className={styles.error}>{error}</p> : null}
          {bundle?.warnings?.length ? (
            <ul className={styles.warningList} data-no-pdf>
              {bundle.warnings.map((w) => (
                <li key={w}>{w}</li>
              ))}
            </ul>
          ) : null}
        </div>
        {viewMode === 'combined' && combinedHero ? (
          <div className={`${boardStyles.heroStatsGrid} ${styles.heroMetrics}`}>
            <div className={`${boardStyles.metricCard} ${styles.metricCardLift}`}>
              <p className={boardStyles.metricLabel}>Total Cost (gabungan)</p>
              <p className={boardStyles.metricValue}>Rp {formatIdr(combinedHero.totalCost)}</p>
              <p className={styles.metricSub}>
                {combinedHero.title} · {combinedHero.campaignCount} kampanye unik
              </p>
            </div>
            <div className={`${boardStyles.metricCard} ${styles.metricCardLift}`}>
              <p className={boardStyles.metricLabel}>Total GMV (gabungan)</p>
              <p className={boardStyles.metricValue}>Rp {formatIdr(combinedHero.totalGrossRevenue)}</p>
              <p className={styles.metricSub}>Σ Gross revenue dari laporan terpilih</p>
            </div>
            <div className={`${boardStyles.metricCard} ${styles.metricCardLift}`}>
              <p className={boardStyles.metricLabel}>ROI campuran (gabungan)</p>
              <p className={boardStyles.metricValue}>{formatRoi(combinedHero.blendedRoi)}</p>
              <p className={styles.metricSub}>Total GMV ÷ total Cost</p>
            </div>
          </div>
        ) : activeWeek ? (
          <div className={`${boardStyles.heroStatsGrid} ${styles.heroMetrics}`}>
            <div className={`${boardStyles.metricCard} ${styles.metricCardLift}`}>
              <p className={boardStyles.metricLabel}>Total Cost (spending)</p>
              <p className={boardStyles.metricValue}>Rp {formatIdr(activeWeek.summary.totalCost)}</p>
              <p className={styles.metricSub}>
                {activeWeek.label} · {activeCampaignList.length} kampanye ber-spend
              </p>
            </div>
            <div className={`${boardStyles.metricCard} ${styles.metricCardLift}`}>
              <p className={boardStyles.metricLabel}>Total GMV</p>
              <p className={boardStyles.metricValue}>
                Rp {formatIdr(activeWeek.summary.totalGrossRevenue)}
              </p>
              <p className={styles.metricSub}>Σ Gross revenue (kolom TikTok)</p>
            </div>
            <div className={`${boardStyles.metricCard} ${styles.metricCardLift}`}>
              <p className={boardStyles.metricLabel}>ROI campuran</p>
              <p className={boardStyles.metricValue}>{formatRoi(activeWeek.summary.blendedRoi)}</p>
              <p className={styles.metricSub}>Total GMV ÷ total Cost</p>
            </div>
          </div>
        ) : (
          <div className={`${boardStyles.heroStatsGrid} ${styles.heroMetrics}`}>
            <div className={`${boardStyles.metricCard} ${styles.placeholderCard}`}>
              <p className={boardStyles.metricLabel}>Total Cost</p>
              <p className={boardStyles.metricValue}>—</p>
            </div>
            <div className={`${boardStyles.metricCard} ${styles.placeholderCard}`}>
              <p className={boardStyles.metricLabel}>Total GMV</p>
              <p className={boardStyles.metricValue}>—</p>
            </div>
            <div className={`${boardStyles.metricCard} ${styles.placeholderCard}`}>
              <p className={boardStyles.metricLabel}>ROI campuran</p>
              <p className={boardStyles.metricValue}>—</p>
            </div>
          </div>
        )}
      </section>

      {weeks.length > 0 ? (
        viewMode === 'combined' ? (
          <TiktokAdsCombinedView
            weeks={selectedCombinedWeeks}
            onEditSelection={openCombinedModal}
            onBackToSingle={() => setViewMode('single')}
          />
        ) : (
          <>
            <div className={styles.weekTabs} data-no-pdf>
              {weeks.map((w, i) => (
                <button
                  key={w.sourceFiles[0] ?? w.weekIndex}
                  type="button"
                  className={i === activeWeekIndex ? styles.weekTabActive : styles.weekTab}
                  onClick={() => setActiveWeekIndex(i)}
                >
                  {w.label}
                </button>
              ))}
            </div>

            {weeks.map((week) => (
              <WeeklyReportSection
                key={week.sourceFiles[0] ?? week.weekIndex}
                week={week}
                hidden={week.weekIndex !== activeWeek?.weekIndex}
                pdfRef={(el) => {
                  if (el) weekPdfRefs.current.set(week.weekIndex, el);
                  else weekPdfRefs.current.delete(week.weekIndex);
                }}
              />
            ))}

            {weeks.length >= 2 ? (
              <>
                <section className={styles.section}>
                  <h3 className={styles.sectionTitle}>Growth antar laporan</h3>
                  <p className={styles.sectionHint}>
                    Perbandingan total Cost, GMV, dan ROI campuran antar periode. Δ% = perubahan vs
                    laporan sebelumnya.
                  </p>
                  <div data-no-pdf>
                    <TiktokAdsGrowthInfographic weeks={weeks} />
                  </div>
                  <div ref={growthPdfRef} className={styles.pdfOnlyBlock} aria-hidden>
                    <h3 className={styles.sectionTitle}>Growth antar laporan</h3>
                    <p className={styles.sectionHint}>
                      {weeks[0]?.startDate} – {weeks[weeks.length - 1]?.endDate}
                    </p>
                    <TiktokAdsGrowthInfographic weeks={weeks} />
                  </div>
                </section>

                <section className={styles.section}>
                  <h3 className={styles.sectionTitle}>Perbandingan antar laporan</h3>
                  <p className={styles.sectionHint}>
                    Cost, GMV, dan ROI per kampanye antar file yang diunggah. Δ% = perubahan vs
                    laporan sebelumnya. Cost turun = hijau; GMV/ROI naik = hijau.
                  </p>
                  <div data-no-pdf>
                    <TiktokAdsWowTable weeks={weeks} />
                  </div>
                  <div ref={wowPdfRef} className={styles.pdfOnlyBlock} aria-hidden>
                    <h3 className={styles.sectionTitle}>Perbandingan antar laporan</h3>
                    <p className={styles.sectionHint}>
                      {weeks[0]?.startDate} – {weeks[weeks.length - 1]?.endDate}
                    </p>
                    <TiktokAdsWowTable weeks={weeks} />
                  </div>
                </section>
              </>
            ) : null}
          </>
        )
      ) : null}

      {combinedModalOpen ? (
        <div className={styles.modalBackdrop} onClick={closeCombinedModal} role="presentation">
          <div
            className={styles.modalCard}
            role="dialog"
            aria-modal="true"
            aria-labelledby="combined-report-title"
            onClick={(e) => e.stopPropagation()}
          >
            <div className={styles.modalHeader}>
              <div>
                <h3 id="combined-report-title" className={styles.sectionTitle}>
                  Buat laporan gabungan
                </h3>
                <p className={styles.sectionHint}>
                  Pilih 1 atau lebih laporan untuk dibandingkan dalam view gabungan.
                </p>
              </div>
              <button type="button" className={styles.modalCloseBtn} onClick={closeCombinedModal}>
                Tutup
              </button>
            </div>

            <div className={styles.modalList}>
              {weeks.map((week) => {
                const key = weekKey(week);
                const checked = combinedDraftSelection.includes(key);
                return (
                  <label key={key} className={styles.modalOption}>
                    <input
                      type="checkbox"
                      checked={checked}
                      onChange={() => toggleDraftSelection(key)}
                    />
                    <span>
                      <strong>{week.label}</strong>
                      <small>{formatFileLabel(week.label)}</small>
                    </span>
                  </label>
                );
              })}
            </div>

            <div className={styles.modalFooter}>
              <span className={styles.modalMeta}>
                {combinedDraftSelection.length} laporan dipilih
              </span>
              <div className={styles.bulkActions}>
                <button
                  type="button"
                  className={`${styles.printBtn} ${styles.printBtnSecondary}`}
                  onClick={closeCombinedModal}
                >
                  Batal
                </button>
                <button
                  type="button"
                  className={styles.printBtn}
                  disabled={combinedDraftSelection.length === 0}
                  onClick={() => void confirmCombinedModal()}
                >
                  Pakai laporan gabungan
                </button>
              </div>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}

function WeeklyReportSection({
  week,
  hidden,
  pdfRef,
}: {
  week: TiktokAdsWeekReport;
  hidden: boolean;
  pdfRef: (el: HTMLDivElement | null) => void;
}) {
  const campaigns = useMemo(() => activeCampaigns(week.summary), [week.summary]);
  const topRoi = useMemo(() => topByRoi(campaigns), [campaigns]);
  const topGmv = useMemo(() => topByGmv(campaigns), [campaigns]);
  const topSpend = useMemo(() => topBySpend(campaigns), [campaigns]);
  const tableRows = useMemo(() => tableRowsFrom(campaigns), [campaigns]);

  return (
    <div
      ref={pdfRef}
      className={hidden ? styles.weekPanelHidden : styles.weekPanelVisible}
      data-week-panel={week.weekIndex}
    >
      <div className={styles.reportBody}>
        <section className={`${styles.section} ${styles.weekHeroSection}`}>
          <p className={boardStyles.heroEyebrow}>Laporan GMV MAX TikTok</p>
          <h3 className={styles.sectionTitle}>{week.label}</h3>
          <div className={styles.weekHeroMetrics}>
            <div className={styles.infoTile}>
              <p className={styles.infoTileLabel}>Total Cost</p>
              <p className={styles.infoTileValue}>Rp {formatIdr(week.summary.totalCost)}</p>
            </div>
            <div className={styles.infoTile}>
              <p className={styles.infoTileLabel}>Total GMV</p>
              <p className={styles.infoTileValue}>Rp {formatIdr(week.summary.totalGrossRevenue)}</p>
            </div>
            <div className={styles.infoTile}>
              <p className={styles.infoTileLabel}>ROI campuran</p>
              <p className={styles.infoTileValue}>{formatRoi(week.summary.blendedRoi)}</p>
            </div>
          </div>
        </section>

        <section className={styles.section}>
          <h3 className={styles.sectionTitle}>Ringkasan</h3>
          <p className={styles.meta}>
            {week.summary.rowCount} baris sumber · {week.summary.campaignCount} nama kampanye ·{' '}
            {campaigns.length} dengan Cost &gt; 0 · Cost per order{' '}
            {week.summary.blendedCostPerOrder !== null
              ? `Rp ${formatIdr(week.summary.blendedCostPerOrder)}`
              : '—'}{' '}
            · Σ daily budget cap Rp {formatIdr(week.summary.sumOfDailyBudgetCaps)}
            {week.sourceFiles.length > 0 ? (
              <> · Sumber: {week.sourceFiles.join(', ')}</>
            ) : null}
          </p>
        </section>

        <section className={styles.section}>
          <h3 className={styles.sectionTitle}>Top 3 — infografik</h3>
          <p className={styles.sectionHint}>
            Hanya kampanye dengan Cost &gt; 0. Panjang batang relatif terhadap #1 di masing-masing
            kartu.
          </p>
          <div className={styles.threeColRankGrid}>
            <TopThreeCard
              title="Top 3 ROI"
              hint="Gross revenue ÷ Cost"
              entries={topRoi}
              maxValue={topRoi[0]?.roi ?? 1}
              formatValue={(c) => formatRoi(c.roi)}
              barClass={styles.barFillRoi}
              pctForRow={(c, max) => (max > 0 && c.roi !== null ? (c.roi / max) * 100 : 0)}
            />
            <TopThreeCard
              title="Top 3 GMV"
              hint="Gross revenue"
              entries={topGmv}
              maxValue={topGmv[0]?.grossRevenue ?? 1}
              formatValue={(c) => `Rp ${formatIdr(c.grossRevenue)}`}
              barClass={styles.barFillGmv}
              pctForRow={(c, max) => (max > 0 ? (c.grossRevenue / max) * 100 : 0)}
            />
            <TopThreeCard
              title="Top 3 spending"
              hint="Cost"
              entries={topSpend}
              maxValue={topSpend[0]?.cost ?? 1}
              formatValue={(c) => `Rp ${formatIdr(c.cost)}`}
              barClass={styles.barFillSpend}
              pctForRow={(c, max) => (max > 0 ? (c.cost / max) * 100 : 0)}
            />
          </div>
        </section>

        <section className={styles.section}>
          <h3 className={styles.sectionTitle}>Per kampanye</h3>
          <p className={styles.sectionHint}>
            Hanya Cost &gt; 0; urut Cost tertinggi. ROI per baris = Gross revenue ÷ Cost.
          </p>
          <CampaignTable rows={tableRows} />
        </section>
      </div>
    </div>
  );
}

function TopThreeCard({
  title,
  hint,
  entries,
  maxValue,
  formatValue,
  barClass,
  pctForRow,
}: {
  title: string;
  hint: string;
  entries: TiktokAdsCampaignAggregate[];
  maxValue: number;
  formatValue: (c: TiktokAdsCampaignAggregate) => string;
  barClass: string;
  pctForRow: (c: TiktokAdsCampaignAggregate, max: number) => number;
}) {
  if (entries.length === 0) {
    return (
      <div className={styles.rankInfographicCard}>
        <h4 className={styles.rankInfographicTitle}>{title}</h4>
        <p className={styles.rankInfographicHint}>{hint}</p>
        <p className={styles.empty}>Tidak ada data (perlu kampanye dengan Cost &gt; 0).</p>
      </div>
    );
  }

  const max = maxValue > 0 ? maxValue : 1;

  return (
    <div className={styles.rankInfographicCard}>
      <h4 className={styles.rankInfographicTitle}>{title}</h4>
      <p className={styles.rankInfographicHint}>{hint}</p>
      {entries.map((c, i) => {
        const pct = Math.min(100, Math.max(0, pctForRow(c, max)));
        return (
          <div key={c.campaignName} className={styles.barRankRow}>
            <div className={styles.barRankHeader}>
              <span className={styles.barRankBadge}>{i + 1}</span>
              <span className={styles.barCampaignName} title={c.campaignName}>
                {c.campaignName}
              </span>
              <span className={styles.barValueCompact}>{formatValue(c)}</span>
            </div>
            <div className={styles.barTrackTall}>
              <div className={barClass} style={{ width: `${pct}%` }} />
            </div>
          </div>
        );
      })}
    </div>
  );
}

function CampaignTable({ rows }: { rows: TiktokAdsCampaignAggregate[] }) {
  if (rows.length === 0) {
    return <p className={styles.empty}>Tidak ada kampanye dengan spending (Cost &gt; 0).</p>;
  }

  return (
    <div className={styles.tableWrap}>
      <table className={styles.table}>
        <thead>
          <tr>
            <th>Campaign</th>
            <th>Cost</th>
            <th>Net Cost</th>
            <th>Daily budget</th>
            <th>GMV</th>
            <th>ROI</th>
            <th>SKU orders</th>
            <th>Cost / order</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((c) => (
            <tr key={c.campaignName}>
              <td>
                <span className={styles.adName}>{c.campaignName}</span>
              </td>
              <td className={styles.numCell}>Rp {formatIdr(c.cost)}</td>
              <td className={styles.numCell}>Rp {formatIdr(c.netCost)}</td>
              <td className={styles.numCell}>Rp {formatIdr(c.currentBudget)}</td>
              <td className={styles.numCell}>Rp {formatIdr(c.grossRevenue)}</td>
              <td className={styles.numCell}>{formatRoi(c.roi)}</td>
              <td className={styles.numCell}>{formatIdr(c.skuOrders)}</td>
              <td className={styles.numCell}>
                {c.costPerOrder !== null ? `Rp ${formatIdr(c.costPerOrder)}` : '—'}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
