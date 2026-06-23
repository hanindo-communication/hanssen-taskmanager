import * as XLSX from 'xlsx';
import { readFileSync, existsSync } from 'fs';

const FILENAME_RANGE_RE = /(\d{4}-\d{2}-\d{2})\s*-\s*(\d{4}-\d{2}-\d{2})/i;

function parseDateRangeFromFilename(fileName) {
  const m = FILENAME_RANGE_RE.exec(fileName);
  if (!m) return null;
  return { start: m[1], end: m[2] };
}

function daysInclusive(startIso, endIso) {
  const start = new Date(`${startIso}T00:00:00Z`);
  const end = new Date(`${endIso}T00:00:00Z`);
  return Math.floor((end - start) / 86400000) + 1;
}

function isExactlySevenDays(startIso, endIso) {
  return daysInclusive(startIso, endIso) === 7;
}

function buildCompleteWeekWindows(startIso, endIso) {
  const windows = [];
  let current = new Date(`${startIso}T00:00:00Z`);
  const endDay = new Date(`${endIso}T00:00:00Z`);

  while (current <= endDay) {
    const windowEnd = new Date(current);
    windowEnd.setUTCDate(windowEnd.getUTCDate() + 6);
    if (windowEnd > endDay) {
      const dayCount = daysInclusive(
        current.toISOString().slice(0, 10),
        endDay.toISOString().slice(0, 10)
      );
      if (dayCount < 7) {
        return {
          windows,
          excludedPartial: {
            from: current.toISOString().slice(0, 10),
            to: endDay.toISOString().slice(0, 10),
            dayCount,
          },
        };
      }
      windows.push({
        start: current.toISOString().slice(0, 10),
        end: endDay.toISOString().slice(0, 10),
      });
      return { windows };
    }
    windows.push({
      start: current.toISOString().slice(0, 10),
      end: windowEnd.toISOString().slice(0, 10),
    });
    current = new Date(windowEnd);
    current.setUTCDate(current.getUTCDate() + 1);
  }
  return { windows };
}

const samplePath =
  'C:/Users/hanss/Downloads/Product campaign data 2026-06-01 - 2026-06-22.xlsx';

console.log('--- Window logic ---');
const w = buildCompleteWeekWindows('2026-06-01', '2026-06-22');
console.log(JSON.stringify(w, null, 2));

console.log('\n--- Aggregate file 1-22 (should reject as weekly) ---');
const aggRange = parseDateRangeFromFilename('Product campaign data 2026-06-01 - 2026-06-22.xlsx');
console.log('range:', aggRange, 'exactly7:', isExactlySevenDays(aggRange.start, aggRange.end));

console.log('\n--- Weekly files ---');
for (const name of [
  'Product campaign data 2026-06-01 - 2026-06-07.xlsx',
  'Product campaign data 2026-06-08 - 2026-06-14.xlsx',
  'Product campaign data 2026-06-15 - 2026-06-21.xlsx',
  'Product campaign data 2026-06-15 - 2026-06-22.xlsx',
]) {
  const r = parseDateRangeFromFilename(name);
  console.log(name, '=>', r, 'valid:', r ? isExactlySevenDays(r.start, r.end) : false);
}

if (existsSync(samplePath)) {
  const wb = XLSX.read(readFileSync(samplePath));
  const sheetName = wb.SheetNames.find((n) => n.toLowerCase() === 'data') ?? wb.SheetNames[0];
  const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '' });
  console.log('\n--- Sample file rows:', raw.length - 1, 'campaigns ---');
} else {
  console.log('\nSample file not found, skipped xlsx read.');
}
