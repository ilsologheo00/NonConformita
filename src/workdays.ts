/** Easter Sunday of a year (anonymous Gregorian algorithm), as [month 0-11, day]. */
function easter(y: number): [number, number] {
  const a = y % 19;
  const b = Math.floor(y / 100);
  const c = y % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return [month - 1, day];
}

// Italian national holidays, as "month-day" (month 1-12).
const FIXED = new Set(["1-1", "1-6", "4-25", "5-1", "6-2", "8-15", "11-1", "12-8", "12-25", "12-26"]);
const mondays = new Map<number, string>();

function isHoliday(d: Date): boolean {
  if (FIXED.has(`${d.getMonth() + 1}-${d.getDate()}`)) return true;
  const y = d.getFullYear();
  let monday = mondays.get(y);
  if (!monday) {
    const [m, day] = easter(y);
    const em = new Date(y, m, day + 1);
    monday = `${em.getMonth()}-${em.getDate()}`;
    mondays.set(y, monday);
  }
  return monday === `${d.getMonth()}-${d.getDate()}`;
}

const isWorkday = (d: Date) => d.getDay() !== 0 && d.getDay() !== 6 && !isHoliday(d);

/** The local calendar day of an ISO date ("2026-09-02" or a SharePoint "2026-09-01T22:00:00Z"). */
function day(iso: string): Date | null {
  const plain = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
  const d = plain ? new Date(+plain[1], +plain[2] - 1, +plain[3]) : new Date(iso);
  if (isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/**
 * Working days from one date to another: the days after "from" up to and including "to",
 * leaving out Saturdays, Sundays and Italian national holidays. Same day -> 0; to before from -> negative.
 */
export function workdaysBetween(fromIso: string, toIso: string): number | null {
  const from = day(fromIso);
  const to = day(toIso);
  if (!from || !to) return null;
  const sign = to < from ? -1 : 1;
  const [start, end] = sign > 0 ? [from, to] : [to, from];
  let n = 0;
  for (const d = new Date(start); d < end; ) {
    d.setDate(d.getDate() + 1);
    if (isWorkday(d)) n++;
  }
  return sign * n;
}
