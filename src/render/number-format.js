/**
 * Excel number format codes ("#,##0.00", "0%", "$#,##0", "#,##0 €",
 * "mmm yy"…), as charts use them for their value axis, data labels and
 * dated categories. Covers what charts commonly carry: decimals, digit
 * grouping, percentages, thousands scaling, literal text and currency
 * signs, separate sections for negative numbers and zero, scientific
 * notation and dates; anything else falls back to plain numbers.
 */

const isNum = (v) => typeof v === 'number' && Number.isFinite(v);

const plain = new Map();
const intl = (min, max, grouping) => {
  const key = `${min}/${max}/${grouping}`;
  if (!plain.has(key)) plain.set(key, new Intl.NumberFormat(undefined, { minimumFractionDigits: min, maximumFractionDigits: max, useGrouping: grouping }));
  return plain.get(key);
};

/** Whether a code is Excel's "General" (or empty): numbers as they are. */
export const isGeneral = (code) => !code || /^\s*general\s*$/i.test(code);

/** The sections of a code (positive;negative;zero;text), split outside quotes. */
function sections(code) {
  const out = [''];
  let quoted = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') quoted = !quoted;
    if (ch === '\\' && !quoted) {
      out[out.length - 1] += ch + (code[++i] ?? '');
      continue;
    }
    if (ch === ';' && !quoted) out.push('');
    else out[out.length - 1] += ch;
  }
  return out;
}

/**
 * A section split into literal text and format characters: quoted text,
 * escaped characters and currency brackets ([$€-407]) become literals;
 * colours and conditions ([Red], [>=100]), padding (_x) and fills (*x) go.
 */
function tokens(section) {
  const out = [];
  const lit = (s) => (out.length && out[out.length - 1].lit != null ? (out[out.length - 1].lit += s) : out.push({ lit: s }));
  for (let i = 0; i < section.length; i++) {
    const ch = section[i];
    if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      lit(section.slice(i + 1, end < 0 ? undefined : end));
      i = end < 0 ? section.length : end;
    } else if (ch === '\\') lit(section[++i] ?? '');
    else if (ch === '_' || ch === '*') i++;
    else if (ch === '[') {
      const end = section.indexOf(']', i);
      const inner = section.slice(i + 1, end < 0 ? undefined : end);
      if (inner.startsWith('$')) lit(inner.slice(1).split('-')[0]);
      i = end < 0 ? section.length : end;
    } else out.push({ ch });
  }
  return out;
}

const DATE_CHARS = /[ydmhs]/i;

/** Whether a code shows dates or times (y, m, d, h or s outside quotes and brackets). */
export function isDateFormat(code) {
  if (isGeneral(code)) return false;
  const bare = tokens(sections(code)[0]).filter((t) => t.ch).map((t) => t.ch).join('');
  return DATE_CHARS.test(bare.replace(/e[+-]/gi, '')) && !/[0#?]/.test(bare.replace(/[ydmhs.:/\s-]/gi, ''));
}

/** An Excel date serial as a Date (UTC), in the 1900 or 1904 date system. */
export const serialDate = (v, date1904 = false) => new Date(Date.UTC(date1904 ? 1904 : 1899, date1904 ? 0 : 11, date1904 ? 1 : 30) + Math.round(v * 86400000));

const monthName = (d, style) => d.toLocaleDateString(undefined, { timeZone: 'UTC', month: style });
const dayName = (d, style) => d.toLocaleDateString(undefined, { timeZone: 'UTC', weekday: style });
const pad = (n, w) => String(n).padStart(w, '0');

/** A date serial in a date code ("d/m/yyyy", "mmm yy", "h:mm AM/PM"…). */
function formatDate(v, list, date1904) {
  const d = serialDate(v, date1904);
  // AM/PM markers, then runs of the same letter as one part ("mmm").
  const parts = [];
  for (let i = 0; i < list.length; i++) {
    const t = list[i];
    const ahead = list.slice(i, i + 5).map((q) => q.ch || '\0').join('');
    const marker = /^am\/pm/i.test(ahead) ? 5 : /^a\/p/i.test(ahead) ? 3 : 0;
    const last = parts[parts.length - 1];
    if (marker) {
      parts.push({ ampm: marker === 5 });
      i += marker - 1;
    } else if (t.ch && DATE_CHARS.test(t.ch) && last?.run && last.run[0].toLowerCase() === t.ch.toLowerCase()) last.run += t.ch;
    else parts.push(t.ch && DATE_CHARS.test(t.ch) ? { run: t.ch } : t);
  }
  const ampm = parts.some((p) => p.ampm != null);
  const runs = parts.filter((p) => p.run);
  let out = '';
  for (const p of parts) {
    if (p.lit != null) out += p.lit;
    else if (p.ampm != null) out += (d.getUTCHours() < 12 ? 'AM' : 'PM').slice(0, p.ampm ? 2 : 1);
    if (!p.run) {
      if (p.ch) out += p.ch;
      continue;
    }
    const r = p.run.toLowerCase();
    const n = r.length;
    const k = runs.indexOf(p);
    if (r[0] === 'y') out += n <= 2 ? pad(d.getUTCFullYear() % 100, 2) : d.getUTCFullYear();
    else if (r[0] === 'd') out += n === 1 ? d.getUTCDate() : n === 2 ? pad(d.getUTCDate(), 2) : dayName(d, n === 3 ? 'short' : 'long');
    else if (r[0] === 'h') {
      const h = ampm ? d.getUTCHours() % 12 || 12 : d.getUTCHours();
      out += n === 1 ? h : pad(h, 2);
    } else if (r[0] === 's') out += pad(d.getUTCSeconds(), 2);
    else if (r[0] === 'm') {
      const minutes = n <= 2 && (runs[k - 1]?.run[0].toLowerCase() === 'h' || runs[k + 1]?.run[0].toLowerCase() === 's');
      if (minutes) out += n === 1 ? d.getUTCMinutes() : pad(d.getUTCMinutes(), 2);
      else out += n === 1 ? d.getUTCMonth() + 1 : n === 2 ? pad(d.getUTCMonth() + 1, 2) : n === 5 ? monthName(d, 'narrow') : monthName(d, n === 3 ? 'short' : 'long');
    }
  }
  return out;
}

/** A number in one section of a code: literal text around a "#,##0.00"-like pattern. */
function formatSection(v, list) {
  const first = list.findIndex((t) => t.ch && /[0#?]/.test(t.ch));
  if (first < 0) return list.map((t) => t.lit ?? t.ch).join('').replace(/@/g, '');
  // The pattern: digits, grouping commas, the decimal point and an exponent.
  let end = first;
  while (end < list.length && list[end].ch && /[0#?,.]/.test(list[end].ch)) end++;
  let exp = null;
  if (list[end]?.ch && /e/i.test(list[end].ch) && /[+-]/.test(list[end + 1]?.ch || '')) {
    let e = end + 2;
    while (list[e]?.ch && /[0#?]/.test(list[e].ch)) e++;
    exp = { sign: list[end + 1].ch, digits: e - end - 2 };
    end = e;
  }
  let pattern = list.slice(first, end).map((t) => t.ch).join('');
  // Commas at the end of the integer part divide by a thousand each.
  const scale = /,+(?=\.|$)/.exec(pattern);
  if (scale) {
    v /= 1000 ** scale[0].length;
    pattern = pattern.replace(/,+(?=\.|$)/, '');
  }
  const [int, frac = ''] = pattern.split('.');
  if (list.some((t) => t.ch === '%')) v *= 100;
  // The sign goes before any currency sign or text ("-$1,234").
  const sign = v < 0 ? '-' : '';
  v = Math.abs(v);
  const min = (frac.match(/0/g) || []).length;
  const max = (frac.match(/[0#?]/g) || []).length;
  let digits;
  if (exp) {
    const [mant, power] = v.toExponential(max).split('e');
    digits = `${mant}E${+power < 0 ? '-' : exp.sign === '+' ? '+' : ''}${pad(Math.abs(+power), exp.digits)}`;
  } else {
    digits = intl(min, max, int.includes(',')).format(v);
    // "#" leaves out a lone zero before the point ("#.00" shows .5).
    if (!/0/.test(int)) digits = digits.replace(/^0(?=\D|$)/, '');
  }
  const before = list.slice(0, first).map((t) => t.lit ?? t.ch).join('');
  const after = list.slice(end).map((t) => t.lit ?? t.ch).join('');
  return sign + before + digits + after;
}

/**
 * A number in an Excel format code. "General" (or no code) shows it with up
 * to `digits` decimals; `date1904` picks the date system for dates.
 */
export function formatValue(v, code, { digits = 2, date1904 = false } = {}) {
  if (!isNum(v)) return '';
  if (isGeneral(code)) return intl(0, digits, true).format(v);
  const parts = sections(code);
  if (isDateFormat(code)) return formatDate(v, tokens(parts[0]), date1904);
  let section = parts[0];
  let n = v;
  if (v < 0 && parts.length > 1 && parts[1] !== '') {
    // The negative section shows the number without its sign ("(#,##0)").
    section = parts[1];
    n = -v;
  } else if (v === 0 && parts.length > 2 && parts[2] !== '') section = parts[2];
  if (isGeneral(section)) return intl(0, digits, true).format(n);
  return formatSection(n, tokens(section));
}

/** Whether a code shows percentages (its numbers times a hundred). */
export const isPercentFormat = (code) => !isGeneral(code) && tokens(sections(code)[0]).some((t) => t.ch === '%');
