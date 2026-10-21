/**
 * PowerPoint's 74 built-in table styles, for files that refer to one by its
 * id without including it in ppt/tableStyles.xml (PowerPoint only writes
 * the ones a deck uses, and other programs often write none). Each family
 * (Themed, Light, Medium and Dark styles 1 to 4) is a pattern of fills,
 * borders and text colours in an accent colour (or dark 1 / text 1 for the
 * plain one); the parameters follow the styles as PowerPoint defines them,
 * with LibreOffice's map of them (oox/source/drawingml/table/
 * predefined-table-styles.cxx) as a reference. The styles are written as
 * a:tblStyle XML, read like any other (see tableLook in pptx-import.js).
 */

const ACCENTS = ['', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];

/** Style ids by family, in accent order (none, accent 1…6). */
const IDS = {
  themed1: ['2D5ABB26-0587-4C30-8999-92F81FD0307C', '3C2FFA5D-87B4-456A-9821-1D502468CF0F', '284E427A-3D55-4303-BF80-6455036E1DE7', '69C7853C-536D-4A76-A0AE-DD22124D55A5', '775DCB02-9BB8-47FD-8907-85C794F793BA', '35758FB7-9AC5-4552-8A53-C91805E547FA', '08FB837D-C827-4EFA-A057-4D05807E0F7C'],
  themed2: ['5940675A-B579-460E-94D1-54222C63F5DA', 'D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7', '18603FDC-E32A-4AB5-989C-0864C3EAD2B8', '306799F8-075E-4A3A-A7F6-7FBC6576F1A4', 'E269D01E-BC32-4049-B463-5C60D7B0CCD2', '327F97BB-C833-4FB7-BDE5-3F7075034690', '638B1855-1B75-4FBE-930C-398BA8C253C6'],
  light1: ['9D7B26C5-4107-4FEC-AEDC-1716B250A1EF', '3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5', '0E3FDE45-AF77-4B5C-9715-49D594BDF05E', 'C083E6E3-FA7D-4D7B-A595-EF9225AFEA82', 'D27102A9-8310-4765-A935-A1911B00CA55', '5FD0F851-EC5A-4D38-B0AD-8093EC10F338', '68D230F3-CF80-4859-8CE7-A43EE81993B5'],
  light2: ['7E9639D4-E3E2-4D34-9284-5A2195B3D0D7', '69012ECD-51FC-41F1-AA8D-1B2483CD663E', '72833802-FEF1-4C79-8D5D-14CF1EAF98D9', 'F2DE63D5-997A-4646-A377-4702673A728D', '17292A2E-F333-43FB-9621-5CBBE7FDCDCB', '5A111915-BE36-4E01-A7E5-04B1672EAD32', '912C8C85-51F0-491E-9774-3900AFEF0FD7'],
  light3: ['616DA210-FB5B-4158-B5E0-FEB733F419BA', 'BC89EF96-8CEA-46FF-86C4-4CE0E7609802', '5DA37D80-6434-44D0-A028-1B22A696006F', '8799B23B-EC83-4686-B30A-512413B5E67A', 'ED083AE6-46FA-4A59-8FB0-9F97EB10719F', 'BDBED569-4797-4DF1-A0F4-6AAB3CD982D8', 'E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3'],
  medium1: ['793D81CF-94F2-401A-BA57-92F5A7B2D0C5', 'B301B821-A1FF-4177-AEE7-76D212191A09', '9DCAF9ED-07DC-4A11-8D7F-57B35C25682E', '1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90', '1E171933-4619-4E11-9A3F-F7608DF75F80', 'FABFCF23-3B69-468F-B69F-88F6DE6A72F2', '10A1B5D5-9B99-4C35-A422-299274C87663'],
  medium2: ['073A0DAA-6AF3-43AB-8588-CEC1D06C72B9', '5C22544A-7EE6-4342-B048-85BDC9FD1C3A', '21E4AEA4-8DFA-4A89-87EB-49C32662AFE0', 'F5AB1C69-6EDB-4FF4-983F-18BD219EF322', '00A15C55-8517-42AA-B614-E9B94910E393', '7DF18680-E054-41AD-8BC1-D1AEF772440D', '93296810-A885-4BE3-A3E7-6D5BEEA58F35'],
  medium3: ['8EC20E35-A176-4012-BC5E-935CFFF8708E', '6E25E649-3F16-4E02-A733-19D2CDBF48F0', '85BE263C-DBD7-4A20-BB59-AAB30ACAA65A', 'EB344D84-9AFB-497E-A393-DC336BA19D2E', 'EB9631B5-78F2-41C9-869B-9F39066F8104', '74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6', '2A488322-F2BA-4B5B-9748-0D474271808F'],
  medium4: ['D7AC3CCA-C797-4891-BE02-D94E43425B78', '69CF1AB2-1976-4502-BF36-3FF5EA218861', '8A107856-5554-42FB-B03E-39F5DBC370BA', '0505E3EF-67EA-436B-97B2-0124C06EBD24', 'C4B1156A-380E-4F78-BDF5-A606A8083BF9', '22838BEF-8BB2-4498-84A7-C5851F593DF1', '16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4'],
  dark1: ['E8034E78-7F5D-4C2E-B375-FC64B27BC917', '125E5076-3810-47DD-B79F-674D7AD40C01', '37CE84F3-28C3-443E-9E96-99CF82512B78', 'D03447BB-5D67-496B-8E87-E561075AD55C', 'E929F9F4-4A8F-4326-A1B4-22849713DDAB', '8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B', 'AF606853-7671-496A-8E4F-DF71F8EC918B'],
  // Dark Style 2 comes plain and in accent pairs: 1/2, 3/4 and 5/6.
  dark2: ['5202B0CA-FC54-4496-8BCA-5EF66A818D29', '0660B408-B3CF-4A94-85FC-2B1E0A45F4A2', null, '91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28', null, '46F890A9-2807-4EBB-B81D-B2AA78EC7F39', null],
};

const BY_ID = new Map(Object.entries(IDS).flatMap(([family, ids]) => ids.map((id, i) => id && [`{${id}}`, [family, i]]).filter(Boolean)));

/**
 * The parts of each family, for accent `a` ('' for the plain style): text
 * colour, bold, fill and borders ({ side: [colour, width in EMU] }), colours
 * being [scheme colour, modifier?, value?].
 */
const SIDES = ['left', 'right', 'top', 'bottom', 'insideH', 'insideV'];
const OUTER = ['left', 'right', 'top', 'bottom'];
const all = (sides, color, w) => Object.fromEntries(sides.map((s) => [s, [color, w]]));
const bold = { bold: true };

const FAMILIES = {
  themed1: (a) => (a ? {
    wholeTbl: { text: ['dk1'], borders: all(SIDES, [a]) },
    band1H: { fill: [a, 'alpha', 40000] }, band1V: { fill: [a, 'alpha', 40000] },
    lastCol: { ...bold, borders: all([...OUTER, 'insideH'], [a]) },
    firstCol: { ...bold, borders: all([...OUTER, 'insideH'], [a]) },
    lastRow: { ...bold, borders: all(OUTER, [a]) },
    firstRow: { ...bold, text: ['lt1'], fill: [a], borders: { ...all(['left', 'right', 'top'], [a]), bottom: [['lt1']] } },
  } : { wholeTbl: { text: ['tx1'] }, firstRow: bold, lastRow: bold, firstCol: bold, lastCol: bold }),
  themed2: (a) => (a ? {
    wholeTbl: { text: ['lt1'], fill: [a], borders: all(OUTER, [a, 'tint', 50000]) },
    band1H: { fill: ['lt1', 'alpha', 20000] }, band1V: { fill: ['lt1', 'alpha', 20000] },
    lastCol: { ...bold, borders: { left: [['lt1']] } },
    firstCol: { ...bold, borders: { right: [['lt1']] } },
    lastRow: { ...bold, borders: { top: [['lt1']] } },
    firstRow: { ...bold, text: ['lt1'], borders: { bottom: [['lt1']] } },
  } : {
    wholeTbl: { text: ['tx1'], borders: { ...all(OUTER, ['tx1']), ...all(['insideH', 'insideV'], ['tx1']) } },
    firstRow: bold, lastRow: bold, firstCol: bold, lastCol: bold,
  }),
  light1: (a, c = a || 'tx1') => ({
    wholeTbl: { text: ['tx1'], borders: all(['top', 'bottom'], [c]) },
    band1H: { fill: [c, 'alpha', 20000] }, band1V: { fill: [c, 'alpha', 20000] },
    lastCol: bold, firstCol: bold,
    lastRow: { ...bold, borders: { top: [[c]] } },
    firstRow: { ...bold, borders: { bottom: [[c]] } },
  }),
  light2: (a, c = a || 'tx1') => ({
    wholeTbl: { text: ['tx1'], borders: all(OUTER, [c]) },
    band1H: { borders: all(['top', 'bottom'], [c]) },
    band1V: { borders: all(['left', 'right'], [c]) },
    band2V: { borders: all(['left', 'right'], [c]) },
    lastCol: bold, firstCol: bold,
    lastRow: { ...bold, borders: { top: [[c], 50800] } },
    firstRow: { ...bold, text: ['bg1'], fill: [c] },
  }),
  light3: (a, c = a || 'tx1') => ({
    wholeTbl: { text: ['tx1'], borders: all(SIDES, [c]) },
    band1H: { fill: [c, 'alpha', 20000] }, band1V: { fill: [c, 'alpha', 20000] },
    lastCol: bold, firstCol: bold,
    lastRow: { ...bold, borders: { top: [[c], 50800] } },
    firstRow: { ...bold, text: [c], borders: { bottom: [[c]] } },
  }),
  medium1: (a, c = a || 'dk1') => ({
    wholeTbl: { text: ['dk1'], fill: ['lt1'], borders: all([...OUTER, 'insideH'], [c]) },
    band1H: { fill: [c, 'tint', 20000] }, band1V: { fill: [c, 'tint', 20000] },
    lastCol: bold, firstCol: bold,
    lastRow: { ...bold, fill: ['lt1'], borders: { top: [[c], 50800] } },
    firstRow: { ...bold, text: ['lt1'], fill: [c] },
  }),
  medium2: (a, c = a || 'dk1') => ({
    wholeTbl: { text: ['dk1'], fill: [c, 'tint', 20000], borders: all(SIDES, ['lt1']) },
    band1H: { fill: [c, 'tint', 40000] }, band1V: { fill: [c, 'tint', 40000] },
    lastCol: { ...bold, text: ['lt1'], fill: [c] },
    firstCol: { ...bold, text: ['lt1'], fill: [c] },
    lastRow: { ...bold, text: ['lt1'], fill: [c], borders: { top: [['lt1'], 38100] } },
    firstRow: { ...bold, text: ['lt1'], fill: [c], borders: { bottom: [['lt1'], 38100] } },
  }),
  medium3: (a, c = a || 'dk1') => ({
    wholeTbl: { text: ['dk1'], fill: ['lt1'], borders: all(['top', 'bottom'], ['dk1'], 25400) },
    band1H: { fill: ['dk1', 'tint', 20000] }, band1V: { fill: ['dk1', 'tint', 20000] },
    lastCol: { ...bold, text: ['lt1'], fill: [c] },
    firstCol: { ...bold, text: ['lt1'], fill: [c] },
    lastRow: { ...bold, fill: ['lt1'], borders: { top: [['dk1'], 50800] } },
    firstRow: { ...bold, text: ['lt1'], fill: [c], borders: { bottom: [['dk1'], 25400] } },
  }),
  medium4: (a, c = a || 'dk1') => ({
    wholeTbl: { text: ['dk1'], fill: [c, 'tint', 20000], borders: all(SIDES, [c]) },
    band1H: { fill: [c, 'tint', 40000] }, band1V: { fill: [c, 'tint', 40000] },
    lastCol: bold, firstCol: bold,
    lastRow: { ...bold, fill: ['dk1', 'tint', 20000], borders: { top: [['dk1'], 25400] } },
    firstRow: { ...bold, text: [c], fill: [c, 'tint', 20000] },
  }),
  dark1: (a) => {
    // Shades of the accent, or tints of dark 1, with light text.
    const c = (v) => (a ? [a, 'shade', v] : ['dk1', 'tint', v]);
    return {
      wholeTbl: { text: ['lt1'], fill: c(20000) },
      band1H: { fill: c(40000) }, band1V: { fill: c(40000) },
      lastCol: { ...bold, fill: c(60000), borders: { left: [['lt1'], 25400] } },
      firstCol: { ...bold, fill: c(60000), borders: { right: [['lt1'], 25400] } },
      lastRow: { ...bold, fill: [a || 'dk1'], borders: { top: [['lt1'], 25400] } },
      firstRow: { ...bold, fill: ['dk1'], borders: { bottom: [['lt1'], 38100] } },
    };
  },
  dark2: (a, c = a || 'dk1') => ({
    wholeTbl: { text: ['dk1'], fill: [c, 'tint', 20000] },
    band1H: { fill: [c, 'tint', 40000] }, band1V: { fill: [c, 'tint', 40000] },
    lastCol: bold, firstCol: bold,
    lastRow: { ...bold, fill: [c, 'tint', 20000], borders: { top: [['dk1'], 50800] } },
    // The header is in the accent's partner (2, 4 or 6).
    firstRow: { ...bold, text: ['lt1'], fill: [a ? `accent${+a.slice(-1) + 1}` : 'dk1'] },
  }),
};

const colorXml = ([name, mod, val]) => `<a:schemeClr val="${name}">${mod ? `<a:${mod} val="${val}"/>` : ''}</a:schemeClr>`;

function partXml(name, { text, bold: b, fill, borders = {} }) {
  const tx = text || b ? `<a:tcTxStyle${b ? ' b="on"' : ''}>${text ? colorXml(text) : ''}</a:tcTxStyle>` : '';
  const bdr = Object.entries(borders).map(([side, [color, w = 12700]]) => `<a:${side}><a:ln w="${w}"${w === 50800 ? ' cmpd="dbl"' : ''}><a:solidFill>${colorXml(color)}</a:solidFill></a:ln></a:${side}>`).join('');
  return `<a:${name}>${tx}<a:tcStyle><a:tcBdr>${bdr}</a:tcBdr>${fill ? `<a:fill><a:solidFill>${colorXml(fill)}</a:solidFill></a:fill>` : ''}</a:tcStyle></a:${name}>`;
}

/** A built-in style's a:tblStyle XML, or null for an id that isn't one. */
export function builtInTableStyleXml(id) {
  const hit = BY_ID.get(String(id || '').toUpperCase());
  if (!hit) return null;
  const [family, i] = hit;
  const parts = FAMILIES[family](ACCENTS[i]);
  const body = Object.entries(parts).map(([name, p]) => partXml(name, p)).join('');
  return `<a:tblStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" styleId="${String(id).toUpperCase()}">${body}</a:tblStyle>`;
}

const parsed = new Map();
/** A built-in style as a parsed a:tblStyle element (parsed once), or null. */
export function builtInTableStyle(id) {
  const key = String(id || '').toUpperCase();
  if (!parsed.has(key)) {
    const text = builtInTableStyleXml(key);
    parsed.set(key, text ? new DOMParser().parseFromString(text, 'application/xml').documentElement : null);
  }
  return parsed.get(key);
}

/** How many built-in styles there are (74). */
export const BUILT_IN_TABLE_STYLES = BY_ID.size;
