import { describe, it, expect, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { mtxToSfnt, lzcompDecode } from '../../src/io/mtx.js';
import { unwrapEot, decodeFontData, sfntTables, checksum, fontNames, toBase64, EOT_XOR, EOT_COMPRESSED } from '../../src/io/font-data.js';
import { insertFontList } from '../../src/io/pptx-fonts.js';
import { buildPptx } from '../../src/io/pptx-export.js';
import { importPptx } from '../../src/io/pptx-import.js';
import { newDeck, newTextBox } from '../../src/model/deck.js';

// A font drawn for these tests, MicroType Express compressed in an EOT the
// way PowerPoint embeds fonts (see tests/fixtures/README.md).
const fixture = () => new Uint8Array(readFileSync(resolve('tests/fixtures/mtx-test.fntdata')));
const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex');
const u16 = (b, o) => (b[o] << 8) | b[o + 1];
const i16 = (b, o) => (u16(b, o) << 16) >> 16;
const hex = (b) => Buffer.from(b).toString('hex');

// The tables of the font the fixture was made from (SHA-256). libeot's
// EOT2ttf_buffer, an independent MTX decoder, unpacks the fixture to the same
// tables but glyf and loca (which it writes differently, with the same
// outlines; it also garbles composite glyphs' instructions) and head (it
// leaves checkSumAdjustment at 0).
const ORIGINAL_TABLES = {
  'OS/2': 'df6697fe5e0b1ddbb7c7ff2c141e4dad0f26a72eec71a37b87f7ea4236585f94',
  cmap: '8857f0cd4fbe0ca99caf14cd23a5df56520f3b871e9e70d71d5d67a2dc952578',
  'cvt ': '94c14e6a36a35f8f3a0a07c8d5a4e28c591668fec6c473cdad8578e9973b2046',
  fpgm: 'c321191e8f162343a27b3fb1b8575df29bb183ee29882760f0ce039a79c96f4c',
  glyf: '4f4e9ad0c89b50304b8b3f89e4f94eb88140e60fb9da512ecc135e3e77ce5a2d',
  hhea: 'f9753872b48e710b8f396d92bfce6d2650286e91ab5d9e6ab3b7d54c912264d1',
  hmtx: '886b9427ccd52dbd0db2984da40ad16b702ebc0f400897e8635b93816361dd16',
  loca: '557e21cd53f99cf594292f6bdc7cb228280b80d162686b2ae757c8c4b07e4080',
  maxp: 'a4cd858f91104a6779c48974479cfd871aaf84f0b355c26a810c78dd5ddbd094',
  name: '1b2cf8bc3625d7e41cc9fbcaefbce7b45907df9818493294f846e6404d4f670e',
  post: 'a1620334253027d7f9258e7787194a94bbca3889ba2fd8da25725d4ef3e6cbc6',
  prep: 'c0b8677f730385ea92fb85717d803020000e264df48ad6ade1a38bd2ce71e0ac',
};
// The original head table with checkSumAdjustment zeroed (it depends on the whole file's layout).
const ORIGINAL_HEAD = 'da3ce98c83e60f6a5f6aaec277527c3fbf57e6ee8224639c456656dc8d46c3ed';

/** Each glyph's record from glyf, by loca (short offsets). */
function glyphRecords(tables) {
  const glyf = tables.get('glyf');
  const loca = tables.get('loca');
  return [...Array(loca.length / 2 - 1).keys()].map((g) => glyf.subarray(u16(loca, g * 2) * 2, u16(loca, g * 2 + 2) * 2));
}

/** A simple glyph's contour count, bounding box, point count and instructions. */
function simpleGlyph(g) {
  const contours = i16(g, 0);
  const points = u16(g, 10 + (contours - 1) * 2) + 1;
  const at = 10 + contours * 2;
  return { contours, bbox: [2, 4, 6, 8].map((o) => i16(g, o)), points, instructions: hex(g.subarray(at + 2, at + 2 + u16(g, at))) };
}

describe('MicroType Express', () => {
  it('unpacks a compressed EOT into the font it was made from', () => {
    const { data, flags, family } = unwrapEot(fixture());
    expect(flags).toBe(EOT_COMPRESSED);
    expect(family).toBe('MTX Test');
    const tables = sfntTables(data);
    expect([...tables.keys()].sort()).toEqual([...Object.keys(ORIGINAL_TABLES), 'head'].sort());
    for (const [tag, digest] of Object.entries(ORIGINAL_TABLES)) expect(sha256(tables.get(tag)), tag).toBe(digest);
    const head = tables.get('head').slice();
    head.fill(0, 8, 12);
    expect(sha256(head)).toBe(ORIGINAL_HEAD);
    expect(checksum(data)).toBe(0xb1b0afba);
    expect(fontNames(tables)).toMatchObject({ family: 'MTX Test', style: 'Regular' });
    expect(decodeFontData(fixture()).mime).toBe('font/ttf');
  });

  it('rebuilds glyph outlines, bounding boxes and instructions', () => {
    const glyphs = glyphRecords(sfntTables(mtxToSfnt(mtxData())));
    // .notdef, space, A, B, C, D, acute, Aacute, Adieresis, E
    expect(glyphs.length).toBe(10);
    expect(glyphs[1].length).toBe(0);
    expect(simpleGlyph(glyphs[0])).toEqual({ contours: 2, bbox: [50, 0, 450, 700], points: 8, instructions: '4010050602010407030005040203060701002f3cdc3c2f3cdc3c2f3cdc3c2f3cdc3c' });
    // Moves of every size, the largest needing 16 bits.
    expect(simpleGlyph(glyphs[2])).toEqual({ contours: 2, bbox: [-4000, -600, 1299, 5003], points: 19, instructions: 'b001b9012cfffeb0072b0001' });
    // A bounding box bigger than the glyph's points is kept.
    expect(simpleGlyph(glyphs[3]).bbox).toEqual([50, -20, 560, 760]);
    expect(simpleGlyph(glyphs[4])).toMatchObject({ contours: 1, bbox: [0, 0, 600, 400], points: 600 });
    // Pushes with words, negative numbers and repeats; then the rest of the code.
    expect(simpleGlyph(glyphs[5]).instructions).toBe(
      '400a09010902090309fbfcfdbb010401f301fe02f34024040004080400040202020202020202020202020202020202020202020202020202020202bcffffff0680007fff03e84b2b1c',
    );
    // A composite glyph's instructions: a byte, a word and a byte pushed, then CALL.
    expect(i16(glyphs[7], 0)).toBe(-1);
    expect(hex(glyphs[7])).toMatch(/0009b10101b8ffecb0352b00$/);
  });

  it('undoes XOR obfuscation before unpacking', () => {
    const eot = fixture();
    const v = new DataView(eot.buffer);
    v.setUint32(12, EOT_COMPRESSED | EOT_XOR, true);
    const start = eot.length - v.getUint32(4, true);
    for (let i = start; i < eot.length; i++) eot[i] ^= 0x50;
    expect(unwrapEot(eot).data).toEqual(unwrapEot(fixture()).data);
  });

  it('unpacks an empty LZCOMP block, and run-length encoded ones', () => {
    expect(lzcompDecode(new Uint8Array(4)).length).toBe(0);
    // The fixture's push data block is run-length encoded (its first bit says so).
    const mtx = mtxData();
    const second = (mtx[4] << 16) | (mtx[5] << 8) | mtx[6];
    const third = (mtx[7] << 16) | (mtx[8] << 8) | mtx[9];
    expect(mtx[second] >> 7).toBe(1);
    const data = lzcompDecode(mtx.subarray(second, third));
    expect(data.length).toBe(85);
    expect(hex(data)).toMatch(/^0506020104070300050402030607010001ff32fa0207.*fc02020202fa01faff00fd8000fd7ffffd03e8030101fa1435$/);
  });

  it('turns down data that isn’t MTX or is damaged, without hanging', () => {
    const data = mtxData();
    expect(() => mtxToSfnt(new Uint8Array(4))).toThrow(/MicroType Express/);
    const version1 = data.slice();
    version1[0] = 1;
    expect(() => mtxToSfnt(version1)).toThrow(/MicroType Express/);
    const badOffsets = data.slice();
    badOffsets.set([0xff, 0xff, 0xff], 7);
    expect(() => mtxToSfnt(badOffsets)).toThrow(/Damaged/);
    // Every truncation fails cleanly.
    for (let n = 0; n < data.length; n += 7) expect(() => mtxToSfnt(data.subarray(0, n)), `${n} bytes`).toThrow(Error);
    // Changed bytes either fail cleanly or still give a font; nothing runs away.
    const started = Date.now();
    let seed = 1;
    const random = () => (seed = (seed * 1103515245 + 12345) % 2 ** 31) / 2 ** 31;
    for (let i = 0; i < 300; i++) {
      const bad = data.slice();
      for (let k = 0; k < 1 + (i % 4); k++) bad[10 + Math.floor(random() * (bad.length - 10))] = Math.floor(random() * 256);
      try {
        expect(sfntTables(mtxToSfnt(bad)).size).toBeGreaterThan(0);
      } catch (err) {
        expect(err).toBeInstanceOf(Error);
      }
    }
    expect(Date.now() - started).toBeLessThan(10000);
  });

  it('reports a compressed font it can’t unpack with code mtx', () => {
    const eot = fixture();
    const start = eot.length - new DataView(eot.buffer).getUint32(4, true);
    eot.fill(0, start + 10, start + 40);
    expect(() => unwrapEot(eot)).toThrow(expect.objectContaining({ code: 'mtx', family: 'MTX Test' }));
  });
});

/** The fixture's MTX data (after the EOT header). */
function mtxData() {
  const eot = fixture();
  return eot.subarray(eot.length - new DataView(eot.buffer).getUint32(4, true));
}

describe('PowerPoint files with compressed fonts', () => {
  it('brings a MicroType Express font into the deck', async () => {
    const deck = newDeck({ slides: [] });
    const box = newTextBox(10, 10);
    box.text = '<p>x</p>';
    box.style.font = 'MTX Test';
    deck.slides.push({ id: 's1', layout: 'blank', background: null, elements: [box], notes: '', transition: { type: 'none', duration: 0.6 }, hidden: false });
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {}); // the bundled fonts can't be fetched here
    const blob = await buildPptx(deck);
    const zip = await JSZip.loadAsync(await (blob.arrayBuffer ? blob.arrayBuffer() : new Response(blob).arrayBuffer()));
    zip.file('ppt/fonts/font1.fntdata', fixture());
    const relsPath = 'ppt/_rels/presentation.xml.rels';
    const rels = await zip.file(relsPath).async('string');
    zip.file(relsPath, rels.replace('</Relationships>', '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font1.fntdata"/></Relationships>'));
    const pres = await zip.file('ppt/presentation.xml').async('string');
    zip.file('ppt/presentation.xml', insertFontList(pres, '<p:embeddedFontLst><p:embeddedFont><p:font typeface="MTX Test" pitchFamily="2" charset="0"/><p:regular r:id="rId99"/></p:embeddedFont></p:embeddedFontLst>'));

    const back = await importPptx(await zip.generateAsync({ type: 'arraybuffer' }), 'x');
    expect(warn.mock.calls.some(([msg]) => /MicroType Express/.test(msg))).toBe(false);
    warn.mockRestore();
    expect(back.fonts['MTX Test']).toEqual({ regular: `data:font/ttf;base64,${toBase64(unwrapEot(fixture()).data)}` });
    expect(back.slides[0].elements[0].style.font).toBe('MTX Test');
  });
});
