import { describe, it, expect, vi, beforeAll, afterEach } from 'vitest';

vi.mock('../../src/storage/db.js', () => ({
  getDeck: async () => null,
  saveDeck: async (deck) => ({ rev: (deck.rev || 0) + 1 }),
  onStorageChange: () => () => {},
}));

import JSZip from 'jszip';
import { buildPptx } from '../../src/io/pptx-export.js';
import { importPptx } from '../../src/io/pptx-import.js';
import { xmlText } from '../../src/io/ooxml.js';
import { sanitizeDeck } from '../../src/io/sanitize.js';
import { renderThumb } from '../../src/render/render.js';
import { EditorScreen } from '../../src/editor/editor-screen.js';
import { newDeck, newSlide, newShape, cloneSlide } from '../../src/model/deck.js';

const toBuffer = (blob) => (blob.arrayBuffer ? blob.arrayBuffer() : new Response(blob).arrayBuffer());
const exported = async (deck) => JSZip.loadAsync(await toBuffer(await buildPptx(deck)));
const reimport = async (zip) => importPptx(await zip.generateAsync({ type: 'arraybuffer' }), 'x');

function deckWith(...elements) {
  const deck = newDeck({ slides: [] });
  deck.slides.push({ id: 's1', layout: 'blank', background: null, elements, notes: '', transition: { type: 'none', duration: 0.6 }, hidden: false });
  return deck;
}

/** Add a relationship to a .rels part of an exported package. */
async function addRel(zip, relsPath, id, type, target) {
  const xml = await zip.file(relsPath).async('string');
  zip.file(relsPath, xml.replace('</Relationships>', `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`));
}

const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

describe('PowerPoint comments: reading', () => {
  it('reads classic comments with their authors, dates, positions and threaded replies', async () => {
    const zip = await exported(deckWith());
    zip.file('ppt/commentAuthors.xml', `<?xml version="1.0"?><p:cmAuthorLst ${P}>
      <p:cmAuthor id="0" name="Ada Lovelace" initials="AL" lastIdx="2" clrIdx="0"/>
      <p:cmAuthor id="1" name="Charles &amp; Co" initials="CB" lastIdx="1" clrIdx="1"/></p:cmAuthorLst>`);
    zip.file('ppt/comments/comment1.xml', `<?xml version="1.0"?><p:cmLst ${P}>
      <p:cm authorId="0" dt="2024-03-05T10:20:30.500" idx="1"><p:pos x="800" y="400"/><p:text>Check &lt;this&gt; &amp; that</p:text>
        <p:extLst><p:ext uri="{C676402C-5697-4E1C-873F-D02D1690AC5C}"><p15:threadingInfo xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" timeZoneBias="-60"/></p:ext></p:extLst></p:cm>
      <p:cm authorId="1" dt="2024-03-05T11:00:00.000" idx="1"><p:pos x="800" y="400"/><p:text>Done</p:text>
        <p:extLst><p:ext uri="{C676402C-5697-4E1C-873F-D02D1690AC5C}"><p15:threadingInfo xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" timeZoneBias="0"><p15:parentCm authorId="0" idx="1"/></p15:threadingInfo></p:ext></p:extLst></p:cm>
      <p:cm authorId="0" idx="2"><p:pos x="0" y="0"/><p:text>Whole slide</p:text></p:cm>
      <p:cm authorId="7" dt="2024-03-06T09:00:00" idx="1"><p:pos x="80" y="16"/><p:text>Who?</p:text></p:cm></p:cmLst>`);
    await addRel(zip, 'ppt/_rels/presentation.xml.rels', 'rIdA', `${REL}/commentAuthors`, 'commentAuthors.xml');
    await addRel(zip, 'ppt/slides/_rels/slide1.xml.rels', 'rIdC', `${REL}/comments`, '../comments/comment1.xml');
    const [first, whole, unknown] = (await reimport(zip)).slides[0].comments;

    // Master units are 1/8 point; the dates were written at UTC+1 (a bias of −60 minutes).
    expect(first).toMatchObject({ author: 'Ada Lovelace', initials: 'AL', text: 'Check <this> & that', x: 100, y: 50, date: '2024-03-05T09:20:30.500Z' });
    expect(first.replies).toHaveLength(1);
    expect(first.replies[0]).toMatchObject({ author: 'Charles & Co', initials: 'CB', text: 'Done', date: '2024-03-05T11:00:00.000Z' });
    expect(first.replies[0].x).toBeUndefined();
    expect(whole.text).toBe('Whole slide');
    expect(whole.x).toBeUndefined();
    expect(whole.date).toBeUndefined();
    expect(unknown).toMatchObject({ author: 'Anonymous', x: 10, y: 2 });
  });

  it('reads Microsoft 365’s threaded comments: replies, resolved threads and comments on an object', async () => {
    const zip = await exported(deckWith(newShape('rect', 120, 80)));
    const slide = await zip.file('ppt/slides/slide1.xml').async('string');
    const shapeId = /<p:sp>[\s\S]*?<p:cNvPr id="(\d+)"/.exec(slide)[1];
    const M = 'xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
    const PC = 'xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command"';
    const AC = 'xmlns:ac="http://schemas.microsoft.com/office/drawing/2013/main/command"';
    const body = (...paras) => `<p188:txBody><a:bodyPr/><a:lstStyle/>${paras.map((p) => `<a:p><a:r><a:rPr lang="en-US"/><a:t>${p}</a:t></a:r></a:p>`).join('')}</p188:txBody>`;
    zip.file('ppt/authors.xml', `<?xml version="1.0"?><p188:authorLst ${M}>
      <p188:author id="{11111111-1111-1111-1111-111111111111}" name="Grace Hopper" initials="GH" userId="grace" providerId="AD"/>
      <p188:author id="{22222222-2222-2222-2222-222222222222}" name="Alan Turing" initials="AT" userId="alan" providerId="AD"/></p188:authorLst>`);
    zip.file('ppt/comments/modernComment_100_1.xml', `<?xml version="1.0"?><p188:cmLst ${M}>
      <p188:cm id="{A}" authorId="{11111111-1111-1111-1111-111111111111}" created="2025-01-02T03:04:05.678" status="resolved">
        <pc:sldMkLst ${PC}><pc:docMk/><pc:sldMk cId="1" sldId="256"/></pc:sldMkLst>
        <p188:pos x="1270000" y="635000"/>
        <p188:replyLst><p188:reply id="{B}" authorId="{22222222-2222-2222-2222-222222222222}" created="2025-01-02T04:00:00Z">${body('Agreed')}</p188:reply></p188:replyLst>
        ${body('Line one', 'Line two')}
      </p188:cm>
      <p188:cm id="{C}" authorId="{22222222-2222-2222-2222-222222222222}" created="2025-01-03T00:00:00.000">
        <ac:deMkLst ${AC} ${PC}><pc:docMk/><pc:sldMk cId="1" sldId="256"/><ac:spMk id="${shapeId}" creationId="{D}"/></ac:deMkLst>
        ${body('On the box')}
      </p188:cm></p188:cmLst>`);
    await addRel(zip, 'ppt/_rels/presentation.xml.rels', 'rIdAu', 'http://schemas.microsoft.com/office/2018/10/relationships/authors', 'authors.xml');
    await addRel(zip, 'ppt/slides/_rels/slide1.xml.rels', 'rIdMc', 'http://schemas.microsoft.com/office/2018/10/relationships/comments', '../comments/modernComment_100_1.xml');
    const [a, b] = (await reimport(zip)).slides[0].comments;
    expect(a).toMatchObject({ author: 'Grace Hopper', initials: 'GH', text: 'Line one\nLine two', date: '2025-01-02T03:04:05.678Z', x: 100, y: 50, resolved: true });
    expect(a.replies).toEqual([expect.objectContaining({ author: 'Alan Turing', text: 'Agreed', date: '2025-01-02T04:00:00.000Z' })]);
    expect(b).toMatchObject({ author: 'Alan Turing', text: 'On the box', x: 120, y: 80 });
    expect(b.resolved).toBeUndefined();
  });

  it('places comments on a resized slide like everything else', async () => {
    // A 4:3 slide twice the size of LibreSlides' (1440 × 1080 pt) is scaled by a half.
    const zip = await exported(deckWith());
    const pres = await zip.file('ppt/presentation.xml').async('string');
    zip.file('ppt/presentation.xml', pres.replace(/<p:sldSz[^>]*\/>/, '<p:sldSz cx="18288000" cy="13716000"/>'));
    zip.file('ppt/comments/comment1.xml', `<p:cmLst ${P}><p:cm authorId="0" idx="1"><p:pos x="800" y="800"/><p:text>x</p:text></p:cm></p:cmLst>`);
    await addRel(zip, 'ppt/slides/_rels/slide1.xml.rels', 'rIdC', `${REL}/comments`, '../comments/comment1.xml');
    const back = await reimport(zip);
    expect(back.size).toBe('4:3');
    expect(back.slides[0].comments[0]).toMatchObject({ x: 50, y: 50 });
  });

  it('keeps the slides when a comments part is damaged', async () => {
    const zip = await exported(deckWith());
    zip.file('ppt/comments/comment1.xml', '<p:cmLst');
    await addRel(zip, 'ppt/slides/_rels/slide1.xml.rels', 'rIdC', `${REL}/comments`, '../comments/comment1.xml');
    const back = await reimport(zip);
    expect(back.slides).toHaveLength(1);
    expect(back.slides[0].comments).toBeUndefined();
  });
});

describe('PowerPoint comments: writing', () => {
  const tricky = 'a < b && "c" \'d\' > e 😀\u0001 end\r\nnext';
  function reviewedDeck() {
    const deck = deckWith();
    deck.slides.push({ ...deck.slides[0], id: 's2', comments: undefined });
    deck.slides[0].comments = [
      { id: 'c1', author: 'Ada Lovelace', initials: 'AL', date: '2024-03-05T09:20:30.500Z', text: tricky, x: 100.4, y: 50,
        replies: [{ id: 'r1', author: 'Bob <B>', date: '2024-03-05T10:00:00.000Z', text: 'Reply 1' }, { id: 'r2', author: 'Ada Lovelace', date: '2024-03-05T11:00:00.000Z', text: 'Reply 2' }] },
      { id: 'c2', author: 'Bob <B>', date: '2024-03-06T00:00:00.000Z', text: 'Done', resolved: true },
    ];
    deck.slides[1].comments = [{ id: 'c3', author: 'Ada Lovelace', date: '2024-03-07T00:00:00.000Z', text: 'Slide two' }];
    return deck;
  }

  it('writes classic comments, their authors, relationships and content types', async () => {
    const zip = await exported(reviewedDeck());
    const types = await zip.file('[Content_Types].xml').async('string');
    expect(types).toContain('PartName="/ppt/commentAuthors.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"');
    expect(types).toContain('PartName="/ppt/comments/comment1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.comments+xml"');
    expect(types).toContain('PartName="/ppt/comments/comment2.xml"');
    expect(await zip.file('ppt/_rels/presentation.xml.rels').async('string')).toContain(`Type="${REL}/commentAuthors" Target="commentAuthors.xml"`);
    expect(await zip.file('ppt/slides/_rels/slide2.xml.rels').async('string')).toContain(`Type="${REL}/comments" Target="../comments/comment2.xml"`);

    const parse = async (p) => new DOMParser().parseFromString(await zip.file(p).async('string'), 'application/xml');
    const authors = await parse('ppt/commentAuthors.xml');
    expect(authors.querySelector('parsererror')).toBeNull();
    const list = [...authors.documentElement.children].map((a) => [a.getAttribute('id'), a.getAttribute('name'), a.getAttribute('initials'), a.getAttribute('lastIdx')]);
    expect(list).toEqual([['0', 'Ada Lovelace', 'AL', '3'], ['1', 'Bob <B>', 'BB', '2']]);

    const cms = await parse('ppt/comments/comment1.xml');
    expect(cms.querySelector('parsererror')).toBeNull();
    const cm = [...cms.documentElement.children];
    expect(cm).toHaveLength(4); // a thread of three, then the resolved one
    expect(cm.map((c) => `${c.getAttribute('authorId')}:${c.getAttribute('idx')}`)).toEqual(['0:1', '1:1', '0:2', '1:2']);
    // Replies sit with their thread and name its first comment.
    expect(cm[1].getElementsByTagNameNS('*', 'pos')[0].getAttribute('x')).toBe('803');
    expect(cm[1].getElementsByTagNameNS('*', 'parentCm')[0].getAttribute('idx')).toBe('1');
    expect(cm[0].getElementsByTagNameNS('*', 'parentCm')).toHaveLength(0);
    // Unanchored: at the slide's corner.
    expect(cm[3].getElementsByTagNameNS('*', 'pos')[0].getAttribute('x')).toBe('0');
    expect(cm[0].getElementsByTagNameNS('*', 'text')[0].textContent).toBe('a < b && "c" \'d\' > e 😀 end\nnext');
  });

  it('reads back what it wrote', async () => {
    const back = await reimport(await exported(reviewedDeck()));
    const [c1, c2] = back.slides[0].comments;
    expect(c1).toMatchObject({ author: 'Ada Lovelace', initials: 'AL', date: '2024-03-05T09:20:30.500Z', text: 'a < b && "c" \'d\' > e 😀 end\nnext', x: 100.38, y: 50 });
    expect(c1.replies.map((r) => [r.author, r.text, r.date])).toEqual([
      ['Bob <B>', 'Reply 1', '2024-03-05T10:00:00.000Z'],
      ['Ada Lovelace', 'Reply 2', '2024-03-05T11:00:00.000Z'],
    ]);
    // The classic format can't say a thread is resolved.
    expect(c2).toMatchObject({ author: 'Bob <B>', text: 'Done' });
    expect(c2.x).toBeUndefined();
    expect(back.slides[1].comments.map((c) => c.text)).toEqual(['Slide two']);
  });

  it('writes nothing for a deck without comments', async () => {
    const zip = await exported(deckWith());
    expect(zip.file('ppt/commentAuthors.xml')).toBeNull();
    expect(await zip.file('[Content_Types].xml').async('string')).not.toContain('comment');
  });

  it('escapes text for XML', () => {
    expect(xmlText('<a href="x">&\'')).toBe('&lt;a href=&quot;x&quot;&gt;&amp;&apos;');
    expect(xmlText('ok\u0000\u0008\u000b\ufffe\ud800 😀\t')).toBe('ok 😀\t');
  });
});

describe('comments in the model', () => {
  it('copies comments with a slide, under new ids', () => {
    const s = newSlide('blank');
    s.comments = [{ id: 'c1', author: 'A', date: '2024-01-01T00:00:00.000Z', text: 'hi', replies: [{ id: 'r1', author: 'B', text: 'yo' }] }];
    const copy = cloneSlide(s);
    expect(copy.comments[0]).toMatchObject({ author: 'A', text: 'hi' });
    expect(copy.comments[0].id).not.toBe('c1');
    expect(copy.comments[0].replies[0].id).not.toBe('r1');
    expect(copy.comments[0].replies[0].text).toBe('yo');
    copy.comments[0].text = 'changed';
    expect(s.comments[0].text).toBe('hi');
  });

  it('keeps only well-formed comments from an untrusted file', () => {
    const long = 'x'.repeat(20000);
    const deck = sanitizeDeck({ slides: [
      { elements: [], comments: [
        { id: 'c1', author: long, text: long, date: '2024-01-01T00:00:00.000Z', x: 10, y: 20, resolved: true, extra: '<script>', replies: [
          { id: 'r1', author: 'B', text: 'ok', date: 'not a date', x: 5 }, 'junk', null,
        ] },
        { author: 5, text: { html: '<b>' }, x: Infinity, y: 3, resolved: 'yes', date: 12, replies: 'no' },
        'junk', 7, null, [],
      ] },
      { elements: [], comments: 'nope' },
      { elements: [], comments: Array.from({ length: 2500 }, () => ({ text: 'a', replies: Array.from({ length: 600 }, () => ({ text: 'b' })) })) },
    ] });
    const [a, b] = deck.slides[0].comments;
    expect(deck.slides[0].comments).toHaveLength(2);
    expect(a.author).toHaveLength(100);
    expect(a.text).toHaveLength(10000);
    expect(a).toMatchObject({ id: 'c1', x: 10, y: 20, resolved: true, date: '2024-01-01T00:00:00.000Z' });
    expect(a.extra).toBeUndefined();
    expect(a.replies).toEqual([{ id: 'r1', author: 'B', text: 'ok' }]);
    expect(b).toEqual({ id: expect.any(String), author: '', text: '' });
    expect(deck.slides[1].comments).toBeUndefined();
    expect(deck.slides[2].comments).toHaveLength(2000);
    expect(deck.slides[2].comments[0].replies).toHaveLength(500);
  });
});

// ------------------------------------------------------------------ the editor

globalThis.ResizeObserver ||= class { observe() {} disconnect() {} unobserve() {} };
globalThis.matchMedia ||= () => ({ matches: false, addEventListener() {}, removeEventListener() {} });
globalThis.CSS ||= { escape: (s) => String(s).replace(/["\\]/g, '\\$&') };
Element.prototype.scrollIntoView ||= () => {};
document.queryCommandState ||= () => false;
document.execCommand ||= () => false;

beforeAll(() => {
  HTMLDialogElement.prototype.showModal ||= function showModal() { this.open = true; };
  HTMLDialogElement.prototype.close ||= function close() {
    this.open = false;
    this.dispatchEvent(new Event('close'));
  };
});

let editor;
function openEditor(deck) {
  const root = document.createElement('div');
  document.body.append(root);
  editor = new EditorScreen(root, { deckId: deck.id });
  editor.deck = deck;
  editor.slideSelection = new Set([deck.slides[0].id]);
  editor.build();
  editor.goToSlide(0);
  editor.filmstrip.render();
  return editor;
}

afterEach(async () => {
  await editor?.destroy({ save: false });
  editor = null;
  document.body.replaceChildren();
  localStorage.clear();
});

const key = (target, init) => target.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init }));
const settle = () => new Promise((r) => setTimeout(r, 0));
const box = (ed, k) => ed.panel.body.querySelector(`[data-composer="${k}"] textarea`);
async function post(ed, k, text) {
  const input = box(ed, k);
  input.value = text;
  input.dispatchEvent(new Event('input'));
  key(input, { key: 'Enter', ctrlKey: true });
  await settle();
}

describe('editor: comments', () => {
  it('adds, replies to, resolves and deletes comments, each one undo step', async () => {
    localStorage.setItem('ls:author', 'Ada Lovelace');
    const shape = newShape('rect', 120, 80);
    const d = newDeck();
    d.slides[0] = newSlide('blank', d);
    d.slides[0].elements.push(shape);
    const ed = openEditor(d);
    ed.setSelection([shape.id]);

    key(document.body, { key: 'm', code: 'KeyM', ctrlKey: true, altKey: true });
    expect(ed.panel.current).toBe('comments');
    expect(document.activeElement).toBe(box(ed, 'new'));
    expect(d.slides).toHaveLength(1); // not a new slide

    await post(ed, 'new', '<img src=x onerror=alert(1)> look');
    const [c] = d.slides[0].comments;
    expect(c).toMatchObject({ author: 'Ada Lovelace', initials: 'AL', text: '<img src=x onerror=alert(1)> look', x: 120, y: 80 });
    expect(Date.parse(c.date)).toBeGreaterThan(0);
    // Plain text on the page, never markup.
    expect(ed.panel.body.querySelector('.comment-text').textContent).toBe('<img src=x onerror=alert(1)> look');
    expect(ed.panel.body.querySelector('img')).toBeNull();
    expect(document.activeElement).toBe(box(ed, 'new'));
    expect(box(ed, 'new').value).toBe('');

    // A marker on the canvas, but not in the slide's picture.
    const marker = ed.canvas.overlay.querySelector('.comment-marker');
    expect(marker.style.left).toBe(`${120 * ed.canvas.scale}px`);
    expect(renderThumb(d.slides[0], d, 160).querySelector('.comment-marker')).toBeNull();

    await post(ed, c.id, 'A reply');
    expect(d.slides[0].comments[0].replies.map((r) => r.text)).toEqual(['A reply']);

    ed.panel.body.querySelector('.comment-actions .btn').click(); // Resolve
    expect(d.slides[0].comments[0].resolved).toBe(true);
    expect(ed.canvas.overlay.querySelector('.comment-marker')).toBeNull();

    ed.undo();
    expect(d.slides[0].comments[0].resolved).toBeUndefined();
    ed.undo();
    expect(d.slides[0].comments[0].replies).toBeUndefined();

    ed.panel.body.querySelector('.comment-actions [aria-label="Delete comment"]').click();
    expect(d.slides[0].comments).toBeUndefined();
    expect(ed.panel.body.textContent).toContain('No comments on this slide yet.');
    ed.undo();
    expect(d.slides[0].comments).toHaveLength(1);
  });

  it('asks for a name the first time, and remembers it', async () => {
    const ed = openEditor(newDeck());
    ed.insertComment();
    const input = box(ed, 'new');
    input.value = 'Hello';
    input.dispatchEvent(new Event('input'));
    key(input, { key: 'Enter', ctrlKey: true });
    await settle();
    const dlg = document.querySelector('dialog');
    expect(dlg.querySelector('input').value).toBe('You');
    dlg.querySelector('input').value = 'Grace';
    dlg.querySelector('form').dispatchEvent(new Event('submit', { cancelable: true }));
    dlg.close();
    await settle();
    await settle();
    expect(ed.deck.slides[0].comments[0]).toMatchObject({ author: 'Grace', text: 'Hello' });
    expect(ed.deck.slides[0].comments[0].x).toBeUndefined();
    expect(localStorage.getItem('ls:author')).toBe('Grace');
    expect(ed.panel.body.textContent).toContain('Commenting as Grace.');
  });

  it('opens the thread from its marker, and shows open comments in the slide list', async () => {
    const d = newDeck();
    d.slides.push(newSlide('blank', d));
    d.slides[1].comments = [
      { id: 'c1', author: 'A', date: '2024-01-01T00:00:00.000Z', text: 'one', x: 10, y: 10 },
      { id: 'c2', author: 'A', date: '2024-01-01T00:00:00.000Z', text: 'two', resolved: true },
    ];
    const ed = openEditor(d);
    const badge = (i) => ed.filmstrip.list.children[i].querySelector('.film-comments');
    expect(badge(0)).toBeNull();
    expect(badge(1).textContent).toBe('1');
    ed.goToSlide(1);
    ed.canvas.overlay.querySelector('.comment-marker').click();
    expect(ed.panel.current).toBe('comments');
    expect(ed.panel.body.querySelector('[data-thread="c1"]').classList.contains('is-active')).toBe(true);
    expect(ed.canvas.overlay.querySelector('.comment-marker').classList.contains('is-active')).toBe(true);
  });
});
