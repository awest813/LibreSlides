import { desc, attr } from './ooxml.js';

/**
 * Slide timelines (p:timing): LibreSlides' entrance animations, and videos
 * and sounds that play by themselves, written the way PowerPoint writes them;
 * and PowerPoint's entrance animations read back. Emphasis, exit and motion
 * path effects have no LibreSlides equivalent and are left out.
 */

// LibreSlides effects as PowerPoint entrance presets (presetID, presetSubtype).
const PRESETS = {
  appear: { id: 1, sub: 0 },
  fade: { id: 10, sub: 0 },
  'fly-up': { id: 2, sub: 4 }, // from the bottom
  'fly-down': { id: 2, sub: 1 }, // from the top
  'fly-left': { id: 2, sub: 8 },
  'fly-right': { id: 2, sub: 2 },
  zoom: { id: 53, sub: 16 },
  wipe: { id: 22, sub: 8 },
  spin: { id: 49, sub: 0 },
};

// Where a fly-in starts, in PowerPoint's slide-relative units.
const FLY_FROM = {
  1: ['ppt_y', '0-#ppt_h/2'], 2: ['ppt_x', '1+#ppt_w/2'], 4: ['ppt_y', '1+#ppt_h/2'], 8: ['ppt_x', '0-#ppt_w/2'],
};

/**
 * The p:timing for a slide. `anims` are { spid, effect, duration (s), start:
 * click|with|after } in order; `media` are { spid, kind, autoplay, loop,
 * duration? } for the slide's videos and sounds. Returns '' when there's nothing to play.
 */
export function timingXml({ anims = [], media = [] }) {
  let n = 1; // the root is 1; the rest count up in document order
  const id = () => ++n;
  const target = (spid) => `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>`;
  const behaviour = (spid, ms, extra = '') => `<p:cBhvr${extra}><p:cTn id="${id()}" dur="${ms}" fill="hold"/>${target(spid)}`;
  const tav = (attrName, from, to, spid, ms) => `<p:anim calcmode="lin" valueType="num">${behaviour(spid, ms, ' additive="base"')}`
    + `<p:attrNameLst><p:attrName>${attrName}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst>`
    + `<p:tav tm="0"><p:val><p:strVal val="${from}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${to}"/></p:val></p:tav></p:tavLst></p:anim>`;
  const fade = (spid, ms) => `<p:animEffect transition="in" filter="fade">${behaviour(spid, ms)}</p:cBhvr></p:animEffect>`;

  /** The behaviours that make one effect: shown (set visible), then how it comes in. */
  const effectBody = (a) => {
    const ms = Math.max(1, Math.round((a.duration ?? 0.5) * 1000));
    const p = PRESETS[a.effect];
    let body = `<p:set>${behaviour(a.spid, 1).replace('fill="hold"/>', 'fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>')}`
      + '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>';
    if (a.effect === 'fade') body += fade(a.spid, ms);
    else if (p.id === 2) {
      const [moving, from] = FLY_FROM[p.sub];
      const still = moving === 'ppt_x' ? 'ppt_y' : 'ppt_x';
      body += tav(moving, from, `#${moving}`, a.spid, ms) + tav(still, `#${still}`, `#${still}`, a.spid, ms);
    } else if (a.effect === 'zoom') body += tav('ppt_w', '0', '#ppt_w', a.spid, ms) + tav('ppt_h', '0', '#ppt_h', a.spid, ms) + fade(a.spid, ms);
    else if (a.effect === 'wipe') body += `<p:animEffect transition="in" filter="wipe(left)">${behaviour(a.spid, ms)}</p:cBhvr></p:animEffect>`;
    else if (a.effect === 'spin') body += tav('style.rotation', '-90', '0', a.spid, ms) + tav('ppt_w', '#ppt_w*0.7', '#ppt_w', a.spid, ms) + tav('ppt_h', '#ppt_h*0.7', '#ppt_h', a.spid, ms) + fade(a.spid, ms);
    return body;
  };
  const effectNode = (a, nodeType, delayMs) => {
    if (a.play) {
      return `<p:par><p:cTn id="${id()}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="${nodeType}">`
        + `<p:stCondLst><p:cond delay="${delayMs}"/></p:stCondLst><p:childTnLst><p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr>`
        + `<p:cTn id="${id()}" dur="${a.duration ? Math.round(a.duration * 1000) : 'indefinite'}" fill="hold"/>${target(a.spid)}`
        + '</p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par>';
    }
    const p = PRESETS[a.effect];
    return `<p:par><p:cTn id="${id()}" presetID="${p.id}" presetClass="entr" presetSubtype="${p.sub}" fill="hold" grpId="0" nodeType="${nodeType}">`
      + `<p:stCondLst><p:cond delay="${delayMs}"/></p:stCondLst><p:childTnLst>${effectBody(a)}</p:childTnLst></p:cTn></p:par>`;
  };

  // Media that plays by itself starts the slide's first step; animations follow in their order.
  const plays = media.filter((m) => m.autoplay).map((m, i) => ({ ...m, play: true, start: i ? 'with' : 'after', duration: m.duration }));
  const items = [...plays, ...anims.filter((a) => PRESETS[a.effect])];
  let seq = '';
  if (items.length) {
    // Steps (a click each, or one that starts by itself), each made of groups that start
    // together, after the one before ends.
    const steps = [];
    for (const it of items) {
      const step = steps[steps.length - 1];
      if (!step || it.start === 'click') steps.push({ auto: !step && it.start !== 'click', groups: [[it]] });
      else if (it.start === 'after') step.groups.push([it]);
      else step.groups[step.groups.length - 1].push(it);
    }
    const main = id();
    const body = steps.map((step) => {
      const stepId = id();
      let at = 0; // ms since the step began
      const groups = step.groups.map((g, gi) => {
        const groupId = id();
        const nodes = g.map((it, k) => effectNode(it, k ? 'withEffect' : gi ? 'afterEffect' : it.start === 'click' ? 'clickEffect' : it.start === 'with' ? 'withEffect' : 'afterEffect', 0)).join('');
        const xml = `<p:par><p:cTn id="${groupId}" fill="hold"><p:stCondLst><p:cond delay="${at}"/></p:stCondLst><p:childTnLst>${nodes}</p:childTnLst></p:cTn></p:par>`;
        at += Math.max(...g.map((it) => (it.play ? 0 : Math.round((it.duration ?? 0.5) * 1000))));
        return xml;
      }).join('');
      // A step that plays by itself begins with the slide (the main sequence).
      const begin = step.auto ? `<p:cond evt="onBegin" delay="0"><p:tn val="${main}"/></p:cond>` : '';
      return `<p:par><p:cTn id="${stepId}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>${begin}</p:stCondLst><p:childTnLst>${groups}</p:childTnLst></p:cTn></p:par>`;
    }).join('');
    seq = `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${main}" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${body}</p:childTnLst></p:cTn>`
      + '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
      + '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>';
  }
  // Each video or sound's media node, which is where looping is set.
  const nodes = media.filter((m) => m.autoplay || m.loop).map((m) => `<p:${m.kind}><p:cMediaNode vol="80000"><p:cTn id="${id()}" fill="hold"${m.loop ? ' repeatCount="indefinite"' : ''}>`
    + `<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst></p:cTn>${target(m.spid)}</p:cMediaNode></p:${m.kind}>`).join('');
  if (!seq && !nodes) return '';
  // Shapes with text are built as a whole.
  const built = [...new Set(anims.filter((a) => PRESETS[a.effect] && a.hasText).map((a) => a.spid))];
  const bldLst = built.length ? `<p:bldLst>${built.map((s) => `<p:bldP spid="${s}" grpId="0" animBg="1"/>`).join('')}</p:bldLst>` : '';
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>${seq}${nodes}</p:childTnLst></p:cTn></p:par></p:tnLst>${bldLst}</p:timing>`;
}

/** Put a slide's timeline in its XML, where PowerPoint expects it (after the transition, before extensions). */
export function insertTiming(xml, timing) {
  if (!timing) return xml;
  xml = xml.replace(/<p:timing>[\s\S]*?<\/p:timing>/, '');
  const ext = xml.lastIndexOf('<p:extLst>');
  const end = xml.lastIndexOf('</p:sld>');
  // A slide-level extension list sits right before </p:sld>; anything else is inside the shapes.
  const at = ext > xml.lastIndexOf('</p:cSld>') ? ext : end;
  return xml.slice(0, at) + timing + xml.slice(at);
}

/**
 * PowerPoint's entrance animations on a slide, as { spid, effect, duration,
 * start, order } in the order they play. Effects LibreSlides has no
 * equivalent of come in as fades; emphasis, exit and motion paths are left out.
 */
export function readAnimations(doc) {
  const timing = desc(doc?.documentElement, 'timing')[0];
  const main = desc(timing, 'cTn').find((c) => attr(c, 'nodeType') === 'mainSeq');
  if (!main) return [];
  const out = [];
  const seen = new Set();
  for (const effect of desc(main, 'cTn')) {
    if (attr(effect, 'presetClass') !== 'entr') continue;
    const spid = attr(desc(effect, 'spTgt')[0], 'spid');
    // A shape built paragraph by paragraph has an effect per paragraph: the first stands for it.
    if (!spid || seen.has(spid)) continue;
    seen.add(spid);
    const id = +attr(effect, 'presetID');
    const sub = +attr(effect, 'presetSubtype') || 0;
    const name = id === 1 ? 'appear' : id === 2 || id === 7 ? ({ 1: 'fly-down', 2: 'fly-right', 8: 'fly-left' }[sub] || 'fly-up')
      : id === 22 ? 'wipe' : [6, 8, 16, 17, 21, 23, 53, 55].includes(id) ? 'zoom' : [19, 45, 49].includes(id) ? 'spin' : 'fade';
    const durations = desc(effect, 'cTn').map((c) => +attr(c, 'dur')).filter((d) => d > 1);
    const type = attr(effect, 'nodeType');
    out.push({
      spid, effect: name,
      duration: Math.min(5, Math.max(0.1, (durations.length ? Math.max(...durations) : 500) / 1000)),
      start: type === 'withEffect' ? 'with' : type === 'afterEffect' ? 'after' : 'click',
      order: out.length + 1,
    });
  }
  return out;
}

