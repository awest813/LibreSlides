import { resolveColor } from '../model/themes.js';
import { mediaUrl } from './render.js';

/**
 * Drawing video and audio elements. While editing, and in thumbnails and
 * print, a video is its poster picture (a real <video> would take the clicks
 * that select and move it) and a sound is a speaker button. In a running
 * slideshow (ctx.live) a video is a <video> with its controls, and a sound
 * a button that plays and pauses it; the slideshow starts the ones set to
 * play by themselves (see present.js).
 */

const PLAY = '<svg viewBox="0 0 48 48" aria-hidden="true"><circle cx="24" cy="24" r="22" fill="rgba(0,0,0,0.55)"/><path d="M19 14.5v19l15-9.5z" fill="#fff"/></svg>';
const SPEAKER = '<svg viewBox="0 0 48 48" aria-hidden="true"><circle cx="24" cy="24" r="22" fill="rgba(0,0,0,0.55)"/>'
  + '<path d="M12 20h6l8-7v22l-8-7h-6z" fill="#fff"/><path d="M31 18.5a8 8 0 0 1 0 11M34.5 15a13 13 0 0 1 0 18" fill="none" stroke="#fff" stroke-width="2.6" stroke-linecap="round"/></svg>';

/** The poster picture, or a plain dark frame where there is none. */
function posterOf(el, ctx) {
  const url = el.poster ? mediaUrl(ctx.deck, el.poster) : '';
  if (!url) {
    const blank = document.createElement('div');
    blank.className = 'el-media-blank';
    return blank;
  }
  const img = document.createElement('img');
  img.className = 'el-img';
  img.src = url;
  img.alt = '';
  img.draggable = false;
  return img;
}

/** The round play (or speaker) mark, sized to the box. */
function badge(el, kind) {
  const mark = document.createElement('div');
  mark.className = 'el-media-badge';
  const side = Math.min(el.w, el.h);
  const size = kind === 'audio' ? side * 0.9 : Math.max(16, Math.min(72, side * 0.3));
  mark.style.width = `${size}px`;
  mark.style.height = `${size}px`;
  mark.innerHTML = kind === 'audio' ? SPEAKER : PLAY;
  return mark;
}

function liveVideo(el, ctx) {
  const video = document.createElement('video');
  video.className = 'el-media-player';
  video.src = mediaUrl(ctx.deck, el.media);
  if (el.poster) video.poster = mediaUrl(ctx.deck, el.poster);
  video.controls = true;
  video.playsInline = true;
  video.preload = 'metadata';
  video.loop = !!el.loop;
  return video;
}

/** A sound in the slideshow: its picture (or speaker) is a button that plays and pauses it. */
function liveAudio(el, ctx) {
  const audio = document.createElement('audio');
  audio.className = 'el-media-player';
  audio.src = mediaUrl(ctx.deck, el.media);
  audio.preload = 'metadata';
  audio.loop = !!el.loop;
  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'el-media-toggle';
  button.setAttribute('aria-label', 'Play sound');
  button.append(el.poster ? posterOf(el, ctx) : badge(el, 'audio'));
  button.addEventListener('click', () => (audio.paused ? audio.play().catch(() => {}) : audio.pause()));
  const label = () => {
    button.setAttribute('aria-label', audio.paused ? 'Play sound' : 'Pause sound');
    button.classList.toggle('is-playing', !audio.paused);
  };
  audio.addEventListener('play', label);
  audio.addEventListener('pause', label);
  return [button, audio];
}

/** The DOM for a media element's box. */
export function renderMedia(el, ctx) {
  const frame = document.createElement('div');
  frame.className = `el-media-frame el-media-${el.kind === 'audio' ? 'audio' : 'video'}`;
  const kind = el.kind === 'audio' ? 'audio' : 'video';
  if (ctx.live) {
    const parts = kind === 'audio' ? liveAudio(el, ctx) : [liveVideo(el, ctx)];
    frame.append(...parts);
    const player = parts[parts.length - 1];
    if (el.autoplay) player.dataset.autoplay = '1';
  } else if (kind === 'video') {
    // While editing the play mark tells a video from a picture.
    frame.append(posterOf(el, ctx));
    if (ctx.editing || !el.poster) frame.append(badge(el, kind));
  } else {
    // A sound with a picture (PowerPoint's speaker icon, usually) shows it.
    frame.append(el.poster ? posterOf(el, ctx) : badge(el, kind));
  }
  if (el.stroke && el.stroke.color !== 'none' && el.stroke.width) {
    frame.style.outline = `${el.stroke.width}px ${el.stroke.dash && el.stroke.dash !== 'solid' ? 'dashed' : 'solid'} ${resolveColor(el.stroke.color, ctx.theme)}`;
    frame.style.outlineOffset = `-${el.stroke.width / 2}px`;
  }
  if (el.shadow) frame.style.filter = 'drop-shadow(0 4px 8px rgba(0,0,0,0.3))';
  return frame;
}
