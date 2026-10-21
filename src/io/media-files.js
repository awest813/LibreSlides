/**
 * Video and sound files: the formats LibreSlides keeps in a deck (the ones
 * browsers play), and reading one from the device, with a still frame of a
 * video to show until it plays.
 */

/** Data URL types kept in decks, and the file extension each is saved as in PowerPoint files. */
export const MEDIA_TYPES = {
  'video/mp4': 'mp4',
  'video/webm': 'webm',
  'video/ogg': 'ogv',
  'video/quicktime': 'mov',
  'audio/mpeg': 'mp3',
  'audio/mp3': 'mp3',
  'audio/mp4': 'm4a',
  'audio/x-m4a': 'm4a',
  'audio/wav': 'wav',
  'audio/x-wav': 'wav',
  'audio/ogg': 'ogg',
};

/** Content types PowerPoint uses for those extensions. */
export const MEDIA_CONTENT_TYPES = { mp4: 'video/mp4', webm: 'video/webm', ogv: 'video/ogg', mov: 'video/quicktime', mp3: 'audio/mpeg', m4a: 'audio/mp4', wav: 'audio/wav', ogg: 'audio/ogg' };

/** Bigger files still work, but they live in the browser's storage and slow saving down. */
export const LARGE_MEDIA_BYTES = 50 * 1024 * 1024;

/** The type of a "data:" URL, lower-cased ("video/mp4"), or ''. */
export const dataType = (dataUrl) => (/^data:([^;,]+)/i.exec(dataUrl || '')?.[1] || '').toLowerCase();

/** A base64 data URL of a video or sound format LibreSlides keeps. */
export const isMediaData = (v) => typeof v === 'string' && /^data:[^;,]+;base64,/i.test(v) && Object.hasOwn(MEDIA_TYPES, dataType(v));

const BY_EXTENSION = { mp4: 'video/mp4', m4v: 'video/mp4', webm: 'video/webm', ogv: 'video/ogg', mov: 'video/quicktime', mp3: 'audio/mpeg', m4a: 'audio/mp4', wav: 'audio/wav', ogg: 'audio/ogg', oga: 'audio/ogg' };

/** The type a file is kept as, from its own type or else its name; null for other files. */
export function mediaTypeOf(file) {
  const type = (file.type || '').toLowerCase();
  if (Object.hasOwn(MEDIA_TYPES, type)) return type;
  return BY_EXTENSION[(/\.([^.]+)$/.exec(file.name || '')?.[1] || '').toLowerCase()] || null;
}

const readAsDataUrl = (blob) => new Promise((resolve, reject) => {
  const r = new FileReader();
  r.onload = () => resolve(r.result);
  r.onerror = () => reject(r.error || new Error('Couldn’t read the file.'));
  r.readAsDataURL(blob);
});

/**
 * A video or sound file as { kind, dataUrl, poster, width, height, duration }.
 * `poster` (a PNG data URL) and the sizes are null where the browser can't
 * decode the file.
 */
export async function readMediaFile(file) {
  const type = mediaTypeOf(file);
  if (!type) throw new Error('That file isn’t a video or sound LibreSlides can play (try MP4, WebM, MP3, M4A or WAV).');
  // Some systems give no type, or another spelling of it: label the data with the one we know.
  const dataUrl = (await readAsDataUrl(file)).replace(/^data:[^;,]*/, `data:${type}`);
  const kind = type.startsWith('video/') ? 'video' : 'audio';
  const probe = await probeMedia(file, kind);
  return { kind, dataUrl, poster: probe?.poster || null, width: probe?.width || null, height: probe?.height || null, duration: probe?.duration || null };
}

/** The length of a file and, for a video, its size and a frame from half a second in. */
function probeMedia(file, kind) {
  return new Promise((resolve) => {
    const media = document.createElement(kind);
    const url = URL.createObjectURL(file);
    const info = {};
    let timer = null;
    const done = () => {
      clearTimeout(timer);
      media.removeAttribute('src');
      media.load?.();
      URL.revokeObjectURL(url);
      resolve(info.duration || info.width ? info : null);
    };
    media.muted = true;
    media.preload = 'auto';
    media.playsInline = true;
    media.addEventListener('loadedmetadata', () => {
      if (Number.isFinite(media.duration)) info.duration = Math.round(media.duration * 1000) / 1000;
      if (kind === 'audio') return done();
      info.width = media.videoWidth || null;
      info.height = media.videoHeight || null;
      // The very first frame is often black; a little way in usually shows what the video is.
      media.currentTime = Math.min(0.5, (info.duration || 1) / 2);
    }, { once: true });
    media.addEventListener('seeked', () => {
      info.poster = drawFrame(media);
      done();
    }, { once: true });
    media.addEventListener('error', done, { once: true });
    // A file the browser can't decode may never say so.
    timer = setTimeout(done, 8000);
    media.src = url;
  });
}

/** The video's current frame as a PNG, at most 960 pixels across. */
function drawFrame(video) {
  try {
    const w = video.videoWidth;
    const h = video.videoHeight;
    if (!w || !h) return null;
    const k = Math.min(1, 960 / Math.max(w, h));
    const canvas = document.createElement('canvas');
    canvas.width = Math.round(w * k);
    canvas.height = Math.round(h * k);
    canvas.getContext('2d').drawImage(video, 0, 0, canvas.width, canvas.height);
    return canvas.toDataURL('image/png');
  } catch {
    return null;
  }
}
