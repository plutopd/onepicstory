// Minimal local server for vibe-cut.
//
// Serves vibe-cut.html as a static file and exposes a small proxy endpoint
// (/api/ai-direct) that forwards the "AI director" request to the Google
// Gemini API (free tier, vision-capable). The Gemini API key lives ONLY
// here, in the server process's environment — it is never sent to or
// stored in the browser.

require('dotenv').config();

const express = require('express');
const { spawn } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const app = express();
app.use(express.json({ limit: '250mb' })); // full-resolution photos for rendering can add up

// Serve vibe-cut.html and any other static assets in this folder.
app.use(express.static(__dirname));
// The deployed link should open the product itself. landing.html is still
// served at its own path, but it predates the current light theme, so pointing
// the root at it would show visitors a different-looking product first.
app.get('/', (req, res) => res.redirect('/vibe-cut.html'));

const API_KEY = process.env.GEMINI_API_KEY;
const MODEL = process.env.GEMINI_MODEL || 'gemini-3.6-flash';

// Frontend sends a provider-agnostic list of parts:
//   { type: 'text', text: '...' }
//   { type: 'image', mediaType: 'image/jpeg', data: '<base64>' }
// This converts them into Gemini's `contents[].parts[]` shape.
function toGeminiParts(content) {
  return content.map(part => {
    if (part.type === 'text') return { text: part.text };
    if (part.type === 'image') {
      return { inline_data: { mime_type: part.mediaType, data: part.data } };
    }
    throw new Error(`Unsupported content part type: ${part.type}`);
  });
}

app.post('/api/ai-direct', async (req, res) => {
  if (!API_KEY) {
    return res.status(500).json({
      error: 'Server is missing GEMINI_API_KEY. Copy .env.example to .env, add your key from https://aistudio.google.com/apikey, and restart the server.'
    });
  }

  const { content, maxOutputTokens } = req.body || {};
  if (!Array.isArray(content) || content.length === 0) {
    return res.status(400).json({ error: 'Request body must include a non-empty "content" array.' });
  }

  let parts;
  try {
    parts = toGeminiParts(content);
  } catch (err) {
    return res.status(400).json({ error: err.message });
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent?key=${API_KEY}`;
  const body = JSON.stringify({
    contents: [{ role: 'user', parts }],
    generationConfig: {
      maxOutputTokens: Number.isFinite(maxOutputTokens) ? Math.min(maxOutputTokens, 4000) : 1000,
      responseMimeType: 'application/json',
      // gemini-3.x models spend hidden "thinking" tokens out of the SAME
      // maxOutputTokens budget before writing the actual answer — on a long,
      // complex prompt (many photos, many rules) it can burn the whole
      // budget thinking and return an empty response. Our tasks are all
      // classification/structured-output, which Google's own docs say
      // needs minimal reasoning, so keep thinking low to leave room for
      // the actual JSON.
      thinkingConfig: { thinkingLevel: 'low' }
    }
  });

  // The free tier occasionally returns 503 "model overloaded" under heavy
  // public demand — that's transient, so retry a few times with backoff
  // before giving up.
  const MAX_ATTEMPTS = 3;
  let lastError;

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      const upstream = await fetch(url, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body
      });

      const data = await upstream.json();

      if (!upstream.ok) {
        const message = (data && data.error && data.error.message) || `Gemini API error (${upstream.status})`;
        if (upstream.status === 503 && attempt < MAX_ATTEMPTS) {
          lastError = { status: upstream.status, message };
          await new Promise(r => setTimeout(r, attempt * 1000));
          continue;
        }
        return res.status(upstream.status).json({ error: message });
      }

      const text = (data.candidates || [])
        .flatMap(c => (c.content && c.content.parts) || [])
        .map(p => p.text || '')
        .join('\n');

      if (!text) {
        return res.status(502).json({ error: 'Gemini returned no text output.' });
      }

      return res.json({ text });
    } catch (err) {
      lastError = { status: 502, message: 'Failed to reach the Gemini API: ' + err.message };
      if (attempt < MAX_ATTEMPTS) {
        await new Promise(r => setTimeout(r, attempt * 1000));
      }
    }
  }

  res.status(lastError.status).json({ error: lastError.message + ' (재시도 3회 후에도 실패했어요. 잠시 후 다시 시도해주세요.)' });
});

// --- Video rendering via ffmpeg ---
//
// Runs entirely on this machine (no cloud upload). The AI is free to choose
// any of these ~50 xfade transition names — this list is FFmpeg's actual
// built-in vocabulary, not something we hand-implemented, so the AI's
// creative choice and what can actually be rendered are the same set.
const XFADE_TRANSITIONS = new Set([
  'fade', 'fadeblack', 'fadewhite', 'distance', 'wipeleft', 'wiperight', 'wipeup', 'wipedown',
  'slideleft', 'slideright', 'slideup', 'slidedown', 'smoothleft', 'smoothright', 'smoothup', 'smoothdown',
  'circlecrop', 'rectcrop', 'circleclose', 'circleopen', 'horzclose', 'horzopen', 'vertclose', 'vertopen',
  'diagbl', 'diagbr', 'diagtl', 'diagtr', 'hlslice', 'hrslice', 'vuslice', 'vdslice', 'hblur', 'fadegrays',
  'wipetl', 'wipetr', 'wipebl', 'wipebr', 'squeezeh', 'squeezev', 'zoomin', 'dissolve', 'pixelize', 'radial',
  'hlwind', 'hrwind', 'vuwind', 'vdwind', 'coverleft', 'coverright', 'coverup', 'coverdown',
  'revealleft', 'revealright', 'revealup', 'revealdown'
]);

// Caption/subtitle burn-in config. textfile= (rather than text=) is used so
// caption content never has to be escaped for the filtergraph parser — only
// the (server-generated, ASCII, punctuation-free) temp file path does.
// drawtext needs a real font file, and the one that ships with the OS differs
// per platform — a macOS-only path meant every captioned render failed outright
// once this ran on a Linux host. Resolved once, from the first candidate that
// actually exists, so the same build works locally and deployed.
const CAPTION_FONT_CANDIDATES = [
  process.env.CAPTION_FONT_FILE,
  '/System/Library/Fonts/AppleSDGothicNeo.ttc',                 // macOS
  '/usr/share/fonts/truetype/nanum/NanumGothic.ttf',            // Debian/Ubuntu: fonts-nanum
  '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',     // Debian: fonts-noto-cjk
  '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
  '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'             // last resort: no Hangul glyphs
].filter(Boolean);
const CAPTION_FONT_FILE = CAPTION_FONT_CANDIDATES.find(f => { try { return fs.existsSync(f); } catch(e){ return false; } }) || null;
const CAPTION_Y_FRAC = { top: 0.12, center: 0.5, bottom: 0.85 };
function escapeDrawtextArg(str) {
  return String(str).replace(/\\/g, '\\\\').replace(/:/g, '\\:').replace(/'/g, "'\\\\''");
}

// Accepts "#rrggbb" or "rgba(r,g,b,a)" (the two forms the client actually
// sends) and converts to ffmpeg's own color syntax ("0xRRGGBB" or
// "0xRRGGBB@alpha"). Falls back to a solid black on anything unrecognized
// so a bad value degrades to a visible-but-plain look rather than
// breaking the whole filtergraph.
function toFfmpegColor(cssColor, fallback) {
  if (typeof cssColor === 'string') {
    const rgbaMatch = cssColor.match(/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*(?:,\s*([\d.]+)\s*)?\)$/i);
    if (rgbaMatch) {
      const hex = [rgbaMatch[1], rgbaMatch[2], rgbaMatch[3]]
        .map(n => Math.max(0, Math.min(255, parseInt(n, 10))).toString(16).padStart(2, '0'))
        .join('');
      const alpha = rgbaMatch[4] !== undefined ? Math.max(0, Math.min(1, parseFloat(rgbaMatch[4]))) : 1;
      return `0x${hex}@${alpha.toFixed(2)}`;
    }
    if (/^#[0-9a-fA-F]{6}$/.test(cssColor)) return '0x' + cssColor.slice(1);
  }
  return fallback;
}

const COLOR_GRADE_FILTERS = {
  none: null,
  warm: 'eq=saturation=1.3:gamma_r=1.05:gamma_b=0.92',
  cool: 'eq=saturation=1.1:gamma_b=1.08:gamma_r=0.95',
  vintage: 'curves=preset=vintage,eq=contrast=0.92:saturation=0.85',
  bw: 'hue=s=0',
  vivid: 'eq=saturation=1.6:contrast=1.15'
};

// Resolve the ffmpeg binary once. Prefers PATH, but falls back to common
// Homebrew install locations in case the process launching this server
// (e.g. a non-login shell) didn't inherit the user's usual PATH.
let ffmpegBinaryPromise = null;
function resolveFfmpegBinary() {
  if (ffmpegBinaryPromise) return ffmpegBinaryPromise;
  // ffmpeg-full (if installed) is preferred over the default Homebrew ffmpeg
  // formula: the default build ships without libfreetype/libass, so it has
  // no "drawtext" filter and can't burn captions into the exported video.
  const candidates = [
    process.env.FFMPEG_PATH,
    '/opt/homebrew/opt/ffmpeg-full/bin/ffmpeg',
    '/usr/local/opt/ffmpeg-full/bin/ffmpeg',
    'ffmpeg',
    '/opt/homebrew/bin/ffmpeg',
    '/usr/local/bin/ffmpeg'
  ].filter(Boolean);
  ffmpegBinaryPromise = (async () => {
    for (const candidate of candidates) {
      const ok = await new Promise(resolve => {
        const proc = spawn(candidate, ['-version']);
        proc.on('error', () => resolve(false));
        proc.on('close', code => resolve(code === 0));
      });
      if (ok) return candidate;
    }
    return null;
  })();
  return ffmpegBinaryPromise;
}

// ffprobe lives alongside whichever ffmpeg binary got resolved (same bin
// directory for both the default Homebrew formula and ffmpeg-full) — reuse
// that instead of re-running the whole PATH/candidate search independently.
let ffprobeBinaryPromise = null;
function resolveFfprobeBinary(ffmpegBinary) {
  if (ffprobeBinaryPromise) return ffprobeBinaryPromise;
  const candidates = [
    process.env.FFPROBE_PATH,
    path.join(path.dirname(ffmpegBinary), 'ffprobe'),
    'ffprobe'
  ].filter(Boolean);
  ffprobeBinaryPromise = (async () => {
    for (const candidate of candidates) {
      const ok = await new Promise(resolve => {
        const proc = spawn(candidate, ['-version']);
        proc.on('error', () => resolve(false));
        proc.on('close', code => resolve(code === 0));
      });
      if (ok) return candidate;
    }
    return null;
  })();
  return ffprobeBinaryPromise;
}

// Whether a video file has its own audio stream — a silent source clip (or
// one ffprobe can't read) needs a synthesized-silence audio segment instead
// of an "[i:a]" map that would make the whole render fail.
function probeHasAudioStream(ffprobeBinary, filePath) {
  return new Promise(resolve => {
    if (!ffprobeBinary) return resolve(false);
    const proc = spawn(ffprobeBinary, ['-v', 'error', '-select_streams', 'a', '-show_entries', 'stream=index', '-of', 'csv=p=0', filePath]);
    let out = '';
    proc.stdout.on('data', d => { out += d.toString(); });
    proc.on('error', () => resolve(false));
    proc.on('close', () => resolve(out.trim().length > 0));
  });
}

function clampNum(v, def, min, max) {
  const n = typeof v === 'number' && Number.isFinite(v) ? v : def;
  return Math.max(min, Math.min(max, n));
}

function runFfmpeg(binary, args) {
  return new Promise((resolve, reject) => {
    const proc = spawn(binary, args);
    let stderr = '';
    proc.stderr.on('data', d => { stderr += d.toString(); });
    proc.on('error', reject);
    proc.on('close', code => {
      if (code === 0) resolve();
      else reject(new Error(`ffmpeg exited with code ${code}: ${stderr.slice(-1500)}`));
    });
  });
}

app.post('/api/render', async (req, res) => {
  const ffmpegBinary = await resolveFfmpegBinary();
  if (!ffmpegBinary) {
    return res.status(500).json({ error: 'ffmpeg is not installed or not found. Install it with "brew install ffmpeg" and restart the server.' });
  }

  const { photos, extraClipTracks, width, height, fps, transition, transitions, transitionDuration, fadeInStart, fadeOutEnd, fadeDuration, colorGrade, audioTracks, captions } = req.body || {};

  if (!Array.isArray(photos) || photos.length === 0) {
    return res.status(400).json({ error: 'Request body must include a non-empty "photos" array.' });
  }
  if (photos.length > 300) {
    return res.status(400).json({ error: 'Too many photos for a single render (max 300).' });
  }
  // Each extra CLIP track (CLIP2, CLIP3, ...) is its own independent
  // photo-only sequence — validate each one's entries the same way.
  const validExtraClipTracks = (Array.isArray(extraClipTracks) ? extraClipTracks : []).map(trackPhotos =>
    (Array.isArray(trackPhotos) ? trackPhotos : []).filter(p =>
      p && typeof p.data === 'string' && Number.isFinite(p.duration) && p.duration > 0
    )
  );
  const validAudioTracks = (Array.isArray(audioTracks) ? audioTracks : []).filter(a =>
    a && typeof a.data === 'string'
  );
  for (const p of photos) {
    if (p.kind === 'video') {
      if (typeof p.videoData !== 'string' || !Number.isFinite(p.duration) || p.duration <= 0) {
        return res.status(400).json({ error: 'Each video clip needs "videoData" (base64) and a positive "duration".' });
      }
      if (!Number.isFinite(p.inPoint) || p.inPoint < 0) {
        return res.status(400).json({ error: 'Each video clip needs a non-negative "inPoint".' });
      }
    } else if (typeof p.data !== 'string' || !Number.isFinite(p.duration) || p.duration <= 0) {
      return res.status(400).json({ error: 'Each photo needs "data" (base64 JPEG) and a positive "duration".' });
    }
    if (p.gapAfter !== undefined && (!Number.isFinite(p.gapAfter) || p.gapAfter < 0)) {
      return res.status(400).json({ error: '"gapAfter", if given, must be a non-negative number.' });
    }
  }

  const W = Number.isFinite(width) ? Math.round(width) : 1280;
  const H = Number.isFinite(height) ? Math.round(height) : 720;
  const FPS = Number.isFinite(fps) ? Math.round(fps) : 30;
  const td = clampNum(transitionDuration, 0.6, 0.05, 3);
  const fd = clampNum(fadeDuration, 0.6, 0.05, 3);
  const colorFilter = COLOR_GRADE_FILTERS.hasOwnProperty(colorGrade) ? COLOR_GRADE_FILTERS[colorGrade] : null;

  // Normalize captions up front; entries that fail validation are silently
  // dropped rather than failing the whole render (they're a decorative
  // overlay, not load-bearing content).
  const validCaptions = (Array.isArray(captions) ? captions : []).filter(c => {
    return c && typeof c.text === 'string' && c.text.trim() &&
      Number.isFinite(c.start) && Number.isFinite(c.end) && c.end > c.start;
  }).map(c => ({
    text: c.text,
    start: Math.max(0, c.start),
    end: c.end,
    color: typeof c.color === 'string' && /^#[0-9a-fA-F]{6}$/.test(c.color) ? c.color : '#ffffff',
    size: Math.round(clampNum(c.size, 36, 12, 120)),
    box: c.box !== false,
    // Free-form normalized position (0-1) — x/y are the newer, primary
    // field; "position" (a legacy top/center/bottom preset) is still
    // accepted as a fallback for any older client build.
    x: clampNum(c.x, 0.5, 0, 1),
    y: clampNum(c.y, CAPTION_Y_FRAC.hasOwnProperty(c.position) ? CAPTION_Y_FRAC[c.position] : CAPTION_Y_FRAC.bottom, 0, 1),
    borderWidth: Math.round(clampNum(c.borderWidth, 0, 0, 20)),
    borderColor: toFfmpegColor(c.borderColor, '0x000000'),
    shadowOffsetX: Math.round(clampNum(c.shadowOffsetX, 0, -30, 30)),
    shadowOffsetY: Math.round(clampNum(c.shadowOffsetY, 0, -30, 30)),
    shadowColor: toFfmpegColor(c.shadowColor, '0x000000@0.7')
  }));

  // The transition used between each specific adjacent pair of photos.
  // "transitions[i]" (if given and valid) overrides the single global
  // "transition" for the junction between photo i and photo i+1 — this is
  // what lets a user ask for "cut" everywhere except a "dissolve" between
  // photo 3 and 4, for example.
  const isValidJunctionName = v => v === 'none' || v === 'cut' || XFADE_TRANSITIONS.has(v);
  const globalJunctionName = isValidJunctionName(transition) ? transition : 'none';
  function junctionTransitionFor(i) {
    const override = Array.isArray(transitions) ? transitions[i] : null;
    return isValidJunctionName(override) ? override : globalJunctionName;
  }

  // Expand photos into a flat render sequence: real photo clips interspersed
  // with synthetic black "gap" clips. A Premiere-style trim that shortens a
  // photo turns the freed time into gapAfter instead of shifting the next
  // photo earlier — this is where that gap actually becomes a rendered clip.
  const segments = [];
  photos.forEach((p, i) => {
    segments.push({ kind: p.kind === 'video' ? 'video' : 'photo', photo: p, photoIndex: i });
    const gapDur = Number.isFinite(p.gapAfter) ? p.gapAfter : 0;
    if (gapDur > 0.01) segments.push({ kind: 'gap', duration: gapDur });
  });
  const segmentDuration = seg => seg.kind === 'gap' ? seg.duration : seg.photo.duration;

  const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'vibecut-'));

  try {
    const inputArgs = [];
    segments.forEach((seg, i) => {
      if (seg.kind === 'gap') {
        inputArgs.push('-f', 'lavfi', '-i', `color=c=black:s=${W}x${H}:d=${seg.duration.toFixed(3)}:r=${FPS}`);
        return;
      }
      if (seg.kind === 'video') {
        const mime = seg.photo.mimeType || '';
        const ext = mime.includes('webm') ? 'webm' : (mime.includes('quicktime') || mime.includes('mov')) ? 'mov' : mime.includes('avi') ? 'avi' : 'mp4';
        const filePath = path.join(workDir, `clip_${i}.${ext}`);
        fs.writeFileSync(filePath, Buffer.from(seg.photo.videoData, 'base64'));
        seg.filePath = filePath; // needed again below to probe for an audio stream
        const inPoint = Math.max(0, seg.photo.inPoint || 0);
        // -ss before -i seeks at the demuxer level (fast, and accurate enough
        // for this); -t as an input option caps how much of the file is read,
        // trimming video AND audio together so they can't drift apart.
        inputArgs.push('-ss', inPoint.toFixed(3), '-t', seg.photo.duration.toFixed(3), '-i', filePath);
        return;
      }
      const filePath = path.join(workDir, `frame_${i}.jpg`);
      fs.writeFileSync(filePath, Buffer.from(seg.photo.data, 'base64'));
      // -r on the INPUT is required: without it, the image2 demuxer loops a
      // still image at a default 25fps regardless of our target fps, which
      // silently desyncs every duration/offset computed below.
      inputArgs.push('-loop', '1', '-r', String(FPS), '-t', seg.photo.duration.toFixed(3), '-i', filePath);
    });

    // Extra CLIP tracks (CLIP2, CLIP3, ...): each an independent sequence of
    // photo-only segments that overlays everything drawn so far wherever it
    // has content. Only non-gap segments become actual ffmpeg inputs — a gap
    // just means nothing is overlaid for that stretch, letting whatever's
    // underneath show through untouched. Track order in the array is z-order:
    // a later track overlays an earlier one, same as higher-numbered tracks
    // in Premiere.
    const extraTrackSegmentsList = validExtraClipTracks.map(trackPhotos => {
      const segs = [];
      let t2 = 0;
      trackPhotos.forEach(p2 => {
        segs.push({ start: t2, end: t2 + p2.duration, photo: p2 });
        t2 += p2.duration + (Number.isFinite(p2.gapAfter) ? Math.max(0, p2.gapAfter) : 0);
      });
      return segs;
    });
    let nextInputIndex = segments.length;
    const extraTrackInputBases = extraTrackSegmentsList.map(segs => {
      const base = nextInputIndex;
      segs.forEach((seg2, i2) => {
        const filePath = path.join(workDir, `overlay_${base}_${i2}.jpg`);
        fs.writeFileSync(filePath, Buffer.from(seg2.photo.data, 'base64'));
        inputArgs.push('-loop', '1', '-r', String(FPS), '-t', seg2.photo.duration.toFixed(3), '-i', filePath);
      });
      nextInputIndex += segs.length;
      return base;
    });

    // Whether each video segment brought its own audio stream — a silent
    // source clip (or a container ffprobe can't read) needs synthesized
    // silence instead, so this has to be known before the audio filter
    // graph below decides which source to pull from.
    const ffprobeBinary = await resolveFfprobeBinary(ffmpegBinary);
    await Promise.all(segments.map(async seg => {
      if (seg.kind === 'video') {
        seg.hasAudioStream = await probeHasAudioStream(ffprobeBinary, seg.filePath);
      }
    }));

    // Every background-music track (AUDIO, AUDIO2, ...) is looped and mixed
    // in additively — audio has no z-order the way video tracks do.
    const audioPaths = validAudioTracks.map((audio, idx) => {
      const mime = audio.mimeType || '';
      const ext = mime.includes('wav') ? 'wav' : mime.includes('ogg') ? 'ogg' : (mime.includes('mp4') || mime.includes('m4a')) ? 'm4a' : 'mp3';
      const audioPath = path.join(workDir, `audio_${idx}.${ext}`);
      fs.writeFileSync(audioPath, Buffer.from(audio.data, 'base64'));
      return audioPath;
    });
    const audioInputIndices = audioPaths.map((_, idx) => nextInputIndex + idx);
    nextInputIndex += audioPaths.length;

    const chains = [];
    // Ken Burns supersampling factor: zoompan rounds its crop rectangle to
    // whole pixels of whatever it's fed each frame. Feeding it the final
    // (small) output size makes slow pans/zooms visibly stutter — each
    // sub-pixel step rounds to the same pixel for several frames, then
    // jumps. Running zoompan at a larger size and scaling down afterward
    // gives it enough sub-pixel headroom to move smoothly.
    const KEN_BURNS_SUPERSAMPLE = 2;
    segments.forEach((seg, i) => {
      if (seg.kind === 'gap') {
        chains.push(`[${i}:v]format=yuv420p[v${i}]`);
        return;
      }
      if (seg.kind === 'video') {
        // Photos are already cover-cropped to WxH client-side before upload
        // (makeCoverJpeg) — a video clip isn't, so the server has to do the
        // equivalent crop itself: scale up until both dimensions cover the
        // frame, then center-crop the overflow. No per-clip crop focus yet,
        // always centered.
        let chain = `[${i}:v]scale=${W}:${H}:force_original_aspect_ratio=increase,crop=${W}:${H},setsar=1,fps=${FPS}`;
        if (colorFilter) chain += `,${colorFilter}`;
        chain += `,format=yuv420p[v${i}]`;
        chains.push(chain);
        return;
      }
      const p = seg.photo;
      let chain;
      if (p.kenBurns) {
        const zf = clampNum(p.kenBurns.zoomFrom, 1, 0.5, 3);
        const zt = clampNum(p.kenBurns.zoomTo, 1.15, 0.5, 3);
        const panX = clampNum(p.kenBurns.panX, 0, -0.3, 0.3);
        const panY = clampNum(p.kenBurns.panY, 0, -0.3, 0.3);
        const totalFrames = Math.max(1, Math.round(p.duration * FPS));
        const ssW = W * KEN_BURNS_SUPERSAMPLE, ssH = H * KEN_BURNS_SUPERSAMPLE;
        chain = `[${i}:v]scale=${ssW}:${ssH},setsar=1,` +
          `zoompan=z='${zf}+(${zt}-${zf})*on/${totalFrames}':x='iw/2-(iw/zoom/2)+${panX}*iw*on/${totalFrames}':y='ih/2-(ih/zoom/2)+${panY}*ih*on/${totalFrames}':d=1:s=${ssW}x${ssH}:fps=${FPS},` +
          `scale=${W}:${H}`;
      } else {
        chain = `[${i}:v]scale=${W}:${H},setsar=1,fps=${FPS}`;
      }
      if (colorFilter) chain += `,${colorFilter}`;
      chain += `,format=yuv420p[v${i}]`;
      chains.push(chain);
    });

    // A video clip's own audio only gets folded into the mix when at least
    // one clip actually has one — otherwise this whole per-segment audio
    // graph is skipped and the render falls back to exactly the old,
    // simpler path (background music mapped straight through, full volume,
    // no needless amix attenuation on the common photos-only case).
    const anyVideoHasAudio = segments.some(seg => seg.kind === 'video' && seg.hasAudioStream);
    if (anyVideoHasAudio) {
      segments.forEach((seg, i) => {
        const dur = segmentDuration(seg).toFixed(3);
        if (seg.kind === 'video' && seg.hasAudioStream) {
          chains.push(`[${i}:a]atrim=duration=${dur},asetpts=PTS-STARTPTS,aformat=sample_rates=44100:channel_layouts=stereo[a${i}]`);
        } else {
          chains.push(`anullsrc=r=44100:cl=stereo,atrim=duration=${dur},asetpts=PTS-STARTPTS[a${i}]`);
        }
      });
    }

    // Walk the segments left to right, deciding per junction (per adjacent
    // pair) whether to hard-cut (concat, no time overlap) or crossfade
    // (xfade, which eats "td" seconds from the running total) — mixing both
    // styles across one video is exactly what makes per-pair transitions
    // work. Any junction touching a gap is always a hard cut — dissolving
    // into solid black doesn't read as a real transition. When a video
    // clip's audio is in play, the audio track is built in lockstep with the
    // same cut/crossfade decisions (concat/acrossfade mirroring concat/xfade)
    // so it never drifts out of sync with — or runs longer than — the video.
    let finalLabel;
    let finalAudioLabel = anyVideoHasAudio ? 'a0' : null;
    let totalDur;
    if (segments.length === 1) {
      finalLabel = 'v0';
      totalDur = segmentDuration(segments[0]);
    } else {
      let runningTotal = segmentDuration(segments[0]);
      let prevLabel = 'v0';
      for (let i = 1; i < segments.length; i++) {
        const segA = segments[i - 1], segB = segments[i];
        // Only a genuine photo-to-photo junction (no gap spliced between
        // them) maps to a "transitions[]" override — segA.photoIndex lines
        // up with that array exactly when no gap was inserted at this point.
        const junction = (segA.kind === 'gap' || segB.kind === 'gap')
          ? 'cut'
          : junctionTransitionFor(segA.photoIndex);
        const isLast = i === segments.length - 1;
        // concat's output timebase doesn't match a plain fps-filtered
        // stream's — feeding that mismatch into a later xfade makes ffmpeg
        // refuse to run, so every junction's output gets re-normalized
        // (fps + format) before it can feed the next junction.
        const rawLabel = isLast ? 'vmerged_raw' : `vx${i}_raw`;
        const outLabel = isLast ? 'vmerged' : `vx${i}`;
        const rawAudioLabel = isLast ? 'amerged_raw' : `ax${i}_raw`;
        const outAudioLabel = isLast ? 'amerged' : `ax${i}`;
        if (junction === 'none' || junction === 'cut') {
          chains.push(`[${prevLabel}][v${i}]concat=n=2:v=1:a=0[${rawLabel}]`);
          if (anyVideoHasAudio) chains.push(`[${finalAudioLabel}][a${i}]concat=n=2:v=0:a=1[${rawAudioLabel}]`);
          runningTotal = runningTotal + segmentDuration(segB);
        } else {
          const offset = Math.max(0, runningTotal - td);
          chains.push(`[${prevLabel}][v${i}]xfade=transition=${junction}:duration=${td}:offset=${offset.toFixed(3)}[${rawLabel}]`);
          if (anyVideoHasAudio) chains.push(`[${finalAudioLabel}][a${i}]acrossfade=d=${td}[${rawAudioLabel}]`);
          runningTotal = runningTotal + segmentDuration(segB) - td;
        }
        chains.push(`[${rawLabel}]fps=${FPS},format=yuv420p[${outLabel}]`);
        if (anyVideoHasAudio) {
          chains.push(`[${rawAudioLabel}]aresample=async=1[${outAudioLabel}]`);
          finalAudioLabel = outAudioLabel;
        }
        prevLabel = outLabel;
      }
      finalLabel = prevLabel;
      totalDur = runningTotal;
    }

    // An extra CLIP track is allowed to run past the end of the base track;
    // when it does, the base video is padded with black so the overlay
    // still has something to sit on for its full length.
    const extraTracksEnd = extraTrackSegmentsList.reduce((max, segs) => {
      const end = segs.length ? segs[segs.length-1].end : 0;
      return Math.max(max, end);
    }, 0);
    if (extraTracksEnd > totalDur + 0.001) {
      const padBy = extraTracksEnd - totalDur;
      chains.push(`[${finalLabel}]tpad=stop_mode=add:stop_duration=${padBy.toFixed(3)}:color=black[vpadded]`);
      finalLabel = 'vpadded';
      totalDur = extraTracksEnd;
    }

    // Composite every extra CLIP track on top, in order: each segment gets
    // scaled/cropped to the frame like any photo, then overlaid onto the
    // running composite only during its own [start,end) window on that
    // track's own timeline — a gap there (or the track ending early) just
    // leaves whatever's underneath showing through untouched. Track order
    // in the array is z-order (a later track overlays an earlier one).
    extraTrackSegmentsList.forEach((segs, trackIdx) => {
      const inputBase = extraTrackInputBases[trackIdx];
      segs.forEach((seg2, i2) => {
        const inputIdx = inputBase + i2;
        const ovLabel = `ov_${trackIdx}_${i2}`;
        chains.push(`[${inputIdx}:v]scale=${W}:${H}:force_original_aspect_ratio=increase,crop=${W}:${H},setsar=1,fps=${FPS},format=yuv420p[${ovLabel}]`);
        const mergedLabel = `vmix_${trackIdx}_${i2}`;
        chains.push(`[${finalLabel}][${ovLabel}]overlay=x=0:y=0:enable='between(t,${seg2.start.toFixed(3)},${Math.min(seg2.end, totalDur).toFixed(3)})'[${mergedLabel}]`);
        finalLabel = mergedLabel;
      });
    });

    const outFilters = [];
    if (fadeInStart) outFilters.push(`fade=t=in:st=0:d=${fd}`);
    if (fadeOutEnd) outFilters.push(`fade=t=out:st=${Math.max(0, totalDur - fd).toFixed(3)}:d=${fd}`);

    // Each caption's text goes into its own temp file so drawtext's textfile=
    // reads it verbatim — sidesteps escaping arbitrary Korean/punctuation
    // text for the filtergraph parser entirely. enable='between(t,...)'
    // times each one to only its own caption window.
    // With no usable font, burning captions in is impossible — drop them and
    // still deliver the video rather than failing the whole export. The header
    // below tells the client so it can say why they are missing.
    const safeFontFile = CAPTION_FONT_FILE ? escapeDrawtextArg(CAPTION_FONT_FILE) : null;
    const captionsToDraw = safeFontFile ? validCaptions : [];
    if(validCaptions.length && !safeFontFile){
      console.warn('No caption font found on this host — captions were skipped. Set CAPTION_FONT_FILE or install fonts-nanum.');
    }
    captionsToDraw.forEach((c, idx) => {
      const txtPath = path.join(workDir, `caption_${idx}.txt`);
      fs.writeFileSync(txtPath, c.text, 'utf8');
      const padX = Math.round(c.size * 0.4);
      const padY = Math.round(c.size * 0.25);
      const fontColorHex = '0x' + c.color.replace('#', '');
      const boxPart = c.box ? `box=1:boxcolor=black@0.55:boxborderw=${padY}|${padX}:` : '';
      const borderPart = c.borderWidth > 0 ? `bordercolor=${c.borderColor}:borderw=${c.borderWidth}:` : '';
      const shadowPart = (c.shadowOffsetX || c.shadowOffsetY) ? `shadowcolor=${c.shadowColor}:shadowx=${c.shadowOffsetX}:shadowy=${c.shadowOffsetY}:` : '';
      outFilters.push(
        `drawtext=textfile='${escapeDrawtextArg(txtPath)}':fontfile='${safeFontFile}':` +
        `fontsize=${c.size}:fontcolor=${fontColorHex}:x=(${c.x}*w)-(text_w/2):y=(${c.y}*h)-(text_h/2):` +
        `${boxPart}${borderPart}${shadowPart}` +
        `enable='between(t,${c.start.toFixed(3)},${c.end.toFixed(3)})'`
      );
    });

    let outputLabel = finalLabel;
    if (outFilters.length) {
      chains.push(`[${finalLabel}]${outFilters.join(',')}[vout]`);
      outputLabel = 'vout';
    }

    // Every background-music track mixes with a video clip's own audio
    // (when there is one) rather than replacing it — amix's default
    // normalization halves each input, which is the right tradeoff when
    // multiple real audio sources are actually being blended, but would
    // needlessly quiet a single source, so amix only ever runs when
    // there's more than one to combine.
    const audioMixInputs = [];
    if (anyVideoHasAudio) audioMixInputs.push(`[${finalAudioLabel}]`);
    audioInputIndices.forEach(idx => audioMixInputs.push(`[${idx}:a]`));

    let finalAudioMapLabel = null;
    let directAudioMapIndex = null;
    if (audioMixInputs.length > 1) {
      chains.push(`${audioMixInputs.join('')}amix=inputs=${audioMixInputs.length}:duration=first:dropout_transition=0[afinal]`);
      finalAudioMapLabel = 'afinal';
    } else if (audioMixInputs.length === 1) {
      if (anyVideoHasAudio) finalAudioMapLabel = finalAudioLabel;
      else directAudioMapIndex = audioInputIndices[0];
    }

    const outputPath = path.join(workDir, 'output.mp4');
    const args = ['-y', ...inputArgs];
    audioPaths.forEach(audioPath => { args.push('-stream_loop', '-1', '-i', audioPath); });
    args.push('-filter_complex', chains.join(';'), '-map', `[${outputLabel}]`);
    if (finalAudioMapLabel) {
      args.push('-map', `[${finalAudioMapLabel}]`, '-c:a', 'aac', '-shortest');
    } else if (directAudioMapIndex != null) {
      args.push('-map', `${directAudioMapIndex}:a`, '-c:a', 'aac', '-shortest');
    }
    args.push('-c:v', 'libx264', '-pix_fmt', 'yuv420p', '-r', String(FPS), '-t', totalDur.toFixed(3), outputPath);

    await runFfmpeg(ffmpegBinary, args);

    if(validCaptions.length && !safeFontFile){
      res.setHeader('X-Caption-Warning', 'no-font-on-server');
    }
    res.setHeader('Content-Type', 'video/mp4');
    res.setHeader('Content-Disposition', 'attachment; filename="vibe-cut.mp4"');
    const stream = fs.createReadStream(outputPath);
    stream.pipe(res);
    stream.on('close', () => fs.rm(workDir, { recursive: true, force: true }, () => {}));
    stream.on('error', () => fs.rm(workDir, { recursive: true, force: true }, () => {}));
  } catch (err) {
    fs.rm(workDir, { recursive: true, force: true }, () => {});
    res.status(500).json({ error: 'Render failed: ' + err.message });
  }
});

const PORT = process.env.PORT || 8787;
app.listen(PORT, () => {
  console.log(`vibe-cut running at http://localhost:${PORT}`);
  if (!API_KEY) {
    console.warn('WARNING: GEMINI_API_KEY is not set — the AI director feature will fail until you set it (see .env.example).');
  }
  resolveFfmpegBinary().then(bin => {
    if (!bin) console.warn('WARNING: ffmpeg was not found — video rendering will fail until you run "brew install ffmpeg" and restart the server.');
  });
});
