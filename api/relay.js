const ALLOWED_ACTIONS = new Set([
  'getLoginOptions',
  'loginStudent',
  'startExam',
  'saveAnswers',
  'reportViolation',
  'submitExam',
]);
const MAX_REQUEST_BYTES = 100000;

function respond(res, status, body) {
  res.setHeader('Cache-Control', 'no-store');
  return res.status(status).json(body);
}

function getAppsScriptUrl() {
  const value = process.env.APPS_SCRIPT_WEB_APP_URL;
  if (!value) return null;

  try {
    const url = new URL(value);
    if (
      url.protocol !== 'https:' ||
      url.hostname !== 'script.google.com' ||
      !/^\/macros\/s\/[^/]+\/exec$/.test(url.pathname) ||
      url.pathname.includes('PASTE_DEPLOYMENT_ID')
    ) {
      return null;
    }
    return url.toString();
  } catch {
    return null;
  }
}

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return respond(res, 405, { ok: false, error: 'Metode permintaan tidak didukung.' });
  }

  const contentLength = Number(req.headers['content-length'] || 0);
  if (contentLength > MAX_REQUEST_BYTES) {
    return respond(res, 413, { ok: false, error: 'Ukuran permintaan terlalu besar.' });
  }

  let payload = req.body;
  if (typeof payload === 'string') {
    try {
      payload = JSON.parse(payload);
    } catch {
      return respond(res, 400, { ok: false, error: 'Format data permintaan tidak valid.' });
    }
  }

  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
    return respond(res, 400, { ok: false, error: 'Format data permintaan tidak valid.' });
  }

  if (Buffer.byteLength(JSON.stringify(payload), 'utf8') > MAX_REQUEST_BYTES) {
    return respond(res, 413, { ok: false, error: 'Ukuran permintaan terlalu besar.' });
  }

  if (typeof payload.action !== 'string' || !ALLOWED_ACTIONS.has(payload.action)) {
    return respond(res, 400, { ok: false, error: 'Permintaan tidak valid.' });
  }
  if (payload.args !== undefined && !Array.isArray(payload.args)) {
    return respond(res, 400, { ok: false, error: 'Format data permintaan tidak valid.' });
  }

  const appsScriptUrl = getAppsScriptUrl();
  if (!appsScriptUrl) {
    return respond(res, 500, {
      ok: false,
      error: 'Atur environment variable APPS_SCRIPT_WEB_APP_URL dengan URL deployment Apps Script yang benar.',
    });
  }

  let upstream;
  try {
    upstream = await fetch(appsScriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json; charset=utf-8' },
      body: JSON.stringify(payload),
    });
  } catch (error) {
    console.error('Apps Script relay failed:', error);
    return respond(res, 502, {
      ok: false,
      error: 'Tidak dapat menghubungi Apps Script. Periksa koneksi dan URL deployment.',
    });
  }

  let result;
  try {
    result = await upstream.json();
  } catch (error) {
    console.error('Unexpected Apps Script response:', error);
    return respond(res, 502, {
      ok: false,
      error: 'Respons Apps Script tidak valid. Pastikan versi deployment sudah diperbarui.',
    });
  }

  if (
    !upstream.ok ||
    !result ||
    typeof result !== 'object' ||
    !Object.prototype.hasOwnProperty.call(result, 'ok')
  ) {
    console.error('Unexpected Apps Script response; HTTP', upstream.status);
    return respond(res, 502, {
      ok: false,
      error: 'Respons Apps Script tidak sesuai. Pastikan versi deployment sudah diperbarui.',
    });
  }

  return respond(res, 200, result);
};
