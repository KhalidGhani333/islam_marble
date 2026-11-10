import {
  deletePass,
  getSettings,
  isAuthenticated,
  listPasses,
  savePass,
  saveSettings,
  sessionCookie,
  verifyPassword,
} from './sheets.js';

async function readBody(req) {
  if (req.body && typeof req.body === 'object') return req.body;
  const chunks = [];
  for await (const chunk of req) chunks.push(Buffer.from(chunk));
  if (!chunks.length) return {};
  return JSON.parse(Buffer.concat(chunks).toString('utf8'));
}

function respondError(res, error) {
  console.error('API request failed:', error);
  if (res.headersSent) return;
  const status = Number(error.status || error.statusCode);
  const statusCode = Number.isInteger(status) && status >= 400 && status <= 599 ? status : 500;
  res.statusCode = statusCode;
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.end(JSON.stringify({ error: error.message || 'Internal server error.' }));
}

export default async function apiHandler(req, res) {
  try {
    const url = new URL(req.url, 'http://localhost');
    const path = url.pathname.replace(/^\/api\/?/, '').replace(/\/+$/, '');

    if (path === 'auth/status' && req.method === 'GET') {
      res.statusCode = 200;
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      res.end(JSON.stringify({
        configured: Boolean(process.env.APP_PASSWORD && process.env.APP_PASSWORD.length >= 12),
        authenticated: isAuthenticated(req),
      }));
      return;
    }

    if (path === 'auth/login' && req.method === 'POST') {
      const body = await readBody(req);
      if (!verifyPassword(body.password)) {
        res.statusCode = 401;
        res.setHeader('Content-Type', 'application/json; charset=utf-8');
        res.end(JSON.stringify({ error: 'Incorrect app password.' }));
        return;
      }
      res.statusCode = 200;
      res.setHeader('Set-Cookie', sessionCookie(process.env.APP_PASSWORD));
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      res.end(JSON.stringify({ authenticated: true }));
      return;
    }

    if (path === 'auth/logout' && req.method === 'POST') {
      const secure = process.env.NODE_ENV === 'production' ? '; Secure' : '';
      res.statusCode = 200;
      res.setHeader('Set-Cookie', `islam_marble_session=; HttpOnly; SameSite=Strict; Path=/; Max-Age=0${secure}`);
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      res.end(JSON.stringify({ authenticated: false }));
      return;
    }

    if (!isAuthenticated(req)) {
      res.statusCode = 401;
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      res.end(JSON.stringify({ error: 'Sign in to access the gate-pass data.' }));
      return;
    }

    if (path === 'bootstrap' && req.method === 'GET') {
      const [settings, passes] = await Promise.all([getSettings(), listPasses()]);
      res.statusCode = 200;
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      res.end(JSON.stringify({ settings, passes }));
      return;
    }

    if (path === 'settings' && req.method === 'PUT') {
      const settings = await saveSettings(await readBody(req));
      res.statusCode = 200;
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      res.end(JSON.stringify(settings));
      return;
    }

    if (path === 'gate-passes' && req.method === 'POST') {
      const pass = await savePass(await readBody(req));
      res.statusCode = 201;
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      res.end(JSON.stringify(pass));
      return;
    }

    const passRoute = path.match(/^gate-passes\/(\d+)(\/duplicate)?$/);
    if (passRoute && passRoute[2] && req.method === 'POST') {
      const id = Number(passRoute[1]);
      const original = (await listPasses()).find((pass) => pass.id === id);
      if (!original) {
        res.statusCode = 404;
        res.setHeader('Content-Type', 'application/json; charset=utf-8');
        res.end(JSON.stringify({ error: 'Gate pass not found in the Google Sheet.' }));
        return;
      }
      const { id: _id, gate_pass_number: _number, ...payload } = original;
      const pass = await savePass({ ...payload, gate_pass_number: '' });
      res.statusCode = 201;
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      res.end(JSON.stringify(pass));
      return;
    }
    if (passRoute && req.method === 'PUT') {
      const pass = await savePass(await readBody(req), Number(passRoute[1]));
      res.statusCode = 200;
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      res.end(JSON.stringify(pass));
      return;
    }
    if (passRoute && !passRoute[2] && req.method === 'DELETE') {
      await deletePass(Number(passRoute[1]));
      res.statusCode = 200;
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      res.end(JSON.stringify({ success: true }));
      return;
    }

    res.statusCode = 404;
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.end(JSON.stringify({ error: 'API route not found.' }));
  } catch (error) {
    respondError(res, error);
  }
}
