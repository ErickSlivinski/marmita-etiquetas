import crypto from 'crypto';
import express from 'express';
import { getAuthStatus, verifyRawMcpToken } from './serverStore.js';

const EPHEMERAL_OAUTH_KEY = crypto.randomBytes(32);
const USED_AUTH_CODES = new Map(); // jti -> expiresAtMs

function getOAuthSigningKey() {
  const envToken = (process.env.MCP_API_TOKEN || '').trim();
  if (envToken) {
    return crypto.createHmac('sha256', 'etiquetafit-oauth-signing-v1').update(envToken).digest();
  }
  return EPHEMERAL_OAUTH_KEY;
}

function base64UrlEncode(bufOrStr) {
  return Buffer.from(bufOrStr)
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/g, '');
}

function base64UrlDecode(str) {
  const padded = String(str || '')
    .replace(/-/g, '+')
    .replace(/_/g, '/')
    .padEnd(Math.ceil(String(str || '').length / 4) * 4, '=');
  return Buffer.from(padded, 'base64');
}

function signTokenPayload(prefix, payload, ttlMs) {
  const body = {
    ...payload,
    iat: Date.now(),
    exp: Date.now() + ttlMs
  };
  const encoded = base64UrlEncode(JSON.stringify(body));
  const sig = base64UrlEncode(
    crypto.createHmac('sha256', getOAuthSigningKey()).update(`${prefix}.${encoded}`).digest()
  );
  return `${prefix}${encoded}.${sig}`;
}

function verifyTokenPayload(token, prefix) {
  if (typeof token !== 'string' || !token.startsWith(prefix)) return null;
  const rest = token.slice(prefix.length);
  const dotIdx = rest.lastIndexOf('.');
  if (dotIdx <= 0) return null;

  const encoded = rest.slice(0, dotIdx);
  const providedSig = rest.slice(dotIdx + 1);
  if (!encoded || !providedSig) return null;

  const expectedSig = base64UrlEncode(
    crypto.createHmac('sha256', getOAuthSigningKey()).update(`${prefix}.${encoded}`).digest()
  );

  const a = Buffer.from(providedSig, 'utf8');
  const b = Buffer.from(expectedSig, 'utf8');
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) {
    return null;
  }

  try {
    const parsed = JSON.parse(base64UrlDecode(encoded).toString('utf8'));
    if (!parsed || typeof parsed.exp !== 'number' || parsed.exp < Date.now()) {
      return null;
    }
    return parsed;
  } catch {
    return null;
  }
}

export function isAllowedRedirectUri(rawUri) {
  if (typeof rawUri !== 'string' || !rawUri.trim()) return false;
  try {
    const parsed = new URL(rawUri.trim());
    if (parsed.username || parsed.password) return false;

    const host = parsed.hostname.toLowerCase();
    if (parsed.protocol === 'https:') {
      return (
        host === 'claude.ai' ||
        host.endsWith('.claude.ai') ||
        host === 'claude.com' ||
        host.endsWith('.claude.com') ||
        host.endsWith('.anthropic.com')
      );
    }

    if (parsed.protocol === 'http:') {
      return host === 'localhost' || host === '127.0.0.1';
    }

    return false;
  } catch {
    return false;
  }
}

function pruneUsedCodes() {
  const now = Date.now();
  for (const [jti, exp] of USED_AUTH_CODES.entries()) {
    if (exp <= now) USED_AUTH_CODES.delete(jti);
  }
}

function escHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function verifyPkceS256(codeVerifier, expectedChallenge) {
  if (typeof codeVerifier !== 'string' || typeof expectedChallenge !== 'string') return false;
  if (codeVerifier.length < 43 || codeVerifier.length > 128) return false;
  const computed = base64UrlEncode(crypto.createHash('sha256').update(codeVerifier).digest());
  const a = Buffer.from(computed, 'utf8');
  const b = Buffer.from(expectedChallenge, 'utf8');
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}

function validateOAuthClientAndRedirect(clientId, redirectUri) {
  if (!isAllowedRedirectUri(redirectUri)) {
    return { ok: false, error: 'redirect_uri não permitido.' };
  }
  if (typeof clientId !== 'string' || !clientId.trim()) {
    return { ok: false, error: 'client_id ausente.' };
  }

  // Stateless registered client created via /oauth/register
  if (clientId.startsWith('etq_cli_')) {
    const payload = verifyTokenPayload(clientId, 'etq_cli_');
    if (!payload || !Array.isArray(payload.r) || !payload.r.includes(redirectUri)) {
      return { ok: false, error: 'client_id inválido ou redirect_uri não registrado.' };
    }
    return { ok: true, clientName: payload.n || 'Claude' };
  }

  // Fallback for manually provided client_id (only allowlisted redirect URIs pass above)
  return { ok: true, clientName: 'Claude' };
}

function renderAuthorizePage({
  clientId,
  redirectUri,
  state,
  codeChallenge,
  codeChallengeMethod,
  scope,
  resource,
  csrfToken,
  clientName,
  tokenConfigured,
  errorMsg = ''
}) {
  const errHtml = errorMsg
    ? `<div class="err" role="alert">${escHtml(errorMsg)}</div>`
    : '';
  const noTokenWarn = !tokenConfigured
    ? `<div class="err">Nenhum token MCP foi configurado no servidor ainda. Configure <code>MCP_API_TOKEN</code> no servidor ou gere um token em <strong>Backup e Dados → Claude (MCP)</strong>.</div>`
    : '';

  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Autorizar Claude — EtiquetaFit</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      background: #f4f6f8; color: #0f172a; min-height: 100vh;
      display: flex; align-items: center; justify-content: center; padding: 20px;
    }
    .card {
      width: min(440px, 100%); background: #fff; border: 1px solid #e2e8f0;
      border-radius: 14px; padding: 26px; box-shadow: 0 16px 40px rgba(15,23,42,0.1);
    }
    h1 { font-size: 19px; font-weight: 700; margin-bottom: 8px; }
    p { font-size: 14px; color: #475569; line-height: 1.5; margin-bottom: 16px; }
    .meta {
      background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px;
      padding: 10px 12px; font-size: 13px; color: #334155; margin-bottom: 16px;
    }
    label { display: block; font-size: 13px; font-weight: 600; color: #334155; margin-bottom: 6px; }
    input[type="password"] {
      width: 100%; height: 42px; padding: 8px 12px; border: 1px solid #cbd5e1;
      border-radius: 8px; font-size: 14px; font-family: monospace; margin-bottom: 14px;
    }
    button {
      width: 100%; height: 44px; border: none; border-radius: 8px;
      background: #0f766e; color: #fff; font-size: 14.5px; font-weight: 600; cursor: pointer;
    }
    button:hover { background: #0d625b; }
    .err {
      background: #fef2f2; border: 1px solid #fecaca; color: #b91c1c;
      border-radius: 8px; padding: 10px 12px; font-size: 13px; margin-bottom: 14px;
    }
    code { font-family: monospace; background: rgba(15,23,42,0.06); padding: 1px 4px; border-radius: 4px; }
  </style>
</head>
<body>
  <main class="card">
    <h1>Autorizar acesso ao EtiquetaFit</h1>
    <p>O conector <strong>${escHtml(clientName)}</strong> está solicitando permissão para consultar seu catálogo de etiquetas e gerar PDFs A4.</p>
    <div class="meta">Destino: <code>${escHtml(new URL(redirectUri).host)}</code></div>
    ${noTokenWarn}
    ${errHtml}
    <form method="POST" action="/oauth/authorize">
      <input type="hidden" name="client_id" value="${escHtml(clientId)}" />
      <input type="hidden" name="redirect_uri" value="${escHtml(redirectUri)}" />
      <input type="hidden" name="state" value="${escHtml(state)}" />
      <input type="hidden" name="code_challenge" value="${escHtml(codeChallenge)}" />
      <input type="hidden" name="code_challenge_method" value="${escHtml(codeChallengeMethod)}" />
      <input type="hidden" name="scope" value="${escHtml(scope)}" />
      <input type="hidden" name="resource" value="${escHtml(resource)}" />
      <input type="hidden" name="csrf" value="${escHtml(csrfToken)}" />
      <label for="apiToken">Digite seu Token MCP (<code>MCP_API_TOKEN</code>)</label>
      <input id="apiToken" name="apiToken" type="password" required autocomplete="off" placeholder="etq_..." />
      <button type="submit">Autorizar e conectar</button>
    </form>
  </main>
</body>
</html>`;
}

export function verifyOAuthBearerToken(authHeader) {
  if (typeof authHeader !== 'string') return false;
  const match = authHeader.match(/^Bearer\s+(.+)$/i);
  if (!match) return false;
  const payload = verifyTokenPayload(match[1].trim(), 'etq_oa_');
  return Boolean(payload && payload.t === 'access');
}

export function issueOAuthCodeAndTokensForTest({ clientId, redirectUri, codeChallenge }) {
  const jti = crypto.randomBytes(12).toString('hex');
  const code = signTokenPayload(
    'etq_oc_',
    { t: 'code', jti, c: clientId, r: redirectUri, ch: codeChallenge },
    5 * 60 * 1000
  );
  return { code, jti };
}

export function exchangeOAuthCode({ code, clientId, redirectUri, codeVerifier }) {
  pruneUsedCodes();
  const payload = verifyTokenPayload(code, 'etq_oc_');
  if (!payload || payload.t !== 'code') {
    return { error: 'invalid_grant', error_description: 'Código de autorização inválido ou expirado.' };
  }
  if (USED_AUTH_CODES.has(payload.jti)) {
    return { error: 'invalid_grant', error_description: 'Código de autorização já utilizado.' };
  }
  if (payload.c !== clientId || payload.r !== redirectUri) {
    return { error: 'invalid_grant', error_description: 'client_id ou redirect_uri não confere.' };
  }
  if (!verifyPkceS256(codeVerifier, payload.ch)) {
    return { error: 'invalid_grant', error_description: 'Falha na verificação PKCE code_verifier.' };
  }

  USED_AUTH_CODES.set(payload.jti, payload.exp);

  const accessToken = signTokenPayload(
    'etq_oa_',
    { t: 'access', c: clientId },
    30 * 24 * 60 * 60 * 1000
  );
  const refreshToken = signTokenPayload(
    'etq_or_',
    { t: 'refresh', c: clientId },
    90 * 24 * 60 * 60 * 1000
  );

  return {
    access_token: accessToken,
    token_type: 'Bearer',
    expires_in: 30 * 24 * 60 * 60,
    refresh_token: refreshToken,
    scope: 'etiquetafit:orders'
  };
}

export function registerOAuthRoutes(app, resolvePublicBaseUrl) {
  const urlEncoded = express.urlencoded({ extended: false });

  // RFC 9728 Protected Resource Metadata
  const sendProtectedResource = (req, res) => {
    const baseUrl = resolvePublicBaseUrl(req);
    res.json({
      resource: `${baseUrl}/mcp`,
      authorization_servers: [baseUrl],
      bearer_methods_supported: ['header'],
      scopes_supported: ['etiquetafit:orders']
    });
  };
  app.get('/.well-known/oauth-protected-resource', sendProtectedResource);
  app.get('/.well-known/oauth-protected-resource/mcp', sendProtectedResource);

  // RFC 8414 Authorization Server Metadata
  const sendAuthServerMeta = (req, res) => {
    const baseUrl = resolvePublicBaseUrl(req);
    res.json({
      issuer: baseUrl,
      authorization_endpoint: `${baseUrl}/oauth/authorize`,
      token_endpoint: `${baseUrl}/oauth/token`,
      registration_endpoint: `${baseUrl}/oauth/register`,
      response_types_supported: ['code'],
      grant_types_supported: ['authorization_code', 'refresh_token'],
      code_challenge_methods_supported: ['S256'],
      token_endpoint_auth_methods_supported: ['none', 'client_secret_post'],
      scopes_supported: ['etiquetafit:orders']
    });
  };
  app.get('/.well-known/oauth-authorization-server', sendAuthServerMeta);
  app.get('/.well-known/openid-configuration', sendAuthServerMeta);

  // RFC 7591 Dynamic Client Registration (DCR)
  app.post('/oauth/register', (req, res) => {
    const body = req.body || {};
    const redirectUris = Array.isArray(body.redirect_uris) ? body.redirect_uris : [];
    if (redirectUris.length === 0 || !redirectUris.every(isAllowedRedirectUri)) {
      return res.status(400).json({
        error: 'invalid_redirect_uri',
        error_description: 'Informe redirect_uris válidos do Claude ou localhost.'
      });
    }

    const clientName = String(body.client_name || 'Claude').slice(0, 64);
    const clientId = signTokenPayload(
      'etq_cli_',
      { n: clientName, r: redirectUris },
      365 * 24 * 60 * 60 * 1000
    );

    return res.status(201).json({
      client_id: clientId,
      client_name: clientName,
      redirect_uris: redirectUris,
      grant_types: ['authorization_code', 'refresh_token'],
      response_types: ['code'],
      token_endpoint_auth_method: 'none'
    });
  });

  // OAuth 2.1 Authorization UI
  app.get('/oauth/authorize', async (req, res) => {
    const {
      response_type: responseType = '',
      client_id: clientId = '',
      redirect_uri: redirectUri = '',
      state = '',
      code_challenge: codeChallenge = '',
      code_challenge_method: codeChallengeMethod = 'S256',
      scope = 'etiquetafit:orders',
      resource = ''
    } = req.query || {};

    if (responseType !== 'code') {
      return res.status(400).send('response_type inválido (esperado: code).');
    }
    const clientCheck = validateOAuthClientAndRedirect(String(clientId), String(redirectUri));
    if (!clientCheck.ok) {
      return res.status(400).send(clientCheck.error);
    }
    if (!codeChallenge || codeChallengeMethod !== 'S256') {
      return res.status(400).send('PKCE code_challenge (S256) é obrigatório.');
    }

    const authStatus = await getAuthStatus();
    const csrfToken = signTokenPayload('etq_csrf_', { c: String(clientId) }, 15 * 60 * 1000);

    return res.status(200).send(
      renderAuthorizePage({
        clientId: String(clientId),
        redirectUri: String(redirectUri),
        state: String(state),
        codeChallenge: String(codeChallenge),
        codeChallengeMethod: String(codeChallengeMethod),
        scope: String(scope),
        resource: String(resource),
        csrfToken,
        clientName: clientCheck.clientName,
        tokenConfigured: authStatus.configured
      })
    );
  });

  app.post('/oauth/authorize', urlEncoded, async (req, res) => {
    const {
      client_id: clientId = '',
      redirect_uri: redirectUri = '',
      state = '',
      code_challenge: codeChallenge = '',
      code_challenge_method: codeChallengeMethod = 'S256',
      scope = 'etiquetafit:orders',
      resource = '',
      csrf = '',
      apiToken = ''
    } = req.body || {};

    const clientCheck = validateOAuthClientAndRedirect(String(clientId), String(redirectUri));
    if (!clientCheck.ok) {
      return res.status(400).send(clientCheck.error);
    }
    if (!codeChallenge || codeChallengeMethod !== 'S256') {
      return res.status(400).send('PKCE code_challenge (S256) é obrigatório.');
    }

    const csrfPayload = verifyTokenPayload(String(csrf), 'etq_csrf_');
    const tokenOk = await verifyRawMcpToken(String(apiToken));
    if (!csrfPayload || csrfPayload.c !== String(clientId) || !tokenOk) {
      const authStatus = await getAuthStatus();
      const nextCsrf = signTokenPayload('etq_csrf_', { c: String(clientId) }, 15 * 60 * 1000);
      return res.status(401).send(
        renderAuthorizePage({
          clientId: String(clientId),
          redirectUri: String(redirectUri),
          state: String(state),
          codeChallenge: String(codeChallenge),
          codeChallengeMethod: String(codeChallengeMethod),
          scope: String(scope),
          resource: String(resource),
          csrfToken: nextCsrf,
          clientName: clientCheck.clientName,
          tokenConfigured: authStatus.configured,
          errorMsg: !tokenOk
            ? 'Token MCP inválido. Confira o valor de MCP_API_TOKEN no servidor.'
            : 'Sessão de autorização expirada. Tente novamente.'
        })
      );
    }

    const jti = crypto.randomBytes(12).toString('hex');
    const code = signTokenPayload(
      'etq_oc_',
      {
        t: 'code',
        jti,
        c: String(clientId),
        r: String(redirectUri),
        ch: String(codeChallenge)
      },
      5 * 60 * 1000
    );

    const callbackUrl = new URL(String(redirectUri));
    callbackUrl.searchParams.set('code', code);
    if (state) callbackUrl.searchParams.set('state', String(state));

    return res.redirect(302, callbackUrl.toString());
  });

  // OAuth 2.1 Token Exchange
  app.post('/oauth/token', urlEncoded, (req, res) => {
    const body = req.body || {};
    const grantType = String(body.grant_type || '');

    if (grantType === 'authorization_code') {
      const result = exchangeOAuthCode({
        code: String(body.code || ''),
        clientId: String(body.client_id || ''),
        redirectUri: String(body.redirect_uri || ''),
        codeVerifier: String(body.code_verifier || '')
      });
      if (result.error) {
        return res.status(400).json(result);
      }
      res.setHeader('Cache-Control', 'no-store');
      return res.json(result);
    }

    if (grantType === 'refresh_token') {
      const refreshPayload = verifyTokenPayload(String(body.refresh_token || ''), 'etq_or_');
      if (!refreshPayload || refreshPayload.t !== 'refresh') {
        return res.status(400).json({ error: 'invalid_grant', error_description: 'refresh_token inválido ou expirado.' });
      }
      const clientId = String(body.client_id || refreshPayload.c || '');
      if (clientId && refreshPayload.c && clientId !== refreshPayload.c) {
        return res.status(400).json({ error: 'invalid_grant', error_description: 'client_id não confere.' });
      }

      const accessToken = signTokenPayload(
        'etq_oa_',
        { t: 'access', c: refreshPayload.c },
        30 * 24 * 60 * 60 * 1000
      );
      const nextRefresh = signTokenPayload(
        'etq_or_',
        { t: 'refresh', c: refreshPayload.c },
        90 * 24 * 60 * 60 * 1000
      );
      res.setHeader('Cache-Control', 'no-store');
      return res.json({
        access_token: accessToken,
        token_type: 'Bearer',
        expires_in: 30 * 24 * 60 * 60,
        refresh_token: nextRefresh,
        scope: 'etiquetafit:orders'
      });
    }

    return res.status(400).json({
      error: 'unsupported_grant_type',
      error_description: 'Suporte apenas para authorization_code e refresh_token.'
    });
  });
}
