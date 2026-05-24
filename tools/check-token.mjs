#!/usr/bin/env node

import { existsSync, readFileSync } from 'node:fs';

const ENDPOINTS = [
  { name: 'me', url: 'https://chatgpt.com/backend-api/me' },
  { name: 'models', url: 'https://chatgpt.com/backend-api/models' },
];

function normalize(value) {
  return String(value || '').trim();
}

function decodeBase64Url(segment = '') {
  const padded = normalize(segment).replace(/-/g, '+').replace(/_/g, '/')
    + '='.repeat((4 - (segment.length % 4)) % 4);
  return Buffer.from(padded, 'base64').toString('utf8');
}

function parseJwtPayload(token = '') {
  const parts = normalize(token).split('.');
  if (parts.length < 2) return {};
  try {
    return JSON.parse(decodeBase64Url(parts[1]));
  } catch {
    return {};
  }
}

function extractToken(input = '') {
  const raw = normalize(input);
  if (!raw) return '';
  try {
    const json = JSON.parse(raw);
    const token = normalize(
      json.access_token
      || json.accessToken
      || json.token?.access_token
      || json.token?.accessToken
    );
    if (token) return token;
  } catch {
    // Fall through to plain-token parsing.
  }
  const jwt = raw.match(/eyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+/);
  return normalize(jwt ? jwt[0] : raw);
}

function maskEmail(email = '') {
  const value = normalize(email);
  const [name, domain] = value.split('@');
  if (!name || !domain) return '';
  return `${name.slice(0, 2)}***@${domain}`;
}

async function fetchWithTimeout(url, options, timeoutMs = 15000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } finally {
    clearTimeout(timer);
  }
}

function bodyHint(text = '') {
  const clean = normalize(text).replace(/\s+/g, ' ');
  if (!clean) return '';
  try {
    const json = JSON.parse(clean);
    return normalize(json.detail || json.error?.message || json.error || json.message || json.code).slice(0, 160);
  } catch {
    return clean.slice(0, 160);
  }
}

async function main() {
  const target = process.argv[2];
  if (!target) {
    throw new Error('Usage: node tools/check-token.mjs <token-or-json-file>');
  }
  const input = existsSync(target) ? readFileSync(target, 'utf8') : target;
  const token = extractToken(input);
  if (!token) throw new Error('No accessToken found.');

  const payload = parseJwtPayload(token);
  const profile = payload['https://api.openai.com/profile'] || {};
  const expMs = Number(payload.exp) ? Number(payload.exp) * 1000 : 0;
  const attempts = [];

  for (const endpoint of ENDPOINTS) {
    try {
      const response = await fetchWithTimeout(endpoint.url, {
        method: 'GET',
        headers: {
          accept: 'application/json',
          authorization: `Bearer ${token}`,
        },
      });
      const text = await response.text().catch(() => '');
      attempts.push({
        endpoint: endpoint.name,
        status: response.status,
        ok: response.ok,
        hint: bodyHint(text),
      });
    } catch (error) {
      attempts.push({
        endpoint: endpoint.name,
        status: 0,
        ok: false,
        hint: error?.name === 'AbortError' ? 'timeout' : normalize(error?.message || error),
      });
    }
  }

  const valid = attempts.some((item) => item.ok);
  const unusable = !valid && attempts.some((item) => [401, 403].includes(item.status));
  console.log(JSON.stringify({
    valid,
    unusable,
    email: maskEmail(profile.email || payload.email),
    expiresAt: expMs ? new Date(expMs).toISOString() : '',
    attempts,
  }, null, 2));
}

main().catch((error) => {
  console.error(error?.message || String(error));
  process.exit(1);
});
