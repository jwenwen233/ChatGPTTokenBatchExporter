(function attachTokenExporterSessionReader() {
  if (window.__TOKEN_EXPORTER_SESSION_READY__) return;
  window.__TOKEN_EXPORTER_SESSION_READY__ = true;
  const HANDLED_TYPES = new Set([
    'TK_SESSION_PING',
    'TK_READ_SESSION',
  ]);

  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    if (!message || !HANDLED_TYPES.has(message.type)) return false;
    Promise.resolve(handleMessage(message))
      .then((result) => sendResponse({ ok: true, ...(result || {}) }))
      .catch((error) => sendResponse({ ok: false, error: error?.message || String(error || '') }));
    return true;
  });

  async function handleMessage(message) {
    if (message.type === 'TK_SESSION_PING') {
      return { ready: true };
    }
    if (message.type === 'TK_READ_SESSION') {
      const response = await fetch('/api/auth/session', { credentials: 'include' });
      const session = await response.json().catch(() => ({}));
      const accessToken = String(session?.accessToken || '').trim();
      if (!response.ok && !accessToken) {
        throw new Error(`读取 ChatGPT session 失败：HTTP ${response.status}`);
      }
      return {
        status: response.status,
        session,
        accessToken,
        email: String(session?.user?.email || session?.email || '').trim(),
      };
    }
    return {};
  }
})();
