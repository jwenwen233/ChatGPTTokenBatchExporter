const $ = (id) => document.getElementById(id);
const POPUP_DRAFT_KEY = 'tokenExporterPopupDraft';

const dom = {
  runState: $('run-state'),
  sessionJsonInput: $('session-json-input'),
  mailProvider: $('mail-provider'),
  mailImportControls: $('mail-import-controls'),
  mailImport: $('mail-import'),
  mailImportFile: $('mail-import-file'),
  microsoftConfig: $('microsoft-config'),
  microsoftLocalConfig: $('microsoft-local-config'),
  microsoftLocalBaseUrl: $('microsoft-local-base-url'),
  luckmailConfig: $('luckmail-config'),
  luckmailApiKey: $('luckmail-api-key'),
  luckmailBaseUrl: $('luckmail-base-url'),
  cloudflareTempConfig: $('cloudflare-temp-config'),
  cloudflareTempBaseUrl: $('cloudflare-temp-base-url'),
  cloudflareTempAdminAuth: $('cloudflare-temp-admin-auth'),
  cloudflareTempCustomAuth: $('cloudflare-temp-custom-auth'),
  cloudflareTempLookupMode: $('cloudflare-temp-lookup-mode'),
  cloudflareTempReceiveMailbox: $('cloudflare-temp-receive-mailbox'),
  btnConvertSessionJson: $('btn-convert-session-json'),
  btnImportMail: $('btn-import-mail'),
  btnSelectCockpitDir: $('btn-select-cockpit-dir'),
  btnPullCockpit401: $('btn-pull-cockpit-401'),
  btnRefreshCockpit: $('btn-refresh-cockpit'),
  btnStopBatch: $('btn-stop-batch'),
  btnToggleAccountsList: $('btn-toggle-accounts-list'),
  btnToggleMailList: $('btn-toggle-mail-list'),
  externalSource: $('external-source'),
  cliproxyConfig: $('cliproxy-config'),
  cliproxyBaseUrl: $('cliproxy-base-url'),
  cliproxyManagementKey: $('cliproxy-management-key'),
  cockpitStatus: $('cockpit-status'),
  cockpitList: $('cockpit-list'),
  btnSaveSettings: $('btn-save-settings'),
  btnClearLogs: $('btn-clear-logs'),
  accountsList: $('accounts-list'),
  mailList: $('mail-list'),
  logs: $('logs'),
  autoCheckEnabled: $('auto-check-enabled'),
  autoCheckMinutes: $('auto-check-minutes'),
  minTokenValidityMinutes: $('min-token-validity-minutes'),
  mailMaxRetries: $('mail-max-retries'),
  mailRetryDelayMs: $('mail-retry-delay-ms'),
  closeTabsAfterRun: $('close-tabs-after-run'),
  activateTabs: $('activate-tabs'),
  directLiveValidateAccessToken: $('direct-live-validate-access-token'),
};

let state = null;
let preserveFormValues = false;
let applyingDraft = false;
let draftSaveTimer = 0;
let cockpitRootHandle = null;
let cockpitLastPulledIds = [];
let cockpitPulledItems = [];
let accountsListExpanded = false;
let mailListExpanded = false;

const EXTERNAL_SOURCES = {
  cockpit: {
    id: 'cockpit',
    label: 'Cockpit',
    pickerId: 'cockpit-token-exporter',
    expectedDir: '~/.antigravity_cockpit',
    accountsDir: 'codex_accounts',
    includeAllFiles: true,
    defaultStatus: '请选择 Cockpit 目录：~/.antigravity_cockpit。',
  },
  cliproxy: {
    id: 'cliproxy',
    label: 'CliProxy',
    apiOnly: true,
    includeAllFiles: true,
    defaultStatus: '填写 CliProxy CPA 地址和管理密钥后，点击“拉取账号”。例如：http://127.0.0.1:8317。',
  },
};

const DRAFT_FIELDS = [
  'sessionJsonInput',
  'externalSource',
  'cliproxyBaseUrl',
  'cliproxyManagementKey',
  'mailProvider',
  'mailImport',
  'microsoftLocalBaseUrl',
  'autoCheckEnabled',
  'autoCheckMinutes',
  'minTokenValidityMinutes',
  'mailMaxRetries',
  'mailRetryDelayMs',
  'closeTabsAfterRun',
  'activateTabs',
  'directLiveValidateAccessToken',
  'luckmailApiKey',
  'luckmailBaseUrl',
  'cloudflareTempBaseUrl',
  'cloudflareTempAdminAuth',
  'cloudflareTempCustomAuth',
  'cloudflareTempLookupMode',
  'cloudflareTempReceiveMailbox',
];

async function send(type, payload = {}) {
  const response = await chrome.runtime.sendMessage({ type, ...payload });
  if (!response?.ok) {
    throw new Error(response?.error || '操作失败');
  }
  return response;
}

async function refresh(options = {}) {
  state = await send('GET_STATE');
  render(options);
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatDate(value) {
  if (!value) return '';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return '';
  return date.toLocaleString();
}

function statusLabel(status) {
  const map = {
    valid: '有效',
    expired: '已过期',
    missing: '缺 token',
    pending: '待处理',
    refreshing: '刷新中',
    checking: '检查中',
    failed: '失败',
    ready: '可用',
    unchecked: '未实测',
  };
  return map[status] || status || '未知';
}

function mailProviderLabel(provider) {
  const map = {
    microsoft: 'Outlook/Hotmail 账号池',
    'luckmail-api': 'LuckMail 查信',
    'cloudflare-temp-email': 'Cloudflare Temp Email 查信',
  };
  return map[provider] || provider || 'Outlook/Hotmail 账号池';
}

function mailImportPlaceholder(provider) {
  const map = {
    microsoft: '每行一个：email----password----clientId----refreshToken',
    'luckmail-api': '每行一个：email----token\n也兼容：email----token----apiKey----https://mails.luckyous.com',
    'cloudflare-temp-email': '只用全局查信配置可留空；也可每行一个：注册邮箱 或 *',
  };
  return map[provider] || map.microsoft;
}

function renderMailImportMode() {
  if (!dom.mailProvider || !dom.mailImport) return;
  const provider = dom.mailProvider.value;
  const usesGlobalLookupOnly = provider === 'cloudflare-temp-email' || provider === 'luckmail-api';
  dom.mailImport.placeholder = mailImportPlaceholder(provider);
  if (dom.microsoftConfig) {
    dom.microsoftConfig.classList.toggle('hidden', provider !== 'microsoft');
  }
  if (dom.mailImportControls) {
    dom.mailImportControls.classList.toggle('hidden', usesGlobalLookupOnly);
  }
  if (dom.luckmailConfig) {
    dom.luckmailConfig.classList.toggle('hidden', provider !== 'luckmail-api');
  }
  if (dom.cloudflareTempConfig) {
    dom.cloudflareTempConfig.classList.toggle('hidden', provider !== 'cloudflare-temp-email');
  }
}

function renderExternalSourceMode({ resetStatus = false } = {}) {
  const config = getExternalSourceConfig();
  const apiOnly = Boolean(config.apiOnly);
  if (dom.cliproxyConfig) {
    dom.cliproxyConfig.classList.toggle('hidden', config.id !== 'cliproxy');
  }
  if (dom.btnSelectCockpitDir) {
    dom.btnSelectCockpitDir.classList.toggle('hidden', apiOnly);
  }
  if (dom.btnPullCockpit401) {
    dom.btnPullCockpit401.textContent = apiOnly ? '拉取账号' : '扫描文件';
  }
  if (dom.btnRefreshCockpit) {
    dom.btnRefreshCockpit.textContent = apiOnly ? '刷新并导入 CliProxy' : '刷新并写回';
  }
  if (resetStatus) {
    setCockpitStatus(config.defaultStatus);
  }
}

function readDraftField(name) {
  const element = dom[name];
  if (!element) return undefined;
  if (element.type === 'checkbox') return Boolean(element.checked);
  return element.value;
}

function writeDraftField(name, value) {
  const element = dom[name];
  if (!element || value === undefined || value === null) return;
  if (element.type === 'checkbox') {
    element.checked = Boolean(value);
  } else {
    element.value = String(value);
  }
}

function collectPopupDraft() {
  return DRAFT_FIELDS.reduce((draft, name) => {
    const value = readDraftField(name);
    if (value !== undefined) draft[name] = value;
    return draft;
  }, {});
}

async function savePopupDraftNow() {
  if (applyingDraft) return;
  await chrome.storage.local.set({ [POPUP_DRAFT_KEY]: collectPopupDraft() });
}

function savePopupDraftSoon() {
  if (applyingDraft) return;
  preserveFormValues = true;
  clearTimeout(draftSaveTimer);
  draftSaveTimer = setTimeout(() => {
    savePopupDraftNow().catch(() => {});
  }, 120);
}

async function restorePopupDraft() {
  const result = await chrome.storage.local.get(POPUP_DRAFT_KEY);
  const draft = result?.[POPUP_DRAFT_KEY];
  if (!draft || typeof draft !== 'object') return;
  applyingDraft = true;
  try {
    DRAFT_FIELDS.forEach((name) => writeDraftField(name, draft[name]));
    preserveFormValues = true;
    renderMailImportMode();
    renderExternalSourceMode({ resetStatus: true });
  } finally {
    applyingDraft = false;
  }
}

function bindPopupDraftAutosave() {
  DRAFT_FIELDS.forEach((name) => {
    const element = dom[name];
    if (!element) return;
    element.addEventListener('input', savePopupDraftSoon);
    element.addEventListener('change', () => {
      preserveFormValues = true;
      renderMailImportMode();
      if (name === 'externalSource') {
        resetExternalSelectionStatus();
        renderExternalSourceMode({ resetStatus: true });
        renderAccounts(state?.accounts || []);
      }
      savePopupDraftNow().catch(() => {});
    });
  });
}

function renderSettings(settings = {}) {
  if (dom.mailProvider) {
    const provider = settings.mailProvider
      || (settings.cloudflareTempEmailBaseUrl ? 'cloudflare-temp-email' : 'microsoft');
    if (dom.mailProvider.querySelector(`option[value="${provider}"]`)) {
      dom.mailProvider.value = provider;
    }
  }
  dom.autoCheckEnabled.checked = Boolean(settings.autoCheckEnabled);
  dom.autoCheckMinutes.value = settings.autoCheckMinutes ?? 60;
  dom.minTokenValidityMinutes.value = settings.minTokenValidityMinutes ?? 20;
  dom.mailMaxRetries.value = settings.mailMaxRetries ?? 18;
  dom.mailRetryDelayMs.value = settings.mailRetryDelayMs ?? 5000;
  dom.closeTabsAfterRun.checked = settings.closeTabsAfterRun !== false;
  dom.activateTabs.checked = settings.activateTabs !== false;
  dom.directLiveValidateAccessToken.checked = settings.directLiveValidateAccessToken !== false;
  dom.microsoftLocalBaseUrl.value = settings.microsoftLocalBaseUrl || 'http://127.0.0.1:17373';
  dom.luckmailApiKey.value = settings.luckmailApiKey || '';
  dom.luckmailBaseUrl.value = settings.luckmailBaseUrl || 'https://mails.luckyous.com';
  dom.cloudflareTempBaseUrl.value = settings.cloudflareTempEmailBaseUrl || '';
  dom.cloudflareTempAdminAuth.value = settings.cloudflareTempEmailAdminAuth || '';
  dom.cloudflareTempCustomAuth.value = settings.cloudflareTempEmailCustomAuth || '';
  dom.cloudflareTempLookupMode.value = settings.cloudflareTempEmailLookupMode || 'receive-mailbox';
  dom.cloudflareTempReceiveMailbox.value = settings.cloudflareTempEmailReceiveMailbox || '';
}

function collectSettingsInput() {
  return {
    autoCheckEnabled: dom.autoCheckEnabled.checked,
    autoCheckMinutes: dom.autoCheckMinutes.value,
    minTokenValidityMinutes: dom.minTokenValidityMinutes.value,
    mailMaxRetries: dom.mailMaxRetries.value,
    mailRetryDelayMs: dom.mailRetryDelayMs.value,
    mailProvider: dom.mailProvider.value,
    closeTabsAfterRun: dom.closeTabsAfterRun.checked,
    activateTabs: dom.activateTabs.checked,
    directLiveValidateAccessToken: dom.directLiveValidateAccessToken.checked,
    microsoftLocalBaseUrl: dom.microsoftLocalBaseUrl.value,
    luckmailApiKey: dom.luckmailApiKey.value,
    luckmailBaseUrl: dom.luckmailBaseUrl.value,
    cloudflareTempEmailBaseUrl: dom.cloudflareTempBaseUrl.value,
    cloudflareTempEmailAdminAuth: dom.cloudflareTempAdminAuth.value,
    cloudflareTempEmailCustomAuth: dom.cloudflareTempCustomAuth.value,
    cloudflareTempEmailLookupMode: dom.cloudflareTempLookupMode.value,
    cloudflareTempEmailReceiveMailbox: dom.cloudflareTempReceiveMailbox.value,
  };
}

function setCockpitStatus(message) {
  if (dom.cockpitStatus) dom.cockpitStatus.textContent = message;
}

function getExternalSourceConfig() {
  return EXTERNAL_SOURCES[dom.externalSource?.value] || EXTERNAL_SOURCES.cockpit;
}

function getExternalSourceLabel() {
  return getExternalSourceConfig().label;
}

function resetExternalSelectionStatus() {
  cockpitRootHandle = null;
  cockpitLastPulledIds = [];
  cockpitPulledItems = [];
  renderCockpitList([]);
  setCockpitStatus(getExternalSourceConfig().defaultStatus);
}

function renderCockpitList(items = cockpitPulledItems) {
  cockpitPulledItems = Array.isArray(items) ? items : [];
  if (!dom.cockpitList) return;
  if (!cockpitPulledItems.length) {
    dom.cockpitList.innerHTML = '';
    return;
  }
  dom.cockpitList.innerHTML = cockpitPulledItems.map((item) => `
    <label class="cockpit-item">
      <input type="checkbox" data-cockpit-account-id="${escapeHtml(item.accountId)}" ${item.checked === false ? '' : 'checked'}>
      <span>
        <span class="cockpit-item-title">${escapeHtml(item.fileName)}</span>
        <span class="cockpit-item-meta">
          ${escapeHtml(item.email)}
          ${item.errorCode ? `<br><span class="cockpit-item-error">${escapeHtml(item.errorCode)}</span>` : ''}
        </span>
      </span>
    </label>
  `).join('');
}

function getSelectedCockpitIds() {
  const selected = Array.from(dom.cockpitList?.querySelectorAll('input[data-cockpit-account-id]:checked') || [])
    .map((input) => input.getAttribute('data-cockpit-account-id'))
    .filter(Boolean);
  return selected.length || cockpitPulledItems.length ? selected : cockpitLastPulledIds.slice();
}

function accountSourceLabel(account = {}) {
  return account.integrationLabel || (account.integrationSource === 'cliproxy' ? 'CliProxy' : 'Cockpit');
}

function normalizeEmailValue(value) {
  const email = String(value || '').trim().toLowerCase();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ? email : '';
}

function decodeBase64UrlJson(segment = '') {
  const normalized = String(segment || '').replace(/-/g, '+').replace(/_/g, '/');
  if (!normalized) return null;
  const padded = normalized + '='.repeat((4 - (normalized.length % 4)) % 4);
  try {
    return JSON.parse(decodeURIComponent(Array.from(atob(padded), (char) => `%${char.charCodeAt(0).toString(16).padStart(2, '0')}`).join('')));
  } catch {
    try { return JSON.parse(atob(padded)); } catch { return null; }
  }
}

function parseJwtPayload(token = '') {
  const parts = String(token || '').split('.');
  return parts.length >= 2 ? decodeBase64UrlJson(parts[1]) : null;
}

function getOpenAiTokenInfo(accessToken = '') {
  const payload = parseJwtPayload(accessToken) || {};
  const auth = payload['https://api.openai.com/auth'] || {};
  const profile = payload['https://api.openai.com/profile'] || {};
  return {
    accountId: auth.chatgpt_account_id || '',
    userId: auth.chatgpt_user_id || auth.user_id || '',
    planType: auth.chatgpt_plan_type || '',
    email: normalizeEmailValue(profile.email || payload.email),
  };
}

function normalizeCliProxyBaseUrl(value = '') {
  const raw = String(value || '').trim() || 'http://127.0.0.1:8317';
  const withProtocol = /^[a-z][a-z0-9+.-]*:\/\//i.test(raw) ? raw : `http://${raw}`;
  let url;
  try {
    url = new URL(withProtocol);
  } catch {
    throw new Error('CliProxy CPA 地址不正确。示例：http://127.0.0.1:8317');
  }
  url.pathname = '';
  url.search = '';
  url.hash = '';
  return url.toString().replace(/\/$/, '');
}

function getCliProxyApiConfig() {
  const baseUrl = normalizeCliProxyBaseUrl(dom.cliproxyBaseUrl?.value || '');
  const managementKey = String(dom.cliproxyManagementKey?.value || '').trim();
  if (!managementKey) {
    throw new Error('请填写 CliProxy 管理密钥。');
  }
  return { baseUrl, managementKey };
}

async function cliProxyFetch(path, options = {}) {
  const { baseUrl, managementKey } = getCliProxyApiConfig();
  const headers = new Headers(options.headers || {});
  headers.set('Authorization', `Bearer ${managementKey}`);
  const response = await fetch(`${baseUrl}${path}`, { ...options, headers });
  if (response.ok) return response;

  let errorText = '';
  try {
    const payload = await response.clone().json();
    errorText = payload?.error || payload?.message || '';
  } catch {
    errorText = await response.text().catch(() => '');
  }
  throw new Error(`CliProxy HTTP ${response.status}${errorText ? `：${errorText}` : ''}`);
}

async function cliProxyFetchJson(path, options = {}) {
  const response = await cliProxyFetch(path, options);
  return response.json();
}

async function listCliProxyAuthFiles() {
  const payload = await cliProxyFetchJson('/v0/management/auth-files');
  return Array.isArray(payload?.files) ? payload.files : [];
}

async function downloadCliProxyAuthFile(name) {
  const safeName = String(name || '').trim();
  if (!safeName || !safeName.endsWith('.json') || safeName.includes('/') || safeName.includes('\\')) {
    throw new Error(`无效 CliProxy 文件名：${safeName || '(空)'}`);
  }
  const response = await cliProxyFetch(`/v0/management/auth-files/download?name=${encodeURIComponent(safeName)}`);
  return JSON.parse(await response.text());
}

async function uploadCliProxyAuthFile(name, payload) {
  const safeName = String(name || '').trim();
  if (!safeName || !safeName.endsWith('.json') || safeName.includes('/') || safeName.includes('\\')) {
    throw new Error(`无效 CliProxy 文件名：${safeName || '(空)'}`);
  }
  await cliProxyFetch(`/v0/management/auth-files?name=${encodeURIComponent(safeName)}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: `${JSON.stringify(payload, null, 2)}\n`,
  });
}

async function deleteCliProxyAuthFile(name) {
  const safeName = String(name || '').trim();
  if (!safeName || !safeName.endsWith('.json') || safeName.includes('/') || safeName.includes('\\')) {
    throw new Error(`无效 CliProxy 文件名：${safeName || '(空)'}`);
  }
  await cliProxyFetch(`/v0/management/auth-files?name=${encodeURIComponent(safeName)}`, {
    method: 'DELETE',
  });
}

function isCliProxyCodexEntry(entry = {}) {
  const name = String(entry.name || '').trim();
  const type = String(entry.type || entry.provider || '').trim().toLowerCase();
  return name.endsWith('.json') && (!type || type === 'codex');
}

function getCliProxyRefreshReason(record = {}, entry = {}) {
  const cockpitReason = getCockpitRefreshReason(record);
  if (cockpitReason) return cockpitReason;
  const text = [
    entry.status,
    entry.status_message,
    entry.error_code,
    entry.error_message,
  ].map((item) => String(item || '')).join(' ').toLowerCase();
  if (/(^|\D)401(\D|$)|token_invalidated|token_expired|unauthorized/.test(text)) {
    return entry.status_message || entry.status || '401/token_invalidated';
  }
  return '可刷新';
}

function cliProxyRecordToImport(record, fileName, entry = {}) {
  const parsed = cockpitRecordToImport(record, fileName);
  parsed.sourceFile = `CliProxy API:${fileName}`;
  parsed.cockpitRelativePath = fileName;
  parsed.cockpitAccountId = record.id || entry.id || fileName;
  parsed.cockpitErrorCode = getCliProxyRefreshReason(record, entry);
  parsed.integrationSource = 'cliproxy';
  parsed.integrationLabel = 'CliProxy';
  return parsed;
}

async function ensureCockpitPermission() {
  if (!cockpitRootHandle) {
    throw new Error(`请先选择 ${getExternalSourceLabel()} 目录：${getExternalSourceConfig().expectedDir}。`);
  }
  if (!cockpitRootHandle.queryPermission || !cockpitRootHandle.requestPermission) return;
  const options = { mode: 'readwrite' };
  if (await cockpitRootHandle.queryPermission(options) === 'granted') return;
  if (await cockpitRootHandle.requestPermission(options) !== 'granted') {
    throw new Error(`没有 ${getExternalSourceLabel()} 目录读写权限。`);
  }
}

async function getCockpitAccountsDirHandle() {
  await ensureCockpitPermission();
  const config = getExternalSourceConfig();
  if (cockpitRootHandle.name === config.accountsDir) return cockpitRootHandle;
  return cockpitRootHandle.getDirectoryHandle(config.accountsDir);
}

function cockpitRelativePath(fileName) {
  const config = getExternalSourceConfig();
  if (config.apiOnly) return fileName;
  return cockpitRootHandle?.name === config.accountsDir ? fileName : `${config.accountsDir}/${fileName}`;
}

async function selectCockpitDirectory() {
  if (getExternalSourceConfig().apiOnly) {
    throw new Error('CliProxy 不使用本机目录，请填写 CPA 地址和管理密钥后拉取。');
  }
  if (!window.showDirectoryPicker) {
    throw new Error('当前浏览器不支持目录读写授权，请用新版 Chrome 打开扩展侧边栏。');
  }
  const config = getExternalSourceConfig();
  const handle = await window.showDirectoryPicker({
    id: config.pickerId,
    mode: 'readwrite',
  });
  cockpitRootHandle = handle;
  await getCockpitAccountsDirHandle();
  cockpitLastPulledIds = [];
  cockpitPulledItems = [];
  renderCockpitList([]);
  setCockpitStatus(`已选择 ${config.label} 目录：${handle.name}。`);
}

async function readJsonFile(fileHandle) {
  const file = await fileHandle.getFile();
  return JSON.parse(await file.text());
}

async function readCockpitAccountIndex() {
  await ensureCockpitPermission();
  const config = getExternalSourceConfig();
  if (cockpitRootHandle.name === config.accountsDir) {
    throw new Error(`请重新选择 ${config.expectedDir} 根目录，不要直接选择 ${config.accountsDir} 子目录。插件需要读取 codex_accounts.json 来同步已删除账号。`);
  }
  const summaryHandle = await cockpitRootHandle.getFileHandle('codex_accounts.json');
  const summary = await readJsonFile(summaryHandle);
  const accounts = Array.isArray(summary.accounts) ? summary.accounts : [];
  return accounts
    .map((item) => ({
      id: String(item.id || '').trim(),
      email: normalizeEmailValue(item.email),
      planType: String(item.plan_type || item.chatgpt_plan_type || '').trim(),
    }))
    .filter((item) => item.id);
}

function getCockpitRefreshReason(record = {}) {
  const quotaError = record.quota_error || {};
  const quota = record.quota || {};
  const text = [
    quotaError.code,
    quotaError.message,
    record.error_code,
    record.error_message,
    quota.status,
    quota.status_code,
  ].map((item) => String(item || '')).join(' ').toLowerCase();
  if (/(^|\D)401(\D|$)|token_invalidated|token_expired|unauthorized|撤销|刷新失败/.test(text)) {
    return quotaError.code || record.error_code || '401/token_invalidated';
  }
  const recordEmail = normalizeEmailValue(record.email || record.name);
  const accessToken = String(record.tokens?.access_token || record.access_token || '').trim();
  const tokenEmail = getOpenAiTokenInfo(accessToken).email;
  if (recordEmail && tokenEmail && recordEmail !== tokenEmail) {
    return `token 邮箱不匹配：${tokenEmail}`;
  }
  const expiresAt = Date.parse(record.expired || record.expiresAt || record.expires || '');
  if (expiresAt && expiresAt <= Date.now()) {
    return '本地 expired 已过期';
  }
  if (!accessToken) {
    return '缺少 accessToken';
  }
  return '';
}

function cockpitRecordToImport(record, fileName) {
  const tokens = record.tokens || {};
  const accessToken = String(tokens.access_token || record.access_token || '').trim();
  const tokenInfo = getOpenAiTokenInfo(accessToken);
  const sourceConfig = getExternalSourceConfig();
  return {
    email: normalizeEmailValue(record.email || record.name) || tokenInfo.email,
    accessToken,
    idToken: tokens.id_token || record.id_token || '',
    refreshToken: tokens.refresh_token || record.refresh_token || '',
    sessionToken: tokens.session_token || record.session_token || '',
    accountId: record.account_id || record.chatgpt_account_id || tokenInfo.accountId || '',
    userId: record.user_id || tokenInfo.userId || '',
    planType: record.plan_type || record.chatgpt_plan_type || tokenInfo.planType || '',
    sourceFile: cockpitRelativePath(fileName),
    cockpitAccountId: record.id || fileName,
    cockpitRelativePath: cockpitRelativePath(fileName),
    cockpitErrorCode: record.quota_error?.code || record.error_code || '',
    integrationSource: sourceConfig.id,
    integrationLabel: sourceConfig.label,
  };
}

async function pullCockpit401Accounts() {
  if (getExternalSourceConfig().id === 'cliproxy') {
    await pullCliProxyAccounts();
    return;
  }
  if (!cockpitRootHandle) await selectCockpitDirectory();
  const config = getExternalSourceConfig();
  const indexedAccounts = await readCockpitAccountIndex();
  const indexedIds = indexedAccounts.map((item) => item.id);
  const accountsDir = await getCockpitAccountsDirHandle();
  const records = [];
  const pendingItems = [];
  let scanned = 0;
  for (const indexedAccount of indexedAccounts) {
    const name = `${indexedAccount.id}.json`;
    scanned += 1;
    let handle = null;
    let record = null;
    try {
      handle = await accountsDir.getFileHandle(name);
      record = await readJsonFile(handle);
    } catch {
      continue;
    }
    const refreshReason = getCockpitRefreshReason(record);
    if (!refreshReason && !config.includeAllFiles) continue;
    const parsed = cockpitRecordToImport(record, name);
    parsed.email = parsed.email || indexedAccount.email;
    parsed.planType = parsed.planType || indexedAccount.planType;
    if (parsed.email && parsed.accessToken) {
      parsed.cockpitErrorCode = refreshReason || '可刷新';
      records.push(parsed);
      pendingItems.push({
        email: parsed.email,
        fileName: name,
        relativePath: parsed.cockpitRelativePath,
        cockpitAccountId: parsed.cockpitAccountId,
        errorCode: refreshReason || '可刷新',
      });
    }
  }
  if (!records.length) {
    await send('IMPORT_COCKPIT_401_ACCOUNTS', {
      accounts: [],
      syncSource: 'cockpit',
      syncAccountIds: indexedIds,
      syncComplete: true,
    });
    await refresh();
    setCockpitStatus(`已按 Cockpit 索引扫描 ${scanned} 个账号，没有找到可导入账号；已清理插件里的失效缓存。`);
    cockpitLastPulledIds = [];
    cockpitPulledItems = [];
    renderCockpitList([]);
    return;
  }
  const response = await send('IMPORT_COCKPIT_401_ACCOUNTS', {
    accounts: records,
    syncSource: 'cockpit',
    syncAccountIds: indexedIds,
    syncComplete: true,
  });
  const accountsByEmail = new Map((response.accounts || []).map((account) => [account.email, account]));
  cockpitPulledItems = pendingItems
    .map((item) => ({ ...item, accountId: accountsByEmail.get(item.email)?.id || '' }))
    .filter((item) => item.accountId);
  cockpitLastPulledIds = cockpitPulledItems.map((item) => item.accountId);
  renderCockpitList(cockpitPulledItems);
  setCockpitStatus(`已按 Cockpit 索引扫描 ${scanned} 个账号，拉取账号 ${records.length} 个。`);
  await checkExternalAccounts(cockpitLastPulledIds, config.label);
}

async function pullCliProxyAccounts() {
  const files = await listCliProxyAuthFiles();
  const sourceIds = files
    .filter((entry) => isCliProxyCodexEntry(entry))
    .map((entry) => String(entry.id || entry.name || '').trim())
    .filter(Boolean);
  const records = [];
  const pendingItems = [];
  let scanned = 0;
  for (const entry of files) {
    if (!isCliProxyCodexEntry(entry)) continue;
    const name = String(entry.name || '').trim();
    scanned += 1;
    let record = null;
    try {
      record = await downloadCliProxyAuthFile(name);
    } catch (error) {
      pendingItems.push({
        email: String(entry.email || ''),
        fileName: name,
        relativePath: name,
        cockpitAccountId: entry.id || name,
        errorCode: `下载失败：${error?.message || error}`,
      });
      continue;
    }
    const parsed = cliProxyRecordToImport(record, name, entry);
    if (parsed.email && parsed.accessToken) {
      records.push(parsed);
      pendingItems.push({
        email: parsed.email,
        fileName: name,
        relativePath: name,
        cockpitAccountId: parsed.cockpitAccountId,
        errorCode: parsed.cockpitErrorCode || '可刷新',
      });
    }
  }
  if (!records.length) {
    await send('IMPORT_COCKPIT_401_ACCOUNTS', {
      accounts: [],
      syncSource: 'cliproxy',
      syncAccountIds: sourceIds,
      syncComplete: true,
    });
    await refresh();
    setCockpitStatus(`已从 CliProxy API 扫描 ${scanned} 个 Codex JSON，没有找到可刷新账号。`);
    cockpitLastPulledIds = [];
    cockpitPulledItems = [];
    renderCockpitList([]);
    return;
  }
  const response = await send('IMPORT_COCKPIT_401_ACCOUNTS', {
    accounts: records,
    syncSource: 'cliproxy',
    syncAccountIds: sourceIds,
    syncComplete: true,
  });
  const accountsByEmail = new Map((response.accounts || []).map((account) => [account.email, account]));
  cockpitPulledItems = pendingItems
    .map((item) => ({ ...item, accountId: accountsByEmail.get(normalizeEmailValue(item.email))?.id || '' }))
    .filter((item) => item.accountId);
  cockpitLastPulledIds = cockpitPulledItems.map((item) => item.accountId);
  renderCockpitList(cockpitPulledItems);
  setCockpitStatus(`已通过 CliProxy API 拉取 ${records.length} 个账号，可选择后刷新并导入。`);
  await checkExternalAccounts(cockpitLastPulledIds, 'CliProxy');
}

async function getFileHandleByRelativePath(relativePath) {
  await ensureCockpitPermission();
  const config = getExternalSourceConfig();
  let parts = String(relativePath || '').split('/').filter(Boolean);
  if (cockpitRootHandle.name === config.accountsDir && parts[0] === config.accountsDir) {
    parts = parts.slice(1);
  }
  if (!parts.length) throw new Error(`无效 Cockpit 文件路径：${relativePath}`);
  let dir = cockpitRootHandle;
  for (const part of parts.slice(0, -1)) {
    dir = await dir.getDirectoryHandle(part);
  }
  return dir.getFileHandle(parts[parts.length - 1]);
}

async function removeCockpitFileByRelativePath(relativePath) {
  await ensureCockpitPermission();
  const config = getExternalSourceConfig();
  let parts = String(relativePath || '').split('/').filter(Boolean);
  if (cockpitRootHandle.name === config.accountsDir && parts[0] === config.accountsDir) {
    parts = parts.slice(1);
  }
  if (!parts.length) throw new Error(`无效 Cockpit 文件路径：${relativePath}`);
  let dir = cockpitRootHandle;
  for (const part of parts.slice(0, -1)) {
    dir = await dir.getDirectoryHandle(part);
  }
  await dir.removeEntry(parts[parts.length - 1]);
}

function buildUpdatedCockpitRecord(existing, account) {
  const tokenInfo = getOpenAiTokenInfo(account.accessToken);
  const updated = { ...existing };
  updated.email = account.email || existing.email;
  updated.account_id = account.accountId || tokenInfo.accountId || existing.account_id || '';
  updated.user_id = account.userId || tokenInfo.userId || existing.user_id || '';
  updated.plan_type = account.planType || tokenInfo.planType || existing.plan_type || '';
  updated.auth_mode = existing.auth_mode || 'oauth';
  updated.token_source_mode = existing.token_source_mode || 'managed';
  updated.token_updated_at = Math.floor(Date.now() / 1000);
  updated.tokens = {
    ...(existing.tokens || {}),
    access_token: account.accessToken,
    id_token: account.idToken || existing.tokens?.id_token || '',
    refresh_token: account.refreshToken || existing.tokens?.refresh_token || '',
    session_token: account.sessionToken || existing.tokens?.session_token || '',
  };
  delete updated.quota_error;
  return updated;
}

async function writeJsonFile(fileHandle, payload) {
  const writable = await fileHandle.createWritable();
  await writable.write(`${JSON.stringify(payload, null, 2)}\n`);
  await writable.close();
}

async function updateCockpitSummary(accountsByCockpitId) {
  try {
    const summaryHandle = await getFileHandleByRelativePath('codex_accounts.json');
    const summary = await readJsonFile(summaryHandle);
    if (!Array.isArray(summary.accounts)) return;
    const seen = new Set();
    summary.accounts = summary.accounts.map((item) => {
      if (item.id) seen.add(item.id);
      const account = accountsByCockpitId.get(item.id);
      if (!account) return item;
      return {
        ...item,
        email: account.email || item.email,
        plan_type: account.planType || item.plan_type,
      };
    });
    const now = Math.floor(Date.now() / 1000);
    for (const [id, account] of accountsByCockpitId.entries()) {
      if (!id || seen.has(id)) continue;
      summary.accounts.push({
        id,
        email: account.email || '',
        plan_type: account.planType || '',
        created_at: now,
        last_used: now,
      });
    }
    await writeJsonFile(summaryHandle, summary);
  } catch {
    // Cockpit summary is optional; individual account JSON is the source of truth.
  }
}

async function removeCockpitSummaryAccounts(accountIds = []) {
  const ids = new Set(accountIds.map(String).filter(Boolean));
  if (!ids.size) return;
  try {
    const summaryHandle = await getFileHandleByRelativePath('codex_accounts.json');
    const summary = await readJsonFile(summaryHandle);
    if (!Array.isArray(summary.accounts)) return;
    const before = summary.accounts.length;
    summary.accounts = summary.accounts.filter((item) => !ids.has(String(item.id || '').trim()));
    if (summary.accounts.length !== before) {
      await writeJsonFile(summaryHandle, summary);
    }
  } catch {
    // Some Cockpit versions can work from individual JSON files without a summary update.
  }
}

function isAccountForExternalSource(account = {}, sourceId = getExternalSourceConfig().id) {
  if (sourceId === 'cliproxy') return account.integrationSource === 'cliproxy';
  return !account.integrationSource || account.integrationSource === 'cockpit';
}

function getVisibleExternalAccounts(accounts = []) {
  const sourceId = getExternalSourceConfig().id;
  return (Array.isArray(accounts) ? accounts : []).filter((account) => (
    account.source === 'cockpit' && isAccountForExternalSource(account, sourceId)
  ));
}

async function checkExternalAccounts(accountIds = [], label = getExternalSourceLabel()) {
  const ids = accountIds.map(String).filter(Boolean);
  if (!ids.length) return;
  const refreshLabel = getExternalSourceConfig().id === 'cliproxy' ? '刷新并导入 CliProxy' : '刷新并写回';
  setCockpitStatus(`已拉取 ${label} 账号 ${ids.length} 个，正在自动验证是否过期...`);
  const response = await send('CHECK_TOKENS', { accountIds: ids });
  state = response;
  const accountsById = new Map((response.accounts || []).map((account) => [account.id, account]));
  cockpitPulledItems = cockpitPulledItems.map((item) => {
    const account = accountsById.get(item.accountId);
    if (!account) return item;
    const needsRefresh = account.status !== 'valid';
    return {
      ...item,
      checked: needsRefresh,
      errorCode: `${statusLabel(account.status)}${account.error ? `：${account.error}` : ''}`,
    };
  });
  renderCockpitList(cockpitPulledItems);
  setCockpitStatus(`已拉取 ${label} 账号 ${ids.length} 个，并完成过期验证。需要刷新时勾选账号后点“${refreshLabel}”。`);
}

async function checkVisibleExternalAccounts() {
  const ids = getVisibleExternalAccounts(state?.accounts || []).map((account) => account.id);
  await checkExternalAccounts(ids, getExternalSourceLabel());
}

function getRecordAccessToken(record = {}) {
  return String(record.tokens?.access_token || record.access_token || record.accessToken || '').trim();
}

async function writeBackCliProxyAccounts(accountIds = [], minRefreshStartedAt = 0) {
  await refresh();
  const ids = new Set(accountIds.map(String));
  const selectedAccounts = (state?.accounts || []).filter((account) => (
    ids.has(account.id) && isAccountForExternalSource(account, 'cliproxy')
  ));
  const accounts = selectedAccounts.filter((account) => {
    const refreshedAt = Date.parse(account.lastRefreshAt || '');
    return account.status === 'valid'
      && account.accessToken
      && account.cockpitRelativePath
      && (!minRefreshStartedAt || (refreshedAt && refreshedAt >= minRefreshStartedAt));
  });
  if (!accounts.length) {
    const details = selectedAccounts
      .map((account) => {
        const refreshedAt = Date.parse(account.lastRefreshAt || '');
        const stale = minRefreshStartedAt && (!refreshedAt || refreshedAt < minRefreshStartedAt);
        const reason = account.error || (stale ? '本次没有刷新成功' : statusLabel(account.status));
        return `${account.email}：${reason}`;
      })
      .filter(Boolean)
      .slice(0, 5)
      .join('；');
    throw new Error(`没有本次刷新成功且可导入 CliProxy 的账号。${details ? `失败原因：${details}` : '请查看日志里的刷新失败原因。'}`);
  }

  const built = await send('BUILD_CPA_AUTH_JSONS', { accountIds: accounts.map((account) => account.id) });
  const authJsonById = new Map((built.items || []).map((item) => [item.id, item.json]));
  let uploaded = 0;
  for (const account of accounts) {
    const fileName = account.cockpitRelativePath;
    const existing = await downloadCliProxyAuthFile(fileName);
    const tokenEmail = getOpenAiTokenInfo(account.accessToken).email;
    if (!tokenEmail || tokenEmail !== normalizeEmailValue(account.email)) {
      throw new Error(`${account.email} 刷新后的 accessToken 邮箱为 ${tokenEmail || '无法解析'}，已阻止导入 CliProxy。`);
    }
    const oldAccessToken = getRecordAccessToken(existing);
    if (oldAccessToken && oldAccessToken === account.accessToken) {
      throw new Error(`${account.email} 刷新后 accessToken 未变化，已阻止把旧 CliProxy JSON 导入。`);
    }
    const cpaJson = authJsonById.get(account.id);
    if (!cpaJson?.access_token) {
      throw new Error(`${account.email} 未能生成 CPA JSON。`);
    }
    await uploadCliProxyAuthFile(fileName, cpaJson);
    uploaded += 1;
  }
  setCockpitStatus(`已通过 CliProxy API 导入刷新后的 JSON：${uploaded} 个，已确认 accessToken 不是原文件旧值。`);
}

async function deleteExternalDeactivatedAccounts(deletedAccounts = []) {
  const sourceId = getExternalSourceConfig().id;
  const accounts = (Array.isArray(deletedAccounts) ? deletedAccounts : [])
    .filter((account) => account?.source === 'cockpit' && isAccountForExternalSource(account, sourceId));
  if (!accounts.length) return [];

  const deletedIds = [];
  const cockpitSummaryIds = [];
  const failures = [];
  for (const account of accounts) {
    deletedIds.push(account.id);
    cockpitPulledItems = cockpitPulledItems.filter((item) => item.accountId !== account.id);
    cockpitLastPulledIds = cockpitLastPulledIds.filter((id) => id !== account.id);
    try {
      if (sourceId === 'cliproxy') {
        await deleteCliProxyAuthFile(account.cockpitRelativePath || account.sourceFile || `${account.cockpitAccountId}.json`);
      } else {
        await removeCockpitFileByRelativePath(account.cockpitRelativePath || account.sourceFile || `${account.cockpitAccountId}.json`);
        if (account.cockpitAccountId) cockpitSummaryIds.push(account.cockpitAccountId);
      }
      setCockpitStatus(`${account.email} 已停用，已从 ${getExternalSourceLabel()} 删除。`);
    } catch (error) {
      failures.push(`${account.email}：${error?.message || error}`);
    }
  }
  if (sourceId !== 'cliproxy') {
    await removeCockpitSummaryAccounts(cockpitSummaryIds);
  }
  renderCockpitList(cockpitPulledItems);
  if (failures.length) {
    setCockpitStatus(`停用账号已从插件账号池删除，但从 ${getExternalSourceLabel()} 删除失败 ${failures.length} 个：${failures.slice(0, 2).join('；')}`);
  }
  return deletedIds;
}

async function writeBackCockpitAccounts(accountIds = [], minRefreshStartedAt = 0) {
  const config = getExternalSourceConfig();
  if (config.id === 'cliproxy') {
    await writeBackCliProxyAccounts(accountIds, minRefreshStartedAt);
    return;
  }
  await ensureCockpitPermission();
  await refresh();
  const ids = new Set(accountIds.map(String));
  const selectedAccounts = (state?.accounts || []).filter((account) => (
    ids.has(account.id) && account.source === 'cockpit' && isAccountForExternalSource(account, 'cockpit')
  ));
  const accounts = (state?.accounts || []).filter((account) => (
    ids.has(account.id)
    && account.source === 'cockpit'
    && isAccountForExternalSource(account, 'cockpit')
    && account.status === 'valid'
    && account.accessToken
    && account.cockpitRelativePath
    && (!minRefreshStartedAt || Date.parse(account.lastRefreshAt || '') >= minRefreshStartedAt)
  ));
  if (!accounts.length) {
    const details = selectedAccounts
      .map((account) => {
        const refreshedAt = Date.parse(account.lastRefreshAt || '');
        const stale = minRefreshStartedAt && (!refreshedAt || refreshedAt < minRefreshStartedAt);
        const reason = account.error || (stale ? '本次没有刷新成功' : statusLabel(account.status));
        return `${account.email}：${reason}`;
      })
      .filter(Boolean)
      .slice(0, 5)
      .join('；');
    throw new Error(`没有本次刷新成功且可回写 Cockpit 的账号。${details ? `失败原因：${details}` : '请查看日志里的刷新失败原因。'}`);
  }
  const byCockpitId = new Map();
  let written = 0;
  for (const account of accounts) {
    const fileHandle = await getFileHandleByRelativePath(account.cockpitRelativePath);
    const existing = await readJsonFile(fileHandle);
    const tokenEmail = getOpenAiTokenInfo(account.accessToken).email;
    if (!tokenEmail || tokenEmail !== normalizeEmailValue(account.email)) {
      throw new Error(`${account.email} 刷新后的 accessToken 邮箱为 ${tokenEmail || '无法解析'}，已阻止写回 Cockpit。`);
    }
    const oldAccessToken = String(existing.tokens?.access_token || existing.access_token || '').trim();
    if (oldAccessToken && oldAccessToken === account.accessToken) {
      throw new Error(`${account.email} 刷新后 accessToken 未变化，已阻止把旧 Cockpit JSON 写回。`);
    }
    await writeJsonFile(fileHandle, buildUpdatedCockpitRecord(existing, account));
    if (account.cockpitAccountId) byCockpitId.set(account.cockpitAccountId, account);
    written += 1;
  }
  await updateCockpitSummary(byCockpitId);
  setCockpitStatus(`已刷新并回写 Cockpit JSON：${written} 个，已确认 accessToken 不是原文件旧值。Cockpit 页面未变化时点一下 Cockpit 的刷新按钮。`);
}

async function refreshCockpit401AndWriteBack() {
  if (!cockpitLastPulledIds.length) {
    await pullCockpit401Accounts();
  }
  const label = getExternalSourceLabel();
  const selectedIds = getSelectedCockpitIds();
  if (!selectedIds.length) {
    throw new Error(`没有可刷新的 ${label} 账号。`);
  }
  await saveSettingsBeforeAction();
  setCockpitStatus(`开始刷新已选择的 ${label} 账号 ${selectedIds.length} 个...`);
  const refreshStartedAt = Date.now();
  const batchResult = await send('START_COCKPIT_BATCH', { accountIds: selectedIds });
  const deletedIds = await deleteExternalDeactivatedAccounts(batchResult.deletedAccounts || []);
  const remainingIds = selectedIds.filter((id) => !deletedIds.includes(id));
  if (!remainingIds.length) {
    await refresh();
    if (!dom.cockpitStatus?.textContent.includes('删除失败')) {
      setCockpitStatus(`已处理 ${label} 账号：停用账号已从插件和本地工具删除。`);
    }
    return;
  }
  await writeBackCockpitAccounts(remainingIds, refreshStartedAt);
}


function renderAccounts(accounts = []) {
  const visibleAccounts = getVisibleExternalAccounts(accounts);
  const label = getExternalSourceLabel();
  renderPoolToggle(dom.btnToggleAccountsList, visibleAccounts.length, accountsListExpanded);
  if (!visibleAccounts.length) {
    const action = getExternalSourceConfig().id === 'cliproxy' ? '填写 CPA 地址和管理密钥后点击“拉取账号”。' : '选择 Cockpit 目录后点击“扫描文件”。';
    dom.accountsList.innerHTML = `<div class="muted">暂无 ${escapeHtml(label)} ChatGPT 账号。${escapeHtml(action)}</div>`;
    return;
  }
  const renderedAccounts = accountsListExpanded ? visibleAccounts : visibleAccounts.slice(0, 1);
  dom.accountsList.innerHTML = renderedAccounts.map((account) => {
    const expires = account.expiresAt ? formatDate(account.expiresAt) : '';
    return `
      <div class="row">
        <div>
          <div class="mono">${escapeHtml(account.email)}</div>
          <div class="muted">
            邮箱：${escapeHtml(account.mailAccountEmail || account.email)}
            ${account.sessionToken ? '<br>session_token：已保存' : ''}
            ${account.sourceFile ? `<br>来源：${escapeHtml(account.sourceFile)}` : ''}
            ${account.source === 'cockpit' ? `<br>${escapeHtml(accountSourceLabel(account))}：${account.cockpitRefreshRequired ? '待刷新' : '已同步'}` : ''}
            ${account.lastLiveValidatedAt ? `<br>实测：${escapeHtml(formatDate(account.lastLiveValidatedAt))}` : ''}
          </div>
        </div>
        <div class="status-${escapeHtml(account.status)}">${escapeHtml(statusLabel(account.status))}</div>
        <div class="muted">${expires ? `过期：${escapeHtml(expires)}` : '未导出'}${account.error ? `<br>${escapeHtml(account.error)}` : ''}</div>
        <button data-delete-account="${escapeHtml(account.id)}">删除</button>
      </div>
    `;
  }).join('');
}

function renderMailAccounts(mailAccounts = []) {
  const visibleMailAccounts = Array.isArray(mailAccounts) ? mailAccounts : [];
  renderPoolToggle(dom.btnToggleMailList, visibleMailAccounts.length, mailListExpanded);
  if (!visibleMailAccounts.length) {
    dom.mailList.innerHTML = '<div class="muted">暂无邮箱账号。</div>';
    return;
  }
  const renderedMailAccounts = mailListExpanded ? visibleMailAccounts : visibleMailAccounts.slice(0, 1);
  dom.mailList.innerHTML = renderedMailAccounts.map((account) => `
    <div class="row">
      <div>
        <div class="mono">${escapeHtml(account.email)}</div>
        <div class="muted">${escapeHtml(mailProviderLabel(account.provider))}</div>
      </div>
      <div class="status-${escapeHtml(account.status)}">${escapeHtml(statusLabel(account.status))}</div>
      <div class="muted">${escapeHtml(mailAccountConfigLabel(account))}${account.error ? `<br>${escapeHtml(account.error)}` : ''}</div>
      <div>
        <button data-test-mail="${escapeHtml(account.id)}">测试</button>
        <button data-delete-mail="${escapeHtml(account.id)}">删除</button>
      </div>
    </div>
  `).join('');
}

function renderPoolToggle(button, count, expanded) {
  if (!button) return;
  button.classList.toggle('hidden', count <= 1);
  button.textContent = expanded ? '收起' : `展开（${Math.max(0, count - 1)}）`;
  button.setAttribute('aria-expanded', String(Boolean(expanded)));
}

function mailAccountConfigLabel(account = {}) {
  if (account.provider === 'luckmail-api') {
    return `${account.token ? 'token 已保存' : '缺 token'}；${account.apiKey ? 'API Key 已保存' : '缺 API Key'}`;
  }
  if (account.provider === 'cloudflare-temp-email') {
    return `${account.baseUrl ? 'Temp API 已保存' : '缺 Temp API'}；${account.adminAuth ? 'Admin Auth 已保存' : '缺 Admin Auth'}`;
  }
  return `${account.clientId ? 'clientId 已保存' : '缺 clientId'}；${account.refreshToken ? 'refreshToken 已保存' : '缺 refreshToken'}`;
}

function renderLogs(logs = []) {
  const autoStickToBottom = !dom.logs
    || dom.logs.scrollHeight - dom.logs.scrollTop - dom.logs.clientHeight < 24;
  dom.logs.textContent = logs.map((item) => {
    const time = new Date(item.at).toLocaleTimeString();
    return `[${time}] [${item.level}] ${item.message}`;
  }).join('\n');
  if (autoStickToBottom) {
    dom.logs.scrollTop = dom.logs.scrollHeight;
  }
}

function render(options = {}) {
  const runtime = state?.runtime || {};
  dom.runState.textContent = runtime.running
    ? (runtime.cancelRequested ? '停止中' : '运行中')
    : '空闲';
  if (dom.btnStopBatch) {
    dom.btnStopBatch.classList.toggle('hidden', !runtime.running);
    dom.btnStopBatch.disabled = Boolean(runtime.cancelRequested);
    dom.btnStopBatch.textContent = runtime.cancelRequested ? '停止中' : '停止';
  }
  if (!preserveFormValues || options.forceForms) {
    renderSettings(state?.settings || {});
  }
  renderMailImportMode();
  renderAccounts(state?.accounts || []);
  renderMailAccounts(state?.mailAccounts || []);
  renderLogs(runtime.logs || []);
}

async function runAction(action) {
  try {
    await action();
    await refresh();
  } catch (error) {
    alert(error?.message || String(error || '操作失败'));
  }
}

dom.btnConvertSessionJson.addEventListener('click', () => runAction(async () => {
  await send('CONVERT_SESSION_JSON_TO_CPA', {
    content: dom.sessionJsonInput.value,
  });
  dom.sessionJsonInput.value = '';
  await savePopupDraftNow();
}));

dom.btnImportMail.addEventListener('click', () => runAction(async () => {
  await send('IMPORT_MAIL_ACCOUNTS', {
    text: dom.mailImport.value,
    provider: dom.mailProvider.value,
    settings: collectSettingsInput(),
  });
  dom.mailImport.value = '';
  dom.mailImportFile.value = '';
  await savePopupDraftNow();
}));

dom.btnSelectCockpitDir.addEventListener('click', () => runAction(selectCockpitDirectory));
dom.btnPullCockpit401.addEventListener('click', () => runAction(async () => {
  await saveSettingsBeforeAction();
  await pullCockpit401Accounts();
}));
dom.btnRefreshCockpit.addEventListener('click', () => runAction(refreshCockpit401AndWriteBack));
dom.btnStopBatch.addEventListener('click', () => runAction(async () => {
  await send('STOP_BATCH');
}));

async function readTextFileToTextarea(fileInput, textarea) {
  const file = fileInput.files?.[0];
  if (!file) return;
  textarea.value = await file.text();
}

dom.mailImportFile.addEventListener('change', () => runAction(async () => {
  await readTextFileToTextarea(dom.mailImportFile, dom.mailImport);
  await savePopupDraftNow();
}));

dom.mailProvider.addEventListener('change', renderMailImportMode);
dom.externalSource.addEventListener('change', () => runAction(async () => {
  await saveSettingsBeforeAction();
  await checkVisibleExternalAccounts();
}));

dom.btnSaveSettings.addEventListener('click', () => runAction(async () => {
  await send('SAVE_SETTINGS', {
    settings: collectSettingsInput(),
  });
  await savePopupDraftNow();
}));

async function saveSettingsBeforeAction() {
  await send('SAVE_SETTINGS', { settings: collectSettingsInput() });
  await savePopupDraftNow();
}

dom.btnClearLogs.addEventListener('click', () => runAction(() => send('CLEAR_LOGS')));
dom.btnToggleAccountsList.addEventListener('click', () => {
  accountsListExpanded = !accountsListExpanded;
  renderAccounts(state?.accounts || []);
});
dom.btnToggleMailList.addEventListener('click', () => {
  mailListExpanded = !mailListExpanded;
  renderMailAccounts(state?.mailAccounts || []);
});

document.addEventListener('click', (event) => {
  const deleteAccount = event.target?.getAttribute?.('data-delete-account');
  const deleteMail = event.target?.getAttribute?.('data-delete-mail');
  const testMail = event.target?.getAttribute?.('data-test-mail');
  if (deleteAccount) {
    runAction(() => send('DELETE_CHATGPT_ACCOUNT', { id: deleteAccount }));
  } else if (deleteMail) {
    runAction(() => send('DELETE_MAIL_ACCOUNT', { id: deleteMail }));
  } else if (testMail) {
    runAction(() => send('TEST_MAIL', { id: testMail }));
  }
});

(async function initPopup() {
  await refresh({ forceForms: true });
  await restorePopupDraft();
  bindPopupDraftAutosave();
  renderMailImportMode();
  renderExternalSourceMode({ resetStatus: true });
  renderAccounts(state?.accounts || []);
  setInterval(() => {
    refresh().catch(() => {});
  }, 2000);
})();
