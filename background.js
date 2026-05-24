importScripts(
  'microsoft-email.js',
  'luckmail-utils.js',
  'cloudflare-temp-email-utils.js'
);

const STORE_KEYS = {
  accounts: 'chatgptAccounts',
  mailAccounts: 'mailAccounts',
  settings: 'settings',
};

const DEFAULT_SETTINGS = {
  autoCheckEnabled: false,
  autoCheckMinutes: 60,
  minTokenValidityMinutes: 20,
  mailMaxRetries: 18,
  mailRetryDelayMs: 5000,
  mailProvider: 'microsoft',
  closeTabsAfterRun: true,
  activateTabs: true,
  directLiveValidateAccessToken: true,
  microsoftLocalBaseUrl: 'http://127.0.0.1:17373',
  luckmailApiKey: '',
  luckmailBaseUrl: 'https://mails.luckyous.com',
  cloudflareTempEmailBaseUrl: '',
  cloudflareTempEmailAdminAuth: '',
  cloudflareTempEmailCustomAuth: '',
  cloudflareTempEmailLookupMode: 'receive-mailbox',
  cloudflareTempEmailReceiveMailbox: '',
};

const OPENAI_COOKIE_ORIGINS = [
  'https://chatgpt.com',
  'https://chat.openai.com',
  'https://auth.openai.com',
  'https://auth0.openai.com',
  'https://accounts.openai.com',
  'https://openai.com',
];
const OPENAI_COOKIE_DOMAINS = [
  'chatgpt.com',
  '.chatgpt.com',
  'chat.openai.com',
  '.chat.openai.com',
  'auth.openai.com',
  '.auth.openai.com',
  'auth0.openai.com',
  '.auth0.openai.com',
  'accounts.openai.com',
  '.accounts.openai.com',
  'openai.com',
  '.openai.com',
];

const AUTH_HOST_RE = /(^|\.)((chatgpt\.com)|(chat\.openai\.com)|(auth\.openai\.com)|(auth0\.openai\.com)|(accounts\.openai\.com))$/i;
const SESSION_HOST_RE = /(^|\.)((chatgpt\.com)|(chat\.openai\.com))$/i;
const AUTO_CHECK_ALARM = 'token-exporter-auto-check';
const MAIL_PROVIDER_MICROSOFT = 'microsoft';
const MAIL_PROVIDER_LUCKMAIL = 'luckmail-api';
const MAIL_PROVIDER_CLOUDFLARE_TEMP_EMAIL = 'cloudflare-temp-email';
const DIRECT_VALIDATION_TIMEOUT_MS = 15000;
const DIRECT_VALIDATION_ENDPOINTS = [
  { name: 'me', url: 'https://chatgpt.com/backend-api/me' },
  { name: 'models', url: 'https://chatgpt.com/backend-api/models' },
];

const runtime = {
  running: false,
  logs: [],
  currentAccountId: '',
  lastRunAt: '',
};

configureSidePanel();

chrome.runtime.onInstalled.addListener(async () => {
  await ensureDefaults();
  await configureSidePanel();
  await syncAutoCheckAlarm();
});

chrome.runtime.onStartup.addListener(async () => {
  await ensureDefaults();
  await configureSidePanel();
  await syncAutoCheckAlarm();
});

chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm?.name === AUTO_CHECK_ALARM) {
    checkTokenStatusOnly().catch((error) => {
      addLog(`自动检查失败：${error?.message || error}`, 'error');
    });
  }
});

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  (async () => handleMessage(message || {}))()
    .then((result) => sendResponse({ ok: true, ...(result || {}) }))
    .catch((error) => sendResponse({ ok: false, error: error?.message || String(error || '') }));
  return true;
});

async function handleMessage(message) {
  switch (message.type) {
    case 'GET_STATE':
      return getUiState();
    case 'SAVE_SETTINGS':
      await saveSettings(message.settings || {});
      return getUiState();
    case 'IMPORT_COCKPIT_401_ACCOUNTS':
      await importCockpit401Accounts(message.accounts || [], {
        syncSource: message.syncSource || '',
        syncAccountIds: message.syncAccountIds || [],
        syncComplete: Boolean(message.syncComplete),
      });
      return getUiState();
    case 'BUILD_CPA_AUTH_JSONS':
      return buildCpaAuthJsonsForAccounts(message.accountIds || []);
    case 'CONVERT_SESSION_JSON_TO_CPA':
      await convertSessionJsonToCpa(message.content || '');
      return getUiState();
    case 'IMPORT_MAIL_ACCOUNTS':
      await importMailAccounts(message.text || '', message.provider || '', message.settings || {});
      return getUiState();
    case 'DELETE_CHATGPT_ACCOUNT':
      await deleteChatGptAccount(message.id || '');
      return getUiState();
    case 'DELETE_MAIL_ACCOUNT':
      await deleteMailAccount(message.id || '');
      return getUiState();
    case 'START_COCKPIT_BATCH':
      return runBatch({
        force: true,
        reason: 'cockpit',
        accountIds: Array.isArray(message.accountIds) ? message.accountIds : [],
      });
    case 'CHECK_TOKENS':
      await checkTokenStatusOnly({ accountIds: message.accountIds || [] });
      return getUiState();
    case 'TEST_MAIL':
      return testMailAccount(message.id || '');
    case 'CLEAR_LOGS':
      runtime.logs = [];
      return getUiState();
    default:
      throw new Error(`未知消息：${message.type || ''}`);
  }
}

async function ensureDefaults() {
  const current = await chrome.storage.local.get([STORE_KEYS.accounts, STORE_KEYS.mailAccounts, STORE_KEYS.settings]);
  const updates = {};
  if (!Array.isArray(current[STORE_KEYS.accounts])) updates[STORE_KEYS.accounts] = [];
  if (!Array.isArray(current[STORE_KEYS.mailAccounts])) updates[STORE_KEYS.mailAccounts] = [];
  updates[STORE_KEYS.settings] = { ...DEFAULT_SETTINGS, ...(current[STORE_KEYS.settings] || {}) };
  await chrome.storage.local.set(updates);
}

async function configureSidePanel() {
  if (!chrome.sidePanel?.setPanelBehavior) return;
  try {
    await chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });
  } catch (error) {
    addLog(`配置侧边栏失败：${error?.message || error}`, 'warn');
  }
}

async function getData() {
  await ensureDefaults();
  const data = await chrome.storage.local.get([STORE_KEYS.accounts, STORE_KEYS.mailAccounts, STORE_KEYS.settings]);
  return {
    accounts: normalizeChatGptAccounts(data[STORE_KEYS.accounts]),
    mailAccounts: normalizeMailAccounts(data[STORE_KEYS.mailAccounts]),
    settings: { ...DEFAULT_SETTINGS, ...(data[STORE_KEYS.settings] || {}) },
  };
}

async function setAccounts(accounts) {
  await chrome.storage.local.set({ [STORE_KEYS.accounts]: normalizeChatGptAccounts(accounts) });
}

async function setMailAccounts(mailAccounts) {
  await chrome.storage.local.set({ [STORE_KEYS.mailAccounts]: normalizeMailAccounts(mailAccounts) });
}

async function saveSettings(input) {
  const previousSettings = (await getData()).settings;
  const settings = {
    ...DEFAULT_SETTINGS,
    ...previousSettings,
    autoCheckEnabled: Boolean(input.autoCheckEnabled),
    autoCheckMinutes: clampInt(input.autoCheckMinutes, 15, 1440, DEFAULT_SETTINGS.autoCheckMinutes),
    minTokenValidityMinutes: clampInt(input.minTokenValidityMinutes, 1, 1440, DEFAULT_SETTINGS.minTokenValidityMinutes),
    mailMaxRetries: clampInt(input.mailMaxRetries, 1, 120, DEFAULT_SETTINGS.mailMaxRetries),
    mailRetryDelayMs: clampInt(input.mailRetryDelayMs, 1000, 60000, DEFAULT_SETTINGS.mailRetryDelayMs),
    mailProvider: normalizeMailProviderId(input.mailProvider) || DEFAULT_SETTINGS.mailProvider,
    closeTabsAfterRun: Boolean(input.closeTabsAfterRun),
    activateTabs: Boolean(input.activateTabs),
    directLiveValidateAccessToken: input.directLiveValidateAccessToken !== false,
    microsoftLocalBaseUrl: normalizeMicrosoftLocalBaseUrl(input.microsoftLocalBaseUrl),
    luckmailApiKey: normalizeString(input.luckmailApiKey),
    luckmailBaseUrl: LuckMailUtils.normalizeLuckmailBaseUrl(input.luckmailBaseUrl || DEFAULT_SETTINGS.luckmailBaseUrl),
    cloudflareTempEmailBaseUrl: CloudflareTempEmailUtils.normalizeCloudflareTempEmailBaseUrl(input.cloudflareTempEmailBaseUrl),
    cloudflareTempEmailAdminAuth: normalizeString(input.cloudflareTempEmailAdminAuth),
    cloudflareTempEmailCustomAuth: normalizeString(input.cloudflareTempEmailCustomAuth),
    cloudflareTempEmailLookupMode: normalizeCloudflareLookupMode(input.cloudflareTempEmailLookupMode),
    cloudflareTempEmailReceiveMailbox: normalizeEmail(input.cloudflareTempEmailReceiveMailbox),
  };
  delete settings.microsoftMailMode;
  await chrome.storage.local.set({ [STORE_KEYS.settings]: settings });
  await syncAutoCheckAlarm();
}

async function syncAutoCheckAlarm() {
  const { settings } = await getData();
  await chrome.alarms.clear(AUTO_CHECK_ALARM);
  if (settings.autoCheckEnabled) {
    chrome.alarms.create(AUTO_CHECK_ALARM, {
      periodInMinutes: settings.autoCheckMinutes,
      delayInMinutes: 1,
    });
  }
}

async function getUiState() {
  const data = await getData();
  return {
    ...data,
    runtime: {
      running: runtime.running,
      currentAccountId: runtime.currentAccountId,
      lastRunAt: runtime.lastRunAt,
      logs: runtime.logs.slice(-300),
    },
  };
}

function addLog(message, level = 'info') {
  const entry = {
    at: new Date().toISOString(),
    level,
    message: String(message || ''),
  };
  runtime.logs.push(entry);
  if (runtime.logs.length > 500) runtime.logs.splice(0, runtime.logs.length - 500);
  console[level === 'error' ? 'error' : level === 'warn' ? 'warn' : 'log']('[TokenExporter]', entry.message);
}

function normalizeString(value) {
  return String(value || '').trim();
}

function normalizeEmail(value) {
  const email = normalizeString(value).toLowerCase();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ? email : '';
}

function normalizeEmailWithoutPlusAlias(value) {
  const email = normalizeEmail(value);
  if (!email) return '';
  const [localPart, domain] = email.split('@');
  const plusIndex = localPart.indexOf('+');
  if (plusIndex <= 0) return email;
  return `${localPart.slice(0, plusIndex)}@${domain}`;
}

function getEmailMatchCandidates(value) {
  const email = normalizeEmail(value);
  if (!email) return [];
  return [...new Set([email, normalizeEmailWithoutPlusAlias(email)].filter(Boolean))];
}

function emailsMatch(left, right) {
  const leftCandidates = new Set(getEmailMatchCandidates(left));
  return getEmailMatchCandidates(right).some((candidate) => leftCandidates.has(candidate));
}

function getPreferredMailboxEmail(value) {
  const candidates = getEmailMatchCandidates(value);
  return candidates[1] || candidates[0] || '';
}

function normalizeMailAccountEmail(value) {
  const text = normalizeString(value).toLowerCase();
  if (text === '*') return '*';
  return normalizeEmail(text);
}

function normalizeMailProviderId(value) {
  const normalized = normalizeString(value).toLowerCase().replace(/_/g, '-');
  switch (normalized) {
    case 'hotmail':
    case 'hotmail-api':
    case 'outlook':
    case 'microsoft':
    case 'ms':
    case 'ms-graph':
    case 'msgraph':
      return MAIL_PROVIDER_MICROSOFT;
    case 'luckmail':
    case 'luckmail-api':
    case 'luckyous':
      return MAIL_PROVIDER_LUCKMAIL;
    case 'cloudflare':
    case 'cloudflare-temp':
    case 'cloudflare-temp-email':
    case 'cf-temp':
    case 'cf-temp-email':
      return MAIL_PROVIDER_CLOUDFLARE_TEMP_EMAIL;
    default:
      return '';
  }
}

function isMailProviderToken(value) {
  return Boolean(normalizeMailProviderId(value));
}

function mailAccountKey(account = {}) {
  return `${normalizeMailProviderId(account.provider) || MAIL_PROVIDER_MICROSOFT}:${normalizeMailAccountEmail(account.email)}`;
}

function normalizeCloudflareLookupMode(value = '') {
  const normalized = normalizeString(value).toLowerCase().replace(/_/g, '-');
  return ['registration-email', 'original-recipient', 'recipient'].includes(normalized)
    ? 'registration-email'
    : 'receive-mailbox';
}

function normalizeMicrosoftLocalBaseUrl(value = '') {
  const raw = normalizeString(value) || DEFAULT_SETTINGS.microsoftLocalBaseUrl;
  try {
    const parsed = new URL(raw);
    if (!['http:', 'https:'].includes(parsed.protocol)) {
      return DEFAULT_SETTINGS.microsoftLocalBaseUrl;
    }
    if (['/messages', '/code'].includes(parsed.pathname)) {
      parsed.pathname = '';
      parsed.search = '';
      parsed.hash = '';
    }
    return parsed.toString().replace(/\/$/, '');
  } catch {
    return DEFAULT_SETTINGS.microsoftLocalBaseUrl;
  }
}

function buildMicrosoftLocalEndpoint(settings = {}, path = '/messages') {
  return new URL(path, `${normalizeMicrosoftLocalBaseUrl(settings.microsoftLocalBaseUrl)}/`).toString();
}

function makeId(prefix = 'id') {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function clampInt(value, min, max, fallback) {
  const numeric = Math.floor(Number(value));
  if (!Number.isFinite(numeric)) return fallback;
  return Math.max(min, Math.min(max, numeric));
}

function splitImportLine(line) {
  const text = normalizeString(line);
  if (!text || text.startsWith('#')) return [];
  if (text.includes('----')) return text.split('----').map((item) => item.trim());
  if (text.includes('\t')) return text.split('\t').map((item) => item.trim());
  if (text.includes('|')) return text.split('|').map((item) => item.trim());
  return text.split(',').map((item) => item.trim());
}

function normalizeChatGptAccount(record = {}) {
  const email = normalizeEmail(record.email);
  if (!email) return null;
  const accessToken = normalizeString(record.accessToken || record.access_token);
  const sessionToken = normalizeString(record.sessionToken || record.session_token);
  const idToken = normalizeString(record.idToken || record.id_token);
  const refreshToken = normalizeString(record.refreshToken || record.refresh_token);
  const expiresAt = Number(record.expiresAt)
    || Number(record.expires_at)
    || Date.parse(record.expired || record.expires || '')
    || getTokenExpiryMs(accessToken);
  return {
    id: normalizeString(record.id) || makeId('cgpt'),
    email,
    password: String(record.password || ''),
    mailAccountEmail: normalizeEmail(record.mailAccountEmail || record.email),
    accessToken,
    sessionToken,
    idToken,
    refreshToken,
    userId: normalizeString(record.userId || record.user_id || record.chatgptUserId || record.chatgpt_user_id),
    accountId: normalizeString(record.accountId || record.account_id || record.chatgptAccountId || record.chatgpt_account_id),
    planType: normalizeString(record.planType || record.plan_type || record.chatgptPlanType || record.chatgpt_plan_type),
    sourceFile: normalizeString(record.sourceFile),
    source: normalizeString(record.source) || (sessionToken ? 'token-json' : 'manual'),
    integrationSource: normalizeString(record.integrationSource || record.integration_source),
    integrationLabel: normalizeString(record.integrationLabel || record.integration_label),
    cockpitAccountId: normalizeString(record.cockpitAccountId || record.cockpit_account_id),
    cockpitRelativePath: normalizeString(record.cockpitRelativePath || record.cockpit_relative_path),
    cockpitImportedAt: normalizeString(record.cockpitImportedAt || record.cockpit_imported_at),
    cockpitErrorCode: normalizeString(record.cockpitErrorCode || record.cockpit_error_code),
    cockpitRefreshRequired: Boolean(record.cockpitRefreshRequired || record.cockpit_refresh_required),
    expiresAt,
    status: normalizeString(record.status) || 'pending',
    lastCheckedAt: normalizeString(record.lastCheckedAt),
    lastRefreshAt: normalizeString(record.lastRefreshAt),
    lastLiveValidatedAt: normalizeString(record.lastLiveValidatedAt),
    lastExportAt: normalizeString(record.lastExportAt),
    error: normalizeString(record.error),
  };
}

function normalizeChatGptAccounts(accounts = []) {
  const seen = new Set();
  const results = [];
  for (const item of Array.isArray(accounts) ? accounts : []) {
    const account = normalizeChatGptAccount(item);
    if (!account || seen.has(account.email)) continue;
    seen.add(account.email);
    results.push(account);
  }
  return results;
}

function normalizeMailAccount(record = {}) {
  const email = normalizeMailAccountEmail(record.email);
  if (!email) return null;
  const provider = normalizeMailProviderId(record.provider) || MAIL_PROVIDER_MICROSOFT;
  return {
    id: normalizeString(record.id) || makeId('mail'),
    provider,
    email,
    password: String(record.password || ''),
    clientId: normalizeString(record.clientId),
    refreshToken: normalizeString(record.refreshToken),
    token: normalizeString(record.token),
    apiKey: normalizeString(record.apiKey),
    baseUrl: normalizeString(record.baseUrl),
    emailType: normalizeString(record.emailType),
    projectCode: normalizeString(record.projectCode),
    adminAuth: normalizeString(record.adminAuth),
    customAuth: normalizeString(record.customAuth),
    adminEmail: normalizeEmail(record.adminEmail),
    adminPassword: String(record.adminPassword || ''),
    domain: normalizeString(record.domain),
    lookupMode: normalizeCloudflareLookupMode(record.lookupMode),
    receiveMailbox: normalizeMailAccountEmail(record.receiveMailbox),
    status: normalizeString(record.status) || 'ready',
    lastCheckedAt: normalizeString(record.lastCheckedAt),
    error: normalizeString(record.error),
  };
}

function normalizeMailAccounts(accounts = []) {
  const seen = new Set();
  const results = [];
  for (const item of Array.isArray(accounts) ? accounts : []) {
    const account = normalizeMailAccount(item);
    const key = account ? mailAccountKey(account) : '';
    if (!account || seen.has(key)) continue;
    seen.add(key);
    results.push(account);
  }
  return results;
}

function extractEmailFromAccessToken(accessToken = '') {
  const payload = parseJwtPayload(accessToken) || {};
  const profile = payload && typeof payload === 'object'
    ? (payload['https://api.openai.com/profile'] || {})
    : {};
  return normalizeEmail(profile.email || payload.email || payload.username || payload.sub);
}

function accountMatchesIntegrationSource(account = {}, source = '') {
  const normalized = normalizeString(source);
  if (normalized === 'cliproxy') return account.source === 'cockpit' && account.integrationSource === 'cliproxy';
  if (normalized === 'cockpit') return account.source === 'cockpit' && (!account.integrationSource || account.integrationSource === 'cockpit');
  return false;
}

async function importCockpit401Accounts(records = [], options = {}) {
  const { accounts } = await getData();
  const byEmail = new Map(accounts.map((item) => [item.email, item]));
  let imported = 0;
  const errors = [];
  const importedAt = new Date().toISOString();
  const syncSource = normalizeString(options.syncSource);
  const syncAccountIds = new Set((Array.isArray(options.syncAccountIds) ? options.syncAccountIds : [])
    .map((item) => normalizeString(item))
    .filter(Boolean));
  for (const record of Array.isArray(records) ? records : []) {
    try {
      const email = normalizeEmail(record.email);
      const accessToken = normalizeString(record.accessToken || record.access_token);
      if (!email) throw new Error('缺少邮箱');
      if (!accessToken) throw new Error(`${email} 缺少 accessToken`);
      const existing = byEmail.get(email) || {};
      const integrationSource = normalizeString(record.integrationSource || record.integration_source) || 'cockpit';
      const integrationLabel = normalizeString(record.integrationLabel || record.integration_label)
        || (integrationSource === 'cliproxy' ? 'CliProxy' : 'Cockpit');
      byEmail.set(email, normalizeChatGptAccount({
        ...existing,
        email,
        accessToken,
        sessionToken: normalizeString(record.sessionToken || record.session_token) || existing.sessionToken || '',
        idToken: normalizeString(record.idToken || record.id_token) || existing.idToken || '',
        refreshToken: normalizeString(record.refreshToken || record.refresh_token) || existing.refreshToken || '',
        accountId: normalizeString(record.accountId || record.account_id) || existing.accountId || '',
        userId: normalizeString(record.userId || record.user_id) || existing.userId || '',
        planType: normalizeString(record.planType || record.plan_type) || existing.planType || '',
        expiresAt: getTokenExpiryMs(accessToken) || existing.expiresAt || 0,
        source: 'cockpit',
        integrationSource,
        integrationLabel,
        sourceFile: normalizeString(record.sourceFile) || existing.sourceFile || '',
        cockpitAccountId: normalizeString(record.cockpitAccountId || record.cockpit_account_id || record.id) || existing.cockpitAccountId || '',
        cockpitRelativePath: normalizeString(record.cockpitRelativePath || record.cockpit_relative_path) || existing.cockpitRelativePath || '',
        cockpitErrorCode: normalizeString(record.cockpitErrorCode || record.cockpit_error_code) || existing.cockpitErrorCode || '',
        cockpitImportedAt: importedAt,
        cockpitRefreshRequired: true,
        mailAccountEmail: existing.mailAccountEmail || email,
        status: 'expired',
        error: `${integrationLabel} 账号等待重新登录刷新。`,
      }));
      imported += 1;
    } catch (error) {
      errors.push(error?.message || String(error || ''));
    }
  }
  let nextAccounts = Array.from(byEmail.values());
  let removed = 0;
  if (syncSource && options.syncComplete) {
    const before = nextAccounts.length;
    nextAccounts = nextAccounts.filter((account) => (
      !accountMatchesIntegrationSource(account, syncSource)
      || syncAccountIds.has(account.cockpitAccountId)
      || syncAccountIds.has(account.id)
    ));
    removed = before - nextAccounts.length;
  }
  await setAccounts(nextAccounts);
  if (imported) addLog(`已从外部工具拉取可刷新账号 ${imported} 个。`, 'info');
  if (removed) addLog(`已同步删除本地缓存中不在外部工具索引里的账号 ${removed} 个。`, 'info');
  if (errors.length) addLog(`外部工具账号导入失败 ${errors.length} 个：${errors.join('；')}`, 'warn');
}

function parseMicrosoftMailImport(email, values = {}) {
  const normalizedEmail = normalizeEmail(email);
  if (!normalizedEmail) {
    throw new Error(`Microsoft 邮箱格式错误：${email || '(空)'}`);
  }
  return {
    provider: MAIL_PROVIDER_MICROSOFT,
    email: normalizedEmail,
    password: normalizeString(values.password),
    clientId: normalizeString(values.clientId),
    refreshToken: normalizeString(values.refreshToken),
  };
}

function parseLuckmailImport(email, values = {}, settings = {}) {
  const normalizedEmail = normalizeEmail(email);
  if (!normalizedEmail) {
    throw new Error(`LuckMail 邮箱格式错误：${email || '(空)'}`);
  }
  const token = normalizeString(values.token);
  if (!token) {
    throw new Error(`LuckMail ${normalizedEmail} 缺少邮箱 token。`);
  }
  return {
    provider: MAIL_PROVIDER_LUCKMAIL,
    email: normalizedEmail,
    token,
    apiKey: normalizeString(values.apiKey) || normalizeString(settings.luckmailApiKey),
    baseUrl: normalizeString(values.baseUrl) || normalizeString(settings.luckmailBaseUrl) || 'https://mails.luckyous.com',
  };
}

function parseCloudflareTempEmailImport(email, values = {}, settings = {}) {
  const normalizedEmail = normalizeMailAccountEmail(email);
  if (!normalizedEmail) {
    throw new Error(`Cloudflare Temp Email 目标邮箱格式错误：${email || '(空)'}`);
  }
  const baseUrl = normalizeString(values.baseUrl) || normalizeString(settings.cloudflareTempEmailBaseUrl);
  if (!baseUrl) {
    throw new Error(`Cloudflare Temp Email ${normalizedEmail} 缺少 Temp API 地址。`);
  }
  return {
    provider: MAIL_PROVIDER_CLOUDFLARE_TEMP_EMAIL,
    email: normalizedEmail,
    baseUrl,
    adminAuth: normalizeString(values.adminAuth) || normalizeString(settings.cloudflareTempEmailAdminAuth),
    customAuth: normalizeString(values.customAuth) || normalizeString(settings.cloudflareTempEmailCustomAuth),
    lookupMode: normalizeCloudflareLookupMode(values.lookupMode || settings.cloudflareTempEmailLookupMode),
    receiveMailbox: normalizeMailAccountEmail(values.receiveMailbox) || normalizeMailAccountEmail(settings.cloudflareTempEmailReceiveMailbox),
  };
}

function parseMailAccountPartsByProvider(provider, parts, settings = {}) {
  const normalizedProvider = normalizeMailProviderId(provider) || MAIL_PROVIDER_MICROSOFT;
  const [email = '', one = '', two = '', three = '', four = '', five = ''] = parts;
  if (normalizedProvider === MAIL_PROVIDER_MICROSOFT) {
    return parseMicrosoftMailImport(email, three
      ? { password: one, clientId: two, refreshToken: three }
      : { password: '', clientId: one, refreshToken: two });
  }
  if (normalizedProvider === MAIL_PROVIDER_LUCKMAIL) {
    return parseLuckmailImport(email, { token: one, apiKey: two, baseUrl: three }, settings);
  }
  if (normalizedProvider === MAIL_PROVIDER_CLOUDFLARE_TEMP_EMAIL) {
    return parseCloudflareTempEmailImport(email, {
      baseUrl: one,
      adminAuth: two,
      customAuth: three,
      lookupMode: four,
      receiveMailbox: five,
    }, settings);
  }
  throw new Error(`不支持的邮箱服务：${provider}`);
}

function parseMailAccountImportLine(line, selectedProvider = '', settings = {}) {
  const parts = splitImportLine(line);
  if (!parts.length) return null;

  const firstProvider = isMailProviderToken(parts[0]) ? normalizeMailProviderId(parts[0]) : '';
  if (firstProvider) {
    return parseMailAccountPartsByProvider(firstProvider, parts.slice(1), settings);
  }

  const secondProvider = isMailProviderToken(parts[1]) ? normalizeMailProviderId(parts[1]) : '';
  if (secondProvider) {
    return parseMailAccountPartsByProvider(secondProvider, [parts[0], ...parts.slice(2)], settings);
  }

  return parseMailAccountPartsByProvider(selectedProvider || MAIL_PROVIDER_MICROSOFT, parts, settings);
}

async function importMailAccounts(text, selectedProvider = '', settingsOverride = {}) {
  const data = await getData();
  const { mailAccounts } = data;
  const settings = { ...data.settings, ...settingsOverride };
  const byKey = new Map(mailAccounts.map((item) => [mailAccountKey(item), item]));
  let imported = 0;
  const errors = [];
  const lines = String(text || '').split(/\r?\n/).filter((line) => {
    const normalized = normalizeString(line);
    return normalized && !normalized.startsWith('#');
  });
  const provider = normalizeMailProviderId(selectedProvider);
  if (!lines.length && [MAIL_PROVIDER_LUCKMAIL, MAIL_PROVIDER_CLOUDFLARE_TEMP_EMAIL].includes(provider)) {
    addLog(`${provider === MAIL_PROVIDER_LUCKMAIL ? 'LuckMail' : 'Cloudflare Temp Email'} 使用全局配置查信，无需导入邮箱池。`, 'info');
    return;
  }
  for (const line of lines) {
    try {
      const parsed = parseMailAccountImportLine(line, selectedProvider, settings);
      if (!parsed) continue;
      const existing = byKey.get(mailAccountKey(parsed)) || {};
      const normalized = normalizeMailAccount({
        ...existing,
        ...parsed,
        password: parsed.password || existing.password || '',
        clientId: parsed.clientId || existing.clientId || '',
        refreshToken: parsed.refreshToken || existing.refreshToken || '',
        token: parsed.token || existing.token || '',
        apiKey: parsed.apiKey || existing.apiKey || '',
        baseUrl: parsed.baseUrl || existing.baseUrl || '',
        emailType: parsed.emailType || existing.emailType || '',
        projectCode: parsed.projectCode || existing.projectCode || '',
        adminAuth: parsed.adminAuth || existing.adminAuth || '',
        customAuth: parsed.customAuth || existing.customAuth || '',
        adminEmail: parsed.adminEmail || existing.adminEmail || '',
        adminPassword: parsed.adminPassword || existing.adminPassword || '',
        domain: parsed.domain || existing.domain || '',
        lookupMode: parsed.lookupMode || existing.lookupMode || '',
        receiveMailbox: parsed.receiveMailbox || existing.receiveMailbox || '',
        status: 'ready',
        error: '',
      });
      byKey.set(mailAccountKey(normalized), normalized);
      imported += 1;
    } catch (error) {
      errors.push(error?.message || String(error || ''));
    }
  }
  await setMailAccounts(Array.from(byKey.values()));
  addLog(`已导入/更新邮箱账号 ${imported} 个。`, 'info');
  if (errors.length) {
    addLog(`邮箱导入有 ${errors.length} 行失败：${errors.join('；')}`, 'warn');
  }
}

async function deleteChatGptAccount(id) {
  const { accounts } = await getData();
  await setAccounts(accounts.filter((item) => item.id !== id));
}

class AccountDeactivatedError extends Error {
  constructor(email, detail = 'account_deactivated') {
    super(`${email || '当前账号'} 已被删除或停用，自动从账号池删除。`);
    this.name = 'AccountDeactivatedError';
    this.code = 'account_deactivated';
    this.detail = detail;
  }
}

function isAccountDeactivatedError(error) {
  return error?.code === 'account_deactivated'
    || /account_deactivated|账号.*(删除|停用|禁用|暂停)|账户.*(删除|停用|禁用|暂停)|已被删除或停用/i.test(String(error?.message || error || ''));
}

async function deleteMailAccount(id) {
  const { mailAccounts } = await getData();
  await setMailAccounts(mailAccounts.filter((item) => item.id !== id));
}

function decodeBase64UrlSegment(segment = '') {
  const normalized = normalizeString(segment).replace(/-/g, '+').replace(/_/g, '/');
  if (!normalized) return '';
  const padded = normalized + '='.repeat((4 - (normalized.length % 4)) % 4);
  try {
    return decodeURIComponent(Array.from(atob(padded), (char) => `%${char.charCodeAt(0).toString(16).padStart(2, '0')}`).join(''));
  } catch {
    try { return atob(padded); } catch { return ''; }
  }
}

function parseJwtPayload(token = '') {
  const parts = normalizeString(token).split('.');
  if (parts.length < 2) return null;
  try {
    return JSON.parse(decodeBase64UrlSegment(parts[1]));
  } catch {
    return null;
  }
}

function getTokenExpiryMs(token = '') {
  const payload = parseJwtPayload(token);
  const exp = Number(payload?.exp);
  return Number.isFinite(exp) && exp > 0 ? exp * 1000 : 0;
}

function isTokenValid(account, minValidityMinutes) {
  const token = normalizeString(account?.accessToken);
  const expiresAt = Number(account?.expiresAt) || getTokenExpiryMs(token);
  return Boolean(token && expiresAt && expiresAt - Date.now() > minValidityMinutes * 60 * 1000);
}

function validationErrorMessage(result = {}) {
  if (result.valid) return '';
  if (result.expiredLocally) return 'JWT exp 已过期。';
  if (result.httpStatus) {
    return `后台接口探活失败：${result.endpoint || 'unknown'} HTTP ${result.httpStatus}。`;
  }
  return result.error || '后台接口探活失败。';
}

function isRefreshableValidationFailure(result = {}) {
  return Boolean(result.expiredLocally || [401, 403].includes(Number(result.httpStatus)));
}

async function fetchWithTimeout(url, options = {}, timeoutMs = DIRECT_VALIDATION_TIMEOUT_MS) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } finally {
    clearTimeout(timer);
  }
}

function summarizeValidationBody(body = '') {
  const text = String(body || '').replace(/\s+/g, ' ').trim();
  if (!text) return '';
  try {
    const json = JSON.parse(text);
    return normalizeString(json.detail || json.error?.message || json.error || json.message || json.code).slice(0, 160);
  } catch {
    return text.slice(0, 160);
  }
}

async function validateAccessTokenDirect(accessToken) {
  const token = normalizeString(accessToken);
  if (!token) {
    return { valid: false, error: 'accessToken 为空。' };
  }
  const expiresAt = getTokenExpiryMs(token);
  if (expiresAt && expiresAt <= Date.now()) {
    return { valid: false, expiredLocally: true, expiresAt };
  }

  const attempts = [];
  for (const endpoint of DIRECT_VALIDATION_ENDPOINTS) {
    try {
      const response = await fetchWithTimeout(endpoint.url, {
        method: 'GET',
        cache: 'no-store',
        credentials: 'omit',
        headers: {
          accept: 'application/json',
          authorization: `Bearer ${token}`,
        },
      });
      const body = await response.text().catch(() => '');
      const attempt = {
        endpoint: endpoint.name,
        httpStatus: response.status,
        ok: response.ok,
        detail: summarizeValidationBody(body),
      };
      attempts.push(attempt);
      if (response.ok) {
        return {
          valid: true,
          endpoint: endpoint.name,
          httpStatus: response.status,
          expiresAt,
          attempts,
        };
      }
    } catch (error) {
      attempts.push({
        endpoint: endpoint.name,
        httpStatus: 0,
        ok: false,
        error: error?.name === 'AbortError' ? '请求超时' : (error?.message || String(error || '请求失败')),
      });
    }
  }

  const decisive = attempts.find((item) => [401, 403].includes(Number(item.httpStatus)))
    || attempts.find((item) => Number(item.httpStatus) > 0)
    || attempts[0]
    || {};
  return {
    valid: false,
    endpoint: decisive.endpoint || '',
    httpStatus: Number(decisive.httpStatus) || 0,
    error: decisive.error || decisive.detail || '后台接口探活未通过。',
    expiresAt,
    attempts,
  };
}

async function validateAndUpdateAccount(account, settings) {
  const now = new Date().toISOString();
  const expiresAt = Number(account.expiresAt) || getTokenExpiryMs(account.accessToken);
  if (settings.directLiveValidateAccessToken && account.accessToken) {
    await updateAccount(account.id, { status: 'checking', lastCheckedAt: now, error: '' });
    const result = await validateAccessTokenDirect(account.accessToken);
    const status = result.valid ? 'valid' : (isRefreshableValidationFailure(result) ? 'expired' : 'failed');
    const integrationPatch = account.source === 'cockpit'
      ? {
          cockpitRefreshRequired: status !== 'valid',
          cockpitErrorCode: status === 'valid' ? '' : validationErrorMessage(result),
        }
      : {};
    await updateAccount(account.id, {
      expiresAt: result.expiresAt || expiresAt,
      status,
      lastCheckedAt: new Date().toISOString(),
      lastLiveValidatedAt: new Date().toISOString(),
      error: validationErrorMessage(result),
      ...integrationPatch,
    });
    addLog(`${account.email}：后台接口探活${result.valid ? '有效' : `不可用（${validationErrorMessage(result)}）`}。`, result.valid ? 'info' : 'warn');
    return result;
  }

  const valid = isTokenValid({ ...account, expiresAt }, settings.minTokenValidityMinutes);
  const localStatus = valid ? 'valid' : (account.accessToken ? 'expired' : 'missing');
  const integrationPatch = account.source === 'cockpit'
    ? {
        cockpitRefreshRequired: localStatus !== 'valid',
        cockpitErrorCode: localStatus === 'valid' ? '' : '本地有效期不足或缺失。',
      }
    : {};
  await updateAccount(account.id, {
    expiresAt,
    status: localStatus,
    lastCheckedAt: now,
    error: valid ? '' : account.error,
    ...integrationPatch,
  });
  return {
    valid,
    expiredLocally: !valid && Boolean(account.accessToken),
    expiresAt,
    error: valid ? '' : '本地有效期不足或缺失。',
  };
}

async function checkTokenStatusOnly(options = {}) {
  const { accounts, settings } = await getData();
  const ids = new Set(Array.isArray(options.accountIds) ? options.accountIds.map(String).filter(Boolean) : []);
  const targetAccounts = ids.size
    ? accounts.filter((account) => ids.has(account.id))
    : accounts;
  for (const account of targetAccounts) {
    await validateAndUpdateAccount(account, settings);
  }
  const scope = ids.size ? `${targetAccounts.length} 个指定账号` : '全部账号';
  addLog(settings.directLiveValidateAccessToken
    ? `已完成${scope}后台 accessToken 实测。`
    : `已完成${scope}本地 token 过期时间检查。`, 'info');
}

async function runBatch(options = {}) {
  if (runtime.running) {
    addLog('已有批量任务正在运行，跳过本次启动。', 'warn');
    return getUiState();
  }
  runtime.running = true;
  runtime.lastRunAt = new Date().toISOString();
  const deletedAccounts = [];
  try {
    const { accounts, settings } = await getData();
    const accountIds = new Set(Array.isArray(options.accountIds) ? options.accountIds.map(String) : []);
    const targetAccounts = accountIds.size
      ? accounts.filter((account) => accountIds.has(account.id))
      : accounts;
    addLog(`开始批量处理 ${targetAccounts.length} 个账号。`, 'info');
    for (const account of targetAccounts) {
      runtime.currentAccountId = account.id;
      if (!options.force && !account.cockpitRefreshRequired) {
        const validation = await validateAndUpdateAccount(account, settings);
        if (validation.valid) {
          addLog(`${account.email} accessToken 当前可用，跳过登录。`, 'info');
          continue;
        }
        if (!isRefreshableValidationFailure(validation)) {
          addLog(`${account.email} accessToken 状态未知，跳过自动登录：${validationErrorMessage(validation)}`, 'warn');
          continue;
        }
        addLog(`${account.email} accessToken 已不可用，开始刷新。`, 'warn');
      }
      const result = await refreshOneAccount(account.id, settings);
      if (result?.deleted) deletedAccounts.push(result.account);
    }
    addLog('批量任务完成。', 'info');
    return { ...(await getUiState()), deletedAccounts };
  } finally {
    runtime.running = false;
    runtime.currentAccountId = '';
  }
}

async function updateAccount(id, patch) {
  const { accounts } = await getData();
  await setAccounts(accounts.map((item) => (item.id === id ? { ...item, ...patch } : item)));
}

async function updateMailAccount(id, patch) {
  const { mailAccounts } = await getData();
  await setMailAccounts(mailAccounts.map((item) => (item.id === id ? { ...item, ...patch } : item)));
}

async function refreshOneAccount(accountId, settings) {
  const { accounts, mailAccounts } = await getData();
  const account = accounts.find((item) => item.id === accountId);
  if (!account) throw new Error(`账号不存在：${accountId}`);

  const loginTab = { id: 0 };
  try {
    await updateAccount(account.id, { status: 'refreshing', error: '', lastCheckedAt: new Date().toISOString() });
    await resetOpenAiAuthState(0, `${account.email}：刷新前清理 ChatGPT 登录态。`);

    const mailAccount = await resolveMailAccountForChatGptAccount(account, mailAccounts, settings);
    if (!mailAccount) {
      throw new Error(`未找到 ${account.email} 对应的查信配置。请检查邮箱服务选择和配置。`);
    }

    addLog(`${account.email}：打开 ChatGPT 首页并点击登录，走网页邮箱验证码登录。`, 'info');
    loginTab.id = await createChatGptLoginTab(settings);
    const requestedAt = Date.now();
    const tokenState = await driveLoginAndReadToken(loginTab, account, mailAccount, requestedAt, settings);
    assertAccessTokenMatchesAccount(account, tokenState.accessToken);
    await updateAccountWithTokenState(account.id, tokenState);
    addLog(`${account.email}：已获取新的 accessToken。`, 'info');
  } catch (error) {
    if (isAccountDeactivatedError(error)) {
      const deletedAccount = { ...account };
      await deleteChatGptAccount(account.id);
      addLog(`${account.email}：账号已被删除或停用，已自动从 ChatGPT 账号池删除。`, 'warn');
      return { deleted: true, account: deletedAccount };
    }
    await updateAccount(account.id, {
      status: 'failed',
      error: error?.message || String(error || ''),
      lastCheckedAt: new Date().toISOString(),
    });
    addLog(`${account.email}：刷新失败：${error?.message || error}`, 'error');
  } finally {
    if (loginTab.id) {
      await chrome.tabs.update(loginTab.id, { url: 'about:blank' }).catch(() => {});
      await sleep(100);
    }
    if (loginTab.id && settings.closeTabsAfterRun) {
      await chrome.tabs.remove(loginTab.id).catch(() => {});
    }
    await resetOpenAiAuthState(0, `${account.email}：流程结束后清理 ChatGPT 登录态。`).catch(() => {});
  }
}

async function createChatGptLoginTab(settings = {}) {
  const tab = await chrome.tabs.create({
    url: 'https://chatgpt.com/',
    active: settings.activateTabs !== false,
  });
  if (!tab?.id) throw new Error('未能创建 ChatGPT 登录标签页。');
  return tab.id;
}

async function replaceChatGptLoginTab(loginTab, settings = {}, reason = '') {
  const previousTabId = loginTab?.id || 0;
  if (previousTabId) {
    await chrome.tabs.remove(previousTabId).catch(() => {});
    loginTab.id = 0;
    await sleep(100);
  }
  await resetOpenAiAuthState(0, reason);
  loginTab.id = await createChatGptLoginTab(settings);
  await sleep(300);
}

async function updateAccountWithTokenState(accountId, tokenState = {}) {
  const session = tokenState.session && typeof tokenState.session === 'object' ? tokenState.session : {};
  await updateAccount(accountId, {
    accessToken: tokenState.accessToken,
    expiresAt: tokenState.expiresAt,
    sessionToken: normalizeString(tokenState.sessionToken || session.sessionToken || session.session_token),
    idToken: normalizeString(tokenState.idToken || session.idToken || session.id_token),
    accountId: normalizeString(tokenState.accountId || session.account?.id || session.account_id),
    userId: normalizeString(tokenState.userId || session.user?.id || session.user_id),
    planType: normalizeString(tokenState.planType || session.account?.planType || session.account?.plan_type),
    cockpitRefreshRequired: false,
    cockpitErrorCode: '',
    status: 'valid',
    lastRefreshAt: new Date().toISOString(),
    lastLiveValidatedAt: new Date().toISOString(),
    lastCheckedAt: new Date().toISOString(),
    error: '',
  });
}

async function resolveMailAccountForChatGptAccount(account, mailAccounts, settings = {}) {
  const selectedProvider = normalizeMailProviderId(settings.mailProvider);
  if (selectedProvider === MAIL_PROVIDER_LUCKMAIL) {
    const runtimeAccount = await buildRuntimeLuckmailMailAccount(account, settings);
    addLog(`${account.email}：使用 LuckMail 全局配置查信。`, 'info');
    return runtimeAccount;
  }
  if (selectedProvider === MAIL_PROVIDER_CLOUDFLARE_TEMP_EMAIL) {
    const runtimeAccount = buildRuntimeCloudflareTempEmailMailAccount(settings);
    if (runtimeAccount) {
      addLog(`${account.email}：使用 Cloudflare Temp Email 全局配置查信。`, 'info');
    }
    return runtimeAccount;
  }

  const importedAccount = findMailAccountForChatGptAccount(account, mailAccounts, selectedProvider);
  if (importedAccount) return importedAccount;

  if (!selectedProvider && normalizeString(settings.cloudflareTempEmailBaseUrl)) {
    const runtimeAccount = buildRuntimeCloudflareTempEmailMailAccount(settings);
    if (runtimeAccount) {
      addLog(`${account.email}：使用 Cloudflare Temp Email 全局配置查信。`, 'info');
    }
    return runtimeAccount;
  }
  return null;
}

function findMailAccountForChatGptAccount(account, mailAccounts, provider = '') {
  const target = normalizeEmail(account.mailAccountEmail || account.email);
  const targets = getEmailMatchCandidates(target);
  const normalizedProvider = normalizeMailProviderId(provider);
  const candidates = normalizedProvider
    ? mailAccounts.filter((item) => normalizeMailProviderId(item.provider) === normalizedProvider)
    : mailAccounts;
  return candidates.find((item) => targets.some((email) => emailsMatch(item.email, email)))
    || candidates.find((item) => emailsMatch(item.email, account.email))
    || candidates.find((item) => targets.some((email) => emailsMatch(item.receiveMailbox, email)))
    || candidates.find((item) => normalizeMailAccountEmail(item.email) === '*')
    || null;
}

async function findLuckmailPurchaseForEmail(email, settings = {}) {
  const target = normalizeEmail(email);
  if (!target) throw new Error('LuckMail 缺少目标邮箱。');
  const lookupTargets = getEmailMatchCandidates(target);
  const apiKey = normalizeString(settings.luckmailApiKey);
  if (!apiKey) throw new Error('LuckMail API Key 为空，请先填写。');
  const runtimeAccount = normalizeMailAccount({
    id: 'luckmail-runtime-lookup',
    provider: MAIL_PROVIDER_LUCKMAIL,
    email: getPreferredMailboxEmail(target) || target,
    apiKey,
    baseUrl: settings.luckmailBaseUrl || DEFAULT_SETTINGS.luckmailBaseUrl,
    status: 'ready',
  });
  const pageSize = 100;
  for (const keyword of lookupTargets) {
    for (let page = 1; page <= 20; page += 1) {
      const payload = await requestLuckmailJson(runtimeAccount, 'GET', '/api/v1/openapi/email/purchases', {
        params: {
          page,
          page_size: pageSize,
          keyword,
        },
      });
      const result = LuckMailUtils.normalizeLuckmailPurchaseListPage(payload);
      const found = result.list.find((item) => emailsMatch(item.email_address, target));
      if (found?.token) return found;
      if (result.total && page * pageSize >= result.total) break;
      if (!result.total && result.list.length < pageSize) break;
    }
  }
  throw new Error(`LuckMail 未找到 ${target} 或其基础邮箱对应的已购邮箱 token。`);
}

async function buildRuntimeLuckmailMailAccount(account, settings = {}) {
  const target = normalizeEmail(account.email);
  const purchase = await findLuckmailPurchaseForEmail(target, settings);
  return normalizeMailAccount({
    id: `luckmail-runtime-${target}`,
    provider: MAIL_PROVIDER_LUCKMAIL,
    email: getPreferredMailboxEmail(target) || target,
    token: purchase.token,
    apiKey: settings.luckmailApiKey,
    baseUrl: settings.luckmailBaseUrl || DEFAULT_SETTINGS.luckmailBaseUrl,
    status: 'ready',
    error: '',
  });
}

function buildRuntimeCloudflareTempEmailMailAccount(settings = {}) {
  const baseUrl = normalizeString(settings.cloudflareTempEmailBaseUrl);
  if (!baseUrl) return null;
  return normalizeMailAccount({
    id: 'cloudflare-temp-email-runtime',
    provider: MAIL_PROVIDER_CLOUDFLARE_TEMP_EMAIL,
    email: '*',
    baseUrl,
    adminAuth: settings.cloudflareTempEmailAdminAuth,
    customAuth: settings.cloudflareTempEmailCustomAuth,
    lookupMode: settings.cloudflareTempEmailLookupMode,
    receiveMailbox: settings.cloudflareTempEmailReceiveMailbox,
    status: 'ready',
    error: '',
  });
}

function isOpenAiBrowserUrl(rawUrl = '') {
  try {
    const host = new URL(rawUrl).hostname;
    return /(^|\.)((chatgpt\.com)|(chat\.openai\.com)|(auth\.openai\.com)|(auth0\.openai\.com)|(accounts\.openai\.com)|(openai\.com))$/i.test(host);
  } catch {
    return false;
  }
}

async function neutralizeOpenAiTabs(exceptTabId = 0) {
  const tabs = await chrome.tabs.query({}).catch(() => []);
  const targets = tabs.filter((tab) => (
    tab?.id
    && tab.id !== exceptTabId
    && isOpenAiBrowserUrl(tab.pendingUrl || tab.url || '')
  ));
  if (!targets.length) return;
  await Promise.all(targets.map((tab) => chrome.tabs.update(tab.id, { url: 'about:blank' }).catch(() => null)));
  addLog(`已跳走 ${targets.length} 个 ChatGPT/OpenAI 标签页，避免旧页面恢复登录态。`, 'info');
  await sleep(100);
}

async function getCookieStoreIds() {
  const stores = await chrome.cookies.getAllCookieStores().catch(() => []);
  const ids = stores.map((store) => store.id).filter((id) => id !== undefined && id !== null);
  return ids.length ? ids : [undefined];
}

function cookieRemovalUrl(cookie = {}, fallbackUrlOrDomain = 'https://chatgpt.com/') {
  try {
    if (fallbackUrlOrDomain.startsWith('http')) {
      const parsed = new URL(fallbackUrlOrDomain);
      const host = String(cookie.domain || parsed.hostname).replace(/^\./, '');
      const protocol = cookie.secure === false ? parsed.protocol : 'https:';
      const path = String(cookie.path || '/');
      return `${protocol}//${host}${path.startsWith('/') ? path : `/${path}`}`;
    }
  } catch {
    // Fall through to domain based URL.
  }
  const host = String(cookie.domain || fallbackUrlOrDomain || 'chatgpt.com').replace(/^\./, '');
  const protocol = cookie.secure ? 'https:' : 'http:';
  const path = String(cookie.path || '/');
  return `${protocol}//${host}${path.startsWith('/') ? path : `/${path}`}`;
}

function cookieKey(cookie = {}) {
  return [
    cookie.storeId || '',
    cookie.name || '',
    cookie.domain || '',
    cookie.path || '',
    JSON.stringify(cookie.partitionKey || null),
  ].join('\n');
}

async function collectOpenAiCookies() {
  const storeIds = await getCookieStoreIds();
  const queries = [];
  for (const storeId of storeIds) {
    for (const origin of OPENAI_COOKIE_ORIGINS) {
      const query = { url: `${origin}/` };
      if (storeId !== undefined) query.storeId = storeId;
      queries.push(chrome.cookies.getAll(query)
        .then((cookies) => ({ cookies, fallback: `${origin}/` }))
        .catch(() => ({ cookies: [], fallback: `${origin}/` })));
    }
    for (const domain of OPENAI_COOKIE_DOMAINS) {
      const query = { domain };
      if (storeId !== undefined) query.storeId = storeId;
      queries.push(chrome.cookies.getAll(query)
        .then((cookies) => ({ cookies, fallback: domain }))
        .catch(() => ({ cookies: [], fallback: domain })));
    }
  }
  const byKey = new Map();
  const results = await Promise.all(queries);
  for (const result of results) {
    result.cookies.forEach((cookie) => byKey.set(cookieKey(cookie), { cookie, fallback: result.fallback }));
  }
  return Array.from(byKey.values());
}

async function removeCookie(cookie = {}, fallback = '') {
  const details = {
    url: cookieRemovalUrl(cookie, fallback),
    name: cookie.name,
  };
  if (cookie.storeId !== undefined && cookie.storeId !== null) details.storeId = cookie.storeId;
  if (cookie.partitionKey) details.partitionKey = cookie.partitionKey;
  const removed = await chrome.cookies.remove(details).catch(() => null);
  if (removed || !details.partitionKey) return removed;
  const fallbackDetails = { ...details };
  delete fallbackDetails.partitionKey;
  return chrome.cookies.remove(fallbackDetails).catch(() => null);
}

async function clearOpenAiCookiesByEnumeration() {
  const entries = await collectOpenAiCookies();
  await Promise.all(entries.map(({ cookie, fallback }) => removeCookie(cookie, fallback)));
  return entries.length;
}

async function countOpenAiCookies() {
  return (await collectOpenAiCookies()).length;
}

async function resetOpenAiAuthState(tabId = 0, reason = '') {
  if (reason) addLog(reason, 'info');
  if (tabId) {
    await chrome.tabs.update(tabId, { url: 'about:blank' }).catch(() => {});
    await sleep(100);
  }
  await neutralizeOpenAiTabs(tabId);
  const removedBeforeSiteData = await clearOpenAiCookiesByEnumeration();
  const remainingCookies = await countOpenAiCookies();
  addLog(`ChatGPT/OpenAI 登录态快速清理完成：已删除 Cookie ${removedBeforeSiteData} 个，剩余 ${remainingCookies} 个。`, remainingCookies ? 'warn' : 'info');
  if (remainingCookies > 0) {
    await sleep(100);
    const retryRemoved = await clearOpenAiCookiesByEnumeration();
    const retryRemaining = await countOpenAiCookies();
    addLog(`ChatGPT/OpenAI Cookie 二次清理：删除 ${retryRemoved} 个，剩余 ${retryRemaining} 个。`, retryRemaining ? 'warn' : 'info');
  }
  await sleep(100);
}

function assertAccessTokenMatchesAccount(account, accessToken) {
  const targetEmail = normalizeEmail(account?.email);
  const tokenEmail = extractEmailFromAccessToken(accessToken);
  if (!tokenEmail) {
    throw new Error(`${targetEmail || '当前账号'}：读取到的 accessToken 无法解析邮箱，已阻止更新。`);
  }
  if (targetEmail && !emailsMatch(tokenEmail, targetEmail)) {
    throw new Error(`${targetEmail}：读取到的登录账号是 ${tokenEmail}，不是当前账号，已阻止用错误 JSON 覆盖。`);
  }
}

function buildTokenStateFromSession(account, session) {
  assertAccessTokenMatchesAccount(account, session?.accessToken);
  return {
    accessToken: session.accessToken,
    expiresAt: getTokenExpiryMs(session.accessToken) || Date.now() + 60 * 60 * 1000,
    session: session.session,
  };
}

async function driveLoginAndReadToken(loginTab, account, mailAccount, requestedAt, settings) {
  const deadline = Date.now() + 5 * 60 * 1000;
  const triedCodes = new Set();
  let emailFilled = false;
  let codeRequestedAt = Number(requestedAt) || Date.now();
  let lastStateLabel = '';
  let wrongSessionResets = 0;
  let codePageLogged = false;
  let codeSubmittedAt = 0;
  let waitingAfterCodeLogged = false;

  while (Date.now() < deadline) {
    const tabId = loginTab?.id || 0;
    const tab = await chrome.tabs.get(tabId).catch(() => null);
    if (!tab?.id) throw new Error('登录标签页已关闭。');
    const url = tab.url || '';
    if (isSessionUrl(url)) {
      const session = await readSessionFromTab(tabId).catch(() => null);
      if (session?.accessToken) {
        const tokenEmail = extractEmailFromAccessToken(session.accessToken);
        if (tokenEmail && !emailsMatch(tokenEmail, account.email) && wrongSessionResets < 3) {
          wrongSessionResets += 1;
          addLog(`${account.email}：当前浏览器仍是 ${tokenEmail} 登录态，清理后重新进入登录流程（${wrongSessionResets}/3）。`, 'warn');
          await replaceChatGptLoginTab(loginTab, settings, `${account.email}：关闭旧标签页并清理 ChatGPT 登录态。`);
          emailFilled = false;
          codePageLogged = false;
          codeSubmittedAt = 0;
          waitingAfterCodeLogged = false;
          lastStateLabel = '';
          codeRequestedAt = Date.now();
          continue;
        }
        return buildTokenStateFromSession(account, session);
      }
    }

    if (!isAuthOrSessionUrl(url)) {
      await sleep(300);
      continue;
    }

    const state = await sendAuthCommand(tabId, 'TK_AUTH_STATE', {}).catch((error) => ({
      error: error?.message || String(error || ''),
    }));
    if (state.fatal && state.fatalReason === 'account_deactivated') {
      throw new AccountDeactivatedError(account.email, state.fatalReason);
    }
    if (state.blocked) {
      throw new Error(`登录页需要人工处理：${state.blockedReason || 'captcha/security check'}`);
    }
    const label = `${state.page || 'unknown'} ${url}`;
    if (label !== lastStateLabel) {
      addLog(`${account.email}：当前登录状态 ${state.page || 'unknown'}。`, 'info');
      lastStateLabel = label;
    }

    if (state.page === 'logged-in') {
      if (emailFilled || codeSubmittedAt) {
        addLog(`${account.email}：网页登录已完成，开始读取新的 accessToken。`, 'info');
        await chrome.tabs.update(tabId, { url: 'https://chatgpt.com/api/auth/session' }).catch(() => {});
        await sleep(700);
        continue;
      }
      if (wrongSessionResets < 3) {
        wrongSessionResets += 1;
        addLog(`${account.email}：检测到已有登录态，清理后重新登录目标账号（${wrongSessionResets}/3）。`, 'warn');
        await replaceChatGptLoginTab(loginTab, settings, `${account.email}：关闭旧标签页并清理 ChatGPT 登录态。`);
        emailFilled = false;
        codePageLogged = false;
        codeSubmittedAt = 0;
        waitingAfterCodeLogged = false;
        lastStateLabel = '';
        codeRequestedAt = Date.now();
      } else {
        await chrome.tabs.update(tabId, { url: 'https://chatgpt.com/' }).catch(() => {});
        await sleep(700);
      }
      continue;
    }
    if (state.page === 'entry') {
      await sendAuthCommand(tabId, 'TK_START_LOGIN', {});
      await sleep(400);
      continue;
    }
    if (state.page === 'email' && !emailFilled) {
      await sendAuthCommand(tabId, 'TK_FILL_EMAIL', { email: account.email });
      emailFilled = true;
      codeRequestedAt = Date.now();
      addLog(`${account.email}：已提交邮箱，等待登录验证码邮件。`, 'info');
      await sleep(600);
      continue;
    }
    if (state.page === 'password') {
      const switched = await sendAuthCommand(tabId, 'TK_USE_EMAIL_CODE', { email: account.email }).catch((error) => ({
        error: error?.message || String(error || ''),
      }));
      if (switched?.error) {
        throw new Error(`登录页要求密码，且没有找到邮箱验证码登录入口：${switched.error}`);
      }
      codeRequestedAt = Date.now();
      addLog(`${account.email}：已切换到邮箱验证码登录。`, 'info');
      await sleep(600);
      continue;
    }
    if (state.page === 'code') {
      if (!codePageLogged) {
        addLog(`${account.email}：已到验证码页，开始按该邮箱查信并填入验证码。`, 'info');
        codePageLogged = true;
      }
      if (codeSubmittedAt && !state.codeError) {
        if (Date.now() - codeSubmittedAt < 25000) {
          if (!waitingAfterCodeLogged) {
            addLog(`${account.email}：验证码已提交，等待页面完成登录，不再重复查信。`, 'info');
            waitingAfterCodeLogged = true;
          }
          await sleep(700);
          continue;
        }
        addLog(`${account.email}：验证码提交后仍停留在验证码页，重新查信尝试下一个验证码。`, 'warn');
        codeSubmittedAt = 0;
        waitingAfterCodeLogged = false;
      }
      if (state.codeError) {
        addLog(`${account.email}：页面提示验证码无效或过期，自动重新发送邮件并等待新验证码。`, 'warn');
        codeSubmittedAt = 0;
        waitingAfterCodeLogged = false;
        const resent = await sendAuthCommand(tabId, 'TK_RESEND_CODE', {}).catch((error) => ({
          error: error?.message || String(error || ''),
        }));
        if (resent?.clicked) {
          codeRequestedAt = Date.now();
          addLog(`${account.email}：已点击重新发送验证码邮件，旧验证码会被排除。`, 'info');
          await sleep(1000);
          continue;
        }
        addLog(`${account.email}：未找到重新发送按钮，继续排除旧验证码后查信。`, 'warn');
      }
      const codeResult = await fetchVerificationCode(mailAccount, codeRequestedAt, Array.from(triedCodes), settings, account);
      triedCodes.add(codeResult.code);
      addLog(`${account.email}：读取到邮箱验证码 ${codeResult.code}。`, 'info');
      await sendAuthCommand(tabId, 'TK_FILL_CODE', { code: codeResult.code });
      codeSubmittedAt = Date.now();
      waitingAfterCodeLogged = false;
      await sleep(800);
      continue;
    }

    await sleep(300);
  }

  throw new Error('登录超时，未能读取到新的 accessToken。');
}

function isAuthOrSessionUrl(rawUrl = '') {
  try {
    const host = new URL(rawUrl).hostname;
    return AUTH_HOST_RE.test(host);
  } catch {
    return false;
  }
}

function isSessionUrl(rawUrl = '') {
  try {
    const host = new URL(rawUrl).hostname;
    return SESSION_HOST_RE.test(host);
  } catch {
    return false;
  }
}

async function sendAuthCommand(tabId, type, payload) {
  await ensureScript(tabId, 'content/auth.js');
  const result = await chrome.tabs.sendMessage(tabId, { type, payload });
  if (result?.ok === false || result?.error) {
    throw new Error(result.error || `${type} 执行失败。`);
  }
  return result || {};
}

async function readSessionFromTab(tabId) {
  await ensureScript(tabId, 'content/session-reader.js');
  const result = await chrome.tabs.sendMessage(tabId, { type: 'TK_READ_SESSION' });
  if (result?.ok === false || result?.error) throw new Error(result.error || '读取 session 失败。');
  if (!result?.accessToken) throw new Error('当前 ChatGPT 页面没有返回 accessToken。');
  return result;
}

async function ensureScript(tabId, file) {
  try {
    const ping = file.includes('session-reader') ? 'TK_SESSION_PING' : 'TK_AUTH_PING';
    await chrome.tabs.sendMessage(tabId, { type: ping });
    return;
  } catch {
    await chrome.scripting.executeScript({ target: { tabId }, files: [file] });
  }
}

async function fetchVerificationCode(mailAccount, requestedAt, excludeCodes, settings, chatGptAccount = {}) {
  const provider = normalizeMailProviderId(mailAccount.provider) || MAIL_PROVIDER_MICROSOFT;
  if (provider === MAIL_PROVIDER_LUCKMAIL) {
    return fetchLuckmailVerificationCode(mailAccount, requestedAt, excludeCodes, settings, chatGptAccount);
  }
  if (provider === MAIL_PROVIDER_CLOUDFLARE_TEMP_EMAIL) {
    return fetchCloudflareTempEmailVerificationCode(mailAccount, requestedAt, excludeCodes, settings, chatGptAccount);
  }
  return fetchMicrosoftVerificationCode(mailAccount, requestedAt, excludeCodes, settings);
}

function buildOpenAiCodeFilters(requestedAt, excludeCodes = []) {
  return {
    filterAfterTimestamp: Math.max(0, Number(requestedAt) - 5000),
    senderFilters: ['openai', 'chatgpt', 'auth0'],
    subjectFilters: ['code', '验证码', 'login', 'sign in', 'openai', 'chatgpt'],
    requiredKeywords: ['openai', 'chatgpt', 'code', '验证码'],
    excludeCodes,
  };
}

function normalizeMicrosoftLocalMessages(messages = []) {
  return (Array.isArray(messages) ? messages : [])
    .map((message) => MultiPageMicrosoftEmail.normalizeMessage(message, message?.mailbox || 'INBOX'));
}

async function requestMicrosoftLocalJson(mailAccount, settings = {}, path = '/messages', extraPayload = {}) {
  if (!mailAccount.email || mailAccount.email === '*') {
    throw new Error('Outlook/Hotmail 本地助手缺少邮箱地址。');
  }
  if (!mailAccount.clientId || !mailAccount.refreshToken) {
    throw new Error(`邮箱 ${mailAccount.email} 缺少 clientId 或 refreshToken。`);
  }
  const helperUrl = normalizeMicrosoftLocalBaseUrl(settings.microsoftLocalBaseUrl);
  let response;
  try {
    response = await fetchWithTimeout(buildMicrosoftLocalEndpoint({ ...settings, microsoftLocalBaseUrl: helperUrl }, path), {
      method: 'POST',
      cache: 'no-store',
      headers: {
        'Content-Type': 'application/json',
        Accept: 'application/json',
      },
      body: JSON.stringify({
        email: mailAccount.email,
        clientId: mailAccount.clientId,
        refreshToken: mailAccount.refreshToken,
        ...extraPayload,
      }),
    }, 45000);
  } catch (error) {
    throw new Error(`Outlook/Hotmail 本地助手请求失败：${error?.message || error}。请先启动 ${helperUrl}。`);
  }

  const text = await response.text();
  let payload = null;
  try {
    payload = text ? JSON.parse(text) : {};
  } catch {
    payload = { raw: text };
  }

  if (!response.ok || payload?.ok === false) {
    const detail = normalizeString(payload?.error || payload?.message || text || `HTTP ${response.status}`);
    throw new Error(`Outlook/Hotmail 本地助手返回失败：${detail}`);
  }
  return payload || {};
}

async function fetchMicrosoftMailboxMessagesViaLocalHelper(mailAccount, settings = {}, options = {}) {
  const payload = await requestMicrosoftLocalJson(mailAccount, settings, '/messages', {
    mailboxes: options.mailboxes || ['INBOX', 'Junk'],
    top: options.top || 10,
  });
  if (payload.nextRefreshToken && payload.nextRefreshToken !== mailAccount.refreshToken) {
    await updateMailAccount(mailAccount.id, {
      refreshToken: payload.nextRefreshToken,
      lastCheckedAt: new Date().toISOString(),
      error: '',
    });
  }
  return {
    messages: normalizeMicrosoftLocalMessages(payload.messages || []),
    nextRefreshToken: normalizeString(payload.nextRefreshToken),
    transport: normalizeString(payload.transport || 'local-helper'),
    mailboxResults: Array.isArray(payload.mailboxResults) ? payload.mailboxResults : [],
  };
}

async function fetchMicrosoftVerificationCodeViaLocalHelper(mailAccount, requestedAt, excludeCodes, settings = {}) {
  const filters = buildOpenAiCodeFilters(requestedAt, excludeCodes);
  const maxRetries = clampInt(settings.mailMaxRetries, 1, 120, DEFAULT_SETTINGS.mailMaxRetries);
  const retryDelayMs = clampInt(settings.mailRetryDelayMs, 1000, 60000, DEFAULT_SETTINGS.mailRetryDelayMs);
  let lastError = null;

  for (let attempt = 1; attempt <= maxRetries; attempt += 1) {
    try {
      const payload = await requestMicrosoftLocalJson(mailAccount, settings, '/code', {
        mailboxes: ['INBOX', 'Junk'],
        top: 10,
        senderFilters: filters.senderFilters,
        subjectFilters: filters.subjectFilters,
        requiredKeywords: filters.requiredKeywords,
        excludeCodes: filters.excludeCodes,
        filterAfterTimestamp: filters.filterAfterTimestamp,
      });
      if (payload.nextRefreshToken && payload.nextRefreshToken !== mailAccount.refreshToken) {
        await updateMailAccount(mailAccount.id, {
          refreshToken: payload.nextRefreshToken,
          lastCheckedAt: new Date().toISOString(),
          error: '',
        });
        mailAccount = { ...mailAccount, refreshToken: payload.nextRefreshToken };
      }
      const code = normalizeString(payload.code);
      if (code) {
        const message = payload.message
          ? MultiPageMicrosoftEmail.normalizeMessage(payload.message, payload.message?.mailbox || 'INBOX')
          : null;
        return {
          code,
          emailTimestamp: MultiPageMicrosoftEmail.getMessageTimestamp(message) || Date.now(),
          messageId: message?.id || '',
          sender: MultiPageMicrosoftEmail.getMessageSender(message),
          subject: normalizeString(message?.subject),
          mailbox: normalizeString(message?.mailbox) || 'INBOX',
          message,
          transport: normalizeString(payload.transport || 'local-helper'),
        };
      }
      lastError = new Error(`Outlook/Hotmail 本地助手暂未返回匹配验证码（${attempt}/${maxRetries}）。`);
    } catch (error) {
      lastError = error;
    }
    if (attempt < maxRetries) await sleep(retryDelayMs);
  }

  throw lastError || new Error('Outlook/Hotmail 本地助手未找到验证码。');
}

async function fetchMicrosoftVerificationCode(mailAccount, requestedAt, excludeCodes, settings) {
  if (!mailAccount.clientId || !mailAccount.refreshToken) {
    throw new Error(`邮箱 ${mailAccount.email} 缺少 clientId 或 refreshToken。`);
  }
  return fetchMicrosoftVerificationCodeViaLocalHelper(mailAccount, requestedAt, excludeCodes, settings);
}

async function requestLuckmailJson(mailAccount, method, path, options = {}) {
  const baseUrl = LuckMailUtils.normalizeLuckmailBaseUrl(mailAccount.baseUrl || 'https://mails.luckyous.com');
  const apiKey = normalizeString(mailAccount.apiKey);
  if (!apiKey) {
    throw new Error(`LuckMail ${mailAccount.email} 缺少 API Key。`);
  }
  const url = new URL(`${baseUrl}${path}`);
  if (options.params && typeof options.params === 'object') {
    for (const [key, value] of Object.entries(options.params)) {
      if (value !== undefined && value !== null && value !== '') url.searchParams.set(key, String(value));
    }
  }
  const response = await fetchWithTimeout(url.toString(), {
    method,
    cache: 'no-store',
    headers: {
      Accept: 'application/json',
      'X-API-Key': apiKey,
    },
  }, options.timeout || 30000);
  const payload = await response.json().catch(() => null);
  if (!response.ok) {
    throw new Error(`LuckMail HTTP ${response.status}：${payload?.message || response.statusText || '请求失败'}`);
  }
  if (!payload || typeof payload !== 'object') {
    throw new Error('LuckMail 返回数据不是有效 JSON。');
  }
  if (payload.code !== 0) {
    throw new Error(`LuckMail 接口失败：${payload.message || 'unknown error'}`);
  }
  return payload.data;
}

async function fetchLuckmailVerificationCode(mailAccount, requestedAt, excludeCodes, settings, chatGptAccount = {}) {
  const token = normalizeString(mailAccount.token);
  if (!token) {
    throw new Error(`LuckMail ${mailAccount.email} 缺少邮箱 token。`);
  }
  const targetEmail = normalizeEmail(chatGptAccount.email || mailAccount.email);
  const filters = buildOpenAiCodeFilters(requestedAt, excludeCodes);
  const maxRetries = Math.max(1, Number(settings.mailMaxRetries) || 1);
  const retryDelayMs = Math.max(5000, Number(settings.mailRetryDelayMs) || 5000);
  let lastError = null;

  for (let attempt = 1; attempt <= maxRetries; attempt += 1) {
    try {
      const tokenCode = LuckMailUtils.normalizeLuckmailTokenCode(await requestLuckmailJson(
        mailAccount,
        'GET',
        `/api/v1/openapi/email/token/${encodeURIComponent(token)}/code`
      ));
      const code = normalizeString(tokenCode.verification_code || tokenCode.mail?.verification_code);
      const remoteEmail = normalizeEmail(tokenCode.email_address);
      if (targetEmail && remoteEmail && !emailsMatch(remoteEmail, targetEmail)) {
        throw new Error(`LuckMail token 对应邮箱与当前账号不一致：当前账号 ${targetEmail}；token 邮箱 ${remoteEmail}`);
      }
      const mail = tokenCode.mail || null;
      const receivedAt = LuckMailUtils.normalizeTimestamp(mail?.received_at);
      if (code && !filters.excludeCodes.includes(code) && (!receivedAt || receivedAt >= filters.filterAfterTimestamp)) {
        return {
          code,
          emailTimestamp: receivedAt || Date.now(),
          messageId: mail?.message_id || '',
          message: mail,
        };
      }

      const mailList = await requestLuckmailJson(
        mailAccount,
        'GET',
        `/api/v1/openapi/email/token/${encodeURIComponent(token)}/mails`
      );
      const match = LuckMailUtils.pickLuckmailVerificationMail(mailList?.mails || [], {
        afterTimestamp: filters.filterAfterTimestamp,
        senderFilters: filters.senderFilters,
        subjectFilters: filters.subjectFilters,
        excludeCodes,
      });
      if (match?.code) {
        return {
          code: match.code,
          emailTimestamp: match.receivedAt || Date.now(),
          messageId: match.mail?.message_id || '',
          message: match.mail,
        };
      }
      lastError = new Error(`LuckMail 暂未返回匹配验证码（${attempt}/${maxRetries}）。`);
    } catch (error) {
      lastError = error;
    }
    if (attempt < maxRetries) await sleep(retryDelayMs);
  }

  throw lastError || new Error('LuckMail 未找到验证码。');
}

function buildCloudflareTempEmailHeaders(mailAccount, options = {}) {
  return CloudflareTempEmailUtils.buildCloudflareTempEmailHeaders({
    adminAuth: mailAccount.adminAuth,
    customAuth: mailAccount.customAuth,
  }, options);
}

async function requestCloudflareTempEmailJson(mailAccount, path, options = {}) {
  const baseUrl = CloudflareTempEmailUtils.normalizeCloudflareTempEmailBaseUrl(mailAccount.baseUrl);
  if (!baseUrl) {
    throw new Error(`Cloudflare Temp Email ${mailAccount.email} 缺少有效 Temp API 地址。`);
  }
  const url = new URL(CloudflareTempEmailUtils.joinCloudflareTempEmailUrl(baseUrl, path));
  if (options.searchParams && typeof options.searchParams === 'object') {
    for (const [key, value] of Object.entries(options.searchParams)) {
      if (value !== undefined && value !== null && value !== '') url.searchParams.set(key, String(value));
    }
  }
  const response = await fetchWithTimeout(url.toString(), {
    method: options.method || 'GET',
    cache: 'no-store',
    headers: buildCloudflareTempEmailHeaders(mailAccount, { acceptJson: true }),
  }, options.timeout || 30000);
  const payload = await response.json().catch(() => null);
  if (!response.ok) {
    if (response.status === 401 || response.status === 403) {
      throw new Error(`Cloudflare Temp Email HTTP ${response.status}：认证失败，请检查 TEMP API、ADMIN AUTH${mailAccount.customAuth ? '、CUSTOM AUTH' : '，如站点启用了访问密码还要填写 CUSTOM AUTH'} 是否正确并已保存。`);
    }
    throw new Error(`Cloudflare Temp Email HTTP ${response.status}：${payload?.message || response.statusText || '请求失败'}`);
  }
  return payload;
}

async function fetchCloudflareTempEmailVerificationCode(mailAccount, requestedAt, excludeCodes, settings, chatGptAccount = {}) {
  const targetEmail = normalizeEmail(chatGptAccount.email)
    || normalizeEmail(mailAccount.email === '*' ? '' : mailAccount.email)
    || normalizeEmail(mailAccount.receiveMailbox);
  if (!targetEmail) {
    throw new Error('Cloudflare Temp Email 缺少目标邮箱。');
  }
  const lookupMode = normalizeCloudflareLookupMode(mailAccount.lookupMode);
  const queryAddresses = lookupMode === 'registration-email' ? [''] : getEmailMatchCandidates(targetEmail);
  const filters = buildCloudflareTempEmailCodeFilters(requestedAt, excludeCodes);
  const maxRetries = Math.max(1, Number(settings.mailMaxRetries) || 1);
  const retryDelayMs = Math.max(3000, Number(settings.mailRetryDelayMs) || 3000);
  let lastError = null;

  for (let attempt = 1; attempt <= maxRetries; attempt += 1) {
    try {
      const directMessages = [];
      for (const queryAddress of queryAddresses) {
        const directPayload = await requestCloudflareTempEmailJson(mailAccount, '/admin/mails', {
          searchParams: { limit: 20, offset: 0, address: queryAddress },
        });
        directMessages.push(...filterCloudflareTempEmailMessagesForTarget(
          CloudflareTempEmailUtils.normalizeCloudflareTempEmailMailApiMessages(directPayload),
          targetEmail,
          queryAddress,
          lookupMode
        ));
      }
      addLog(`${targetEmail}：Temp 查信 ${attempt}/${maxRetries}，按邮箱查询返回 ${directMessages.length} 封。`, 'info');
      let matchResult = pickVerificationCodeWithTimeFallback(directMessages, filters);
      let match = matchResult.match;
      if (!match?.code) {
        const recentPayload = await requestCloudflareTempEmailJson(mailAccount, '/admin/mails', {
          searchParams: { limit: 50, offset: 0 },
        });
        const recentMessages = filterCloudflareTempEmailMessagesForTarget(
          CloudflareTempEmailUtils.normalizeCloudflareTempEmailMailApiMessages(recentPayload),
          targetEmail,
          '',
          lookupMode
        );
        addLog(`${targetEmail}：Temp 最近邮件本地匹配 ${recentMessages.length} 封。`, 'info');
        const sample = summarizeCloudflareTempEmailMessagesForLog(recentMessages);
        if (sample) addLog(`${targetEmail}：Temp 最近邮件样本：${sample}`, 'info');
        matchResult = pickVerificationCodeWithTimeFallback(recentMessages, filters);
        match = matchResult.match;
      }
      if (match?.code) {
        if (matchResult.usedTimeFallback) {
          addLog(`${targetEmail}：严格时间窗口未命中，已用时间回退匹配到验证码。`, 'warn');
        }
        return match;
      }
      lastError = new Error(`Cloudflare Temp Email 暂未返回匹配验证码（${attempt}/${maxRetries}）。`);
    } catch (error) {
      lastError = error;
    }
    if (attempt < maxRetries) await sleep(retryDelayMs);
  }

  throw lastError || new Error('Cloudflare Temp Email 未找到验证码。');
}

function buildCloudflareTempEmailCodeFilters(requestedAt, excludeCodes = []) {
  return {
    filterAfterTimestamp: Math.max(0, Number(requestedAt) - 5000),
    senderFilters: [],
    subjectFilters: [],
    requiredKeywords: [],
    codePatterns: [],
    excludeCodes,
  };
}

function pickVerificationCodeWithTimeFallback(messages = [], filters = {}) {
  const strictMatch = MultiPageMicrosoftEmail.extractVerificationCodeFromMessages(messages, filters);
  if (strictMatch?.code) {
    return { match: strictMatch, usedTimeFallback: false };
  }
  const timeFallbackMatch = MultiPageMicrosoftEmail.extractVerificationCodeFromMessages(messages, {
    ...filters,
    filterAfterTimestamp: 0,
    requiredKeywords: [],
  });
  return {
    match: timeFallbackMatch || null,
    usedTimeFallback: Boolean(timeFallbackMatch?.code),
  };
}

function summarizeCloudflareTempEmailMessagesForLog(messages = []) {
  return (Array.isArray(messages) ? messages : [])
    .slice()
    .sort((left, right) => {
      const leftTime = Date.parse(left.receivedDateTime || '') || 0;
      const rightTime = Date.parse(right.receivedDateTime || '') || 0;
      return rightTime - leftTime;
    })
    .slice(0, 3)
    .map((message) => {
      const receivedAt = message?.receivedDateTime || '未知时间';
      const sender = message?.from?.emailAddress?.address || '未知发件人';
      const subject = message?.subject || '（无主题）';
      const preview = String(message?.bodyPreview || '').replace(/\s+/g, ' ').trim().slice(0, 80);
      const address = message?.address || '未知地址';
      return `[${address}] ${receivedAt} | ${sender} | ${subject} | ${preview}`;
    })
    .join(' || ');
}

function filterCloudflareTempEmailMessagesForTarget(messages = [], targetEmail = '', queryAddress = '', lookupMode = '') {
  const target = normalizeEmail(targetEmail);
  const targetCandidates = getEmailMatchCandidates(target);
  const queryCandidates = getEmailMatchCandidates(queryAddress);
  return (Array.isArray(messages) ? messages : []).filter((message) => {
    const originalRecipient = normalizeEmail(message.originalRecipient);
    if (originalRecipient) {
      return targetCandidates.some((email) => emailsMatch(originalRecipient, email));
    }
    if (lookupMode === 'registration-email') {
      return cloudflareTempEmailMessageContainsTarget(message, target);
    }
    const address = normalizeEmail(message.address);
    return !address
      || targetCandidates.some((email) => emailsMatch(address, email))
      || queryCandidates.some((email) => emailsMatch(address, email))
      || cloudflareTempEmailMessageContainsTarget(message, target);
  });
}

function cloudflareTempEmailMessageContainsTarget(message = {}, targetEmail = '') {
  const targets = getEmailMatchCandidates(targetEmail);
  if (!targets.length) return false;
  const haystack = [
    message.address,
    message.originalRecipient,
    message.subject,
    message.bodyPreview,
    message.raw,
  ].map((value) => String(value || '').toLowerCase()).join('\n');
  return targets.some((target) => haystack.includes(target));
}

async function testMailAccount(id) {
  const { mailAccounts, settings } = await getData();
  const account = mailAccounts.find((item) => item.id === id);
  if (!account) throw new Error('邮箱账号不存在。');
  const provider = normalizeMailProviderId(account.provider) || MAIL_PROVIDER_MICROSOFT;
  let count = 0;
  if (provider === MAIL_PROVIDER_MICROSOFT) {
    const result = await fetchMicrosoftMailboxMessagesViaLocalHelper(account, settings, {
      mailboxes: ['INBOX', 'Junk'],
      top: 3,
    });
    count = result.messages.length;
  } else if (provider === MAIL_PROVIDER_LUCKMAIL) {
    const token = normalizeString(account.token);
    if (!token) throw new Error(`LuckMail ${account.email} 缺少邮箱 token。`);
    const result = await requestLuckmailJson(account, 'GET', `/api/v1/openapi/email/token/${encodeURIComponent(token)}/mails`);
    count = LuckMailUtils.normalizeLuckmailTokenMails(result?.mails || []).length;
  } else if (provider === MAIL_PROVIDER_CLOUDFLARE_TEMP_EMAIL) {
    const payload = await requestCloudflareTempEmailJson(account, '/admin/mails', {
      searchParams: { limit: 3, offset: 0, address: account.email === '*' ? '' : account.email },
    });
    count = CloudflareTempEmailUtils.normalizeCloudflareTempEmailMailApiMessages(payload).length;
  }
  await updateMailAccount(account.id, {
    status: 'ready',
    lastCheckedAt: new Date().toISOString(),
    error: '',
  });
  addLog(`邮箱 ${account.email}（${provider}）测试成功，读取 ${count} 封邮件。`, 'info');
  return { count, settings };
}

function sanitizeFileName(value, fallback = 'account') {
  return normalizeString(value)
    .replace(/[\\/:*?"<>|]+/g, '-')
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-+|-+$/g, '') || fallback;
}

async function downloadJson(filename, payload) {
  const json = `${JSON.stringify(payload, null, 2)}\n`;
  const url = `data:application/json;charset=utf-8,${encodeURIComponent(json)}`;
  await chrome.downloads.download({
    url,
    filename,
    saveAs: false,
    conflictAction: 'overwrite',
  });
}

async function buildCpaAuthJsonsForAccounts(accountIds = []) {
  const ids = new Set((Array.isArray(accountIds) ? accountIds : []).map(String).filter(Boolean));
  const { accounts } = await getData();
  const targetAccounts = accounts.filter((account) => ids.has(account.id));
  return {
    items: targetAccounts.map((account) => {
      const sessionAuth = buildGuJumpgateCpaSessionAuthJson(
        buildGuJumpgateSessionStateFromAccount(account),
        { now: new Date() }
      );
      return {
        id: account.id,
        email: account.email,
        fileName: sessionAuth.fileName,
        json: sessionAuth.authJson,
      };
    }),
  };
}

async function convertSessionJsonToCpa(content = '') {
  let session = null;
  try {
    session = JSON.parse(String(content || ''));
  } catch {
    throw new Error('ChatGPT session JSON 不是有效 JSON。');
  }
  const sessionAuth = buildGuJumpgateCpaSessionAuthJson({ session }, { now: new Date() });
  await downloadJson(`chatgpt-session-cpa/${sessionAuth.fileName}`, sessionAuth.authJson);
  if (!sessionAuth.hasRefreshToken) {
    addLog('当前 SESSION 未包含 refresh_token，导出的 CPA JSON 无法自动续期。', 'warn');
  }
  addLog(`${sessionAuth.email || sessionAuth.fileName}：已按 GuJumpgate SESSION 最后一步转换为 CPA JSON。`, 'info');
}

function buildGuJumpgateSyntheticCodexIdToken(email, accountId, planType, userId, expiresAt) {
  if (!accountId) return '';
  const now = Math.trunc(Date.now() / 1000);
  const authInfo = { chatgpt_account_id: accountId };
  const expires = epochSecondsFromValue(expiresAt) || now + 90 * 24 * 60 * 60;
  if (planType) authInfo.chatgpt_plan_type = planType;
  if (userId) {
    authInfo.chatgpt_user_id = userId;
    authInfo.user_id = userId;
  }
  const payload = {
    iat: now,
    exp: expires,
    'https://api.openai.com/auth': authInfo,
  };
  if (email) payload.email = email;
  return `${encodeBase64UrlJson({ alg: 'none', typ: 'JWT', cpa_synthetic: true })}.${encodeBase64UrlJson(payload)}.synthetic`;
}

function buildGuJumpgateCpaSessionAuthJson(state = {}, options = {}) {
  const session = state?.session && typeof state.session === 'object' && !Array.isArray(state.session)
    ? state.session
    : {};
  const accessToken = normalizeString(state?.accessToken || session?.accessToken || session?.access_token);
  if (!accessToken) {
    throw new Error('未读取到可导入的 ChatGPT accessToken。');
  }
  const inputIdToken = firstNonEmptyString(
    state?.idToken,
    state?.id_token,
    session?.idToken,
    session?.id_token
  );
  const refreshToken = firstNonEmptyString(
    state?.refreshToken,
    state?.refresh_token,
    session?.refreshToken,
    session?.refresh_token
  );
  const sessionToken = firstNonEmptyString(
    state?.sessionToken,
    state?.session_token,
    session?.sessionToken,
    session?.session_token
  );
  const accessPayload = parseJwtPayload(accessToken) || {};
  const idPayload = parseJwtPayload(inputIdToken) || {};
  const accessAuth = getOpenAiAuthSection(accessPayload);
  const idAuth = getOpenAiAuthSection(idPayload);
  const profile = getOpenAiProfileSection(accessPayload);
  const accountIdentifierEmail = normalizeString(state?.accountIdentifierType).toLowerCase() === 'email'
    ? normalizeEmail(state?.accountIdentifier)
    : '';
  const expiresAt = firstNonEmptyString(
    timestampFromUnixSeconds(accessPayload.exp),
    normalizeTimestamp(session?.expires),
    normalizeTimestamp(session?.expiresAt),
    normalizeTimestamp(session?.expired),
    normalizeTimestamp(session?.expires_at)
  );
  const email = firstNonEmptyString(
    normalizeEmail(session?.user?.email),
    normalizeEmail(session?.email),
    normalizeEmail(state?.email),
    accountIdentifierEmail,
    normalizeEmail(profile?.email),
    normalizeEmail(idPayload?.email),
    normalizeEmail(accessPayload?.email)
  );
  const accountId = firstNonEmptyString(
    session?.account?.id,
    session?.account_id,
    accessAuth.chatgpt_account_id,
    idAuth.chatgpt_account_id
  );
  const userId = firstNonEmptyString(
    session?.user?.id,
    session?.user_id,
    accessAuth.chatgpt_user_id,
    accessAuth.user_id,
    idAuth.chatgpt_user_id,
    idAuth.user_id
  );
  const planType = firstNonEmptyString(
    session?.account?.planType,
    session?.account?.plan_type,
    session?.planType,
    session?.plan_type,
    accessAuth.chatgpt_plan_type,
    idAuth.chatgpt_plan_type
  );
  const exportedAt = normalizeTimestamp(options.now || new Date()) || new Date().toISOString();
  const syntheticIdToken = inputIdToken
    ? ''
    : buildGuJumpgateSyntheticCodexIdToken(email, accountId, planType, userId, expiresAt);
  const idToken = inputIdToken || syntheticIdToken;
  const authJson = Object.fromEntries(Object.entries({
    type: 'codex',
    account_id: accountId,
    chatgpt_account_id: accountId,
    email,
    name: firstNonEmptyString(email, state?.email, 'ChatGPT Account'),
    plan_type: planType,
    chatgpt_plan_type: planType,
    id_token: idToken,
    id_token_synthetic: syntheticIdToken ? true : undefined,
    access_token: accessToken,
    refresh_token: refreshToken || '',
    session_token: sessionToken,
    last_refresh: exportedAt,
    expired: expiresAt,
    disabled: session?.disabled === true ? true : undefined,
  }).filter(([, value]) => value !== undefined && value !== null && value !== ''));
  return {
    authJson,
    accountId,
    email,
    expiresAt,
    fileName: buildCpaAuthFileName(authJson),
    hasRefreshToken: Boolean(refreshToken),
  };
}

function firstNonEmptyString(...values) {
  for (const value of values) {
    const normalized = normalizeString(value);
    if (normalized) return normalized;
  }
  return '';
}

function getOpenAiAuthSection(payload = {}) {
  const auth = payload && typeof payload === 'object' ? payload['https://api.openai.com/auth'] : null;
  return auth && typeof auth === 'object' && !Array.isArray(auth) ? auth : {};
}

function getOpenAiProfileSection(payload = {}) {
  const profile = payload && typeof payload === 'object' ? payload['https://api.openai.com/profile'] : null;
  return profile && typeof profile === 'object' && !Array.isArray(profile) ? profile : {};
}

function timestampFromUnixSeconds(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return '';
  const date = new Date(numeric * 1000);
  return Number.isNaN(date.getTime()) ? '' : date.toISOString();
}

function normalizeTimestamp(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString();
  if (typeof value === 'number' && Number.isFinite(value)) {
    const date = new Date(value > 1e11 ? value : value * 1000);
    return Number.isNaN(date.getTime()) ? '' : date.toISOString();
  }
  const date = new Date(String(value || ''));
  return Number.isNaN(date.getTime()) ? '' : date.toISOString();
}

function epochSecondsFromValue(value) {
  if (value === undefined || value === null || value === '') return 0;
  const numeric = Number(value);
  if (Number.isFinite(numeric)) return Math.trunc(numeric > 1e11 ? numeric / 1000 : numeric);
  const parsed = Date.parse(String(value));
  return Number.isFinite(parsed) ? Math.trunc(parsed / 1000) : 0;
}

function encodeBase64UrlJson(value) {
  const json = JSON.stringify(value);
  const bytes = new TextEncoder().encode(json);
  let binary = '';
  for (let index = 0; index < bytes.length; index += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(index, index + 0x8000));
  }
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

function normalizePlanTypeForFileName(planType = '') {
  return normalizeString(planType)
    .split(/[^a-zA-Z0-9]+/)
    .map((part) => part.trim().toLowerCase())
    .filter(Boolean)
    .join('-');
}

function buildCpaAuthFileName(authJson = {}) {
  const email = sanitizeFileName(authJson.email || '', '');
  const planType = normalizePlanTypeForFileName(authJson.plan_type || authJson.chatgpt_plan_type);
  const accountId = sanitizeFileName(authJson.account_id || authJson.chatgpt_account_id || '', '');
  if (email && planType) return `codex-${email}-${planType}.json`;
  if (email) return `codex-${email}.json`;
  if (accountId && planType) return `codex-${accountId}-${planType}.json`;
  if (accountId) return `codex-${accountId}.json`;
  return `codex-${Date.now()}.json`;
}

function buildGuJumpgateSessionStateFromAccount(account = {}) {
  return {
    accessToken: account.accessToken,
    idToken: account.idToken,
    refreshToken: account.refreshToken,
    sessionToken: account.sessionToken,
    email: account.email,
    session: {
      accessToken: account.accessToken,
      idToken: account.idToken,
      refreshToken: account.refreshToken,
      sessionToken: account.sessionToken,
      email: account.email,
      expiresAt: account.expiresAt,
      user: {
        id: account.userId,
        email: account.email,
      },
      account: {
        id: account.accountId,
        planType: account.planType,
      },
    },
  };
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, Math.max(0, Number(ms) || 0)));
}
