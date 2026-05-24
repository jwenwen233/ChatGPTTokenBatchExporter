(function attachTokenExporterAuthContent() {
  if (window.__TOKEN_EXPORTER_AUTH_READY__) return;
  window.__TOKEN_EXPORTER_AUTH_READY__ = true;
  const HANDLED_TYPES = new Set([
    'TK_AUTH_PING',
    'TK_AUTH_STATE',
    'TK_START_LOGIN',
    'TK_FILL_EMAIL',
    'TK_FILL_PASSWORD',
    'TK_USE_EMAIL_CODE',
    'TK_FILL_CODE',
  ]);

  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    if (!message || !HANDLED_TYPES.has(message.type)) return false;
    Promise.resolve(handleMessage(message))
      .then((result) => sendResponse({ ok: true, ...(result || {}) }))
      .catch((error) => sendResponse({ ok: false, error: error?.message || String(error || '') }));
    return true;
  });

  async function handleMessage(message) {
    const payload = message.payload || {};
    switch (message.type) {
      case 'TK_AUTH_PING':
        return { ready: true };
      case 'TK_AUTH_STATE':
        return inspectState();
      case 'TK_START_LOGIN':
        return startLogin();
      case 'TK_FILL_EMAIL':
        return fillEmail(payload.email);
      case 'TK_FILL_PASSWORD':
        return fillPassword(payload.password);
      case 'TK_USE_EMAIL_CODE':
        return useEmailCodeLogin();
      case 'TK_FILL_CODE':
        return fillCode(payload.code);
      default:
        return {};
    }
  }

  function text() {
    return String(document.body?.innerText || '').replace(/\s+/g, ' ').trim();
  }

  function isVisible(el) {
    if (!el) return false;
    const style = getComputedStyle(el);
    const rect = el.getBoundingClientRect();
    return style.display !== 'none'
      && style.visibility !== 'hidden'
      && rect.width > 0
      && rect.height > 0;
  }

  function findInput(selectors) {
    for (const selector of selectors) {
      const found = Array.from(document.querySelectorAll(selector)).find(isVisible);
      if (found) return found;
    }
    return null;
  }

  function findEmailInput() {
    return findInput([
      'input[type="email"]',
      'input[name="email"]',
      'input#email-input',
      'input[autocomplete="username"]',
      'input[placeholder*="email" i]',
      'input[aria-label*="email" i]',
    ]);
  }

  function findPasswordInput() {
    return findInput([
      'input[type="password"]',
      'input[name="password"]',
      'input[autocomplete="current-password"]',
      'input[placeholder*="password" i]',
      'input[aria-label*="password" i]',
    ]);
  }

  function isCodeLikePage() {
    return /检查你的收件箱|验证码|verification\s*code|one[-\s]*time|passcode|check\s+your\s+inbox|sent\s+to/i.test(text());
  }

  function findTextEntryInputs() {
    return Array.from(document.querySelectorAll('input, textarea, [contenteditable="true"]'))
      .filter(isVisible)
      .filter((input) => {
        if (input.matches?.('[contenteditable="true"]')) return true;
        const type = String(input.type || 'text').toLowerCase();
        return !['hidden', 'email', 'password', 'checkbox', 'radio', 'submit', 'button', 'file'].includes(type);
      })
      .filter((input) => input !== findEmailInput() && input !== findPasswordInput());
  }

  function findCodeInputs() {
    const split = Array.from(document.querySelectorAll('input[maxlength="1"], input[data-testid*="code"]')).filter(isVisible);
    if (split.length >= 6) {
      return { type: 'split', inputs: split.slice(0, 6) };
    }
    const active = document.activeElement;
    if (isCodeLikePage() && active && isVisible(active) && findTextEntryInputs().includes(active)) {
      return { type: 'single', input: active };
    }
    const single = findInput([
      'input[autocomplete="one-time-code"]',
      'input[name="code"]',
      'input[name*="code" i]',
      'input[name="otp"]',
      'input[name*="otp" i]',
      'input[name*="verification" i]',
      'input[inputmode="numeric"]',
      'input[type="tel"][maxlength="6"]',
      'input[type="text"][maxlength="6"]',
      'input[placeholder*="code" i]',
      'input[placeholder*="验证码"]',
      'input[placeholder*="一次性"]',
      'input[aria-label*="code" i]',
      'input[aria-label*="验证码"]',
      'input[aria-label*="一次性"]',
    ]);
    if (single) return { type: 'single', input: single };

    if (!isCodeLikePage()) {
      return null;
    }
    const fallback = findTextEntryInputs()[0];
    return fallback ? { type: 'single', input: fallback } : null;
  }

  function delay(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  function buttonText(el) {
    return String(el?.innerText || el?.value || el?.getAttribute?.('aria-label') || '').replace(/\s+/g, ' ').trim();
  }

  function isSocialLoginButton(el) {
    return /google|apple|microsoft|github|sso|phone|电话|手机号|使用\s*(google|apple|电话|手机)/i.test(buttonText(el));
  }

  function findButtons(options = {}) {
    return Array.from(document.querySelectorAll('button, a, [role="button"], input[type="submit"]'))
      .filter(isVisible)
      .filter((el) => !options.excludeSocial || !isSocialLoginButton(el));
  }

  function findButtonByText(pattern, options = {}) {
    return findButtons(options).find((el) => pattern.test(buttonText(el)));
  }

  function findSubmitButton() {
    return findButtonByText(/continue|next|submit|log\s*in|sign\s*in|verify|继续|下一步|登录|验证/i, { excludeSocial: true })
      || Array.from(document.querySelectorAll('button[type="submit"], input[type="submit"]')).filter(isVisible).find((el) => !isSocialLoginButton(el))
      || Array.from(document.querySelectorAll('button')).filter(isVisible).filter((el) => !isSocialLoginButton(el)).at(-1)
      || null;
  }

  function findEmailContinueButton(input) {
    const pattern = /^(continue|next|继续|下一步)$/i;
    const formButton = Array.from(input.closest('form')?.querySelectorAll('button, input[type="submit"]') || [])
      .filter(isVisible)
      .filter((el) => !isSocialLoginButton(el))
      .find((el) => pattern.test(buttonText(el)) || /continue|next|继续|下一步/i.test(buttonText(el)));
    if (formButton) return formButton;

    const inputRect = input.getBoundingClientRect();
    return findButtons({ excludeSocial: true })
      .filter((el) => {
        const rect = el.getBoundingClientRect();
        return rect.top >= inputRect.bottom - 4;
      })
      .sort((a, b) => a.getBoundingClientRect().top - b.getBoundingClientRect().top)
      .find((el) => pattern.test(buttonText(el)) || /continue|next|继续|下一步/i.test(buttonText(el)))
      || findSubmitButton();
  }

  function findEmailCodeButton() {
    return findButtonByText(/one[-\s]*time|passcode|verification\s*code|email\s*(me|a)?\s*code|send.*code|use.*code|sign\s*in.*code|验证码|一次性|邮箱.*码|发送.*码/i);
  }

  function findOtherMethodButton() {
    return findButtonByText(/another\s*(way|method)|other\s*(way|method|options)|try\s*another|不同方式|其他方式|更多方式|换一种/i);
  }

  function findLoginButton() {
    return findButtonByText(/^(log\s*in|sign\s*in|登录)$/i, { excludeSocial: true })
      || findButtonByText(/\b(log\s*in|sign\s*in)\b|登录/i, { excludeSocial: true });
  }

  function isLikelyLoggedInChatPage(bodyText) {
    if (!/chatgpt\.com|chat\.openai\.com/i.test(location.hostname)) return false;
    if (/auth|login/i.test(location.href.toLowerCase())) return false;
    if (findLoginButton() || findEmailInput() || findPasswordInput() || findCodeInputs()) return false;
    if (!String(bodyText || '').trim()) return false;
    return Boolean(document.querySelector([
      '[data-testid="composer-textarea"]',
      '[data-testid="prompt-textarea"]',
      '#prompt-textarea',
      'textarea[placeholder*="Message" i]',
      '[contenteditable="true"][data-testid*="composer" i]',
    ].join(','))) || /new chat|message chatgpt|chatgpt can make mistakes|upgrade plan|library|explore gpts|新聊天|消息 chatgpt/i.test(bodyText);
  }

  function setValue(input, value) {
    input.focus();
    input.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Backspace' }));
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
    const textAreaSetter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')?.set;
    const normalized = String(value || '');
    if (input.matches?.('[contenteditable="true"]')) {
      input.textContent = normalized;
    } else if (input instanceof HTMLTextAreaElement && textAreaSetter) {
      textAreaSetter.call(input, normalized);
    } else if (setter) {
      setter.call(input, normalized);
    } else {
      input.value = normalized;
    }
    input.dispatchEvent(new InputEvent('beforeinput', { bubbles: true, inputType: 'insertText', data: normalized }));
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
    input.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: normalized.slice(-1) || '0' }));
  }

  function click(el) {
    if (!el) return false;
    el.scrollIntoView?.({ block: 'center', inline: 'center' });
    el.click();
    return true;
  }

  function inspectState() {
    const bodyText = text();
    const lowerUrl = location.href.toLowerCase();
    const blocked = /captcha|verify you are human|cloudflare|unusual activity|suspicious|too many attempts|安全检查|人机|真人|验证你是真人/i.test(bodyText);
    const codeError = /invalid\s*(code|verification|passcode)|incorrect\s*(code|verification|passcode)|expired\s*(code|verification|passcode)|code\s*(is\s*)?(invalid|incorrect|expired)|验证码.*(错误|无效|过期)|错误的验证码|验证码不正确|验证码已过期/i.test(bodyText);
    if (blocked) {
      return { page: 'blocked', blocked: true, blockedReason: 'captcha/security check', url: location.href };
    }
    if (findCodeInputs()) return { page: 'code', codeError, url: location.href };
    if (findPasswordInput()) return { page: 'password', url: location.href };
    if (findEmailInput()) return { page: 'email', url: location.href };
    if (findLoginButton()) return { page: 'entry', url: location.href };
    if (/chatgpt\.com|chat\.openai\.com/i.test(location.hostname) && /\/auth\/login/i.test(location.pathname)) {
      return { page: 'entry', url: location.href };
    }
    if (isLikelyLoggedInChatPage(bodyText)) {
      return { page: 'logged-in', url: location.href };
    }
    if (/log\s*in|sign\s*in|登录|continue/i.test(bodyText)) return { page: 'entry', url: location.href };
    return { page: 'unknown', url: location.href };
  }

  async function startLogin() {
    const button = findLoginButton();
    if (button) {
      click(button);
      return { clicked: true };
    }
    location.href = 'https://chatgpt.com/';
    return { navigated: true };
  }

  async function fillEmail(email) {
    const input = findEmailInput();
    if (!input) throw new Error('未找到邮箱输入框。');
    setValue(input, email);
    await delay(80);
    const button = findEmailContinueButton(input);
    if (!button) throw new Error('未找到邮箱输入框下方的继续按钮。');
    click(button);
    return { filled: true };
  }

  async function fillPassword(password) {
    const input = findPasswordInput();
    if (!input) throw new Error('未找到密码输入框。');
    setValue(input, password);
    click(findSubmitButton());
    return { filled: true };
  }

  async function useEmailCodeLogin() {
    if (findCodeInputs()) return { alreadyCodePage: true };
    const direct = findEmailCodeButton();
    if (direct) {
      click(direct);
      return { clicked: true, method: 'email-code' };
    }
    const other = findOtherMethodButton();
    if (other) {
      click(other);
      return { clicked: true, method: 'other-methods' };
    }
    throw new Error('未找到“邮箱验证码/一次性代码”按钮。');
  }

  async function fillCode(code) {
    const normalized = String(code || '').trim();
    if (!/^[A-Za-z0-9]{4,12}$/.test(normalized)) throw new Error('验证码格式不正确。');
    const target = findCodeInputs();
    if (!target) throw new Error('未找到验证码输入框。');
    if (target.type === 'split') {
      target.inputs.forEach((input, index) => setValue(input, normalized[index] || ''));
    } else {
      setValue(target.input, normalized);
    }
    await delay(150);
    click(findSubmitButton());
    return { filled: true };
  }
})();
