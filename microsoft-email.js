(function attachMicrosoftEmailHelpers(globalScope) {
  const CODE_PATTERN = /\b(\d{6})\b/;

  function normalizeMailboxLabel(mailbox = 'INBOX') {
    return /^junk(?:\s*e-?mail|\s*email)?$/i.test(String(mailbox || '').trim()) ? 'Junk' : 'INBOX';
  }

  function normalizeRecipientAddress(rawValue) {
    if (!rawValue) return '';
    if (typeof rawValue === 'string') {
      return rawValue.trim();
    }
    if (typeof rawValue === 'object') {
      const emailAddress = rawValue.EmailAddress || rawValue.emailAddress || {};
      return String(
        emailAddress.Address
        || emailAddress.address
        || rawValue.Address
        || rawValue.address
        || rawValue.email
        || ''
      ).trim();
    }
    return '';
  }

  function normalizeRecipientList(rawValue) {
    const source = Array.isArray(rawValue)
      ? rawValue
      : (rawValue ? [rawValue] : []);
    const results = [];
    const seen = new Set();
    for (const item of source) {
      const address = normalizeRecipientAddress(item);
      const key = address.toLowerCase();
      if (!key || seen.has(key)) {
        continue;
      }
      seen.add(key);
      results.push(address);
    }
    return results;
  }

  function normalizeMessage(message, mailbox = 'INBOX') {
    const sender = message?.From || message?.from || {};
    const emailAddress = sender?.EmailAddress || sender?.emailAddress || {};
    const recipients = {
      to: normalizeRecipientList(message?.ToRecipients || message?.toRecipients || message?.to),
      cc: normalizeRecipientList(message?.CcRecipients || message?.ccRecipients || message?.cc),
      bcc: normalizeRecipientList(message?.BccRecipients || message?.bccRecipients || message?.bcc),
    };
    recipients.all = [...new Set([...recipients.to, ...recipients.cc, ...recipients.bcc])];
    return {
      mailbox: normalizeMailboxLabel(mailbox || message?.mailbox),
      from: {
        emailAddress: {
          address: String(emailAddress?.Address || emailAddress?.address || '').trim(),
          name: String(emailAddress?.Name || emailAddress?.name || '').trim(),
        },
      },
      subject: String(message?.Subject || message?.subject || '').trim(),
      receivedDateTime: String(message?.ReceivedDateTime || message?.receivedDateTime || '').trim(),
      bodyPreview: String(message?.BodyPreview || message?.bodyPreview || '').trim(),
      body: {
        content: String(message?.Body?.Content || message?.body?.content || '').trim(),
      },
      recipients,
      id: String(message?.Id || message?.id || message?.internetMessageId || '').trim(),
    };
  }

  function normalizeFilterValue(value) {
    return String(value || '').trim().toLowerCase();
  }

  function normalizeRulePatternList(patterns = []) {
    return Array.isArray(patterns) ? patterns : [];
  }

  function extractCodeByRulePatterns(text, patterns = []) {
    const normalizedText = String(text || '');
    for (const pattern of normalizeRulePatternList(patterns)) {
      try {
        const source = String(pattern?.source || '').trim();
        if (!source) {
          continue;
        }
        const flags = String(pattern?.flags || '').replace(/[^dgimsuvy]/g, '');
        const match = normalizedText.match(new RegExp(source, flags));
        if (!match) {
          continue;
        }
        for (let index = 1; index < match.length; index += 1) {
          const candidate = String(match[index] || '').trim();
          if (candidate) {
            return candidate;
          }
        }
        if (String(match[0] || '').trim()) {
          return String(match[0] || '').trim();
        }
      } catch (_) {
        // Runtime rule patterns are user-provided; skip invalid patterns.
      }
    }
    return null;
  }

  function extractVerificationCode(text, options = {}) {
    const source = String(text || '');
    const matchedByRule = extractCodeByRulePatterns(source, options?.codePatterns);
    if (matchedByRule) return matchedByRule;

    const matchCn = source.match(/(?:代码为|验证码[^0-9]*?)[\s：:]*(\d{6})/i);
    if (matchCn) return matchCn[1];

    const matchLoginCode = source.match(/(?:log-?in\s+code|enter\s+this\s+code)[^0-9]{0,24}(\d{6})/i);
    if (matchLoginCode) return matchLoginCode[1];

    const matchEn = source.match(/code(?:\s+is|[\s:])+(\d{6})/i);
    if (matchEn) return matchEn[1];

    const matchStandalone = source.match(CODE_PATTERN);
    return matchStandalone?.[1] || '';
  }

  function getMessageSender(message) {
    return String(
      message?.from?.emailAddress?.address
      || message?.sender?.emailAddress?.address
      || ''
    ).trim();
  }

  function getMessageTimestamp(message) {
    const value = Date.parse(message?.receivedDateTime || message?.createdDateTime || '');
    return Number.isFinite(value) ? value : 0;
  }

  function getMessageSearchText(message) {
    return [
      message?.subject,
      message?.bodyPreview,
      message?.body?.content,
      getMessageSender(message),
    ]
      .map((value) => String(value || ''))
      .join('\n');
  }

  function extractVerificationCodeFromMessages(messages, options = {}) {
    const filterAfterTimestamp = Number(options.filterAfterTimestamp || 0) || 0;
    const senderFilters = (options.senderFilters || []).map(normalizeFilterValue).filter(Boolean);
    const subjectFilters = (options.subjectFilters || []).map(normalizeFilterValue).filter(Boolean);
    const requiredKeywords = (options.requiredKeywords || []).map(normalizeFilterValue).filter(Boolean);
    const excludedCodes = new Set((options.excludeCodes || []).map((value) => String(value || '').trim()).filter(Boolean));
    const hasExplicitFilters = senderFilters.length > 0 || subjectFilters.length > 0 || requiredKeywords.length > 0;

    const sortedMessages = (Array.isArray(messages) ? messages : [])
      .map((raw) => normalizeMessage(raw, raw?.mailbox))
      .sort((left, right) => getMessageTimestamp(right) - getMessageTimestamp(left));

    for (const message of sortedMessages) {
      const receivedAt = getMessageTimestamp(message);
      if (receivedAt && receivedAt < filterAfterTimestamp) {
        continue;
      }

      const sender = normalizeFilterValue(getMessageSender(message));
      const subject = normalizeFilterValue(message?.subject);
      const preview = normalizeFilterValue(message?.bodyPreview);
      const searchText = normalizeFilterValue(getMessageSearchText(message));
      const code = extractVerificationCode(getMessageSearchText(message), {
        codePatterns: options.codePatterns,
      });
      if (!code || excludedCodes.has(code)) {
        continue;
      }

      const senderMatched = senderFilters.length === 0
        ? false
        : senderFilters.some((filter) => sender.includes(filter) || preview.includes(filter) || searchText.includes(filter));
      const subjectMatched = subjectFilters.length === 0
        ? false
        : subjectFilters.some((filter) => subject.includes(filter) || preview.includes(filter) || searchText.includes(filter));
      const keywordMatched = requiredKeywords.length === 0
        ? false
        : requiredKeywords.some((filter) => preview.includes(filter) || searchText.includes(filter));
      if (hasExplicitFilters && !senderMatched && !subjectMatched && !keywordMatched) {
        continue;
      }

      return {
        code,
        emailTimestamp: receivedAt || Date.now(),
        messageId: message?.id || null,
        sender: getMessageSender(message),
        subject: String(message?.subject || ''),
        mailbox: message?.mailbox || 'INBOX',
        message,
      };
    }

    return null;
  }

  const api = {
    CODE_PATTERN,
    extractVerificationCode,
    extractVerificationCodeFromMessages,
    getMessageSender,
    getMessageTimestamp,
    normalizeMailboxLabel,
    normalizeMessage,
  };

  globalScope.MultiPageMicrosoftEmail = api;

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api;
  }
})(typeof globalThis !== 'undefined' ? globalThis : this);
