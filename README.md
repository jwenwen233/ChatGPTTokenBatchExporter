# ChatGPT Token Refresh Helper

独立 Chrome MV3 扩展，用于从 Cockpit 或 CliProxy 拉取 ChatGPT auth JSON，后台实测 `accessToken`，对失效账号走 ChatGPT 网页邮箱验证码登录，读取新的 `/api/auth/session`，再写回 Cockpit 或 CliProxy。

## 加载方式

1. 打开 `chrome://extensions`
2. 打开右上角“开发者模式”
3. 点击“加载已解压的扩展程序”
4. 选择扩展目录

## 本地工具对接

### Cockpit

选择 Cockpit 根目录：

```text
~/.antigravity_cockpit
```

然后点击“扫描文件”。扫描后会列出 Cockpit 账号池里的账号，并自动实测 token 是否可用。勾选需要刷新的账号后点击“刷新并写回”。

### CliProxy

选择 `CliProxy`，填写：

```text
CPA 地址：http://127.0.0.1:8317
管理密钥：你的 CliProxy management key
```

然后点击“拉取账号”。刷新成功后会通过 CliProxy 管理 API 写回新的 CPA JSON，不读写本机 CliProxy 文件。

## 查信配置

刷新失效账号时，插件会按当前 ChatGPT 账号邮箱查收验证码。

### Outlook/Hotmail

固定使用“本地助手”：

```text
助手地址：http://127.0.0.1:17373
邮箱格式：email----password----clientId----refreshToken
```

启动本地助手：

```bash
./start-hotmail-helper.command
```

`password` 只做本地记录，实际查信使用 `clientId + refreshToken`。

### LuckMail

填写 LuckMail `API Key` 和 `Base URL`。插件会按当前 ChatGPT 账号邮箱查找已购邮箱 token，再读取验证码邮件。

### Cloudflare Temp Email

填写 Temp API 和 Admin Auth。一般选择“按 accessToken 邮箱”，备用邮箱留空。

## Session JSON 转 CPA JSON

把 `https://chatgpt.com/api/auth/session` 返回的 JSON 粘贴到“Session JSON 转 CPA JSON”，点击转换即可下载一个 CPA JSON。

## 注意

- 后台实测 `accessToken` 默认开启，会真实请求 ChatGPT 后端接口；即使 JWT 过期时间未到，也可能因为服务端撤销而返回不可用。
- 刷新账号前后会清理 ChatGPT / OpenAI 登录态，避免用上一个账号的缓存覆盖当前账号。
- 遇到人机验证、二次验证、风控页时，该账号会标记失败，需要人工处理。
- 没有可查收验证码的邮箱配置时，失效账号无法自动刷新。
