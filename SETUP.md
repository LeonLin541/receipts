# 收据 App 设置步骤

## 1. 在 Microsoft Entra 登记 App（约 5 分钟，需要管理员）

1. 打开 https://entra.microsoft.com → **应用注册 App registrations** → **新注册 New registration**
2. 名称：`JML Receipts`
3. 支持的账户类型：**仅此组织目录中的账户（单租户）**
4. 重定向 URI：平台选 **单页应用程序 (SPA)**，地址填 App 的网址（例如 `https://xxx.github.io/receipts/`）
   - 再加一条 `http://localhost:8080/`，用于在电脑上测试
5. 注册完成后，在「概述」页复制：
   - **应用程序(客户端) ID** → 填进 `config.js` 的 `MS_CLIENT_ID`
   - **目录(租户) ID** → 填进 `config.js` 的 `MS_TENANT_ID`
6. **API 权限** → 添加 → Microsoft Graph → 委托的权限 → 勾选 `Files.ReadWrite` 和 `User.Read` → 点 **代表 JML 授予管理员同意**

App 只用你自己的登录身份访问你自己的 OneDrive，没有服务器，也不保存任何密码。

## 2. Claude API key

1. https://console.anthropic.com → API Keys → Create Key，名称 `receipts-phone`
2. 建议在 Limits 里设置每月消费上限（例如 £10），每张收据大约 2–3 便士
3. 第一次打开 App 时粘贴到「设置」里。key 只保存在这台手机上

## 3. 在安卓手机上安装

1. 用 Chrome 打开 App 网址 → 用公司 Microsoft 账号登录
2. Chrome 菜单 ⋮ → **添加到主屏幕 / 安装应用**
3. 之后点桌面图标就能直接拍照

## 修改卡号、物业

编辑 OneDrive 里的 `0 Receipts Inbox/_settings/receipt-config.json`，然后在 App 设置里点「刷新卡号/目录设置」。电脑端的收据处理规则也读这个文件。
