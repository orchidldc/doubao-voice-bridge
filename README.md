# DouBao Voice Bridge

[![English README](https://img.shields.io/badge/README-English-blue)](README.en.md)

豆包跨屏输入桥接工具。手机端用豆包或任意语音输入法把文字写入飞书/Lark Docx 文档，Windows 端读取文档里新增的文字，并粘贴到当前输入框。

这个公开仓库是展示和发布仓库，不包含应用源码、真实配置、日志、截图、飞书文档链接或任何账号密钥。

## 下载

请从 GitHub Releases 下载 Windows x64 版本：

[下载 DouBao Voice Bridge](https://github.com/orchidldc/doubao-voice-bridge/releases/latest)

当前发布包名：

```text
DouBaoVoiceBridge-v0.2.0-windows-x64.zip
```

发布包中包含：

- `DouBaoVoiceBridge.exe`
- `config.example.json`
- `README-USER.md`
- `RELEASE_NOTES.txt`
- `THIRD_PARTY_NOTICES.txt`
- `DISCLAIMER.txt`
- `NOTICE.txt`

## 它解决什么问题

有些桌面输入场景不方便直接使用手机语音输入。这个工具把飞书 Docx 文档当作中转：

1. 手机端打开一个专用飞书 Docx 文档。
2. 你用豆包或手机输入法把语音转成文字写进去。
3. 电脑端运行 DouBao Voice Bridge。
4. 把光标放到目标输入框。
5. 按 `F8` 建立 baseline。
6. 工具只粘贴 `F8` 之后新增的文字。

它不做语音识别，不调用豆包 API，不抓浏览器 DOM，不截图，不上传本地数据。

## 图形界面

`v0.2.0` 开始，主程序是图形化软件：

- 保存配置
- 检查飞书连接
- 单次读取
- 测试粘贴
- 启动监听
- 停止监听
- 打开配置
- 打开日志目录

底层桥接 CLI 已嵌入 GUI 资源中，用户不需要管理多个 exe。

## 快速开始

1. 解压 `DouBaoVoiceBridge-v0.2.0-windows-x64.zip`。
2. 双击 `DouBaoVoiceBridge.exe`。
3. 填写你的飞书 Docx 链接。
4. 确认本机 `lark-cli` 已登录：

```powershell
lark-cli auth status --json
lark-cli auth login --scope "docx:document:readonly offline_access"
```

5. 点击“保存配置”。
6. 点击“检查连接”。
7. 点击“启动监听”。
8. 把光标放到目标输入框。
9. 按 `F8` 建立 baseline。
10. 手机端继续往飞书文档输入，电脑端会粘贴新增文字。

## 热键

- `F8`: 开始监听并建立 baseline
- `F9`: 暂停
- `F10`: 重设 baseline
- `F12`: 退出桥接进程

## 开头空行处理

默认启用智能开头空行处理。若两段语音输入间隔超过阈值，工具会去掉飞书段落带来的片段开头换行，避免每次粘贴都先出现一个空行。短时间连续输入仍会保留换行。

界面中的“开头换行阈值秒”默认是 `2.0`。

## 目标窗口模式

默认模式是 `any`，适合 VS Code、浏览器输入框、聊天窗口等场景，光标在哪里就粘贴到哪里。它更顺手，但也要求你确认当前光标位置正确。

可选模式：

- `locked`: 只粘贴到按 `F8` 时捕获的同一窗口。
- `process`: 允许同一进程的其他窗口。
- `any`: 粘贴到当前前台窗口，最顺手，误粘风险也最高。

## 隐私和安全

- 不要上传真实 `doubao_voice_bridge_config.json`。
- 不要公开真实飞书文档链接。
- 不要公开 token、App Secret、日志或含隐私内容的截图。
- 不要在密码框、浏览器地址栏、支付表单、后台管理危险输入框里使用。

详见：

- [隐私说明](PRIVACY.md)
- [安全策略](SECURITY.md)
- [免责声明](DISCLAIMER.md)

## 文档

- [快速开始](docs/quick-start.md)
- [飞书/Lark 设置](docs/setup-feishu.md)
- [故障排查](docs/troubleshooting.md)
- [设计说明](docs/design.md)

## 权利声明

本仓库不授予应用源码或编译程序的开源许可证。详见 [NOTICE](NOTICE)。

