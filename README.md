# autoWeComLite

跨平台微信/企业微信自动化助手

## 项目简介

autoWeComLite 是一个基于 wxPython + pyautogui 的桌面自动化工具，支持 Windows 和 macOS，专注于微信/企业微信的消息自动发送。

## 主要特性
- 跨平台（Windows/macOS）
- 高鲁棒性自动化（窗口/控件名判据）
- 简洁 UI，支持联系人、消息输入与发送
- 日志与异常反馈

## 依赖安装

```bash
pip install -r requirements.txt
```

## 运行方式

```bash
python main.py
```

## 目录结构（建议）
```
autoWeComLite/
├── core/           # 核心业务逻辑
├── ui/             # UI 相关代码
├── automation/     # 自动化实现
├── config/         # 配置文件
├── assets/         # 静态资源
├── requirements.txt
├── README.md
├── main.py         # 程序入口
```

## 平台适配说明
- Windows 下依赖 pywinauto 获取控件名
- macOS 下可用 AppleScript 辅助窗口/控件名判据

## 许可证
MIT 