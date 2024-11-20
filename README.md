## 文件扫描助手

### 简介

文件扫描助手是一款基于Vite + Vue 3 + Electron技术栈开发的跨平台桌面应用程序。它提供了强大的文件内容搜索功能，支持Word、Excel、PDF、PPT等常见办公文档格式。用户可以通过关键词快速定位到包含特定内容的文件，极大地提高了文件管理和查找效率。

### 功能特点

- 多格式支持：支持Word（.doc, .docx）、Excel（.xls, .xlsx）、PDF（.pdf）、PPT（.ppt, .pptx）等多种文件格式的全文本搜索。
- 高效搜索：使用先进的文本处理算法，实现快速准确的搜索体验。
- 跨平台兼容：支持Windows、macOS和Linux操作系统，满足不同用户的使用需求。
- 简洁界面：采用现代简约设计风格，提供友好直观的操作体验。
- 配置灵活：允许用户自定义搜索路径、关键字高亮等个性化设置。

### 技术栈

前端框架：Vue 3
构建工具：Vite
桌面应用框架：Electron
样式框架：unocss

### Install

```bash
$ pnpm install
```

### Development

```bash
$ pnpm dev
```

### Build

```bash
# For windows
$ pnpm build:win

# For macOS
$ pnpm build:mac

# For Linux
$ pnpm build:linux
```
