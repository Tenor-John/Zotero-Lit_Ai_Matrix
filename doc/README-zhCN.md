# Zotero Lit AI Matrix

<div align="center">

[English](../README.md) | [简体中文](README-zhCN.md)

</div>

一个面向科研阅读流程的 Zotero 智能文献矩阵插件。

它把 Zotero 元数据（标题、作者、期刊、年份、标签）与结构化 AI 笔记整合到一个可检索、可排序、可导出的矩阵页面中。

## 功能特色

- 智能文献矩阵视图
  - 直接在 Zotero 内嵌页面使用。
  - 同时展示基础元数据列与 AI 结构化字段列。
- AI 笔记自动解析
  - 自动解析 note 中的结构化字段并写入矩阵缓存。
  - 适配 Zotero GPT 一类的笔记生成流程。
- GitHub 风格热图
  - 展示每日使用频率。
  - 点击某天方块可按天过滤矩阵，再点一次取消。
- 一键跳转阅读
  - 点击标题直接跳转 PDF 阅读界面。
- 提效工具
  - 全库重建缓存、批量阅读状态更新、CSV/Markdown 导出。

## 安装方式

### 方式 1：正式版（推荐）

1. 打开 GitHub Releases 下载最新 `.xpi`。
2. Zotero：`工具` -> `插件` -> 右上角齿轮 -> `Install Plugin From File...`。
3. 选择 `.xpi` 安装并重启 Zotero。

仓库地址：<https://github.com/Tenor-John/Zotero-Lit_Ai_Matrix>

### 方式 2：开发模式

```bash
npm install
npm start
```

请正确配置 `.env`：

- `ZOTERO_PLUGIN_ZOTERO_BIN_PATH`
- `ZOTERO_PLUGIN_PROFILE_PATH`

## 使用教程

### 1. 打开矩阵页面

可通过以下入口打开：

- 工具栏 AI Matrix 图标
- 文件菜单中的矩阵入口

### 2. 生成可解析的 AI 笔记

建议字段如下（每个字段单独一行，`::` 分隔）：

- `领域基础知识::`
- `研究背景::`
- `作者的问题意识::`
- `研究意义::`
- `研究结论::`
- `未来研究方向提及::`
- `未来研究方向思考::`

### 3. 更新矩阵缓存

- 选中文献右键：重建矩阵缓存
- 文件菜单：重建全库矩阵缓存

### 4. 筛选与分析

- 支持关键词、状态、年份、期刊、标签筛选
- 支持列头排序
- 点击标题打开 PDF（无 PDF 时定位条目）

### 5. 热图按天定位

- 点击有数据的日期方块：矩阵过滤为当天文献
- 再次点击同一天：恢复全部矩阵内容

## 与 Zotero GPT 协同建议

1. 先对文献生成结构化 note。
2. 保证字段名与矩阵字段一致。
3. 执行缓存重建。
4. 在矩阵中做检索、对比和导出。

## 常见问题

### 为什么列标题不显示？

- 确认安装的是最新 release。
- 查看矩阵右上角 `UI:xxxx` 版本号。
- 若版本号较旧，请重新安装最新 `.xpi` 并重启 Zotero。

### 为什么 `npm start` 提示找不到 Zotero？

检查 `.env`：

- `ZOTERO_PLUGIN_ZOTERO_BIN_PATH` 指向 Zotero 可执行文件
- `ZOTERO_PLUGIN_PROFILE_PATH` 指向 Zotero profile 目录

## 隐私说明

- 插件主要处理本地 Zotero 条目与 note 数据。
- 不会主动上传你的文献库内容。
- 是否调用外部 AI 服务取决于你自己的 GPT 插件配置。

## 构建与发布

```bash
npm run build
npm run release
```

建议发布流程：

1. 在功能分支开发和测试
2. 通过 PR 合并到 `main`
3. 使用语义化版本发布（例如 `v0.1.5`）

## 许可证

AGPL-3.0-or-later

## 作者

Tenor-John
