<div align="center">

# 📊 Excel Agent

**基于 Cloudflare Workers、Mastra、Cloudflare Workers AI 与 Cloudflare Workflows 构建的企业级开源 AI 表格智能体 Worker**

[![官方站点](https://img.shields.io/badge/官方网站-excelgen.app-2563eb?style=for-the-badge&logo=googlechrome&logoColor=white)](https://www.excelgen.app)
[![开源协议](https://img.shields.io/badge/License-Apache_2.0-blue.svg?style=for-the-badge)](LICENSE)
[![Cloudflare Workers](https://img.shields.io/badge/Cloudflare-Workers-F38020?style=for-the-badge&logo=cloudflare&logoColor=white)](https://workers.cloudflare.com/)
[![Workers AI](https://img.shields.io/badge/AI-Cloudflare%20Workers%20AI-F38020?style=for-the-badge&logo=cloudflare&logoColor=white)](https://developers.cloudflare.com/workers-ai/)
[![Mastra](https://img.shields.io/badge/开发框架-Mastra%20AI-black?style=for-the-badge)](https://mastra.ai)

[English](README.md) | [简体中文](README.zh-CN.md) | [在线演示 (https://www.excelgen.app)](https://www.excelgen.app)

</div>

---

## 🌟 项目简介

**Excel Agent** 是 [ExcelGen.app](https://www.excelgen.app) 官方采用的开源核心后端智能体引擎。它直接运行在 Cloudflare 全球边缘网络上，提供多轮智能表格对话、确定性数据统计分析、公式自动推导以及交互式图表生成。

与市面上直接将大批杂乱数据喂给大语言模型并容易产生幻觉的方案不同，**Excel Agent** 采用严格的**结构驱动、确定性计算管道**：
1. **结构感知与意图规划**：大模型（借助 Cloudflare Workers AI 驱动流式推理与工具调用）解析表格真实 Schema，仅生成类型安全的 `AnalysisPlan` 执行计划。
2. **零幻觉执行**：数值求和、平均值、分组聚合、数据清洗、公式计算全由 Worker 内部确定的计算执行器和 SheetJS 完成，杜绝模型瞎编数字与错误公式。
3. **高可用长任务调度**：大型工作簿的复杂生成与导出任务由 **Cloudflare Workflows** 承载，具备多步重试与状态断点续跑能力。
4. **边缘存储持久化**：多轮会话记录与工作簿 Context 缓存持久化保存在 **Cloudflare D1** 与 **Cloudflare R2**。

---

## 🚀 架构设计

```text
               客户端 / Web 应用 (https://www.excelgen.app)
                                 │
                     HTTP / SSE 流式通信 (POST /api/chat)
                                 ▼
                     ┌──────────────────────┐
                     │  Excel Agent Worker  │
                     └──────────┬───────────┘
                                │
        ┌───────────────────────┼────────────────────────┐
        ▼                       ▼                        ▼
┌──────────────────────┐ ┌──────────────────┐     ┌──────────────────┐
│ Cloudflare Workers AI│ │ Cloudflare D1    │     │ Cloudflare R2    │
│ 高性能大模型与工具调用 │ │ 会话历史、状态   │     │ 工作簿与文件     │
│ 流式推理 (Streaming) │ │ 任务与上下文     │     │ Context 缓存     │
└──────────────────────┘ └──────────────────┘     └──────────────────┘
                                │
                                ▼
                     ┌──────────────────────┐
                     │ Cloudflare Workflows │
                     │ 异步长时间工作流     │
                     │ SheetJS 格式与处理   │
                     └──────────────────────┘
```

---

## ✨ 核心特性

- **🗣️ 多轮工作簿智能对话**：围绕同一份 Excel 文件连续追问，支持表格版本回溯与增量分析。
- **📈 确定性统计与图表**：精准生成柱状图、折线图、饼图、散点图，且指标结果经代码严格核算。
- **📑 智能从零生成表格**：一键生成符合格式规范的全新 `.xlsx` 工作簿，内置带样式的真实公式。
- **⚡ 原生集成 Cloudflare Workers AI**：借助原生 `env.AI` 绑定直接调度边缘大语言模型，生产环境零外部 API 密钥。
- **🔄 Cloudflare Workflows 长任务**：支持异步导出大型报表，具备任务容灾和断点续传。
- **🛡️ 纯 TypeScript 与边缘原生**：无需配置 Python 环境或重量级虚拟机，开箱即用。

---

## 🛠️ 快速启动

### 环境要求
- **Node.js**：`>= 22.0.0`
- **包管理器**：`pnpm` (`npm install -g pnpm`)
- **Cloudflare 账号**：[免费注册](https://dash.cloudflare.com)

### 1. 克隆并安装依赖

```bash
git clone https://github.com/TankCJZ/excel-agent.git
cd excel-agent
pnpm install
```

### 2. 配置本地环境变量

复制 `.dev.vars.example` 为 `.dev.vars`：

```bash
cp .dev.vars.example .dev.vars
```

填写你的 Cloudflare 凭据：
```dotenv
CLOUDFLARE_ACCOUNT_ID=your_cloudflare_account_id
CLOUDFLARE_API_TOKEN=your_cloudflare_api_token
```

### 3. 初始化本地数据库

在本地 Miniflare 环境执行初始迁移：

```bash
pnpm db:migrate:local
```

### 4. 启动本地开发服务

```bash
# 模式 A：使用内置 Mock 模型（无需调用真实大模型，0 推理费用，极速测试）
pnpm dev:mock

# 模式 B：连接 Cloudflare Workers AI 运行
pnpm dev
```

服务将在 `http://localhost:8788` 启动。测试健康检查：

```bash
curl http://localhost:8788/api/health
```

---

## 🧪 执行自动化测试

运行全部单元与集成测试（包含模型协议、执行计划器、表格变更、公式校验等）：

```bash
pnpm test
pnpm typecheck
```

---

## 🚢 部署上线

1. **在 Cloudflare 创建资源**：
   ```bash
   # 创建 D1 数据库
   npx wrangler d1 create excelgen

   # 创建 R2 存储桶
   npx wrangler r2 bucket create excelgen-media
   ```

2. **推送远端数据库迁移**：
   ```bash
   pnpm db:migrate:remote
   ```

3. **部署 Worker**：
   ```bash
   pnpm deploy
   ```

---

## 🔗 官方站点与社区

- **官方产品**：[https://www.excelgen.app](https://www.excelgen.app)
- **问题反馈**：[GitHub Issues](https://github.com/TankCJZ/excel-agent/issues)
- **讨论社区**：[GitHub Discussions](https://github.com/TankCJZ/excel-agent/discussions)

---

## 📄 开源许可证

本项目基于 [Apache-2.0 开源许可证](LICENSE) 授权。
