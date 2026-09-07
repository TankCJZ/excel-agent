<div align="center">

# 📊 Excel Agent

**Production-ready AI Excel Agent Worker built with Cloudflare Workers, Mastra, OpenAI GPT-5.6 Luna, and Cloudflare Workflows.**

[![Official Website](https://img.shields.io/badge/Official%20Website-excelgen.app-2563eb?style=for-the-badge&logo=googlechrome&logoColor=white)](https://www.excelgen.app)
[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg?style=for-the-badge)](LICENSE)
[![Cloudflare Workers](https://img.shields.io/badge/Cloudflare-Workers-F38020?style=for-the-badge&logo=cloudflare&logoColor=white)](https://workers.cloudflare.com/)
[![OpenAI GPT-5.6](https://img.shields.io/badge/OpenAI-GPT--5.6%20Luna-412991?style=for-the-badge&logo=openai&logoColor=white)](https://developers.cloudflare.com/ai/models/openai/gpt-5.6-luna/)
[![Mastra](https://img.shields.io/badge/Framework-Mastra%20AI-black?style=for-the-badge)](https://mastra.ai)

[English](README.md) | [简体中文](README.zh-CN.md) | [Live Demo (https://www.excelgen.app)](https://www.excelgen.app)

</div>

---

## 🌟 Overview

**Excel Agent** is the open-source backend engine powering [ExcelGen.app](https://www.excelgen.app). It executes multi-turn conversational intelligence, deterministic spreadsheet analysis, automated formula synthesis, and interactive data visualization directly at Cloudflare's edge.

Unlike naive LLM wrappers that paste raw data into model prompts, **Excel Agent** employs a strict **schema-first, deterministic execution pipeline**:
1. **Intelligent Intent & Plan Generation**: The model (`openai/gpt-5.6-luna` via Cloudflare Responses API) inspects workbook structures and outputs a strongly typed `AnalysisPlan`.
2. **Zero-Hallucination Execution**: Calculations (aggregations, sums, groupings, filters, data quality audits) are computed deterministically by the Worker and SheetJS engine—preventing hallucinated figures or broken formulas.
3. **Resilient Long Tasks**: Large workbooks and multi-step mutations run on **Cloudflare Workflows** with automatic retries and checkpointing.
4. **Edge Persistence**: Multi-turn history, context caches, and memory are stored in **Cloudflare D1** and **Cloudflare R2**.

---

## 🚀 Architecture

```text
               User / Web Client (https://www.excelgen.app)
                                 │
                     HTTP / SSE Streaming (POST /api/chat)
                                 ▼
                     ┌──────────────────────┐
                     │  Excel Agent Worker  │
                     └──────────┬───────────┘
                                │
        ┌───────────────────────┼────────────────────────┐
        ▼                       ▼                        ▼
┌──────────────┐      ┌──────────────────┐     ┌──────────────────┐
│ Workers AI   │      │ Cloudflare D1    │     │ Cloudflare R2    │
│ GPT-5.6 Luna │      │ Conversations    │     │ Workbooks        │
│ Responses API│      │ Tasks & Memory   │     │ Cached Contexts  │
└──────────────┘      └──────────────────┘     └──────────────────┘
                                │
                                ▼
                     ┌──────────────────────┐
                     │ Cloudflare Workflows │
                     │  Async Batch Engine  │
                     │  SheetJS Processing  │
                     └──────────────────────┘
```

---

## ✨ Key Features

- **🗣️ Multi-Turn Workbook Dialogue**: Ask follow-up questions about complex workbooks with full revision control and context caching.
- **📈 Deterministic Analytics & Charts**: Produces bar, line, pie, and scatter charts with exact formulas and verified calculations.
- **📑 New Spreadsheet Generation**: Generates clean `.xlsx` spreadsheets from scratch with native styles, conditional formatting, and verified formulas.
- **⚡ OpenAI GPT-5.6 Luna on Cloudflare**: Uses Cloudflare Responses API with direct `env.AI` binding—zero external API keys required in production.
- **🔄 Async Cloudflare Workflows**: Handles large-scale workbook exports and complex transforms with step-level fault tolerance.
- **🛡️ Type-Safe & Zero External Runtime**: Runs natively on Cloudflare Workers edge runtime with TypeScript.

---

## 🛠️ Quick Start

### Prerequisites
- **Node.js**: `>= 22.0.0`
- **Package Manager**: `pnpm` (`npm install -g pnpm`)
- **Cloudflare Account**: [Sign up free](https://dash.cloudflare.com)

### 1. Clone & Install

```bash
git clone https://github.com/TankCJZ/excel-agent.git
cd excel-agent
pnpm install
```

### 2. Environment Setup

Copy `.dev.vars.example` to `.dev.vars`:

```bash
cp .dev.vars.example .dev.vars
```

Edit `.dev.vars`:
```dotenv
CLOUDFLARE_ACCOUNT_ID=your_cloudflare_account_id
CLOUDFLARE_API_TOKEN=your_cloudflare_api_token
```

### 3. Initialize Local Database

Apply the initial D1 migration to your local Miniflare database:

```bash
pnpm db:migrate:local
```

### 4. Start Local Development

```bash
# Option A: Fast mock model for testing (zero token cost)
pnpm dev:mock

# Option B: Real OpenAI GPT-5.6 Luna via Cloudflare
pnpm dev
```

The Worker will start on `http://localhost:8788`. Verify health:

```bash
curl http://localhost:8788/api/health
```

---

## 🧪 Testing

Run the full automated test suite (including model protocol, plan executor, workbook mutations, and formula calculations):

```bash
pnpm test
pnpm typecheck
```

---

## 🚢 Production Deployment

1. **Create Cloudflare Resources**:
   ```bash
   # Create D1 Database
   npx wrangler d1 create excelgen

   # Create R2 Bucket
   npx wrangler r2 bucket create excelgen-media
   ```

2. **Apply Migrations to Remote D1**:
   ```bash
   pnpm db:migrate:remote
   ```

3. **Deploy Worker**:
   ```bash
   pnpm deploy
   ```

---

## 🔗 Official Links & Community

- **Official Product**: [https://www.excelgen.app](https://www.excelgen.app)
- **Bug Reports & Issues**: [GitHub Issues](https://github.com/TankCJZ/excel-agent/issues)
- **Discussions & Feedback**: [GitHub Discussions](https://github.com/TankCJZ/excel-agent/discussions)

---

## 📄 License

Excel Agent is licensed under the [Apache-2.0 License](LICENSE).
