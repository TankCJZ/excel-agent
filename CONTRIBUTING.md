# Contributing to Excel Agent

Thank you for your interest in contributing to **Excel Agent**! We welcome bug reports, feature requests, documentation improvements, and code contributions.

## Development Setup

1. **Fork and Clone**:
   ```bash
   git clone https://github.com/TankCJZ/excel-agent.git
   cd excel-agent
   ```

2. **Install Dependencies**:
   ```bash
   pnpm install
   ```

3. **Configure Local Environment**:
   ```bash
   cp .dev.vars.example .dev.vars
   ```

4. **Start Development Server**:
   ```bash
   # Run with mock planner (no API key needed):
   pnpm dev:mock

   # Or run with Cloudflare Workers AI models:
   pnpm dev
   ```

5. **Run Tests**:
   ```bash
   pnpm test
   pnpm typecheck
   ```

## Pull Request Guidelines

1. Create a descriptive feature branch from `main` (`git checkout -b feat/my-feature`).
2. Make sure all unit tests and typechecks pass (`pnpm test` & `pnpm typecheck`).
3. Add unit tests for new capabilities or bug fixes.
4. Open a pull request against `main`.

## Code of Conduct

Please review and adhere to our [Code of Conduct](CODE_OF_CONDUCT.md).
