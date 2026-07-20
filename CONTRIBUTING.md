# Contributing

Thank you for your interest in contributing to `@bcibibi/docx-html-converter`!

## Getting Started

1. Fork the repository
2. Clone your fork:
   ```bash
   git clone https://github.com/<your-username>/docx-html-converter.git
   cd docx-html-converter
   ```
3. Install dependencies:
   ```bash
   npm ci
   ```

## Development Workflow

- **Build:** `npm run build`
- **Test:** `npm test`
- **Verify modules:** `npm run verify:modules`

## Submitting Changes

1. Create a branch from `main`:
   ```bash
   git checkout -b feat/my-feature
   ```
2. Make your changes and ensure all tests pass.
3. Commit using [Conventional Commits](https://www.conventionalcommits.org/):
   ```
   feat: add support for table conversion
   fix: handle missing image gracefully
   ```
4. Push your branch and open a Pull Request against `main`.

## Code Style

- TypeScript only — no plain JavaScript in `src/`
- Keep converters focused: one file per HTML element
- Add or update tests in `test/` for any changed behavior

## Reporting Issues

Please use [GitHub Issues](https://github.com/bcibibi/docx-html-converter/issues) and include:
- A minimal reproducible example
- Expected vs actual behavior
- Node.js and package version
