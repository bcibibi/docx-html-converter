# Changelog

All notable changes to this project will be documented in this file.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased][Unreleased]

## [1.0.1][1.0.1] - 2026-07-27

### Changed

- #7 : new options for default IRunOptions
- #9 : new options fileReader, replacement of fileProvider
- Improved package compatibility for TypeScript/VS Code import resolution in consumer projects.
- Documentation refresh for current package name and converter options.

## [1.0.0][1.0.0] - 2026-07-20

### Added

- Initial release
- HTML to DOCX conversion support
- Converters: `b`, `i`, `u`, `strike`, `sub`, `sup`, `span`, `p`, `li`, `ol`, `ul`, `img`, `text`
- CSS color and style parsing
- ESM and CJS dual build output
- `DefaultDocxHtmlFileProvider` for file resolution

[Unreleased]: https://github.com/bcibibi/docx-html-converter/compare/v1.0.1...HEAD
[1.0.1]: https://github.com/bcibibi/docx-html-converter/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/bcibibi/docx-html-converter/releases/tag/v1.0.0
