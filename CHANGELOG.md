# Changelog

All notable changes to PDF Toolkit will be documented in this file.

The format is based on Keep a Changelog, with practical release notes for GitHub users.

## [0.3.1] - 2026-03-08

### Added

- Windows installer packaging with Inno Setup and matching GitHub release automation.
- First-run `Start Here` experience with bundled workflow templates.
- Recent task, recent file, and repeat-last-task shortcuts in the desktop UI.
- LLM-ready extraction bundles in Markdown, JSON, and JSONL.
- Public GitHub launch docs, release template, screenshots, and contributor/community files.

### Changed

- Reframed the product as an installable, offline-first Windows desktop app.
- Improved diagnostics so core-ready and optional add-ons are clearly separated.
- Updated repo metadata and links for the `AboveWireless/PDF-ToolKit` repository.

### Fixed

- Prevented LLM analysis from running on empty scan-heavy extraction bundles.
- Kept optional LLM and OCR dependencies from failing broad health checks.
- Fixed the center-pane scroll layout so the whole workspace stack scrolls correctly.
- Added GUI regression coverage for scroll ownership, operation switching, and template application.

## [0.3.0] - 2026-03-08

### Added

- Initial public-release packaging flow for the portable Windows app.
- GitHub issue templates, PR template, security policy, and contributor docs.
- Workflow-template product positioning, docs, and Windows install guidance.

### Changed

- Cleaned up the repo for public GitHub launch with Windows-first messaging.

