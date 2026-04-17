# Changelog

All notable changes to the Sierra Speakers Toastmasters Meeting Management Macros will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/), and this project adheres to [Semantic Versioning](https://semver.org/).

---

## [1.2.0] — 2026-04-16

### Added
- **Format-aware club hype email**: Club email dialog now includes a Meeting Format selector (Hybrid / In Person / Virtual / Undecided). Email copy adapts to the selected format — attendance prompt is shown only for Hybrid meetings, and location/Zoom details adjust accordingly.
- **"Did You Know?" facts deduplication**: AI-generated fun facts about speakers are now deduplicated, preventing the same fact from appearing more than once in the club email.
- **Comprehensive tutorial documentation**: Full setup-and-usage tutorial with 19 real screenshots, published in three formats — Markdown (`docs/TUTORIAL.md`), PDF (`docs/Sierra_Speakers_Macro_Tutorial.pdf`), and PowerPoint (`docs/Sierra_Speakers_Macro_Tutorial.pptx`).
- **Mobile PWA app design**: Concept and documentation for a mobile Progressive Web App companion to the scheduling spreadsheet.
- **Comprehensive code review**: Full audit of the codebase with targeted fixes and improvements.

### Fixed
- **Email sign-off uses sender name instead of Toastmaster name**: Confirmation emails are now signed by the person sending them, not always by the Toastmaster of the Day.
- **Future-week club email data anchoring (`perDateData`)**: Club hype emails now correctly anchor to the chosen week's row data when generating emails for future meetings, instead of pulling from the current week.
- **Added "In Person" meeting format option**: In Person is now a first-class format option alongside Hybrid and Virtual, with matching email copy.
- **`dateStr` undefined in `generateAgenda`**: Agenda generation no longer crashes on a ReferenceError when the date string wasn't set in certain code branches.
- **Club email TypeError (`speakersJson` comment placement)**: Fixed a syntax/scope bug where a misplaced comment around `speakersJson` caused a TypeError when drafting the club-wide email.
- **Agenda alternating row colors with odd evaluator count**: Alternating-row banding in the agenda now renders correctly when the evaluator count is odd (previously skipped or doubled a row).
- **WOD_Memory sheet auto-hiding**: The hidden WOD_Memory sheet could become the active tab and fail to hide; now calls `setActiveSheet` on the main sheet before `hideSheet` so the tab is reliably hidden.
- **Meeting format email logic**: Attendance prompt ("attending in person or via Zoom") is now shown only for Hybrid meetings; In Person and Virtual meetings get format-specific copy.
- **Removed leftover `testBug1Fix` function**: Cleaned up debug/test code that was accidentally left in the production script.

### Changed
- **WOD_Memory hidden sheet**: Now stores persistent Word of the Day selections with full metadata (date, word, definition, pronunciation, part of speech, example, source, theme used, Gemini model used).
- **Smart WOD caching**: Word of the Day caching is now theme-aware, model-aware, and respects user selections — regenerates only when theme changes or a stronger model becomes available.
- **Deploy to Another Sheet**: Menu option now includes script property transfer to the target spreadsheet.
- **Gemini model auto-discovery and resilience** (`GeminiResilience.gs`): Discovers available Gemini models via the v1beta API at runtime, ranks them, fails over on 404/400/429/5xx, caches dead models for 24 hours, and gracefully degrades to templated content.
- **Update AI Models menu item**: Manually refreshes the discovered Gemini model list from the Toastmasters menu.

---

## [1.1.0] — 2026-04-13

### Added
- **WOD_Memory hidden sheet**: Persistent Word of the Day selection log (date, word, definition, pronunciation, part of speech, example, source, theme, Gemini model used) so past WOTDs are never repeated.
- **Smart WOD caching**: Word of the Day now caches per theme and per Gemini model; regenerates only when theme changes or a stronger model becomes available.
- **Deploy to Another Sheet menu option**: One-click deployment of the current script to another target spreadsheet.
- **Gemini model auto-discovery and resilience** (new `GeminiResilience.gs`): Automatically discovers available Gemini models via the v1beta API, ranks them, falls back across models on 404/400/429/5xx, caches the dead-model list for 24 hours, and gracefully degrades to templated content when no model is usable.
- **Update AI Models menu item**: Manually refreshes the discovered Gemini model list from the menu.
- **"In Person" meeting format**: Added as a first-class option alongside Hybrid and Virtual, with matching email copy for in-person-only meetings.

### Fixed
- **Email sign-off uses sender name instead of Toastmaster**: Confirmation emails are now signed by the person sending them, not always by the Toastmaster of the Day.
- **Future-week club email data anchoring**: Club hype emails now correctly anchor to the chosen week's row data when generating emails for future meetings, instead of pulling from the current week.
- **`dateStr` undefined in `generateAgenda`**: Agenda generation no longer crashes on a ReferenceError when the date string wasn't set in certain branches.
- **Club email TypeError (speakersJson comment placement)**: Fixed a syntax/scope bug where a misplaced comment around `speakersJson` caused a TypeError when drafting the club-wide email.
- **Agenda alternating row colors with odd evaluator count**: Alternating-row banding in the agenda now renders correctly when the evaluator count is odd (previously skipped/doubled a row).
- **WOD_Memory sheet auto-hiding**: The hidden WOD_Memory sheet could become the active tab and fail to hide; now calls `setActiveSheet` on the main sheet before `hideSheet` so the tab is reliably hidden.

### Changed
- Split resilience/fallback logic out of `Code.gs` into a dedicated `GeminiResilience.gs` for clarity.

---

## [1.0.0] — 2026-04-11

### Initial Release

This is the first versioned release, capturing the current state of the project as it has been used in production by Sierra Speakers Toastmasters (Club #3767, San Francisco).

### Features
- **Role Confirmation Emails**: Automated email drafting for all meeting roles with color-coded status tracking (green/red/yellow cell backgrounds)
- **Meeting Agenda Generation**: Timed agenda creation as Google Docs with speaker info extraction from Gmail
- **Club Hype Email Drafting**: AI-powered club-wide meeting promotion emails
- **Gemini AI Integration**: Smart Word of the Day suggestions, speech detail extraction, and email copywriting using Google Gemini (2.5 Flash Lite / Flash / Gemma with automatic fallback)
- **Merriam-Webster API Integration**: Dictionary lookups for Word of the Day with definitions and example sentences
- **Meeting Format Support**: Hybrid, virtual, and in-person meeting configurations with appropriate email content
- **Speaker Introductions Document**: Optional companion document generated alongside the agenda
- **Fuzzy Name Matching**: Handles slight name variations between schedule and member roster
- **Gmail Draft Workflow**: All emails created as drafts for manual review before sending
- **Custom UI Dialogs**: In-spreadsheet dialogs for configuration, email review, and definition selection
- **Continuation-Based Flow**: Avoids Apps Script 6-minute timeout by using dialog-driven continuation between steps
- **Smart Model Rotation**: Daily rate tracking for Gemini API calls with automatic model fallback
- **Debug Utilities**: Built-in diagnostic functions for API keys, script properties, and speaker scanning

### Technical Details
- Google Apps Script, V8 runtime
- Bound to a Google Sheets scheduling spreadsheet
- OAuth scopes: Spreadsheets, Gmail, Drive, Script UI, External Requests
- Single-file architecture (`Code.gs`, ~4200 lines) — later split into `Code.gs` + `GeminiResilience.gs` in v1.1.0
