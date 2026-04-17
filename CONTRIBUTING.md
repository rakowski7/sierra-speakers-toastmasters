# Contributing to Sierra Speakers Toastmasters Macros

Thanks for your interest in contributing! This project automates meeting management for Toastmasters clubs using Google Apps Script. Whether you're a fellow Toastmaster or a developer, your contributions are welcome.

---

## Getting Started

1. **Fork this repository** on GitHub
2. **Clone your fork** locally:
   ```bash
   git clone https://github.com/YOUR_USERNAME/sierra-speakers-toastmasters.git
   cd sierra-speakers-toastmasters
   ```
3. **Create a feature branch**:
   ```bash
   git checkout -b feature/your-feature-name
   ```

## Development Setup

Since this is a Google Apps Script project, you have two options for development:

### Option A: Edit directly in Apps Script (easier)
1. Make a copy of the scheduling spreadsheet
2. Open `Extensions > Apps Script`
3. Edit `Code.gs` directly in the browser-based editor
4. Test by running functions from the editor or using the Toastmasters menu in the spreadsheet
5. Copy your changes back into `src/Code.gs` in your local repo

### Option B: Use clasp (advanced)
1. Install the Google Apps Script CLI:
   ```bash
   npm install -g @google/clasp
   clasp login
   ```
2. Update `.clasp.json` with your test project's script ID
3. Push changes:
   ```bash
   clasp push
   ```
4. Pull changes:
   ```bash
   clasp pull
   ```

## Making Changes

### Code Style
- Use JSDoc comments for all public functions (`@param`, `@return`, description)
- Private/internal functions should end with an underscore (e.g., `saveConfirmationState_`)
- Keep the single-file structure (`Code.gs`) — Apps Script works best this way for bound scripts
- Use `const` and `let` (V8 runtime), not `var`

### Testing
- Always test in a **copy** of the spreadsheet, never in the production sheet
- The script has `SCHEDULING_SHEET_URL_` at the top — set it to your test sheet URL during development, set to `null` for production
- Test each flow end-to-end:
  - Role Confirmations: verify emails are generated correctly and cell colors update
  - Agenda Generation: verify the generated Google Doc has correct timing and content
  - Club Email: verify the Gmail draft looks correct

### What to Contribute
- Bug fixes (please describe what was broken and how you fixed it)
- New role types or meeting formats
- Improved email templates
- Better error handling or user feedback
- Documentation improvements
- Support for other Toastmasters club formats

## Submitting Changes

1. Commit your changes with a clear message:
   ```bash
   git commit -m "Add support for custom meeting locations"
   ```
2. Push to your fork:
   ```bash
   git push origin feature/your-feature-name
   ```
3. Open a Pull Request against the `main` branch
4. Describe what you changed and why
5. Include screenshots of any UI changes

## Reporting Issues

If you find a bug or have a feature request, please open a GitHub Issue with:
- What you expected to happen
- What actually happened
- Steps to reproduce (if it's a bug)
- Your meeting format (hybrid/virtual/in-person) if relevant

## Code of Conduct

Be kind, be respectful, and remember that this project is built by Toastmasters for Toastmasters. We're all here to grow and help each other.

---

Thank you for contributing!
