# Sierra Speakers Toastmasters — Meeting Management Macros

A Google Apps Script project that automates meeting management for [Sierra Speakers Toastmasters](https://www.toastmasters.org/) (Club #3767, San Francisco). Built as a bound script inside a Google Sheets scheduling spreadsheet, it handles role confirmations, agenda generation, and club email drafting — with optional Gemini AI integration for smart content suggestions.

---

## What It Does

This script turns a standard Google Sheets meeting schedule into a full meeting-management dashboard. Club officers (typically the Toastmaster of the Day or VP Education) can run everything from a custom **Toastmasters** menu that appears in the spreadsheet toolbar.

### Features

**Role Confirmation Emails**
- Reads the scheduling sheet to find who is assigned to each role for a given meeting date
- Detects confirmation status via cell background colors (green = confirmed, red = unable, yellow = email sent)
- Generates personalized confirmation emails for each role (Toastmaster, speakers, evaluators, grammarian, joke master, table topics master, etc.)
- Supports hybrid, virtual, and in-person meeting formats with appropriate details
- Creates Gmail drafts (not auto-sends) so you can review before sending
- Handles "thank you" emails for already-confirmed members
- Fuzzy name matching when member names don't exactly match the roster

**Meeting Agenda Generation**
- Builds a complete meeting agenda with timed segments
- Scans speaker emails (via Gmail) to extract speech titles, project paths, and timing
- Generates a formatted Google Doc agenda with proper Toastmasters structure
- Optionally builds a companion Speaker Introductions document
- Supports Merriam-Webster dictionary lookup for Word of the Day definitions

**Club Meeting Hype Email**
- Drafts a club-wide email promoting the upcoming meeting
- Uses Gemini AI to write engaging, personalized email copy
- Includes meeting theme, Word of the Day, speaker lineup, and logistics
- Generates "Did You Know?" fun facts about speakers using AI
- Supports CC/BCC for guest invitations

**Gemini AI Integration**
- Word of the Day suggestions with definitions and example sentences
- Speech detail extraction from email threads
- Club email copywriting
- Smart model rotation (Gemini 2.5 Flash Lite → Flash → Gemma) with daily rate tracking
- **Auto-discovery & resilience**: Discovers available Gemini models via the v1beta API at runtime, ranks them, fails over across models on 404/400/429/5xx, caches the dead-model list for 24 hours, and gracefully degrades to templated content when no model is usable (see `src/GeminiResilience.gs`).
- **Update AI Models** menu item to manually refresh the discovered model list.

**Persistent Word of the Day (WOD Memory)**
- Hidden `WOD_Memory` sheet logs every Word of the Day picked (date, word, definition, pronunciation, part of speech, example, source, theme, model used).
- Smart caching avoids regenerating the WOD when the theme hasn't changed, but auto-refreshes when a stronger Gemini model becomes available.

**Deploy to Another Sheet**
- One-click menu option to deploy the current script to another target scheduling spreadsheet.

**Meeting Formats**
- Hybrid, Virtual, and **In Person** are all first-class formats with tailored email copy.
- Club email dialog includes a Meeting Format selector so the hype email adapts its copy to the chosen format (attendance prompts, location/Zoom details).

---

## Documentation

A full **step-by-step tutorial** with 19 real screenshots walks through every feature — from first open to generated agenda. Available in three formats:

- **Markdown**: [`docs/TUTORIAL.md`](docs/TUTORIAL.md)
- **PDF**: [`docs/Sierra_Speakers_Macro_Tutorial.pdf`](docs/Sierra_Speakers_Macro_Tutorial.pdf)
- **PowerPoint**: [`docs/Sierra_Speakers_Macro_Tutorial.pptx`](docs/Sierra_Speakers_Macro_Tutorial.pptx)

---

## Project Structure

```
sierra-speakers-toastmasters/
├── src/
│   ├── Code.gs                              # Main Apps Script source (~4850 lines)
│   └── GeminiResilience.gs                  # Gemini model auto-discovery & failover
├── docs/
│   ├── TUTORIAL.md                          # Step-by-step setup tutorial (Markdown)
│   ├── Sierra_Speakers_Macro_Tutorial.pdf   # Tutorial (PDF version)
│   └── Sierra_Speakers_Macro_Tutorial.pptx  # Tutorial (PowerPoint version)
├── images/                                  # 19 tutorial screenshots
├── appsscript.json                          # Apps Script manifest (OAuth scopes, runtime)
├── .clasp.json                              # Google clasp CLI config (push/pull)
├── .gitignore                               # Git ignore rules
├── CHANGELOG.md                             # Version history
├── CONTRIBUTING.md                          # Contribution guidelines
├── SETUP_GITHUB.md                          # Instructions for pushing to GitHub
├── LICENSE                                  # MIT License
└── README.md                                # This file
```

---

## Setup Instructions

### Prerequisites
- A Google account with access to Google Sheets and Gmail
- A copy of the Sierra Speakers scheduling spreadsheet (or your own Toastmasters schedule formatted similarly)

### Quick Start

1. **Copy the spreadsheet template**
   Open the scheduling spreadsheet and go to `File > Make a copy`. This gives you your own editable version with the script already attached.

2. **Open the script editor**
   In your copied spreadsheet, go to `Extensions > Apps Script`. You should see `Code.gs` with all the functions.

3. **Authorize the script**
   Run any function (e.g., `onOpen`) from the editor. Google will prompt you to review and grant permissions for:
   - Reading/writing spreadsheets
   - Sending email / creating Gmail drafts
   - Making external API requests (for Gemini and Merriam-Webster)
   - Accessing Google Drive (for agenda document creation)

4. **Reload the spreadsheet**
   Close and reopen the spreadsheet. You should see a **Toastmasters** menu in the toolbar with three options:
   - Start Role Confirmations
   - Generate Meeting Agenda
   - Draft Club Meeting Email
   - Update AI Models
   - Deploy to Another Sheet

### Configuration

The script uses **Script Properties** (stored server-side, not in the code) for sensitive configuration. Set these via `Extensions > Apps Script > Project Settings > Script Properties`:

| Property | Required | Description |
|---|---|---|
| `GEMINI_API_KEY` | Optional | Your Google Gemini API key for AI features. Get one free at [aistudio.google.com](https://aistudio.google.com/). |
| `MW_API_KEY` | Optional | Merriam-Webster Dictionary API key for Word of the Day lookups. Get one at [dictionaryapi.com](https://dictionaryapi.com/). |

You can also run `setGeminiKey()` from the script editor to store your API key programmatically.

### Spreadsheet Format Requirements

The script expects the scheduling sheet to follow this layout:

- **Rows 1–N**: Active club members with columns for First Name, Last Name, Pronunciation, Role/Notes, and Email
- **"Roles" header row**: Contains meeting dates across the top (as proper Date values)
- **Role rows below**: Each row is a role (Toastmaster, Speech 1, Evaluator 1, Grammarian, etc.) with member names in the date columns
- **Cell colors**: Green = confirmed, Red = unable to attend, Yellow = email sent, White = needs confirmation
- **Theme row**: One row above the Roles header, same columns as dates

---

## How It Works

### Role Confirmations Flow
1. Script reads the schedule sheet and identifies the next meeting date
2. Builds a list of all assigned roles and their confirmation status (based on cell background colors)
3. Shows a summary dialog where you enter meeting theme, Word of the Day, sender name, and meeting format
4. Walks through each unconfirmed role one at a time, showing a draft email you can edit
5. Creates Gmail drafts for each email (you send them manually from Gmail)
6. Colors cells yellow after emails are drafted

### Agenda Generation Flow
1. Prompts for meeting date and loads role assignments
2. Optionally scans Gmail for speaker emails to extract speech titles and project info
3. Uses Gemini AI to parse speech details from email threads
4. Looks up Word of the Day definitions from Merriam-Webster (with Gemini fallback)
5. Generates a timed agenda as a Google Doc with all meeting segments

### Club Email Flow
1. Collects meeting details (theme, speakers, WOTD, agenda URL)
2. Sends details to Gemini AI to draft an engaging club-wide email
3. Generates fun facts about speakers using AI
4. Creates a Gmail draft with formatted HTML email

---

## Tech Stack

- **Google Apps Script** (V8 runtime)
- **Google Sheets API** — reading/writing the schedule
- **Gmail API** — creating drafts, scanning speaker emails
- **Google Drive API** — creating agenda documents
- **Google Gemini API** — AI-powered content generation
- **Merriam-Webster Dictionary API** — Word of the Day definitions
- **HTML Service** — custom dialog UIs within Sheets

---

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.

---

## Contributing

Contributions are welcome! See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

---

## Author

**Mateusz Rakowski** — Sierra Speakers Toastmasters, San Francisco

Built to make Toastmasters meeting prep less tedious and more fun.
