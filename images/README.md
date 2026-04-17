# Toastmasters Google Sheets Macro Tutorial - Diagrams

This directory contains 12 professional annotated diagram images (800x500px each) for the Toastmasters Google Sheets macro tutorial.

## Image Manifest

1. **01-main-sheet-overview.png** - Main Google Sheets interface mockup showing Sierra Speakers TM Schedule with the custom Toastmasters menu highlighted

2. **02-toastmasters-menu.png** - Dropdown menu showing all macro commands:
   - Start Role Confirmations
   - Generate Meeting Agenda
   - Draft Club Meeting Email
   - Deploy Code to Another Sheet
   - Update AI Models

3. **03-role-schedule-colors.png** - Role schedule section with color-coded legend:
   - Green = Confirmed
   - Yellow = Email Sent (Contacted)
   - Red = Needs Reassignment
   - White = Not Yet Contacted

4. **04-sheet-confirm-dialog.png** - Dialog confirming which sheet to process for role confirmations

5. **05-date-picker-dialog.png** - Date picker showing available meeting dates (4/16/2026, 4/23/2026, 4/30/2026)

6. **06-pending-confirmations.png** - Status dialog showing pending role assignments with form fields for:
   - Meeting Theme
   - Meeting Format (Hybrid/Virtual/Undecided)
   - Word of the Day
   - Email sender selection

7. **07-email-review-dialog.png** - Email preview/review dialog showing individual email template with To, Subject, and Body fields

8. **08-agenda-wod-picker.png** - Word of the Day selection dialog with Merriam-Webster, AI-Generated, and custom definition options

9. **09-agenda-generated.png** - Success notification showing "Meeting Agenda Created!" with link to open document

10. **10-hype-email-dialog.png** - Club meeting email composition dialog with fields for:
    - Meeting Date
    - Theme
    - Word of the Day
    - Agenda URL
    - Guest CC/BCC

11. **11-deploy-dialog.png** - Code deployment dialog for copying macro to another spreadsheet

12. **12-wod-memory-sheet.png** - Hidden sheet diagram showing Word of the Day memory table with columns:
    - Date, Word, Definition, Pronunciation, Example, Source, Theme, Model

## Design Specifications

- **Dimensions**: 800x500 pixels (all images)
- **Format**: PNG, RGB color, 8-bit
- **Background**: White (#FFFFFF)
- **Color Palette**:
  - Dark Gray Text: #333333
  - Light Gray Borders: #DADCE0
  - Medium Gray Background: #F5F5F5
  - Google Blue (highlighted/clickable): #4285F4
  - Google Green (confirmed): #34A853
  - Google Yellow (contacted): #FBBC04
  - Google Red (needs attention): #EA4335
- **Font**: DejaVu Sans (fallback to system default)

## Usage

These images are designed to be embedded in tutorial documentation, guides, or presentations to illustrate the Toastmasters Google Sheets macro workflow.

Generated using Python PIL/Pillow library.
