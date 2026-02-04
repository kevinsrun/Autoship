# Scholarship Tools (Automation Suite)

A Google Apps Script + Google Sheets workflow that helps me track scholarships, pull personalized scholarship opportunities, draft recommendation-letter / editor outreach emails, and sync deadlines to Google Calendar.

## Why this exists
I’m applying to a lot of scholarships with different requirements, portals, and deadlines. This project keeps everything in one place (a sheet), automates repetitive work (triage + document creation), and will eventually find new scholarships that match my profile.

---

## Current Features
- **Scholarship tracker sheet** with standardized headers
- **Triage + status workflow** (track where each scholarship is in the pipeline)
- **Portal / requirements scraping (basic)** to extract key info from scholarship pages
- **Doc generation** (optional): create Google Docs from templates or prompts (if configured)

> Note: Some features depend on Apps Script permissions (Drive, UrlFetchApp, Calendar, etc.).

---

## Roadmap (Planned)
### 1) Personalized Scholarship Discovery
Goal: Automatically suggest scholarships based on my profile (grade, state, intended major, leadership, research, demographics where applicable, etc.).

Planned approach:
- Maintain a **Profile** object / sheet (facts + keywords)
- Use **free scholarship sources** (RSS, websites, public lists) + basic crawlers
- Match opportunities using:
  - keyword scoring (major, orgs, leadership, service)
  - eligibility filters (deadline, location, grade, requirements)
- Auto-add results to the sheet with:
  - name, URL, due date, award, requirements, portal link, notes

### 2) Email Drafting (Recommendation Letters / Essay Editing)
Goal: Draft personalized emails to:
- ask for recommendation letters
- request essay edits / feedback
- follow up politely if no response

Planned approach:
- Store contacts in a **Contacts** tab (name, email, relationship, context)
- Store email templates in a **Templates** tab (rec letter, edits, follow-up)
- Generate a draft email from:
  - scholarship name + due date
  - recommender relationship/context
  - doc links (resume/brag sheet, essay doc, scholarship portal)
- Optional: write drafts into a **Gmail Drafts** folder or a “Draft Emails” sheet
  - (If Gmail draft API isn’t used, we still generate copy/paste-ready emails.)

### 3) Google Calendar Sync (Deadlines + Reminders)
Goal: Keep scholarship due dates on Google Calendar automatically.

Planned approach:
- For each scholarship with a due date:
  - create/update an event: `Scholarship Due: <Scholarship Name>`
  - attach links in the description (portal, docs, sheet row link)
  - add reminders (ex: 14 days, 7 days, 2 days, 1 day)
- Store `calendar_event_id` back into the sheet so edits update instead of duplicating

---

## Data Model (Google Sheet)
### Required Headers (Main Sheet)
These headers must exist **exactly** for the script to work reliably:

- Scholarship Name
- Theme
- Triage
- Difficulty
- Status
- Start date
- Due Date
- Application Portal
- Requirements
- Application File
- Additional Application Essay File
- Received
- Max Award Value
- Notes

Recommended additional columns (for upcoming features):
- Scholarship URL
- Contact(s)
- Prompt Text
- Word Limit
- Calendar Event ID
- Last Checked
- Source

---

## Setup
### 1) Create the Google Sheet
- Make a new Google Sheet and create your main tracker tab.
- Add the required headers (exact spelling/case).

### 2) Apps Script
- Extensions → Apps Script
- Paste the project code into the script editor
- Set constants like:
  - `MAIN_SHEET_NAME`
  - template Doc IDs (if used)
  - folder IDs (if used)

### 3) Enable Services / Permissions
Depending on what you’re running, you may need:
- **Google Drive** (Docs creation, file cleanup)
- **UrlFetchApp** (web page fetch for scholarship pages)
- **CalendarApp** (calendar sync)
- **Advanced Google Services** (optional)
  - Gmail API (draft emails)
  - Drive API (better trash/delete + file operations)

Run a small test function first to trigger the permissions prompt.

---

## How It Works (High Level)
1. You add/edit a scholarship row in the sheet.
2. Script can fetch the scholarship page, extract useful text, and fill fields.
3. Script can generate docs or link to required files.
4. Script can create draft emails to recommenders/editors.
5. Script syncs due dates to Google Calendar and updates events when dates change.

---

## Safety / Guardrails
- Web scraping is best-effort (pages vary; some block bots; PDFs need special handling).
- Calendar sync avoids duplicates by storing `Calendar Event ID`.
- Email drafting does **not** send mail automatically (unless explicitly enabled later).

---

## Contributing / Notes
This project is designed around my workflow, but it’s modular:
- `scrape/` = page fetch + prompt extraction
- `docs/` = doc creation + file cleanup
- `email/` = template fill + draft generation
- `calendar/` = event create/update/delete

---

## Next Milestones
- [ ] Add a `Profile` tab + JSON profile config
- [ ] Implement scholarship “source crawlers” (free lists + RSS)
- [ ] Add Gmail draft creation (optional)
- [ ] Add Calendar sync with reminders + update-by-ID
- [ ] Improve prompt extraction + PDF handling fallback

