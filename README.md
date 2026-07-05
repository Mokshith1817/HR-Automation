# HR-Automation
AI-powered HR letter automation system built for Atoms Digital Solutions
# HR Letter Automation System
### Atoms Digital Solutions Pvt. Ltd.

An AI-powered HR letter automation system that eliminates manual letter writing and sending. Built as a Google Apps Script web app integrated with Google Sheets, Gmail, and Google Drive — actively used by the HR team at Atoms Digital Solutions.

---

## The Problem It Solved

Previously, HR had to manually write, format, and send every letter or certificate individually. This was time-consuming and error-prone. This system automates the entire process — from filling employee details to generating a PDF and sending it via email — in under a minute.

---

## Three Ways to Use It

| Version | Interface | Best For |
|---|---|---|
| Version 1 — Chat | Conversational chatbot | Step-by-step guided experience |
| Version 2 — Form | Multi-step slide form | Structured data entry |
| Version 3 — Quick | Paste-and-go template | Fast bulk sending |

---

## Letter Types Supported

- Offer Letter
- Offer Letter with Probation
- Internship Offer Letter
- Experience Letter
- Reliving Letter
- Termination Letter
- Certification of Appreciation
- Certification of Completion
- Best Employee of the Month
- Best Employee of the Year

---

## Key Features

- **Auto PDF generation** — pulls from Google Docs templates, replaces placeholders, exports as PDF
- **Gmail integration** — sends branded HTML emails with PDF attached automatically
- **Google Drive storage** — organizes sent PDFs by letter type and year automatically
- **Dynamic fields** — form fields change based on letter type selected
- **Salary hike logic** — handles conditional probation salary hike scenarios
- **Logs sheet** — every sent letter is logged with timestamp, status, and Drive link
- **Beautify sheet** — one-click formatting for the data sheet with color-coded statuses

---

## Tech Stack

- **Google Apps Script** — backend logic and automation
- **HTML / CSS / JavaScript** — frontend web app UI
- **Google Sheets** — data source and logging
- **Gmail API** — email delivery with HTML templates
- **Google Drive API** — PDF storage and folder organization
- **Google Docs** — letter templates with `{{placeholder}}` system

---

## Impact

- **20+ letters sent** since deployment
- Actively used by HR at Atoms Digital Solutions
- Reduced letter sending time from ~15 minutes manual effort to under 1 minute
- Zero formatting errors due to template-based PDF generation

---

## Project Structure

```
├── Code.gs        # Backend — all logic, PDF generation, email sending, Drive storage
└── Index.html     # Frontend — Chat, Form, and Quick versions of the UI
```

---

## How It Works

1. HR opens the web app and selects a letter type
2. Fills in employee details via Chat, Form, or Quick paste
3. System fetches the matching Google Doc template
4. Replaces all `{{placeholders}}` with actual employee data
5. Exports the filled document as a PDF
6. Sends a branded HTML email with the PDF attached
7. Saves the PDF to Google Drive under the correct folder
8. Logs the action in the Logs sheet

---

## Built By

**Mokshith** — AI/ML Intern, Atoms Digital Solutions Pvt. Ltd.  
B.Tech Computer Science, 2nd Year (2026)

---

> *Built during internship at Atoms Digital Solutions, Guntur, Andhra Pradesh.*
