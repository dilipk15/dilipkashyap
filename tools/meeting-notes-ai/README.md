# AI Meeting Notes

A free Otter.ai alternative for Google Workspace users. One Google Apps Script that watches your Google Meet Drive folder, summarises new transcripts with Gemini, drops a formatted Doc in `Meeting Notes (AI)/`, and emails the summary to attendees.

Free. Open source. No monthly SaaS fee. You keep all your data in your own Drive.

---

## Requirements

- **Google Workspace Business Standard or higher.** Google Meet transcripts are only available on these paid Workspace plans — this is a Google limitation, not something the script can work around. Personal Gmail accounts and Business Starter cannot use this.
- **A free Gemini API key.** Get one at https://aistudio.google.com/apikey (30 seconds, no credit card). The free tier (1,500 requests/day) is far more than any team will use.
- **A Google Sheet** to install into. Any Sheet works — a fresh empty one is cleanest.

---

## Install (one-click)

1. **Open a Google Sheet.** A fresh, empty one is easiest.
2. **Extensions → Apps Script.** Delete any placeholder code in the editor.
3. **Paste the entire contents of `Code.gs`** into the editor. Save (Ctrl+S / Cmd+S).
4. **Reload the Sheet tab** in your browser.
5. **Menu → AI Meeting Notes → One-Click Setup (First Time).**
6. Approve the permissions prompt when it appears. It asks for:
   - Google Drive (to read your Meet transcripts and write summaries)
   - Gmail (to email you and your attendees the summaries)
   - Calendar (to find who was on the meeting)
   - External requests (to call the Gemini API)
7. When prompted, paste your **Gemini API key** and click OK.
8. Done. The script now runs every 15 minutes in the background.

---

## What happens next

Every 15 minutes, the script checks your Drive's `Google Meet/` folder for new transcript Docs. For each new transcript it finds:

1. Reads the full transcript text.
2. Looks at your Calendar for that day to find the matching meeting and its attendees.
3. Sends the transcript to Gemini and asks for a structured summary (summary, action items with owners, decisions made, open questions, highlight quotes).
4. Creates a formatted Google Doc in `Meeting Notes (AI)/` — titled `Notes - [Meeting Name] (YYYY-MM-DD)`.
5. Emails the summary to all meeting attendees (or just you, if you turn that off in Settings).
6. Logs the row in the `Log` tab of the Sheet.

You can trigger it manually any time: **Menu → Process New Meetings Now.**

---

## Settings

Edit the `Settings` tab of the Sheet:

| Setting | Default | What it does |
|---|---|---|
| Email attendees | Yes | Yes = send to everyone on the Calendar invite. No = only email you. |
| Send HTML email | Yes | Yes = pretty formatting. No = plain text. |
| Include full transcript link | Yes | Adds a "View original transcript" link at the bottom of every email. |
| Meeting-type prompt add-on | *(empty)* | Optional extra instruction to Gemini (e.g. `Focus on client objections and next steps.`). |
| Skip meetings shorter than (minutes) | 3 | Skip very short calls. Set to 0 to summarise everything. |

---

## Troubleshooting

**"Google Meet folder not found" during setup.**
The folder is auto-created by Google after your first Workspace Meet call with transcripts enabled. Run one short Meet call with transcripts on (three-dot menu → "Turn on transcript"), wait ~2 minutes, then re-run Setup.

**Nothing happens after a meeting.**
- Confirm transcripts are enabled by default on your calls (Workspace Admin → Meet → Transcripts).
- Wait up to 15 minutes (the trigger interval).
- Menu → Process New Meetings Now — this runs the check immediately.
- Check the `Log` tab for status rows.

**"Gemini API returned 429".**
You hit the free-tier rate limit (10 requests per minute, 1,500 per day). This resets at midnight Pacific time. If it happens often, the free tier is genuinely too small for your call volume — enable billing on your Gemini API project.

**"Gemini API returned 403".**
Your API key is invalid or was regenerated. Menu → Set Gemini API Key → paste the new one.

**Summary is wrong or missed things.**
Add a note in `Meeting-type prompt add-on` telling Gemini what to focus on. E.g. `This was a technical design discussion. Preserve all technical decisions verbatim.`

---

## Privacy

- Your transcripts stay in your own Google Drive.
- The transcript is sent to the Gemini API *only* when a new one appears. No transcript is stored on any third-party server.
- Gemini's data handling for API calls is documented at https://ai.google.dev/gemini-api/terms — Google states that free-tier API prompts and outputs may be used to improve their models. If that matters to you, enable billing (paid tier does not train on your data).
- Nothing about this script is affiliated with or endorsed by Google.

---

## Uninstall

Menu → AI Meeting Notes → Uninstall (Remove Trigger).

Your Sheet, Drive folder, and existing summary Docs stay put. Only the automatic trigger is removed. Delete the Sheet + folder manually if you want a clean slate.

---

## Known limits

- Only reads transcripts saved by Google Meet's native transcript feature. Recordings, chat exports, and third-party transcripts (Otter, Fireflies, Fathom) are ignored.
- Transcripts longer than ~200,000 characters are truncated. A 3-hour call fits comfortably; a full-day workshop may lose the tail.
- The Calendar match is a best-effort heuristic. If two meetings on the same day have similar names, the wrong one's attendees may be emailed. Check the `Log` tab.
- Recurring-meeting subfolders in Drive are shared across instances. Each instance's transcript is a distinct file, so the script handles this correctly, but the folder name reflects the series, not the specific date.

---

Built by [Dilip Kashyap](https://dilipkashyap15.medium.com/). Every automation tool I ship is free and open source. If this saved you a monthly subscription, consider [following me on Medium](https://dilipkashyap15.medium.com/) — I publish one build like this every month.
