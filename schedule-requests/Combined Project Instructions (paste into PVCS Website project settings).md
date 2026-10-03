# Combined project instructions

Paste the text below into the PVCS Website project's instructions (Project settings → Instructions), merging with what's already there. I can't read the Website project's existing instructions from here, so keep anything in them that isn't covered below.

---

This project covers the PVCS website and the staff schedule-request system (merged from the former "PVCS - Schedule Requests" project on Aug 8, 2026).

Your working folder is /Users/Mike/Library/CloudStorage/GoogleDrive-doctormikeharrington@gmail.com/My Drive/0 Claude Projects/PVCS Website. Connect to it at the start of each task before reading or writing files.

Website: the folder is a git repo for the PVCS website (GitHub Pages, encrypted with StatiCrypt). Publish changes with update-site.command / push-to-github.command.

Schedule requests: the schedule-requests/ subfolder holds the forms and Apps Script code used to collect staff scheduling requests — psychiatrist on-call request forms (psychiatrist-oncall-form*.gs), on-call schedulers (oncall-scheduler*.gs, latest v4), the PA extra-shift scheduler (pa-crc-scheduler.gs), and the PA additional-shift request email/PDFs. Use the latest version of each script unless told otherwise. A record of prior chats on this work is in schedule-requests/Schedule Requests - Chat Record.md.
