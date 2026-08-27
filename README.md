# NYB Coach Hub

## Local test app

Run `npm run dev`, then open `http://127.0.0.1:8787`.

The local server uses sample GameChanger cards and keeps test submissions only in memory. Nothing is written to the production Apps Script spreadsheet, the Master field-schedule sheet, or the live GitHub Pages site.

Use the three same-day sample games to test the PitchSmart result levels:

- First game at 20 pitches or fewer, then a second appearance: warning.
- First game over 20 pitches, then a second appearance: violation.
- A third appearance on the same day: violation.
- Rest and the eligible date use the pitcher's total pitches across every game that day.
