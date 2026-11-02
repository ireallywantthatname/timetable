# NSBM Computer Security timetable sync

Downloads the NSBM SharePoint timetable, extracts Computer Security events, and writes an iCalendar file for Proton Calendar.

## Calendar URL

Add this subscription URL in Proton Calendar:

```text
https://<username>.github.io/<repo>/timetable.ics
```

## Local run

```bash
bun install
bun run typecheck
bun test
bun run sync
```
