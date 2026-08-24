export const SAMPLE_BRIEF = `# Quarterly Product Brief

**LS.md** · 24 August 2026  
A Lean Studio add-in for the Windows 11 desktop rollout

## Purpose

This page is a live Word document. Type Markdown in the pane — headings, lists, tables, and quotes land here as they would in Microsoft Word.

Keep **Live update** on to push every change. Switch it off when you want to draft first, then press **Apply to document**.

## What ships

1. A task pane for Word on Windows 11
2. Live Markdown → document sync
3. Native Word styles for headings, lists, and tables
4. Pull the current document back into Markdown

### Editing conventions

- \`**bold**\` and \`*italic*\` map to Word emphasis
- Headings (\`#\` … \`######\`) become Heading 1–6
- Tables, quotes, and fenced code survive the round-trip
- \`~~strikethrough~~\` and \`[links](https://example.com)\` are supported

- [x] Sideload the add-in into Word
- [x] Write the draft in Markdown
- [ ] Publish the org catalog

> Write the draft in Markdown. Let Word handle the typesetting.

## Schedule

| Milestone | Date | Owner |
| --- | --- | --- |
| Sideload build | 24 Aug | Design |
| Desktop QA | 28 Aug | Eng |
| Org catalog | 4 Sep | IT |

## Apply from script

\`\`\`ts
await Word.run(async (context) => {
  context.document.body.insertHtml(html, "Start");
  await context.sync();
});
\`\`\`

See the [Install for Word](/install) guide to pin LS.md on the Home ribbon.

---

*LS.md · lean.studio · Markdown for Word*
`;

export const SAMPLE_MEETING = `# Weekly standup

**Attendees:** Alex, Priya, Jordan  
**Date:** 24 August 2026

## Wins

- Closed the sideload path for Word on Windows 11
- Live sync now replaces the document body in one pass

## Blockers

1. Need a signed catalog URL for IT
2. Confirm WordApi 1.3 on the floor machines

## Actions

- [ ] Priya: ship the Shared Folder catalog
- [ ] Jordan: QA nested lists in Word 365
- [ ] Alex: write the one-pager for support

## Notes

> Keep updates to three bullets. Park detail in the appendix.

### Appendix

| Topic | Note |
| --- | --- |
| Ribbon | Button lives on the Home tab |
| Permissions | ReadWriteDocument |
`;

export const SAMPLE_LETTER = `# Letter

24 August 2026

Dear team,

I am writing to confirm that **LS.md** is ready for the Windows 11 desktop rollout. Authors can keep the draft in Markdown and let Word carry the styles.

Please review the attached schedule and reply with any blockers before Friday.

Yours sincerely,  
Terry

---

| Item | Detail |
| --- | --- |
| Product | LS.md |
| Host | Microsoft Word |
| Channel | Lean Studio task pane |
`;

export const SAMPLE_SPEC = `# Technical spec

## Summary

LS.md is a Lean Studio Office.js task pane. Markdown in the pane is converted to Word-friendly HTML and written into the active document.

## Flow

1. Author types GitHub-flavored Markdown
2. \`marked\` produces HTML
3. Tags are normalized (\`<b>\`, \`<i>\`, bordered tables)
4. \`Word.run\` replaces or inserts at the cursor

## Constraints

- Requires WordApi 1.3 (Microsoft 365 or Word 2021+)
- Source location must be HTTPS, except localhost
- Nested HTML CSS is limited — prefer semantic tags

## Error cases

| Case | Handling |
| --- | --- |
| Not in Word | Simulator preview only |
| \`insertHtml\` fails | Status line + retry via Apply |
| Empty draft | Clears the document body |

\`\`\`md
# Title
Write as you usually would.
\`\`\`
`;

export const TEMPLATES = [
  { id: "brief", label: "Product brief", markdown: SAMPLE_BRIEF },
  { id: "meeting", label: "Meeting notes", markdown: SAMPLE_MEETING },
  { id: "letter", label: "Letter", markdown: SAMPLE_LETTER },
  { id: "spec", label: "Technical spec", markdown: SAMPLE_SPEC },
  { id: "blank", label: "Blank page", markdown: "# \n\n" },
] as const;
