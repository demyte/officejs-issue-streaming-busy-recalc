# Captioned reproduction

[Watch the video](reproduction-captioned.mp4) (2 minutes 28 seconds, MP4/H.264, approximately 1.9 MB). [Caption transcript](reproduction.srt).

Recorded on 24 September 2026 with the supplied `repro.xlsx` and standalone add-in already loaded in desktop Excel. The clip shows a healthy public cell-state read, one Ctrl+Alt+F9 recalculation, and both copies of `=REPRO.MEMBER("Account","GP")` remaining Busy in **Sheet2!B10 and B26**.

## Sequence

| Video time | Action or observation |
| --- | --- |
| 00:00 | Workbook and add-in are already open; both formula blocks display their codes. |
| 00:13 | Read cell states reports no Busy cells. Snapshot time: 13:03:33.802 UTC. |
| 00:39 | Select B4 to return focus to the worksheet. |
| 00:54 | Press Ctrl+Alt+F9. Both GP cells retain calculating icons. |
| 01:18 | Read cell states reports Busy in B10 and B26. Snapshot time: 13:04:38.382 UTC. |
| 01:35 | Select B10 and show its formula. |
| 01:51 | Select B26 and show the identical formula. |
| 02:08 | Read again without another recalculation. Both are still Busy. Snapshot time: 13:05:28.423 UTC, about 74 seconds after the keypress. |

The taskpane reader uses the public `Range.valuesAsJson` API. Its status and JSON are visible in the recording. The complete JSON from this run was not exported. No DevTools, private Office hooks, automated recalculation loops, code edits or workbook edits were used during the capture.

## Editing and environment

The finished clip is one continuous, real-time excerpt from the capture. The first 210 seconds were trimmed: those contained seven earlier Ctrl+Alt+F9 attempts and one Ctrl+Alt+Shift+F9 attempt, after which the cells remained healthy. There are no internal cuts or speed changes. Captions and orange outlines around the affected cells were added afterward; the worksheet content was not altered. There is no audio.

Excel version: **16.0.20430.20092**, Windows 11 Pro Insider Preview build 28120. Other add-ins were installed on this machine, as noted in the main README. The cause remains unconfirmed, and failure timing varies between runs.

The workbook hash was verified before this run:

```text
SHA256 FFBB52E74F8590815851D1653DBA545D3CBBBB8C07252D3B6909B474CC2B5FA2
```
