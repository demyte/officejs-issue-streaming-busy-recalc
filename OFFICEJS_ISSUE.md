# Excel streaming custom functions with duplicate arguments remain #BUSY! after full recalculation

## Your environment

- Platform: PC desktop.
- Host: Excel.
- Office version: 16.0.20430.20092, reported by `Office.context.diagnostics.version`.
- Operating system: Windows 11 Pro Insider Preview, build 28120.
- Browser: not Office on the web; the desktop add-in uses WebView2.
- Office.js: hosted `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`.
- Runtime: long-lived shared runtime.
- Other add-ins were installed during the capture, including a Premier Construction Software ribbon tab. This has not been repeated with those add-ins removed.

## Expected behavior

A streaming custom function that immediately calls `invocation.setResult([[code]])` should resolve its cells after full recalculation, including multiple cells with identical arguments.

## Current behavior

After repeated Ctrl+Alt+F9 calculations, both copies of one argument set can remain `#BUSY!`. The previous string stays visible with a calculating icon.

In the initial standalone snapshot run, all 30 formulas initially resolved. The third full recalculation left Sheet2!B5 and B21 Busy. Both contain:

```excel
=REPRO.MEMBER("Account","PL")
```

A subsequent public `Range.valuesAsJson` read returned `type: "Error"`, `basicValue: "#BUSY!"`, `errorType: "Busy"` and previous string PL for those two cells. The other 28 formulas returned resolved strings. The failure is intermittent and the affected arguments can vary.

A separate [captioned video](https://github.com/demyte/officejs-issue-streaming-busy-recalc/blob/main/evidence/video/reproduction-captioned.mp4) shows a healthy baseline followed by both `GP` cells at B10 and B26 remaining Busy. A second public read confirms they are still Busy 74 seconds after the last recalculation. The video is a continuous real-time excerpt; earlier healthy attempts were trimmed. [Capture notes and timeline](https://github.com/demyte/officejs-issue-streaming-busy-recalc/blob/main/evidence/video/README.md).

## Steps to reproduce

1. Clone [the standalone reproduction](https://github.com/demyte/officejs-issue-streaming-busy-recalc).
2. Run `npm ci`, `npm run certs`, and `npm start`. The server runs at `https://localhost:3443`.
3. In another terminal in the repository, run `npm run sideload`.
4. Save other work, close Excel completely, reopen it normally, and open the supplied `repro.xlsx`. Select Sheet2.
5. Open the Stream repro taskpane and wait until Ready. Click **Read cell states** and confirm that the baseline has no Busy cells.
6. Focus a blank worksheet cell, press Ctrl+Alt+F9, and wait at least five seconds. Repeat until calculating icons persist.
7. Stop recalculating and click **Read cell states**. Its JSON records the actual Excel cell states.

The supplied workbook has 15 argument sets duplicated once. The observed run failed after three recalculations, but the exact count is not guaranteed.

## Link to example

[Source, manifest, workbook and run instructions](https://github.com/demyte/officejs-issue-streaming-busy-recalc).

This is a locally sideloaded desktop add-in, not a publicly hosted add-in endpoint.

## Additional details

The complete function is:

```js
function member(hierarchyOrMember, memberCode, relation, ordinal, invocation) {
  invocation.onCanceled = () => {};
  invocation.setResult([[String(memberCode?.[0]?.[0] ?? "")]]);
}

Office.onReady(() => CustomFunctions.associate("MEMBER", member));
```

There is no application backend, cache, data request, timer, promise, or tracing in this handler. Metadata retains two matrix arguments, optional string/number arguments, a matrix result, and streaming. No invocation-per-cell option is enabled.

The workbook was created with the standalone add-in. Its exact bytes are preserved. It also contains a blank Sheet1 and a very-hidden `_REFERENCE888777888` sheet with eight generic report headers and no records. Other installed add-ins and the Windows Insider environment have not been eliminated as possible factors. The evidence does not establish workbook corruption or a particular native defect.

## Context

A production streaming custom-function add-in intermittently leaves duplicate formulas calculating after full recalculation. This sample reproduces the symptom with an immediate string result and no application logic.

## Useful logs

- [Captioned reproduction video](https://github.com/demyte/officejs-issue-streaming-busy-recalc/blob/main/evidence/video/reproduction-captioned.mp4) and [caption transcript](https://github.com/demyte/officejs-issue-streaming-busy-recalc/blob/main/evidence/video/reproduction.srt).
- [Full public cell-state JSON](https://github.com/demyte/officejs-issue-streaming-busy-recalc/blob/main/evidence/busy-snapshot.json).
- [Screenshot of the failure](https://github.com/demyte/officejs-issue-streaming-busy-recalc/blob/main/evidence/busy.png).
- [Run record and workbook hash](https://github.com/demyte/officejs-issue-streaming-busy-recalc/blob/main/evidence/observation.json).
- [Reproduction workbook](https://github.com/demyte/officejs-issue-streaming-busy-recalc/blob/main/repro.xlsx).

DevTools was opened only after the failure and the public read, to extract the already captured JSON. No private Office API hooks or automated recalculation loops were active. The baseline's healthy status was observed in the pane, but its complete JSON was not exported.

> Written by Codex on behalf of James.
> _Codex gpt-6-astra xhigh_
