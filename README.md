# Excel streaming custom functions remain Busy after recalculation

A standalone Office.js reproduction for desktop Excel. One streaming function immediately supplies a string, but full recalculation can leave both cells with identical arguments stuck at `#BUSY!`.

Observed on Excel for Windows **16.0.20430.20092**. In the recorded run, all 30 formulas initially resolved. The third Ctrl+Alt+F9 left Sheet2!B5 and B21 Busy, both using `=REPRO.MEMBER("Account","PL")`. The other 28 formulas resolved.

## Run the reproduction

Requires desktop Excel for Windows, Node.js and npm. The add-in uses a shared runtime and Microsoft's hosted Office.js. No application backend or sign-in is needed.

```powershell
git clone https://github.com/demyte/officejs-issue-streaming-busy-recalc.git
cd officejs-issue-streaming-busy-recalc
npm ci
npm run certs
npm start
```

The static server listens on `https://localhost:3443`. The certificate command uses Microsoft's localhost development certificate tool; skip it if a trusted certificate already exists in `%USERPROFILE%\.office-addin-dev-certs`.

In a second terminal, from the same folder:

```powershell
npm run sideload
```

1. Save other work and close Excel completely. Reopen Excel normally while the server remains running.
2. Open [repro.xlsx](repro.xlsx) through Excel's Open dialog and select **Sheet2**.
3. Open **Home > Stream repro > Open repro** if the pane is hidden. Wait for Ready and for the formulas to resolve.
4. Click **Read cell states**. The baseline should report no Busy cells.
5. Click a blank worksheet cell, press **Ctrl+Alt+F9**, and wait at least five seconds. Repeat until calculating icons remain.
6. Stop recalculating and click **Read cell states** again. Copy its JSON to capture the actual cell values.

The failure is intermittent. Three recalculations triggered the recorded failure, but this is not a fixed threshold. The affected argument set can vary.

**Expected:** every formula resolves to its supplied code after each recalculation.

**Actual:** both copies of one argument set remain Busy. Excel can still display their previous text with a calculating icon; the public `Range.valuesAsJson` read reports `errorType: "Busy"`.

For a fresh-workbook comparison, **Create reproduction sheet** adds the same two blocks to a new sheet. The supplied saved workbook is the confirmed case. A newly created sheet stayed healthy in some earlier runs.

If the pane does not initialize, close all Excel instances and reopen Excel normally. Calculation observations are only meaningful once the pane is Ready and the formulas have resolved.

## Function and workbook

The complete handler is:

```js
function member(hierarchyOrMember, memberCode, relation, ordinal, invocation) {
  invocation.onCanceled = () => {};
  invocation.setResult([[String(memberCode?.[0]?.[0] ?? "")]]);
}

Office.onReady(() => CustomFunctions.associate("MEMBER", member));
```

There is no cache, data request, timer, promise or instrumentation. The metadata declares streaming and a matrix result. Identical arguments retain Excel's normal shared-stream behavior.

Sheet2 contains 15 argument sets at B4:B18, repeated at B20:B34. The workbook was created with this add-in and is preserved byte for byte from the recorded run. It also contains a blank Sheet1 and a very-hidden `_REFERENCE888777888` sheet with eight generic report headers and no records. The workbook retains its creator name and original local folder in document metadata.

The capture machine had other add-ins installed, including a Premier Construction Software ribbon tab. This has not been reproduced on a machine with those add-ins removed. The cause and behavior on other Excel builds remain unconfirmed.

## Evidence and bug report

- [Full public cell snapshot](evidence/busy-snapshot.json), captured at 2026-09-24 12:16:46.908 UTC.
- [Screenshot](evidence/busy.png) showing both calculating cells and the matching public reader output.
- [Run record](evidence/observation.json), including baseline time, recalculation count, environment and unchanged workbook SHA256.
- [Office.js issue draft](OFFICEJS_ISSUE.md), ready to paste into an OfficeDev/office-js bug report.

The snapshot contains two Error/Busy cells with previous string PL and 28 resolved strings. The baseline's healthy status was observed in the pane; its complete JSON was not exported.

DevTools was opened only after the failure and public read to extract the existing snapshot. No private Office hooks or automated recalculation loops were used. No tests or builds were run.

## Repository contents

| File | Purpose |
| --- | --- |
| `functions.js`, `functions.json` | Single handler and its metadata |
| `manifest.xml` | Separate REPRO namespace and long-lived shared runtime |
| `index.html`, `taskpane.js` | Manual sheet creation and cell-state reader |
| `server.mjs` | Static localhost HTTPS server |
| `repro.xlsx` | Recorded reproduction workbook |
| `evidence/` | Failure capture and run record |
| `OFFICEJS_ISSUE.md` | Bug report draft |

## Stop and unregister

Stop the server with Ctrl+C. To unregister this add-in:

```powershell
npm run unregister
```
