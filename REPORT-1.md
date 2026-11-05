# REPORT-1 — Stability pass

Branch: `ds/task-1-stability` (based on `main` @ 25bbbe0)
Only `src/OfficeTabsLite.cs` was changed. No other file touched. One commit per item:

| Item | Commit | Subject |
| --- | --- | --- |
| 1 | a2ee505 | Log errors swallowed by Office callbacks and timer ticks |
| 2 | 8bbecd1 | Unsubscribe timer handlers on shutdown and when a strip closes |
| 3 | a7414cb | Release temporary COM objects (Windows collection, opened and added files, taskbar) |
| 4 | 9d25843 | Keep every timer tick from letting an exception escape |
| — | this commit | REPORT-1.md |

## Changed files

- `src/OfficeTabsLite.cs` — all four items (logging helper + call sites, event/handler cleanup, conservative COM release, timer hardening). Nothing else changed.
- `REPORT-1.md` — this report.

## Item 1 — Logging for swallowed errors

Added, lines 37–73: `internal static class Log` with one method `Log.Write(string where, Exception e)`.
It appends `yyyy-MM-dd HH:mm:ss  where  Type: message` to `%LOCALAPPDATA%\OfficeTabsLite\log.txt`,
creates the folder, deletes the file when it passes 256 KB, and never throws (its own body is guarded).
It also writes the *same* message at most once a minute: the sync tick runs 5×/second and an
un-throttled log would otherwise fill 256 KB in seconds — that throttle is what keeps this from
turning a broken COM call into a disk-thrashing loop.

`catch { }` → `catch (Exception ex) { Log.Write(...); }` at 39 sites: 12 were bare `catch { }` and 27
already had a statement on the same line. Every one is inside an Office callback, a WinForms event
handler or a timer tick:

| Site (original line) | where |
| --- | --- |
| OnConnection.Name (477) | `OnConnection.Name` |
| timer tick (516) | `Timer.Tick` |
| Later() tick (534) | `Later` |
| OnSettings (594) | `OnSettings` |
| OnNew PowerPoint fallback (630) | `OnNew.PowerPoint` |
| ActivateTab (643) | `ActivateTab` |
| CloseTab clean close (663) / FileClose (673) | `CloseTab`, `CloseTab.FileClose` |
| Sync pane visibility (792) | `Sync.PaneVisible` |
| SaveTab Save (841) / FileSave (846) | `SaveTab.Save`, `SaveTab.FileSave` |
| RenameFile delete old (873) | `RenameFile.DeleteOld` |
| OpenFile (1107) | `OpenFile` |
| RestoreSession open (1146) / method (1149) | `RestoreSession.Open`, `RestoreSession` |
| SetTaskbarButton (1281), RestoreTaskbar (1288) | `SetTaskbarButton`, `RestoreTaskbar.AddTab` |
| FitHeight (1310), CreatePane (1325) | `FitHeight`, `CreatePane` |
| GetWindows (1365) | `GetWindows` |
| ActiveIndex (1471), ActivateAt (1483), Invalidate (1499) | same names |
| CloseWatcher ctor (1515) / WndProc (1542) | `CloseWatcher.AssignHandle`, `CloseWatcher.WndProc` |
| KeyHook.Proc (1585) | `KeyHook.Proc` |
| FreezeFrame.Cover (1750) / tick (1795) | `FreezeFrame.Cover`, `FreezeFrame.Tick` |
| TabStrip anim tick (1899), new file tick (2019) | `TabStrip.AnimTick`, `TabStrip.NewDocument` |
| Relayout (2047), IconFor (2164), OnPaint (2174) | same names |
| OnMouseMove (2482), OnMouseDown (2497), OnMouseUp (2549), OnMouseWheel (2611), OnMouseDoubleClick (2624) | `TabStrip.Mouse*` |
| Defer (2630, both catches) | `Defer`, `Defer.BeginInvoke` |

Deliberately NOT logged (26 of the original `catch { }` left in place, unchanged, plus one new guard
inside `FreezeFrame.OnTick` that wraps `Finish()` so the finish path itself cannot reach the loop):

- Dispose paths: `CloseWatcher.Release`, `FreezeFrame.Finish` (`DestroyHandle`), `RemovePane` (`pane.Delete`).
- Expected control flow that relies on an exception: `DocOf`, `PathOf`, `IsModified`, `CaptionOf`,
  `ExtOf`, `KeyOf` (they ask every app's object model for Word/Excel/PowerPoint-only properties and
  the wrong ones always throw), `SaveDocAs` (`SaveAs2` → `SaveAs` fallback), `DwmSetWindowAttribute`
  (not available on older Windows), `ActiveWindow` reads in `StepTab`/`GroupWindows`, `StripOf` cast.
  Logging these would write a line per window per 200 ms tick with no diagnostic value.
- Registry helpers `PinStore`, `TabColorStore`, `FileLists`, `TabsSettings.Load` — not Office callbacks
  and not event handlers, so out of this item's scope.

## Item 2 — Event cleanup

Finding: the add-in has **no Office COM event subscriptions at all** (no `+=` on Application, Window,
Document, Workbook or Presentation); it polls with the 200 ms timer and watches window messages with
`CloseWatcher` (a `NativeWindow` assigned to the frame handle). So there was nothing to unhook in the
Office object model — only WinForms timers, which is what was actually leaking.

Every subscription in the file and where it is now removed:

| # | Subscription | Removed at |
| --- | --- | --- |
| 1 | `_timer.Tick += OnSyncTick` (CTPFactoryAvailable) | `StopTimer()`: `_timer.Tick -= OnSyncTick` before Stop/Dispose — called from both `OnBeginShutdown` and `OnDisconnection` |
| 2 | `t.Tick += tick` (Later) | inside the tick itself: `t.Tick -= tick` before Stop/Dispose |
| 3 | `wait.Tick += tick` (TabStrip.AnimateThenNew) | inside the tick itself: `wait.Tick -= tick`, Stop/Dispose in `finally` |
| 4 | `_timer.Tick += OnTick` (FreezeFrame.RevealWhenReady) | `FreezeFrame.Finish()`: `_timer.Tick -= OnTick` (the cover always finishes within 2 s) |
| 5 | `_animTimer.Tick += OnAnimTick` (TabStrip ctor) | new `TabStrip.Dispose(bool)` override: unsubscribe, Stop, Dispose, dispose the tooltip |
| 6 | `all.Click`, `one.Click` (CloseAskForm) | no change needed: the form is created in `using` in `DecideClose`; the controls die with it |
| 7 | `ok.Click` (RenameForm) | no change needed: form in `using` in `PromptRename` |
| 8 | menu item `Click` handlers (AddItem / Recently Closed / Tab Color) | no change needed: the ContextMenuStrip is disposed on `menu.Closed` |
| 9 | `menu.Closed +=` (ShowTabMenu) | self-removing: handler disposes the menu, the menu disposes its handlers |

`CloseWatcher` (Office window hook) was already released correctly: `StopWatching()` on shutdown, and in
`Sync()` for every frame handle where `IsWindow` is false; `_watchers` is a dictionary of live handles,
so each window gets exactly one watcher.

Also changed: `OnBeginShutdown` now also calls `RemoveAllPanes()` and `RestoreTaskbar()` (it previously
only stopped the timer and the hook, relying on `OnDisconnection`; off both callbacks may not arrive).
Both calls are idempotent, so the normal BeginShutdown → Disconnection sequence runs them twice safely.

Functions added: `Connect.OnSyncTick`, `FreezeFrame.OnTick`, `TabStrip.OnAnimTick`, `TabStrip.Dispose(bool)`.
Functions modified: `CTPFactoryAvailable`, `Later`, `StopTimer`, `OnBeginShutdown`, `RevealWhenReady`, `Finish`, `TabStrip..ctor`, `AnimateThenNew`.

## Item 3 — COM object release (conservative)

Added `Connect.ReleaseCom(object)` — no-op for null/managed objects, guarded, logs a failure.
Four release points, all objects that exist only for one call and are stored nowhere:

| Object | Where obtained | Released in |
| --- | --- | --- |
| `Application.Windows` collection | `GetWindows()` | `finally` in `GetWindows()` (the window objects it returns stay with the caller) |
| Document/Workbook/Presentation returned by `Open` | `OpenInApp()` | `finally` in `OpenInApp()` |
| Document/Workbook/Presentation returned by `Add` | `OnNew()` | after the fallback chain in `OnNew()` |
| `ITaskbarList` (`_taskbar`) | `SetTaskbarButton()` | `OnDisconnection()`, after `RestoreTaskbar()`, then `_taskbar = null` |

Deliberately NOT released (kept/shared, releasing them would risk `InvalidComObjectException`):
`_app`, `_addInInst`, `_factory`, `_ribbon`, every `CustomTaskPane` in `_panes`, every Document/Workbook/
Presentation object returned by `DocOf()`/`Application.ActiveWindow`, the `TabStrip` returned by
`pane.ContentControl`, and the Window objects in `Sync()`'s `wins` dictionary (they are stored in a local
dictionary and passed to `CreateCTP` during the same call).

## Item 4 — Timer safety

Five timers exist; every tick is now fully guarded:

- `Connect._timer` (200 ms) → `OnSyncTick`, try/catch + log.
- `Later()` one-shot timer → whole body inside try/catch, so Stop/Dispose/action cannot leak an exception.
- `FreezeFrame._timer` (15 ms) → `OnTick`, catch logs and calls `Finish()` inside its own try/catch (Finish also deletes the layered window).
- `TabStrip._animTimer` (15 ms) → `OnAnimTick`, already guarded; still stops itself on failure.
- `TabStrip` new-file timer in `AnimateThenNew` → had no guard; whole body now in try/catch/finally, and the
  placeholder is cleared even when `NewDocument()` fails (`finally` stops and disposes the timer).

## Build result

`powershell -ExecutionPolicy Bypass -File src\build.ps1` run after each item (4 runs): all succeeded,
`built add-in DLL`, `built Center app`, `built installer`, exit code 0. No warnings-to-error, no compiler errors.

## Manual test checklist (Word / Excel / PowerPoint)

1. Word, three documents: tabs appear, click each tab, `Ctrl+Tab` and `Ctrl+Shift+Tab` switch, `×` and middle-click close.
2. Pin a tab, then click its `×`: the "tab is pinned" dialog appears; unpin from the tab menu and close it.
3. Close a Word window with 2+ files: "Close all / Close this file" dialog, including "Don't ask me again".
4. Drag a tab to reorder; right-click → Tab Color, Pin, Save All, Reopen Closed File all still work.
5. `+` and double-click on empty strip space open a new file with the slide-in animation.
6. Excel with two workbooks: one taskbar button, tabs switch, `Save All` saves only modified files.
7. PowerPoint: open two presentations with the same file name from different folders — they must be two distinct tabs.
8. Delete `%LOCALAPPDATA%\OfficeTabsLite\log.txt`, use all three apps for a few minutes, then check the file:
   it should exist and stay small (a repeated failure is written at most once a minute).
9. Leave Office open ~2 minutes switching tabs: no hang, no flicker, `Task Manager` CPU near zero when idle.
10. With "reopen last time's files" enabled in the Center, close and reopen Office: the previous files come back,
    no duplicate windows and no leftover tab panes.
