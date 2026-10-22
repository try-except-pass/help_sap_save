# help_sap_save

Saves SAP GUI exports automatically from Excel VBA when SAP opens the native
Windows **Save As** dialog, which SAP GUI scripting can't control.

## The problem

When a SAP GUI script presses an export button that opens the Windows Save As
dialog, the VBA call (`.Press`) blocks until the dialog closes. VBA can't fill
the dialog while it's blocked, and the usual workaround (an external script
using `AppActivate` + `SendKeys`) only works some of the time, because it
depends on window focus.

## How this works

1. Before pressing the export button, your macro calls `StartWatcher`.
2. `StartWatcher` writes a small `.vbs` to `%TEMP%` and runs it. The `.vbs`
   starts a hidden Excel, which opens `Watcher.xlsm` read-only and runs
   `modWatcher.WatchSaveAs`.
3. The watcher polls for a visible Save As dialog. When it finds one, it writes
   the target path directly into the File name box and triggers Save using
   Windows messages (`WM_SETTEXT` + `WM_COMMAND`). No focus or keystrokes are
   involved, and it reads the text back before saving, so it is reliable.
4. Your macro waits for the file to appear, closes the hidden Excel
   (`KillWatcher`) and opens the exported file in its own Excel instance
   (`OpenHere`).

Only Excel, VBA and Windows Script Host are needed. No `.exe` files and no
PowerShell.

## Files

| File | Goes into | Purpose |
|---|---|---|
| `modWatcher.bas` | `Watcher.xlsm` (new, small workbook) | Finds and fills the Save As dialog. Also has diagnostics. |
| `modExport.bas` | Your SAP report workbook | Runs the export, starts and stops the watcher, opens the result. |

A separate small `Watcher.xlsm` is used so the hidden Excel starts quickly,
even if your report workbook is large.

## Setup

1. Create a new workbook, import `modWatcher.bas` (VBA editor → File → Import
   File) and save it as `Watcher.xlsm` in the same folder as your report
   workbook.
2. In your report workbook, import `modExport.bas`.
3. In `ExportReport`, set `TARGET_DIR` / `TARGET_FILE`, and replace the
   placeholder SAP navigation and the `wnd[0]/...` export button with your own
   recorded code.
4. Save both workbooks.

To keep `Watcher.xlsm` somewhere else, set `WATCHER_PATH` in `modExport` to its
full path. Do this too if your report workbook lives on OneDrive or SharePoint,
because the hidden Excel can't open a workbook from a web address.

## Usage

```vb
Dim wbData As Workbook
Set wbData = ExportReport(session)
' work with wbData here
```

`ExportReport` returns the exported workbook opened in the Excel instance
running your macro. If SAP's own save dialog appears instead of the Windows
one, `ExportReport` fills that one directly.

## Notes

- **The hidden Excel takes a few seconds to start**, so call `StartWatcher`
  early (it is already at the top of `ExportReport`). The watcher keeps looking
  for up to 300 seconds.
- **SAP may open the exported file by itself.** With the hidden Excel idle,
  that copy often lands in it. `KillWatcher` closes it with the hidden Excel,
  and `OpenHere` opens the file in your own instance. If SAP's copy opened in
  another visible Excel instead, `OpenHere` opens a read-only copy.
- **Leftover hidden Excels** from earlier runs can be closed with
  `CleanupHiddenExcel`. Only Excel processes started by automation
  (`/automation` in the command line) are closed; Excel windows you opened
  yourself are never touched.
- **The target file is deleted before each export** so no "replace existing
  file?" prompt appears.
- **Macro security:** if macros don't run in the hidden Excel, put
  `Watcher.xlsm` in a Trusted Location.
- **WMI blocked?** `KillWatcher` uses WMI. If that is blocked, replace
  `p.Terminate` with `Shell "taskkill /F /PID " & pid, vbHide`.

## Troubleshooting

Open `Watcher.xlsm`, open the Save As dialog you want to automate, and run from
the Immediate window:

- `DiagDialogs` lists every visible dialog with its title and whether the
  Save button and File name box were found (`nameEdit` must be non-zero).
- `DiagTree` writes the full control tree of every visible dialog to
  `%TEMP%\dialog_tree.txt`. Type something recognisable in the File name box
  first to spot it. Use this if `nameEdit=0` on your version of Windows.
- `WatchSaveAs "C:\temp\test.txt", 5` runs the watcher directly in your own
  Excel, to test the dialog handling without the hidden Excel.

The dialog title is matched case-insensitively against `DLG_TITLE`
(`Save As`). Change it if your dialog uses a different (e.g. localized) title.
