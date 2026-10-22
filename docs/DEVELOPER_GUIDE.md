# Developer editing and protection

The published `vba/SIDRAAutomation.bas` was extracted statically from the current
`SIDRA_Automation_Unrestricted.xlsm`. Its 488-line module includes layout/print
changes absent from the older loose `.bas`. The originals were not changed.

## Editing the source

Edit the public `.bas` in a text editor, or import it into a workbook you own through
Excel's VBA editor. Add the installed SIDRA 9 API reference and compile there.
Public text source does not require a protection password to edit. Preserve the
module name `SIDRAAutomation` and the public entry point `SIDRA_RunModels` unless
you also update button assignments and documentation.

## Owner-only protection details

A redacted [SIDRA_Owner_Only.txt](../SIDRA_Owner_Only.txt) explains advanced editing
of the protected workbook. **The existing editing password is unchanged in the
owner's local original and is not distributed in this public copy.** Obtain it
privately from the owner when needed. The owner states
that advanced VBA/workbook features require the password. The note specifically
documents worksheet/workbook-structure protection; a separate VBA-project unlock
procedure has not been independently verified.

For advanced workbook editing, consult that note and the workbook's actual
protection settings before editing protected sheets or structure. Retain
applicable copyright and attribution notices; the note does not add restrictions
to the Apache-2.0 code license.
The remaining workbook has four protected sheets, an unprotected About sheet,
and unprotected workbook structure despite its Unrestricted filename. No protection
was removed or changed during preparation. The password remains only in the
owner's local original, not in the public developer note or VBA module. The note retains its original workbook filename;
that named formatted workbook is not included in this source distribution.
The note's historical verification statement is retained as supplied by the owner;
it is not evidence of a new Excel/SIDRA runtime test for this publication.

## Verification and side effects

- `SIDRA_RunModels` performs real automation; `SIDRA_TestRun` also runs real models despite its name.
- The macro clears output rows, updates/saves the hosting workbook and creates copied model outputs.
- It reuses the SIDRA API session between runs and preserves baseline model factors.
- The baseline integrity check uses file size/modification time, not a cryptographic fingerprint.
- The distributed module has no protected-sheet helper equivalent to the earlier ZM variant; use writable working sheets.
- CSV audit cells can contain user/model text; inspect untrusted logs before opening them in spreadsheet applications.

State Excel/SIDRA versions and bitness when testing changes. Review one synthetic
case's inputs, turn/class allocations, saved model, results and error behavior.
No macro execution, Excel compilation, API connection or licensed model processing
was used for the current publication review. Compiled VBA/source equivalence and
runtime compatibility remain unverified.
