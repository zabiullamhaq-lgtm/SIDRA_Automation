# Assemble a synthetic workbook

Create a new `.xlsm` that you own. The source uses four working sheets and can
tolerate unrelated additional sheets. Required sheet names are Paths, Scenarios
(legacy `Sceanrios` is also accepted), Inputs and Outputs; trailing spaces are
normalized. Create writable, unprotected working sheets for the distributed module.
No populated original workbook is supplied.

| Sheet | Synthetic assignments |
| --- | --- |
| Paths | B2 = `Sidra File Directory`; C2 = your own saved baseline `.sip9` path; B3 = `Year`; C3 = `2035` |
| Scenarios | B1 = `Scenarios`; C1 = `1`; B2 = `Scenario 1`; C2 = `Demo`; B4 = `Peaks`; C4 = `AM`; C5 = `PM` |
| Inputs, settings | A3 = `HV %`; B3 = numeric `0.05`, displayed as `5%` |
| Inputs, junction header | A5 = `2035`; B5 = `Junction 1`; C5 = `Peak`; D5:G5 = `NB`, `EB`, `SB`, `WB` |
| Inputs, demand | C6 = `AM`; D6:G6 = `120`, `180`, `240`, `260`; C7 = `PM`; D7:G7 = `180`, `240`, `300`, `280` |
| Inputs, mapping header | I5:M5 = `Junction`, `NB`, `EB`, `SB`, `WB` |
| Inputs, mapping row | I6 = `Junction 1`; J6:M6 = `South`, `West`, `North`, `East` |
| Outputs | B2:E2 = `Scenario`, `V/c`, `Delay`, `LOS`; leave output rows blank; F2/G2 will become `Status`/`Model` |

All example counts are invented. Keep peak labels in Inputs column C because the
macro uses it to locate demand rows. Repeat junction blocks and mapping rows for
additional Sites, each matching exactly one baseline Site after normalization.
Every declared peak needs a demand row for every junction. The same peak-specific
demand is repeated across scenarios; years/labels add no growth.

The origin mapping must match your geometry; it describes geographic origin legs,
not the destination of traffic. Diagonal leg names can be parsed, but the source
expects the Site's legs to fit the four mapped origins. Use 60-minute volume units,
enabled light/heavy classes and baseline turning proportions. No Network processing
is supplied. Prepare geometry, timings, calibration and licensing separately.

Import the module, add the installed SIDRA 9 API reference, review/compile in Excel,
then invoke `SIDRA_RunModels`. With one scenario, one junction and two peaks, there
are two cases. Inspect the Outputs statuses, model copies, `Volume_Audit.csv` and
`Run_Log.csv`; no numerical v/c, delay or LOS is promised by this example.

The macro clears old output rows and saves its hosting workbook. Keep backups and
work in a disposable copy. Review all diagnostics before using engineering results.
