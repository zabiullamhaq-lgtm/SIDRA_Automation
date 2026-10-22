# SIDRA_Automation

**Excel VBA workflows for prepared SIDRA INTERSECTION 9 or above models, with results and volume audits. Suggestions welcome.**

The current VBA targets SIDRA 9. Newer versions require API/model compatibility
checks and any necessary adaptations; support for those versions is unverified.

An independent project by **Zabiulla M Haq** for traffic engineers working with
prepared, licensed SIDRA INTERSECTION models. This project is not affiliated with
or endorsed by SIDRA SOLUTIONS or Microsoft.

## What it does

- Reads scenario labels, peak-specific NB/EB/SB/WB approach counts and a common heavy-vehicle percentage from Excel.
- Uses a configurable direction-to-origin mapping and the baseline's turning proportions.
- Creates a copied, single-Site model for each scenario, junction and peak.
- Processes individual Sites through the installed SIDRA 9 API.
- Writes v/c, control delay, LOS, processing status and model links into the hosting workbook.
- Saves model copies, a volume audit, a processing log and a workbook snapshot in a new output folder.

Scenario years are labels: the macro adds no forecast growth. Existing growth,
flow-scale and peak-flow settings in the baseline remain in effect. It does not
create geometry, calibration or signal plans, and it does not process Networks.

## Requirements

- Windows, desktop Microsoft Excel with VBA, and separately licensed **SIDRA INTERSECTION 9** with its API installed.
- Compatible Excel/SIDRA bitness and an installed VBA reference exposing `SIDRASolutions_SI_API`.
- A saved `.sip9` baseline with prepared independent Sites, matching junction names, 60-minute volume units and enabled light/heavy vehicle classes.

The source targets SIDRA 9.0; 9.1 and later versions are unverified. Exact supported
Excel builds and bitness combinations have not been established by this review.
Commercial software, API libraries and project models are not supplied.

## Get started

1. Download this repository using **Code → Download ZIP**, or clone it with Git.
2. Work in a new local Excel workbook; save it as `.xlsm` and assemble the sheets using the [workbook schema](docs/WORKBOOK_SCHEMA.md).
3. Import [SIDRAAutomation.bas](vba/SIDRAAutomation.bas) in the VBA editor and add your installed SIDRA 9 API reference.
4. Review and compile the module; prepare a baseline model that you own and close it in SIDRA.
5. Review the inputs, then assign a button to `SIDRA_RunModels` or run that macro manually.
6. Inspect output statuses and logs before using the results.

This is a **source distribution**: no populated `.xlsm` or `.sip9` is included.
No Python installation is required for this VBA workflow. See the
[developer guide](docs/DEVELOPER_GUIDE.md) for source editing and protection details.

## Safe example and expected outputs

The schema guide uses invented AM counts of 120/180/240/260 vehicles and PM counts
of 180/240/300/280, with 5% heavy vehicles, one junction and one scenario.
That schedules two cases. For a baseline turn ratio of 2:3:5, the AM northbound
total of 120 splits into 24/36/60 vehicles, including 1.2/1.8/3 heavy vehicles.
These are arithmetic examples, not field data or predicted SIDRA results.

Expected output types are copied `.sip9` models, `Volume_Audit.csv`, `Run_Log.csv`,
and `SIDRA_Results.xlsm` in a timestamped folder beside the workbook. The hosting
workbook receives the results as well. Numerical v/c, delay and LOS depend on your
model and are not supplied as expected test values.

## Limitations and troubleshooting

**The macro clears old results and saves its hosting workbook, including on some
error paths. Use a disposable working copy and keep backups.**

| Issue | What to check |
| --- | --- |
| VBA type/reference error | Add the installed SIDRA 9 API reference; check version and matching bitness. |
| Missing or duplicate Site | Match each junction name to exactly one baseline Site. |
| Missing turning proportions | Set nonzero baseline turn totals on approaches with positive demand. |
| Protected-sheet error | Use an editable working copy and follow the published developer-access note where applicable. |
| Input rejected | Enter nonnegative numbers, explicit zeros and HV as a percentage; retain the required sheet layout. |
| Failed or warning result | Read `Run_Log.csv` and SIDRA diagnostics; review the saved model. |
| Output folder already exists | Choose a fresh run/output folder. |

Blank/negative inputs, missing mappings, unsupported extra-class demand and
Network-containing baselines are rejected. An all-zero enabled lane may still be
rejected by SIDRA; the macro does not alter geometry to suppress that diagnostic.
The VBA reader uses evaluated Excel values; use literal inputs for reproducibility.
Logs may contain model names, paths and input data: review them before sharing.

## Verification

The current workbook's VBA source was extracted and inspected without executing
macros. Source structure, procedure names, reference requirements, file provenance,
documentation links, ignore rules and staged contents were checked locally.
**Excel compilation and licensed SIDRA processing were not run for this review.**
There is no automated VBA test suite in the current directory. Previous Python
unit-test results do not validate this VBA-only publication candidate.

## Open for suggestions

Feature ideas, usability improvements and documentation corrections are welcome:
open a **Suggestion** issue and describe your use case, expected behavior and a
synthetic example. For bugs, use **Bug report**. Contributions should follow
[CONTRIBUTING.md](CONTRIBUTING.md); security concerns follow [SECURITY.md](SECURITY.md).
The [developer-access note](SIDRA_Owner_Only.txt) explains advanced workbook/VBA
editing, with the password redacted from this public copy. The existing password
remains unchanged in the owner's local original; obtain it privately from the owner
when required. Do not upload passwords, client traffic data, licence keys or private paths.
Support is community best effort; no response time or engineering acceptance is promised.

## License and acknowledgements

Copyright 2026 Zabiulla M Haq. Original code and documentation are offered under
[Apache-2.0](LICENSE); see [NOTICE](NOTICE). SIDRA INTERSECTION and Microsoft Excel
remain separately licensed third-party products. No rights in their names,
logos or software are claimed. The repository name describes the software this
tool automates; see [TRADEMARKS.md](TRADEMARKS.md).
