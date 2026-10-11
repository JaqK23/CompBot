# CompBot Insider

The Insider edition of [CompBot](../README.md): everything in CompBot, plus commands that need Excel features
most machines do not have yet. Commands graduate into CompBot when their feature reaches general release.

**You need** Excel for Windows on the Microsoft 365 **Insider Beta Channel** with in-cell arrays live
(Version 2610, Build 20520 or later, and `=FLATTEN(1)` must give 1 rather than #NAME?; the rollout is
gradual per machine, so File > Account > Update Now if it is missing). Without the feature the Insider
commands refuse with a plain status-bar message and change nothing. **Competition machines will not have
these features**: practise with Insider, compete with CompBot.

**Get it**: register `CompBot Insider.xlsm` from this folder in OA Robot (local file, or sync this folder's
address). Register ONE edition at a time, or disable the other: both carry the same commands and launch
codes, and the palette would show each twice.

## The Alves Method

Named after **Fredson Alves**, whose workflow it follows. Build a spill-based solve wide on the example row,
with room for every array, the way you normally would. Then, standing on the final answer cell:

| Command | Launch codes | What it does |
| --- | --- | --- |
| Alves Nest and Pack | `NP`, `FA`, `ALV` | Nest then Pack, in one go. Then run Save From Example from the answer. |
| Alves Unpack and Unnest | `UU`, `UNP`, `UFA`, `ALV` | Unpack then Unnest, in one go: one packed row laid out with every spill visible, to debug. |
| Alves Nest | `NST`, `ALV` | Every spilling step the answer depends on becomes one cell (`={...}`); references lose their `#`. Switches the workbook to compatibility version 3 (in-cell arrays on). |
| Alves Pack | `PCK`, `ALV` | Moves the nested steps onto the answer's row, right of the inputs, in dependency order, answer last. |
| Alves Unpack | `UPK`, `ALV` | Copies one packed row's steps to a scratch area with room for each array, still nested. The packed row is untouched. |
| Alves Unnest | `UNN`, `ALV` | The reverse of Nest: every nested step spills again; references get their `#` back. |

No selection is needed: the solve region is everything the answer depends on, on that sheet, right of the
level's inputs (the same rule Save From Example uses). Every command reports on the status bar, never in a
dialog, and a refusal starts `ALVES ... FAILED:` with the reason and the cell.

Why it matters: spilled helpers cannot be filled down, because the spills collide. Nested, each step is one
cell, so a whole solve chain copies down like an ordinary row, and you still get to build and debug with
visible spills.

Version: built from CompBot 2026.10.2. Definitions: [CompBot Insider Readme](CompBot%20Insider.xlsm%20README/CompBot%20Insider.xlsm%20README.md).
