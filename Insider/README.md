# CompBot Insider

The Insider edition of [CompBot](../README.md): everything in CompBot, plus commands that need Excel features
most machines do not have yet. Commands graduate into CompBot when their feature reaches general release.

**You need** Excel for Windows on the Microsoft 365 **Insider Beta Channel** with in-cell arrays live
(Version 2610, Build 20520 or later, and `=FLATTEN(1)` must give 1 rather than #NAME?; the rollout is
gradual per machine, so File > Account > Update Now if it is missing). Without the feature the Insider
commands refuse with a plain status-bar message and change nothing. **The competition machines in
Las Vegas do not have these features**, so if you plan to compete there, practise with CompBot.
Insider is for using in-cell array commands, like the Alves Method, now.

**Get it**: register `CompBot Insider.xlsm` from this folder in OA Robot (local file, or sync this folder's
address). Register ONE edition at a time, or disable the other: both carry the same commands and launch
codes, and the palette would show each twice.

## The Alves Method

By **Fredson Alves**: this is his workflow, as commands. Build a spill-based solve wide on the example row,
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
level's inputs (the same rule Save From Example uses). A helper spilled on another row (below the level, say)
comes along too when it depends on the example row; a block that does not (worked out once, shared by every
game) is left where it is and named on the status bar. Every command reports on the status bar, never in a
dialog, and a refusal starts `ALVES ... FAILED:` with the reason and the cell.

Why it matters: spilled helpers cannot be filled down, because the spills collide. Nested, each step is one
cell, so a whole solve chain copies down like an ordinary row, and you still get to build and debug with
visible spills.

Version: 2026.10.5 Insider, built from CompBot 2026.10.5. Definitions: [CompBot Insider Readme](CompBot%20Insider.xlsm%20README/CompBot%20Insider.xlsm%20README.md).

What changed: 2026.10.5 (22 Oct) Create Data Table finds every input column, including one the
example leaves blank but the games fill in (issue #15); Run Data Table keeps text answers such as
2-5-00 as text (issue #16), recalculates the whole sheet for every game when a cell stays on an old
value, and warns when every game gives the same answer (issue #17); Clear Unused Lambdas keeps a
lambda passed by name, as in MAP (issue #18). 2026.10.4 (16 Oct) the MAXLOOKUP and MINLOOKUP lambdas (issue #13); a Credits column
on the Info sheets (issue #11); Full Setup Case copes with cases full of images (issue #14); the six
Alves commands listed on Info_Commands; the tutorial gains an Insider section with a lesson for
each. 2026.10.3 (15 Oct) Run Data Table recalculates until Excel has finished, so a game
run straight after a new SIDE or ANSWER formula no longer comes out #NUM! or #N/A (issue #12).
2026.10.2 (11 Oct) a helper on another row that depends on the example row is nested and packed
wherever it sits; one that does not is left alone and named. 2026.10.1 (11 Oct) Alves Unpack and Alves Unnest
work on any game row, not only the first few below the level's header. 2026.10 (11 Oct) first release.
