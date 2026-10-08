# <h1 id="oa-robot-definitions">OA Robot Definitions</h1>

\*\*CompBot.xlsm\*\* contains definitions for:

[104 Robot Commands](#command-definitions)<BR>[33 Robot Texts](#text-definitions)<BR>

<BR>

## Available Robot Commands

[Array](#array) | [Bonus](#bonus) | [Color](#color) | [Convert](#convert) | [DataTable](#datatable) | [Fill](#fill) | [Formatting](#formatting) | [GoTo](#goto) | [LAMBDA](#lambda) | [Lookup](#lookup) | [Maintenance](#maintenance) | [Map](#map) | [MM](#mm) | [Name](#name) | [Navigation](#navigation) | [Paste](#paste) | [Prep](#prep) | [Settings](#settings) | [WrapWith](#wrapwith) | [Other](#other)

### Array

| Name | Description |
| --- | --- |
| [Find Address Of Value On Map](#find-address-of-value-on-map) | Returns the cell ADDRESS of each value you are looking for on a map, the start point the distance commands need |
| [Find Nearest On Map](#find-nearest-on-map) | Stand on the start (e.g. the S) in a map spill, with the target cell COPIED: lists the nearest targets with their addresses and distances, to the right of the sheet's contents |
| [First Match By Row](#first-match-by-row) | For each row, the position of the first cell equal to the COPIED cell's value, or a fallback (blank \= never) where the row never matches |
| [Flip Array](#flip-array) | Reverses the active cell's text (stressed becomes desserts). For a block use FLR (left to right) or FTB (top to bottom) |
| [Flip Array Left to Right](#flip-array-left-to-right) | Mirrors the active array left to right: the last column comes first |
| [Flip Array Top to Bottom](#flip-array-top-to-bottom) | Mirrors the active array top to bottom: the last row comes first |
| [Label Connected Regions (4 ways)](#label-connected-regions-4-ways) | Numbers each connected blob on a map, cells joining only up, down, left and right |
| [Label Connected Regions (8 ways)](#label-connected-regions-8-ways) | Numbers each connected blob on a map, a diagonal (corner) touch joining cells too |
| [List Combinations Of Array](#list-combinations-of-array) | Every ordering of all the items in the active array, each used once (10 20 30, 10 30 20, 20 10 30 ...) |
| [List Pairings Of Two Lists](#list-pairings-of-two-lists) | Every pairing of the COPIED list with the active spill (each item of one with each item of the other), written to the right of the sheet's contents |
| [Map Distance From Cell (4 ways)](#map-distance-from-cell-4-ways) | How far every cell on a map is from the COPIED start cell(s), moving only up, down, left and right, routing AROUND obstacles |
| [Map Distance From Cell (8 ways)](#map-distance-from-cell-8-ways) | How far every cell on a map is from the COPIED start cell(s), diagonal moves allowed, routing AROUND obstacles |
| [Rotate Array](#rotate-array) | Turns the active array a quarter turn clockwise; run it again for another quarter turn |
| [Stack Sheets From Clipboard](#stack-sheets-from-clipboard) | Stacks the same block from many sheets into one table, driven by a sheet\-name list in the clipboard |
| [Stack Sheets From Clipboard, Hide Blanks](#stack-sheets-from-clipboard-hide-blanks) | As Stack Sheets From Clipboard, but drops empty cells. WARNING: it drops zeros too |
| [Wrap Flat List Into Grid](#wrap-flat-list-into-grid) | Folds a single row or column back into a grid of a chosen width, the missing half of Reshape To One Row\/Column |

### Bonus

| Name | Description |
| --- | --- |
| [Clear Bonus From Dock](#clear-bonus-from-dock) | Hides one bonus from the dock by hand (type 2, B4, B or Bonus 2) |
| [Clear Status Bar](#clear-status-bar) | Clears the status bar |
| [Next Bonus In Status Bar](#next-bonus-in-status-bar) | Shows the next open bonus question on the status bar (no\-dock fallback) |
| [Previous Bonus In Status Bar](#previous-bonus-in-status-bar) | Shows the previous open bonus question on the status bar |
| [Record Walk Route](#record-walk-route) | Start recording a route: every orthogonal move adds the cells you pass over. Diagonal click finishes |
| [Record Walk Route Here](#record-walk-route-here) | As Record Walk Route, but the route table goes on this sheet, level with where the walk started and to its right, in the first space clear enough to hold it |
| [Restore Cleared Bonuses](#restore-cleared-bonuses) | Brings every hand\-cleared bonus back into the dock |
| [Save Answer To Bonus 1](#save-answer-to-bonus-1) | Links the active cell (any sheet) into the Bonus 1 answer cell, goes there and copies it for the submission site. |
| [Save Answer To Bonus 2](#save-answer-to-bonus-2) | Links the active cell (any sheet) into the Bonus 2 answer cell, goes there and copies it for the submission site. |
| [Save Answer To Bonus 3](#save-answer-to-bonus-3) | Links the active cell (any sheet) into the Bonus 3 answer cell, goes there and copies it for the submission site. |
| [Save Answer To Bonus 4](#save-answer-to-bonus-4) | Links the active cell (any sheet) into the Bonus 4 answer cell, goes there and copies it for the submission site. |
| [Save Answer To Bonus 5](#save-answer-to-bonus-5) | Links the active cell (any sheet) into the Bonus 5 answer cell, goes there and copies it for the submission site. |
| [Show Bonus Dock](#show-bonus-dock) | Shows the open bonus questions in a dock on the right; answered ones drop off |
| [Stop Walk Route](#stop-walk-route) | Ends a walk route by hand and writes it out, for when a diagonal click is awkward to reach |

### Color

| Name | Description |
| --- | --- |
| [Copy Map And Fill From Legend](#copy-map-and-fill-from-legend) | Copy a color\-coded map sheet: stand on a colored legend cell; copies the sheet and fills every map cell on the copy with its color's legend value (white is never a key); each named range on the sheet gets a clr\_ twin on the copy |

### Convert

| Name | Description |
| --- | --- |
| [Round to 0](#round-to-0) | Wraps the active formula in ROUND(...,0) for a whole number. Round once at the END of a chain; rounding each step compounds the error |

### DataTable

| Name | Description |
| --- | --- |
| [Link Answer Cell](#link-answer-cell) | Points the level's data table ANSWER cell at the active cell (your calculation's output); then run RDT |
| [Link Side Cell](#link-side-cell) | Adds a SIDE column to the level's data table, pointed at the active cell, so RDT also records a second value per game for a bonus |

### Fill

| Name | Description |
| --- | --- |
| [Copy Map And Fill From Legend](#copy-map-and-fill-from-legend) | Copy a color\-coded map sheet: stand on a colored legend cell; copies the sheet and fills every map cell on the copy with its color's legend value (white is never a key); each named range on the sheet gets a clr\_ twin on the copy |
| [Fill Similar Background Color](#fill-similar-background-color) | Copies the selected cells' values over every cell sharing their fill color; turns a color\-coded map into data in one run |

### Formatting

| Name | Description |
| --- | --- |
| [Copy Map And Fill From Legend](#copy-map-and-fill-from-legend) | Copy a color\-coded map sheet: stand on a colored legend cell; copies the sheet and fills every map cell on the copy with its color's legend value (white is never a key); each named range on the sheet gets a clr\_ twin on the copy |
| [Fill Similar Background Color](#fill-similar-background-color) | Copies the selected cells' values over every cell sharing their fill color; turns a color\-coded map into data in one run |
| [Merged To Centre Across Selection](#merged-to-centre-across-selection) | Merged Cells changed to Centre Across Selection |

### GoTo

| Name | Description |
| --- | --- |
| [Goto Similar Background Color](#goto-similar-background-color) | Selects every cell in the selection that shares the active cell's fill color, so a color group can be seen or counted at once |

### LAMBDA

| Name | Description |
| --- | --- |
| [Clear Lambda Library](#clear-lambda-library) | Forgets your saved lambda library, so ILL and Full Setup Case stop loading it |
| [Import Case Lambdas](#import-case-lambdas) | Your library (if set) and CompBot's lambdas, the winner first, into the active workbook. Chained after Full Setup Case |
| [Import Lambda Library](#import-lambda-library) | Loads YOUR own lambda library (set once with SLL) into the active workbook: its lambdas and its stored values (named constants) |
| [Import Lambdas From CompBot](#import-lambdas-from-compbot) | Loads CompBot's own lambdas into the active workbook, replacing older copies unless SCS says your library wins |
| [Set Lambda Library](#set-lambda-library) | Points CompBot at YOUR lambda library workbook, once; ILL and Full Setup Case then load it |

### Lookup

| Name | Description |
| --- | --- |
| [Lookup Array from Column N](#lookup-array-from-column-n) | Looks every value of the active array up in a named range or table and returns the column you choose |
| [Lookup Value by Row](#lookup-value-by-row) | Returns RC location of a list of values in a range |

### Maintenance

| Name | Description |
| --- | --- |
| [Clear Lambdas](#clear-lambdas) | DEV ONLY, DISABLED. Clears lambdas out of CompBot itself (never the active file) that are not in the Lambdas table or needed by a command. Would delete any lambda library a user has pulled into CompBot. |

### Map

| Name | Description |
| --- | --- |
| [Copy Map And Fill From Legend](#copy-map-and-fill-from-legend) | Copy a color\-coded map sheet: stand on a colored legend cell; copies the sheet and fills every map cell on the copy with its color's legend value (white is never a key); each named range on the sheet gets a clr\_ twin on the copy |

### MM

| Name | Description |
| --- | --- |
| [MM](#mm) | Applies \-\- to an array, turning TRUE\/FALSE into 1\/0 so it can be summed or multiplied |

### Name

| Name | Description |
| --- | --- |
| [Name All Used Ranges](#name-all-used-ranges) | Renames used range in all sheets |
| [Name Used Ranges](#name-used-ranges) | Names used range in all sheets starting with the prefix as provided in the active cell |

### Navigation

| Name | Description |
| --- | --- |
| [Exclude Examples](#exclude-examples) | Trims the worked\-example rows off the top of the selection, leaving just the real questions |
| [Go To Example](#go-to-example) | Goes to the level's example row, two cells right of its last input, ready to solve |

### Paste

| Name | Description |
| --- | --- |
| [Paste Flattened List With Formatting](#paste-flattened-list-with-formatting) | Turns a copied range into a list, one row per cell, carrying fill color, font color and border flags |
| [Save Answers To Left](#save-answers-to-left) | Saves references to the selected cells in the green answer cells to the left on the same row. |

### Prep

| Name | Description |
| --- | --- |
| [Backup Sheets in Workbook](#backup-sheets-in-workbook) | Copy all sheets in workbook for backup |
| [Clear Lambda Library](#clear-lambda-library) | Forgets your saved lambda library, so ILL and Full Setup Case stop loading it |
| [Clear Solve Folder](#clear-solve-folder) | Forgets your Solve folder, so Full Setup Case saves the \_Solve copy next to the case again |
| [Create Blank Sheet](#create-blank-sheet) | Creates a blank sheet named based on cell value (Sht otherwise) |
| [Create Bonus Sheet](#create-bonus-sheet) | Creates bonus sheet "B" with bonus questions |
| [Create Case Inputs Sheet](#create-case-inputs-sheet) | Creates a case inputs sheet: one row per game with its Game and Level, then every input; names each level L01\_Inputs, L02\_Inputs, ... |
| [Create Data Table](#create-data-table) | Builds (or finds) the data table block for the level you are in, beside its inputs, and lands on the last input; fill it with Run Data Table (RDT) |
| [Create Level Sheets](#create-level-sheets) | Creates a sheet for each level in the Case sheet |
| [Full Setup Case](#full-setup-case) | Sets up the case: saves a \_Solve copy, backs up the sheets, creates the level, bonus and Case inputs sheets, then imports lambdas. Choose the steps with SCS |
| [Import Case Lambdas](#import-case-lambdas) | Your library (if set) and CompBot's lambdas, the winner first, into the active workbook. Chained after Full Setup Case |
| [Import Lambda Library](#import-lambda-library) | Loads YOUR own lambda library (set once with SLL) into the active workbook: its lambdas and its stored values (named constants) |
| [Import Lambdas From CompBot](#import-lambdas-from-compbot) | Loads CompBot's own lambdas into the active workbook, replacing older copies unless SCS says your library wins |
| [Level Inputs](#level-inputs) | Type a level, or levels like 4\-5 or 3,5\-6: their inputs land as an array to work on, with headers and game\/level beside them (blank \= this sheet's level) |
| [Load Game Into Calculation](#load-game-into-calculation) | Loads the selected data table game into the example calculation, or restores the example |
| [Merged To Centre Across Selection](#merged-to-centre-across-selection) | Merged Cells changed to Centre Across Selection |
| [Rename Sheets](#rename-sheets) | Shortens multi\-word sheet names to their initials, keeping numbers (Level 1 Data becomes L1D); backups keep their names |
| [Run Data Table](#run-data-table) | Fills a Create Data Table block with static results game by game, recalculating only when run |
| [Save Copy of File](#save-copy-of-file) | Enable editing and save copy of file with suffix based on active cell (otherwise Working) |
| [Save From Example](#save-from-example) | Select your answer formula on the example row; the working right of the inputs is copied to every question and the answers linked |
| [Set Lambda Library](#set-lambda-library) | Points CompBot at YOUR lambda library workbook, once; ILL and Full Setup Case then load it |
| [Set Solve Folder](#set-solve-folder) | Choose the folder Full Setup Case saves your \_Solve copy into, once (e.g. OneDrive) |
| [Setup Case Settings](#setup-case-settings) | Opens CompBot's Setup Settings sheet: choose which steps Full Setup Case runs, and whose lambdas win a name clash |
| [Unmerge Multi\-Row Merges](#unmerge-multi-row-merges) | Unmerges every merged area spanning more than one row (the selection, or the whole sheet), leaving the value in the top\-left cell |

### Settings

| Name | Description |
| --- | --- |
| [Get Current Settings](#get-current-settings) | Get current settings for this computer. |
| [Revert Settings](#revert-settings) | Reverts settings to those loaded into the Loaded column. |
| [Setup Case Settings](#setup-case-settings) | Opens CompBot's Setup Settings sheet: choose which steps Full Setup Case runs, and whose lambdas win a name clash |
| [Toggle Calculation Mode](#toggle-calculation-mode) | Toggles calculation mode and places current mode notice in StatusBar |
| [Toggle Iterative Calculation](#toggle-iterative-calculation) | Toggles iterative calculation and sets status in status bar |
| [Update Settings](#update-settings) | Updates default settings to those in Regional Settings sheet |

### WrapWith

| Name | Description |
| --- | --- |
| [Align Array to Right](#align-array-to-right) | Aligns array to the right with blanks to left |
| [Apply ARROWSHIFT\_byEmilieWilliams Lambda](#apply-arrowshift_byemiliewilliams-lambda) | Apply ARROWSHIFT\_byEmilieWilliams lambda function to active cell, using direction in clipboard. |
| [Difference of Array Columns by Row](#difference-of-array-columns-by-row) | Returns first column minus last column in array |
| [Find Above In Left](#find-above-in-left) | Select the empty grid between two lists: marks 1 where the item ABOVE each column is found in the item LEFT of each row |
| [Keep Cell of Array](#keep-cell-of-array) | Reduces a spilled array to just the one cell you selected, so a single value can be inspected or reused |
| [Lookup Array from Column N](#lookup-array-from-column-n) | Looks every value of the active array up in a named range or table and returns the column you choose |
| [Lookup Value by Row](#lookup-value-by-row) | Returns RC location of a list of values in a range |
| [Modulo Array](#modulo-array) | Wraps the active array with MOD; turns cumulative movement into a position on a looping board |
| [Modulo Array 1\-Based](#modulo-array-1-based) | Wraps the active array in BoardGameMove\_byHadynWiseman: a board numbered 1 to N, so an exact lap lands on N rather than 0 |
| [Negative Values](#negative-values) | Negative values. with errors as blank |
| [Running Total](#running-total) | Wraps the active array in a running total (RunningTotal\_byJaqKennedy with nothing else set) |
| [Running Total with Cap](#running-total-with-cap) | Running total that stops at the COPIED limit and holds there (Cap \= 1). Change the final 1 to 0 to keep the crossing row's full value instead |
| [Running Total with Reset](#running-total-with-reset) | Running total that restarts wherever the COPIED flag range holds 1; asks what it restarts at (blank \= 0). The flagged row's own value is included |
| [Sequence of Row Count](#sequence-of-row-count) | Sequence of row count of array variable |
| [Split Text by Delimiter Above](#split-text-by-delimiter-above) | Splits the active cell's text DOWN into rows on the delimiter in the cell above (linked, so changing that cell re\-splits), and trims each piece |
| [Split Text by Semicolon](#split-text-by-semicolon) | Splits the active cell's text on semicolons into an array, trimming each piece |
| [Wrap in ABS](#wrap-in-abs) | Wraps the active formula in ABS() to make the result positive: distances, differences, gaps, reflections |
| [Wrap in Concat](#wrap-in-concat) | Wraps the active formula in CONCAT() to join an array into one string with no separator, rebuilding a word from its letters |
| [Wrap in Drop First Row](#wrap-in-drop-first-row) | Wraps the active formula in DROP(...,1) to remove a header row from a spilled array |
| [Wrap in Take by Copied Cell Columns](#wrap-in-take-by-copied-cell-columns) | Wrap in take by copied cell columns. |
| [Wrap in UNICHAR](#wrap-in-unichar) | Wraps the active formula in UNICHAR() to turn code numbers back into characters: decoding glyphs, emoji or dice faces |
| [Wrap in UNICODE](#wrap-in-unicode) | Wraps the active formula in UNICODE() to turn characters into their code numbers, the first step in most cipher and glyph puzzles |
| [Wrap in UNIQUE](#wrap-in-unique) | Wraps the active formula in UNIQUE() to strip duplicates: distinct values, distinct colors, distinct answers |
| [Wrap with IFERROR TEXTBEFORE space](#wrap-with-iferror-textbefore-space) | Wraps current formula with TEXTBEFORE space with IFERROR in case space doesn't exist |

### Other

| Name | Description |
| --- | --- |
| [Clear Unused Lambdas](#clear-unused-lambdas) | Deletes lambdas from the active workbook that nothing refers to, so the case file you submit does not carry every lambda you own |
| [Clear Unused Lambdas with Report](#clear-unused-lambdas-with-report) | As Clear Unused Lambdas, but writes out what it removed and what it kept; use this one when you want to check before trusting it |
| [Eggy](#eggy) | Easter Egg Fun |
| [Exclude Blanks](#exclude-blanks) | Filters blank cells out of the active array, leaving only cells that hold something |
| [Extract Characters](#extract-characters) | Split the active cell's text into its individual characters, one per row. Keeps line breaks, accented letters and emoji intact. |
| [List Lambdas Used](#list-lambdas-used) | Lists every lambda the active workbook's formulas actually call; worth running before clearing anything |
| [Sum False](#sum-false) | Counts the FALSE (0) values in an array: how many rows fail the test. Anything non\-zero, such as 2 from an OR built with +, counts as TRUE |
| [Sum True](#sum-true) | Counts the TRUE values in an array (applies \-\- and sums it): the usual way to answer 'how many rows pass this test' |

<BR>

## Available Robot Texts

| Name | Description |
| --- | --- |
| [ADDRESSES\_byDiarmuidEarly.lambda](#addresses_bydiarmuidearlylambda) | Definition of ADDRESSES\_byDiarmuidEarly lambda function. |
| [ARROWSHIFT.lambda](#arrowshiftlambda) | Definition of ARROWSHIFT lambda function. |
| [ARROWSHIFT\_byEmilieWilliams.lambda](#arrowshift_byemiliewilliamslambda) | Definition of ARROWSHIFT\_byEmilieWilliams lambda function. |
| [BiCol\_byHadynWiseman.lambda](#bicol_byhadynwisemanlambda) | Definition of BiCol\_byHadynWiseman lambda function. |
| [BiRow\_byHadynWiseman.lambda](#birow_byhadynwisemanlambda) | Definition of BiRow\_byHadynWiseman lambda function. |
| [BoardGameMove\_byHadynWiseman.lambda](#boardgamemove_byhadynwisemanlambda) | Definition of BoardGameMove\_byHadynWiseman lambda function. |
| [ClosestOnMap\_byHadynWiseman.lambda](#closestonmap_byhadynwisemanlambda) | Definition of ClosestOnMap\_byHadynWiseman lambda function. |
| [ColNum\_byHadynWiseman.lambda](#colnum_byhadynwisemanlambda) | Definition of ColNum\_byHadynWiseman lambda function. |
| [Combinations\_byHadynWiseman.lambda](#combinations_byhadynwisemanlambda) | Definition of Combinations\_byHadynWiseman lambda function. |
| [Dice\_byHadynWiseman.lambda](#dice_byhadynwisemanlambda) | Definition of Dice\_byHadynWiseman lambda function. |
| [DiffByRow\_byJaqKennedy.lambda](#diffbyrow_byjaqkennedylambda) | Definition of DiffByRow\_byJaqKennedy lambda function. |
| [Distance\_byHadynWiseman.lambda](#distance_byhadynwisemanlambda) | Definition of Distance\_byHadynWiseman lambda function. |
| [Exists\_byHadynWiseman.lambda](#exists_byhadynwisemanlambda) | Definition of Exists\_byHadynWiseman lambda function. |
| [Extract\_byJaqKennedy.lambda](#extract_byjaqkennedylambda) | Definition of Extract\_byJaqKennedy lambda function. |
| [FilterArray\_byErikOehm.lambda](#filterarray_byerikoehmlambda) | Definition of FilterArray\_byErikOehm lambda function. |
| [FindInMap\_byHadynWiseman.lambda](#findinmap_byhadynwisemanlambda) | Definition of FindInMap\_byHadynWiseman lambda function. |
| [Flip\_byHadynWiseman.lambda](#flip_byhadynwisemanlambda) | Definition of Flip\_byHadynWiseman lambda function. |
| [GetAddress.lambda](#getaddresslambda) | Definition of GetAddress lambda function. |
| [GetAddressesByLookup.lambda](#getaddressesbylookuplambda) | Definition of GetAddressesByLookup lambda function. |
| [GRIDTOCOL\_byLiannaGerrish.lambda](#gridtocol_byliannagerrishlambda) | Definition of GRIDTOCOL\_byLiannaGerrish lambda function. |
| [IFBLANK.lambda](#ifblanklambda) | Definition of IFBLANK lambda function. |
| [IsInList\_byErikOehm.lambda](#isinlist_byerikoehmlambda) | Definition of IsInList\_byErikOehm lambda function. |
| [LkpRC.lambda](#lkprclambda) | Definition of LkpRC lambda function. |
| [LkpRCByRow.lambda](#lkprcbyrowlambda) | Definition of LkpRCByRow lambda function. |
| [MazeDistance\_byHadynWiseman.lambda](#mazedistance_byhadynwisemanlambda) | Definition of MazeDistance\_byHadynWiseman lambda function. |
| [MM\_byHaDang.lambda](#mm_byhadanglambda) | Definition of MM\_byHaDang lambda function. |
| [PerCom\_byBoRydobon.lambda](#percom_byborydobonlambda) | Definition of PerCom\_byBoRydobon lambda function. |
| [Regions\_byHadynWiseman.lambda](#regions_byhadynwisemanlambda) | Definition of Regions\_byHadynWiseman lambda function. |
| [RightAlignedArray\_byJaqKennedy.lambda](#rightalignedarray_byjaqkennedylambda) | Definition of RightAlignedArray\_byJaqKennedy lambda function. |
| [Rotate\_byHadynWiseman.lambda](#rotate_byhadynwisemanlambda) | Definition of Rotate\_byHadynWiseman lambda function. |
| [RowNum\_byHadynWiseman.lambda](#rownum_byhadynwisemanlambda) | Definition of RowNum\_byHadynWiseman lambda function. |
| [RunningTotal\_byJaqKennedy.lambda](#runningtotal_byjaqkennedylambda) | Definition of RunningTotal\_byJaqKennedy lambda function. |
| [SplitText\_byHadynWiseman.lambda](#splittext_byhadynwisemanlambda) | Definition of SplitText\_byHadynWiseman lambda function. |

<BR>

## Command Definitions

<BR>

### Align Array to Right

*Aligns array to the right with blanks to left*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=RightAlignedArray\_byJaqKennedy(IFBLANK(\[\[ActiveCell::Formula\]\],""))</code> |
| Formula Dependencies | <ol><li>[RightAlignedArray_byJaqKennedy.lambda](#rightalignedarray_byjaqkennedylambda)</li><li>[IFBLANK.lambda](#ifblanklambda)</li></ol> |
| Launch Codes | <code>ar</code> |

[^Top](#oa-robot-definitions)

<BR>

### Apply ARROWSHIFT\_byEmilieWilliams Lambda

*Apply ARROWSHIFT\_byEmilieWilliams lambda function to active cell, using direction in clipboard.*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=ARROWSHIFT\_byEmilieWilliams(\[\[ActiveCell::Formula\]\], \[\[Clipboard\]\])</code> |
| Formula Dependencies | [ARROWSHIFT_byEmilieWilliams.lambda](#arrowshift_byemiliewilliamslambda) |
| Launch Codes | <code>ARR</code> |

[^Top](#oa-robot-definitions)

<BR>

### Backup Sheets in Workbook

*Copy all sheets in workbook for backup*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Prep`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCaseSetup.Backup](./VBA/modCaseSetup.bas#L418)()</code> |
| Macro Workbook Connection | ThisWorkbook |
| Launch Codes | <code>BU</code> |

[^Top](#oa-robot-definitions)

<BR>

### Clear Bonus From Dock

*Hides one bonus from the dock by hand (type 2, B4, B or Bonus 2)*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Bonus`</sup>

> \*\*Note:\*\* For bonuses skipped or answered somewhere the dock cannot see. Answered bonuses drop off by themselves. Remembered in the case workbook as a hidden name BonusDock\_Cleared; Restore Cleared Bonuses (BQR) brings them all back. Refreshes the dock afterwards.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modBonusDock.ClearBonusFromDock](./VBA/modBonusDock.bas#L176)({{bonus_to_clear}})</code> |
| Parameters | <ol><li>[bonus_to_clear](#clear-bonus-from-dock--bonus_to_clear)</li></ol> |
| Command After | [Show Bonus Dock](#show-bonus-dock) |
| Launch Codes | <code>BQX</code> |

<BR>

#### Clear Bonus From Dock \>\> bonus\_to\_clear

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Bonus to clear from the dock (e.g. 2, B4, B or Bonus 2)</code> |
| Data Type | String |
| Caching Policy | CacheForDurationOfCommandExecution |

[^Top](#oa-robot-definitions)

<BR>

### Clear Lambda Library

*Forgets your saved lambda library, so ILL and Full Setup Case stop loading it*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#LAMBDA` `#Prep`</sup>

> \*\*Note:\*\* Removes the per\-user setting Set Lambda Library (SLL) saved. Changes no workbook and no lambda: the library file itself is untouched. After it, Full Setup Case loads CompBot's lambdas only and ILL reports that no library is set. To SWAP one library for another you do not need this: run SLL again and pick the new file. Added 2026\-09\-24 (Jaq).

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdas.ClearLambdaLibrary](./VBA/modLambdas.bas#L140)()</code> |
| Launch Codes | <code>CLL</code> |

[^Top](#oa-robot-definitions)

<BR>

### Clear Lambdas

*DEV ONLY, DISABLED. Clears lambdas out of CompBot itself (never the active file) that are not in the Lambdas table or needed by a command. Would delete any lambda library a user has pulled into CompBot.*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Maintenance`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdas.ClearLambdas](./VBA/modLambdas.bas#L656)()</code> |
| Enabled | ☐Yes ☑No |

[^Top](#oa-robot-definitions)

<BR>

### Clear Solve Folder

*Forgets your Solve folder, so Full Setup Case saves the \_Solve copy next to the case again*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Prep`</sup>

> \*\*Note:\*\* Removes the per\-user setting Set Solve Folder (SSF) saved. Changes no folder and no file. To switch to a different folder you do not need this: run SSF again. Added 2026\-10\-08 (GitHub \#4).

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modSetupSettings.ClearSolveFolder](./VBA/modSetupSettings.bas#L135)()</code> |
| Launch Codes | <code>CSF</code> |

[^Top](#oa-robot-definitions)

<BR>

### Clear Status Bar

*Clears the status bar*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Bonus`</sup>

> \*\*Note:\*\* Clears the status bar (Application.StatusBar \= False), so Excel's own 'Ready' and 'Average \/ Count \/ Sum' come back. Several CompBot commands report on the status bar rather than in a dialog, and a message stays until something clears it. Renamed 2026\-09\-24 from Clear Bonus Status Bar; the old BQS code was dropped the same day (Jaq: nobody else has used it).

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modBonusDock.ClearBonusStatusBar](./VBA/modBonusDock.bas#L361)()</code> |
| Launch Codes | <code>CSB</code> |

[^Top](#oa-robot-definitions)

<BR>

### Clear Unused Lambdas

*Deletes lambdas from the active workbook that nothing refers to, so the case file you submit does not carry every lambda you own*

<sup>`@CompBot.xlsm` `!VBA Macro Command` </sup>

> \*\*Note:\*\* Creates a report on used lambdas and clears out unused lambdas from the workbook

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdas.ClearUnusedLambdas](./VBA/modLambdas.bas#L1575)()</code> |

[^Top](#oa-robot-definitions)

<BR>

### Clear Unused Lambdas with Report

*As Clear Unused Lambdas, but writes out what it removed and what it kept; use this one when you want to check before trusting it*

<sup>`@CompBot.xlsm` `!VBA Macro Command` </sup>

> \*\*Note:\*\* Creates a report on used lambdas and clears out unused lambdas from the workbook, giving a report on deleted lambdas

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdas.ClearUnusedLambdas](./VBA/modLambdas.bas#L1575)(1)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Copy Map And Fill From Legend

*Copy a color\-coded map sheet: stand on a colored legend cell; copies the sheet and fills every map cell on the copy with its color's legend value (white is never a key); each named range on the sheet gets a clr\_ twin on the copy*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Fill` `#Formatting` `#Color` `#Map`</sup>

> \*\*Note:\*\* 2026\-10\-05 (Jaq): copy + Fill Similar Background Color + clr\_ names in one run. Legend \= the unbroken run of colored cells through the active cell (down the column, else along the row), or the selection when several cells are selected. A swatch with no value takes the value to its right. Exact color match, one pass, white\/no\-fill excluded.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modGoToSpecial.CopyMapFillFromLegend](./VBA/modGoToSpecial.bas#L137)()</code> |
| Launch Codes | <code>CFL</code> |

[^Top](#oa-robot-definitions)

<BR>

### Create Blank Sheet

*Creates a blank sheet named based on cell value (Sht otherwise)*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Prep`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modMisc.CreateBlankSheet](./VBA/modMisc.bas#L191)([[ActiveCell]])</code> |
| Macro Workbook Connection | ThisWorkbook |
| Launch Codes | <ol><li><code>BS</code></li><li><code>CBS</code></li><li><code>S</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Create Bonus Sheet

*Creates bonus sheet "B" with bonus questions*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Prep`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCaseSetup.CreateBonusSheet](./VBA/modCaseSetup.bas#L859)()</code> |
| Launch Codes | <ol><li><code>BQC</code></li><li><code>CB</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Create Case Inputs Sheet

*Creates a case inputs sheet: one row per game with its Game and Level, then every input; names each level L01\_Inputs, L02\_Inputs, ...*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Prep`</sup>

> \*\*Note:\*\* Run by Full Setup Case. Column B is the game number and column C its LEVEL (from the case's own Level column; a case with none gets no levels), so the header row's filter shows one level at a time; the inputs follow from column D. Each level's rows are also named L01\_Inputs, L02\_Inputs, ... (Game, Level and the input columns that level uses), for the Name Box and for Level Inputs (LI), which drops one or more levels' inputs into the active cell. Level column and names added 2026\-10\-08 (GitHub \#3, Ben de Leon). The same day its unused MacroWorkbookConnection reference ('ThisWorkbook', which named no Connection) was removed, so it runs like every other CompBot macro command.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCaseSetup.CreateCaseInputsSheet](./VBA/modCaseSetup.bas#L1329)([[ActiveCell]])</code> |
| Launch Codes | <ol><li><code>CIS</code></li><li><code>IS</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Create Data Table

*Builds (or finds) the data table block for the level you are in, beside its inputs, and lands on the last input; fill it with Run Data Table (RDT)*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Prep`</sup>

> \*\*Note:\*\* Run from any cell in a level: on the Case sheet, on an L\#\# sheet, or on the level's Example\# cell on any sheet. The new block goes level with the example row, as close as it can to two columns right of the level's last input, moved right only as far as it needs to overlap nothing. If the level already has a block, nothing is built: it takes you to that block's last input. Either way the view scrolls to the block, and you land on the last input cell, ready to build the example calculation beside it; then point ANSWER at your output and run RDT. Refuses when the level's answers already link to a block on another sheet (solve it there). Static results, no What\-If table. Merged 2026\-09\-24 (Jaq): Solve Level's finding and placement under this command's name, codes and end state; it no longer asks you to point at a target cell.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modDataTable.CreateDataTable](./VBA/modDataTable.bas#L476)()</code> |
| Launch Codes | <ol><li><code>CDT</code></li><li><code>DT</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Create Level Sheets

*Creates a sheet for each level in the Case sheet*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Prep`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCaseSetup.CreateLevelSheets](./VBA/modCaseSetup.bas#L576)()</code> |
| Launch Codes | <ol><li><code>CL</code></li><li><code>CLS</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Difference of Array Columns by Row

*Returns first column minus last column in array*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=DiffByRow\_byJaqKennedy(\[\[ActiveCell::Formula\]\])</code> |
| Formula Dependencies | [DiffByRow_byJaqKennedy.lambda](#diffbyrow_byjaqkennedylambda) |
| Launch Codes | <ol><li><code>diff</code></li><li><code>dbr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Eggy

*Easter Egg Fun*

<sup>`@CompBot.xlsm` `!VBA Macro Command` </sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modUtilities.Eggy](./VBA/modUtilities.bas#L247)()</code> |

[^Top](#oa-robot-definitions)

<BR>

### Exclude Blanks

*Filters blank cells out of the active array, leaving only cells that hold something*

<sup>`@CompBot.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=FILTER(\[\[ActiveCell::Formula\]\],\[\[ActiveCell::Formula\]\]\<\>"")</code> |
| Launch Codes | <code>EB</code> |

[^Top](#oa-robot-definitions)

<BR>

### Exclude Examples

*Trims the worked\-example rows off the top of the selection, leaving just the real questions*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Navigation`</sup>

> \*\*Note:\*\* Promoted into CompBot from A\-ZTraining 2026\-09\-22: the command review found it needed in 23 of 67 surveyed cases, the 2nd most\-needed of Jaq's own commands, while living in the training collection rather than the competition one. NOT a straight copy; the A\-Z original drops exactly two rows, assuming one example row plus one blank spacer. This walks the rows instead, dropping example\-label rows and blank rows until it reaches real content, so two or three worked rows per level (Example3a \/ 3b, common in 2026 cases) and a missing spacer all behave. It recognises Example \/ Exemple \/ Exemplo \/ Sample via the shared list in modCaseSetup, spaces and all. If it cannot read the label column it falls back to dropping two rows, exactly as the original did. LAUNCH CODE: deliberately the SAME code (XE) as A\-ZTraining's copy. Jaq, 2026\-09\-22: 'it doesn't matter which it runs from, and I'm one of few who'll have both collections.' If you ever need to tell them apart, this is the one that handles multiple example rows.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCaseNav.ExcludeExamples](./VBA/modCaseNav.bas#L47)()</code> |
| Launch Codes | <code>XE</code> |

[^Top](#oa-robot-definitions)

<BR>

### Extract Characters

*Split the active cell's text into its individual characters, one per row. Keeps line breaks, accented letters and emoji intact.*

<sup>`@CompBot.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SplitText\_byHadynWiseman(\[\[ActiveCell::Formula\]\])</code> |
| Formula Dependencies | [SplitText_byHadynWiseman.lambda](#splittext_byhadynwisemanlambda) |

[^Top](#oa-robot-definitions)

<BR>

### Fill Similar Background Color

*Copies the selected cells' values over every cell sharing their fill color; turns a color\-coded map into data in one run*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Fill` `#Formatting`</sup>

> \*\*Note:\*\* Fills each cell on the sheet that has the same background color with either the value in the cell, or if empty, the value to the right of the cell.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modGoToSpecial.FillSimilarBackgroundColor](./VBA/modGoToSpecial.bas#L75)()</code> |
| Launch Codes | <code>FBC</code> |

[^Top](#oa-robot-definitions)

<BR>

### Find Above In Left

*Select the empty grid between two lists: marks 1 where the item ABOVE each column is found in the item LEFT of each row*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=1\-ISERROR(FIND(\[\[Selection.Rows(1).Offset(\-1,0)::Address\]\],\[\[Selection.Columns(1).Offset(0,\-1)::Address\]\]))</code> |
| Destination Range Address | <code>\[\[Selection.Cells(1,1)\]\]</code> |
| Launch Codes | <code>fal</code> |

[^Top](#oa-robot-definitions)

<BR>

### Find Address Of Value On Map

*Returns the cell ADDRESS of each value you are looking for on a map, the start point the distance commands need*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#Array`</sup>

> \*\*Note:\*\* Point it at a map and it gives back WHERE the things you asked for are, as addresses like "H12". One match comes back as a single address; several come back as a column of them; nothing found gives "Not Found". WHY IT MATTERS MORE THAN IT LOOKS: Find Nearest On Map (FNM) and Map Distance From Cell (MDC) both need a STARTING ADDRESS as text, and on a real case you rarely know it by eye; the start is 'wherever the S is'. This is the command that produces it, so the three chain together: FAM to locate the start, then MDC or FNM from there. ARGUMENT ORDER, if you ever type the lambda by hand: it is FindInMap(items, map), so the things you are looking for come FIRST, the map second, which is the opposite way round from most of the map lambdas. Wraps FindInMap\_byHadynWiseman (which itself uses GridToCol\_byLiannaGerrish); owned by CompBot since the lambda review, exposed as a command 2026\-09\-22.

| Property | Value |
| --- | --- |
| Formula | <code>\=LET(\_l,"{{LookFor}}",FindInMap\_byHadynWiseman(IFERROR(\-\-\_l,\_l),\[\[ActiveCell::Formula\]\]))</code> |
| Formula Dependencies | <ol><li>[FindInMap_byHadynWiseman.lambda](#findinmap_byhadynwisemanlambda)</li><li>[GRIDTOCOL_byLiannaGerrish.lambda](#gridtocol_byliannagerrishlambda)</li></ol> |
| Parameters | <ol><li>[LookFor](#find-address-of-value-on-map--lookfor)</li></ol> |
| User Context Filter | ExcelSelectionIsSingleCell |
| Launch Codes | <code>FAM</code> |

<BR>

#### Find Address Of Value On Map \>\> LookFor

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Value to find, typed plainly (e.g. S, no quotes)</code> |
| Data Type | String |

[^Top](#oa-robot-definitions)

<BR>

### Find Nearest On Map

*Stand on the start (e.g. the S) in a map spill, with the target cell COPIED: lists the nearest targets with their addresses and distances, to the right of the sheet's contents*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#Array`</sup>

> \*\*Note:\*\* Returns three columns (Value, Address, Distance) for every matching cell within MaxSteps of the start, SORTED NEAREST FIRST, so the first row is the answer to 'what is closest'. WHAT IT IS FOR: 'which portal is nearest', 'what is the closest resource', 'how far to the nearest wall'. The 2024 and 2025 MEWC Portal cases are the worked example; this answers their level 5 outright. MaxSteps keeps it fast by only looking at a window around the start rather than the whole map. Wraps ClosestOnMap\_byHadynWiseman, which in turn uses Distance\_byHadynWiseman. Both have been in CompBot's library since the lambda review; no command exposed them until 2026\-09\-22. THE REASON A WRAPPER MATTERS HERE: the lambda takes ELEVEN parameters. It is not something anyone reconstructs correctly from memory in a timed round, so the capability was effectively out of reach even though it was already owned and tested.

| Property | Value |
| --- | --- |
| Formula | <code>\=LET(\_m,\[\[ActiveCell.SpillParent::Formula\]\],ClosestOnMap\_byHadynWiseman(\_m,FindInMap\_byHadynWiseman(\[\[ActiveCell\]\],\_m),ROWS(\_m)+COLUMNS(\_m),\[\[Clipboard::Address\]\]))</code> |
| Destination Range Address | <code>\[\[AvailableCellToRight\]\]</code> |
| Formula Dependencies | <ol><li>[ClosestOnMap_byHadynWiseman.lambda](#closestonmap_byhadynwisemanlambda)</li><li>[RowNum_byHadynWiseman.lambda](#rownum_byhadynwisemanlambda)</li><li>[ColNum_byHadynWiseman.lambda](#colnum_byhadynwisemanlambda)</li><li>[ADDRESSES_byDiarmuidEarly.lambda](#addresses_bydiarmuidearlylambda)</li><li>[Distance_byHadynWiseman.lambda](#distance_byhadynwisemanlambda)</li><li>[FindInMap_byHadynWiseman.lambda](#findinmap_byhadynwisemanlambda)</li><li>[GRIDTOCOL_byLiannaGerrish.lambda](#gridtocol_byliannagerrishlambda)</li></ol> |
| User Context Filter | ExcelSelectionIsSingleCell |
| Launch Codes | <code>FNM</code> |

[^Top](#oa-robot-definitions)

<BR>

### First Match By Row

*For each row, the position of the first cell equal to the COPIED cell's value, or a fallback (blank \= never) where the row never matches*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#Array`</sup>

> \*\*Note:\*\* Promoted into CompBot from A\-ZTraining 2026\-09\-22. Answers 'on which step did this row first hit X': a set completed, a target reached, a state entered. Returns the 1\-based column position; an r x c array becomes r x 1. 2026\-09\-28 (Jaq): the value to find is the COPIED cell (linked, no typing, no quotes); the one popup is the not\-found result, blank \= never. WHY IT EARNS A PLACE IN THE COMPETITION COLLECTION: chained with Max Of Array By Row (Array Robot, free) it gives a STABLE ARGMAX, highest value, ties going to the earliest position. Needed in 7 of 67 surveyed cases. Self\-contained formula, no lambda dependency.

| Property | Value |
| --- | --- |
| Formula | <code>\=LET(\_a,(\[\[ActiveCell::Formula\]\]),\_v,\[\[Clipboard::Address\]\],\_nf,"{{IfNotFound}}",BYROW(\_a,LAMBDA(\_r,IFERROR(XMATCH(\_v,\_r),IF(\_nf\="","never",IFERROR(\-\-\_nf,\_nf))))))</code> |
| Parameters | <ol><li>[IfNotFound](#first-match-by-row--ifnotfound)</li></ol> |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <code>FMR</code> |

<BR>

#### First Match By Row \>\> IfNotFound

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Result where the row never matches (leave blank for never)</code> |
| Data Type | String |

[^Top](#oa-robot-definitions)

<BR>

### Flip Array

*Reverses the active cell's text (stressed becomes desserts). For a block use FLR (left to right) or FTB (top to bottom)*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#Array`</sup>

> \*\*Note:\*\* Mirrors a grid. Horizontal 1 flips left\-right, Vertical 1 flips top\-bottom, and 0 for BOTH gives the 180\-degree point mirror (every cell swapped with its opposite). Unlike Rotate Array the shape never changes. WHAT IT IS FOR: a board read from the other side, a reflected path, a mirrored tile. ONE NICE QUIRK: on a SINGLE CELL it reverses the TEXT of that cell, so it doubles as a reverse\-a\-string command. TWO THINGS NOT REACHABLE FROM THIS COMMAND, type the lambda by hand for them: DiagonalTopLeft (a transpose) and DiagonalTopRight (an anti\-transpose) are the 3rd and 4th arguments, and a 6th argument AddressMoves returns WHERE A GIVEN CELL MOVED TO rather than the flipped grid: Flip\_byHadynWiseman(array, 1, 0, , , "H12"). Wraps Flip\_byHadynWiseman; exposed as a command 2026\-09\-22.

| Property | Value |
| --- | --- |
| Formula | <code>\=Flip\_byHadynWiseman(\[\[ActiveCell::Formula\]\])</code> |
| Formula Dependencies | <ol><li>[Flip_byHadynWiseman.lambda](#flip_byhadynwisemanlambda)</li><li>[ADDRESSES_byDiarmuidEarly.lambda](#addresses_bydiarmuidearlylambda)</li></ol> |
| User Context Filter | ExcelSelectionIsSingleCell |
| Launch Codes | <code>FLP</code> |

[^Top](#oa-robot-definitions)

<BR>

### Flip Array Left to Right

*Mirrors the active array left to right: the last column comes first*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#Array`</sup>

> \*\*Note:\*\* 2026\-09\-28 (Jaq): one of the Flip family (FLP text, FLR left\-right, FTB top\-bottom), replacing Flip Array's two popups.

| Property | Value |
| --- | --- |
| Formula | <code>\=Flip\_byHadynWiseman(\[\[ActiveCell::Formula\]\],1)</code> |
| Formula Dependencies | <ol><li>[Flip_byHadynWiseman.lambda](#flip_byhadynwisemanlambda)</li><li>[ADDRESSES_byDiarmuidEarly.lambda](#addresses_bydiarmuidearlylambda)</li></ol> |
| User Context Filter | ExcelSelectionIsSingleCell |
| Launch Codes | <code>FLR</code> |

[^Top](#oa-robot-definitions)

<BR>

### Flip Array Top to Bottom

*Mirrors the active array top to bottom: the last row comes first*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#Array`</sup>

> \*\*Note:\*\* 2026\-09\-28 (Jaq): one of the Flip family (FLP text, FLR left\-right, FTB top\-bottom).

| Property | Value |
| --- | --- |
| Formula | <code>\=Flip\_byHadynWiseman(\[\[ActiveCell::Formula\]\],,1)</code> |
| Formula Dependencies | <ol><li>[Flip_byHadynWiseman.lambda](#flip_byhadynwisemanlambda)</li><li>[ADDRESSES_byDiarmuidEarly.lambda](#addresses_bydiarmuidearlylambda)</li></ol> |
| User Context Filter | ExcelSelectionIsSingleCell |
| Launch Codes | <code>FTB</code> |

[^Top](#oa-robot-definitions)

<BR>

### Full Setup Case

*Sets up the case: saves a \_Solve copy, backs up the sheets, creates the level, bonus and Case inputs sheets, then imports lambdas. Choose the steps with SCS*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Prep`</sup>

> \*\*Note:\*\* Steps, in order: save a working copy as \<name\>\_Solve (skipped if the file has never been saved or is already a \_Solve \/ \_Working copy), back up every sheet, name the used ranges, create the level sheets, the bonus sheet B and the Case inputs sheet; then Import Case Lambdas runs (your library and CompBot's, the winner first). Any step can be switched off per user with Setup Case Settings (SCS); a switched\-off step is reported as skipped. A failed step is reported on the status bar and the rest still run. Does NOT unmerge or rename sheets: run Unmerge Multi\-Row Merges, M2C or Rename Sheets (rs) yourself first if the case needs them. Save copy step added 2026\-09\-24 (Jaq); step switches added 2026\-09\-24.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCaseSetup.Setup](./VBA/modCaseSetup.bas#L62)()</code> |
| Keyboard Shortcut | <code>^+g</code> |
| Command After | [Import Case Lambdas](#import-case-lambdas) |
| Launch Codes | <code>SC</code> |

[^Top](#oa-robot-definitions)

<BR>

### Get Current Settings

*Get current settings for this computer.*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Settings`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modMisc.GetCurrentSettings](./VBA/modMisc.bas#L367)()</code> |
| Launch Codes | <code>GS</code> |

[^Top](#oa-robot-definitions)

<BR>

### Go To Example

*Goes to the level's example row, two cells right of its last input, ready to solve*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Navigation`</sup>

> \*\*Note:\*\* From anywhere in a level (the Case sheet, an L\#\# sheet, or standing on an Example\# label), selects the cell two columns right of the level's last input on its example row, where a live solve starts, and scrolls to it. Finds the level exactly as Create Data Table does (shared LocateLevel) and the inputs by the one input rule (LevelInputColumns). Status bar report, no dialogs. Added 2026\-09\-24 to replace Select Example And Questions (SEQ), which suited A\-ZTraining replays rather than a live solve (Jaq).

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modDataTable.GoToExample](./VBA/modDataTable.bas#L736)()</code> |
| Launch Codes | <code>GE</code> |

[^Top](#oa-robot-definitions)

<BR>

### Goto Similar Background Color

*Selects every cell in the selection that shares the active cell's fill color, so a color group can be seen or counted at once*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#GoTo`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modGoToSpecial.GotoSimilarBackgroundColor](./VBA/modGoToSpecial.bas#L17)()</code> |
| Launch Codes | <code>SBC</code> |

[^Top](#oa-robot-definitions)

<BR>

### Import Case Lambdas

*Your library (if set) and CompBot's lambdas, the winner first, into the active workbook. Chained after Full Setup Case*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#LAMBDA` `#Prep`</sup>

> \*\*Note:\*\* Hidden: it exists to be Full Setup Case's CommandAfter. Loads the user's own lambda library (if Set Lambda Library has been run) and CompBot's lambdas, each only if switched on in Setup Case Settings (SCS). The winner (SCS; CompBot by default) goes first and replaces same\-named lambdas; the other goes second and never overwrites, so where names clash the winner's version is the one left. Its report is appended to Full Setup Case's own status\-bar line rather than replacing it. Added 2026\-09\-24; winner and switches 2026\-09\-24.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdas.ImportCaseLambdas](./VBA/modLambdas.bas#L252)()</code> |
| Visibility | Hidden |

[^Top](#oa-robot-definitions)

<BR>

### Import Lambda Library

*Loads YOUR own lambda library (set once with SLL) into the active workbook: its lambdas and its stored values (named constants)*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#LAMBDA` `#Prep`</sup>

> \*\*Note:\*\* Everyone keeps their own lambda library in a workbook of their own. Set Lambda Library (SLL) remembers where it is, once; this copies its lambdas, and its STORED VALUES (named constants and arrays such as \=0.05, \="Red", \={1,2,3}, and formula names with no cell reference), into whatever workbook is active. Names that point at ranges in the library are not copied (they would link back to the library file); the status bar counts them. It opens the library read\-only with its macros kept quiet and closes it again (or uses it as\-is if you already have it open). By default a lambda ALREADY in the file is skipped (left as it is) and listed on the status bar, so your copy does not replace CompBot's improved version; if Setup Case Settings (SCS) says your library wins, it replaces instead. Stored values follow the same rule; ILL marks the values it copies with \[CompBot ILL\] in their Name Manager comment, and those are the ones it updates. Refuses to run with CompBot or the library itself active. No dialogs; failures start ILL FAILED on the status bar. WHY NOT KEEP YOUR LAMBDAS INSIDE COMPBOT: a GitHub update replaces CompBot.xlsm wholesale, and most people run it hidden and read\-only anyway. The library location is a per\-user Windows setting, so it survives every update. Stored values added 2026\-10\-08 (GitHub \#5, Fabian Sjöblom).

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdas.ImportLambdaLibrary](./VBA/modLambdas.bas#L175)()</code> |
| Launch Codes | <code>ILL</code> |

[^Top](#oa-robot-definitions)

<BR>

### Import Lambdas From CompBot

*Loads CompBot's own lambdas into the active workbook, replacing older copies unless SCS says your library wins*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#LAMBDA` `#Prep`</sup>

> \*\*Note:\*\* Copies every lambda CompBot carries into the active workbook, with its description. By default a lambda of the same name already there is REPLACED, so an older case file picks up CompBot's improved versions; if Setup Case Settings (SCS) says your library wins, a same\-named lambda is skipped instead. A name that is not a lambda (the case's own range or constant) is never touched. Refuses to run with CompBot itself active. Reports on the status bar; no dialogs. Uses CompBot's own importer (modLambdas.CopyLambdas), shared with Import Lambda Library (ILL), so there is no dependency on another collection. Full Setup Case runs it for you through Import Case Lambdas.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdas.ImportLambdasFromCompBot](./VBA/modLambdas.bas#L211)()</code> |
| Launch Codes | <code>ILC</code> |

[^Top](#oa-robot-definitions)

<BR>

### Keep Cell of Array

*Reduces a spilled array to just the one cell you selected, so a single value can be inspected or reused*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=CHOOSEROWS(CHOOSECOLS(\[\[ActiveCell.SpillParent::Formula\]\],{{Selected\_Column\_Indexes\_In\_Spilling\_Range}}),{{Selected\_Row\_Indexes\_In\_Spilling\_Range}})</code> |
| Destination Range Address | <code>\[\[ActiveCell.SpillParent\]\]</code> |
| Launch Codes | <code>KTC</code> |

[^Top](#oa-robot-definitions)

<BR>

### Label Connected Regions (4 ways)

*Numbers each connected blob on a map, cells joining only up, down, left and right*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#Array`</sup>

> \*\*Note:\*\* 2026\-09\-28 (Jaq): replaces Label Connected Regions' 0\/1 popup. LCR8 lets a corner touch join two cells.

| Property | Value |
| --- | --- |
| Formula | <code>\=Regions\_byHadynWiseman(\[\[ActiveCell::Formula\]\],1)</code> |
| Formula Dependencies | <ol><li>[Regions_byHadynWiseman.lambda](#regions_byhadynwisemanlambda)</li><li>[Exists_byHadynWiseman.lambda](#exists_byhadynwisemanlambda)</li></ol> |
| User Context Filter | ExcelSelectionIsSingleCell |
| Launch Codes | <code>LCR4</code> |

[^Top](#oa-robot-definitions)

<BR>

### Label Connected Regions (8 ways)

*Numbers each connected blob on a map, a diagonal (corner) touch joining cells too*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#Array`</sup>

> \*\*Note:\*\* 2026\-09\-28 (Jaq): the diagonal twin of LCR4.

| Property | Value |
| --- | --- |
| Formula | <code>\=Regions\_byHadynWiseman(\[\[ActiveCell::Formula\]\],0)</code> |
| Formula Dependencies | <ol><li>[Regions_byHadynWiseman.lambda](#regions_byhadynwisemanlambda)</li><li>[Exists_byHadynWiseman.lambda](#exists_byhadynwisemanlambda)</li></ol> |
| User Context Filter | ExcelSelectionIsSingleCell |
| Launch Codes | <code>LCR8</code> |

[^Top](#oa-robot-definitions)

<BR>

### Level Inputs

*Type a level, or levels like 4\-5 or 3,5\-6: their inputs land as an array to work on, with headers and game\/level beside them (blank \= this sheet's level)*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Prep`</sup>

> \*\*Note:\*\* For solving bonuses that ask about one level's games: from any sheet, type the levels the way you would say them (4, 4\-5, 3,5\-6; any number of levels, so 5\-, 8\- and 10\-level cases work). From the active cell (say Z5): Game \| Level headers in Z5 with the game and level numbers as their own array in Z6\#; the INPUTS as an array of their own two columns right, in AB6\#, with their headers above in AB5. You finish on AB6, ready to work on the inputs. Input columns blank for every chosen level are left out, empty inputs stay blank rather than showing 0, and columns holding nothing but the result are auto\-fitted. The formulas are built to be read: LET names levels, caseInputs, levelRows (a FILTER using IsInList\_byErikOehm, which this command loads as a dependency), tidyRows, inputs, usedCols. The data is the CaseInputs table that Create Case Inputs Sheet (CIS) builds during Full Setup Case, with its Level column; a case with no Level column has no levels tagged and this command says so. Left blank on a level sheet (L06) it uses that level. It checks that the four anchor cells are empty before writing, and if a spill would run into anything the result is taken back out. Every outcome is on the status bar; a failure starts LEVEL INPUTS FAILED. Launch code LI (the name's initials too); LLI also works. Added 2026\-10\-08 (GitHub \#3, Ben de Leon).

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCaseNav.LoadLevelInputs](./VBA/modCaseNav.bas#L480)({{Levels}})</code> |
| Formula Dependencies | [IsInList_byErikOehm.lambda](#isinlist_byerikoehmlambda) |
| Parameters | <ol><li>[Levels](#level-inputs--levels)</li></ol> |
| Launch Codes | <ol><li><code>LI</code></li><li><code>LLI</code></li></ol> |

<BR>

#### Level Inputs \>\> Levels

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Levels, e.g. 4 or 4\-5 or 3,5\-6 (blank \= this sheet's level)</code> |
| Data Type | String |

[^Top](#oa-robot-definitions)

<BR>

### Link Answer Cell

*Points the level's data table ANSWER cell at the active cell (your calculation's output); then run RDT*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#DataTable`</sup>

> \*\*Note:\*\* The step between Create Data Table (CDT) and Run Data Table (RDT). Stand on the cell holding the example calculation's OUTPUT and run it: the ANSWER cell of this level's block (3 rows below the block's input cell, 1 column right) becomes \=that cell, absolute. The level is found exactly as CDT finds it, so on a Case sheet with several blocks the right one is used. Refuses when the level has no block yet, or when you are standing on the block itself. Copies nothing, moves nothing. Added 2026\-09\-24 (Jaq) as CompBot's own take on A\-ZTraining's ANS, which searched for the first green ANSWER cell of an Excel What\-If table.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modDataTable.LinkAnswerCell](./VBA/modDataTable.bas#L810)()</code> |
| Launch Codes | <code>ANS</code> |

[^Top](#oa-robot-definitions)

<BR>

### Link Side Cell

*Adds a SIDE column to the level's data table, pointed at the active cell, so RDT also records a second value per game for a bonus*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#DataTable`</sup>

> \*\*Note:\*\* For bonuses that need a DIFFERENT per\-game number from the level's calculation (a count, a per\-player score), then SUM \/ MAX \/ COUNTIF over the games. Stand on the cell holding that number and run it: the next SIDE column of this level's block goes in the first free column right of the block. Check row: SIDE k; ANSWER row: \=that cell, light blue; game rows: \#N\/A until RDT, which recalculates each game once and fills the answer and every SIDE column from the same pass. Up to 5 SIDE columns. If the column is not empty from the Check row down, nothing is written and an orange note above it says which cell is in the way; the next successful LSC clears it. Refuses on the block itself. Then write the bonus formula over the SIDE results and save it with Save Answer To Bonus N (B1 to B5). Added 2026\-09\-24 (Jaq); spec commandtools\\SPEC\-side\-bonus\-column\-260924.md.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modDataTable.LinkSideCell](./VBA/modDataTable.bas#L902)()</code> |

[^Top](#oa-robot-definitions)

<BR>

### List Combinations Of Array

*Every ordering of all the items in the active array, each used once (10 20 30, 10 30 20, 20 10 30 ...)*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#Array`</sup>

> \*\*Note:\*\* THE ONE THAT SOLVES SUBSET\-SUM. Set HowMany to the number of items, IgnoreOrder 1, AllowRepeats 0 and AllSizesUpTo 1, and you get EVERY SUBSET of the array; then filter for the ones summing to your target. The 2023 Cribbage case's note calls that job 'not native Excel in any clean form' and proposes building a new lambda for it; CompBot already owned this one, unexposed. Also does plain combinations and permutations, with or without repeats, which covers 'how many ways', 'list every pairing', 'every ordering of these tokens'. WARNING: the output grows fast; every subset of 15 items is 32,767 rows, and of 20 items over a million. Start small and check the shape before trusting it on a big list. Wraps Combinations\_byHadynWiseman, which calls PerCom\_byBoRydobon. Exposed as a command 2026\-09\-22.

| Property | Value |
| --- | --- |
| Formula | <code>\=LET(\_a,\[\[ActiveCell::Formula\]\],Combinations\_byHadynWiseman(\_a,COUNTA(\_a)))</code> |
| Formula Dependencies | <ol><li>[Combinations_byHadynWiseman.lambda](#combinations_byhadynwisemanlambda)</li><li>[PerCom_byBoRydobon.lambda](#percom_byborydobonlambda)</li></ol> |
| User Context Filter | ExcelSelectionIsSingleCell |
| Launch Codes | <code>LCA</code> |

[^Top](#oa-robot-definitions)

<BR>

### List Lambdas Used

*Lists every lambda the active workbook's formulas actually call; worth running before clearing anything*

<sup>`@CompBot.xlsm` `!VBA Macro Command` </sup>

> \*\*Note:\*\* Creates a sheet showing all the lambdas actively used in the active workbook

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdas.ListLambdasUsed](./VBA/modLambdas.bas#L943)()</code> |

[^Top](#oa-robot-definitions)

<BR>

### List Pairings Of Two Lists

*Every pairing of the COPIED list with the active spill (each item of one with each item of the other), written to the right of the sheet's contents*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#Array`</sup>

> \*\*Note:\*\* 2026\-09\-28 (Jaq): the other common 'list them all' ask next to List Combinations Of Array. Copy list 1, stand on the spill of list 2, run: two columns, list 1 x list 2, in the next blank cell right of it.

| Property | Value |
| --- | --- |
| Formula | <code>\=LET(\_a,TOCOL(\[\[Clipboard::Address\]\],1),\_b,TOCOL(\[\[ActiveCell.SpillParent::Formula\]\],1),\_n,ROWS(\_a),\_m,ROWS(\_b),\_i,SEQUENCE(\_n\*\_m)\-1,HSTACK(INDEX(\_a,INT(\_i\/\_m)+1),INDEX(\_b,MOD(\_i,\_m)+1)))</code> |
| Destination Range Address | <code>\[\[AvailableCellToRight\]\]</code> |
| User Context Filter | ExcelSelectionIsSingleCell |
| Launch Codes | <code>LP</code> |

[^Top](#oa-robot-definitions)

<BR>

### Load Game Into Calculation

*Loads the selected data table game into the example calculation, or restores the example*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Prep`</sup>

> \*\*Note:\*\* For edge\-case testing and troubleshooting. Select a game row in a Create Data Table block (game number or result) and run it: that game goes into the input cell and the calc block shows it. Select the input cell and run again to restore Example\#. Run Data Table also restores it.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modDataTable.LoadGameIntoCalculation](./VBA/modDataTable.bas#L369)()</code> |
| Launch Codes | <ol><li><code>LDT</code></li><li><code>LG</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Lookup Array from Column N

*Looks every value of the active array up in a named range or table and returns the column you choose*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith` `#Lookup`</sup>

> \*\*Note:\*\* Promoted into CompBot from A\-ZTraining 2026\-09\-22: needed in 7 of 67 surveyed cases. Emits the range NAME rather than a resolved address, so the formula stays readable and survives the table being moved. Exact match, wraps in place. WHERE IT EARNS ITS KEEP: joining a split list back to a stats table, such as 20 racer numbers to their speeds, glyphs to a legend, codes to names. Array Robot's free Paste Lookup does a similar job but is fixed to the LAST column of the copied range; this one lets you pick the column, which is the difference that matters on a wide table. Self\-contained formula, no lambda dependency. Same launch code as A\-ZTraining's copy, deliberately.

| Property | Value |
| --- | --- |
| Formula | <code>\=VLOOKUP(\[\[ActiveCell::Formula\]\],{{LookupRange}},{{ColumnNumber}},FALSE)</code> |
| Parameters | <ol><li>[ColumnNumber](#lookup-array-from-column-n--columnnumber)</li><li>[LookupRange](#lookup-array-from-column-n--lookuprange)</li></ol> |
| User Context Filter | ExcelSelectionIsSingleCell |
| Launch Codes | <code>LKN</code> |

<BR>

#### Lookup Array from Column N \>\> ColumnNumber

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Column number to return</code> |
| Data Type | Integer |

<BR>

#### Lookup Array from Column N \>\> LookupRange

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Named range or table to look up in</code> |
| Data Type | String |

[^Top](#oa-robot-definitions)

<BR>

### Lookup Value by Row

*Returns RC location of a list of values in a range*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#Lookup` `#WrapWith`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=LkpRCByRow(\[\[ActiveCell::Formula\]\], \[\[Clipboard::Address\]\])</code> |
| Formula Dependencies | <ol><li>[LkpRCByRow.lambda](#lkprcbyrowlambda)</li><li>[BiRow_byHadynWiseman.lambda](#birow_byhadynwisemanlambda)</li><li>[LkpRC.lambda](#lkprclambda)</li><li>[FilterArray_byErikOehm.lambda](#filterarray_byerikoehmlambda)</li><li>[IsInList_byErikOehm.lambda](#isinlist_byerikoehmlambda)</li><li>[GRIDTOCOL_byLiannaGerrish.lambda](#gridtocol_byliannagerrishlambda)</li></ol> |
| Launch Codes | <code>lkp</code> |

[^Top](#oa-robot-definitions)

<BR>

### Map Distance From Cell (4 ways)

*How far every cell on a map is from the COPIED start cell(s), moving only up, down, left and right, routing AROUND obstacles*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#Array`</sup>

> \*\*Note:\*\* 2026\-09\-28 (Jaq): replaces Map Distance From Cell's two popups. Copy the cell holding the start address (e.g. E26; a range of starts works too), stand on the formula giving the map, run. MDC8 allows diagonals.

| Property | Value |
| --- | --- |
| Formula | <code>\=MazeDistance\_byHadynWiseman(\[\[ActiveCell::Formula\]\],\[\[Clipboard::Address\]\],,1)</code> |
| Formula Dependencies | <ol><li>[MazeDistance_byHadynWiseman.lambda](#mazedistance_byhadynwisemanlambda)</li><li>[RowNum_byHadynWiseman.lambda](#rownum_byhadynwisemanlambda)</li><li>[ColNum_byHadynWiseman.lambda](#colnum_byhadynwisemanlambda)</li></ol> |
| User Context Filter | ExcelSelectionIsSingleCell |
| Launch Codes | <code>MDC4</code> |

[^Top](#oa-robot-definitions)

<BR>

### Map Distance From Cell (8 ways)

*How far every cell on a map is from the COPIED start cell(s), diagonal moves allowed, routing AROUND obstacles*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#Array`</sup>

> \*\*Note:\*\* 2026\-09\-28 (Jaq): the diagonal twin of MDC4. Copy the start cell, stand on the map formula, run.

| Property | Value |
| --- | --- |
| Formula | <code>\=MazeDistance\_byHadynWiseman(\[\[ActiveCell::Formula\]\],\[\[Clipboard::Address\]\],,0)</code> |
| Formula Dependencies | <ol><li>[MazeDistance_byHadynWiseman.lambda](#mazedistance_byhadynwisemanlambda)</li><li>[RowNum_byHadynWiseman.lambda](#rownum_byhadynwisemanlambda)</li><li>[ColNum_byHadynWiseman.lambda](#colnum_byhadynwisemanlambda)</li></ol> |
| User Context Filter | ExcelSelectionIsSingleCell |
| Launch Codes | <code>MDC8</code> |

[^Top](#oa-robot-definitions)

<BR>

### Merged To Centre Across Selection

*Merged Cells changed to Centre Across Selection*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Prep` `#Formatting`</sup>

> \*\*Note:\*\* Unmerges merged cells across a single row and applies centre across selection formatting instead, so the layout looks the same but the cells select, fill and spill normally. Multi\-row merges are left for Unmerge Multi\-Row Merges. SCOPE: a selection of more than one cell (or one merged cell) limits it to the merged blocks the selection touches; one ordinary cell means the whole active sheet. Status bar report, no dialogs. Promoted from Jaq's own WiE Robot collection 2026\-09\-24 with its name, description and launch codes; WiE Robot itself is unchanged.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modMisc.MergedToCAS](./VBA/modMisc.bas#L37)()</code> |
| Launch Codes | <ol><li><code>Unmerge</code></li><li><code>M2C</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### MM

*Applies \-\- to an array, turning TRUE\/FALSE into 1\/0 so it can be summed or multiplied*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#MM`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=MM\_byHaDang(\[\[ActiveCell::Formula\]\])</code> |
| Formula Dependencies | [MM_byHaDang.lambda](#mm_byhadanglambda) |
| Launch Codes | <code>MM</code> |

[^Top](#oa-robot-definitions)

<BR>

### Modulo Array

*Wraps the active array with MOD; turns cumulative movement into a position on a looping board*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

> \*\*Note:\*\* Promoted into CompBot from A\-ZTraining 2026\-09\-22: needed in 5 of 67 surveyed cases. Use this where the spaces are numbered 0..N\-1 and a full lap returns to 0. Where they are numbered 1..N and a full lap should land on N, use Modulo Array 1\-Based (MOD1) instead; getting that wrong is an off\-by\-one that looks plausible all the way to the answer. Self\-contained formula, no lambda dependency. Same launch code as A\-ZTraining's copy, deliberately.

| Property | Value |
| --- | --- |
| Formula | <code>\=MOD(\[\[ActiveCell::Formula\]\],{{Divisor}})</code> |
| Parameters | <ol><li>[Divisor](#modulo-array--divisor)</li></ol> |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <code>MODA</code> |

<BR>

#### Modulo Array \>\> Divisor

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Divide by (e.g. board size)</code> |
| Data Type | Integer |

[^Top](#oa-robot-definitions)

<BR>

### Modulo Array 1\-Based

*Wraps the active array in BoardGameMove\_byHadynWiseman: a board numbered 1 to N, so an exact lap lands on N rather than 0*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

> \*\*Note:\*\* Promoted into CompBot from A\-ZTraining 2026\-09\-22. The 1\-based partner to Modulo Array (MODA): use this where board spaces are numbered 1..N and a full lap lands on N, and MODA where they are numbered 0..N\-1 and a lap returns to 0. 2026\-10\-05 (Jaq): now writes BoardGameMove\_byHadynWiseman, which is exactly MOD(x\-1,d)+1, instead of inlining the arithmetic; MOD1\_byJaqKennedy (the same thing) was dropped from the lambda set.

| Property | Value |
| --- | --- |
| Formula | <code>\=BoardGameMove\_byHadynWiseman(\[\[ActiveCell::Formula\]\],{{Divisor}})</code> |
| Formula Dependencies | [BoardGameMove_byHadynWiseman.lambda](#boardgamemove_byhadynwisemanlambda) |
| Parameters | <ol><li>[Divisor](#modulo-array-1-based--divisor)</li></ol> |
| User Context Filter | ExcelActiveCellIsSpillParent AND ExcelSelectionIsSingleCell |
| Launch Codes | <code>MOD1</code> |

<BR>

#### Modulo Array 1\-Based \>\> Divisor

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Divide by (e.g. board size)</code> |
| Data Type | Integer |

[^Top](#oa-robot-definitions)

<BR>

### Name All Used Ranges

*Renames used range in all sheets*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNames.NameAllUsedRanges](./VBA/modNames.bas#L30)()</code> |
| Launch Codes | <ol><li><code>NA</code></li><li><code>NAUR</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Name Used Ranges

*Names used range in all sheets starting with the prefix as provided in the active cell*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNames.NameUsedRanges](./VBA/modNames.bas#L11)([[ActiveCell]])</code> |
| Launch Codes | <code>NURS</code> |

[^Top](#oa-robot-definitions)

<BR>

### Negative Values

*Negative values. with errors as blank*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=IFERROR(\-(\[\[ActiveCell::Formula\]\]),"")</code> |
| Launch Codes | <code>neg</code> |

[^Top](#oa-robot-definitions)

<BR>

### Next Bonus In Status Bar

*Shows the next open bonus question on the status bar (no\-dock fallback)*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Bonus`</sup>

> \*\*Note:\*\* Each run steps to the next open bonus and wraps round: '\[2\/3 open\] Bonus 2 \- 20 pts \- Relates to L3: question'. Same detection as Show Bonus Dock; answered and cleared bonuses are skipped. Text is cut at Excel's 255\-character status bar limit. Clear Status Bar (CSB) gives the bar back to Excel.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modBonusDock.StepBonusStatus](./VBA/modBonusDock.bas#L287)(1)</code> |
| Launch Codes | <code>BQN</code> |

[^Top](#oa-robot-definitions)

<BR>

### Paste Flattened List With Formatting

*Turns a copied range into a list, one row per cell, carrying fill color, font color and border flags*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Paste`</sup>

> \*\*Note:\*\* COPY the range first, then select the cell you want the list to start at, then run this. Columns: Items \| Row \# \| Column \# \| Address \| RCRef \| Fill \| Font \| BdrT \| BdrR \| BdrB \| BdrL. WHAT IT IS FOR: reading a map that is drawn in formatting rather than written in values, such as color\-coded terrain, walls drawn as borders, a legend you need to decode. Flatten it, then XLOOKUP the Fill column against the legend, or UNIQUE the Fill column to discover what colors a map actually uses when it has no legend at all. NOTES: colors come back as \#RRGGBB hex, not Excel's raw BGR Long, so they are readable and XLOOKUP\-able. BLANK CELLS ARE INCLUDED; a blank cell with a fill is the whole point. A border between two cells is one line with two possible owners, so each edge flag ORs both sides; a wall drawn from the neighbour's side still shows. RCRef is packed 10^6\*row+column, matching GridToCol. Like a paste, it overwrites the cell you start from; if anything else is in the way (a value or a merged cell in the output block), the list goes to a new sheet next to this one and the start cell becomes a link to it, 'Unable to fit here: see \<sheet\>', and you stay where you were (2026\-10\-05, Jaq). Values are read with .Value2, so dates arrive as serial numbers. A CUT range (Ctrl+X) resolves too, not just a copy. Built and tested 2026\-09\-22; 3000 cells in 0.36s.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFlatten.PasteFlattenedListWithFormatting](./VBA/modFlatten.bas#L58)()</code> |
| Launch Codes | <code>PFF</code> |

[^Top](#oa-robot-definitions)

<BR>

### Previous Bonus In Status Bar

*Shows the previous open bonus question on the status bar*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Bonus`</sup>

> \*\*Note:\*\* Mirror of Next Bonus In Status Bar (BQN).

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modBonusDock.StepBonusStatus](./VBA/modBonusDock.bas#L287)(-1)</code> |
| Launch Codes | <code>BQP</code> |

[^Top](#oa-robot-definitions)

<BR>

### Record Walk Route

*Start recording a route: every orthogonal move adds the cells you pass over. Diagonal click finishes*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Bonus`</sup>

> \*\*Note:\*\* Select the starting cell, run this, then walk. ARROW, CTRL+ARROW and an ORTHOGONAL CLICK all count as a move; clicking is often easier when there is no content to Ctrl+Arrow over. Each move adds every cell between the last stop and the new one. A DIAGONAL CLICK ends it (that cell is not part of the route) and writes the route to a new sheet. Stop Walk Route (SWR) ends it by hand. WHAT IT IS FOR: capturing a path that exists only visually (a route drawn on a map in fills or glyphs and recorded nowhere) and turning it into data you can measure. Output columns: Step \| Items \| Row \# \| Column \# \| Address \| RCRef \| Fill \| Font \| BdrT \| BdrR \| BdrB \| BdrL \| Repeat, the same shape as Paste Flattened List With Formatting, plus the step order. NOTES: a corner cell is recorded ONCE, not twice, where two runs meet. A route that crosses itself keeps EVERY visit in order and flags the revisits in the Repeat column rather than dropping them, because order matters more than uniqueness for a route. A click\-and\-drag takes the last cell of the selection. This command deliberately does NOT use vbaInit\/vbaFin: it needs Application events switched ON to see you move, and it puts the previous setting back when it finishes. Built and tested 2026\-09\-22.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modWalk.StartWalkRoute](./VBA/modWalk.bas#L53)()</code> |
| Launch Codes | <code>RWR</code> |

[^Top](#oa-robot-definitions)

<BR>

### Record Walk Route Here

*As Record Walk Route, but the route table goes on this sheet, level with where the walk started and to its right, in the first space clear enough to hold it*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Bonus`</sup>

> \*\*Note:\*\* 2026\-09\-28 (Jaq): wanted the route beside the working area as well as the new\-sheet version. 2026\-10\-05 (Jaq): the table goes level with the walk's START cell, at the first position from two columns to its right where a block the size of the table holds no values, no fill and no merges (modWalk.ClearSpotRightOf); only if there is none does it go one column after the sheet's contents. The view scrolls to it. Same day: FinishWalk stops the watcher before writing, which ended a self\-retriggering loop that crashed Excel. Stop Walk Route (SWR) ends either mode.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modWalk.StartWalkRouteHere](./VBA/modWalk.bas#L84)()</code> |
| Launch Codes | <code>RWH</code> |

[^Top](#oa-robot-definitions)

<BR>

### Rename Sheets

*Shortens multi\-word sheet names to their initials, keeping numbers (Level 1 Data becomes L1D); backups keep their names*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Prep`</sup>

> \*\*Note:\*\* One character per word (split on spaces and underscores), plus any digits later in that word: Level 1 Data and Level1 Data both become L1D, Level 10 Data L10D. Leaves Case, Case\-Varsity, Answers and backup sheets (BU\_...) alone, so the original names stay readable on the backups. A name already taken gets \_2, \_3... Excel updates every formula that refers to a renamed sheet. A short result that looks like a cell address (L1) still needs quotes in formulas. Status bar lists the renames. Rewritten 2026\-09\-24 (Jaq's C2 notes): it used to drop digits inside a word (Level1 Data and Level2 Data both wanted LD) and renamed the backups.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCaseSetup.RenameSht](./VBA/modCaseSetup.bas#L292)()</code> |
| Launch Codes | <code>rs</code> |

[^Top](#oa-robot-definitions)

<BR>

### Restore Cleared Bonuses

*Brings every hand\-cleared bonus back into the dock*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Bonus`</sup>

> \*\*Note:\*\* Removes the case workbook's hidden BonusDock\_Cleared name, then refreshes the dock.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modBonusDock.RestoreClearedBonuses](./VBA/modBonusDock.bas#L245)()</code> |
| Command After | [Show Bonus Dock](#show-bonus-dock) |
| Launch Codes | <code>BQR</code> |

[^Top](#oa-robot-definitions)

<BR>

### Revert Settings

*Reverts settings to those loaded into the Loaded column.*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Settings`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modMisc.RevertSettings](./VBA/modMisc.bas#L310)()</code> |
| Launch Codes | <code>RVS</code> |

[^Top](#oa-robot-definitions)

<BR>

### Rotate Array

*Turns the active array a quarter turn clockwise; run it again for another quarter turn*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#Array`</sup>

> \*\*Note:\*\* Rotates a grid. A rows\-by\-columns block becomes columns\-by\-rows on a 90 or 270, and keeps its shape on a 180. WHAT IT IS FOR: cases where the board is shown one way up and read another, such as isometric grids, boards viewed from the opposite player's side, tiles that must be turned to fit. ALSO WORTH KNOWING, and not reachable from this command: the lambda takes a third argument, InputAddresses, which returns WHERE A GIVEN CELL MOVED TO after the rotation rather than the rotated grid itself: Rotate\_byHadynWiseman(array, turns, "H12"). That is the piece the 2025 Bigger Portals case needs for its reflection rule, and its own note proposed building a new lambda for it, not knowing this existed. Type it by hand when you need it. Wraps Rotate\_byHadynWiseman; exposed as a command 2026\-09\-22.

| Property | Value |
| --- | --- |
| Formula | <code>\=Rotate\_byHadynWiseman(\[\[ActiveCell::Formula\]\],1)</code> |
| Formula Dependencies | <ol><li>[Rotate_byHadynWiseman.lambda](#rotate_byhadynwisemanlambda)</li><li>[BiCol_byHadynWiseman.lambda](#bicol_byhadynwisemanlambda)</li><li>[BiRow_byHadynWiseman.lambda](#birow_byhadynwisemanlambda)</li><li>[ADDRESSES_byDiarmuidEarly.lambda](#addresses_bydiarmuidearlylambda)</li></ol> |
| User Context Filter | ExcelSelectionIsSingleCell |
| Launch Codes | <code>ROT</code> |

[^Top](#oa-robot-definitions)

<BR>

### Round to 0

*Wraps the active formula in ROUND(...,0) for a whole number. Round once at the END of a chain; rounding each step compounds the error*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#Convert`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=ROUND(\[\[ActiveCell::Formula\]\],0)</code> |
| Launch Codes | <code>r0</code> |

[^Top](#oa-robot-definitions)

<BR>

### Run Data Table

*Fills a Create Data Table block with static results game by game, recalculating only when run*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Prep`</sup>

> \*\*Note:\*\* Replacement for Excel What\-If data tables. Point the ANSWER cell at the example's output, then run from any cell on the sheet. Last\-run time and any volatile\-function warning go to the StatusBar.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modDataTable.RunDataTable](./VBA/modDataTable.bas#L55)()</code> |
| Launch Codes | <code>RDT</code> |

[^Top](#oa-robot-definitions)

<BR>

### Running Total

*Wraps the active array in a running total (RunningTotal\_byJaqKennedy with nothing else set)*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

> \*\*Note:\*\* 2026\-09\-28 (Jaq): one of three commands replacing Running Total With Reset Or Stop. The lambda's other arguments (\[ResetFlags\], \[ResetTo\], \[StopAt\], \[Cap\], \[IncFlaggedVal\]) can be typed into the formula afterwards.

| Property | Value |
| --- | --- |
| Formula | <code>\=RunningTotal\_byJaqKennedy(\[\[ActiveCell::Formula\]\])</code> |
| Formula Dependencies | [RunningTotal_byJaqKennedy.lambda](#runningtotal_byjaqkennedylambda) |
| User Context Filter | ExcelSelectionIsSingleCell |
| Launch Codes | <code>RT</code> |

[^Top](#oa-robot-definitions)

<BR>

### Running Total with Cap

*Running total that stops at the COPIED limit and holds there (Cap \= 1). Change the final 1 to 0 to keep the crossing row's full value instead*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

> \*\*Note:\*\* 2026\-09\-28 (Jaq). Copy the limit first. No popup: StopAt comes from the clipboard, Cap is set to 1.

| Property | Value |
| --- | --- |
| Formula | <code>\=RunningTotal\_byJaqKennedy(\[\[ActiveCell::Formula\]\],,,\[\[Clipboard\]\],1)</code> |
| Formula Dependencies | [RunningTotal_byJaqKennedy.lambda](#runningtotal_byjaqkennedylambda) |
| User Context Filter | ExcelSelectionIsSingleCell |
| Launch Codes | <code>RTC</code> |

[^Top](#oa-robot-definitions)

<BR>

### Running Total with Reset

*Running total that restarts wherever the COPIED flag range holds 1; asks what it restarts at (blank \= 0). The flagged row's own value is included*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

> \*\*Note:\*\* 2026\-09\-28 (Jaq). Copy the flag column first (1 \= restart here). To leave the flagged row's own value out, add ,,,0 after ResetTo in the formula (IncFlaggedVal \= 0).

| Property | Value |
| --- | --- |
| Formula | <code>\=RunningTotal\_byJaqKennedy(\[\[ActiveCell::Formula\]\],\[\[Clipboard::Address\]\],{{ResetTo}})</code> |
| Formula Dependencies | [RunningTotal_byJaqKennedy.lambda](#runningtotal_byjaqkennedylambda) |
| Parameters | <ol><li>[ResetTo](#running-total-with-reset--resetto)</li></ol> |
| User Context Filter | ExcelSelectionIsSingleCell |
| Launch Codes | <code>RTR</code> |

<BR>

#### Running Total with Reset \>\> ResetTo

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>What the total restarts at (leave blank for 0)</code> |
| Data Type | String |

[^Top](#oa-robot-definitions)

<BR>

### Save Answer To Bonus 1

*Links the active cell (any sheet) into the Bonus 1 answer cell, goes there and copies it for the submission site.*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Bonus`</sup>

> \*\*Note:\*\* Target: sheet B if it has the bonus (Case links to it), else the sheet with 'Bonus Questions'; the MEWC green cell on the 'Bonus 1' row. Writes a live link, selects the answer cell, copies it, reports on the StatusBar. Writes nothing if the row or green cell is missing. Promoted from A\-ZTraining 2026\-09\-24 (A\-ZTraining keeps its own copy).

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCaseNav.SaveAnswerToBonus](./VBA/modCaseNav.bas#L246)(1)</code> |
| User Context Filter | ExcelSelectionIsSingleCell |

[^Top](#oa-robot-definitions)

<BR>

### Save Answer To Bonus 2

*Links the active cell (any sheet) into the Bonus 2 answer cell, goes there and copies it for the submission site.*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Bonus`</sup>

> \*\*Note:\*\* See Save Answer To Bonus 1.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCaseNav.SaveAnswerToBonus](./VBA/modCaseNav.bas#L246)(2)</code> |
| User Context Filter | ExcelSelectionIsSingleCell |

[^Top](#oa-robot-definitions)

<BR>

### Save Answer To Bonus 3

*Links the active cell (any sheet) into the Bonus 3 answer cell, goes there and copies it for the submission site.*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Bonus`</sup>

> \*\*Note:\*\* See Save Answer To Bonus 1.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCaseNav.SaveAnswerToBonus](./VBA/modCaseNav.bas#L246)(3)</code> |
| User Context Filter | ExcelSelectionIsSingleCell |

[^Top](#oa-robot-definitions)

<BR>

### Save Answer To Bonus 4

*Links the active cell (any sheet) into the Bonus 4 answer cell, goes there and copies it for the submission site.*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Bonus`</sup>

> \*\*Note:\*\* See Save Answer To Bonus 1.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCaseNav.SaveAnswerToBonus](./VBA/modCaseNav.bas#L246)(4)</code> |
| User Context Filter | ExcelSelectionIsSingleCell |

[^Top](#oa-robot-definitions)

<BR>

### Save Answer To Bonus 5

*Links the active cell (any sheet) into the Bonus 5 answer cell, goes there and copies it for the submission site.*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Bonus`</sup>

> \*\*Note:\*\* See Save Answer To Bonus 1.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCaseNav.SaveAnswerToBonus](./VBA/modCaseNav.bas#L246)(5)</code> |
| User Context Filter | ExcelSelectionIsSingleCell |

[^Top](#oa-robot-definitions)

<BR>

### Save Answers To Left

*Saves references to the selected cells in the green answer cells to the left on the same row.*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Paste`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCaseSetup.SaveAnswersToLeft](./VBA/modCaseSetup.bas#L1693)()</code> |
| Launch Codes | <code>SAL</code> |

[^Top](#oa-robot-definitions)

<BR>

### Save Copy of File

*Enable editing and save copy of file with suffix based on active cell (otherwise Working)*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Prep`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCaseSetup.SaveCopy](./VBA/modCaseSetup.bas#L2307)([[ActiveCell]])</code> |
| Launch Codes | <code>SA</code> |

[^Top](#oa-robot-definitions)

<BR>

### Save From Example

*Select your answer formula on the example row; the working right of the inputs is copied to every question and the answers linked*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Prep`</sup>

> \*\*Note:\*\* Select ONE cell: your answer formula on the example row. Every calculation on that row to the right of the case's last input comes with it, on either side of the answer, spills and typed values included, out to the last filled cell (a gap of 3+ empty columns, or a column with typed values on the question rows such as a case table, ends it). It is copied to every question row of the level and the answer cells are linked to your answer column, ready to paste from the clipboard. Re\-run it after correcting the example: links it wrote earlier are replaced. Before copying it checks the cells it will fill, and stops and says so if any of them hold case content. Every outcome is on the status bar; a failure starts SAVE FROM EXAMPLE FAILED. Modelled on MEWC Robot's command of the same name, but finds the first question by its empty answer cell (any number of example rows) and the answer column by its 'Answer' header, falling back to the sheet's own answer colour. Added 2026\-09\-22; working and re\-run rules 2026\-10\-08 (GitHub \#2).

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modCaseSetup.SaveFromExample](./VBA/modCaseSetup.bas#L1926)()</code> |
| User Context Filter | ExcelActiveCellIsNotEmpty AND ExcelSelectionIsSingleCell |
| Launch Codes | <code>SFE</code> |

[^Top](#oa-robot-definitions)

<BR>

### Sequence of Row Count

*Sequence of row count of array variable*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=SEQUENCE(ROWS(\[\[ActiveCell::Formula\]\]),1,1,1)</code> |
| Launch Codes | <code>src</code> |

[^Top](#oa-robot-definitions)

<BR>

### Set Lambda Library

*Points CompBot at YOUR lambda library workbook, once; ILL and Full Setup Case then load it*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#LAMBDA` `#Prep`</sup>

> \*\*Note:\*\* A setup\-time command, run once before competition (and again only if the library moves). Opens a file picker; choose the workbook your own lambdas live in. The location is saved as a per\-user Windows setting (HKCU, VB and VBA Program Settings\\CompBot\\Lambdas), NOT inside CompBot, so it survives every CompBot update. Refuses CompBot itself. Import Lambda Library (ILL) and Full Setup Case read it from then on. Added 2026\-09\-24.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modLambdas.SetLambdaLibrary](./VBA/modLambdas.bas#L94)()</code> |
| Launch Codes | <code>SLL</code> |

[^Top](#oa-robot-definitions)

<BR>

### Set Solve Folder

*Choose the folder Full Setup Case saves your \_Solve copy into, once (e.g. OneDrive)*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Prep`</sup>

> \*\*Note:\*\* A setup\-time command, run once (and again only if the folder moves). Opens a folder picker. The folder is saved as a per\-user Windows setting (HKCU, VB and VBA Program Settings\\CompBot\\Setup), NOT inside CompBot, so it survives every CompBot update. A OneDrive or SharePoint folder is stored as its local synced path. If the folder is missing when you run Full Setup Case, the copy goes next to the case and the status bar says so. Undo with Clear Solve Folder (CSF). Added 2026\-10\-08 (GitHub \#4).

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modSetupSettings.SetSolveFolder](./VBA/modSetupSettings.bas#L86)()</code> |
| Launch Codes | <code>SSF</code> |

[^Top](#oa-robot-definitions)

<BR>

### Setup Case Settings

*Opens CompBot's Setup Settings sheet: choose which steps Full Setup Case runs, and whose lambdas win a name clash*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Prep` `#Settings`</sup>

> \*\*Note:\*\* Shows CompBot's Setup Settings sheet (making a hidden CompBot window visible; View \> Hide puts it away). A tick box for each Full Setup Case step (save the \_Solve copy, backup, name used ranges, level sheets, bonus sheet, Case inputs sheet, import your library, import CompBot's lambdas) and a drop\-down for whose lambdas win a name clash: CompBot (default) or Your library. The winner's import replaces a same\-named lambda and the other's never does; ILL and ILC follow it too. Also shows the library SLL points at. The choices are a per\-user Windows setting (HKCU, VB and VBA Program Settings\\CompBot\\Setup), NOT cells in CompBot, so they survive every CompBot update and work with CompBot read\-only; a new machine needs them set again. A change is saved at once. The tick boxes are in\-cell checkboxes; an Excel without them shows TRUE\/FALSE, which works the same.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modSetupSettings.ShowSetupSettings](./VBA/modSetupSettings.bas#L185)()</code> |
| Launch Codes | <code>SCS</code> |

[^Top](#oa-robot-definitions)

<BR>

### Show Bonus Dock

*Shows the open bonus questions in a dock on the right; answered ones drop off*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Bonus`</sup>

> \*\*Note:\*\* Reads the active case workbook: 'Bonus 1' \/ 'Bonus A' labels in column B, header row above with Answer \/ Points \/ Question. Sheet B (Create Bonus Sheet) wins, else Case, else the active sheet. Answer sheets are never read. A bonus is answered when its answer cell holds a value or a real formula; a bare link to an empty cell (Case → B) counts as open. Answered on either B or Case counts. Hand\-cleared bonuses (Clear Bonus From Dock) go to the footer. Output: OA Robot's task pane, title 'Bonus Questions', Text content, wrapped at 55 characters by the VBA (HTML renders in a tiny fixed box, Markdown rendered blank; tested 2026\-09\-16). Re\-run to refresh; chain it with CommandAfter behind commands that save bonus answers. Moved into CompBot from the BonusDock collection 2026\-09\-22.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modBonusDock.BonusDockText](./VBA/modBonusDock.bas#L76)()</code> |
| Outputs | Bonus dock pane |
| Launch Codes | <code>BQ</code> |

<BR>

#### Show Bonus Dock \>\> Bonus dock pane

<sup>`!Excel Task Pane Output` </sup>

| Property | Value |
| --- | --- |
| Task Pane Title | <code>Bonus Questions</code> |
| Scope | Application |

[^Top](#oa-robot-definitions)

<BR>

### Split Text by Delimiter Above

*Splits the active cell's text DOWN into rows on the delimiter in the cell above (linked, so changing that cell re\-splits), and trims each piece*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TRIM(TEXTSPLIT(\[\[ActiveCell::Formula\]\],,\[\[ActiveCell.Offset(\-1,0)::Address\]\],TRUE))</code> |
| Launch Codes | <code>ST</code> |

[^Top](#oa-robot-definitions)

<BR>

### Split Text by Semicolon

*Splits the active cell's text on semicolons into an array, trimming each piece*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

> \*\*Note:\*\* Promoted into CompBot from A\-ZTraining 2026\-09\-22: needed in 9 of 67 surveyed cases. Semicolons are far and away the commonest delimiter in the case library: lists of racer numbers, dice rolls, move sequences, team rosters. NOTE THE BASE LEVEL DOES NOT COVER THIS: Text Robot splits by CHARACTER and by WORD, but nothing free splits on a delimiter you choose, so this is real added coverage rather than a duplicate. Self\-contained formula, no lambda dependency. Same launch code as A\-ZTraining's copy, deliberately.

| Property | Value |
| --- | --- |
| Formula | <code>\=TRIM(TEXTSPLIT(\[\[ActiveCell::Formula\]\],";",,TRUE))</code> |
| Launch Codes | <code>S;</code> |

[^Top](#oa-robot-definitions)

<BR>

### Stack Sheets From Clipboard

*Stacks the same block from many sheets into one table, driven by a sheet\-name list in the clipboard*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Array`</sup>

> \*\*Note:\*\* COPY the sheet\-name list first, then select the output cell and run this. The clipboard may hold EITHER two cells, the FIRST and LAST sheet name, taking every sheet between them in tab order (so 64 Area sheets need two cells, not 64), OR the exact list of sheets to stack, in the order given. A row or a column both work. It then finds the BOUNDING USED RANGE across those sheets (the smallest block covering every sheet's data) and writes a StackSheets\_byJaqKennedy formula into the active cell. Output columns: Sheet \| Value \| Address \| Row \| Col \| RCRef, the same shape as GridToCol and Paste Flattened List With Formatting, so the three are interchangeable downstream. WHAT IT IS FOR: cases that split their data across one sheet per area, region or round; the 2024 MEWC Portal case has 64 of them, and nothing can be answered until they are one table. Added 2026\-09\-22.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modStackSheets.StackSheetsFromClipboard](./VBA/modStackSheets.bas#L36)()</code> |
| Launch Codes | <code>STK</code> |

[^Top](#oa-robot-definitions)

<BR>

### Stack Sheets From Clipboard, Hide Blanks

*As Stack Sheets From Clipboard, but drops empty cells. WARNING: it drops zeros too*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Array`</sup>

> \*\*Note:\*\* Identical to Stack Sheets From Clipboard (STK) except it passes HideBlanks\=1, which is right when the sheets are a sparse map and the empty cells are noise. READ THIS BEFORE USING IT ON NUMBERS: the lambda's filter is (val\<\>"")\*(val\<\>0), so it hides ZEROS as well as blanks (measured during the 2026\-09 lambda review). On a numeric grid a real zero silently disappears and your row count is wrong with no error shown. Use the plain STK version on anything where zero is a meaningful value. Added 2026\-09\-22.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modStackSheets.StackSheetsFromClipboardHideBlanks](./VBA/modStackSheets.bas#L44)()</code> |
| Launch Codes | <code>STKH</code> |

[^Top](#oa-robot-definitions)

<BR>

### Stop Walk Route

*Ends a walk route by hand and writes it out, for when a diagonal click is awkward to reach*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Bonus`</sup>

> \*\*Note:\*\* Does the same as the diagonal click that normally ends Record Walk Route (RWR): writes the route to a new sheet, unhooks the event watcher and restores the previous Application.EnableEvents setting. Safe to run when no route is being recorded; it does nothing.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modWalk.StopWalkRoute](./VBA/modWalk.bas#L94)()</code> |
| Launch Codes | <code>SWR</code> |

[^Top](#oa-robot-definitions)

<BR>

### Sum False

*Counts the FALSE (0) values in an array: how many rows fail the test. Anything non\-zero, such as 2 from an OR built with +, counts as TRUE*

<sup>`@CompBot.xlsm` `!Excel Formula Command` </sup>

> \*\*Note:\*\* Added 2026\-09\-28 (Jaq) as the partner of Sum True: MM turns the array into numbers, \=0 flags the FALSE \/ zero results, and MM with sum counts them.

| Property | Value |
| --- | --- |
| Formula | <code>\=MM\_byHaDang(MM\_byHaDang(\[\[ActiveCell::Formula\]\])\=0,1)</code> |
| Formula Dependencies | [MM_byHaDang.lambda](#mm_byhadanglambda) |
| Launch Codes | <code>SF</code> |

[^Top](#oa-robot-definitions)

<BR>

### Sum True

*Counts the TRUE values in an array (applies \-\- and sums it): the usual way to answer 'how many rows pass this test'*

<sup>`@CompBot.xlsm` `!Excel Formula Command` </sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=MM\_byHaDang(\[\[ActiveCell::Formula\]\],1)</code> |
| Formula Dependencies | [MM_byHaDang.lambda](#mm_byhadanglambda) |
| Launch Codes | <ol><li><code>ST</code></li><li><code>MMS</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Toggle Calculation Mode

*Toggles calculation mode and places current mode notice in StatusBar*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Settings`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modMisc.ToggleCalculationMode](./VBA/modMisc.bas#L420)()</code> |
| Launch Codes | <code>TC</code> |

[^Top](#oa-robot-definitions)

<BR>

### Toggle Iterative Calculation

*Toggles iterative calculation and sets status in status bar*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Settings`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modMisc.ToggleIterativeCalculation](./VBA/modMisc.bas#L437)()</code> |
| Launch Codes | <code>IC</code> |

[^Top](#oa-robot-definitions)

<BR>

### Unmerge Multi\-Row Merges

*Unmerges every merged area spanning more than one row (the selection, or the whole sheet), leaving the value in the top\-left cell*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Prep`</sup>

> \*\*Note:\*\* Promoted into CompBot from A\-ZTraining 2026\-09\-22: needed in 7 of 67 surveyed cases, and EVERY sighting was a standalone competition workbook (Tic Tac Toe's 104 three\-row merges, Battleship's 408 two\-row merges, Where's Wally's 142), never an A\-Z training case. WHAT IT IS FOR: vertically merged cells break almost everything; array formulas refuse to spill past them, and the value is only really in the top\-left cell. SINGLE\-ROW merges are left for Merged To Centre Across Selection (M2C), which replaces them without destroying the layout. SCOPE (2026\-09\-24, Jaq): a selection of more than one cell, or one merged cell, limits it to the merged blocks the selection touches (a block partly selected is taken whole); one ordinary cell means the whole active sheet, with UsedRange expanded back out to A1. It collects every area first and unmerges in one hit. Status bar report, no dialogs. NO LAUNCH CODE: typing Unmerge also finds M2C, whose launch code is Unmerge.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modMisc.UnmergeMultiRowMerges](./VBA/modMisc.bas#L23)()</code> |

[^Top](#oa-robot-definitions)

<BR>

### Update Settings

*Updates default settings to those in Regional Settings sheet*

<sup>`@CompBot.xlsm` `!VBA Macro Command` `#Settings`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modMisc.UpdateSettings](./VBA/modMisc.bas#L251)()</code> |
| Launch Codes | <code>US</code> |

[^Top](#oa-robot-definitions)

<BR>

### Wrap Flat List Into Grid

*Folds a single row or column back into a grid of a chosen width, the missing half of Reshape To One Row\/Column*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#Array`</sup>

> \*\*Note:\*\* Takes a flat list and lays it out row by row into a grid the width you ask for. WHY IT EXISTS: Array Robot ships Reshape To One (1) Column and Reshape To One (1) Row; both go from a GRID to a LIST. Nothing anywhere went back the other way, in any collection or lambda: the family stopped one member short. Found needed in two surveyed cases by different authors in different years: a 50\-column block that is really 5 reels x 10 turns (2023 A Story About the Reels), and a flat 9\- or 25\-token row that is really a 3x3 or 5x5 board (2026 Sunlight). Ragged input is padded with blanks rather than erroring. Uses native WRAPROWS, so there is no lambda to maintain. Added 2026\-09\-22.

| Property | Value |
| --- | --- |
| Formula | <code>\=WRAPROWS(TOCOL(\[\[ActiveCell::Formula\]\],3),{{RowWidth}},"")</code> |
| Parameters | <ol><li>[RowWidth](#wrap-flat-list-into-grid--rowwidth)</li></ol> |
| User Context Filter | ExcelSelectionIsSingleCell |
| Launch Codes | <code>WFG</code> |

<BR>

#### Wrap Flat List Into Grid \>\> RowWidth

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>How many cells per row in the result</code> |
| Data Type | String |

[^Top](#oa-robot-definitions)

<BR>

### Wrap in ABS

*Wraps the active formula in ABS() to make the result positive: distances, differences, gaps, reflections*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=ABS(\[\[ActiveCell::Formula\]\])</code> |
| Launch Codes | <code>ABS</code> |

[^Top](#oa-robot-definitions)

<BR>

### Wrap in Concat

*Wraps the active formula in CONCAT() to join an array into one string with no separator, rebuilding a word from its letters*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=CONCAT(\[\[ActiveCell::Formula\]\])</code> |
| Launch Codes | <code>con</code> |

[^Top](#oa-robot-definitions)

<BR>

### Wrap in Drop First Row

*Wraps the active formula in DROP(...,1) to remove a header row from a spilled array*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=DROP(\[\[ActiveCell::Formula\]\],1)</code> |
| Launch Codes | <code>DF</code> |

[^Top](#oa-robot-definitions)

<BR>

### Wrap in Take by Copied Cell Columns

*Wrap in take by copied cell columns.*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=TAKE(\[\[ActiveCell::Formula\]\],,\[\[Clipboard\]\])</code> |
| Launch Codes | <code>TCC</code> |

[^Top](#oa-robot-definitions)

<BR>

### Wrap in UNICHAR

*Wraps the active formula in UNICHAR() to turn code numbers back into characters: decoding glyphs, emoji or dice faces*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=UNICHAR(\[\[ActiveCell::Formula\]\])</code> |
| Launch Codes | <code>char</code> |

[^Top](#oa-robot-definitions)

<BR>

### Wrap in UNICODE

*Wraps the active formula in UNICODE() to turn characters into their code numbers, the first step in most cipher and glyph puzzles*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=UNICODE(\[\[ActiveCell::Formula\]\])</code> |
| Launch Codes | <code>code</code> |

[^Top](#oa-robot-definitions)

<BR>

### Wrap in UNIQUE

*Wraps the active formula in UNIQUE() to strip duplicates: distinct values, distinct colors, distinct answers*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=UNIQUE(\[\[ActiveCell::Formula\]\])</code> |
| Launch Codes | <code>U</code> |

[^Top](#oa-robot-definitions)

<BR>

### Wrap with IFERROR TEXTBEFORE space

*Wraps current formula with TEXTBEFORE space with IFERROR in case space doesn't exist*

<sup>`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=IFERROR(TEXTBEFORE(\[\[ActiveCell::Formula\]\]," "),\[\[ActiveCell::Formula\]\])</code> |
| Launch Codes | <code>tbse</code> |

[^Top](#oa-robot-definitions)

<BR>

## Text Definitions

<BR>

### ADDRESSES\_byDiarmuidEarly.lambda

*Definition of ADDRESSES\_byDiarmuidEarly lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [ADDRESSES_byDiarmuidEarly.lambda](<./Text/ADDRESSES_byDiarmuidEarly.lambda.txt>) |
| Value | <code>\/\*Returns the cell address of a range\*\/</code><br><code>Addresses\_byDiarmuidEarly \= LAMBDA(array,ADDRESS(ROW(array),COLUMN(array),4));</code> |
| Content Type | ExcelFormula |
| Location | <code>ADDRESSES\_byDiarmuidEarly</code> |

[^Top](#oa-robot-definitions)

<BR>

### ARROWSHIFT.lambda

*Definition of ARROWSHIFT lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [ARROWSHIFT.lambda](<./Text/ARROWSHIFT.lambda.txt>) |
| Value | ExcelNameText [ARROWSHIFT.lambda] was unable to find name [ARROWSHIFT] in workbook [CompBot]. |
| Content Type | ExcelLambda |
| Location | <code>ARROWSHIFT</code> |

[^Top](#oa-robot-definitions)

<BR>

### ARROWSHIFT\_byEmilieWilliams.lambda

*Definition of ARROWSHIFT\_byEmilieWilliams lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [ARROWSHIFT_byEmilieWilliams.lambda](<./Text/ARROWSHIFT_byEmilieWilliams.lambda.txt>) |
| Value | <code>ARROWSHIFT\_byEmilieWilliams \= LAMBDA(starting\_cell,arrow, LET(</code><br><code> Arrows, VSTACK("↗", "↓", "↖", "←", "→", "↙", "↘", "↑"),</code><br><code> Cards, VSTACK("NE", "S", "NW", "W", "E", "SW", "SE", "N"),</code><br><code> Arrowsx, VSTACK(1, 0, \-1, \-1, 1, \-1, 1, 0),</code><br><code> Arrowsy, VSTACK(\-1, 1, \-1, 0, 0, 1, 1, \-1),</code><br><code> CardInd, ISERROR(XMATCH(RIGHT(arrow), Arrows, 0)),</code><br><code> CardComment, "Multiplier Can Only Be Used On Arrows",<... |
| Content Type | ExcelLambda |
| Location | <code>ARROWSHIFT\_byEmilieWilliams</code> |

[^Top](#oa-robot-definitions)

<BR>

### BiCol\_byHadynWiseman.lambda

*Definition of BiCol\_byHadynWiseman lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [BiCol_byHadynWiseman.lambda](<./Text/BiCol_byHadynWiseman.lambda.txt>) |
| Value | <code>BiCol\_byHadynWiseman \= LAMBDA(array,function,IF(COLUMNS(array)\=1,IF(ROWS(array)\=1,function(@array),function(array)),HSTACK(BiCol\_byHadynWiseman(TAKE(array,,COLUMNS(array)\/2),function),BiCol\_byHadynWiseman(DROP(array,,COLUMNS(array)\/2),function))));</code> |
| Content Type | ExcelLambda |
| Location | <code>BiCol\_byHadynWiseman</code> |

[^Top](#oa-robot-definitions)

<BR>

### BiRow\_byHadynWiseman.lambda

*Definition of BiRow\_byHadynWiseman lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [BiRow_byHadynWiseman.lambda](<./Text/BiRow_byHadynWiseman.lambda.txt>) |
| Value | <code>BiRow\_byHadynWiseman \= LAMBDA(array,function,IF(ROWS(array)\=1,IF(COLUMNS(array)\=1,function(@array),function(array)),VSTACK(BiRow\_byHadynWiseman(TAKE(array,ROWS(array)\/2),function),BiRow\_byHadynWiseman(DROP(array,ROWS(array)\/2),function))));</code> |
| Content Type | ExcelLambda |
| Location | <code>BiRow\_byHadynWiseman</code> |

[^Top](#oa-robot-definitions)

<BR>

### BoardGameMove\_byHadynWiseman.lambda

*Definition of BoardGameMove\_byHadynWiseman lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [BoardGameMove_byHadynWiseman.lambda](<./Text/BoardGameMove_byHadynWiseman.lambda.txt>) |
| Value | <code>BoardGameMove\_byHadynWiseman \= LAMBDA(MoveValue,EndSpace,MOD(MoveValue \- 1, EndSpace) + 1);</code> |
| Content Type | ExcelLambda |
| Location | <code>BoardGameMove\_byHadynWiseman</code> |

[^Top](#oa-robot-definitions)

<BR>

### ClosestOnMap\_byHadynWiseman.lambda

*Definition of ClosestOnMap\_byHadynWiseman lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [ClosestOnMap_byHadynWiseman.lambda](<./Text/ClosestOnMap_byHadynWiseman.lambda.txt>) |
| Value | <code>ClosestOnMap\_byHadynWiseman \= LAMBDA(mp,start,steps,\[Items\],\[DistFunc\],\[ShowBlanks\],\[RCaa\_CRaa\_RCad\_CRad\_RCda\_CRda\_RCdd\_CRdd\],LET(st,start,rng,steps,brng,rng\*10,o,RCaa\_CRaa\_RCad\_CRad\_RCda\_CRda\_RCdd\_CRdd,mr,ROWS(mp),mc,COLUMNS(mp), vs,INDEX(mp,MAX(RowNum\_byHadynWiseman(st)\-MIN(ROW(mp))+1\-brng,1),MAX(ColNum\_byHadynWiseman(st)\-MIN(COLUMN(mp))+1\-brng,1)):INDEX(mp,MIN(RowNum\_byHadynWiseman(st)\-MIN(ROW(mp))+1+brng,mr),MIN(ColNum\_byHadynWiseman(st)\-MIN(COLUMN... |
| Content Type | ExcelLambda |
| Location | <code>ClosestOnMap\_byHadynWiseman</code> |

[^Top](#oa-robot-definitions)

<BR>

### ColNum\_byHadynWiseman.lambda

*Definition of ColNum\_byHadynWiseman lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [ColNum_byHadynWiseman.lambda](<./Text/ColNum_byHadynWiseman.lambda.txt>) |
| Value | <code>ColNum\_byHadynWiseman \= LAMBDA(address,IF(ISNUMBER(address),MOD(address,10^6), LET(chars,UPPER(REGEXEXTRACT(address,"\[A\-Za\-z\]+")), L,LEN(chars), val\_1,CODE(RIGHT(chars,1))\-64, val\_2,IF(L\>1,(CODE(LEFT(RIGHT(chars,2),1))\-64)\*26,0), val\_3,IF(L\=3,(CODE(LEFT(chars,1))\-64)\*676, 0), val\_1+val\_2+val\_3)));</code> |
| Content Type | ExcelLambda |
| Location | <code>ColNum\_byHadynWiseman</code> |

[^Top](#oa-robot-definitions)

<BR>

### Combinations\_byHadynWiseman.lambda

*Definition of Combinations\_byHadynWiseman lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [Combinations_byHadynWiseman.lambda](<./Text/Combinations_byHadynWiseman.lambda.txt>) |
| Value | <code>Combinations\_byHadynWiseman \= LAMBDA(Array,c,\[RemoveOrder\],\[AllowRepeatItems\],\[Cumulative\],IFERROR(XLOOKUP(PerCom\_byBoRydobon(COUNTA(Array), c, IF(RemoveOrder \= 1, 1), IF(AllowRepeatItems \= 1, 1), IF(Cumulative \= 1, 1)), TAKE(HSTACK(SEQUENCE(COUNTA(Array)), TOCOL(Array)), , 1), TAKE(HSTACK(SEQUENCE(COUNTA(Array)), TOCOL(Array)), , \-1)), ""));</code> |
| Content Type | ExcelLambda |
| Location | <code>Combinations\_byHadynWiseman</code> |

[^Top](#oa-robot-definitions)

<BR>

### Dice\_byHadynWiseman.lambda

*Definition of Dice\_byHadynWiseman lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [Dice_byHadynWiseman.lambda](<./Text/Dice_byHadynWiseman.lambda.txt>) |
| Value | <code>Dice\_byHadynWiseman \= LAMBDA(Dice,\[add\],\[join\],LET(fin, MAP(Dice, LAMBDA(a, LET(l, IFERROR(SplitText\_byHadynWiseman(a), ""), m, IF(add, 0, IF(join, "", l)), rs, IFERROR(XLOOKUP(l, {"⚀";"⚁";"⚂";"⚃";"⚄";"⚅"}, {1;2;3;4;5;6}), m), IF(add, SUM(rs), CONCAT(rs))))), IF(COUNTA(fin) \= 1, @fin, fin)));</code> |
| Content Type | ExcelLambda |
| Location | <code>Dice\_byHadynWiseman</code> |

[^Top](#oa-robot-definitions)

<BR>

### DiffByRow\_byJaqKennedy.lambda

*Definition of DiffByRow\_byJaqKennedy lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [DiffByRow_byJaqKennedy.lambda](<./Text/DiffByRow_byJaqKennedy.lambda.txt>) |
| Value | <code>DiffByRow\_byJaqKennedy \= LAMBDA(Input,LET(\\\\LambdaName, "DiffByRow", \\\\CommandName, "Difference of array columns by row", \\\\Description, "Returns first column \- last column in array", \\\\Source, "Jaq Kennedy", TAKE(Input, , 1) \- TAKE(Input, , \-1)));</code> |
| Content Type | ExcelFormula |
| Location | <code>DiffByRow\_byJaqKennedy</code> |

[^Top](#oa-robot-definitions)

<BR>

### Distance\_byHadynWiseman.lambda

*Definition of Distance\_byHadynWiseman lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [Distance_byHadynWiseman.lambda](<./Text/Distance_byHadynWiseman.lambda.txt>) |
| Value | <code>Distance\_byHadynWiseman \= LAMBDA(start,end,\[h\],\[v\],\[d\],\[pythag\],\[knight\], LET(rd,ABS(RowNum\_byHadynWiseman(start)\-RowNum\_byHadynWiseman(end)),cd,ABS(ColNum\_byHadynWiseman(start)\-ColNum\_byHadynWiseman(end)),dd,IF(rd\>cd,rd,cd),md,IF(rd\<cd,rd,cd), hvd,IF(ISOMITTED(h),0,IF(ISOMITTED(d),rd\*v+cd\*h,md\*d+(rd\-md)\*v+(cd\-md)\*h)), py,IF(pythag,SQRT(rd^2+cd^2),0), kn,IF(knight,LET(oo,dd\/2,ot,(dd+md)\/3,km,ROUNDUP(IF(oo\>ot,oo,ot),0),parity,IF(ISODD(km)\=ISODD(dd+md),... |
| Content Type | ExcelLambda |
| Location | <code>Distance\_byHadynWiseman</code> |

[^Top](#oa-robot-definitions)

<BR>

### Exists\_byHadynWiseman.lambda

*Definition of Exists\_byHadynWiseman lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [Exists_byHadynWiseman.lambda](<./Text/Exists_byHadynWiseman.lambda.txt>) |
| Value | <code>Exists\_byHadynWiseman \= LAMBDA(Item,List,\[SearchMode\],LET(\_Mode,IF(ISOMITTED(SearchMode)+ISBLANK(SearchMode),0,SearchMode),ISNUMBER(IF(\_Mode,SEARCH(Item,List),XMATCH(Item,List)))));</code> |
| Content Type | ExcelLambda |
| Location | <code>Exists\_byHadynWiseman</code> |

[^Top](#oa-robot-definitions)

<BR>

### Extract\_byJaqKennedy.lambda

*Definition of Extract\_byJaqKennedy lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [Extract_byJaqKennedy.lambda](<./Text/Extract_byJaqKennedy.lambda.txt>) |
| Value | ExcelNameText [Extract_byJaqKennedy.lambda] was unable to find name [Extract_byJaqKennedy] in workbook [CompBot]. |
| Content Type | ExcelLambda |
| Location | <code>Extract\_byJaqKennedy</code> |

[^Top](#oa-robot-definitions)

<BR>

### FilterArray\_byErikOehm.lambda

*Definition of FilterArray\_byErikOehm lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [FilterArray_byErikOehm.lambda](<./Text/FilterArray_byErikOehm.lambda.txt>) |
| Value | <code>FilterArray\_byErikOehm \= LAMBDA(data,column\_indexes,filter\_values,LET(\\\\LambdaName, "FilterArray", FILTER(data, BYROW(IsInList\_byErikOehm(CHOOSECOLS(data, column\_indexes), filter\_values), LAMBDA(x, AND(x))))));</code> |
| Content Type | ExcelFormula |
| Location | <code>FilterArray\_byErikOehm</code> |

[^Top](#oa-robot-definitions)

<BR>

### FindInMap\_byHadynWiseman.lambda

*Definition of FindInMap\_byHadynWiseman lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [FindInMap_byHadynWiseman.lambda](<./Text/FindInMap_byHadynWiseman.lambda.txt>) |
| Value | <code>FindInMap\_byHadynWiseman \= LAMBDA(items,mp,LET(g,DROP(GridToCol\_byLiannaGerrish(mp,1,,,,,1),1),incells,CHOOSECOLS(g,1),addresses,CHOOSECOLS(g,2),out,FILTER(addresses,ISNUMBER(XMATCH(incells,items)),"Not Found"),IF((ROWS(out)\*COLUMNS(out))\=1,@out,out)));</code> |
| Content Type | ExcelLambda |
| Location | <code>FindInMap\_byHadynWiseman</code> |

[^Top](#oa-robot-definitions)

<BR>

### Flip\_byHadynWiseman.lambda

*Definition of Flip\_byHadynWiseman lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [Flip_byHadynWiseman.lambda](<./Text/Flip_byHadynWiseman.lambda.txt>) |
| Value | <code>Flip\_byHadynWiseman \= LAMBDA(Array,\[Horizontal\],\[Vertical\],\[DiagonalTopLeft\],\[DiagonalTopRight\],\[AddressMoves\],LET(a, Array, c, COLUMNS(a), r, ROWS(a), ta, AddressMoves, rs, LAMBDA(x, IF(r \* c \= 1, CONCAT(MID(x, SEQUENCE(LEN(x), , LEN(x), \-1), 1)), IF(DiagonalTopLeft, TRANSPOSE(x), IF(Horizontal, SORTBY(x, SEQUENCE(, c, c, \-1)), IF(Vertical, SORTBY(x, SEQUENCE(r, , r, \-1)), IF(DiagonalTopRight, TRANSPOSE(INDEX(x, SEQUENCE(r, , r, \-1), SEQUENCE(, c, c, \-1))), INDEX(x, SEQ... |
| Content Type | ExcelLambda |
| Location | <code>Flip\_byHadynWiseman</code> |

[^Top](#oa-robot-definitions)

<BR>

### GetAddress.lambda

*Definition of GetAddress lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [GetAddress.lambda](<./Text/GetAddress.lambda.txt>) |
| Value | ExcelNameText [GetAddress.lambda] was unable to find name [GetAddress] in workbook [CompBot]. |
| Content Type | ExcelLambda |
| Location | <code>GetAddress</code> |

[^Top](#oa-robot-definitions)

<BR>

### GetAddressesByLookup.lambda

*Definition of GetAddressesByLookup lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [GetAddressesByLookup.lambda](<./Text/GetAddressesByLookup.lambda.txt>) |
| Value | ExcelNameText [GetAddressesByLookup.lambda] was unable to find name [GetAddressesByLookup] in workbook [CompBot]. |
| Content Type | ExcelLambda |
| Location | <code>GetAddressesByLookup</code> |

[^Top](#oa-robot-definitions)

<BR>

### GRIDTOCOL\_byLiannaGerrish.lambda

*Definition of GRIDTOCOL\_byLiannaGerrish lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [GRIDTOCOL_byLiannaGerrish.lambda](<./Text/GRIDTOCOL_byLiannaGerrish.lambda.txt>) |
| Value | <code>GridToCol\_byLiannaGerrish \= LAMBDA(arr,\[Address\],\[RowCol\],\[RCRef\],\[ByColumn\],\[HideZeros\],\[ShowBlanks\],LET(nv,NOT(ISREF(arr)),bycol,IF(ISOMITTED(ByColumn),0,ByColumn),r0,IF(nv,1,@ROW(arr)),c0,IF(nv,1,@COLUMN(arr)),rws,TOCOL(SEQUENCE(,COLUMNS(arr))\*0&SEQUENCE(ROWS(arr),,r0),,bycol)\*1,cls,TOCOL(SEQUENCE(ROWS(arr))\*0&SEQUENCE(,COLUMNS(arr),c0),,bycol)\*1,empt,TOCOL(IF(nv,arr\="",ISBLANK(arr)),,bycol),unit,empt\*0+1,items,IF(empt,"",TOCOL(arr,,bycol)),wA,IF(ISOMITTED(Address),0... |
| Content Type | ExcelFormula |
| Location | <code>GRIDTOCOL\_byLiannaGerrish</code> |

[^Top](#oa-robot-definitions)

<BR>

### IFBLANK.lambda

*Definition of IFBLANK lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [IFBLANK.lambda](<./Text/IFBLANK.lambda.txt>) |
| Value | <code>IFBLANK \= LAMBDA(value,value\_if\_blank, IF(ISBLANK(value), value\_if\_blank, value));</code> |
| Content Type | ExcelFormula |
| Location | <code>IFBLANK</code> |

[^Top](#oa-robot-definitions)

<BR>

### IsInList\_byErikOehm.lambda

*Definition of IsInList\_byErikOehm lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [IsInList_byErikOehm.lambda](<./Text/IsInList_byErikOehm.lambda.txt>) |
| Value | <code>IsInList\_byErikOehm \= LAMBDA(array,list,MAP(array,LAMBDA(x,OR(list\=x))));</code> |
| Content Type | ExcelFormula |
| Location | <code>IsInList\_byErikOehm</code> |

[^Top](#oa-robot-definitions)

<BR>

### LkpRC.lambda

*Definition of LkpRC lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [LkpRC.lambda](<./Text/LkpRC.lambda.txt>) |
| Value | <code>LkpRC \= LAMBDA(ToFind,Range,LET(\\\\LambdaName, "LkpRC", \\\\CommandName, "Lookup RC coordinates ", \\\\Description, "Looks up RC coordinates of matching cell(s) in range", CHOOSECOLS(FilterArray\_byErikOehm(GridToCol\_byLiannaGerrish(Range), {1}, ToFind), {2,3})));</code> |
| Content Type | ExcelFormula |
| Location | <code>LkpRC</code> |

[^Top](#oa-robot-definitions)

<BR>

### LkpRCByRow.lambda

*Definition of LkpRCByRow lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [LkpRCByRow.lambda](<./Text/LkpRCByRow.lambda.txt>) |
| Value | <code>LkpRCByRow \= LAMBDA(ToFind,Range,LET(\\\\LambdaName, "LkpRCByRow", \\\\CommandName, "Lookup Value by Row", \\\\Description, "Returns RC location of a list of values in a range", \\\\Source, "Jaq Kennedy", BiRow\_byHadynWiseman(ToFind, LAMBDA(a, IFERROR(CHOOSEROWS(LkpRC(a, Range), 1), {"",""})))));</code> |
| Content Type | ExcelFormula |
| Location | <code>LkpRCByRow</code> |

[^Top](#oa-robot-definitions)

<BR>

### MazeDistance\_byHadynWiseman.lambda

*Definition of MazeDistance\_byHadynWiseman lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [MazeDistance_byHadynWiseman.lambda](<./Text/MazeDistance_byHadynWiseman.lambda.txt>) |
| Value | <code>MazeDistance\_byHadynWiseman \= LAMBDA(mp,st,\[Targets\],\[No\_Diag\],\[Allowmp\],\[Show\_Steps\],\[Cus\_Dir\],\[Mazelam\],\[ReferenceMap\],LET(M,10^6,rm,ReferenceMap,nv,NOT(ISREF(mp)),mis,ISOMITTED(rm),IF(nv\*mis,"Map is not a range, add reference map",LET( um,IF(mis,mp,rm),rn,ROW(um),cn,COLUMN(um),rcn,rn\*M+cn, dir,{1000000,\-1000000,1,\-1,1000001,\-1000001,999999,\-999999},xRC,IF(ISOMITTED(Cus\_Dir),TAKE(dir,,8\-4\*No\_Diag),TOROW(Cus\_Dir,3)),ir,@rn\-1,ic,@cn\-1,na,HSTACK(0,0,0),rcst,T... |
| Content Type | ExcelLambda |
| Location | <code>MazeDistance\_byHadynWiseman</code> |

[^Top](#oa-robot-definitions)

<BR>

### MM\_byHaDang.lambda

*Definition of MM\_byHaDang lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [MM_byHaDang.lambda](<./Text/MM_byHaDang.lambda.txt>) |
| Value | <code>\/\*Applies \-\- to values, with option to sum output. \*\/</code><br><code>MM\_byHaDang \= LAMBDA(value,\[add\_up\], LET(</code><br><code> \\\\LambdaName, "MM\_byHaDang",</code><br><code> \\\\Description, "Applies \-\- to values, with option to sum output",</code><br><code> result, IFERROR(\-\-(value), value & ""),</code><br><code> IF(add\_up, SUM(result), result)</code><br><code>));</code> |
| Content Type | ExcelLambda |
| Location | <code>MM\_byHaDang</code> |

[^Top](#oa-robot-definitions)

<BR>

### PerCom\_byBoRydobon.lambda

*Definition of PerCom\_byBoRydobon lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [PerCom_byBoRydobon.lambda](<./Text/PerCom_byBoRydobon.lambda.txt>) |
| Value | <code>PerCom\_byBoRydobon \= LAMBDA(n,c,\[P0C1\],\[rep\],\[cum\],LET(s,UNICHAR(SEQUENCE(,n,20001)), t,SEQUENCE(,c), re,REDUCE("",t,LAMBDA(x,w,LET(b,IF(cum,FILTER(x,LEN(x)\=w\-1),x), d,TOCOL(IFS(IF(P0C1,IF(rep,RIGHT(b)\<\=s,RIGHT(b)\<s),IF(rep,1,ISERR(FIND(s,b)))),b&s),3),IF(cum,VSTACK(x,d),d)))), IFERROR(UNICODE(MID(re,t,1))\-20000,"")));</code> |
| Content Type | ExcelLambda |
| Location | <code>PerCom\_byBoRydobon</code> |

[^Top](#oa-robot-definitions)

<BR>

### Regions\_byHadynWiseman.lambda

*Definition of Regions\_byHadynWiseman lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [Regions_byHadynWiseman.lambda](<./Text/Regions_byHadynWiseman.lambda.txt>) |
| Value | <code>Regions\_byHadynWiseman \= LAMBDA(mp,\[No\_Diag\],\[Allowmp\],\[Cus\_Dir\],\[Mazelam\],LET(M,10^6,rn,SEQUENCE(ROWS(mp)),cn,SEQUENCE(,COLUMNS(mp)),rcn,rn\*M+cn, dir,{1000000,\-1000000,1,\-1,1000001,\-1000001,999999,\-999999},xRC,IF(ISOMITTED(Cus\_Dir),TAKE(dir,,8\-4\*No\_Diag),HSTACK(TOROW(Cus\_Dir,3),TOROW(\-Cus\_Dir,3))),na,HSTACK(0,0,0),mpA,IF(ISOMITTED(Allowmp),mp\=IF(Exists\_byHadynWiseman("",TOCOL(mp)),"",0),Allowmp),rcst,FILTER(TOCOL(rcn),TOCOL(mpA)), fx, LAMBDA(lfx,pool,ft,n, ... |
| Content Type | ExcelLambda |
| Location | <code>Regions\_byHadynWiseman</code> |

[^Top](#oa-robot-definitions)

<BR>

### RightAlignedArray\_byJaqKennedy.lambda

*Definition of RightAlignedArray\_byJaqKennedy lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [RightAlignedArray_byJaqKennedy.lambda](<./Text/RightAlignedArray_byJaqKennedy.lambda.txt>) |
| Value | <code>\/\*Aligns contents of array to the right, ignoring blanks\*\/</code><br><code>RightAlignedArray\_byJaqKennedy \= LAMBDA(Input,LET(\\\\LambdaName, "RightAlignedArray", \\\\CommandName, "Align array to right", \\\\Description, "Aligns array to the right with blanks to left", \\\\Source, "Jaq Kennedy", \_ColsByRow, BYROW(Input, COUNT), \_Rows, ROWS(Input), \_Cols, COLUMNS(Input), \_ColIndex, IF(MOD(SEQUENCE(\_Rows, \_Cols) \- 1, \_Cols) + 1 \- \_Cols + \_ColsByRow \<\= 0, \-1, MOD(SEQUENCE(\... |
| Content Type | ExcelFormula |
| Location | <code>RightAlignedArray\_byJaqKennedy</code> |

[^Top](#oa-robot-definitions)

<BR>

### Rotate\_byHadynWiseman.lambda

*Definition of Rotate\_byHadynWiseman lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [Rotate_byHadynWiseman.lambda](<./Text/Rotate_byHadynWiseman.lambda.txt>) |
| Value | <code>Rotate\_byHadynWiseman \= LAMBDA(Array,\[Rotation\],\[InputAddresses\],LET(ar,IF(ISOMITTED(Rotation),1,MOD(Rotation\-1,4)+1),rev,LAMBDA(a,INDEX(a,SEQUENCE(ROWS(a),,ROWS(a),\-1),SEQUENCE(,COLUMNS(a),COLUMNS(a),\-1))),rs,CHOOSE(ar,TRANSPOSE(BiCol\_byHadynWiseman(Array,rev)),rev(Array),TRANSPOSE(BiRow\_byHadynWiseman(Array,rev)),Array),IF(ISOMITTED(InputAddresses),rs,LET(sr,@ROW(Array),sc,@COLUMN(Array),oad,Addresses\_byDiarmuidEarly(Array),nad,IF(ISODD(ar),ADDRESS(SEQUENCE(COLUMNS(Array),,sr... |
| Content Type | ExcelLambda |
| Location | <code>Rotate\_byHadynWiseman</code> |

[^Top](#oa-robot-definitions)

<BR>

### RowNum\_byHadynWiseman.lambda

*Definition of RowNum\_byHadynWiseman lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [RowNum_byHadynWiseman.lambda](<./Text/RowNum_byHadynWiseman.lambda.txt>) |
| Value | <code>RowNum\_byHadynWiseman \= LAMBDA(addresses,IF(ISNUMBER(addresses),INT(addresses\/10^6),1\*REGEXEXTRACT(addresses,"\\d+")));</code> |
| Content Type | ExcelLambda |
| Location | <code>RowNum\_byHadynWiseman</code> |

[^Top](#oa-robot-definitions)

<BR>

### RunningTotal\_byJaqKennedy.lambda

*Definition of RunningTotal\_byJaqKennedy lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [RunningTotal_byJaqKennedy.lambda](<./Text/RunningTotal_byJaqKennedy.lambda.txt>) |
| Value | <code>\/\*Running total that can RESTART at flagged rows and CAP or STOP at a limit. Only values is required; with nothing else it is a plain running total. ResetFlags: 1 where the total restarts at ResetTo (default 0); the flagged row's own value is included unles\*\/</code><br><code>RunningTotal\_byJaqKennedy \= LAMBDA(values,\[ResetFlags\],\[ResetTo\],\[StopAt\],\[Cap\],\[IncFlaggedVal\],LET(v,TOCOL(values),n,ROWS(v),f,IF(ISOMITTED(ResetFlags),SEQUENCE(n)\*0,IFERROR(\-\-TOCOL(ResetFlags),0)),... |
| Content Type | ExcelLambda |
| Location | <code>RunningTotal\_byJaqKennedy</code> |

[^Top](#oa-robot-definitions)

<BR>

### SplitText\_byHadynWiseman.lambda

*Definition of SplitText\_byHadynWiseman lambda function.*

<sup>`@CompBot.xlsm` `!Excel Name Text` </sup>

| Property | Value |
| --- | --- |
| Text | [SplitText_byHadynWiseman.lambda](<./Text/SplitText_byHadynWiseman.lambda.txt>) |
| Value | <code>SplitText\_byHadynWiseman \= LAMBDA(Text,\[Groups\],\[Overlap\],LET(t,CONCAT(Text),\_groups,MAX(Groups,1),IF(Overlap,MID(t,SEQUENCE(LEN(t)\-\_groups+1),\_groups),IF(ISOMITTED(Groups),TOCOL(REGEXEXTRACT(t,"\\X",1)),MID(t,SEQUENCE(ROUNDUP(LEN(t)\/\_groups,0),,,\_groups),\_groups)))));</code> |
| Content Type | ExcelLambda |
| Location | <code>SplitText\_byHadynWiseman</code> |

[^Top](#oa-robot-definitions)
