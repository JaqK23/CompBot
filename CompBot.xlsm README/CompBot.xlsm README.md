# OA Robot Definitions

**CompBot.xlsm** contains definitions for:

[38 Robot Commands](#command-definitions)  
[10 Robot Texts](#text-definitions)  

  

## Available Robot Commands

[GoTo](#goto) | [LAMBDA](#lambda) | [Maintenance](#maintenance) | [Name](#name) | [Paste](#paste) | [Prep](#prep) | [Settings](#settings) | [WrapWith](#wrapwith) | [Other](#other)

### GoTo

| Name | Description |
| --- | --- |
| [Goto Similar Background Color](#goto-similar-background-color) | Goto similar background color. |

### LAMBDA

| Name | Description |
| --- | --- |
| [Import Lambda Library](#import-lambda-library) | Imports CompBot's lambda collection into active workbook |
| [Import Lambdas From](#import-lambdas-from) | Import lambdas from a specified workbook |

### Maintenance

| Name | Description |
| --- | --- |
| [Clear Lambdas](#clear-lambdas) | Clear lambdas that aren't in Lambdas table - for Comp Bot maintenance |

### Name

| Name | Description |
| --- | --- |
| [Name All Used Ranges](#name-all-used-ranges) | Renames used range in all sheets |
| [Name Used Ranges](#name-used-ranges) | Names used range in all sheets starting with the prefix as provided in the active cell |

### Paste

| Name | Description |
| --- | --- |
| [Save Answers To Left](#save-answers-to-left) | Saves references to the selected cells in the green answer cells to the left on the same row. |

### Prep

| Name | Description |
| --- | --- |
| [Backup Sheets in Workbook](#backup-sheets-in-workbook) | Copy all sheets in workbook for backup |
| [Create Blank Sheet](#create-blank-sheet) | Creates a blank sheet named based on cell value (Sht otherwise) |
| [Create Bonus Sheet](#create-bonus-sheet) | Creates bonus sheet "B" with bonus questions |
| [Create Case Inputs Sheet](#create-case-inputs-sheet) | Creates a case inputs sheet with the inputs for the current case |
| [Create Data Table](#create-data-table) | Creates a data table on the current sheet for level sheets "L#" or active cell = "Example#" |
| [Create Level Sheets](#create-level-sheets) | Creates a sheet for each level in the Case sheet |
| [Full Setup Case](#full-setup-case) | Setup case by running Backup and Level sheet creation |
| [Import Lambda Library](#import-lambda-library) | Imports CompBot's lambda collection into active workbook |
| [Import Lambdas From](#import-lambdas-from) | Import lambdas from a specified workbook |
| [Rename Sheets](#rename-sheets) | Rename multi-word sheet names |
| [Save Copy of File](#save-copy-of-file) | Enable editing and save copy of file with suffix based on active cell (otherwise Working) |

### Settings

| Name | Description |
| --- | --- |
| [Default Settings](#default-settings) | Updates default settings to those in Default Settings sheet |
| [Toggle Calculation Mode](#toggle-calculation-mode) | Toggles calculation mode and places current mode notice in StatusBar |
| [Toggle Iterative Calculation](#toggle-iterative-calculation) | Toggles iterative calculation and sets status in status bar |

### WrapWith

| Name | Description |
| --- | --- |
| [Align Array to Right](#align-array-to-right) | Aligns array to the right with blanks to left |
| [Difference of Array Columns by Row](#difference-of-array-columns-by-row) | Returns first column - last column in array |
| [Find Above In Left](#find-above-in-left) | Find above value(s) in left value(s). |
| [Keep Cell of Array](#keep-cell-of-array) | Keep selected cell of array. |
| [Lookup Value by Row](#lookup-value-by-row) | Returns RC location of a list of values in a range |
| [Negative Values](#negative-values) | Negative values. with errors as blank |
| [Sequence of Row Count](#sequence-of-row-count) | Sequence of row count of array variable |
| [Split Text by Delimiter Above](#split-text-by-delimiter-above) | Split text by delimiter in the cell above and trim output. |
| [Substitute Blank for 0](#substitute-blank-for-0) | Substitute blank values for 0 |
| [Substitute Blank for 1](#substitute-blank-for-1) | Substitute blank cells for 1 |
| [Wrap in ABS](#wrap-in-abs) | Wrap in abs. |
| [Wrap in Concat](#wrap-in-concat) | Wrap in concat. |
| [Wrap in Drop First Row](#wrap-in-drop-first-row) | Wrap in drop first row |
| [Wrap in IFERROR with space](#wrap-in-iferror-with-space) | Wrap in iferror. |
| [Wrap in Take by Copied Cell Columns](#wrap-in-take-by-copied-cell-columns) | Wrap in take by copied cell columns. |
| [Wrap in UNICHAR](#wrap-in-unichar) | Wrap in unichar. |
| [Wrap in UNICODE](#wrap-in-unicode) | Wrap in unicode. |
| [Wrap with IFERROR TEXTBEFORE space](#wrap-with-iferror-textbefore-space) | Wraps current formula with TEXTBEFORE space with IFERROR in case space doesn't exist |

### Other

| Name | Description |
| --- | --- |
| [Eggy](#eggy) | Easter Egg Fun |

  

## Available Robot Texts

| Name | Description |
| --- | --- |
| [ADDRESSES_byDiarmuidEarly.lambda](#addresses_bydiarmuidearly.lambda) | Definition of ADDRESSES_byDiarmuidEarly lambda function. |
| [BiRow_byPeterBartholomew.lambda](#birow_bypeterbartholomew.lambda) | Definition of BiRow_byPeterBartholomew lambda function. |
| [DiffByRow_byJaqKennedy.lambda](#diffbyrow_byjaqkennedy.lambda) | Definition of DiffByRow_byJaqKennedy lambda function. |
| [FilterArray_byErikOehm.lambda](#filterarray_byerikoehm.lambda) | Definition of FilterArray_byErikOehm lambda function. |
| [GRIDTOCOL_byLiannaGerrish.lambda](#gridtocol_byliannagerrish.lambda) | Definition of GRIDTOCOL_byLiannaGerrish lambda function. |
| [IFBLANK.lambda](#ifblank.lambda) | Definition of IFBLANK lambda function. |
| [IsInList_byErikOehm.lambda](#isinlist_byerikoehm.lambda) | Definition of IsInList_byErikOehm lambda function. |
| [LkpRC.lambda](#lkprc.lambda) | Definition of LkpRC lambda function. |
| [LkpRCByRow.lambda](#lkprcbyrow.lambda) | Definition of LkpRCByRow lambda function. |
| [RightAlignedArray_byJaqKennedy.lambda](#rightalignedarray_byjaqkennedy.lambda) | Definition of RightAlignedArray_byJaqKennedy lambda function. |

  

## Command Definitions

  

### Align Array to Right

*Aligns array to the right with blanks to left*

`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`

| Property | Value |
| --- | --- |
| Formula | ``` =RightAlignedArray_byJaqKennedy(IFBLANK([[ActiveCell::Formula]],"")) ``` |
| Formula Dependencies | 1. [RightAlignedArray_byJaqKennedy.lambda](#rightalignedarray_byjaqkennedy.lambda) 2. [IFBLANK.lambda](#ifblank.lambda) |
| Launch Codes | ``` ar ``` |

[^Top](#oa-robot-definitions)

  

### Backup Sheets in Workbook

*Copy all sheets in workbook for backup*

`@CompBot.xlsm` `!VBA Macro Command` `#Prep`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.Backup](./VBA/modCaseSetup.bas#L75)() |
| Macro Workbook Connection | ThisWorkbook |
| Launch Codes | ``` BU ``` |

[^Top](#oa-robot-definitions)

  

### Clear Lambdas

*Clear lambdas that aren't in Lambdas table - for Comp Bot maintenance*

`@CompBot.xlsm` `!VBA Macro Command` `#Maintenance`

| Property | Value |
| --- | --- |
| Macro Expression | [modLambdas.ClearLambdas](./VBA/modLambdas.bas#L93)() |
| Launch Codes | ``` CL ``` |

[^Top](#oa-robot-definitions)

  

### Create Blank Sheet

*Creates a blank sheet named based on cell value (Sht otherwise)*

`@CompBot.xlsm` `!VBA Macro Command` `#Prep`

| Property | Value |
| --- | --- |
| Macro Expression | [modMisc.CreateBlankSheet](./VBA/modMisc.bas#L11)([[ActiveCell]]) |
| Macro Workbook Connection | ThisWorkbook |
| Launch Codes | 1. ``` BS ``` 2. ``` CBS ``` 3. ``` S ``` |

[^Top](#oa-robot-definitions)

  

### Create Bonus Sheet

*Creates bonus sheet "B" with bonus questions*

`@CompBot.xlsm` `!VBA Macro Command` `#Prep`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.CreateBonusSheet](./VBA/modCaseSetup.bas#L312)() |
| Launch Codes | 1. ``` CB ``` 2. ``` CBS ``` |

[^Top](#oa-robot-definitions)

  

### Create Case Inputs Sheet

*Creates a case inputs sheet with the inputs for the current case*

`@CompBot.xlsm` `!VBA Macro Command` `#Prep`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.CreateCaseInputsSheet](./VBA/modCaseSetup.bas#L678)([[ActiveCell]]) |
| Macro Workbook Connection | ThisWorkbook |
| Launch Codes | 1. ``` CIS ``` 2. ``` IS ``` |

[^Top](#oa-robot-definitions)

  

### Create Data Table

*Creates a data table on the current sheet for level sheets "L#" or active cell = "Example#"*

`@CompBot.xlsm` `!VBA Macro Command` `#Prep`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.CreateDataTable](./VBA/modCaseSetup.bas#L486)([[NewTableTargetToRight]]) |
| Launch Codes | 1. ``` CDT ``` 2. ``` DT ``` |

[^Top](#oa-robot-definitions)

  

### Create Level Sheets

*Creates a sheet for each level in the Case sheet*

`@CompBot.xlsm` `!VBA Macro Command` `#Prep`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.CreateLevelSheets](./VBA/modCaseSetup.bas#L106)() |
| Launch Codes | 1. ``` CL ``` 2. ``` CLS ``` |

[^Top](#oa-robot-definitions)

  

### Default Settings

*Updates default settings to those in Default Settings sheet*

`@CompBot.xlsm` `!VBA Macro Command` `#Settings`

| Property | Value |
| --- | --- |
| Macro Expression | [modMisc.DefaultSettings](./VBA/modMisc.bas#L71)() |
| Launch Codes | ``` DS ``` |

[^Top](#oa-robot-definitions)

  

### Difference of Array Columns by Row

*Returns first column - last column in array*

`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`

| Property | Value |
| --- | --- |
| Formula | ``` =DiffByRow_byJaqKennedy([[ActiveCell::Formula]]) ``` |
| Formula Dependencies | [DiffByRow_byJaqKennedy.lambda](#diffbyrow_byjaqkennedy.lambda) |
| Launch Codes | 1. ``` diff ``` 2. ``` dbr ``` |

[^Top](#oa-robot-definitions)

  

### Eggy

*Easter Egg Fun*

`@CompBot.xlsm` `!VBA Macro Command`  

| Property | Value |
| --- | --- |
| Macro Expression | [modUtilities.Eggy](./VBA/modUtilities.bas#L244)() |

[^Top](#oa-robot-definitions)

  

### Find Above In Left

*Find above value(s) in left value(s).*

`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`

| Property | Value |
| --- | --- |
| Formula | ``` =1-ISERROR(FIND([[ActiveCell.Offset(-1,0).SpillParent.SpillingToRange]],[[ActiveCell.Offset(0,-1).SpillParent.SpillingToRange]])) ``` |
| Launch Codes | ``` fal ``` |

[^Top](#oa-robot-definitions)

  

### Full Setup Case

*Setup case by running Backup and Level sheet creation*

`@CompBot.xlsm` `!VBA Macro Command` `#Prep`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.Setup](./VBA/modCaseSetup.bas#L18)() |
| Keyboard Shortcut | ``` ^+g ``` |
| Command After | [Import Lambda Library](#import-lambda-library) |
| Launch Codes | ``` SC ``` |

[^Top](#oa-robot-definitions)

  

### Goto Similar Background Color

*Goto similar background color.*

`@CompBot.xlsm` `!VBA Macro Command` `#GoTo`

| Property | Value |
| --- | --- |
| Macro Expression | [modGoToSpecial.GotoSimilarBackgroundColor](./VBA/modGoToSpecial.bas#L15)() |
| Launch Codes | ``` SBC ``` |

[^Top](#oa-robot-definitions)

  

### Import Lambda Library

*Imports CompBot's lambda collection into active workbook*

`@CompBot.xlsm` `!VBA Macro Command` `#LAMBDA` `#Prep`

| Property | Value |
| --- | --- |
| Macro Expression | [modImportLambdas.ImportAllLambdas](./VBA/modImportLambdas.bas)("CompBot.xlsm") |
| Macro Workbook Connection | Lambda Robot |
| Launch Codes | ``` IL ``` |

[^Top](#oa-robot-definitions)

  

### Import Lambdas From

*Import lambdas from a specified workbook*

`@CompBot.xlsm` `!VBA Macro Command` `#LAMBDA` `#Prep`

| Property | Value |
| --- | --- |
| Macro Expression | [modLambdas.ImportLambdasFrom](./VBA/modLambdas.bas#L40)() |
| Launch Codes | ``` ILF ``` |

[^Top](#oa-robot-definitions)

  

### Keep Cell of Array

*Keep selected cell of array.*

`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`

| Property | Value |
| --- | --- |
| Formula | ``` =CHOOSEROWS(CHOOSECOLS([[ActiveCell.SpillParent::Formula]],{{Selected_Column_Indexes_In_Spilling_Range}}),{{Selected_Row_Indexes_In_Spilling_Range}}) ``` |
| Destination Range Address | ``` [[ActiveCell.SpillParent]] ``` |
| Launch Codes | ``` KTC ``` |

[^Top](#oa-robot-definitions)

  

### Lookup Value by Row

*Returns RC location of a list of values in a range*

`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`

| Property | Value |
| --- | --- |
| Formula | ``` =LkpRCByRow([[ActiveCell::Formula]], [[Clipboard::Address]]) ``` |
| Formula Dependencies | 1. [LkpRCByRow.lambda](#lkprcbyrow.lambda) 2. [BiRow_byPeterBartholomew.lambda](#birow_bypeterbartholomew.lambda) 3. [LkpRC.lambda](#lkprc.lambda) 4. [FilterArray_byErikOehm.lambda](#filterarray_byerikoehm.lambda) 5. [IsInList_byErikOehm.lambda](#isinlist_byerikoehm.lambda) 6. [GRIDTOCOL_byLiannaGerrish.lambda](#gridtocol_byliannagerrish.lambda) 7. [ADDRESSES_byDiarmuidEarly.lambda](#addresses_bydiarmuidearly.lambda) |
| Launch Codes | ``` lkp ``` |

[^Top](#oa-robot-definitions)

  

### Name All Used Ranges

*Renames used range in all sheets*

`@CompBot.xlsm` `!VBA Macro Command` `#Name`

| Property | Value |
| --- | --- |
| Macro Expression | [modNames.NameAllUsedRanges](./VBA/modNames.bas#L30)() |
| Launch Codes | 1. ``` NA ``` 2. ``` NAUR ``` |

[^Top](#oa-robot-definitions)

  

### Name Used Ranges

*Names used range in all sheets starting with the prefix as provided in the active cell*

`@CompBot.xlsm` `!VBA Macro Command` `#Name`

| Property | Value |
| --- | --- |
| Macro Expression | [modNames.NameUsedRanges](./VBA/modNames.bas#L11)([[ActiveCell]]) |
| Launch Codes | ``` NURS ``` |

[^Top](#oa-robot-definitions)

  

### Negative Values

*Negative values. with errors as blank*

`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`

| Property | Value |
| --- | --- |
| Formula | ``` =IFERROR(-([[ActiveCell::Formula]]),"") ``` |
| Launch Codes | ``` neg ``` |

[^Top](#oa-robot-definitions)

  

### Rename Sheets

*Rename multi-word sheet names*

`@CompBot.xlsm` `!VBA Macro Command` `#Prep`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.RenameSht](./VBA/modCaseSetup.bas#L37)() |
| Launch Codes | ``` rs ``` |

[^Top](#oa-robot-definitions)

  

### Save Answers To Left

*Saves references to the selected cells in the green answer cells to the left on the same row.*

`@CompBot.xlsm` `!VBA Macro Command` `#Paste`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.SaveAnswersToLeft](./VBA/modCaseSetup.bas#L844)() |
| Launch Codes | ``` SAL ``` |

[^Top](#oa-robot-definitions)

  

### Save Copy of File

*Enable editing and save copy of file with suffix based on active cell (otherwise Working)*

`@CompBot.xlsm` `!VBA Macro Command` `#Prep`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.SaveCopy](./VBA/modCaseSetup.bas#L897)([[ActiveCell]]) |
| Launch Codes | ``` SA ``` |

[^Top](#oa-robot-definitions)

  

### Sequence of Row Count

*Sequence of row count of array variable*

`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`

| Property | Value |
| --- | --- |
| Formula | ``` =SEQUENCE(ROWS([[ActiveCell::Formula]]),1,1,1) ``` |
| Launch Codes | ``` src ``` |

[^Top](#oa-robot-definitions)

  

### Split Text by Delimiter Above

*Split text by delimiter in the cell above and trim output.*

`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`

| Property | Value |
| --- | --- |
| Formula | ``` =TRIM(TEXTSPLIT([[ActiveCell::Formula]],,[[ActiveCell.Offset(-1,0)]],TRUE)) ``` |
| Launch Codes | ``` ST ``` |

[^Top](#oa-robot-definitions)

  

### Substitute Blank for 0

*Substitute blank values for 0*

`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`

| Property | Value |
| --- | --- |
| Formula | ``` =SUBSTITUTE([[ActiveCell::Formula]],"",0) ``` |
| Launch Codes | ``` sb0 ``` |

[^Top](#oa-robot-definitions)

  

### Substitute Blank for 1

*Substitute blank cells for 1*

`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`

| Property | Value |
| --- | --- |
| Formula | ``` =SUBSTITUTE([[ActiveCell::Formula]],"",1) ``` |
| Launch Codes | ``` sb1 ``` |

[^Top](#oa-robot-definitions)

  

### Toggle Calculation Mode

*Toggles calculation mode and places current mode notice in StatusBar*

`@CompBot.xlsm` `!VBA Macro Command` `#Settings`

| Property | Value |
| --- | --- |
| Macro Expression | [modMisc.ToggleCalculationMode](./VBA/modMisc.bas#L122)() |
| Launch Codes | ``` TC ``` |

[^Top](#oa-robot-definitions)

  

### Toggle Iterative Calculation

*Toggles iterative calculation and sets status in status bar*

`@CompBot.xlsm` `!VBA Macro Command` `#Settings`

| Property | Value |
| --- | --- |
| Macro Expression | [modMisc.ToggleIterativeCalculation](./VBA/modMisc.bas#L139)() |
| Launch Codes | ``` IC ``` |

[^Top](#oa-robot-definitions)

  

### Wrap in ABS

*Wrap in abs.*

`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`

| Property | Value |
| --- | --- |
| Formula | ``` =ABS([[ActiveCell::Formula]]) ``` |
| Launch Codes | ``` ABS ``` |

[^Top](#oa-robot-definitions)

  

### Wrap in Concat

*Wrap in concat.*

`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`

| Property | Value |
| --- | --- |
| Formula | ``` =CONCAT([[ActiveCell::Formula]]) ``` |
| Launch Codes | ``` con ``` |

[^Top](#oa-robot-definitions)

  

### Wrap in Drop First Row

*Wrap in drop first row*

`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`

| Property | Value |
| --- | --- |
| Formula | ``` =DROP([[ActiveCell::Formula]],1) ``` |
| Launch Codes | ``` DF ``` |

[^Top](#oa-robot-definitions)

  

### Wrap in IFERROR with space

*Wrap in iferror.*

`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`

| Property | Value |
| --- | --- |
| Formula | ``` =IFERROR([[ActiveCell::Formula]]," ") ``` |
| Launch Codes | ``` err ``` |

[^Top](#oa-robot-definitions)

  

### Wrap in Take by Copied Cell Columns

*Wrap in take by copied cell columns.*

`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`

| Property | Value |
| --- | --- |
| Formula | ``` =TAKE([[ActiveCell::Formula]],,[[Clipboard]]) ``` |
| Launch Codes | ``` TCC ``` |

[^Top](#oa-robot-definitions)

  

### Wrap in UNICHAR

*Wrap in unichar.*

`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`

| Property | Value |
| --- | --- |
| Formula | ``` =UNICHAR([[ActiveCell::Formula]]) ``` |
| Launch Codes | ``` char ``` |

[^Top](#oa-robot-definitions)

  

### Wrap in UNICODE

*Wrap in unicode.*

`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`

| Property | Value |
| --- | --- |
| Formula | ``` =UNICODE([[ActiveCell::Formula]]) ``` |
| Launch Codes | ``` code ``` |

[^Top](#oa-robot-definitions)

  

### Wrap with IFERROR TEXTBEFORE space

*Wraps current formula with TEXTBEFORE space with IFERROR in case space doesn't exist*

`@CompBot.xlsm` `!Excel Formula Command` `#WrapWith`

| Property | Value |
| --- | --- |
| Formula | ``` =IFERROR(TEXTBEFORE([[ActiveCell::Formula]]," "),[[ActiveCell::Formula]]) ``` |
| Launch Codes | ``` tbse ``` |

[^Top](#oa-robot-definitions)

  

## Text Definitions

  

### ADDRESSES_byDiarmuidEarly.lambda

*Definition of ADDRESSES_byDiarmuidEarly lambda function.*

`@CompBot.xlsm` `!Excel Name Text`  

| Property | Value |
| --- | --- |
| Text | [ADDRESSES_byDiarmuidEarly.lambda](./Text/ADDRESSES_byDiarmuidEarly.lambda.txt) |
| Value | ``` ADDRESSES_byDiarmuidEarly = LAMBDA(rng, ADDRESS(ROW(rng), COLUMN(rng), 4)); ``` |
| Content Type | ExcelFormula |
| Location | ``` ADDRESSES_byDiarmuidEarly ``` |

[^Top](#oa-robot-definitions)

  

### BiRow_byPeterBartholomew.lambda

*Definition of BiRow_byPeterBartholomew lambda function.*

`@CompBot.xlsm` `!Excel Name Text`  

| Property | Value |
| --- | --- |
| Text | [BiRow_byPeterBartholomew.lambda](./Text/BiRow_byPeterBartholomew.lambda.txt) |
| Value | ``` BiRow_byPeterBartholomew = LAMBDA(array,function,LET(\\LambdaName, "BiRow_byPeterBartholomew", \\Description, "Perform array function by row", \\Source, "attributed to Peter Bartholomew: https://techcommunity.microsoft.com/discussions/excelgeneral/recursive-lambda-implementation-of-excels-reduce-function-/3949754/replies/3952653#M207809", IF(ROWS(array) = 1, function(array), VSTACK(BiRow_byPeterBartholomew(TAKE(array, ROWS(array) / 2), function), BiRow_byPeterBartholomew(DROP(array, ROWS(arr...``` |
| Content Type | ExcelFormula |
| Location | ``` BiRow_byPeterBartholomew ``` |

[^Top](#oa-robot-definitions)

  

### DiffByRow_byJaqKennedy.lambda

*Definition of DiffByRow_byJaqKennedy lambda function.*

`@CompBot.xlsm` `!Excel Name Text`  

| Property | Value |
| --- | --- |
| Text | [DiffByRow_byJaqKennedy.lambda](./Text/DiffByRow_byJaqKennedy.lambda.txt) |
| Value | ``` DiffByRow_byJaqKennedy = LAMBDA(Input,LET(\\LambdaName, "DiffByRow", \\CommandName, "Difference of array columns by row", \\Description, "Returns first column - last column in array", \\Source, "Jaq Kennedy", TAKE(Input, , 1) - TAKE(Input, , -1))); ``` |
| Content Type | ExcelFormula |
| Location | ``` DiffByRow_byJaqKennedy ``` |

[^Top](#oa-robot-definitions)

  

### FilterArray_byErikOehm.lambda

*Definition of FilterArray_byErikOehm lambda function.*

`@CompBot.xlsm` `!Excel Name Text`  

| Property | Value |
| --- | --- |
| Text | [FilterArray_byErikOehm.lambda](./Text/FilterArray_byErikOehm.lambda.txt) |
| Value | ``` FilterArray_byErikOehm = LAMBDA(data,column_indexes,filter_values,LET(\\LambdaName, "FilterArray", FILTER(data, BYROW(IsInList_byErikOehm(CHOOSECOLS(data, column_indexes), filter_values), LAMBDA(x, AND(x)))))); ``` |
| Content Type | ExcelFormula |
| Location | ``` FilterArray_byErikOehm ``` |

[^Top](#oa-robot-definitions)

  

### GRIDTOCOL_byLiannaGerrish.lambda

*Definition of GRIDTOCOL_byLiannaGerrish lambda function.*

`@CompBot.xlsm` `!Excel Name Text`  

| Property | Value |
| --- | --- |
| Text | [GRIDTOCOL_byLiannaGerrish.lambda](./Text/GRIDTOCOL_byLiannaGerrish.lambda.txt) |
| Value | ``` GRIDTOCOL_byLiannaGerrish = LAMBDA(grid,[ShowAll_0], LET( \\LambdaName, "GRIDTOCOL", items_grid, TOCOL(grid), addr_col, TOCOL(ADDRESSES_byDiarmuidEarly(grid)), rows, ROW(INDIRECT(addr_col)) * 1, cols, COLUMN(INDIRECT(addr_col)) * 1, rcnum, TOCOL(1000000 * rows + 1000 + cols), headers, HSTACK("Items", "Row #", "Column #", "Address", "RCRef"), data_grid, HSTACK(items_grid, rows, cols, addr_col, rcnum), IF( OR(ISOMITTED(ShowAll_0), ShowAll_0 <> 0), VSTACK(...``` |
| Content Type | ExcelFormula |
| Location | ``` GRIDTOCOL_byLiannaGerrish ``` |

[^Top](#oa-robot-definitions)

  

### IFBLANK.lambda

*Definition of IFBLANK lambda function.*

`@CompBot.xlsm` `!Excel Name Text`  

| Property | Value |
| --- | --- |
| Text | [IFBLANK.lambda](./Text/IFBLANK.lambda.txt) |
| Value | ``` IFBLANK = LAMBDA(value,value_if_blank, IF(ISBLANK(value), value_if_blank, value)); ``` |
| Content Type | ExcelFormula |
| Location | ``` IFBLANK ``` |

[^Top](#oa-robot-definitions)

  

### IsInList_byErikOehm.lambda

*Definition of IsInList_byErikOehm lambda function.*

`@CompBot.xlsm` `!Excel Name Text`  

| Property | Value |
| --- | --- |
| Text | [IsInList_byErikOehm.lambda](./Text/IsInList_byErikOehm.lambda.txt) |
| Value | ``` IsInList_byErikOehm = LAMBDA(array,list,LET(\\LambdaName, "IsInList", MAP(array, LAMBDA(x, OR(list = x))))); ``` |
| Content Type | ExcelFormula |
| Location | ``` IsInList_byErikOehm ``` |

[^Top](#oa-robot-definitions)

  

### LkpRC.lambda

*Definition of LkpRC lambda function.*

`@CompBot.xlsm` `!Excel Name Text`  

| Property | Value |
| --- | --- |
| Text | [LkpRC.lambda](./Text/LkpRC.lambda.txt) |
| Value | ``` LkpRC = LAMBDA(ToFind,Range,LET(\\LambdaName, "LkpRC", \\CommandName, "Lookup RC coordinates ", \\Description, "Looks up RC coordinates of matching cell(s) in range", CHOOSECOLS(FilterArray_byErikOehm(GRIDTOCOL_byLiannaGerrish(Range), {1}, ToFind), {2,3}))); ``` |
| Content Type | ExcelFormula |
| Location | ``` LkpRC ``` |

[^Top](#oa-robot-definitions)

  

### LkpRCByRow.lambda

*Definition of LkpRCByRow lambda function.*

`@CompBot.xlsm` `!Excel Name Text`  

| Property | Value |
| --- | --- |
| Text | [LkpRCByRow.lambda](./Text/LkpRCByRow.lambda.txt) |
| Value | ``` LkpRCByRow = LAMBDA(ToFind,Range,LET(\\LambdaName, "LkpRCByRow", \\CommandName, "Lookup Value by Row", \\Description, "Returns RC location of a list of values in a range", \\Source, "Jaq Kennedy", BiRow_byPeterBartholomew(ToFind, LAMBDA(a, IFERROR(CHOOSEROWS(LkpRC(a, Range), 1), {"",""}))))); ``` |
| Content Type | ExcelFormula |
| Location | ``` LkpRCByRow ``` |

[^Top](#oa-robot-definitions)

  

### RightAlignedArray_byJaqKennedy.lambda

*Definition of RightAlignedArray_byJaqKennedy lambda function.*

`@CompBot.xlsm` `!Excel Name Text`  

| Property | Value |
| --- | --- |
| Text | [RightAlignedArray_byJaqKennedy.lambda](./Text/RightAlignedArray_byJaqKennedy.lambda.txt) |
| Value | ``` /*Aligns contents of array to the right - ignoring blanks*/ RightAlignedArray_byJaqKennedy = LAMBDA(input,LET(\\LambdaName, "RightAlignedArray", \\CommandName, "Align array to right", \\Description, "Aligns array to the right with blanks to left", \\Source, "Jaq Kennedy", _ColsByRow, BYROW(input, COUNT), _Rows, ROWS(input), _Cols, COLUMNS(input), _ColIndex, IF(MOD(SEQUENCE(_Rows, _Cols) - 1, _Cols) + 1 - _Cols + _ColsByRow <= 0, -1, MOD(SEQUENCE(_Rows, _Cols) - 1, _Cols) + 1 - _Cols + _Cols...``` |
| Content Type | ExcelFormula |
| Location | ``` RightAlignedArray_byJaqKennedy ``` |

[^Top](#oa-robot-definitions)
