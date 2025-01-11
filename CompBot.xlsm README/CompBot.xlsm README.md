# OA Robot Definitions

**CompBot.xlsm** contains definitions for:

[18 Robot Commands](#command-definitions)  

  

## Available Robot Commands

[GoTo](#goto) | [LAMBDA](#lambda) | [Maintenance](#maintenance) | [Name](#name) | [Paste](#paste) | [Prep](#prep) | [Settings](#settings)

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
| [Import Lambda Library](#import-lambda-library) | Imports CompBot's lambda collection into active workbook |
| [Import Lambdas From](#import-lambdas-from) | Import lambdas from a specified workbook |
| [Save Copy of File](#save-copy-of-file) | Enable editing and save copy of file with suffix based on active cell (otherwise Working) |
| [Setup Case](#setup-case) | Setup case by running Backup and Level sheet creation |

### Settings

| Name | Description |
| --- | --- |
| [Default Settings](#default-settings) | Updates default settings to those in Default Settings sheet |
| [Toggle Calculation Mode](#toggle-calculation-mode) | Toggles calculation mode and places current mode notice in StatusBar |
| [Toggle Iterative Calculation](#toggle-iterative-calculation) | Toggles iterative calculation and sets status in status bar |

  

## Command Definitions

  

### Backup Sheets in Workbook

*Copy all sheets in workbook for backup*

`@CompBot.xlsm` `!VBA Macro Command` `#Prep`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.Backup](./VBA/modCaseSetup.bas#L35)() |
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
| Macro Expression | [modCaseSetup.CreateBonusSheet](./VBA/modCaseSetup.bas#L262)() |
| Launch Codes | 1. ``` CB ``` 2. ``` CBS ``` |

[^Top](#oa-robot-definitions)

  

### Create Case Inputs Sheet

*Creates a case inputs sheet with the inputs for the current case*

`@CompBot.xlsm` `!VBA Macro Command` `#Prep`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.CreateCaseInputsSheet](./VBA/modCaseSetup.bas#L624)([[ActiveCell]]) |
| Macro Workbook Connection | ThisWorkbook |
| Launch Codes | 1. ``` CIS ``` 2. ``` IS ``` |

[^Top](#oa-robot-definitions)

  

### Create Data Table

*Creates a data table on the current sheet for level sheets "L#" or active cell = "Example#"*

`@CompBot.xlsm` `!VBA Macro Command` `#Prep`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.CreateDataTable](./VBA/modCaseSetup.bas#L432)([[NewTableTargetToRight]]) |
| Launch Codes | 1. ``` CDT ``` 2. ``` DT ``` |

[^Top](#oa-robot-definitions)

  

### Create Level Sheets

*Creates a sheet for each level in the Case sheet*

`@CompBot.xlsm` `!VBA Macro Command` `#Prep`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.CreateLevelSheets](./VBA/modCaseSetup.bas#L66)() |
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

  

### Save Answers To Left

*Saves references to the selected cells in the green answer cells to the left on the same row.*

`@CompBot.xlsm` `!VBA Macro Command` `#Paste`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.SaveAnswersToLeft](./VBA/modCaseSetup.bas#L786)() |
| Launch Codes | ``` SAL ``` |

[^Top](#oa-robot-definitions)

  

### Save Copy of File

*Enable editing and save copy of file with suffix based on active cell (otherwise Working)*

`@CompBot.xlsm` `!VBA Macro Command` `#Prep`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.SaveCopy](./VBA/modCaseSetup.bas#L839)([[ActiveCell]]) |
| Launch Codes | ``` SA ``` |

[^Top](#oa-robot-definitions)

  

### Setup Case

*Setup case by running Backup and Level sheet creation*

`@CompBot.xlsm` `!VBA Macro Command` `#Prep`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.Setup](./VBA/modCaseSetup.bas#L18)() |
| Launch Codes | ``` SC ``` |

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
