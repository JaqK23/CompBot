# OA Robot Definitions

**CompBot.xlsm** contains definitions for:

[9 Robot Commands](#command-definitions)  

  

## Available Robot Commands

[Backup](#backup) | [Blank](#blank) | [Case](#case) | [Comp](#comp) | [Create](#create) | [Data](#data) | [Default](#default) | [External](#external) | [Generate](#generate) | [Import](#import) | [Inputs](#inputs) | [International](#international) | [Lambda](#lambda) | [Language](#language) | [levels](#levels) | [Levels](#levels) | [Library](#library) | [Settings](#settings) | [Setup](#setup) | [Sheet](#sheet) | [Sheets](#sheets) | [Table](#table) | [Update](#update) | [Workbook](#workbook) | [WorkSheets](#worksheets)

### Backup

| Name | Description |
| --- | --- |
| [Backup Sheets in Workbook](#backup-sheets-in-workbook) | Copy all sheets in workbook for backup |

### Blank

| Name | Description |
| --- | --- |
| [Create Blank Sheet](#create-blank-sheet) | Creates a blank sheet named based on cell value (Sht otherwise) |

### Case

| Name | Description |
| --- | --- |
| [Create Case Inputs Sheet](#create-case-inputs-sheet) | Creates a case inputs sheet with the inputs for the current case |
| [Setup Case](#setup-case) | Setup case by running Backup and Level sheet creation |

### Comp

| Name | Description |
| --- | --- |
| [Create Case Inputs Sheet](#create-case-inputs-sheet) | Creates a case inputs sheet with the inputs for the current case |

### Create

| Name | Description |
| --- | --- |
| [Create Data Table](#create-data-table) | Creates a data table on the current sheet for level sheets "L#" or active cell = "Example#" |

### Data

| Name | Description |
| --- | --- |
| [Create Data Table](#create-data-table) | Creates a data table on the current sheet for level sheets "L#" or active cell = "Example#" |

### Default

| Name | Description |
| --- | --- |
| [Default Settings](#default-settings) | Updates default settings to those in Default Settings sheet |

### External

| Name | Description |
| --- | --- |
| [Import Lambdas From](#import-lambdas-from) | Import lambdas from a specified workbook |

### Generate

| Name | Description |
| --- | --- |
| [Create Data Table](#create-data-table) | Creates a data table on the current sheet for level sheets "L#" or active cell = "Example#" |

### Import

| Name | Description |
| --- | --- |
| [Import Lambda Library](#import-lambda-library) | Imports CompBot's lambda collection into active workbook |
| [Import Lambdas From](#import-lambdas-from) | Import lambdas from a specified workbook |

### Inputs

| Name | Description |
| --- | --- |
| [Create Case Inputs Sheet](#create-case-inputs-sheet) | Creates a case inputs sheet with the inputs for the current case |

### International

| Name | Description |
| --- | --- |
| [Default Settings](#default-settings) | Updates default settings to those in Default Settings sheet |

### Lambda

| Name | Description |
| --- | --- |
| [Import Lambda Library](#import-lambda-library) | Imports CompBot's lambda collection into active workbook |
| [Import Lambdas From](#import-lambdas-from) | Import lambdas from a specified workbook |

### Language

| Name | Description |
| --- | --- |
| [Default Settings](#default-settings) | Updates default settings to those in Default Settings sheet |

### levels

| Name | Description |
| --- | --- |
| [Create Level Sheets](#create-level-sheets) | Creates a sheet for each level in the Case sheet |

### Levels

| Name | Description |
| --- | --- |
| [Setup Case](#setup-case) | Setup case by running Backup and Level sheet creation |

### Library

| Name | Description |
| --- | --- |
| [Import Lambda Library](#import-lambda-library) | Imports CompBot's lambda collection into active workbook |
| [Import Lambdas From](#import-lambdas-from) | Import lambdas from a specified workbook |

### Settings

| Name | Description |
| --- | --- |
| [Default Settings](#default-settings) | Updates default settings to those in Default Settings sheet |

### Setup

| Name | Description |
| --- | --- |
| [Create Blank Sheet](#create-blank-sheet) | Creates a blank sheet named based on cell value (Sht otherwise) |
| [Create Level Sheets](#create-level-sheets) | Creates a sheet for each level in the Case sheet |
| [Setup Case](#setup-case) | Setup case by running Backup and Level sheet creation |

### Sheet

| Name | Description |
| --- | --- |
| [Create Blank Sheet](#create-blank-sheet) | Creates a blank sheet named based on cell value (Sht otherwise) |
| [Create Case Inputs Sheet](#create-case-inputs-sheet) | Creates a case inputs sheet with the inputs for the current case |
| [Create Level Sheets](#create-level-sheets) | Creates a sheet for each level in the Case sheet |

### Sheets

| Name | Description |
| --- | --- |
| [Backup Sheets in Workbook](#backup-sheets-in-workbook) | Copy all sheets in workbook for backup |
| [Setup Case](#setup-case) | Setup case by running Backup and Level sheet creation |

### Table

| Name | Description |
| --- | --- |
| [Create Data Table](#create-data-table) | Creates a data table on the current sheet for level sheets "L#" or active cell = "Example#" |

### Update

| Name | Description |
| --- | --- |
| [Default Settings](#default-settings) | Updates default settings to those in Default Settings sheet |

### Workbook

| Name | Description |
| --- | --- |
| [Import Lambda Library](#import-lambda-library) | Imports CompBot's lambda collection into active workbook |
| [Import Lambdas From](#import-lambdas-from) | Import lambdas from a specified workbook |

### WorkSheets

| Name | Description |
| --- | --- |
| [Backup Sheets in Workbook](#backup-sheets-in-workbook) | Copy all sheets in workbook for backup |

  

## Command Definitions

  

### Backup Sheets in Workbook

*Copy all sheets in workbook for backup*

`@CompBot.xlsm` `!VBA Macro Command` `#Backup` `#WorkSheets` `#Sheets`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.Backup](./VBA/modCaseSetup.bas#L32)() |
| Macro Workbook Connection | ThisWorkbook |
| Launch Codes | ``` BS ``` |

[^Top](#oa-robot-definitions)

  

### Create Blank Sheet

*Creates a blank sheet named based on cell value (Sht otherwise)*

`@CompBot.xlsm` `!VBA Macro Command` `#Sheet` `#Setup` `#Blank`

| Property | Value |
| --- | --- |
| Macro Expression | [modMisc.CreateBlankSheet](./VBA/modMisc.bas#L11)([[ActiveCell]]) |
| Macro Workbook Connection | ThisWorkbook |
| Launch Codes | 1. ``` BS ``` 2. ``` CBS ``` 3. ``` S ``` |

[^Top](#oa-robot-definitions)

  

### Create Case Inputs Sheet

*Creates a case inputs sheet with the inputs for the current case*

`@CompBot.xlsm` `!VBA Macro Command` `#Inputs` `#Case` `#Comp` `#Sheet`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.CreateCaseInputsSheet](./VBA/modCaseSetup.bas#L363)([[ActiveCell]]) |
| Macro Workbook Connection | ThisWorkbook |
| Launch Codes | 1. ``` CIS ``` 2. ``` IS ``` |

[^Top](#oa-robot-definitions)

  

### Create Data Table

*Creates a data table on the current sheet for level sheets "L#" or active cell = "Example#"*

`@CompBot.xlsm` `!VBA Macro Command` `#Data` `#Table` `#Generate` `#Create`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.CreateDataTable](./VBA/modCaseSetup.bas#L175)([[NewTableTargetToRight]]) |
| Launch Codes | 1. ``` CDT ``` 2. ``` DT ``` |

[^Top](#oa-robot-definitions)

  

### Create Level Sheets

*Creates a sheet for each level in the Case sheet*

`@CompBot.xlsm` `!VBA Macro Command` `#levels` `#Sheet` `#Setup`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.CreateLevelSheets](./VBA/modCaseSetup.bas#L63)() |
| Launch Codes | 1. ``` CL ``` 2. ``` CLS ``` |

[^Top](#oa-robot-definitions)

  

### Default Settings

*Updates default settings to those in Default Settings sheet*

`@CompBot.xlsm` `!VBA Macro Command` `#Update` `#Default` `#Settings` `#International` `#Language`

| Property | Value |
| --- | --- |
| Macro Expression | [modMisc.DefaultSettings](./VBA/modMisc.bas#L71)() |
| Launch Codes | ``` DS ``` |

[^Top](#oa-robot-definitions)

  

### Import Lambda Library

*Imports CompBot's lambda collection into active workbook*

`@CompBot.xlsm` `!VBA Macro Command` `#Import` `#Lambda` `#Library` `#Workbook`

| Property | Value |
| --- | --- |
| Macro Expression | [modImportLambdas.ImportAllLambdas](./VBA/modImportLambdas.bas)("CompBot.xlsb") |
| Macro Workbook Connection | Lambda Robot |
| Launch Codes | ``` IL ``` |

[^Top](#oa-robot-definitions)

  

### Import Lambdas From

*Import lambdas from a specified workbook*

`@CompBot.xlsm` `!VBA Macro Command` `#Import` `#Lambda` `#External` `#Workbook` `#Library`

| Property | Value |
| --- | --- |
| Macro Expression | [modLambdas.ImportLambdasFrom](./VBA/modLambdas.bas#L40)() |
| Launch Codes | ``` ILF ``` |

[^Top](#oa-robot-definitions)

  

### Setup Case

*Setup case by running Backup and Level sheet creation*

`@CompBot.xlsm` `!VBA Macro Command` `#Setup` `#Case` `#Levels` `#Sheets`

| Property | Value |
| --- | --- |
| Macro Expression | [modCaseSetup.Setup](./VBA/modCaseSetup.bas#L18)() |
| Launch Codes | ``` SC ``` |

[^Top](#oa-robot-definitions)
