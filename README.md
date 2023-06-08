# The Excel Modules POSH Handler Module #

### Table of Contents: ###

* [Overview](#overview)
* [Installation](#installation)


## Overview ##

The ExcelModulesPoshHandler is a [PowerShell](https://learn.microsoft.com/en-us/powershell/) module which contains functions that helps you manage an Excel VBA Modules.  
By help is means:

  * [Exporting](https://github.com/BaronProsimir/ExcelModulesPoshHandler/wiki/Export_ExcelModulesAll) the members of VBAProject ( modules, classes and form ) from the passed Excel file/s,
  * Comming soon...

## Installation ##

ExcelModulesPoshHandler is currently **not** in [PowerShell Gallery](https://www.powershellgallery.com).
That's mean you have to download it first and then install it to your computer. By do this, follow the following steps:

  1. [Download](https://github.com/BaronProsimir/ExcelModulesPoshHandler/archive/refs/heads/master.zip) the ExcelModulesPoshHandler.
  2. Extract the downloaded folder:

  ```PowerShell
    Expand-Archive -Path "$env:USERPROFILE\Downloads\ExcelModulesPoshHandler-master.zip" -DestinationPath "$( ($env:PSModulePath -split ';')[0] )";
  ```

  More info about the command you can find [here.](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.archive/expand-archive?view=powershell-5.1)

  3. Add the following line to your [$PROFILE](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_profiles?view=powershell-5.1#the-profile-variable) file:

  ```PowerShell
    Import-Module -Name "$( ($env:PSModulePath -split ';')[0] )\ExcelModulesPoshHandler-master\ExcelModulesPoshHandler.psd1";
  ```

  4. **🎉🎉 DONE. Your are ready to use the module. 🎉🎉** 
