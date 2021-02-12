# BankHoliday PowerShell Module

BankHoliday is a Windows PowerShell module that can be used tp add, remove, or list bank holiday entries from mailboxes. The module uses the EWS Managed API 2.2 to connect to mailboxes, in conjunction with MSAL.Net for OAuth authentication. 

**IMPORTANT** The current version of the module was tested against Microsoft 365 (Exchange Online) maioboxes only. It is not guaranteed to work against on premises mailboxes. 

The current version of the module exports 4 functions:
- Add-BankHoliday
- Get-BankHoliday
- Remove-BankHoliday
- Get-CountryNames

As the name suggests, *Add-BankHoliday* can be used to add bank holidays to mailboxes, *Get-BankHoliday* can be used to list bank holidays in mailboxes, and *Remove-BankHoliday* can be used to remove bank holidays from mailboxes. 

# Installing the module

To install the module, run Windows PowerShell and run the following command to list the paths to the PowerShell module folders in your system:
```sh
$env:PSModulePath -split ';'
```

You should see a result similar to:

```sh
C:\Users\User1\Documents\WindowsPowerShell\Modules
C:\Program Files\WindowsPowerShell\Modules
C:\WINDOWS\system32\WindowsPowerShell\v1.0\Modules
```

Download and extract the contents of the [v1.0 release](https://github.com/andreighita/BankHolidayPowershellModule/files/5972306/BankHoliday_v1.0.zip) into one of the folders listed above. 
If you want the module to be available to the current user only, then extract the archive content into the *C:\Users* location, alternatively, use one of the otehr folders to make the module available to everyone logged on to the system. 

Do not extract the contents of the *BankHoliday* folder directly under *Modules*. Please the whole *BankHoliday* folder in the *Modules* folder instead.

![Modules Folder](https://github.com/andreighita/BankHolidayPowershellModule/blob/main/Artefacts/ModulesFolder.PNG?raw=true)

# Module folder contents

The module folder contains the .psm1 PowerShell module and the .psd1 PowerShell module manifest. 

In addition, the folder Assemblies contains the [Ews Managed API 2.2](https://www.microsoft.com/en-us/download/details.aspx?id=42951) assemblies and the [MSAL.Net](https://www.nuget.org/packages/Microsoft.Identity.Client) version 4.25 assembly. 

The HOL file contains 41 csv files created based on the .HOL files the Microsoft Outlook ships with. These contain bank holiday details until year 2026. Each file contains the country name and bank holiday name for a different locale. 

The locales included are *Arabic (Saudi Arabia)*, *Bulgarian (Bulgaria)*, *Chinese (Simplified,  PRC)*, *Chinese (Traditional,  Taiwan)*, *Croatian (Croatia)*, *Czech (Czech Republic)*, *Danish (Denmark)*, *Dutch (Netherlands)*, *English (United States)*, *Estonian (Estonia)*, *Finnish (Finland)*, *French (France)*, *German (Germany)*, *Greek (Greece)*, *Hebrew (Israel)*, *Hindi (India)*, *Hungarian (Hungary)*, *Indonesian (Indonesia)*, *Italian (Italy)*, *Japanese (Japan)*, *Kazakh (Kazakhstan)*, *Korean (Korea)*, *Latvian (Latvia)*, *Lithuanian (Lithuania)*, *Malay (Malaysia)*, *Norwegian,  Bokmål (Norway)*, *Polish (Poland)*, *Portuguese (Brazil)*, *Portuguese (Portugal)*, *Romanian (Romania)*, *Russian (Russia)*, *Serbian (Latin,  Serbia and Montenegro (Former))*, *Serbian (Latin,  Serbia)*, *Slovak (Slovakia)*, *Slovenian (Slovenia)*, *Spanish (Spain)*, *Swedish (Sweden)*, *Thai (Thailand)*, *Turkish (Turkey)*, *Ukrainian (Ukraine)*, *Vietnamese (Vietnam)*.

![Modules Folder Contents](https://github.com/andreighita/BankHolidayPowershellModule/blob/main/Artefacts/ModulesFolderContents.PNG?raw=true)

# Using the functions exposed by the module

For information on how to run the three bank holiday related functions please run the following commands to read the help information and examples for each function:

```sh
Get-Help Add-BankHoliday -ShowWindow
Get-Help Get-BankHoliday -ShowWindow
Get-Help Remove-BankHoliday -ShowWindow
```
![Get Help](https://github.com/andreighita/BankHolidayPowershellModule/blob/main/Artefacts/Add-BankHoliday_Get-Help.PNG)
For a GUI - visual editor of the commands and to make sure you match the various parameters to the correct parameter sets please use one of the following commands:

```sh
Show-Command Add-BankHoliday
Show-Command Get-BankHoliday
Show-Command Remove-BankHoliday
```
![Show Command](https://github.com/andreighita/BankHolidayPowershellModule/blob/main/Artefacts/Add-BankHoliday_ShowCommand.PNG)
# Optional Parameters

If you have deployed the module files in one of the Module folders on your machine as per the instructions, you do not need to specify the *EWSManagedAPIPath*, *MSALNetDllPath*, *HolFolderPath* parameters or their values. The code behind each function will look up the required files automatically in the Module folders on your system.  
# License
MIT

# Disclaimer

This is NOT something supported by Microsoft Customer Service and Support. If you're on this page, using this script, you're on your own!
