# ITAMTools

This is a Powershell module created to equip IWU IT Staff to update the ITAM Asset Management System. For more information about PowerShell Modules, please refere to: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_modules?view=powershell-7.2

To install:
Register-PSRepository -name Datacenter -SourceLocation "\\\\iwufiles\common\uit\datacenter\PowershellGallery"

install-module ITAMTools

To update:
update-module ITAMTools

To use in a new Powershell window:
import-module ITAMTools
Connect-ITAMServices

To publish updates:
Publish-Module -Path *Working Location* -Repository Datacenter
