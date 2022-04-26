#Import-Module VMware.PowerCLI
#Import-Module SnipeitPS
#Import-Module MicrosoftPowerBIMgmt

enum outputTypes{
  Accessory
  Activity
  asset
  Category
  Company
  Component
  Consumable
  Department
  License
  Location
  Manufacturer
  Model
  Status
  Supplier
  User
}

enum credTypes{
  ITAM
  VMWare
  PowerBI
}

function Connect-ITAMServices {
  param(
    #validate parameter is path that exists
    [Parameter()][ValidateScript({
        if (-not ($_ | Test-Path)) {
          throw "File/folder does not exist"
        }
        if (-not ($_ | Test-Path -PathType Container)) {
          throw "The path argument must be a folder."
        }
        return $true
      })] [System.IO.FileInfo]$path = (Get-Location).path
  )

  $itamUrl = "http://itassets.iwunet.indwes.edu"
  $viServerUrl = "marn-vb-vmsa01.iwunet.indwes.edu" 

  #while Snipe connection is not valid, enter new API Key
  $validConnect = $false
  while (!$validConnect) {
    try {
      Connect-SnipeitPS -siteCred (Import-Clixml -Path ($path.FullName + "\ITAM.xml"))
      $validConnect = $true
    }
    catch {
      $host.ui.PromptForCredential("Store API key for ITAM", "Please enter a new API Key as password.", "$itamUrl", "ITAM") | Export-Clixml -Path ($path.FullName + "\ITAM.xml")
    }
  }

  #While VMWare connection is not valid, replace vmware credential
  $validConnect = $false
  while (!$validConnect) {
    try {
      Connect-VIServer -Server $viServerUrl -Credential (Import-Clixml -Path ($path.FullName + "\VMWare.xml"))
      $validConnect = $true
    }
    catch {
      Get-Credential -Message "Please provide your login for VMWare" | Export-Clixml -Path ($path.FullName + "\VMWare.xml")
    }

  }


  #Without an PowerBI admin account to create API key, automated PowerBI login is not possible:
  #https://stackoverflow.com/questions/61662906/powershell-automated-connection-to-power-bi-service-without-hardcoding-passwor
  #Connect-PowerBIServiceAccount
  $validConnect = $false
  while (!$validConnect) {
    try {
      Get-PowerBIAccessToken
      $validConnect = $true
    }
    catch {
      Connect-PowerBIServiceAccount
    }
  }

  #$gscn = New-Object System.Data.SqlClient.SqlConnection


  

}

function Sync-ITAMfromiSupport {
  #Connect to Sql
  $sqlServer = 'prd-sql01'
  $db = 'RP-ISUPPORT'
  
  $query = "SELECT 
[CUSTOMERS].LOGIN, [CUSTOMERS].EMAIL,
[ASSET].NAME, [ASSET_TYPES].TYPE, [ASSET].MFG, [ASSET].MODEL, [ASSET].SERIAL_NUMBER, [ASSET].TAG_NUMBER, [ASSET].LOCATION, [ASSET].DT_PURCHASE, [ASSET].DT_WAR_END, [ASSET].COMMENTS
FROM [RP-ISUPPORT].[dbo].[ASSET_OWNERS], [RP-ISUPPORT].[dbo].[CUSTOMERS], [RP-ISUPPORT].[dbo].[ASSET], [RP-ISUPPORT].[dbo].[ASSET_TYPES]
WHERE [ASSET_OWNERS].ID_ASSET_OWNERS = [CUSTOMERS].ID and [ASSET_OWNERS].ID_ASSET = [ASSET].ID and [ASSET_TYPES].ID = [ASSET].ID_TYPE"

  #Uses windows authentication with the current user.
  $ownedAssets = Invoke-Sqlcmd -ServerInstance $sqlServer -Database $db -Query $query

  $ownedWorkstations = $ownedAssets | Where-Object { $_.type -eq "Workstation" }  
  $workstationCategoryID = (Get-SnipeitCategory -search "Workstation" | Where-Object { $_.name -eq "Workstation" }).id
  $fieldsetID = ((Get-SnipeitFieldset) | Where-Object { $_.name -eq "Asset with AD and iSupport Sync" }).id

  #Verify that all workstation manufacturers exist in ITAM
  $workstationManufacturers = $ownedWorkstations | Select-Object -Property 'MFG' -Unique
  foreach ($manu in $workstationManufacturers) {
    if (!((Get-SnipeitManufacturer -all).name -eq $manu) ) {
      New-SnipeitManufacturer -Name $manu
    }
  }

  #Verify that all workstation models exist in ITAM
  $workstationModels = $ownedWorkstations | Select-Object -Property 'MODEL', 'MFG' -Unique
  foreach ($model in $workstationModels) {
    if (!((Get-SnipeitModel -all).name -eq ($model.MODEL -replace '\"', '')) ) {
      New-SnipeitModel -name ($model.MODEL -replace '\"', '') -model_number ($model.MODEL -replace '\"', '') -category_id $workstationCategoryID -manufacturer_id (Get-SnipeitManufacturer -search $model.mfg).id -fieldset_id $fieldsetID 
    }
  }

  #Verify that all locations exist in ITAM
  $workstationLocations = $ownedWorkstations | Select-Object -Property 'LOCATION' -Unique
  foreach ($location in $workstationLocations) {
    if (!((Get-SnipeitLocation -all).name -eq $location.location) ) {
      New-SnipeitLocation -name $location.location
    }
  }

  $defaultStatus = Get-SnipeitStatus -search 'Ready to Deploy'
  $waitTime = 1000
  $num = 0

  #Update/add each workstation
  foreach ($ownedWorkstation in $ownedWorkstations) {
    $model = (Get-SnipeitModel -search ($ownedWorkstation.MODEL -replace '\"', '')) | Where-Object { $_.name -eq ($ownedWorkstation.MODEL -replace '\"', '') }
    #Handle Too many Requests
    while (!$model -or $model.GetType().name -eq 'HttpWebRequest') {
      Start-Sleep -Milliseconds $waitTime
      $model = (Get-SnipeitModel -search ($ownedWorkstation.MODEL -replace '\"', '')) | Where-Object { $_.name -eq ($ownedWorkstation.MODEL -replace '\"', '') }
    }
      
    $asset = Get-SnipeitAsset -asset_tag $ownedWorkstation.NAME

    #Handle Too many Requests
    while ($asset.GetType().name -eq 'HttpWebRequest') {
      Start-Sleep -Milliseconds $waitTime
      $asset = Get-SnipeitAsset -asset_tag $ownedWorkstation.NAME
      $asset
    }


    #Add the asset if it doesn't already exist
    if (!$asset) {
      $tempAsset = New-SnipeitAsset -asset_tag $ownedWorkstation.NAME -model_id $model.id -status_id $defaultStatus.id -name $ownedWorkstation.NAME
      while ($tempAsset.GetType().name -eq 'HttpWebRequest') {
        Start-Sleep -Milliseconds $waitTime
        $tempAsset = New-SnipeitAsset -asset_tag $ownedWorkstation.NAME -model_id $model.id -status_id $defaultStatus.id -name $ownedWorkstation.NAME
        $tempAsset
      }
      
    }

    $asset = Get-SnipeitAsset -asset_tag $ownedWorkstation.NAME

    #Handle Too many Requests
    while (!$asset -or $asset.GetType().name -eq 'HttpWebRequest') {
      Start-Sleep -Milliseconds $waitTime
      $asset = Get-SnipeitAsset -asset_tag $ownedWorkstation.NAME
      $asset
    }

    $setSnipeitAssetString = 'Set-SnipeitAsset -id (' + $asset.id + ')'
    $set = $false

    #Update Model
    if ($model.id -ne $asset.model.id) {

      $trySet = Set-SnipeitAsset -id $asset.id -model_id $model.id

      #Handle Too many Requests
      while (!$trySet -or $trySet.GetType().name -eq 'HttpWebRequest') {
        Start-Sleep -Milliseconds $waitTime
        $trySet = Set-SnipeitAsset -id $asset.id -model_id $model.id
      }


      $asset = Get-SnipeitAsset -asset_tag $ownedWorkstation.NAME

      #Handle Too many Requests
      while (!$asset -or $asset.GetType().name -eq 'HttpWebRequest') {
        Start-Sleep -Milliseconds $waitTime
        $asset = Get-SnipeitAsset -asset_tag $ownedWorkstation.NAME
      }

    }


    $purchaseDateSet = ""

    #Update Purchase Date and warranty if purchase date doesn't exist
    if ($null -ne $ownedWorkstation.DT_PURCHASE -and ($purchaseDate = [datetime]$ownedWorkstation.DT_PURCHASE) -and (!$asset.purchase_date -or ($asset.purchase_date.date -ne $purchaseDate.ToString("yyyy-MM-dd"))) ) {
      $purchaseDateSet = $purchaseDate.ToString("yyyy-MM-dd")

      $setSnipeitAssetString = $setSnipeitAssetString + ' -purchase_date ($purchaseDateSet)'
      $set = $true


      if ($null -ne $ownedWorkstation.DT_WAR_END -and ($warrantyDate = [datetime]$ownedWorkstation.DT_WAR_END)) {

        $warrantyMonths = ($warrantyDate.Year - $purchaseDate.Year) * 12 + ($warrantyDate.Month - $purchaseDate.Month)
        $setSnipeitAssetString = $setSnipeitAssetString + ' -warranty_months (' + $warrantyMonths + ')'
        $set = $true

      }
    }

    $customFields = @()

    if ($null -ne $ownedWorkstation.SERIAL_NUMBER -and !$asset.custom_fields.'Serial Number'.value) {

      $customFields = @{
        $asset.custom_fields.'Serial Number'.field = $ownedWorkstation.SERIAL_NUMBER
      }
      $setSnipeitAssetString = $setSnipeitAssetString + ' -customfields ($customFields)'
      $set = $true

    }

    $owner = Get-SnipeitUser -username ($ownedWorkstation.LOGIN -Split '\\')[1]
    #Handle Too many Requests
    if ($owner) {
      while ($owner.GetType().name -eq 'HttpWebRequest') {
        Start-Sleep -Milliseconds $waitTime
        $owner = Get-SnipeitUser -username ($ownedWorkstation.LOGIN -Split '\\')[1]
      }
    }

    #Update Asset with Owner
    if ($owner) {
      #Update the owner of the asset with the location
      
      ##################################################################DOES NOT HAVE TOO MANY REQUEST CHECKING
      if (($owner.location) -and ($owner.location.name -ne $ownedWorkstation.LOCATION) -and ($location = Get-SnipeitLocation -search $ownedWorkstation.LOCATION)) {

        $trySet = Set-SnipeitUser -id $owner.id -location_id $location.id
        #Handle Too many Requests
        while (!$trySet -or $trySet.GetType().name -eq 'HttpWebRequest') {
          Start-Sleep -Milliseconds $waitTime
          $trySet = Set-SnipeitUser -id $owner.id -location_id $location.id
          $trySet
        }
      }

      #assign asset to owner if the asset can be deployed and isn't already deployed to someone
      if ($asset.status_label.status_type -eq 'deployable' -and $asset.status_label.status_meta -ne 'deployed' ) {
        $trySet = Set-SnipeitAssetOwner -id $asset.id -assigned_id $owner.id
        #Handle Too many Requests
        while (!$trySet -or $trySet.GetType().name -eq 'HttpWebRequest') {
          Start-Sleep -Milliseconds $waitTime
          $trySet = Set-SnipeitAssetOwner -id $asset.id -assigned_id $owner.id
          $trySet
        }
      }
    }

    if ($set) {
      if ($asset.id) {
        $trySet = Invoke-Expression -Command $setSnipeitAssetString
      }
      else {
        "Asset does not exist for: " + $ownedWorkstation.NAME
      }
      #Handle Too many Requests
      while (!$trySet -or $trySet.GetType().name -eq 'HttpWebRequest') {
        Start-Sleep -Milliseconds $waitTime
        if ($asset.id) {
          $trySet = Invoke-Expression -Command $setSnipeitAssetString
        }
        else {
          "Asset does not exist for: " + $ownedWorkstation.NAME
        }
        $trySet
      }
    }
    $num++
    ($num.toString() + ' out of ' + $ownedWorkstations.Count)
    $ownedWorkstation.NAME
  }


}



function Reset-ITAMServiceCredential {
  param(
    #validate parameter is path that exists
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()] [credTypes]$type,
    [Parameter()][ValidateScript({
        if (-not ($_ | Test-Path)) {
          throw "File/folder does not exist"
        }
        if (-not ($_ | Test-Path -PathType Container)) {
          throw "The path argument must be a folder."
        }
        return $true
      })] [System.IO.FileInfo]$path = (Get-Location)
  )

  #Reset credential
  if ($type -eq [credTypes]::ITAM) {
    $host.ui.PromptForCredential("Store API key for ITAM", "Please enter a new API Key as password.", "http://itassets.iwunet.indwes.edu", "ITAM") | Export-Clixml -Path ($path.FullName + "\ITAM.xml")
  }
  elseif ($type -eq ([credTypes]::VMWare)) {
    Get-Credential -Message "Please provide your login for vm Ware" | Export-Clixml -Path ($path.FullName + "\VMWare.xml")
  }
  elseif ($type -eq ([credTypes]::PowerBI)) {
    Connect-PowerBIServiceAccount
  }

  Connect-ITAMServices
}


function Add-ITAMVM {

  param(
    [Parameter()] [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$vm
  )

  $onStatusID = (Get-SnipeitStatus -Search "Powered On").id
  $offStatusID = (Get-SnipeitStatus -Search "Powered Off").id

  $vmModelID = (Get-SnipeitModel -Search "Virtual Machine").id

  $vmGuest = $vm | Get-VMGuest
  $ipString = ""
  foreach ($ip in $vmGuest.IPAddress) {
    $ipString += $ip + ", "
  }

  $ipString = $ipString.TrimEnd(", ")


  $customFields = @{
    "_snipeit_number_of_cpus_2" = $vm.NumCpu
    "_snipeit_memory_gb_3"      = $vm.MemoryGB
    "_snipeit_ip_addresses_4"   = $ipString
    "_snipeit_os_5"             = $vmGuest.OSFullName
  }

  $powerStatusID
  if ($vm.PowerState -eq "PoweredOn") {
    $powerStatusID = $onStatusID
  }
  else {
    $powerStatusID = $offStatusID
  }

  #Creates the asset in ITAM with the updated data.
  New-SnipeitAsset -Name $vm.Name -model_id $vmModelID -status_id $powerStatusID -customfields $customFields -asset_tag $vm.id

  #Gets an updated reference to the asset from ITAM
  $asset = Get-SnipeitAsset -asset_tag $vm.id

  #Gets a vm's parents and if it exists, assign the asset to the computer
  $parentComputer = Get-SnipeitAsset -asset_tag $asset.name
  if ($parentComputer -and !$asset.assigned_to) {
    Set-SnipeitAssetOwner -id $asset.id -assigned_id $parentComputer.id -checkout_to_type asset
  }
}

function Add-ITAMComputer {

  param(
    [Parameter()] [Microsoft.ActiveDirectory.Management.ADComputer]$computer
  )

  $computer = $computer | Get-ADComputer -Properties CN, Created, IPv4Address, IPv6Address, LastLogonDate, Modified, OperatingSystem, OperatingSystemVersion

  $powerStatusID = (Get-SnipeitStatus -Search "Powered On").id
  $modelID = (Get-SnipeitModel -Search "Computer").id

  $customFields = @{
    "_snipeit_os_5"               = $computer.OperatingSystem
    "_snipeit_os_version_9"       = $computer.OperatingSystemVersion
    "_snipeit_modified_10"        = $computer.Modified.ToString()
    "_snipeit_created_11"         = $computer.Created.ToString()
    "_snipeit_last_logon_date_12" = $computer.LastLogonDate.ToString()
    "_snipeit_sid_13"             = $computer.SID.Value
  }

  if ($computer.IPv4Address) {
    $customFields += @{ "_snipeit_ipv4_address_6" = $computer.IPv4Address }
  }
  if ($computer.IPv6Address) {
    $customFields += @{ "_snipeit_ipv6_address_7" = $computer.IPv6Address }
  }

  New-SnipeitAsset -Name $computer.CN -model_id $modelID -status_id $powerStatusID -customfields $customFields -asset_tag $computer.CN
}

function Add-AllITAMComputers {
  $computers = Get-ADComputer -Filter '*' -Properties CN | Where-Object { ($_.DistinguishedName -notlike "*OU=Archived,*") -and ($_.DistinguishedName -notlike "*OU=Unmanaged,*") }

  foreach ($computer in $computers) {
    if (!(Get-SnipeitAsset -asset_tag $computer.CN )) {
      Add-ITAMComputer -Computer $computer
      $computer.CN
    }
  }
}


function Add-AllITAMVMs {
  $vms = Get-VM

  foreach ($vm in $vms) {
    Add-ITAMVM -VM $vm
  }
}

function Remove-AllITAMVMs {
  $vmModelID = (Get-SnipeitModel -Search "Virtual Machine").id
  Remove-SnipeitAsset -Id (Get-SnipeitAsset -All -model_id $vmModelID).id
}

function Remove-AllITAMComputers {
  $modelID = (Get-SnipeitModel -Search "Computer").id
  Remove-SnipeitAsset -Id (Get-SnipeitAsset -All -model_id $modelID).id
}

function Update-ITAMVM {
  param(
    [Parameter()] [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$vm,
    [Parameter()] $assetTag
  )

  $asset = Get-SnipeitAsset -asset_tag $assetTag

  #if($vm.Name -ne $asset.name){
  #Set-SnipeitAsset -id $assetID -name $vm.Name
  #}

  $onStatusID = (Get-SnipeitStatus -Search "Powered On").id
  $offStatusID = (Get-SnipeitStatus -Search "Powered Off").id
  $vmModelID = (Get-SnipeitModel -Search "Virtual Machine").id

  $vmGuest = Get-VMGuest -VM $vm
  $ipString = ""
  foreach ($ip in $vmGuest.IPAddress) {
    $ipString += $ip + ", "
  }

  $ipString = $ipString.TrimEnd(", ")


  $customFields = @{
    "_snipeit_number_of_cpus_2" = $vm.NumCpu
    "_snipeit_memory_gb_3"      = $vm.MemoryGB
    "_snipeit_ip_addresses_4"   = $ipString
    "_snipeit_os_5"             = $vmGuest.OSFullName
  }

  $powerStatusID
  if ($vm.PowerState -eq "PoweredOn") {
    $powerStatusID = $onStatusID
  }
  else {
    $powerStatusID = $offStatusID
  }

  Set-SnipeitAsset -id $asset.id -status_id $powerStatusID -customfields $customFields

  #Gets te vm's parents and if it exists, assign the asset to the computer
  $parentComputer = Get-SnipeitAsset -asset_tag $asset.name
  if ($parentComputer -and !$asset.assigned_to) {
    Set-SnipeitAssetOwner -id $asset.id -assigned_id $parentComputer.id -checkout_to_type asset
  }
  


}

function Archive-ITAMAsset {
  param(
    [Parameter()] $assetTag
  )
  $archivedStatus = 3

  #Archives an asset in Snipe based on the asset tag passed in.
  $asset = Get-SnipeitAsset -asset_tag $assetTag
  Set-SnipeitAsset -id $asset.id -status_id $archivedStatus -archived $true

}

function Archive-ITAMVMs {
  $vms = Get-VM
  $itamVMs = Get-SnipeitAsset -model_id (Get-SnipeitModel -Search "Virtual Machine").id -All


  foreach ($itamVM in $itamVMs) {
    if ($vms.id -notcontains $itamVM.asset_tag) {
      Archive-ITAMAsset -assetTag $itamVM.asset_tag
    }
  }
}

function Archive-ITAMComputers {
  #gets the up to date computer data from Active Directory
  $computers = Get-ADComputer -Filter * -Properties CN

  $itamComputers = Get-SnipeitAsset -model_id (Get-SnipeitModel -Search "Computer").id -All

  foreach ($computer in $itamComputers) {
    #If the computer does not exist in the Active Directory Dataset, archive it on Snipe
    if ($computers.CN -notcontains $computer.asset_tag) {
      Archive-ITAMAsset -assetTag $computer.asset_tag
    }
  }
}

function Update-ITAMComputer {
  param(
    [Parameter()] $assetTag
  )

  #Gets the computer asset in Snipe
  $asset = Get-SnipeitAsset -asset_tag $assetTag
  #Gets the corrosponding computer data from Active Directory
  $computer = Get-ADComputer -Filter 'CN -like $assetTag' -Properties CN, Created, IPv4Address, IPv6Address, LastLogonDate, Modified, OperatingSystem, OperatingSystemVersion
  $lastModifiedITAM = $null

  if ($asset.custom_fields.Modified.Value) {
    $lastModifiedITAM = [datetime]$asset.custom_fields.Modified.Value
  }

  $status = Get-SnipeitStatus -Id 2

  $lastModifiedAD = [datetime]$computer.Modified

  #If the asset has been modified since the last update, update all of the data fields
  if (!$lastModifiedITAM -or ($lastModifiedITAM -lt $lastModifiedAD)) {
    $customFields = @{
      "_snipeit_os_5"               = $computer.OperatingSystem
      "_snipeit_os_version_9"       = $computer.OperatingSystemVersion
      "_snipeit_modified_10"        = $computer.Modified.ToString()
      "_snipeit_created_11"         = $computer.Created.ToString()
      "_snipeit_last_logon_date_12" = $computer.LastLogonDate.ToString()
    }

    if ($computer.IPv4Address) {
      $customFields += @{ "_snipeit_ipv4_address_6" = $computer.IPv4Address }
    }
    if ($computer.IPv6Address) {
      $customFields += @{ "_snipeit_ipv6_address_7" = $computer.IPv6Address }
    }

    #Push the updates to Snipe
    Set-SnipeitAsset -Id $asset.id -customfields $customFields -status_id $status.id

  }

}

function Update-AllITAMVMs {

  $vms = Get-VM

  foreach ($vm in $vms) {

    $asset = Get-SnipeitAsset -asset_tag $vm.id
    if ($asset) {
      Update-ITAMVM -VM $vm -assetTag $vm.id
    }
    else {
      Add-ITAMVM $vm
    }

  }

  #Archives any vms on Snipe that no longer exist in VMWare
  Archive-ITAMVMs
}

function Get-ITAMAsset {
  param(
    [Parameter(Mandatory = $true)] [string]$assetTag,
    [string]$properties,
    [switch]$excludeEmptyProperties
  )

  $snipeAsset = Get-SnipeitAsset -asset_tag $assetTag

  #If the snipe asset is null (passed asset tag is not real), throw exception
  if (!($snipeAsset)) {
    throw "asset Does not exist in Snipe"
  }

  $itamAsset = [pscustomobject]@{}

  #If the parameter properties is null or *, return object with all properties
  if ([string]::IsNullOrWhiteSpace($properties) -or ($properties -eq "*")) {
    foreach ($property in $snipeAsset.PSObject.properties) {
      $itamAsset | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
    }
  }
  else {
    foreach ($property in ($properties.Split(", "))) {
      #If a property provided does not exist, an exception will be thrown
      if ($property) {
        if (!($snipeAsset.PSObject.properties.Name -eq $property)) {
          throw "property provided does not exist: $property"
        }

        $itamAsset | Add-Member -MemberType NoteProperty -Name $property -Value $snipeAsset.$property
      }
    }
  }

  #Expand the custom fields property to include custom field value pairs in the object
  if ($itamAsset.custom_fields) {
    foreach ($property in $itamAsset.custom_fields.PSObject.properties) {
      $itamAsset | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value.Value
    }
  }

  #If the user chooses to remove the properties that don't have value, iterate and remove empty valued pairs.
  if ($excludeEmptyProperties) {
    foreach ($property in $itamAsset.PSObject.properties) {
      if ([string]::IsNullOrWhiteSpace($property.Value)) {
        $itamAsset.PSObject.properties.Remove($property.Name)
      }
    }
  }

  return $itamAsset
}

function Update-AllITAMComputers {
  $computers = Get-ADComputer -Filter * -Properties CN

  foreach ($computer in $computers) {
    if (Get-SnipeitAsset -asset_tag $computer.CN) {

      Update-ITAMComputer -assetTag $computer.CN
    }
    else {
      Add-ITAMComputer -Computer $computer
    }
  }

  Archive-ITAMComputer

}


#Prints a pie chart of the Assets in Snipe by Model
function Out-ITAMAssetsbyModel {
  param(
    #sets default value for num parameter if the parameter was not provided.
    [Parameter()][ValidateRange(1, 10)] [int]$num = 4,
    [Parameter()][bool]$testPrint = $false
  )

  $assets = Get-SnipeitAsset -All
  $assetModels = $assets.Model | Group-Object -Property name | Sort-Object -Descending Count
  $printData = $assetModels | Select-Object -First $num Name, Count
  $count = $null

  foreach ($assetModel in $assetModels) {
    if (!(($printData.Name).Contains($assetModel.Name))) {
      $count += $assetModel.Count
    }
  }

  if ($count) {
    $printData += [pscustomobject]@{
      Name  = 'Other'
      Count = $count
    }
  }
  
  if (!$testPrint) { 
    $printData | Out-PieChart -PieChartTitle "Assets by Model" -DisplayToScreen
  }

}

#Prints a report to PowerBI
function Out-ITAMPowerBIReport {
    
  param(
    [Parameter(Mandatory = $true)] [string]$name,
    [Parameter(Mandatory = $true)] [string]$search

  )

  $workspace = Get-PowerBIWorkspace -Name 'IWU ITAM'

  # PowerBI Data Types: Int64, Double, Boolean, DateTime, String
  $columnMap = @{
    'name'       = 'String'
    'asset_tag'  = 'String'
    'model'      = 'String'
    'status'     = 'String'
    'category'   = 'String'
    'created_at' = 'DateTime'
    'updated_at' = 'DateTime'
  }

  $columns = @()
     
  #Create an array of PowerBI Column Objects
  $columnMap.GetEnumerator() | ForEach-Object {
    $columns += New-PowerBIColumn -Name $_.Key -DataType $_.Value
  }

  $table = New-PowerBITable -Name "DefaultTable" -Columns $columns
  $dataSet = New-PowerBIDataSet -Name $name -Tables $table
  $dataSetResult = Add-PowerBIDataSet -DataSet $dataSet -WorkspaceId $workspace.Id

  $assets = @()
      
  #Get the relevant data fields from Snipe
    
  $assets = Get-SnipeitAsset -search $search

  if (!$assets) {
    throw "Assets do not exist"
  }

  foreach ($asset in $assets) {
    $Row =
    @{
      'name'       = $asset.name
      'asset_tag'  = $asset.asset_tag
      'model'      = $asset.model.name
      'status'     = $asset.status_label.name
      'category'   = $asset.category.name
      'created_at' = $asset.created_at.datetime
      'updated_at' = $asset.updated_at.datetime
    }
   
    Add-PowerBIRow -DatasetId $dataSetResult.Id -TableName $table.Name -Row $Row -WorkspaceId $workspace.Id
  }
}

#Function for printing Pie Charts is from: https://4sysops.com/archives/convert-powershell-csv-output-into-a-pie-chart/
function Out-PieChart {
  [CmdletBinding()]
  param(
    [Parameter(ValueFromPipeline)]
    [psobject]$inputObject,
    [Parameter()]
    [string]$PieChartTitle,
    [Parameter()]
    [int]$ChartWidth = 800,
    [Parameter()]
    [int]$ChartHeight = 400,
    [Parameter()]
    [string[]]$NameProperty,
    [Parameter()]
    [string]$ValueProperty,
    [Parameter()]
    [switch]$Pie3D,
    [Parameter()]
    [switch]$DisplayToScreen,
    [Parameter()]
    [string]$saveImage
  )
  begin {
    Add-Type -AssemblyName System.Windows.Forms.DataVisualization
    # Frame
    $Chart = [System.Windows.Forms.DataVisualization.Charting.Chart]@{
      Width       = $ChartWidth
      Height      = $ChartHeight
      BackColor   = 'White'
      BorderColor = 'Black'
    }
    # Body
    $null = $Chart.Titles.Add($PieChartTitle)
    $Chart.Titles[0].Font = "segoeuilight,20pt"
    $Chart.Titles[0].Alignment = "TopCenter"
    # Create Chart Area
    $ChartArea = [System.Windows.Forms.DataVisualization.Charting.ChartArea]::new()
    $ChartArea.Area3DStyle.Enable3D = $Pie3D.ToBool()
    $ChartArea.Area3DStyle.Inclination = 50
    $Chart.ChartAreas.Add($ChartArea)
    # Define Chart Area
    $null = $Chart.Series.Add("Data")
    $Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Pie
    # Chart style
    $Chart.Series["Data"]["PieLabelStyle"] = "Outside"
    $Chart.Series["Data"]["PieLineColor"] = "Black"
    $Chart.Series["Data"]["PieDrawingStyle"] = "Concave"

    $chart.Series["Data"].IsValueShownAsLabel = $true
    $chart.Series["Data"].Label = "#PERCENT\n#VALX"
    # Set ArrayList
    $XColumn = [System.Collections.ArrayList]::new()
    $yColumn = [System.Collections.ArrayList]::new()
  }
  process {
    if (-not $valueProperty) {
      $numericProperties = foreach ($property in $inputObject.PSObject.properties) {
        if ([double]::TryParse($property.Value, [ref]$null)) {
          $property.Name
        }
      }
      if (@($numericProperties).Count -eq 1) {
        $valueProperty = $numericProperties
      }
      else {
        throw 'Unable to automatically determine properties to graph'
      }
    }
    if (-not $LabelProperty) {
      if ($inputObject.PSObject.properties.Count -eq 2) {
        $LabelProperty = $inputObject.properties.Name -ne $valueProperty
      }
      elseif ($inputObject.PSObject.properties.Item('Name')) {
        $LabelProperty = 'Name'
      }
      else {
        throw 'Cannot convert Data'
      }
    }
    # Bind chart columns
    $null = $yColumn.Add($InputObject.$valueProperty)
    $null = $xColumn.Add($inputObject.$LabelProperty)
  }
  end {
    # Add data to chart
    $Chart.Series["Data"].Points.DataBindXY($xColumn, $yColumn)
    # Save file
    if ($psboundparameters.ContainsKey('saveImage')) {
      try {
        if (Test-Path (Split-Path $saveImage -Parent)) {
          $SaveImage = $pscmdlet.GetUnresolvedProviderPathFromPSPath($saveImage)
          $Chart.SaveImage($saveImage, "png")
        }
        else {
          throw 'Invalid path, the parent directory must exist'
        }
      }
      catch {
        throw
      }
    }
    # Display Chart to screen
    if ($DisplayToScreen.ToBool()) {
      $Form = [Windows.Forms.Form]@{
        Width           = 800
        Height          = 450
        AutoSize        = $true
        FormBorderStyle = "FixedDialog"
        MaximizeBox     = $false
        MinimizeBox     = $false
        KeyPreview      = $true
      }
      $Form.controls.Add($Chart)
      $Chart.Anchor = 'Bottom, Right, Top, Left'
      $Form.Add_KeyDown({
          if ($_.KeyCode -eq "Escape") { $Form.Close() }
        })
      $Form.Add_Shown({ $Form.Activate() })
      $Form.ShowDialog() | Out-Null
    }
  }
}

function Get-ITAMData {
  param(
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()] [outputTypes]$type
  )
  $command = "Get-Snipeit$type -all"

  $data = Invoke-Expression $command
  return $data
}

function Out-ITAMReport {
  param(
    [Parameter(Mandatory = $true)][ValidateScript({
        if ($_ | Test-Path) {
          throw "File/Folder alreadys exists"
        }
        if (-not ($_ | Split-Path | Test-Path)) {
          throw "Folder path to new file does not exist."
        }
        if ($_ -notmatch "(\.csv)") {
          throw "The file specified in the path argument must be a .csv"
        }
        return $true
      })] [System.IO.FileInfo]$path,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()] [outputTypes]$type
  )
  Get-ITAMData -Type $type | Export-Csv -Path $path

}

function Out-AllITAMReports {

  param(
    [Parameter(Mandatory = $true)][ValidateScript({
        if (-not ($_ | Test-Path)) {
          throw "File/folder does not exist"
        }
        if (-not ($_ | Test-Path -PathType Container)) {
          throw "The path argument must be a folder."
        }
        return $true
      })] [System.IO.FileInfo]$path
  )

  $date = (Get-Date -Format yyyy-MM-dd_HH-mm-ss)
  $folderName = "InventoryReport - $date"
  New-Item -Path $path -ItemType "directory" -Name $folderName

  $path = "$path\$folderName"
  foreach ($type in [outputTypes].GetEnumNames()) {
    $newPath = "$path\$type.csv"
    $newPath
    Out-ITAMReport -Path $newPath -Type $type
  }
}

#Export all functions that can be executed by a user
Export-ModuleMember -Function 'Add-ITAMVM'
Export-ModuleMember -Function 'Add-AllITAMVMs'
Export-ModuleMember -Function 'Remove-AllITAMVMs'
Export-ModuleMember -Function 'Update-ITAMVM'
Export-ModuleMember -Function 'Update-AllITAMVMs'
Export-ModuleMember -Function 'Archive-ITAMAsset'
Export-ModuleMember -Function 'Out-ITAMPowerBIReport'


Export-ModuleMember -Function 'Add-ITAMComputer'
Export-ModuleMember -Function 'Add-AllITAMComputers'
Export-ModuleMember -Function 'Remove-AllITAMComputers'
Export-ModuleMember -Function 'Update-ITAMComputer'
Export-ModuleMember -Function 'Update-AllITAMComputers'

Export-ModuleMember -Function 'Out-ITAMAssetsbyModel'
Export-ModuleMember -Function 'Connect-ITAMServices'
Export-ModuleMember -Function 'Reset-ITAMServiceCredential'

Export-ModuleMember -Function 'Out-AllITAMReports'
Export-ModuleMember -Function 'Out-ITAMReport'
Export-ModuleMember -Function 'Get-ITAMData'
Export-ModuleMember -Function 'Get-ITAMAsset'
Export-ModuleMember -Function 'Sync-ITAMfromiSupport'

