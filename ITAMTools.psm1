Import-Module VMware.PowerCLI
Import-Module SnipeitPS
Import-Module MicrosoftPowerBIMgmt

$SnipeURL = "http://itassets.iwunet.indwes.edu"
$SnipeApiKey = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiYTIyNDM2ZjM1Y2M0NWJiZTE4YjI2NzU5OGZkZDM0NGU4YThiNjJkYzc2OTFmOWJiMGRiNmFmNDlhMWM1NDYyYzdkODllNmQwOGU4NTk4NmIiLCJpYXQiOjE2MzkwNjU4NjYsIm5iZiI6MTYzOTA2NTg2NiwiZXhwIjoyMTEyNDUxNDY2LCJzdWIiOiI2Iiwic2NvcGVzIjpbXX0.fbV87YVigpVrbeGtfXGFU17dkeRQ1s5ADyZuAdmMtWp5uX_pAWSDB3-Sbzc6IQMxz4uu-ssbNzInmAt-iDuinVjAoWMcJroEwCzQEzZuADNFfEOARsUAKIMOmL9FMNNZbBsg7xQwMSq_EAKnOFYLAYXBE-STN31oCSjh7d1mfs0r6GRINcc2djSDP6GLeSVaQEWFVivd_ki9pQ3_Gp-3_Shd7XooHDF6zZziNUGgoJvvxVtPK6OFsLRnP2PYsuw3xNHVwlMuxDgwmvqU_MOwzW0WWMXG4iI7-qZOGAOAaDWIhRr7hwQNR_hKqYpCN3K9JxwbkFAdV9eJKPP2ok5aF0AL84vz2wnp3y4CaSt0fidpL_K6uCOabu0_sboLQztNS94DH2vo1uEQ3Ej1hByUYA2qKV4GSrSPBRZWAs32mU9c-NYagDVxSXjiMeLv18CqcowAB25jVeLEtPCNxiqt2RXtRAZC7ydlmEKg8F4YXA2gCsDtN_vYVtXWU1nga_mhmy8nYhy_ONj5BCTkdkjWfope5KFtHvCIKFWtQiSGh00p_iYTsNUIU_5A17CtnE09Y_yAlsYdxDFpB71tec_Sa0poQnymt4M8-01GjIkowLSp_Sx8De2iGYFP4pfPrpHaPyZ1QPZ7SWfZ4VY-LuXQ6HUDFfyK8AabNeRPHcNJAIM"

$VIServer = "marn-vb-vmsa01.iwunet.indwes.edu"

function Connect-ITAM {
  #New-VICredentialStoreItem -Host marn-vb-vmsa01.iwunet.indwes.edu -User b.batman-stwk@indwes.edu -Password *
  Connect-SnipeitPS -url $SnipeURL -apiKey $SnipeApiKey
  #Connect-VIServer -Server $VIServer
  #Connect-PowerBIServiceAccount

}

Connect-ITAM


enum outputTypes{
  Accessory
  Activity
  Asset
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

function Add-ITAMVM {

  param(
    [Parameter()] [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$VM
  )

  #Connect-ITAM


  $onStatusID = (Get-SnipeitStatus -Search "Powered On").id
  $offStatusID = (Get-SnipeitStatus -Search "Powered Off").id
  $VMModelID = (Get-SnipeitModel -Search "Virtual Machine").id

  $VMGuest = $VM | Get-VMGuest
  $ipString = ""
  foreach ($IP in $VMGuest.IPAddress) {
    $ipString += $IP + ", "
  }

  $ipString = $ipString.TrimEnd(", ")


  $customFields = @{
    "_snipeit_number_of_cpus_2" = $VM.NumCpu
    "_snipeit_memory_gb_3"      = $VM.MemoryGB
    "_snipeit_ip_addresses_4"   = $ipString
    "_snipeit_os_5"             = $VMGuest.OSFullName
  }

  $powerStatusID
  if ($VM.PowerState -eq "PoweredOn") {
    $powerStatusID = $onStatusID
  }
  else {
    $powerStatusID = $offStatusID
  }

  New-SnipeitAsset -Name $VM.Name -model_id $VMModelID -status_id $powerStatusID -customfields $customFields -asset_tag $VM.id

  $Asset = Get-SnipeitAsset -asset_tag $VM.id
  $ParentComputer = Get-SnipeitAsset -asset_tag $Asset.name
    if($ParentComputer -and !$Asset.assigned_to){
      Set-SnipeitAssetOwner -id $Asset.id -assigned_id $ParentComputer.id -checkout_to_type asset
    }
}

function Add-ITAMComputer {

  param(
    [Parameter()] [Microsoft.ActiveDirectory.Management.ADComputer]$Computer
  )

  #Connect-ITAM


  $Computer = $Computer | Get-ADComputer -Properties CN, Created, IPv4Address, IPv6Address, LastLogonDate, Modified, OperatingSystem, OperatingSystemVersion

  $powerStatusID = (Get-SnipeitStatus -Search "Powered On").id
  $ModelID = (Get-SnipeitModel -Search "Computer").id

  $customFields = @{
    "_snipeit_os_5"               = $Computer.OperatingSystem
    "_snipeit_os_version_9"       = $Computer.OperatingSystemVersion
    "_snipeit_modified_10"        = $Computer.Modified.ToString()
    "_snipeit_created_11"         = $Computer.Created.ToString()
    "_snipeit_last_logon_date_12" = $Computer.LastLogonDate.ToString()
    "_snipeit_sid_13"             = $Computer.SID.Value
  }

  if ($Computer.IPv4Address) {
    $customFields += @{ "_snipeit_ipv4_address_6" = $Computer.IPv4Address }
  }
  if ($Computer.IPv6Address) {
    $customFields += @{ "_snipeit_ipv6_address_7" = $Computer.IPv6Address }
  }

  New-SnipeitAsset -Name $Computer.CN -model_id $ModelID -status_id $powerStatusID -customfields $customFields -asset_tag $Computer.CN
}

function Add-AllITAMComputers {
  #Connect-ITAM


  $Computers = Get-ADComputer -Filter '*' | Where-Object { ($_.DistinguishedName -notlike "*OU=Archived,*") -and ($_.DistinguishedName -notlike "*OU=Unmanaged,*") }

  foreach ($Computer in $Computers) {
    Add-ITAMComputer -Computer $Computer
  }
}


function Add-AllITAMVMs {
  Connect-ITAM


  $VMs = Get-VM

  foreach ($VM in $VMs) {
    Add-ITAMVM -VM $VM
  }
}

function Remove-AllITAMVMs {
  #Connect-ITAM

  $VMModelID = (Get-SnipeitModel -Search "Virtual Machine").id
  Remove-SnipeitAsset -Id (Get-SnipeitAsset -All -model_id $VMModelID).id
}

function Remove-AllITAMComputers {
  #Connect-ITAM

  $ModelID = (Get-SnipeitModel -Search "Computer").id
  Remove-SnipeitAsset -Id (Get-SnipeitAsset -All -model_id $ModelID).id
}

function Update-ITAMVM {
  param(
    [Parameter()] [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$VM,
    [Parameter()] $assetTag
  )

    $Asset = Get-SnipeitAsset -asset_tag $assetTag

    #if($VM.Name -ne $Asset.name){
    #Set-SnipeitAsset -id $assetID -name $VM.Name
    #}

    $onStatusID = (Get-SnipeitStatus -Search "Powered On").id
    $offStatusID = (Get-SnipeitStatus -Search "Powered Off").id
    $VMModelID = (Get-SnipeitModel -Search "Virtual Machine").id

    $VMGuest = Get-VMGuest -VM $VM
    $ipString = ""
    foreach ($IP in $VMGuest.IPAddress) {
      $ipString += $IP + ", "
    }

    $ipString = $ipString.TrimEnd(", ")


    $customFields = @{
      "_snipeit_number_of_cpus_2" = $VM.NumCpu
      "_snipeit_memory_gb_3"      = $VM.MemoryGB
      "_snipeit_ip_addresses_4"   = $ipString
      "_snipeit_os_5"             = $VMGuest.OSFullName
    }

    $powerStatusID
    if ($VM.PowerState -eq "PoweredOn") {
      $powerStatusID = $onStatusID
    }
    else {
      $powerStatusID = $offStatusID
    }

    Set-SnipeitAsset -id $Asset.id -status_id $powerStatusID -customfields $customFields

    
    $ParentComputer = Get-SnipeitAsset -asset_tag $Asset.name
    if($ParentComputer -and !$Asset.assigned_to){
      Set-SnipeitAssetOwner -id $Asset.id -assigned_id $ParentComputer.id -checkout_to_type asset
    }
  

  #Archives any VMs on Snipe that no longer exist in VMWare
  Archive-ITAMVMs
}

function Archive-ITAMAsset {
  param(
    [Parameter()] $assetTag
  )
  $archivedStatus = 3

  #Archives an asset in Snipe based on the asset tag passed in.
  $Asset = Get-SnipeitAsset -asset_tag $assetTag
  Set-SnipeitAsset -id $Asset.id -status_id $archivedStatus -archived $true

}

function Archive-ITAMVMs {
  $VMs = Get-VM
  $SnipeVMs = Get-SnipeitAsset -model_id (Get-SnipeitModel -Search "Virtual Machine").id -All


  foreach ($SnipeVM in $SnipeVMs) {
    if ($VMs.id -notcontains $SnipeVM.asset_tag) {
      Archive-ITAMAsset -assetTag $SnipeVm.asset_tag
    }
  }
}

function Archive-ITAMComputers {
  #gets the up to date computer data from Active Directory
  $Computers = Get-ADComputer -Filter * -Properties CN

  $SnipeComputers = Get-SnipeitAsset -model_id (Get-SnipeitModel -Search "Computer").id -All

  foreach ($Computer in $SnipeComputers) {
    #If the computer does not exist in the Active Directory Dataset, archive it on Snipe
    if ($Computers.CN -notcontains $Computer.asset_tag) {
      Archive-ITAMAsset -assetTag $Computer.asset_tag
    }
  }
}

function Update-ITAMComputer {
  param(
    [Parameter()] $assetTag
  )

  #Connect-ITAM


  try {
    #Gets the Computer asset in Snipe
    $Asset = Get-SnipeitAsset -asset_tag $assetTag
    #Gets the corrosponding Computer data from Active Directory
    $Computer = Get-ADComputer -Filter 'CN -like $assetTag' -Properties CN, Created, IPv4Address, IPv6Address, LastLogonDate, Modified, OperatingSystem, OperatingSystemVersion
    $LastModifiedSnipe = [datetime]$Asset.custom_fields.Modified.Value

    $statusID = Get-SnipeitStatus -Id 2

    $LastModifiedAD = [datetime]$Computer.Modified

    #If the Asset has been modified since the last update, update all of the data fields
    if ($LastModifiedSnipe -lt $LastModifiedAD) {
      $customFields = @{
        "_snipeit_os_5"               = $Computer.OperatingSystem
        "_snipeit_os_version_9"       = $Computer.OperatingSystemVersion
        "_snipeit_modified_10"        = $Computer.Modified.ToString()
        "_snipeit_created_11"         = $Computer.Created.ToString()
        "_snipeit_last_logon_date_12" = $Computer.LastLogonDate.ToString()
      }

      if ($Computer.IPv4Address) {
        $customFields += @{ "_snipeit_ipv4_address_6" = $Computer.IPv4Address }
      }
      if ($Computer.IPv6Address) {
        $customFields += @{ "_snipeit_ipv6_address_7" = $Computer.IPv6Address }
      }

      #Push the updates to Snipe
      Set-SnipeitAsset -Id $Asset.id -customfields $customFields -status_id $statusID

    }

  }
  catch {
    throw "Asset does not exist"
  }


}

function Update-AllITAMVMs {

  $VMs = Get-VM

  foreach ($VM in $VMs) {

    $Asset = Get-SnipeitAsset -asset_tag $VM.id
    if ($Asset) {
      Update-ITAMVM -VM $VM -assetTag $VM.id
    }
    else {
      Add-ITAMVM $VM
    }

  }
}

function Get-ITAMAsset {
  param(
    [Parameter(Mandatory = $true)] [string]$assetTag,
    [string]$Properties,
    [switch]$excludeEmptyProperties
  )

  $SnipeAsset = Get-SnipeitAsset -asset_tag $assetTag

  #If the snipe asset is null (passed asset tag is not real), throw exception
  if (!($SnipeAsset)) {
    throw "Asset Does not exist in Snipe"
  }

  $IWUITAsset = [pscustomobject]@{}

  #If the parameter Properties is null or *, return object with all Properties
  if ([string]::IsNullOrWhiteSpace($Properties) -or ($Properties -eq "*")) {
    foreach ($Property in $SnipeAsset.PSObject.Properties) {
      $IWUITAsset | Add-Member -MemberType NoteProperty -Name $Property.Name -Value $Property.Value
    }
  }
  else {
    foreach ($Property in ($Properties.Split(", "))) {
      #If a property provided does not exist, an exception will be thrown
      if ($Property) {
        if (!($SnipeAsset.PSObject.Properties.Name -eq $Property)) {
          throw "Property provided does not exist: $Property"
        }

        $IWUITAsset | Add-Member -MemberType NoteProperty -Name $Property -Value $SnipeAsset.$Property
      }
    }
  }

  #Expand the custom fields property to include custom field value pairs in the object
  if ($IWUITAsset.custom_fields) {
    foreach ($Property in $IWUITAsset.custom_fields.PSObject.Properties) {
      $IWUITAsset | Add-Member -MemberType NoteProperty -Name $Property.Name -Value $Property.Value.Value
    }
  }

  #If the user chooses to remove the properties that don't have value, iterate and remove empty valued pairs.
  if ($excludeEmptyProperties) {
    foreach ($Property in $IWUITAsset.PSObject.Properties) {
      if ([string]::IsNullOrWhiteSpace($Property.Value)) {
        $IWUITAsset.PSObject.Properties.Remove($Property.Name)
      }
    }
  }

  return $IWUITAsset
}

function Update-AllITAMComputers {
  $Computers = Get-ADComputer -Filter * -Properties CN

  foreach ($Computer in $Computers) {
    if (Get-SnipeitAsset -asset_tag $Computer.CN) {
      try {
        Update-ITAMComputer -assetTag $Computer.CN
      }
      catch {}
    }
    else {
      Add-ITAMComputer -Computer $Computer
    }
  }

}


#Prints a pie chart of the Assets in Snipe by Model
function Out-ITAMAssetsbyModel {
  param(
    #sets default value for num parameter if the parameter was not provided.
    [Parameter()][ValidateRange(1, 10)] [int]$num = 4,
    [Parameter()][bool]$testPrint = $false
  )

  Connect-ITAM


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
  
  if(!$testPrint){ 
      $printData | Out-PieChart -PieChartTitle "Assets by Model" -DisplayToScreen
  }

}

#Prints a report to PowerBI
function Out-ITAMPowerBIReport {
    
  param(
    [Parameter(Mandatory = $true)] [string]$name,
    [Parameter(Mandatory = $true)] [string]$search

  )

  $Workspace = Get-PowerBIWorkspace -Name 'IWU ITAM'

  # PowerBI Data Types: Int64, Double, Boolean, DateTime, String
  $ColumnMap = @{
    'name'       = 'String'
    'asset_tag'  = 'String'
    'model'      = 'String'
    'status'     = 'String'
    'category'   = 'String'
    'created_at' = 'DateTime'
    'updated_at' = 'DateTime'
  }

  $Columns = @()
     
  #Create an array of PowerBI Column Objects
  $ColumnMap.GetEnumerator() | ForEach-Object {
    $Columns += New-PowerBIColumn -Name $_.Key -DataType $_.Value
  }

  $Table = New-PowerBITable -Name "DefaultTable" -Columns $Columns
  $DataSet = New-PowerBIDataSet -Name $name -Tables $Table
  $DataSetResult = Add-PowerBIDataSet -DataSet $DataSet -WorkspaceId $Workspace.Id

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
   
    Add-PowerBIRow -DatasetId $DataSetResult.Id -TableName $Table.Name -Row $Row -WorkspaceId $Workspace.Id
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
      $numericProperties = foreach ($property in $inputObject.PSObject.Properties) {
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
      if ($inputObject.PSObject.Properties.Count -eq 2) {
        $LabelProperty = $inputObject.Properties.Name -ne $valueProperty
      }
      elseif ($inputObject.PSObject.Properties.Item('Name')) {
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
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()] [outputTypes]$Type
  )
  $command = "Get-Snipeit$Type -all"

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
      })] [System.IO.FileInfo]$Path,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()] [outputTypes]$Type
  )
  Get-ITAMData -Type $Type | Export-Csv -Path $Path

}

function Out-AllITAMReports {

  param(
    [Parameter(Mandatory = $true)][ValidateScript({
        if (-not ($_ | Test-Path)) {
          throw "File/folder does not exist"
        }
        if (-not ($_ | Test-Path -PathType Container)) {
          throw "The Path argument must be a folder."
        }
        return $true
      })] [System.IO.FileInfo]$Path
  )

  $date = (Get-Date -Format yyyy-MM-dd_HH-mm-ss)
  $folderName = "InventoryReport - $date"
  New-Item -Path $Path -ItemType "directory" -Name $folderName

  $Path = "$Path\$folderName"
  foreach ($type in [outputTypes].GetEnumNames()) {
    $NewPath = "$Path\$type.csv"
    $NewPath
    Out-ITAMReport -Path $NewPath -Type $type
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
Export-ModuleMember -Function 'Connect-ITAM'

Export-ModuleMember -Function 'Out-AllITAMReports'
Export-ModuleMember -Function 'Out-ITAMReport'
Export-ModuleMember -Function 'Get-ITAMData'
Export-ModuleMember -Function 'Get-ITAMAsset'
