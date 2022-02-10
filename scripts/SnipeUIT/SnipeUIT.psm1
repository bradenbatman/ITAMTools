Import-module VMware.PowerCLI
Import-module SnipeitPS

$SnipeURL = "http:\\localhost"
$SnipeApiKey = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiYTIyNDM2ZjM1Y2M0NWJiZTE4YjI2NzU5OGZkZDM0NGU4YThiNjJkYzc2OTFmOWJiMGRiNmFmNDlhMWM1NDYyYzdkODllNmQwOGU4NTk4NmIiLCJpYXQiOjE2MzkwNjU4NjYsIm5iZiI6MTYzOTA2NTg2NiwiZXhwIjoyMTEyNDUxNDY2LCJzdWIiOiI2Iiwic2NvcGVzIjpbXX0.fbV87YVigpVrbeGtfXGFU17dkeRQ1s5ADyZuAdmMtWp5uX_pAWSDB3-Sbzc6IQMxz4uu-ssbNzInmAt-iDuinVjAoWMcJroEwCzQEzZuADNFfEOARsUAKIMOmL9FMNNZbBsg7xQwMSq_EAKnOFYLAYXBE-STN31oCSjh7d1mfs0r6GRINcc2djSDP6GLeSVaQEWFVivd_ki9pQ3_Gp-3_Shd7XooHDF6zZziNUGgoJvvxVtPK6OFsLRnP2PYsuw3xNHVwlMuxDgwmvqU_MOwzW0WWMXG4iI7-qZOGAOAaDWIhRr7hwQNR_hKqYpCN3K9JxwbkFAdV9eJKPP2ok5aF0AL84vz2wnp3y4CaSt0fidpL_K6uCOabu0_sboLQztNS94DH2vo1uEQ3Ej1hByUYA2qKV4GSrSPBRZWAs32mU9c-NYagDVxSXjiMeLv18CqcowAB25jVeLEtPCNxiqt2RXtRAZC7ydlmEKg8F4YXA2gCsDtN_vYVtXWU1nga_mhmy8nYhy_ONj5BCTkdkjWfope5KFtHvCIKFWtQiSGh00p_iYTsNUIU_5A17CtnE09Y_yAlsYdxDFpB71tec_Sa0poQnymt4M8-01GjIkowLSp_Sx8De2iGYFP4pfPrpHaPyZ1QPZ7SWfZ4VY-LuXQ6HUDFfyK8AabNeRPHcNJAIM"


$VIServer = "marn-vb-vmsa01.iwunet.indwes.edu"

function Connect-SnipeUIT{
#New-VICredentialStoreItem -Host marn-vb-vmsa01.iwunet.indwes.edu -User b.batman-stwk@indwes.edu -Password *
Connect-SnipeitPS -url $SnipeURL -apiKey $SnipeApiKey
#Connect-VIServer -Server $VIServer

}

Connect-SnipeUIT


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

function Add-SnipeVM{

    param(
        [Parameter()] [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl] $VM
    )

    #Connect-SnipeUIT

    $onStatusID = (Get-SnipeitStatus -Search "Powered On").id
    $offStatusID = (Get-SnipeitStatus -Search "Powered Off").id
    $VMModelID = (Get-SnipeitModel -search "Virtual Machine").id

    $VMGuest = $VM | Get-VMGuest
    $ipString = ""
    foreach ($IP in $VMGuest.IPAddress){
        $ipString += $IP + ", "
    }

    $ipString =$ipString.TrimEnd(", ")


    $customFields = @{
        "_snipeit_number_of_cpus_2" = $VM.NumCpu
        "_snipeit_memory_gb_3" = $VM.MemoryGB
        "_snipeit_ip_addresses_4" = $ipString
        "_snipeit_os_5" = $VMGuest.OSFullName
    }

    $powerStatusID
    if($VM.PowerState -eq "PoweredOn"){
        $powerStatusID=$onStatusID
    }
    else{
        $powerStatusID=$offStatusID
    }

    New-SnipeitAsset -name $VM.Name -model_id $VMModelID -status_id $powerStatusID -customfields $customFields -asset_tag $VM.Id
}

function Add-SnipeComputer{

    param(
        [Parameter()][Microsoft.ActiveDirectory.Management.ADComputer] $Computer
    )

    #Connect-SnipeUIT

    $Computer = $Computer | Get-ADComputer -Properties CN, Created, IPv4Address, IPv6Address, LastLogonDate, Modified, OperatingSystem, OperatingSystemVersion

    $powerStatusID = (Get-SnipeitStatus -Search "Powered On").id
    $ModelID = (Get-SnipeitModel -search "Computer").id

    $customFields = @{
        "_snipeit_os_5" = $Computer.OperatingSystem
        "_snipeit_os_version_9" = $Computer.OperatingSystemVersion
        "_snipeit_modified_10" = $Computer.Modified.ToString()
        "_snipeit_created_11" = $Computer.Created.ToString()
        "_snipeit_last_logon_date_12" = $Computer.LastLogonDate.ToString()
        "_snipeit_sid_13" = $Computer.SID.Value
    }

    if($Computer.IPv4Address){
        $customFields += @{"_snipeit_ipv4_address_6" = $Computer.IPv4Address}
    }
    if($Computer.IPv6Address){
        $customFields += @{"_snipeit_ipv6_address_7" = $Computer.IPv6Address}
    }    

    New-SnipeitAsset -name $Computer.CN -model_id $ModelID -status_id $powerStatusID -customfields $customFields -asset_tag $Computer.CN
}

function Add-AllComputerstoSnipe{
    #Connect-SnipeUIT

    $Computers = get-adcomputer -filter '*' | where { ($_.DistinguishedName -notlike "*OU=Archived,*") -and ($_.DistinguishedName -notlike "*OU=Unmanaged,*")}

    Foreach ($Computer in $Computers){
        Add-SnipeComputer -Computer $Computer
    }
}


function Add-AllVMstoSnipe{
    Connect-SnipeUIT

    $VMs = Get-VM

    Foreach ($VM in $VMs){
        Add-SnipeVM -VM $VM
    }
}

function Remove-AllSnipeVMs{
    #Connect-SnipeUIT
    $VMModelID = (Get-SnipeitModel -search "Virtual Machine").id
    Remove-SnipeitAsset -id (Get-SnipeitAsset -all -model_id $VMModelID).id
}

function Remove-AllSnipeComputers{
    #Connect-SnipeUIT
    $ModelID = (Get-SnipeitModel -search "Computer").id
    Remove-SnipeitAsset -id (Get-SnipeitAsset -all -model_id $ModelID).id
}

function Update-SnipeVM{
    param(
        [Parameter()] [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl] $VM,
        [Parameter()] $assetTag
    )

    Connect-SnipeUIT

    try {
        $Asset = Get-SnipeitAsset -asset_tag $assetTag
        $id = $Asset.id

        #if($VM.Name -ne $Asset.name){
        #Set-SnipeitAsset -id $assetID -name $VM.Name
        #}

        $onStatusID = (Get-SnipeitStatus -Search "Powered On").id
    $offStatusID = (Get-SnipeitStatus -Search "Powered Off").id
    $VMModelID = (Get-SnipeitModel -search "Virtual Machine").id

    $VMGuest = $VM | Get-VMGuest
    $ipString = ""
    foreach ($IP in $VMGuest.IPAddress){
        $ipString += $IP + ", "
    }

    $ipString =$ipString.TrimEnd(", ")


    $customFields = @{
        "_snipeit_number_of_cpus_2" = $VM.NumCpu
        "_snipeit_memory_gb_3" = $VM.MemoryGB
        "_snipeit_ip_addresses_4" = $ipString
        "_snipeit_os_5" = $VMGuest.OSFullName
    }

    $powerStatusID
    if($VM.PowerState -eq "PoweredOn"){
        $powerStatusID=$onStatusID
    }
    else{
        $powerStatusID=$offStatusID
    }

    $Asset.asset_tag | Set-SnipeitAsset  -status_id $powerStatusID -customfields $customFields


    }
    catch {
        "Asset does not exist on SnipeIT"
    }
    
    #Archives any VMs on Snipe that no longer exist in VMWare
    Archive-SnipeVMs
}

function Archive-SnipeAsset{
    param(
        [Parameter()] $assetTag
    )

    Get-SnipeitAsset -asset_tag $assetTag | Set-SnipeitAsset -status_id 3 -archived $true

}

function Archive-SnipeVMs{
    $VMs = Get-VM
    $SnipeVMs = Get-SnipeitAsset -model_id (Get-SnipeitModel -search "Virtual Machine").id -all
    
    foreach($SnipeVM in $SnipeVMs){
    if($VMs.id -notcontains $SnipeVM.asset_tag ){
       Archive-SnipeAsset -assetTag  $SnipeVm.asset_tag
    }
    }
}

function Archive-SnipeComputers{
    $Computers = Get-ADComputer -Filter * -Properties CN

    $SnipeComputers = Get-SnipeitAsset -model_id (Get-SnipeitModel -search "Computer").id -all
    
    foreach($Computer in $SnipeComputers){
    if($Computers.CN -notcontains $Computer.asset_tag){
       Archive-SnipeAsset -assetTag  $Computer.asset_tag
    }
    }
}

function Update-SnipeComputer{ 
param(
        [Parameter()] $assetTag
    )

    #Connect-SnipeUIT

    try {
        $Asset = Get-SnipeitAsset -asset_tag $assetTag
        $Computer = Get-ADComputer -Filter 'CN -like $assetTag' -Properties CN, Created, IPv4Address, IPv6Address, LastLogonDate, Modified, OperatingSystem, OperatingSystemVersion
        $LastModifiedSnipe = [DateTime]$Asset.custom_fields.Modified.value

        $statusID = Get-SnipeitStatus -id 2

        $LastModifiedAD = [DateTime]$Computer.Modified 
 
        if($LastModifiedSnipe -lt $LastModifiedAD){ 
 
            $customFields = @{ 
            "_snipeit_os_5" = $Computer.OperatingSystem  
            "_snipeit_os_version_9" = $Computer.OperatingSystemVersion 
            "_snipeit_modified_10" = $Computer.Modified.ToString() 
            "_snipeit_created_11" = $Computer.Created.ToString() 
            "_snipeit_last_logon_date_12" = $Computer.LastLogonDate.ToString() 
            } 
 
            if($Computer.IPv4Address){ 
                $customFields += @{"_snipeit_ipv4_address_6" = $Computer.IPv4Address} 
            } 
            if($Computer.IPv6Address){ 
                $customFields += @{"_snipeit_ipv6_address_7" = $Computer.IPv6Address} 
            } 
            Set-SnipeitAsset -id $Asset.id -customfields $customFields -status_id $statusID
 
        } 
        
    } 
    catch { 
        throw "Asset does not exist" 
    } 
 

}

function Update-AllSnipeVMs{

    $VMs = Get-VM
    
    foreach($VM in $VMs){
        
        $Asset = Get-SnipeitAsset -asset_tag $VM.Id
        if($Asset){
            Update-SnipeVM -VM $VM -assetTag $VM.Id
        }
        else{
            Add-SnipeVM $VM
        }
        
    }
}

function Update-AllSnipeComputers{
$Computers = Get-ADComputer -Filter * -Properties CN

Foreach ($Computer in $Computers){
        if(Get-SnipeitAsset -asset_tag $Computer.CN){
            try{
            Update-SnipeComputer -assetTag $Computer.CN
            }
            catch{}
        }
        else{
            Add-SnipeComputer -Computer $Computer
        }
    }

}


#Prints a pie chart of the Assets in Snipe by Model
Function Out-SnipeAssetsbyModel{
    param(
        #sets default value for num parameter if the parameter was not provided.
        [Parameter()][ValidateRange(1,10)][Int]$num= 4
    )

    Connect-SnipeUIT

    $assets = Get-SnipeitAsset -all
    $assetModels = $assets.model | Group-Object -Property name |Sort-Object -Descending Count
    $printData = $assetModels | Select-Object -First $num Name, Count
    $count

    foreach ($assetModel in $assetModels){
        if (!(($printData.name).Contains($assetModel.Name))) {
            $count += $assetModel.Count
        }
    }

    if($count){
        $printData += [PSCustomObject]@{
            Name     = 'Other'
            Count = $count
        }
    }

    $printData | Out-PieChart -PieChartTitle "Assets by Model" -DisplayToScreen
}





#Function for printing Pie Charts is from: https://4sysops.com/archives/convert-powershell-csv-output-into-a-pie-chart/
Function Out-PieChart {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline)]
        [psobject] $inputObject,
        [Parameter()]
        [string] $PieChartTitle,
        [Parameter()]
        [int] $ChartWidth = 800,
        [Parameter()]
        [int] $ChartHeight = 400,
        [Parameter()]
        [string[]] $NameProperty,
        [Parameter()]
        [string] $ValueProperty,
        [Parameter()]
        [switch] $Pie3D,
        [Parameter()]
        [switch] $DisplayToScreen,
        [Parameter()]
        [string] $saveImage
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
        $chart.series["Data"].Label = "#PERCENT\n#VALX"
        # Set ArrayList
        $XColumn = [System.Collections.ArrayList]::new()
        $yColumn = [System.Collections.ArrayList]::new()
    }
    process {
        if (-not $valueProperty) {
            $numericProperties = foreach ($property in $inputObject.PSObject.Properties) {
                if ([Double]::TryParse($property.Value, [Ref]$null)) {
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
            try{
                if (Test-Path (Split-Path $saveImage -Parent)) {
                    $SaveImage = $pscmdlet.GetUnresolvedProviderPathFromPSPath($saveImage)
                    $Chart.SaveImage($saveImage, "png")
                } else {
                    throw 'Invalid path, the parent directory must exist'
                }
            } catch {
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
            $Form.controls.add($Chart)
            $Chart.Anchor = 'Bottom, Right, Top, Left'
            $Form.Add_KeyDown({
                if ($_.KeyCode -eq "Escape") { $Form.Close() }
            })
            $Form.Add_Shown( {$Form.Activate()})
            $Form.ShowDialog() | Out-Null
        }
    }
}

Function Get-SnipeITData{
param (
		[Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][outputTypes]$Type
	)
    $command = "Get-Snipeit$Type -all"

    $data = Invoke-Expression $command
    return $data
}

Function Out-SnipeITReport{
param (
        [Parameter(Mandatory=$true)][ValidateScript({
            if($_ | Test-Path ){
                throw "File/Folder alreadys exists" 
            }
            if( -Not ( $_ | Split-Path | Test-Path) ){
                throw "Folder path to new file does not exist." 
            }
            if($_ -notmatch "(\.csv)"){
                throw "The file specified in the path argument must be a .csv"
            }
            return $true
        })][System.IO.FileInfo]$Path,
		[Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][outputTypes]$Type
	)
    Get-SnipeITData -Type $Type | Export-Csv -Path $Path

}

Function Out-SnipeITAllReports{

param (
        [Parameter(Mandatory=$true)][ValidateScript({
            if(-Not($_ | Test-Path )){
                throw "File/folder does not exist" 
            }
            if(-Not ($_ | Test-Path -PathType Container) ){
                throw "The Path argument must be a folder."
            }
            return $true
        })][System.IO.FileInfo]$Path
	)

$date = (get-date -format yyyy-MM-dd_HH-mm-ss)
$folderName = "InventoryReport - $date"
New-Item -Path $Path -ItemType "directory" -Name $folderName

$Path = "$Path\$folderName" 
foreach ($type in [outputTypes].GetEnumNames()){
    $NewPath = "$Path\$type.csv"
    $NewPath
    Out-SnipeITReport -Path $NewPath -Type $type
    }
}





#Export all functions that can be executed by a user
Export-ModuleMember -Function 'Add-SnipeVM'
Export-ModuleMember -Function 'Add-AllVMstoSnipe'
Export-ModuleMember -Function 'Remove-AllSnipeVMs'
Export-ModuleMember -Function 'Update-SnipeVM'
Export-ModuleMember -Function 'Update-AllSnipeVMs'
Export-ModuleMember -Function 'Archive-SnipeAsset'


Export-ModuleMember -Function 'Add-SnipeComputer'
Export-ModuleMember -Function 'Add-AllComputerstoSnipe'
Export-ModuleMember -Function 'Remove-AllSnipeComputers'
Export-ModuleMember -Function 'Update-SnipeComputer'
Export-ModuleMember -Function 'Update-AllSnipeComputers'

Export-ModuleMember -Function 'Out-SnipeAssetsbyModel'
Export-ModuleMember -Function 'Connect-SnipeUIT'

Export-ModuleMember -Function 'Out-SnipeITAllReports'
Export-ModuleMember -Function 'Out-SnipeITReport'
Export-ModuleMember -Function 'Get-SnipeITData'