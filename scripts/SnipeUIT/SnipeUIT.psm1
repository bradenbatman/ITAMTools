Import-module VMware.PowerCLI
Import-module SnipeitPS

$SnipeURL = "http:\\localhost"
$SnipeApiKey = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIzIiwianRpIjoiOTlmMzhiNDI0NmI0YWEwNTBiZTlhODY1ZDc0ZGRhNjVmNGI3YTk0NmJiZWE1OTA4NTMzY2VjMjZkNDE2MTU4M2IwYjEyY2ZlODUwYzYyZWMiLCJpYXQiOjE2MzQxODYzMDIsIm5iZiI6MTYzNDE4NjMwMiwiZXhwIjoyMTA3NTcxOTAxLCJzdWIiOiIxIiwic2NvcGVzIjpbXX0.PONNBVZ5qtK39fL58Osh4Qh6ZvdYPlJQnAZ1HUInyMDa0IF23u1XfdNMowGgg5eTgNP71ID6doH4ZHRPUq30EsQOp7xKSUPCNtMdZlPVVsm7dj8Vi6DYPAmyBIgHiXPHMc_0o3QGU_0B6LosFfpdXgDim3v7_ibqoo7__7JVNs8IyS4VupHjBGVFcbLSQfzwcFLFcIaj1AiqApj9bewmVA9L9m51ak-7cLxyep3fSBWZN4YJ-booXLKX4oxKyj8iP34zYpf4E1RjVHWRhC3jXhbFyxTLQzlnQfRg3Za9dUp4n8TVl9mxhwuXy3oEqX3UiBMLpoJKwgnPrPQLIAZ_GV6C1hjqkE0VGiNnw41ToBUx8fF0BOuaBdllVeZJCLD827DXsrHgE6Zumd1v95SuO8pjJ5OzUAzXrFqvVjbyTTxvEeqrORvOTDPDRKfkC-5HL4EdDE6NHdKeQOiVRWmiGYNZa5YVHAowz-wFVJGq20ENgVMp5dvWrvDMaqB8cvrdjhE9n_LKuGDe_IfPU3vI6lYAO0KGY3MXe8EOh1FLzilpT2SJWcD3CEurKoovDG7I3m7BG-PFzPsEyAc3IWCCvTS81Dt_t_kWNg7A4r1Cd467pRAca_oC2NYZ49A1hFtJyrfBgA6ZeZVzMLgyQMSdQyfvCROGmIV5XaeMbFXPOmA"

$VIServer = "marn-vb-vmsa01.iwunet.indwes.edu"

function Connect-SnipeUIT{
#New-VICredentialStoreItem -Host marn-vb-vmsa01.iwunet.indwes.edu -User b.batman-stwk@indwes.edu -Password *
Connect-SnipeitPS -url $SnipeURL -apiKey $SnipeApiKey
Connect-VIServer -Server $VIServer


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

    Connect-SnipeUIT

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
        "_snipeit_number_of_cpus_6" = $VM.NumCpu
        "_snipeit_memory_gb_7" = $VM.MemoryGB
        "_snipeit_ip_addresses_8" = $ipString
        "_snipeit_os_10" = $VMGuest.OSFullName
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


function Add-AllVMstoSnipe{
    Connect-SnipeUIT

    $VMs = Get-VM

    Foreach ($VM in $VMs){
        Add-SnipeVM -VM $VM
    }
}

function Remove-AllSnipeVMs{
    Connect-SnipeUIT
    Remove-SnipeitAsset -id (Get-SnipeitAsset -model_id $VMModelID).id
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

        if($VM.PowerState -like "PoweredOn" -and $Asset.status_label.name -ne "Powered On"){
            Set-SnipeitAsset -id $id -status_id $onStatusID
        }
        elseif($VM.PowerState -like "PoweredOff" -and $Asset.status_label.name -ne "Powered Off"){
            Set-SnipeitAsset -id $id -status_id $offStatusID
        }
    }
    catch {
        "Asset does not exist on SnipeIT"
    }
    
}


function Update-AllSnipeVMs{
    Connect-SnipeUIT

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

Export-ModuleMember -Function 'Out-SnipeAssetsbyModel'
Export-ModuleMember -Function 'Connect-SnipeUIT'

Export-ModuleMember -Function 'Out-SnipeITAllReports'
Export-ModuleMember -Function 'Out-SnipeITReport'
Export-ModuleMember -Function 'Get-SnipeITData'