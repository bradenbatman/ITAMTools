$SnipeURL = "http:\\localhost"
$SnipeApiKey = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIzIiwianRpIjoiOTlmMzhiNDI0NmI0YWEwNTBiZTlhODY1ZDc0ZGRhNjVmNGI3YTk0NmJiZWE1OTA4NTMzY2VjMjZkNDE2MTU4M2IwYjEyY2ZlODUwYzYyZWMiLCJpYXQiOjE2MzQxODYzMDIsIm5iZiI6MTYzNDE4NjMwMiwiZXhwIjoyMTA3NTcxOTAxLCJzdWIiOiIxIiwic2NvcGVzIjpbXX0.PONNBVZ5qtK39fL58Osh4Qh6ZvdYPlJQnAZ1HUInyMDa0IF23u1XfdNMowGgg5eTgNP71ID6doH4ZHRPUq30EsQOp7xKSUPCNtMdZlPVVsm7dj8Vi6DYPAmyBIgHiXPHMc_0o3QGU_0B6LosFfpdXgDim3v7_ibqoo7__7JVNs8IyS4VupHjBGVFcbLSQfzwcFLFcIaj1AiqApj9bewmVA9L9m51ak-7cLxyep3fSBWZN4YJ-booXLKX4oxKyj8iP34zYpf4E1RjVHWRhC3jXhbFyxTLQzlnQfRg3Za9dUp4n8TVl9mxhwuXy3oEqX3UiBMLpoJKwgnPrPQLIAZ_GV6C1hjqkE0VGiNnw41ToBUx8fF0BOuaBdllVeZJCLD827DXsrHgE6Zumd1v95SuO8pjJ5OzUAzXrFqvVjbyTTxvEeqrORvOTDPDRKfkC-5HL4EdDE6NHdKeQOiVRWmiGYNZa5YVHAowz-wFVJGq20ENgVMp5dvWrvDMaqB8cvrdjhE9n_LKuGDe_IfPU3vI6lYAO0KGY3MXe8EOh1FLzilpT2SJWcD3CEurKoovDG7I3m7BG-PFzPsEyAc3IWCCvTS81Dt_t_kWNg7A4r1Cd467pRAca_oC2NYZ49A1hFtJyrfBgA6ZeZVzMLgyQMSdQyfvCROGmIV5XaeMbFXPOmA"

$VIServer = "marn-vb-vmsa01.iwunet.indwes.edu"

#New-VICredentialStoreItem -Host marn-vb-vmsa01.iwunet.indwes.edu -User b.batman-stwk@indwes.edu -Password *

Connect-VIServer -Server $VIServer
Connect-SnipeitPS -url $SnipeURL -apiKey $SnipeApiKey

$onStatusID = (Get-SnipeitStatus -Search "Powered On").id
$offStatusID = (Get-SnipeitStatus -Search "Powered Off").id
$VMModelID = (Get-SnipeitModel -search "Virtual Machine").id

function Add-SnipeVM{

    param(
        [Parameter()] [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl] $VM
    )

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
    $VMs = Get-VM

    Foreach ($VM in $VMs){
        Add-SnipeVM -VM $VM
    }
}

function Remove-AllSnipeVMs{
    Remove-SnipeitAsset -id (Get-SnipeitAsset -model_id $VMModelID).id
}

function Update-SnipeVM{
    param(
        [Parameter()] [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl] $VM,
        [Parameter()] $assetTag
    )

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