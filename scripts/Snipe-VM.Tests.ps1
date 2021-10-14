Import-Module Pester

Describe 'Test: Each VM is on Snipe-IT' {
    It 'Check Snipe-IT VM count against VM Server count' {
        . C:\inetpub\wwwroot\snipe-uit\Snipe-UIT\scripts\Snipe-VM.ps1
        $SnipeVMCount = (Get-SnipeitModel -search "Virtual Machine").assets_count
        $VMs = Get-VM
        $SnipeVMCount | Should -Be $VMs.count
    }
}

Describe 'Test: Remove-SnipeVM' {
    It 'Check that removing a VM is reflected on Snipe' {
        . C:\inetpub\wwwroot\snipe-uit\Snipe-UIT\scripts\Snipe-VM.ps1
        $VM= Get-VM -Name Snipe-IT_Demo
        Remove-SnipeitAsset -id (Get-SnipeitAsset -asset_tag $VM.Id).id

        Get-SnipeitAsset -asset_tag $VM.Id | Should -Be $null
    }
}

Describe 'Test: Add-SnipeVM' {
    It 'Check that adding a VM shows up on Snipe' {
        . C:\inetpub\wwwroot\snipe-uit\Snipe-UIT\scripts\Snipe-VM.ps1
        $VM= Get-VM -Name Snipe-IT_Demo
        Add-SnipeVM -VM $VM

        (Get-SnipeitAsset -asset_tag $VM.Id).asset_tag | Should -Be $VM.Id
    }
}

Describe 'Test: Update-SnipeVM Error' {
    It 'Attempt to update a VM on Snipe that does not exist' {
        . C:\inetpub\wwwroot\snipe-uit\Snipe-UIT\scripts\Snipe-VM.ps1
        $errorThrown = $false
        try {
            $VM = Get-VM -Name Snipe-IT_Demo
            Update-SnipeVM -VM 1234 -$assetTag $VM.Id 
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $true
    }
}

Describe 'Test: Add-SnipeVM Error' {
    It 'Attempt to add a VM that does not exist on Snipe' {
        . C:\inetpub\wwwroot\snipe-uit\Snipe-UIT\scripts\Snipe-VM.ps1
        $errorThrown = $false
        try {
            $VM = Get-VM -Name Snipe-IT_Demos
            Add-SnipeVM -VM $VM
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $true
    }
}