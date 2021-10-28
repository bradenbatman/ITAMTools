
#Update-Module SnipeUIT
Import-Module SnipeUIT

Import-Module Pester

Describe 'Test: Remove-SnipeVM' {
    It 'Check that removing a VM is reflected on Snipe' {
        $VM= Get-VM -Name Snipe-IT_Demo
        Remove-SnipeitAsset -id (Get-SnipeitAsset -asset_tag $VM.Id).id

        Get-SnipeitAsset -asset_tag $VM.Id | Should -Be $null
    }
}

Describe 'Test: Add-SnipeVM' {
    It 'Check that adding a VM shows up on Snipe' {
        $VM= Get-VM -Name Snipe-IT_Demo
        Add-SnipeVM -VM $VM

        (Get-SnipeitAsset -asset_tag $VM.Id).asset_tag | Should -Be $VM.Id
    }
}

Describe 'Test: Update-SnipeVM Error' {
    It 'Attempt to update a VM on Snipe that does not exist in VMware' {
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

Describe 'Test: Out-SnipeAssetsbyModel Parameter issue' {
    It 'Attempt to supply parameter outside of expeceted range' {
        $errorThrown = $false
        try {
            Out-SnipeAssetsbyModel -num 15
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $true
    }
}

Describe 'Test: Out-SnipeAssetsbyModel' {
    It 'Attempt to run command Out-SnipeAssetsbyModel' {
        $errorThrown = $false
        try {
            Out-SnipeAssetsbyModel -num 3
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $false
    }
}