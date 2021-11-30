Import-module VMware.PowerCLI
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
    It 'Attempt to supply parameter outside of expected range' {
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

#Failure test for Path validation
Describe 'Test: Out-SnipeITReport Fail 1' {
    It 'Attempt to run command Out-SnipeITReport with path: existing file' {
        $errorThrown = $false
        try {
            Out-SnipeITReport -Type Accessory -Path C:\Users\b.batman-stwk\Documents\testing\test.csv
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $true
    }
}

#Failure test for Path validation
Describe 'Test: Out-SnipeITReport Fail 2' {
    It 'Attempt to run command Out-SnipeITReport with path: path does not exist' {
        $errorThrown = $false
        try {
            Out-SnipeITReport -Type Accessory -Path C:\Users\b.ba\out.csv
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $true
    }
}

#Failure test for Path validation
Describe 'Test: Out-SnipeITReport Fail 3' {
    It 'Attempt to run command Out-SnipeITReport with path: new file is not .csv' {
        $errorThrown = $false
        try {
            Out-SnipeITReport -Type Accessory -Path C:\Users\b.batman-stwk\Documents\testing\out
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $true
    }
}

#Happy path test for Out-SnipeITReport
Describe 'Test: Out-SnipeITReport Pass' {
    It 'Attempt to run command Out-SnipeITReport' {
        $errorThrown = $false
        try {
            $date = (get-date -format yyyy-MM-dd_HH-mm-ss)
            Out-SnipeITReport -Type Accessory -Path "C:\Users\b.batman-stwk\Documents\testing\out$date.csv"
        }
        catch {
            $errorThrown = $true
        }

        $errorThrown | Should -Be $false
    }
}

#Failure test for Path validation
Describe 'Test: Out-SnipeITAllReports Fail 1' {
    It 'Attempt to run command Out-SnipeITAllReports with path: path does not exist' {
        $errorThrown = $false
        try {
            Out-SnipeITAllReports -Path C:\Users\b.ba\out
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $true
    }
}

#Failure test for Path validation
Describe 'Test: Out-SnipeITAllReports Fail 2' {
    It 'Attempt to run command Out-SnipeITAllReports with path: The Path argument must be a folder' {
        $errorThrown = $false
        try {
            Out-SnipeITAllReports -Path C:\Users\b.batman-stwk\Documents\testing\test.csv
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $true
    }
}


#Failure test for Type param validation
Describe 'Test: Get-SnipeITData Fail' {
    It 'Attempt to run command Get-SnipeITData with incorrect type' {
        $errorThrown = $false
        try {
            Get-SnipeITData -Type null
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $true
    }
}

#Happy path test for Type param validation
Describe 'Test: Get-SnipeITData pass' {
    It 'Attempt to run command Get-SnipeITData' {
        $errorThrown = $false
        try {
            Get-SnipeITData -Type Accessory
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $false
    }
}

Describe 'Test: Remove-SnipeComputer' {
    It 'Check that removing a Computer is reflected on Snipe' {
        $asset = Get-SnipeitAsset -search DOXIE
        Remove-SnipeitAsset -id $asset.id
        
        Get-SnipeitAsset -asset_tag $asset.asset_tag | Should -Be $null
    }
}

Describe 'Test: Add-SnipeComputer' {
    It 'Check that adding a Computer shows up on Snipe' {
        $Computer = Get-ADComputer -Filter 'Name -like "DOXIE"'
        Add-SnipeComputer -Computer $Computer

        (Get-SnipeitAsset -asset_tag $Computer.Sid).asset_tag | Should -Be $Computer.Sid
    }
}

Describe 'Test: All ADComputers exist in Snipe' {
    It 'Check that all AD Computers have been added to Snipe' {

        $allAssetsExist = $false
        
        if ((Get-SnipeitModel -search "Computer").assets_count -eq (get-adcomputer -Filter *).count) {
            $allAssetsExist = $true
        }
               
        $allAssetsExist | Should -Be $true
    }
}

Describe 'Test: Update-SnipeComputer Error' {
    It 'Attempt to update a Computer on Snipe that does not exist on AD' {
        $errorThrown = $false
        try {
            $Computer = Get-ADComputer -Filter 'Name -like "12345"'
            Update-SnipeComputer -assetTag $Computer.SIN
        }
        catch {
            $errorThrown = $true
        } 

        $errorThrown | Should -Be $true
    }
}

Describe 'Test: Update-SnipeComputer' {
    It 'Verify that a Computer can be updated on Snipe' {
        $Computer = Get-ADComputer -Filter 'Name -like "DOXIE"' -Properties Modified
        Update-SnipeComputer -assetTag $Computer.SID

        (Get-SnipeitAsset -asset_tag $Computer.Sid).custom_fields.Modified.value | Should -Be $Computer.Modified.toString()
    }
}



Describe 'Test: Add-SnipeComputer Error' {
    It 'Attempt to add a Computer that does not exist on Snipe' {
        $errorThrown = $false
        try {
            $Computer = Get-ADComputer -Filter 'Name -like "12345"'
            Add-SnipeComputer -Computer $Computer
        }
        catch {
            $errorThrown = $true
        }
               
        $errorThrown | Should -Be $true
    }
}