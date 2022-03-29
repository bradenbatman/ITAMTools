Import-module VMware.PowerCLI
Import-Module Pester

Describe 'Test: Remove-SnipeVM' {
    It 'Check that removing a VM is reflected on ITAM' {
        $VM= Get-VM -Name Snipe-IT_Demo
        Remove-SnipeitAsset -id (Get-SnipeitAsset -asset_tag $VM.Id).id
        
        Get-SnipeitAsset -asset_tag $VM.Id | Should -Be $null
    }
}

Describe 'Test: Add-ITAMVM' {
    It 'Check that adding a VM shows up on ITAM' {
        $VM= Get-VM -Name Snipe-IT_Demo
        Add-ITAMVM -VM $VM

        (Get-SnipeitAsset -asset_tag $VM.Id).asset_tag | Should -Be $VM.Id
    }
}

Describe 'Test: Update-ITAMVM Error' {
    It 'Attempt to update a VM on ITAM that does not exist in VMware' {
        $errorThrown = $false
        try {
            $VM = Get-VM -Name Snipe-IT_Demo
            Update-ITAMVM -VM 1234 -$assetTag $VM.Id 
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $true
    }
}

Describe 'Test: Add-ITAMVM Error' {
    It 'Attempt to add a VM that does not exist on ITAM' {
        $errorThrown = $false
        try {
            $VM = Get-VM -Name Snipe-IT_Demos
            Add-ITAMVM -VM $VM
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $true
    }
}

Describe 'Test: Out-ITAMAssetsbyModel Parameter issue' {
    It 'Attempt to supply parameter outside of expected range' {
        $errorThrown = $false
        try {
            Out-ITAMAssetsbyModel -num 15 -testPrint $true
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $true
    }
}

Describe 'Test: Out-ITAMAssetsbyModel' {
    It 'Attempt to run command Out-ITAMAssetsbyModel' {
        $errorThrown = $false
        try {
            Out-ITAMAssetsbyModel -num 3 -testPrint $true
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $false
    }
}

#Failure test for Path validation
Describe 'Test: Out-ITAMReport Fail 1' {
    It 'Attempt to run command Out-ITAMReport with path: existing file' {
        $errorThrown = $false
        try {
            Out-ITAMReport -Type Accessory -Path C:\Users\b.batman-stwk\Documents\testing\test.csv
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $true
    }
}

#Failure test for Path validation
Describe 'Test: Out-ITAMReport Fail 2' {
    It 'Attempt to run command Out-ITAMReport with path: path does not exist' {
        $errorThrown = $false
        try {
            Out-ITAMReport -Type Accessory -Path C:\Users\b.ba\out.csv
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $true
    }
}

#Failure test for Path validation
Describe 'Test: Out-ITAMReport Fail 3' {
    It 'Attempt to run command Out-ITAMReport with path: new file is not .csv' {
        $errorThrown = $false
        try {
            Out-ITAMReport -Type Accessory -Path C:\Users\b.batman-stwk\Documents\testing\out
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $true
    }
}

#Happy path test for Out-ITAMReport
Describe 'Test: Out-ITAMReport Pass' {
    It 'Attempt to run command Out-ITAMReport' {
        $errorThrown = $false
        try {
            $date = (get-date -format yyyy-MM-dd_HH-mm-ss)
            Out-ITAMReport -Type Accessory -Path "C:\Users\b.batman-stwk\Documents\testing\out$date.csv"
        }
        catch {
            $errorThrown = $true
        }

        $errorThrown | Should -Be $false
    }
}

#Failure test for Path validation
Describe 'Test: Out-AllITAMReports Fail 1' {
    It 'Attempt to run command Out-AllITAMReports with path: path does not exist' {
        $errorThrown = $false
        try {
            Out-AllITAMReports -Path C:\Users\b.ba\out
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $true
    }
}

#Failure test for Path validation
Describe 'Test: Out-AllITAMReports Fail 2' {
    It 'Attempt to run command Out-AllITAMReports with path: The Path argument must be a folder' {
        $errorThrown = $false
        try {
            Out-AllITAMReports -Path C:\Users\b.batman-stwk\Documents\testing\test.csv
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $true
    }
}


#Failure test for Type param validation
Describe 'Test: Get-ITAMData Fail' {
    It 'Attempt to run command Get-ITAMData with incorrect type' {
        $errorThrown = $false
        try {
            Get-ITAMData -Type null
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $true
    }
}

#Happy path test for Type param validation
Describe 'Test: Get-ITAMData pass' {
    It 'Attempt to run command Get-ITAMData' {
        $errorThrown = $false
        try {
            Get-ITAMData -Type Accessory
        }
        catch {
            $errorThrown = $true
        }
               

        $errorThrown | Should -Be $false
    }
}

Describe 'Test: Remove-SnipeComputer' {
    It 'Check that removing a Computer is reflected on ITAM' {
        $asset = Get-SnipeitAsset -search DEV-INV01
        Remove-SnipeitAsset -id $asset.id
        
        Get-SnipeitAsset -asset_tag $asset.asset_tag | Should -Be $null
    }
}

Describe 'Test: Add-ITAMComputer' {
    It 'Check that adding a Computer shows up on ITAM' {
        $Computer = Get-ADComputer -Filter 'Name -like "DEV-INV01"'
        Add-ITAMComputer -Computer $Computer

        (Get-SnipeitAsset -asset_tag $Computer.Name).asset_tag | Should -Be $Computer.Name
    }
}

Describe 'Test: All ADComputers exist in ITAM' {
    It 'Check that all AD Computers have been added to ITAM' {

        $allAssetsExist = $false
        
        if ((Get-SnipeitModel -search "Computer").assets_count -le (get-adcomputer -Filter *).count) {
            $allAssetsExist = $true
        }
               
        $allAssetsExist | Should -Be $true
    }
}

Describe 'Test: Update-ITAMComputer Error' {
    It 'Attempt to update a Computer on ITAM that does not exist on AD' {
        $errorThrown = $false
        try {
            $Computer = Get-ADComputer -Filter 'Name -like "12345"'
            Update-ITAMComputer -assetTag $Computer.SIN
        }
        catch {
            $errorThrown = $true
        } 

        $errorThrown | Should -Be $true
    }
}

Describe 'Test: Add-ITAMComputer Error' {
    It 'Attempt to add a Computer that does not exist on ITAM' {
        $errorThrown = $false
        try {
            $Computer = Get-ADComputer -Filter 'Name -like "12346785"'
            Add-ITAMComputer -Computer $Computer
        }
        catch {
            $errorThrown = $true
        }
               
        $errorThrown | Should -Be $true
    }
}

Describe 'Test: Archive-SnipeVm Archive Happy Path' {
    It 'Attempt to arhive a VM' {
        $vm = Get-SnipeitAsset -asset_tag DEV-INV01
        $status = $vm.status_label
        $isArchived = $false
        try {
            Archive-ITAMAsset -assetTag $vm.asset_tag
            $vmUpdated = Get-SnipeitAsset -asset_tag DEV-INV01

            if($vmUpdated.status_label.name -eq "Archived"){
                $isArchived = $true
            }

            Set-SnipeitAsset -id $vmUpdated.id -status_id $status.id -archived $false
        }
        catch {
            $isArchived = $false
        }
               
        $isArchived | Should -Be $true
    }
}

Describe 'Test: Update-ITAMVM Archive Error Path' {
    It 'Attempt to arhive a VM that does not exist' {
        try {
            $vm = Get-SnipeitAsset -asset_tag '1243345234'
            $status = $vm.status_label
            $isArchived = $false
            Archive-ITAMAsset -assetTag $vm.asset_tag
            $vmUpdated = Get-SnipeitAsset -asset_tag '1243345234'

            if($vmUpdated.status_label.name -eq "Archived"){
                $isArchived = $true
            }

            Set-SnipeitAsset -id $vmUpdated.id -status_id $status.id -archived $false
        }
        catch {
            $isArchived = $false
        }
               
        $isArchived | Should -Be $false
    }
}

Describe 'Test: Archive-SnipeComputer Archive Happy Path' {
    It 'Attempt to archive a Computer' {
        $Computer = Get-SnipeitAsset -asset_tag DEV-INV01
        $status = $Computer.status_label
        $isArchived = $false
        try {
            Archive-ITAMAsset -assetTag $Computer.asset_tag
            $ComputerUpdated = Get-SnipeitAsset -asset_tag DEV-INV01

            if($ComputerUpdated.status_label.name -eq "Archived"){
                $isArchived = $true
            }

            Set-SnipeitAsset -id $ComputerUpdated.id -status_id $status.id -archived $false
        }
        catch {
            $isArchived = $false
        }
               
        $isArchived | Should -Be $true
    }
}

Describe 'Test: Update-ITAMComputer Happy Path' {
    It 'Attempt to update a Computer' {
        $updateFailed = $false
        
        try {
            $Computer = Get-SnipeitAsset -asset_tag DEV-INV01
            Update-ITAMComputer -assetTag $Computer.asset_tag
        }
        catch {
            $updateFailed = $true
        }
               
        $updateFailed | Should -Be $false
    }
}

Describe 'Test: Update-ITAMComputer Archive Error Path' {
    It 'Attempt to arhive a Computer that does not exist' {
        try {
            $Computer = Get-SnipeitAsset -asset_tag 'NONREALASSET'
            $status = $Computer.status_label
            $isArchived = $false
            Archive-ITAMAsset -assetTag $Computer.asset_tag
            $ComputerUpdated = Get-SnipeitAsset -asset_tag 'NONREALASSET'

            if($ComputerUpdated.status_label.name -eq "Archived"){
                $isArchived = $true
            }

            Set-SnipeitAsset -id $vmUpdated.id -status_id $status.id -archived $false
        }
        catch {
            $isArchived = $false
        }
               
        $isArchived | Should -Be $false
    }
}

Describe 'Test: Update-ITAMComputer Happy Path' {
    It 'Attempt to update a Computer' {
        $updateFailed = $false
        
        try {
            $Computer = Get-SnipeitAsset -asset_tag "NONREALASSET"
            Update-ITAMComputer -assetTag $Computer.asset_tag
        }
        catch {
            $updateFailed = $true
        }
               
        $updateFailed | Should -Be $true
    }
}

Describe 'Test: Get-ITAMAsset Happy Path' {
    It 'get-asset with all parameters' {
        $exceptionThrown = $false
        
        try {
            $Asset = Get-ITAMAsset DEV-INV01
        }
        catch {
            $exceptionThrown = $true
        }
               
        $exceptionThrown | Should -Be $false
    }
}

Describe 'Test: Get-ITAMAsset Happy Path 2' {
    It 'get-asset with limited parameters' {
        $exceptionThrown = $false
        
        try {
            $Asset = Get-ITAMAsset DEV-INV01 -Properties "id"
        }
        catch {
            $exceptionThrown = $true
        }
               
        $exceptionThrown | Should -Be $false
    }
}

Describe 'Test: Get-ITAMAsset Fail Path' {
    It 'get-asset with nonreal asset' {
        $exceptionThrown = $false
        
        try {
            $Asset = Get-ITAMAsset "NONREALASSET"
        }
        catch {
            $exceptionThrown = $true
        }
               
        $exceptionThrown | Should -Be $true
    }
}

Describe 'Test: Get-ITAMAsset Fail Path 2' {
    It 'get-asset with nonreal parameters' {
        $exceptionThrown = $false
        
        try {
            $Asset = Get-ITAMAsset DEV-INV01 -Properties "id, nonrealprop"
        }
        catch {
            $exceptionThrown = $true
        }
               
        $exceptionThrown | Should -Be $true
    }
}

Describe 'Test: Out-ITAMPowerBIReport Happy Path' {
    It 'Out-ITAMPowerBIReport, verifying that a report is successfully created.' {
        $DataSetExists = $false
        
        $name = 'testReport'
        Out-ITAMPowerBIReport -name $name -search 'IWU'


        try {
            $DataSet = Get-PowerBIDataset -name $name 
            if($DataSet){
                $DataSetExists = $true
            }
        }
        catch {
            $DataSetExists = $false
        }

        $DataSetExists | Should -Be $true
    }
}

Describe 'Test: Out-ITAMPowerBIReport Happy Path: Exception not thrown' {
    It 'Out-ITAMPowerBIReport, verifying that exception is not thrown' {
        $exceptionThrown = $false
        
        $name = 'testReport'

        try {
            Out-ITAMPowerBIReport -name $name -search 'IWU'
        }
        catch {
            $exceptionThrown = $true
        }

        $exceptionThrown | Should -Be $false
    }
}

Describe 'Test: Out-ITAMPowerBIReport Fail Path' {
    It 'Out-ITAMPowerBIReport with missing parameters' {
        $exceptionThrown = $false
        
        try {
            Out-ITAMPowerBIReport -name $null -search $null
        }
        catch {
            $exceptionThrown = $true
        }
               
        $exceptionThrown | Should -Be $true
    }
}

Describe 'Test: Out-ITAMPowerBIReport Fail Path 2' {
    It 'Out-ITAMPowerBIReport with non real asset' {
        $exceptionThrown = $false
        
        $name = 'testReport'
        try {
            Out-ITAMPowerBIReport -name $name -search '12345821749587'
        }
        catch {
            $exceptionThrown = $true
        }
               
        $exceptionThrown | Should -Be $true
    }
}

Describe 'Test: Update-ITAMVM VM Assignment Happy Path' {
    It 'Update-ITAMVM with test values' {
        

        $VM = Get-VM -Name 'PRD-WEB04'
        $Computer = Get-SnipeitAsset -asset_tag PRD-WEB04
        
        Update-ITAMVM -VM $VM -assetTag $VM.id

        $updatedVM = Get-SnipeitAsset -asset_tag $VM.id

        if($Computer){
            $correctAssignment = ($updatedVM.assigned_to.id -eq $Computer.id)
        }
        else{
            $correctAssignment = $true
        }
              
        $correctAssignment | Should -Be $true
    }

}


Describe 'Test: Update-ITAMVM VM Assignment Happy Path 2' {
    It 'Update-ITAMVM with test values' {
        
        $exceptionThrown = $false

        $VM = Get-VM -Name 'PRD-WEB04'

        try{
            Update-ITAMVM -VM $VM -assetTag $VM.id
        }
        catch{
           $exceptionThrown = $true;
        }
              
        $exceptionThrown | Should -Be $false
    }

}

Describe 'Test: Add-ITAMVM VM Assignment Fail Path' {
    It 'Update-ITAMVM without existing parent computer' {
        
        $VMname = 'DEV-INF01'
        $VM = Get-VM -Name $VMname
        $Computer = Get-SnipeitAsset -asset_tag $VMname
        
        Update-ITAMVM -VM $VM -assetTag $VM.id

        $updatedVM = Get-SnipeitAsset -asset_tag $VM.id

        #If the computer does not exists and the vm is assigned, assignment is incorrect
        if(!$updatedVM.assigned_to){
            $correctAssignment = $false
        }
        else{
            $correctAssignment = $true
        }
              
        $correctAssignment | Should -Be $false
    }
}


Describe 'Test: Add-ITAMVM VM Assignment Fail Path 2' {
    It 'Update-ITAMVM where parent computer is null ' {
        
        $VMname = 'DEV-INF01'
        $VM = Get-VM -Name $VMname
        $Computer = Get-SnipeitAsset -asset_tag $VMname
        
        Update-ITAMVM -VM $VM -assetTag $VM.id

        $updatedVM = Get-SnipeitAsset -asset_tag $VM.id

        (!$updatedVM.assigned_to -and !$Computer) | Should -Be $true
    }
}