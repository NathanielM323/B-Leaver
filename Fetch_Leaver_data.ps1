[System.Threading.Thread]::CurrentThread.ApartmentState = "STA"
#Install-Module AzureAD
connect-exchangeonline
try {Get-AzureADCurrentSessionInfo}

catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException]
{
Connect-AzureAD
}
#function to retrieve BBB users 
    function Get-BBBUserGroupsExport
    {
    param ([String]$username)
    #retrieves groups from AAD
    $username = 'jasleen.haer' 
    $accountName = Get-ADUser -Filter {SamAccountname -eq $username} l
     
    try{$membership = Get-AzureADUserMembership -ObjectId $accountName.UserPrincipalName | Where-Object {$_.ObjectType -eq "Group"} | Select-Object DisplayName

    #Sort membership by alphabetical order
    $membership = $membership |Sort-Object DisplayName
        # Create Directory to save
        $directoryPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\B-Leaver\Leaver Data\$username"
        if (-not (Test-Path -Path $directoryPath)) {
            New-Item -Path $directoryPath -ItemType "directory" | Out-Null
        }
    
        # Define CSV file path with the username in the filename
        $csvFilePath = "$directoryPath\($username) Group memberships.csv"

        $membership | Export-Csv -Path $csvFilePath -NoTypeInformation
        #[System.Windows.Forms.MessageBox]::Show("Group memberships exported as CSV to $directoryPath successfully.", "Export Complete")
    }
      
      catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException] {
    [System.Windows.Forms.MessageBox]::Show("You must connect to AzureAD/Identity before calling Azure commands", "ERROR_AZURE_AUTHENTICATION_REQUIRED_101")
}
      catch {
        if ($_.Exception.Message -like "*Cannot bind argument to parameter 'ObjectId' because it is null.*") {
            [System.Windows.Forms.MessageBox]::Show("The username '$username' is invalid or does not exist.", "ERROR_INVALID_USERNAME_012")
        }
        else {
            #[System.Windows.Forms.MessageBox]::Show("An error occurred: $($_.Exception.Message)", "Error")
        }
}
}
    function Get-SulcoUserGroupsExport
    {
    param ([String]$username)
    #retrieves groups from AAD   
    try{$membership = Get-AzureADUserMembership -ObjectId $username@startuploans.co.uk | Where-Object {$_.ObjectType -eq "Group"} | Select-Object DisplayName

    #Sort membership by alphabetical order
    $membership = $membership |Sort-Object DisplayName

    $directoryPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\B-Leaver\Leaver Data\$username"
        if (-not (Test-Path -Path $directoryPath)) {
            New-Item -Path $directoryPath -ItemType "directory" | Out-Null
        }
    
        # Define CSV file path with the username in the filename
        $csvFilePath = "$directoryPath\($username) Group memberships.csv"

        $membership | Export-Csv -Path $csvFilePath -NoTypeInformation
        #[System.Windows.Forms.MessageBox]::Show("Group memberships exported as CSV to $directoryPath successfully.", "Export Complete")
    }
      
      catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException] {
    [System.Windows.Forms.MessageBox]::Show("You must connect to AzureAD/Identity before calling Azure commands", "ERROR_AZURE_AUTHENTICATION_REQUIRED_101")
}

      catch {
        if ($_.Exception.Message -like "*Cannot bind argument to parameter 'ObjectId' because it is null.*") {
            [System.Windows.Forms.MessageBox]::Show("The username '$username' is invalid or does not exist.", "ERROR_INVALID_USERNAME_012")
        }
        else {
            #[System.Windows.Forms.MessageBox]::Show("An error occurred: $($_.Exception.Message)", "Error")
            }
        }
    }

    function Get-BBBUserDLExport
    {
       param ([String]$username)
       $accountName = Get-ADUser -Filter {SamAccountname -eq $username} -Properties EmailAddress
       $BBBaddress = $accountName.EmailAddress

        try {
            $allDistributionGroups = Get-DistributionGroup -ResultSize Unlimited
            $totalDLs = $allDistributionGroups.Count
            $progress = 0
            # Array to store DLs user has access to
            $DLswithAccess = @()
            # Loop through each DL to see if user has access
            foreach ($DL in $allDistributionGroups) {
                $members = Get-DistributionGroupMember -Identity $DL.Identity
                if ($members.PrimarySmtpAddress -contains $BBBaddress) {
                    $DLswithAccess += $DL.PrimarySmtpAddress
                }
                $progress++
                Write-Progress -Activity "Checking Distribution Groups" -Status "$progress out of $totalDLs" -PercentComplete (($progress / $totalDLs) * 100)
            }
            $directoryPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\B-Leaver\Leaver Data\$username"
            if (-not (Test-Path -Path $directoryPath)) {
                New-Item -Path $directoryPath -ItemType "directory" | Out-Null
            }

            # Define CSV file path with the username in the filename
            $csvFilePath = "$directoryPath\($username) Distribution_Lists.csv"
    
            # Create a custom object for each distribution list
            $DLObjects = $DLswithAccess | ForEach-Object {
                [PSCustomObject]@{
                    DistributionList = $_
                }
            }
    
            # Export the custom objects to CSV
            $DLObjects | Export-Csv -Path $csvFilePath -NoTypeInformation
    
            #[System.Windows.Forms.MessageBox]::Show("Group memberships exported as CSV to $directoryPath successfully.", "Export Complete")
        }

         catch [System.Management.Automation.CommandNotFoundException] 
    
            {
                if ($_.Exception.CommandName -eq 'Get-DistributionGroup') {
                    Write-Error "The Exchange Online PowerShell module is not loaded or not connected. Please ensure you've connected to Exchange Online before running this script."
                    [System.Windows.Forms.MessageBox]::Show("You must connect to MSExchange Online", "ERROR_MSE_AUTHENTICATION_REQUIRED_113")
                    # Optionally, you can try to connect to Exchange Online here
                    # Connect-ExchangeOnline
                }
                else {
                    # Re-throw if it's a different CommandNotFoundException
                    throw
                }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error occurred: $($_.Exception.Message)", "Error")
        }
    }
   
    function Get-SulcoUserDLExport
    {
     param ([String]$username)
     $sulcoAddress = "$username@startuploans.co.uk"
 
           try {
                $allDistributionGroups = Get-DistributionGroup -ResultSize Unlimited
                $totalDLs = $allDistributionGroups.Count
                $progress = 0
                # Array to store DLs user has access to
                $DLswithAccess = @()
                # Loop through each DL to see if user has access
                foreach ($DL in $allDistributionGroups) {
                    $members = Get-DistributionGroupMember -Identity $DL.Identity
                    if ($members.PrimarySmtpAddress -contains $sulcoAddress) {
                        $DLswithAccess += $DL.PrimarySmtpAddress
                    }
                    $progress++
                    Write-Progress -Activity "Checking Distribution Groups" -Status "$progress out of $totalDLs" -PercentComplete (($progress / $totalDLs) * 100)
                }
                $directoryPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\B-Leaver\Leaver Data\$username"
                if (-not (Test-Path -Path $directoryPath)) {
                    New-Item -Path $directoryPath -ItemType "directory" | Out-Null
                }

                # Define CSV file path with the username in the filename
                $csvFilePath = "$directoryPath\($username) Distribution_Lists.csv"
    
                # Create a custom object for each distribution list
                $DLObjects = $DLswithAccess | ForEach-Object {
                    [PSCustomObject]@{
                        DistributionList = $_
                    }
                }
    
                # Export the custom objects to CSV
                $DLObjects | Export-Csv -Path $csvFilePath -NoTypeInformation
    
                #[System.Windows.Forms.MessageBox]::Show("Group memberships exported as CSV to $directoryPath successfully.", "Export Complete")
            }

 catch [System.Management.Automation.CommandNotFoundException]    
    {
        if ($_.Exception.CommandName -eq 'Get-DistributionGroup') {
            Write-Error "The Exchange Online PowerShell module is not loaded or not connected. Please ensure you've connected to Exchange Online before running this script."
            # Optionally, you can try to connect to Exchange Online here
            # Connect-ExchangeOnline
        }
        else {
            # Re-throw if it's a different CommandNotFoundException
            throw
        }
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Error occurred: $($_.Exception.Message)", "Error")
}

    }

    function Get-BBBSharedMailbox {
    param ([String]$username)
    $accountName = Get-ADUser -Filter {SamAccountname -eq $username} -Properties EmailAddress
    $BBBaddress = $accountName.EmailAddress

    
    try {
        # Get all mailboxes
        $AllMailboxes = Get-Mailbox -ResultSize Unlimited
        $totalMailboxes = $AllMailboxes.Count
        $progress = 0

        # Array to store mailboxes the user has access to
        $MailboxesWithAccess = @()

        # Loop through each mailbox and check if the user has access
        foreach ($Mailbox in $AllMailboxes) {
            $MailboxPermissions = Get-MailboxPermission -Identity $Mailbox.DistinguishedName | Where-Object { $_.User -like $BBBeaddress -and $_.AccessRights -like "FullAccess" }
            if ($MailboxPermissions) {
                $MailboxesWithAccess += $Mailbox
            }

            # Update progress
            $progress++
            $percentComplete = ($progress / $totalMailboxes) * 100
            Write-Progress -Activity "Checking Mailbox Permissions" -Status "Progress: $percentComplete% Complete" -PercentComplete $percentComplete
        }

        # Create Directory to save
        $directoryPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\B-Leaver\Leaver Data\$username"
        if (-not (Test-Path -Path $directoryPath)) {
            New-Item -Path $directoryPath -ItemType "directory" | Out-Null
        }
    
        # Define CSV file path with the username in the filename
        $csvFilePath = "$directoryPath\($username) Mailbox access.csv"

        # Export mailboxes to CSV
        $MailboxesWithAccess | Export-Csv -Path $csvFilePath -NoTypeInformation
        #[System.Windows.Forms.MessageBox]::Show("Shared mailboxes exported as CSV to $directoryPath successfully.", "Export Complete")
        } catch {
        [System.Windows.Forms.MessageBox]::Show("Error occurred.", "Error")
            }
    }
    # Uncomment the below line if you want to display the mailboxes in the console
    # $MailboxesWithAccess | Select-Object DisplayName, UserPrincipalName | Out-Host
     

    function Get-SulcoMailbox
    {      
         param ([String]$username)
         $sulcoAddress = "$username@startuploans.co.uk"        
                 
          try {
            # Get all mailboxes
            $AllMailboxes = Get-Mailbox -ResultSize Unlimited
            $totalMailboxes = $AllMailboxes.Count
            $progress = 0

            # Array to store mailboxes the user has access to
            $MailboxesWithAccess = @()

            # Loop through each mailbox and check if the user has access
            foreach ($Mailbox in $AllMailboxes) {
                $MailboxPermissions = Get-MailboxPermission -Identity $Mailbox.DistinguishedName | Where-Object { $_.User -like $sulcoAddress -and $_.AccessRights -like "FullAccess" }
                if ($MailboxPermissions) {
                    $MailboxesWithAccess += $Mailbox
                }

                # Update progress
                $progress++
                $percentComplete = ($progress / $totalMailboxes) * 100
                Write-Progress -Activity "Checking Mailbox Permissions" -Status "Progress: $percentComplete% Complete" -PercentComplete $percentComplete
            }

            # Create Directory to save
            $directoryPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\B-Leaver\Leaver Data\$username"
            if (-not (Test-Path -Path $directoryPath)) {
                New-Item -Path $directoryPath -ItemType "directory" | Out-Null
            }
    
            # Define CSV file path with the username in the filename
            $csvFilePath = "$directoryPath\($username) Mailbox access.csv"

            # Export mailboxes to CSV
            $MailboxesWithAccess | Export-Csv -Path $csvFilePath -NoTypeInformation
            #[System.Windows.Forms.MessageBox]::Show("Shared mailboxes exported as CSV to $directoryPath successfully.", "Export Complete")
            } catch {
            [System.Windows.Forms.MessageBox]::Show("Error occurred.", "Error")
                }
    
    }

Add-Type -AssemblyName System.Windows.Forms

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "AzureAD User Membership"
$form.Size = New-Object System.Drawing.Size(600, 400)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false

# Create a drop-down menu (combo box)
$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(180, 50)
$comboBox.Size = New-Object System.Drawing.Size(250, 100)
$comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboBox.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
$comboBox.Items.Add("BBB") 
$comboBox.Items.Add("SULCO")
$form.Controls.Add($comboBox)

# Create a label for the drop-down menu
$Dropdownlabel = New-Object System.Windows.Forms.Label
$Dropdownlabel.Location = New-Object System.Drawing.Point(20, 50)
$Dropdownlabel.Size = New-Object System.Drawing.Size(90, 20)
$Dropdownlabel.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Regular)
$Dropdownlabel.Text = "Type of user:"
$form.Controls.Add($Dropdownlabel)

# Create a label for the username input
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(20, 100)
$label.Size = New-Object System.Drawing.Size(120, 20)
$label.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Regular)
$label.Text = "Username:"
$form.Controls.Add($label)

# Create a text box for username input
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(180, 95)
$textBox.Size = New-Object System.Drawing.Size(350, 60)
$textBox.Font = New-Object System.Drawing.Font("Arial", 20, [System.Drawing.FontStyle]::Regular)
$form.Controls.Add($textBox)

# Create a button to export as CSV
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(350, 180)
$exportButton.Size = New-Object System.Drawing.Size(180, 40)
$exportButton.Text = "Export as CSV"
$exportButton.Add_Click({
    $username = $textBox.Text

    #sees what option is selected
    if($comboBox.SelectedItem -eq "BBB")
    {
    Get-BBBUserGroupsExport -username $username
    }

    else
    {
    Get-SulcoUserGroupsExport -username $username
    }
    
    #Calls function to display save dialog box 
   
})
$form.Controls.Add($exportButton)

# Create a button to fetch user groups
$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(60, 180)
$button.Size = New-Object System.Drawing.Size(180, 40)
$button.Text = "Get O365 Groups"
$button.Add_Click({
    $username = $textBox.Text
    
    
    #displays groups for BBB users 
    if($comboBox.SelectedItem -eq "BBB")
    {
    try{$membership = Get-AzureADUserMembership -ObjectId $username@british-business-bank.co.uk | Where-Object {$_.ObjectType -eq "Group"} | Select-Object DisplayName}
    catch 
    {
    [System.Windows.Forms.MessageBox]::Show("Invalid username or error occurred.", "Error")
    }
    $result = "Groups for user: $($username)`n" + ($membership.DisplayName -join "`n")
    [System.Windows.Forms.MessageBox]::Show($result, "AzureAD User Membership")
    }

    else
    {
    try{$membership = Get-AzureADUserMembership -ObjectId $username@startuploans.co.uk | Where-Object {$_.ObjectType -eq "Group"} | Select-Object DisplayName}
    catch 
    {
    [System.Windows.Forms.MessageBox]::Show("Invalid username or error occurred.", "Error")
    }
    $result = "Groups for user: $($username)`n" + ($membership.DisplayName -join "`n")
    [System.Windows.Forms.MessageBox]::Show($result, "AzureAD User Membership")
    }   
})
$form.Controls.Add($button)

#create button to fetch Distribution Groups
$DLbutton = New-Object System.Windows.Forms.Button
$DLbutton.Location = New-Object System.Drawing.Point(60,260)
$DLbutton.Size = New-Object System.Drawing.Size(180, 40)
$DLbutton.Text = "Export DLs"
$DLbutton.Add_Click({
    $username = $textBox.Text
    if ($comboBox.SelectedItem -eq "BBB") {
        
        Get-BBBUserDLExport -username $username     
    }

    else
    {
        Get-SulcoUserDLExport -username $username
    }
})
$form.controls.Add($DLbutton)

#create button to fetch shared mailboxes
$SMbutton = New-Object System.Windows.Forms.Button
$SMbutton.Location = New-Object System.Drawing.Point(350,260)
$SMbutton.Size = New-Object System.Drawing.Size(180, 40)
$SMbutton.Text = "Export Shared Mailboxes"
$SMbutton.Add_Click({
    $username =$textBox.Text
    if($comboBox.SelectedItem -eq "BBB"){
        Get-BBBSharedMailbox -username $username
    }

    else 
    {
        Get-SulcoMailbox -username $username
    }

})
$form.controls.Add($SMbutton)


# Show the form
$form.ShowDialog() | Out-Null

#Disconnect-AzureAD
#Disconnect-ExchangeOnline

