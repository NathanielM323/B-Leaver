[System.Threading.Thread]::CurrentThread.ApartmentState = "STA"
#Install-Module AzureAD
#connect-exchangeonline
try {Get-AzureADCurrentSessionInfo}

catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException]
{
Connect-AzureAD
}
#function to retrieve BBB users 
function Get-BBBUserGroupsExport {
    param ([String]$username)
    $operationSuccess = $false  # Initialize a flag for operation success

    try {
        # Retrieve the account name
        $accountName = Get-ADUser -Filter {SamAccountname -eq $username}

        if ($null -eq $accountName) {
            throw "The username '$username' is invalid or does not exist."
        }

        # Retrieve groups from AAD
        $membership = Get-AzureADUserMembership -ObjectId $accountName.UserPrincipalName | Where-Object {$_.ObjectType -eq "Group"} | Select-Object DisplayName

        # Sort membership by alphabetical order
        $membership = $membership | Sort-Object DisplayName

        # Create Directory to save the CSV file
        $directoryPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\B-Leaver\Leaver Data\$username"
        if (-not (Test-Path -Path $directoryPath)) {
            New-Item -Path $directoryPath -ItemType "directory" | Out-Null
        }

        # Define CSV file path with the username in the filename
        $csvFilePath = "$directoryPath\($username) Group memberships.csv"
        $membership | Export-Csv -Path $csvFilePath -NoTypeInformation

        [System.Windows.Forms.MessageBox]::Show("Group memberships exported as CSV to $directoryPath successfully.", "Export Complete")
        $operationSuccess = $true  # Set the flag to true if the operation succeeds
    } catch {
        if ($_.Exception.Message -like "*invalid or does not exist*") {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "ERROR_INVALID_USERNAME_012")
        } else {
            [System.Windows.Forms.MessageBox]::Show("An error occurred: $($_.Exception.Message)", "Error")
        }
        throw  # Re-throw the exception to propagate it to the main script
    }

    return $operationSuccess  # Return the success flag
    start-sleep -seconds 15
}




    function Get-SulcoUserGroupsExport {
    param ([String]$username)
    $operationSuccess = $false  # Initialize a flag for operation success

    try {
        # Attempt to retrieve groups from AAD   
        $membership = Get-AzureADUserMembership -ObjectId "$username@startuploans.co.uk" | Where-Object {$_.ObjectType -eq "Group"} | Select-Object DisplayName

        # Sort membership by alphabetical order
        $membership = $membership | Sort-Object DisplayName

        # Create directory to save the CSV file
        $directoryPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\B-Leaver\Leaver Data\$username"
        if (-not (Test-Path -Path $directoryPath)) {
            New-Item -Path $directoryPath -ItemType "directory" | Out-Null
        }

        # Define CSV file path with the username in the filename
        $csvFilePath = "$directoryPath\($username) Group memberships.csv"

        # Export memberships to CSV
        $membership | Export-Csv -Path $csvFilePath -NoTypeInformation
        [System.Windows.Forms.MessageBox]::Show("Group memberships exported as CSV to $directoryPath successfully.", "Export Complete")
        $operationSuccess = $true  # Set the flag to true if the operation succeeds

    } catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException] {
        [System.Windows.Forms.MessageBox]::Show("You must connect to AzureAD/Identity before calling Azure commands", "ERROR_AZURE_AUTHENTICATION_REQUIRED_101")
        throw  # Re-throw the exception to propagate it to the main script

    } catch {
        if ($_.Exception.Message -like "*Cannot bind argument to parameter 'ObjectId' because it is null.*") {
            [System.Windows.Forms.MessageBox]::Show("The username '$username' is invalid or does not exist.", "ERROR_INVALID_USERNAME_012")
        } else {
            [System.Windows.Forms.MessageBox]::Show("An error occurred: $($_.Exception.Message)", "Error")
        }
        throw  # Re-throw the exception to propagate it to the main script
    }

    return $operationSuccess  # Return the success flag
}


 function Get-BBBUserDLExport {
    param (
        [String]$username,
        [System.Windows.Forms.ProgressBar]$progressBar,
        [System.Windows.Forms.Label]$statusLabel  # Added the status label as a parameter
    )

    $operationSuccess = $false  # Initialize a flag for operation success

    try {
        # Retrieve the account name and email address
        $accountName = Get-ADUser -Filter {SamAccountname -eq $username} -Properties EmailAddress
        if ($null -eq $accountName) {
            throw "The username '$username' is invalid or does not exist."
        }
        $BBBaddress = $accountName.EmailAddress

        # Update status label
        $statusLabel.Text = "Status: Fetching Distribution Lists for $username..."
        
        # Retrieve all distribution groups
        $allDistributionGroups = Get-DistributionGroup -ResultSize Unlimited
        $totalDLs = $allDistributionGroups.Count
        $progress = 0

        # Array to store DLs the user has access to
        $DLswithAccess = @()

        # Loop through each DL to see if the user has access
        foreach ($DL in $allDistributionGroups) {
            $members = Get-DistributionGroupMember -Identity $DL.Identity
            if ($members.PrimarySmtpAddress -contains $BBBaddress) {
                $DLswithAccess += $DL.PrimarySmtpAddress
            }
            $progress++

            # Update progress bar
            $percentComplete = ($progress / $totalDLs) * 100
            $progressBar.Value = [math]::Round($percentComplete)

            # Optionally update status with progress
            $statusLabel.Text = "Status: Fetching Distribution Lists - $progress out of $totalDLs..."
        }

        # Create directory to save the CSV file
        $directoryPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\B-Leaver\Leaver Data\$username"
        if (-not (Test-Path -Path $directoryPath)) {
            New-Item -Path $directoryPath -ItemType "directory" | Out-Null
        }

        # Define CSV file path with the username in the filename
        $csvFilePath = "$directoryPath\($username) Distribution_Lists.csv"

        # Create a custom object for each distribution list and export to CSV
        $DLObjects = $DLswithAccess | ForEach-Object {
            [PSCustomObject]@{
                DistributionList = $_
            }
        }
        $DLObjects | Export-Csv -Path $csvFilePath -NoTypeInformation

        $operationSuccess = $true  # Set the flag to true if the operation succeeds

    } catch [System.Management.Automation.CommandNotFoundException] {
        if ($_.Exception.CommandName -in @('Get-DistributionGroup', 'Get-Mailbox', 'Get-DistributionGroupMember')) {
            # Handle error if the Exchange module is not connected
            [System.Windows.Forms.MessageBox]::Show("The Exchange Online module is not connected. Please connect to Exchange Online and try again.", "Exchange Online Connection Error")
            Write-Error "The Exchange Online PowerShell module is not loaded or not connected. Please ensure you've connected to Exchange Online before running this script."
        } else {
            throw
        }
    } catch {
        if ($_.Exception.Message -like "*invalid or does not exist*") {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "ERROR_INVALID_USERNAME_012")
        } else {
            [System.Windows.Forms.MessageBox]::Show("An error occurred: $($_.Exception.Message)", "Error")
        }
        throw  # Re-throw the exception to propagate it to the main script
    } finally {
        if ($operationSuccess) {
            $statusLabel.Text = "Status: Completed fetching Distribution Lists for $username."
        } else {
            $statusLabel.Text = "Status: Failed to fetch Distribution Lists for $username."
        }
    }
    
    start-sleep -seconds 15
    return $operationSuccess  # Return the success flag
}

   
function Get-SulcoUserDLExport {
    param (
        [String]$username,
        [System.Windows.Forms.ProgressBar]$progressBar,
        [System.Windows.Forms.Label]$statusLabel
    )

    try {
        # Update status label before starting
        $statusLabel.Text = "Status: Fetching Distribution Lists for $username..."
        
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

            # Update progress bar
            $percentComplete = ($progress / $totalDLs) * 100
            $progressBar.Value = [math]::Round($percentComplete)

            # Optionally update status with progress
            $statusLabel.Text = "Status: Fetching Distribution Lists - $progress out of $totalDLs..."
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

    } catch [System.Management.Automation.CommandNotFoundException] {
        if ($_.Exception.CommandName -eq 'Get-DistributionGroup') {
            Write-Error "The Exchange Online PowerShell module is not loaded or not connected. Please ensure you've connected to Exchange Online before running this script."
        } else {
            throw
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error occurred: $($_.Exception.Message)", "Error")
    } finally {
        # Update status label after completion
        $statusLabel.Text = "Status: Completed fetching Distribution Lists for $username."
    }
}


function Get-BBBSharedMailbox {
    param (
        [String]$username,
        [System.Windows.Forms.ProgressBar]$progressBar,
        [System.Windows.Forms.Label]$statusLabel
    )
    $accountName = Get-ADUser -Filter {SamAccountname -eq $username} -Properties EmailAddress
    $BBBaddress = $accountName.EmailAddress
    write-host $BBBaddress
    try {
        # Update status label before starting
        $statusLabel.Text = "Status: Fetching Shared Mailbox Permissions for $username..."

        # Get all mailboxes
        $AllMailboxes = Get-Mailbox -ResultSize Unlimited
        $totalMailboxes = $AllMailboxes.Count
        $progress = 0

        # Array to store mailboxes the user has access to
        $MailboxesWithAccess = @()

        # Loop through each mailbox and check if the user has access
        foreach ($Mailbox in $AllMailboxes) {
            $MailboxPermissions = Get-MailboxPermission -Identity $Mailbox.DistinguishedName | Where-Object { $_.User -like $BBBaddress -and $_.AccessRights -like "FullAccess" }
            if ($MailboxPermissions) {
                $MailboxesWithAccess += $Mailbox
            }

            # Update progress
            $progress++
            $percentComplete = ($progress / $totalMailboxes) * 100
            $progressBar.Value = [math]::Round($percentComplete)

            # Optionally update status with progress
            $statusLabel.Text = "Status: Checking Mailbox Permissions - $progress out of $totalMailboxes..."
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

    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error occurred: $($_.Exception.Message)", "Error")
    } finally {
        # Update status label after completion
        $statusLabel.Text = "Status: Completed fetching Shared Mailbox Permissions for $username."
        write-host "mailboxes retrieved"
        start-sleep -seconds 15
    }
}

    function Get-SulcoMailbox {
    param (
        [String]$username,
        [System.Windows.Forms.ProgressBar]$progressBar,
        [System.Windows.Forms.Label]$statusLabel
    )
    
    $sulcoAddress = "$username@startuploans.co.uk"        

    try {
        # Update status label before starting
        $statusLabel.Text = "Status: Fetching Shared Mailbox Permissions for $username..."

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

            # Update progress bar and status label
            $progress++
            $percentComplete = ($progress / $totalMailboxes) * 100
            $progressBar.Value = [math]::Round($percentComplete)

            # Optionally update status with progress
            $statusLabel.Text = "Status: Checking Mailbox Permissions - $progress out of $totalMailboxes..."
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

    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error occurred: $($_.Exception.Message)", "Error")
    } finally {
        # Update status label after completion
        $statusLabel.Text = "Status: Completed fetching Shared Mailbox Permissions for $username."
    }
}

