function BBBDisable-ADAccount {
    param (
        [string]$username)
        $operationSuccess = $false  # Initialize a flag for operation success
    
   # Attempt to disable the account
    try {
        Write-Output "Attempting to disable account: $username"
         #Check if account is already disabled
            $account = Get-ADUser -Identity $username -Property "userAccountControl" -ErrorAction Stop
            if (($account.userAccountControl -band 2) -eq 2) {
            [System.Windows.Forms.MessageBox]::Show("Account Already Disabled", "Warning")
            return
            }

            else{
             $statusLabel.Text = "Status: Disabling AD Account..."
            Disable-ADAccount -Identity $account -ErrorAction Stop        
            $operationSuccess = $true
            }      
    } catch {
        Write-Warning "Failed to disable account: $_"
        [System.Windows.Forms.MessageBox]::Show("Failed to disable account: $_", "Warning")
        throw
        return
    }

    # Wait for AD replication
    Start-Sleep -Seconds 5

    # Retrieve the account status
    try {
        $account = Get-ADUser -Identity $account -Property "userAccountControl" -ErrorAction Stop
    } catch {
        Write-Warning "Failed to retrieve account status: $_"
        return
    }

    # Check the ACCOUNTDISABLE flag (value 2)    
    if (($account.userAccountControl -band 2) -eq 2) {
        Write-Host "Account has been disabled"
        $status = "Account Disabled"
    } else {
        Write-Host "Account is still enabled"
        $status = "Account Enabled"
    }

    # Logging to CSV
    $logPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\Logs\TestLog.csv"
    $logEntry = "$username,$status,$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Add-Content -Path $logPath -Value $logEntry

    Start-Sleep -Seconds 5
    return $operationSuccess
}

function SULCODisable-ADAccount {
    param (
        [string]$username)
$operationSuccess = $false  # Initialize a flag for operation success
$domainController = "startuploans.local"
    
   # Attempt to disable the account
    try {
        Write-Output "Attempting to disable account: $username"
         #Check if account is already disabled
            $account = Get-ADUser -Identity $username -Server $domainController -Property "userAccountControl" -ErrorAction Stop
            if (($account.userAccountControl -band 2) -eq 2) {
            [System.Windows.Forms.MessageBox]::Show("Account Already Disabled", "Warning")
            return
            }

            else{
             $statusLabel.Text = "Status: Disabling AD Account..."
            Disable-ADAccount -Identity $account -ErrorAction Stop        
            $operationSuccess = $true
            }      
    } catch {
        Write-Warning "Failed to disable account: $_"
        [System.Windows.Forms.MessageBox]::Show("Failed to disable account: $_", "Warning")
        throw
        return
    }

    # Wait for AD replication
    Start-Sleep -Seconds 5

    # Retrieve the account status
    try {
        $account = Get-ADUser -Identity $account -Property "userAccountControl" -ErrorAction Stop
    } catch {
        Write-Warning "Failed to retrieve account status: $_"
        return
    }

    # Check the ACCOUNTDISABLE flag (value 2)    
    if (($account.userAccountControl -band 2) -eq 2) {
        Write-Host "Account has been disabled"
        $status = "Account Disabled"
    } else {
        Write-Host "Account is still enabled"
        $status = "Account Enabled"
    }

    # Logging to CSV
    $logPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\Logs\TestLog.csv"
    $logEntry = "$username,$status,$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Add-Content -Path $logPath -Value $logEntry

    Start-Sleep -Seconds 5
    return $operationSuccess

}
