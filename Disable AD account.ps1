function Disable-ADAccountAndCheck {
    param (
        [string]$accountName
    )
    
   # Attempt to disable the account
    try {
        Write-Output "Attempting to disable account: $accountName"
         #Check if account is already disabled
            $account = Get-ADUser -Identity $accountName -Property "userAccountControl" -ErrorAction Stop
            if (($account.userAccountControl -band 2) -eq 2) {
            [System.Windows.Forms.MessageBox]::Show("Account Already Disabled", "Warning")
            return
            }

            else{
            Disable-ADAccount -Identity $accountName -ErrorAction Stop
            }      
    } catch {
        Write-Warning "Failed to disable account: $_"
        return
    }

    # Wait for AD replication
    Start-Sleep -Seconds 5

    # Retrieve the account status
    try {
        $account = Get-ADUser -Identity $accountName -Property "userAccountControl" -ErrorAction Stop
    } catch {
        Write-Warning "Failed to retrieve account status: $_"
        return
    }

    # Check the ACCOUNTDISABLE flag (value 2)    
    if (($account.userAccountControl -band 2) -eq 2) {
        Write-Host "Account has been disabled"
        $status = "Disabled"
    } else {
        Write-Host "Account is still enabled"
        $status = "Enabled"
    }

    # Logging to CSV
    $logPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\Logs\TestLog.csv"
    $logEntry = "$accountName,$status,$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Add-Content -Path $logPath -Value $logEntry

    Start-Sleep -Seconds 5

}
# Example usage:
Disable-ADAccountAndCheck -accountName "simran.haroonraja"
