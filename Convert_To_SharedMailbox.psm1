function BBB-ConvertMailbox{
    param ([string]$username)
         
    $operationSuccess = $false

     try {$accountName = Get-ADUser -Filter {SamAccountname -eq $username}
             
               $userEmail = $accountName.UserPrincipalName
               $userObject = Get-AzureADUser -ObjectId $userEmail
               $userObjectId = $userObject.ObjectId

               $mailbox = Get-mailbox -Identity $userEmail
         }

     catch {
             Write-Error "An unexpected error occurred: $($_.Exception.Message)"
             throw
           }

     try {Set-Mailbox -Identity $mailbox -Type Shared
             
              start-sleep -Seconds 30
              # Logging to CSV
                $logPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\Logs\TestLog.csv" 
                $logEntry = "$username,Converted to Shared Mailbox,$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                Add-Content -Path $logPath -Value $logEntry

                $operationSuccess = $true
         }

     catch 
            {
                 [System.Windows.Forms.MessageBox]::Show($($_.Exception.Message))
            }

            return $operationSuccess
} 

function SULCO-ConvertMailbox{
    param ([string]$username)
         
    $domainController = "startuploans.local"
    $operationSuccess = $false
    $username = 'samina.khan'

     try {$accountName = Get-ADUser -Filter {SamAccountname -eq $username} -Server $domainController
             
               $userEmail = $accountName.UserPrincipalName
               $userObject = Get-AzureADUser -ObjectId $userEmail
               $userObjectId = $userObject.ObjectId

               $mailbox = Get-mailbox -Identity $userEmail
         }

     catch {
             Write-Error "An unexpected error occurred: $($_.Exception.Message)"
             throw
           }

     try {Set-Mailbox -Identity $mailbox -Type Shared
             
              start-sleep -Seconds 30
              # Logging to CSV
                $logPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\Logs\TestLog.csv" 
                $logEntry = "$username,Converted to Shared Mailbox,$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                Add-Content -Path $logPath -Value $logEntry

                $operationSuccess = $true
         }

     catch 
            {
                 [System.Windows.Forms.MessageBox]::Show($($_.Exception.Message))
            }

            return $operationSuccess
} 