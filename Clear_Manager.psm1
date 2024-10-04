 function BBB-ClearManager{
 param ([string]$username)
         

         $operationSuccess = $false

            try {$accountName = Get-ADUser -Filter {SamAccountname -eq $username}
             
                  $userEmail = $accountName.UserPrincipalName
                  $userObject = Get-AzureADUser -ObjectId $userEmail
                  $userObjectId = $userObject.ObjectId
            }
            
            catch {
                  Write-Error "An unexpected error occurred: $($_.Exception.Message)"
                  throw
                  }
            
            #Try statement to clear the manager field
            try{Set-ADUser -Identity $accountname -Clear manager
                 # Logging to CSV
                $logPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\Logs\TestLog.csv" 
                $logEntry = "$username,Manager field cleared,$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                Add-Content -Path $logPath -Value $logEntry

                $operationSuccess = $true
            }

            catch 
            {
                 [System.Windows.Forms.MessageBox]::Show($($_.Exception.Message))
            }

            return $operationSuccess
}

 function SULCO-ClearManager{
 param ([string]$username)
         
         $domainController = "startuploans.local"
         $operationSuccess = $false

            try {$accountName = Get-ADUser -Filter {SamAccountname -eq $username}
                             
                  $userEmail = $accountName.UserPrincipalName
                  $userObject = Get-AzureADUser -ObjectId $userEmail
                  $userObjectId = $userObject.ObjectId
            }
            
            catch {
                  Write-Error "An unexpected error occurred: $($_.Exception.Message)"
                  throw
                  }
            
            #Try statement to clear the manager field
            try{Set-ADUser -Server $domainController  -Identity $accountName.SamAccountName -Clear manager
            
            # Logging to CSV
                $logPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\Logs\TestLog.csv" 
                $logEntry = "$username,Manager field cleared,$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                Add-Content -Path $logPath -Value $logEntry

                   $operationSuccess = $true
            }
                            
            catch 
            {
                 [System.Windows.Forms.MessageBox]::Show($($_.Exception.Message))
            }

            return $operationSuccess
}
