function BBB-DisabledOU{
  param ([string]$username)
            $operationSuccess = $false
           
            try {$accountName = Get-ADUser -Filter {SamAccountname -eq $username}
             
              $userEmail = $accountName.UserPrincipalName
              $userObject = Get-AzureADUser -ObjectId $userEmail
              $userObjectId = $userObject.ObjectId
            }

            catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException] {
            # Handle the specific authentication exception
             [System.Windows.Forms.MessageBox]::Show("You must authenticate with AzureAD before running this action", "ERROR_AZUREAD_AUTHENTICATION_REQUIRED_105")
            }
            catch {
                        Write-Error "An unexpected error occurred: $($_.Exception.Message)"
                        throw
                  }


            # Define the paths for the source and destination containers
            #$sourcePath = "OU=SBSUsers,OU=Users,OU=MyBusiness,DC=CFEL,DC=local"

            $sourcePath = $accountName.DistinguishedName -replace "^CN=.*?,", ""
            $destinationPath = "OU=Disabled Users,OU=Users,OU=MyBusiness,DC=CFEL,DC=local"

            $object = Get-ADUser -Filter {SamAccountName -eq $accountName.SamAccountName} -SearchBase $sourcePath

            if ($object) {
                # Move the AD object to the destination container
                Move-ADObject -Identity $object.DistinguishedName -TargetPath $destinationPath
                Write-Host "Object moved successfully."
               
                # Logging to CSV
                $logPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\Logs\TestLog.csv" 
                $logEntry = "$username,Moved to Disabled OU,$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                Add-Content -Path $logPath -Value $logEntry

                $operationSuccess = $true
           
            } else {
                Write-Host "Object not found in the specified source path."
            }
}

function SULCO-DisabledOU{

param ([string]$username)
            
            $operationSuccess = $false
            $domainController = "startuploans.local"

             try {$accountName = Get-ADUser -Filter {SamAccountname -eq $username}
             
              $userEmail = $accountName.UserPrincipalName
              $userObject = Get-AzureADUser -ObjectId $userEmail
              $userObjectId = $userObject.ObjectId
            }

            catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException] {
            # Handle the specific authentication exception
             [System.Windows.Forms.MessageBox]::Show("You must authenticate with AzureAD before running this action", "ERROR_AZUREAD_AUTHENTICATION_REQUIRED_105")
            }
            catch {
                        Write-Error "An unexpected error occurred: $($_.Exception.Message)"
                        throw
                  }

            # Define the paths for the source and destination containers
            #$sourcePath = "OU=SBSUsers,OU=Users,OU=MyBusiness,DC=CFEL,DC=local"

            $sourcePath = $accountName.DistinguishedName -replace "^CN=.*?,", ""
            $destinationPath = "OU=Disabled Users,OU=RDS Users,OU=Users,OU=StartUpLoans,DC=startuploans,DC=local"

            $object = Get-ADUser -Filter {SamAccountName -eq $accountName.SamAccountName} -Server $domainController -SearchBase $sourcePath

            if ($object) {
                # Move the AD object to the destination container
                Move-ADObject -Identity $object.DistinguishedName -TargetPath $destinationPath -Server $domainController
                Write-Host "Object moved successfully."
               
                # Logging to CSV
                $logPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\Logs\TestLog.csv" 
                $logEntry = "$username,Moved to Disabled OU,$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                Add-Content -Path $logPath -Value $logEntry

                $operationSuccess = $true
           
            } else {
                Write-Host "Object not found in the specified source path."
            }
}