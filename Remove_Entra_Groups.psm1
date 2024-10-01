Function BBB-RemoveEntraGroups{
    
    param ([string]$username)
            $operationSuccess = $false

    #$username = 'Sarah.deRancourt'

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

    #store AAD groups in an array 
    try{ $membership = Get-AzureADUserMembership -ObjectId $userEmail | Where-Object {$_.ObjectType -eq "Group"} | Select-Object ObjectID}

    catch {[System.Windows.Forms.MessageBox]::Show("An error occurred: $($_.Exception.Message)", "Error")
            throw
          }         
           $counter = 0
           foreach ($group in $membership){
                
               try {
                    Remove-AzureADGroupMember -ObjectId $group.ObjectId -MemberId $userObjectId -ErrorAction Stop
                    # Logging to CSV
                    $logPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\Logs\TestLog.csv" 
                    $logEntry = "$username,Entra Groups Removed($group),$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                    Add-Content -Path $logPath -Value $logEntry
                }
                catch {
                    if ($_.Exception.Message -like "*Cannot Update a mail-enabled security groups and or distribution list*") {
                        # Ignore this specific error
                        # Optionally, log the error or perform other actions
                    }
                    elseif ($_.Exception.Message -like "*Insufficient privileges to complete the operation*") {
                        # Ignore this specific error
                        # Optionally, log the error or perform other actions
                    }
                    else {
                        # Re-throw other exceptions to handle them elsewhere
                        throw
                    }
                }



           $counter ++
           write-host $counter
        }
    write-host $membership 
           
}

Function SULCO-RemoveEntraGroups{
    
    param ([string]$username)
            $operationSuccess = $false

    #$username = 'leah.walland'

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

    #store AAD groups in an array 
    try{ $membership = Get-AzureADUserMembership -ObjectId $userEmail | Where-Object {$_.ObjectType -eq "Group"} | Select-Object ObjectID}

    catch {[System.Windows.Forms.MessageBox]::Show("An error occurred: $($_.Exception.Message)", "Error")
            throw
          }         
           $counter = 0
           foreach ($group in $membership){
                
                try {Remove-AzureADGroupMember -ObjectId $group.ObjectId -MemberId $userObjectId -ErrorAction Stop}

                catch {
                    if ($_.Exception.Message -like "*Cannot Update a mail-enabled security groups and or distribution list*") {
                        # Ignore this specific error
                        # Optionally, you can log the error or perform other actions
                    }
                    else {
                        # Re-throw other exceptions to handle them elsewhere
                        throw
                    }
                        }


           $counter ++
           write-host $counter
        }
    write-host $membership 
    #Turn on automatic replies
}