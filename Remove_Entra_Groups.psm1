Function BBB-RemoveEntraGroups{
    
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
         [System.Windows.Forms.MessageBox]::Show("You must call the Connect-AzureAD cmdlet before calling any other cmdlets.", "ERROR_AZUREAD_AUTHENTICATION_REQUIRED_105")
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
           Remove-AzureADGroupMember -ObjectId $group.ObjectId -MemberId $userObjectId
           $counter ++
           write-host $counter
        }
    write-host $membership 
    #Turn on automatic replies
}