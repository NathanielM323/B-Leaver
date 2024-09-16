function BBB-RemoveADGroups{
param ([string]$username)
        $operationSuccess = $false

        #$username = 'George.Elysee'
            try{ $groups = Get-ADUser $username -Properties MemberOf |Select-Object -ExpandProperty MemberOf}

            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]{
                    [System.Windows.Forms.MessageBox]::Show("Error: The group '$group' or user '$username' was not found in AD.","ERROR_INVALID_USERNAME_OR_Group_013")
            }
            catch {
            if ($_.Exception.Message -like "*Cannot bind argument to parameter 'ObjectId' because it is null.*") {
                [System.Windows.Forms.MessageBox]::Show("The username '$username' is invalid or does not exist.", "ERROR_INVALID_USERNAME_012")
            } else {
                [System.Windows.Forms.MessageBox]::Show("An error occurred: $($_.Exception.Message)", "Error")
            }
            throw  # Re-throw the exception to propagate it to the main script
            }
                Write-Host $groups
        foreach ($group in $groups)
        {
        try {Remove-ADGroupMember -Identity $group -Members $username -confirm:$false}

        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]{
                    [System.Windows.Forms.MessageBox]::Show("Error: The group '$group' or user '$username' was not found in AD.","ERROR_INVALID_USERNAME_OR_Group_013")
            }
        }

        catch {
            if ($_.Exception.Message -like "*Cannot bind argument to parameter 'ObjectId' because it is null.*") {
                [System.Windows.Forms.MessageBox]::Show("The username '$username' is invalid or does not exist.", "ERROR_INVALID_USERNAME_012")
            } else {
                [System.Windows.Forms.MessageBox]::Show("An error occurred: $($_.Exception.Message)", "Error")
            }
            throw  # Re-throw the exception to propagate it to the main script
            }

        start-sleep -seconds 10

        }