function GALHideBBB {
    param ([String]$username)
    $operationSuccess = $false  # Initialize a flag for operation success

    try {
        # Attempt to hide the user from the Global Address List
        Set-ADUser -Identity $username -Replace @{msExchHideFromAddressLists = $true}

        # Check if the user is successfully hidden from GAL
        $user = Get-ADUser -Filter {SamAccountname -eq $username} -Properties msExchHideFromAddressLists

        if ($user.msExchHideFromAddressLists -eq $true) {
            Write-Host "$username is hidden from address lists."
            $operationSuccess = $true  # Set the flag to true if the operation succeeds
        } else {
            throw "Failed to hide $username from the Global Address List."
        }

    } catch {
        # Handle any errors that occur during the process
        Write-Host "An error occurred while trying to hide $username from the GAL: $($_.Exception.Message)"
        throw  # Re-throw the exception to be caught by the main script
    }

    return $operationSuccess  # Return the success flag
}




function GALHideSULCO {
param ([String]$username)
$domainController = "startuploans.local"

write-Host $username , $domainController

Set-ADUser -Identity $username -Replace @{msExchHideFromAddressLists = $true} -Server $domainController
#check if user is hidden from GAL
$user = Get-ADUser -Filter{SamAccountname -eq $username} -Properties msExchHideFromAddressLists -Server $domainController

if ($user.msExchHideFromAddressLists -eq $true) {
     [System.Windows.Forms.MessageBox]::Show("$username hidden from GAL", "Successfully hidden from GAL")
} else {
     [System.Windows.Forms.MessageBox]::Show("Error occured when hiding $username from GAL", "Error hiding from GAL")
}
}