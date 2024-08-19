function GALHideBBB{
param ([String]$username)

Set-ADUser -Identity $username -Replace @{msExchHideFromAddressLists = $true}
write-host "User hidden from GAL"
#check if user is hidden from GAL
$user = Get-ADUser -Filter{SamAccountname -eq $username} -Properties msExchHideFromAddressLists

if ($user.msExchHideFromAddressLists -eq $true) {
    [System.Windows.Forms.MessageBox]::Show("$username hidden from GAL", "Successfully hidden from GAL")
} else {
    [System.Windows.Forms.MessageBox]::Show("Error occured when hiding $username from GAL", "Error hiding from GAL")
}

} 

function GALHideSULCO {
param ([String]$username)
$domainController = "startuploans.local"

Set-ADUser -Identity $accountName.SamAccountName -Replace @{msExchHideFromAddressLists = $true} -Server $domainController
#check if user is hidden from GAL
$user = Get-ADUser -Filter{SamAccountname -eq $accountName.SamAccountName} -Properties msExchHideFromAddressLists -Server $domainController

if ($user.msExchHideFromAddressLists -eq $true) {
     [System.Windows.Forms.MessageBox]::Show("$username hidden from GAL", "Successfully hidden from GAL")
} else {
     [System.Windows.Forms.MessageBox]::Show("Error occured when hiding $username from GAL", "Error hiding from GAL")
}
}