function BBB-OOO {
param ([string]$username)
        $operationSuccess = $false

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

try {$autoReplyConfig = Get-MailboxAutoReplyConfiguration -Identity $username@british-business-bank.co.uk}

catch [System.Management.Automation.CommandNotFoundException] {
                Write-host "The Set-MailboxAutoReplyConfiguration cmdlet was not found. Ensure you have the necessary Exchange management tools installed."
                [System.Windows.Forms.MessageBox]::Show("Ensure you are connected to MS Exchange online", "ERROR_MSE_AUTHENTICATION_REQUIRED_104")
                }

if($autoReplyConfig.AutoReplyState -eq "Enabled"){
    $InternalMessage = $autoReplyConfig.InternalMessage
    
    function Remove-HTMLTags {
        param([string]$html)
        $text = $html -replace '<[^>]+>', ''
        $text = [System.Web.HttpUtility]::HtmlDecode($text)
        return $text.Trim()
    }
    
    $cleanMessage = Remove-HTMLTags -html $InternalMessage
    
    $OOOform = New-Object System.Windows.Forms.Form
    $OOOform.Text = "Out of Office Reply"
    $OOOform.StartPosition = "CenterScreen"
    
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Internal Out of Office Message:`n`n$cleanMessage"
    $label.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
    $label.AutoSize = $true
    $label.MaximumSize = New-Object System.Drawing.Size(460, 0)
    $label.Location = New-Object System.Drawing.Point(20, 20)

    #Add button to leave message as is
    $stayButton = New-Object System.Windows.Forms.Button
    $stayButton.Location = New-Object System.Drawing.Point(50, 170)
    $stayButton.Size = New-Object System.Drawing.Size(180, 40)
    $stayButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
    $stayButton.Text = "Leave as is"

    #Add button to replace OOO mesage with template
    $templateButton = New-Object System.Windows.Forms.Button
    $templateButton.Location = New-Object System.Drawing.Point(320, 170)
    $templateButton.Size = New-Object System.Drawing.Size(180, 40)
    $templateButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
    $templateButton.Text = "Replace with template"
    
    $OOOform.Controls.Add($label)
    $OOOform.Controls.Add($stayButton)
    $OOOform.Controls.Add($templateButton)
    
    $stayButton.Add_Click({
         $OOOform.Close()
    })

    $templateButton.Add_click({
        try { Set-MailboxAutoReplyConfiguration -Identity $username@british-business-bank.co.uk -AutoReplyState Enabled -InternalMessage "Dear Sir / Madam,

            I have now left the British Business Bank. If your query is urgent, please contact another member of my team.

            Kind Regards." -ExternalMessage "Dear Sir / Madam,

            I have now left the British Business Bank. If your query is urgent, please contact another member of my team.

            Kind Regards."
            
            $operationSuccess = $true

            # Logging to CSV
            $logPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\Logs\TestLog.csv"
            $logEntry = "$username,Automatic Reply Enabled,$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            Add-Content -Path $logPath -Value $logEntry
            }

         catch [System.Management.Automation.ParameterBindingException] 
                {
                Write-Error "There was an issue with one of the parameters. Check that the username and email address are correct."
                }

         catch {
                Write-Error "An unexpected error occurred: $($_.Exception.Message)"
                throw
               }

                $OOOform.Close()

    })
    
    # Calculate form size based on label size
    $formWidth = [Math]::Min(500, $label.PreferredWidth + 60)  # Max width of 500, or label width + padding
    $formHeight = $label.PreferredHeight + 130  # Label height + padding
    $OOOform.ClientSize = New-Object System.Drawing.Size($formWidth, $formHeight)
    
    $OOOform.ShowDialog()
    }

    else {
        try { Set-MailboxAutoReplyConfiguration -Identity $username@british-business-bank.co.uk -AutoReplyState Enabled -InternalMessage "Dear Sir / Madam,

            I have now left the British Business Bank. If your query is urgent, please contact another member of my team.

            Kind Regards." -ExternalMessage "Dear Sir / Madam,

            I have now left the British Business Bank. If your query is urgent, please contact another member of my team.

            Kind Regards."
            
            $operationSuccess = $true

            # Logging to CSV
            $logPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\Logs\TestLog.csv"
            $logEntry = "$username,Automatic Reply Enabled,$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            Add-Content -Path $logPath -Value $logEntry
            }

         catch [System.Management.Automation.ParameterBindingException] 
                {
                Write-Error "There was an issue with one of the parameters. Check that the username and email address are correct."
                }

         catch {
                Write-Error "An unexpected error occurred: $($_.Exception.Message)"
                throw
               }
     return $operationSuccess  # Return the success flag
    }
}

function SULCO-OOO {
param ([string]$username)
        $operationSuccess = $false

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

try {$autoReplyConfig = Get-MailboxAutoReplyConfiguration -Identity $username@startuploans.co.uk}

catch [System.Management.Automation.CommandNotFoundException] {
                Write-host "The Set-MailboxAutoReplyConfiguration cmdlet was not found. Ensure you have the necessary Exchange management tools installed."
                [System.Windows.Forms.MessageBox]::Show("Ensure you are connected to MS Exchange online", "ERROR_MSE_AUTHENTICATION_REQUIRED_104")
                }

if($autoReplyConfig.AutoReplyState -eq "Enabled"){
    $InternalMessage = $autoReplyConfig.InternalMessage
    
    function Remove-HTMLTags {
        param([string]$html)
        $text = $html -replace '<[^>]+>', ''
        $text = [System.Web.HttpUtility]::HtmlDecode($text)
        return $text.Trim()
    }
    
    $cleanMessage = Remove-HTMLTags -html $InternalMessage
    
    $OOOform = New-Object System.Windows.Forms.Form
    $OOOform.Text = "Out of Office Reply"
    $OOOform.StartPosition = "CenterScreen"
    
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Internal Out of Office Message:`n`n$cleanMessage"
    $label.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
    $label.AutoSize = $true
    $label.MaximumSize = New-Object System.Drawing.Size(460, 0)
    $label.Location = New-Object System.Drawing.Point(20, 20)

    #Add button to leave message as is
    $stayButton = New-Object System.Windows.Forms.Button
    $stayButton.Location = New-Object System.Drawing.Point(50, 170)
    $stayButton.Size = New-Object System.Drawing.Size(180, 40)
    $stayButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
    $stayButton.Text = "Leave as is"

    #Add button to replace OOO mesage with template
    $templateButton = New-Object System.Windows.Forms.Button
    $templateButton.Location = New-Object System.Drawing.Point(320, 170)
    $templateButton.Size = New-Object System.Drawing.Size(180, 40)
    $templateButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
    $templateButton.Text = "Replace with template"
    
    $OOOform.Controls.Add($label)
    $OOOform.Controls.Add($stayButton)
    $OOOform.Controls.Add($templateButton)
    
    $stayButton.Add_Click({
         $OOOform.Close()
    })

    $templateButton.Add_click({
        try { Set-MailboxAutoReplyConfiguration -Identity $username@startuploans.co.uk -AutoReplyState Enabled -InternalMessage "Dear Madam / Sir,

            I have now left the British Business Bank. If your query is urgent, please contact another member of my team.

            Kind Regards." -ExternalMessage "Dear Sir / Madam,

            I have now left the British Business Bank. If your query is urgent, please contact another member of my team.

            Kind Regards."
            
            $operationSuccess = $true

            # Logging to CSV
            $logPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\Logs\TestLog.csv"
            $logEntry = "$username,Automatic Reply Enabled,$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            Add-Content -Path $logPath -Value $logEntry
            }

         catch [System.Management.Automation.ParameterBindingException] 
                {
                Write-Error "There was an issue with one of the parameters. Check that the username and email address are correct."
                }

         catch {
                Write-Error "An unexpected error occurred: $($_.Exception.Message)"
                throw
               }

                $OOOform.Close()

    })
    
    # Calculate form size based on label size
    $formWidth = [Math]::Min(500, $label.PreferredWidth + 60)  # Max width of 500, or label width + padding
    $formHeight = $label.PreferredHeight + 130  # Label height + padding
    $OOOform.ClientSize = New-Object System.Drawing.Size($formWidth, $formHeight)
    
    $OOOform.ShowDialog()
    }

    else {
        try { Set-MailboxAutoReplyConfiguration -Identity $username@startuploans.co.uk -AutoReplyState Enabled -InternalMessage "Dear Madam / Sir,

            I have now left the British Business Bank. If your query is urgent, please contact another member of my team.

            Kind Regards." -ExternalMessage "Dear Madam / Sir,

            I have now left the British Business Bank. If your query is urgent, please contact another member of my team.

            Kind Regards."
            
            $operationSuccess = $true

            # Logging to CSV
            $logPath = "\\cfel.local\dfsroot\group\ICT\Nathaniel\Leaver\Leaver Automation\Logs\TestLog.csv"
            $logEntry = "$username,Automatic Reply Enabled,$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            Add-Content -Path $logPath -Value $logEntry
            }

         catch [System.Management.Automation.ParameterBindingException] 
                {
                Write-Error "There was an issue with one of the parameters. Check that the username and email address are correct."
                }

         catch {
                Write-Error "An unexpected error occurred: $($_.Exception.Message)"
                throw
               }
     return $operationSuccess  # Return the success flag
    }
}