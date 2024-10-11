cd '\\cfel.local\dfsroot\ICT\Nathaniel\Leaver\Leaver Automation\B-Leaver\B-Leaver(git)'

$currentDirectory = Get-Location
$modulePath = Join-Path $currentDirectory "Fetch_Leaver_data.psm1"

# Import 'Fetch_Leaver_data' using the module using the path
Import-Module $modulePath -Force
Get-Module -Name Fetch_Leaver_data

# Import 'Hide_From_GAL' using the module using the path
$modulePath = Join-Path $currentDirectory "Hide_From_GAL.psm1"
Import-Module $modulePath -Force
Get-Module -Name Hide_From_GAL

# Import 'Disable_AD_account' using the module using the path
$modulePath = Join-Path $currentDirectory "Disable_AD_account.psm1"
Import-Module $modulePath -Force
Get-Module -Name Disable_AD_Account

$modulePath = Join-Path $currentDirectory "OOO_Reply.psm1"
Import-Module $modulePath -Force
Get-Module -Name OOO_Reply

$modulePath = Join-Path $currentDirectory "Remove_AD_Groups.psm1"
Import-Module $modulePath -Force
Get-Module -Name Remove_AD_Groups

$modulePath = Join-Path $currentDirectory "Remove_Entra_Groups.psm1"
Import-Module $modulePath -Force
Get-Module -Name Remove_Entra_Groups

$modulePath = Join-Path $currentDirectory "Move_To_DisabledOU.psm1"
Import-Module $modulePath -Force
Get-Module -Name Move_To_DisabledOU

$modulePath = Join-Path $currentDirectory "Clear_Manager.psm1"
Import-Module $modulePath -Force
Get-Module -Name Clear_Manager

$modulePath = Join-Path $currentDirectory "Convert_To_SharedMailbox.psm1"
Import-Module $modulePath -Force
Get-Module -Name Convert_To_SharedMailbox



#--------------------------------------------------------

if (Get-Module -Name "Fetch_Leaver_data") {
    Write-Host "Module 'Fetch_Leaver_data' is successfully imported."
    Get-Command -Module "Fetch_Leaver_data" | ForEach-Object { Write-Host "Function available: $($_.Name)" }
} else {
    Write-Host "Module 'Fetch_Leaver_data' is not imported."
}

# Load Windows Forms Assembly
Add-Type -AssemblyName System.Windows.Forms

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "BLeaver"
$form.Size = New-Object System.Drawing.Size(1000, 1000)
$form.StartPosition = "CenterScreen"

# Add a label to display current process
$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Text = "Status: Idle"
$StatusLabel.AutoSize = $true
$StatusLabel.Location = New-Object System.Drawing.Point(50, 440)
$form.Controls.Add($StatusLabel)

# Add a progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(170,460)
$progressBar.Size = New-Object System.Drawing.Size(800, 30)
$form.Controls.Add($progressBar)

# Create a label for the username input
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(20, 100)
$label.Size = New-Object System.Drawing.Size(120, 20)
$label.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Regular)
$label.Text = "Username:"
$form.Controls.Add($label)

# Create a text box for username input
$userTextBox = New-Object System.Windows.Forms.TextBox
$userTextBox.Location = New-Object System.Drawing.Point(180, 95)
$userTextBox.Size = New-Object System.Drawing.Size(350, 60)
$userTextBox.Font = New-Object System.Drawing.Font("Arial", 20, [System.Drawing.FontStyle]::Regular)
$form.Controls.Add($userTextBox)
$username = $userTextBox.Text

# Create a drop-down menu (combo box)
$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(180, 50)
$comboBox.Size = New-Object System.Drawing.Size(250, 100)
$comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboBox.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
$comboBox.Items.Add("BBB") 
$comboBox.Items.Add("SULCO")
$form.Controls.Add($comboBox)

# Create a label for the drop-down menu
$Dropdownlabel = New-Object System.Windows.Forms.Label
$Dropdownlabel.Location = New-Object System.Drawing.Point(20, 50)
$Dropdownlabel.Size = New-Object System.Drawing.Size(90, 20)
$Dropdownlabel.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Regular)
$Dropdownlabel.Text = "Domain:"
$form.Controls.Add($Dropdownlabel)

# Add checkboxes
$checkboxOptions = @("Select All", "Retrieve O365 groups", "Retrieve Distribution lists", "Retrieve Shared Mailboxes" , "Hide from GAL" , "Disable AD Account", 
                    "Enable OOO Reply", "Remove AD groups" , "Remove Entra Groups", "Move to Disabled OU","Clear 'Manager' field","Convert to Shared Mailbox")
$checkboxes = @()
$startY = 160  # Starting Y position for the first checkbox
$indent = 23   # Indentation for sub-options

foreach ($index in 0..($checkboxOptions.Count - 1)) {
    $checkbox = New-Object System.Windows.Forms.CheckBox
    $checkbox.Text = $checkboxOptions[$index]
    
    if ($index -notin 1, 2, 3) {
        # Position for non-indented checkboxes
        $checkbox.Location = New-Object System.Drawing.Point(180, $startY)
    } else {
        # Position for indented checkboxes
        $checkbox.Location = New-Object System.Drawing.Point((180 + $indent), $startY)
    }

    
    $checkbox.Size = New-Object System.Drawing.Size(200, 20)
    $checkbox.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Regular)
    $form.Controls.Add($checkbox)
    $checkboxes += $checkbox
    $startY += 23  # Increment Y position for the next checkbox
}

# Add event handler for "Retrieve all leaver data" checkbox
$checkboxes[0].Add_CheckedChanged({
    if ($checkboxes[0].Checked) {
        # Check all other checkboxes
        foreach ($i in 1..($checkboxOptions.Count - 1)) {
            $checkboxes[$i].Checked = $true
        }
    } else {
        # Uncheck all other checkboxes
        foreach ($i in 1..($checkboxOptions.Count - 1)) {
            $checkboxes[$i].Checked = $false
        }
    }
})


# Create a function mapping that asscoiates the function with the checkbox name
$functionMap = @{
    "BBB" = @{
        "Retrieve O365 groups" = { param($username, $progressBar, $statusLabel) Get-BBBUserGroupsExport -username $username -progressBar $progressBar -statusLabel $statusLabel }
        "Retrieve Distribution lists" = { param($username, $progressBar, $statusLabel) Get-BBBUserDLExport -username $username -progressBar $progressBar -statusLabel $statusLabel }
        "Retrieve Shared Mailboxes" = { param($username, $progressBar, $statusLabel) Get-BBBSharedMailbox -username $username -progressBar $progressBar -statusLabel $statusLabel }
        "Hide from GAL" = {param($username) GALHideBBB -username $username}
        "Disable AD Account" = {param($username, $progressBar, $statusLabel) BBBDisable-ADAccount -username $username -progressBar $progressBar -statusLabel $statusLabel}
        "Enable OOO Reply" = {param($username, $progressBar, $statusLabel) BBB-OOO -username $username -progressBar $progressBar -statusLabel $statusLabel}
        "Remove AD groups" = {param($username, $progressBar, $statusLabel) BBB-RemoveADGroups -username $username -progressBar $progressBar -statusLabel $statusLabel}
        "Remove Entra Groups" = {param($username, $progressBar, $statusLabel) BBB-RemoveEntraGroups -username $username -progressBar $progressBar -statusLabel $statusLabel}
        "Move to Disabled OU" = {param($username, $progressBar, $statusLabel) BBB-DisabledOU -username $username -progressBar $progressBar -statusLabel $statusLabel}
        "Clear 'Manager' field" = {param($username, $progressBar, $statusLabel) BBB-ClearManager -username $username -progressBar $progressBar -statusLabel $statusLabel}
        "Convert to Shared Mailbox" = {param($username, $progressBar, $statusLabel) BBB-ConvertMailbox -username $username -progressBar $progressBar -statusLabel $statusLabel}
    }
    "SULCO" = @{
        "Retrieve O365 groups" = { param($username, $progressBar, $statusLabel) Get-SulcoUserGroupsExport -username $username -progressBar $progressBar -statusLabel $statusLabel }
        "Retrieve Distribution lists" = { param($username, $progressBar, $statusLabel) Get-SulcoUserDLExport -username $username -progressBar $progressBar -statusLabel $statusLabel }
        "Retrieve Shared Mailboxes" = { param($username, $progressBar, $statusLabel) Get-SulcoMailbox -username $username -progressBar $progressBar -statusLabel $statusLabel }
        "Hide from GAL" = {param($username) GALHideSULCO -username $username}
        "Disable AD Account" = {param($username, $progressBar, $statusLabel) SULCODisable-ADAccount -username $username -progressBar $progressBar -statusLabel $statusLabel}
        "Enable OOO Reply" = {param($username, $progressBar, $statusLabel) SULCO-OOO -username $username -progressBar $progressBar -statusLabel $statusLabel}
        "Remove AD groups" = {param($username, $progressBar, $statusLabel) SULCO-RemoveADGroups -username $username -progressBar $progressBar -statusLabel $statusLabel}
        "Remove Entra Groups" = {param($username, $progressBar, $statusLabel) SULCO-RemoveEntraGroups -username $username -progressBar $progressBar -statusLabel $statusLabel}
        "Move to Disabled OU" = {param($username, $progressBar, $statusLabel) SULCO-DisabledOU -username $username -progressBar $progressBar -statusLabel $statusLabel}
        "Clear 'Manager' field" = {param($username, $progressBar, $statusLabel) SULCO-ClearManager -username $username -progressBar $progressBar -statusLabel $statusLabel}
        "Convert to Shared Mailbox" = {param($username, $progressBar, $statusLabel) SULCO-ConvertMailbox -username $username -progressBar $progressBar -statusLabel $statusLabel}
    }
}

# Create a button to fetch user groups
$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(760, 380)
$button.Size = New-Object System.Drawing.Size(180, 40)
$button.Text = "Run"
$button.Add_Click({
    $username = $userTextBox.Text
    $selectedDomain = $comboBox.SelectedItem
    $selectedFunctions = $functionMap[$selectedDomain]

    foreach ($index in 1..($checkboxes.Count - 1)) {
        if ($checkboxes[$index].Checked) {
            $checkboxText = $checkboxes[$index].Text
            $operationSuccess = $false  # Initialize a flag for operation success

            try {
                # Invoke the function with username, progressBar, and statusLabel as parameters
                $operationSuccess = $selectedFunctions[$checkboxText].Invoke($username, $progressBar, $statusLabel)
                Write-Host "Operation success for $checkboxText $operationSuccess"
                write-Host $operationSuccess
            } catch {
                # Log the error in the main text box
                $largeTextBox.AppendText("Error during $checkboxText for $selectedDomain user: $username. Error: $($_.Exception.Message)`r`n")
                Write-Host "Error during $checkboxText $($_.Exception.Message)"
                throw $_
            }

            if ($operationSuccess -eq $true) {
                # Only append success message if the operation succeeded
                $largeTextBox.AppendText("$checkboxText for $selectedDomain user: $username completed successfully.`r`n")
            }
        }
    }

    $statusLabel.Text = "Status: Leaver Process Completed"
})


$form.Controls.Add($button)

# Add a KeyPress event to prevent spaces
$userTextBox.Add_KeyPress({
    param($sender, $e)
    if ($e.KeyChar -eq [char]::Parse(" ")) {
        $e.Handled = $true  # Suppress the space character
    }
})

# Add a large read-only text box beneath the progress bar
$largeTextBox = New-Object System.Windows.Forms.TextBox
$largeTextBox.Location = New-Object System.Drawing.Point(50, 510)
$largeTextBox.Size = New-Object System.Drawing.Size(900, 400)
$largeTextBox.Multiline = $true
$largeTextBox.ScrollBars = "Vertical"
$largeTextBox.ReadOnly = $true
$largeTextBox.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
$form.Controls.Add($largeTextBox)

# Show the form
# This base64 string holds the bytes that make up the orange 'G' icon (just an example for a 32x32 pixel image)
$iconBase64      = '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'
$iconBytes       = [Convert]::FromBase64String($iconBase64)
# initialize a Memory stream holding the bytes
$stream          = [System.IO.MemoryStream]::new($iconBytes, 0, $iconBytes.Length)
$Form.Icon       = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))

$form.ShowDialog()
