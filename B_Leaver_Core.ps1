cd '\\cfel.local\dfsroot\ICT\Nathaniel\Leaver\Leaver Automation\B-Leaver\B-Leaver(git)'

$currentDirectory = Get-Location
$modulePath = Join-Path $currentDirectory "Fetch_Leaver_data.psm1"

# Import 'Fetch_LEaver_data' using the module using the path
Import-Module $modulePath -Force
Get-Module -Name Fetch_Leaver_data

# Import 'Hide_From_GAL' using the module using the path
$modulePath = Join-Path $currentDirectory "Hide_From_GAL.psm1"
Import-Module $modulePath -Force
Get-Module -Name Hide_From_GAL



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
$form.Text = "Leaver Process Automation"
$form.Size = New-Object System.Drawing.Size(1200, 1200)
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
$checkboxOptions = @("Retrieve all leaver data", "Retrieve O365 groups", "Retrieve Distribution lists", "Retrieve Shared Mailboxes" , "Hide from GAL")
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
    }
    "SULCO" = @{
        "Retrieve O365 groups" = { param($username, $progressBar, $statusLabel) Get-SulcoUserGroupsExport -username $username -progressBar $progressBar -statusLabel $statusLabel }
        "Retrieve Distribution lists" = { param($username, $progressBar, $statusLabel) Get-SulcoUserDLExport -username $username -progressBar $progressBar -statusLabel $statusLabel }
        "Retrieve Shared Mailboxes" = { param($username, $progressBar, $statusLabel) Get-SulcoMailbox -username $username -progressBar $progressBar -statusLabel $statusLabel }
        "Hide from GAL" = {param($username) GALHideSULCO -username $username}
    }
}

# Create a button to fetch user groups
$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(60, 180)
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
            } catch {
                # Log the error in the main text box
                $largeTextBox.AppendText("Error during $checkboxText for $selectedDomain user: $username. Error: $($_.Exception.Message)`r`n")
                Write-Host "Error during $checkboxText $($_.Exception.Message)"
            }

            if ($operationSuccess) {
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
$largeTextBox.Size = New-Object System.Drawing.Size(1100, 400)
$largeTextBox.Multiline = $true
$largeTextBox.ScrollBars = "Vertical"
$largeTextBox.ReadOnly = $true
$largeTextBox.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
$form.Controls.Add($largeTextBox)

# Show the form
$form.ShowDialog()
