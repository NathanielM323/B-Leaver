#Import Fetch Leaver Data module using reletive file path

$exePath = [System.IO.Path]::GetDirectoryName([System.Reflection.Assembly]::GetExecutingAssembly().Location)
$modulePath = Join-Path $exePath "Fetch_Leaver_data.psm1"

# Import the module using the path
Import-Module $modulePath -Force

Get-Module -Name Fetch_Leaver_data


# Load Windows Forms Assembly
Add-Type -AssemblyName System.Windows.Forms

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Leaver Process Automation"
$form.Size = New-Object System.Drawing.Size(1200, 1200)
$form.StartPosition = "CenterScreen"

# Add a label to display current process
$label = New-Object System.Windows.Forms.Label
$label.Text = "Status: Idle"
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point(50, 460)
$form.Controls.Add($label)

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
$checkboxOptions = @("Retrieve all leaver data", "Retrieve O365 groups", "Retrieve Distribution lists", "Retrieve Shared Mailboxes")
$checkboxes = @()
$startY = 160  # Starting Y position for the first checkbox
$indent = 23   # Indentation for sub-options

foreach ($index in 0..($checkboxOptions.Count - 1)) {
    $checkbox = New-Object System.Windows.Forms.CheckBox
    $checkbox.Text = $checkboxOptions[$index]
    
    if ($index -eq 0) {
        # First option (Option 1) is not indented
        $checkbox.Location = New-Object System.Drawing.Point(180, $startY)
    } else {
        # Other options are indented
        $checkbox.Location = New-Object System.Drawing.Point((180 + $indent), $startY)
    }
    
    $checkbox.Size = New-Object System.Drawing.Size(200, 20)
    $checkbox.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Regular)
    $form.Controls.Add($checkbox)
    $checkboxes += $checkbox
    $startY += 23  # Increment Y position for the next checkbox
}

# Add this after creating the checkboxes
$checkboxO365Groups = $checkboxes[1]  # Assuming "Retrieve O365 groups" is the second checkbox

# Modify the button click event
# Create a button to fetch user groups
$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(60, 180)
$button.Size = New-Object System.Drawing.Size(180, 40)
$button.Text = "Run"
$button.Add_Click({
    $username = $userTextBox.Text

    if ($comboBox.SelectedItem -eq "BBB") {
        if ($checkboxO365Groups.Checked) {
            try {
                Get-BBBUserGroupsExport -username $username
                $largeTextBox.AppendText("O365 groups exported for BBB user: $username`r`n")
            } catch {
                $largeTextBox.AppendText("Error exporting O365 groups for BBB user: $username. Error: $($_.Exception.Message)`r`n")
            }
        }
    } elseif ($comboBox.SelectedItem -eq "SULCO") {
        if ($checkboxO365Groups.Checked) {
            try {
                Get-SulcoUserGroupsExport -username $username
                $largeTextBox.AppendText("O365 groups exported for SULCO user: $username`r`n")
            } catch {
                $largeTextBox.AppendText("Error exporting O365 groups for SULCO user: $username. Error: $($_.Exception.Message)`r`n")
            }
        }
    }

    # Update progress
    $progressBar.Value += 25
    $label.Text = "Status: O365 Groups Export Complete"
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
