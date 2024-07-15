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
