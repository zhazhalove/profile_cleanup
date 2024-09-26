# Load the required .NET assemblies for Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to accurately calculate the size of a directory
function Get-AccurateDirectorySize {
    param (
        [string]$Path
    )

    $totalSize = 0
    if (Test-Path $Path) {
        Get-ChildItem -Path $Path -File -Recurse -Force | ForEach-Object {
            try {
                $fileInfo = [System.IO.FileInfo]::new($_.FullName)
                $totalSize += $fileInfo.Length
            }
            catch {
                Write-Verbose "Could not access file: $_.FullName"
            }
        }
    }
    return $totalSize
}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "User Profile Size Viewer"
$form.Size = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "CenterScreen"

# Set the AutoScaleMode to DPI to ensure DPI awareness
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi

# Create a DataGridView to display profile information
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Size = New-Object System.Drawing.Size(760, 400)
$dataGridView.Location = New-Object System.Drawing.Point(10, 10)
$dataGridView.AutoGenerateColumns = $true
$dataGridView.AllowUserToAddRows = $false
$dataGridView.ReadOnly = $true
$dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dataGridView.MultiSelect = $false
$dataGridView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
$form.Controls.Add($dataGridView)

# Create a button to load user profiles
$loadButton = New-Object System.Windows.Forms.Button
$loadButton.Text = "Load Profiles"
$loadButton.Size = New-Object System.Drawing.Size(100, 30)
$loadButton.Location = New-Object System.Drawing.Point(10, 420)
$loadButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($loadButton)

# Create a button to delete the selected profile
$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Text = "Delete"
$deleteButton.Size = New-Object System.Drawing.Size(100, 30)
$deleteButton.Location = New-Object System.Drawing.Point(120, 420)
$deleteButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$deleteButton.Enabled = $false
$form.Controls.Add($deleteButton)

# Create a ProgressBar to show progress
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Size = New-Object System.Drawing.Size(760, 20)
$progressBar.Location = New-Object System.Drawing.Point(10, 460)
$progressBar.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$progressBar.Minimum = 0
$form.Controls.Add($progressBar)



$loadButton.Add_Click({
    $dataGridView.DataSource = $null   # Clear existing data
    $deleteButton.Enabled = $false     # Disable delete button

    # Retrieve user profiles using WMI
    $userProfiles = Get-WmiObject -Class Win32_UserProfile | Where-Object {
        -not $_.Special -and
        $_.LocalPath -like "C:\Users\*" -and
        $_.SID -match "^S-1-5-21-" -and
        -not $_.Loaded
    }

    # Set the progress bar maximum value
    $progressBar.Maximum = $userProfiles.Count
    $progressBar.Value = 0

    # Prepare data for the DataGridView
    $profileData = @()

    foreach ($userProfile in $userProfiles) {
        $profileSize = Get-AccurateDirectorySize -Path $userProfile.LocalPath

        try {
            $userName = (New-Object System.Security.Principal.SecurityIdentifier($userProfile.SID)).Translate([System.Security.Principal.NTAccount]).ToString()
        } catch {
            Write-Verbose "Invalid or missing SID: $($userProfile.SID)"
            $userName = "unknown_SID"
        }
        

        $profileData += [PSCustomObject]@{
            UserName     = $userName
            LocalPath    = $userProfile.LocalPath
            LastUseTime  = [Management.ManagementDateTimeConverter]::ToDateTime($userProfile.LastUseTime).ToString()
            ProfileSizeMB = [math]::Round($profileSize / 1MB, 2).ToString("N2") # Convert to a string with 2 decimal places
            SID          = $userProfile.SID # Store the SID for deletion
        }

        # Update the progress bar
        $form.Invoke([Action]{
            $progressBar.Value += 1
        })
    }

    # Sort the profile data by ProfileSizeMB in descending order
    $sortedProfileData = $profileData | Sort-Object -Property {[double]$_.ProfileSizeMB} -Descending

    # Update the DataGridView using the UI thread
    $form.Invoke([Action]{
        $bindingList = [System.ComponentModel.BindingList[object]]::new()
        $sortedProfileData | ForEach-Object { $bindingList.Add($_) }
        $dataGridView.DataSource = $bindingList
        $dataGridView.AutoResizeColumns()
        $dataGridView.Refresh()
    })

    # Reset the progress bar
    $form.Invoke([Action]{
        $progressBar.Value = 0
    })
})


# Enable the Delete button when a row is selected
$dataGridView.add_SelectionChanged({
    if ($dataGridView.SelectedRows.Count -eq 1) {
        $deleteButton.Enabled = $true
    } else {
        $deleteButton.Enabled = $false
    }
})

# Event handler for the Delete button click
$deleteButton.Add_Click({
    if ($dataGridView.SelectedRows.Count -eq 1) {
        $selectedProfile = $dataGridView.SelectedRows[0].DataBoundItem
        $userName = $selectedProfile.UserName
        $profileSID = $selectedProfile.SID

        # Display a confirmation dialog
        $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete $($userName)?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                # Disable the Load and Delete buttons
                $deleteButton.Enabled = $false
                $loadButton.Enabled = $false

                # Delete the profile using WMI
                $profileToDelete = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.SID -eq $profileSID }
                
                if ($profileToDelete) {
                    $profileToDelete.Delete()
                    [System.Windows.Forms.MessageBox]::Show("$($userName) has been deleted.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

                    # Remove the selected row from DataGridView
                    $dataGridView.Rows.Remove($dataGridView.SelectedRows[0])
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Profile not found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
            finally {
                # Re-enable the buttons
                $deleteButton.Enabled = $true
                $loadButton.Enabled = $true
                # $loadButton.PerformClick() # Reload profiles
            }
        }
    }
})


# Run the form
[void]$form.ShowDialog()
