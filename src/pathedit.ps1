<#
.SYNOPSIS
    A PowerShell GUI tool for managing Windows PATH environment variables.

.DESCRIPTION
    PATHedit provides a graphical interface to view and modify System, User, and
    Temporary PATH environment variables. It features path validation, reordering,
    and automatic backups before changes.

.NOTES
    Requirements:
    - Windows PowerShell 5.1 or PowerShell Core 7.x
    - Windows OS
    - Administrator privileges (for System PATH modifications)

    Restrictions:
    - Must be run with administrator privileges to modify System PATH
    - Can only modify PATH environment variables
    - Windows-only compatibility

.EXAMPLE
    .\pathedit.ps1
#>

# Suppress progress bars and non-terminating errors
$ProgressPreference = 'SilentlyContinue'
$ErrorActionPreference = 'SilentlyContinue'

# Check if script is running with admin privileges, if not, restart with elevation
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    try {
        $scriptPath = $MyInvocation.MyCommand.Path
        if (!$scriptPath) { $scriptPath = $PSCommandPath }
        Start-Process pwsh -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
    }
    catch {
        Write-Error "Failed to elevate privileges: $($_.Exception.Message)"
    }
    exit
}

# Import required Windows Forms assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Backup-EnvironmentVariables {
    # Creates a backup of current system and user PATH variables
    [CmdletBinding()]
    param()

    $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    # Use Documents folder as fallback location if script path cannot be determined
    $ScriptDir = Split-Path -Parent -Path ($MyInvocation.MyCommand.Path)
    if (!$ScriptDir) {
        $ScriptDir = Split-Path -Parent -Path $PSCommandPath
        if (!$ScriptDir) {
            $ScriptDir = Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath 'PATHedit'
        }
    }
    $BackupPath = Join-Path -Path $ScriptDir -ChildPath "PathBackup_$Timestamp.txt"

    try {
        # Create directory if it doesn't exist
        $BackupDir = Split-Path -Parent -Path $BackupPath
        if (!(Test-Path -Path $BackupDir)) {
            New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null
        }

        "$((Get-Date).ToString()) - Backing up environment variables" | Out-File -Path $BackupPath
        "----------------- System Path -----------------" | Out-File -Path $BackupPath -Append
        ([Environment]::GetEnvironmentVariable("Path", "Machine") | Out-String) | Out-File -Path $BackupPath -Append
        "------------------ User Path ------------------" | Out-File -Path $BackupPath -Append
        ($env:Path | Out-String) | Out-File -Path $BackupPath -Append
        Write-Host "Environment variable backup created at: $BackupPath" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to create environment variable backup: $($_.Exception.Message)"
    }
}

function Show-FolderBrowser {
    # Displays a folder browser dialog
    # Returns: Selected folder path or null if cancelled
    [CmdletBinding()]
    param(
        [string]$Description = "Select a folder",
        [string]$InitialDirectory = "C:\"
    )

    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = $Description
    $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::MyComputer
    $folderBrowser.ShowNewFolderButton = $true
    $folderBrowser.UseDescriptionForTitle = $true

    if ($InitialDirectory -and (Test-Path $InitialDirectory)) {
        $folderBrowser.SelectedPath = $InitialDirectory
    }

    if ($folderBrowser.ShowDialog($script:objForm) -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    }
    return $null
}

function Show-InputDialog {
    # Displays a dialog for text input with browse capability
    # Returns: Entered text or null if cancelled
    [CmdletBinding()]
    param(
        [string]$Title = "Input",
        [string]$Message = "Enter value:",
        [string]$DefaultValue = ""
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(400,150)
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.Owner = $script:objForm

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,10)
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.Text = $Message
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,30)
    $textBox.Size = New-Object System.Drawing.Size(360,20)
    $textBox.Text = $DefaultValue
    $form.Controls.Add($textBox)

    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Location = New-Object System.Drawing.Point(290,60)
    $browseButton.Size = New-Object System.Drawing.Size(80,23)
    $browseButton.Text = "Browse..."
    $browseButton.Add_Click({
        $initialPath = if ($textBox.Text -and (Test-Path $textBox.Text)) {
            $textBox.Text
        } else {
            "C:\"
        }

        $selectedPath = Show-FolderBrowser -Description "Select a folder" -InitialDirectory $initialPath
        if ($selectedPath) {
            $textBox.Text = $selectedPath
            $textBox.Select($textBox.Text.Length, 0)
            $textBox.Focus()
        }
    })
    $form.Controls.Add($browseButton)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(200,60)
    $okButton.Size = New-Object System.Drawing.Size(80,23)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    }
    return $null
}

function Show-ChangeConfirmation {
    # Shows a confirmation dialog with pending PATH changes
    # Returns: DialogResult indicating user's choice
    [CmdletBinding()]
    param(
        [hashtable]$Changes
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Confirm Changes"
    $form.Size = New-Object System.Drawing.Size(500,400)
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,10)
    $textBox.Size = New-Object System.Drawing.Size(460,300)
    $textBox.Multiline = $true
    $textBox.ScrollBars = "Vertical"
    $textBox.ReadOnly = $true

    $message = "The following changes will be made:`n`n"
    foreach ($key in $Changes.Keys) {
        $message += "$key Path:`n$($Changes[$key])`n`n"
    }
    $textBox.Text = $message
    $form.Controls.Add($textBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(290,320)
    $okButton.Size = New-Object System.Drawing.Size(80,30)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(380,320)
    $cancelButton.Size = New-Object System.Drawing.Size(80,30)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancelButton)

    return $form.ShowDialog()
}

function Show-ChangeResult {
    # Displays results of PATH variable updates
    [CmdletBinding()]
    param(
        [hashtable]$Results
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Update Results"
    $form.Size = New-Object System.Drawing.Size(500,400)
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,10)
    $textBox.Size = New-Object System.Drawing.Size(460,300)
    $textBox.Multiline = $true
    $textBox.ScrollBars = "Vertical"
    $textBox.ReadOnly = $true

    $message = ""
    foreach ($key in $Results.Keys) {
        $result = $Results[$key]
        $message += "$($result.Message)`n"
        if ($result.Error) {
            $message += "Error: $($result.Error)`n"
        }
        if ($result.NewValue) {
            $message += "New Value: $($result.NewValue)`n"
        }
        $message += "`n"
    }
    $textBox.Text = $message
    $form.Controls.Add($textBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(380,320)
    $okButton.Size = New-Object System.Drawing.Size(80,30)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)

    [void]$form.ShowDialog()
}

function Update-PathColors {
    # Refreshes the visual state of the ListBox
    [CmdletBinding()]
    param(
        [System.Windows.Forms.ListBox]$ListBox
    )
    $ListBox.Refresh()
}

function Add-DoubleClickHandler {
    # Adds double-click editing capability to a ListBox
    [CmdletBinding()]
    param (
        [System.Windows.Forms.ListBox]$ListBox
    )

    $ListBox.Add_DoubleClick({
        if ($this.SelectedItem) {
            $path = $this.SelectedItem.ToString()
            if (Test-Path -Path $path -PathType Container) {
                $selectedPath = Show-FolderBrowser -Description "Select a folder" -InitialDirectory $path
                if ($selectedPath) {
                    $selectedIndex = $this.SelectedIndex
                    $this.Items.RemoveAt($selectedIndex)
                    $this.Items.Insert($selectedIndex, $selectedPath)
                    $this.SelectedIndex = $selectedIndex
                    Update-PathColors -ListBox $this
                }
            } else {
                $editedPath = Show-InputDialog -Title "Edit Path" -Message "Edit path:" -DefaultValue $path
                if ($editedPath) {
                    $selectedIndex = $this.SelectedIndex
                    $this.Items.RemoveAt($selectedIndex)
                    $this.Items.Insert($selectedIndex, $editedPath.Trim())
                    $this.SelectedIndex = $selectedIndex
                    Update-PathColors -ListBox $this
                }
            }
        }
    })
}

function Update-ListBoxAppearance {
    # Configures ListBox visual appearance and custom drawing
    [CmdletBinding()]
    param(
        [System.Windows.Forms.ListBox]$ListBox
    )

    $ListBox.BackColor = [System.Drawing.Color]::White
    $ListBox.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $ListBox.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed

    $drawHandler = {
        param($sender, $e)

        if ($e.Index -lt 0) { return }

        $backColor = if ($e.Index % 2 -eq 0) {
            [System.Drawing.Color]::White
        } else {
            [System.Drawing.Color]::FromArgb(240, 240, 240)
        }

        $isSelected = ($e.State -band [System.Windows.Forms.DrawItemState]::Selected) -eq [System.Windows.Forms.DrawItemState]::Selected

        $backBrush = New-Object System.Drawing.SolidBrush($(
            if ($isSelected) {
                [System.Drawing.Color]::FromArgb(0, 120, 215)
            } else {
                $backColor
            }
        ))

        $e.Graphics.FillRectangle($backBrush, $e.Bounds)
        $backBrush.Dispose()

        $path = $sender.Items[$e.Index].ToString()
        $pathExists = Test-Path -Path $path -PathType Container

        $textBrush = New-Object System.Drawing.SolidBrush($(
            if ($isSelected) {
                if ($pathExists) {
                    [System.Drawing.Color]::White
                } else {
                    [System.Drawing.Color]::FromArgb(255, 150, 150)
                }
            } else {
                if ($pathExists) {
                    [System.Drawing.SystemColors]::WindowText
                } else {
                    [System.Drawing.Color]::Red
                }
            }
        ))

        $e.Graphics.DrawString($path, $e.Font, $textBrush,
            ($e.Bounds.Left + 2),
            ($e.Bounds.Top + 2))
        $textBrush.Dispose()
    }

    $ListBox.Tag = $drawHandler
    $ListBox.Add_DrawItem($drawHandler)
}

function Update-ButtonAppearance {
    # Sets consistent button styling
    [CmdletBinding()]
    param(
        [System.Windows.Forms.Button]$Button,
        [switch]$IsOKButton
    )

    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard

    if ($IsOKButton) {
        $Button.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
        $Button.ForeColor = [System.Drawing.Color]::White
    } else {
        $Button.BackColor = [System.Drawing.SystemColors]::Control
        $Button.UseVisualStyleBackColor = $false
    }
}

# Create main form and controls
$script:objForm = New-Object System.Windows.Forms.Form
$objForm.Text = "PATHedit"
$objForm.Size = New-Object System.Drawing.Size(600, 400)
$objForm.StartPosition = "CenterScreen"
$objForm.FormBorderStyle = 'FixedSingle'
$objForm.MaximizeBox = $false

$panelMain = New-Object System.Windows.Forms.Panel
$panelMain.Location = New-Object System.Drawing.Point(10, 10)
$panelMain.Size = New-Object System.Drawing.Size(460, 300)
$objForm.Controls.Add($panelMain)

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = 'Fill'
$panelMain.Controls.Add($tabControl)

$tabSystem = New-Object System.Windows.Forms.TabPage
$tabSystem.Text = "System Path"
$tabUser = New-Object System.Windows.Forms.TabPage
$tabUser.Text = "User Path"
$tabTemp = New-Object System.Windows.Forms.TabPage
$tabTemp.Text = "Temporary Path"

$tabControl.Controls.AddRange(@($tabSystem, $tabUser, $tabTemp))

$listSystem = New-Object System.Windows.Forms.ListBox
$listSystem.Dock = 'Fill'
$listUser = New-Object System.Windows.Forms.ListBox
$listUser.Dock = 'Fill'
$listTemp = New-Object System.Windows.Forms.ListBox
$listTemp.Dock = 'Fill'

$tabSystem.Controls.Add($listSystem)
$tabUser.Controls.Add($listUser)
$tabTemp.Controls.Add($listTemp)

# Define button configurations
$buttons = @{
    'New' = @{ Location = @(480, 33); Text = "New..." }
    'Edit' = @{ Location = @(480, 73); Text = "Edit..." }
    'Delete' = @{ Location = @(480, 113); Text = "Delete" }
    'MoveUp' = @{ Location = @(480, 153); Text = "Move Up" }
    'MoveDown' = @{ Location = @(480, 193); Text = "Move Down" }
    'Cancel' = @{ Location = @(380, 320); Text = "Cancel" }
    'OK' = @{ Location = @(480, 320); Text = "OK" }
}

# Create status label with GitHub link
$statusLabel = New-Object System.Windows.Forms.LinkLabel
$statusLabel.Location = New-Object System.Drawing.Point(10, 325)
$statusLabel.Size = New-Object System.Drawing.Size(360, 20)
$statusLabel.Text = "https://github.com/fl4pj4ck/pathedit"
$statusLabel.LinkColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$statusLabel.ActiveLinkColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$statusLabel.VisitedLinkColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$statusLabel.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline
$statusLabel.Add_LinkClicked({
    Start-Process $statusLabel.Text
})
$objForm.Controls.Add($statusLabel)

# Create and configure buttons
$buttonObjects = @{}
foreach ($btn in $buttons.GetEnumerator()) {
    $button = New-Object System.Windows.Forms.Button
    $button.Location = New-Object System.Drawing.Point($btn.Value.Location[0], $btn.Value.Location[1])
    $button.Size = New-Object System.Drawing.Size(90, 30)
    $button.Text = $btn.Value.Text
    $buttonObjects[$btn.Key] = $button
    $objForm.Controls.Add($button)
}

foreach ($btn in $buttonObjects.GetEnumerator()) {
    Update-ButtonAppearance -Button $btn.Value -IsOKButton:($btn.Key -eq 'OK')
}

# Initialize ListBox appearances and handlers
Update-ListBoxAppearance $listSystem
Update-ListBoxAppearance $listUser
Update-ListBoxAppearance $listTemp

Add-DoubleClickHandler -ListBox $listSystem
Add-DoubleClickHandler -ListBox $listUser
Add-DoubleClickHandler -ListBox $listTemp

# Define environment variable loading logic
$LoadEnvironmentVariables = {
    # Load System PATH
    $systemPath = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
    $listSystem.Items.Clear()
    [void]($systemPath.Split(';') | Where-Object { $_ } | ForEach-Object {
        $listSystem.Items.Add($_)
    })

    $userPath = [System.Environment]::GetEnvironmentVariable("Path", "User")
    $listUser.Items.Clear()
    [void]($userPath.Split(';') | Where-Object { $_ } | ForEach-Object {
        $listUser.Items.Add($_)
    })

    $listTemp.Items.Clear()
    [void]($env:Path.Split(';') | Where-Object { $_ } | ForEach-Object {
        $listTemp.Items.Add($_)
    })

    Update-PathColors -ListBox $listSystem
    Update-PathColors -ListBox $listUser
    Update-PathColors -ListBox $listTemp
}

# Button click event handlers
$buttonObjects['New'].Add_Click({
    $selectedListBox = switch ($tabControl.SelectedIndex) {
        0 {$listSystem}
        1 {$listUser}
        2 {$listTemp}
    }
    $newPath = Show-InputDialog -Title "New Path" -Message "Enter new path:"
    if ($newPath) {
        $selectedListBox.Items.Add($newPath.Trim())
        Update-PathColors -ListBox $selectedListBox
    }
})

$buttonObjects['Edit'].Add_Click({
    $selectedListBox = switch ($tabControl.SelectedIndex) {
        0 {$listSystem}
        1 {$listUser}
        2 {$listTemp}
    }
    if ($selectedListBox.SelectedItem) {
        $oldPath = $selectedListBox.SelectedItem
        $editedPath = Show-InputDialog -Title "Edit Path" -Message "Edit path:" -DefaultValue $oldPath
        if ($editedPath) {
            $selectedIndex = $selectedListBox.SelectedIndex
            $selectedListBox.Items.RemoveAt($selectedIndex)
            $selectedListBox.Items.Insert($selectedIndex, $editedPath.Trim())
            Update-PathColors -ListBox $selectedListBox
        }
    }
})

$buttonObjects['Delete'].Add_Click({
    $selectedListBox = switch ($tabControl.SelectedIndex) {
        0 {$listSystem}
        1 {$listUser}
        2 {$listTemp}
    }
    if ($selectedListBox.SelectedItem) {
        $selectedListBox.Items.Remove($selectedListBox.SelectedItem)
        Update-PathColors -ListBox $selectedListBox
    }
})

$buttonObjects['MoveUp'].Add_Click({
    $selectedListBox = switch ($tabControl.SelectedIndex) {
        0 {$listSystem}
        1 {$listUser}
        2 {$listTemp}
    }
    if ($selectedListBox.SelectedItem -and $selectedListBox.SelectedIndex -gt 0) {
        $currentIndex = $selectedListBox.SelectedIndex
        $selectedValue = $selectedListBox.SelectedItem
        $selectedListBox.Items.RemoveAt($currentIndex)
        $selectedListBox.Items.Insert($currentIndex - 1, $selectedValue)
        $selectedListBox.SelectedIndex = $currentIndex - 1
    }
})

$buttonObjects['MoveDown'].Add_Click({
    $selectedListBox = switch ($tabControl.SelectedIndex) {
        0 {$listSystem}
        1 {$listUser}
        2 {$listTemp}
    }
    if ($selectedListBox.SelectedItem -and $selectedListBox.SelectedIndex -lt $selectedListBox.Items.Count - 1) {
        $currentIndex = $selectedListBox.SelectedIndex
        $selectedValue = $selectedListBox.SelectedItem
        $selectedListBox.Items.RemoveAt($currentIndex)
        $selectedListBox.Items.Insert($currentIndex + 1, $selectedValue)
        $selectedListBox.SelectedIndex = $currentIndex + 1
    }
})

$buttonObjects['OK'].Add_Click({
    $changes = @{}

    $newSystemPath = ($listSystem.Items | ForEach-Object { $_.ToString() }) -join ';'
    if ($newSystemPath -ne [System.Environment]::GetEnvironmentVariable("Path", "Machine")) {
        $changes['System'] = $newSystemPath
    }

    $newUserPath = ($listUser.Items | ForEach-Object { $_.ToString() }) -join ';'
    if ($newUserPath -ne [System.Environment]::GetEnvironmentVariable("Path", "User")) {
        $changes['User'] = $newUserPath
    }

    $newTempPath = ($listTemp.Items | ForEach-Object { $_.ToString() }) -join ';'
    if ($newTempPath -ne $env:Path) {
        $changes['Temp'] = $newTempPath
    }

    if ($changes.Count -gt 0) {
        $result = Show-ChangeConfirmation -Changes $changes
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            Backup-EnvironmentVariables

            $results = @{}

            foreach ($key in $changes.Keys) {
                try {
                    switch ($key) {
                        'System' {
                            [System.Environment]::SetEnvironmentVariable("Path", $changes[$key], "Machine")
                            $results[$key] = @{
                                Message = "System Path updated successfully."
                                Error = $null
                                NewValue = $changes[$key]
                            }
                        }
                        'User' {
                            [System.Environment]::SetEnvironmentVariable("Path", $changes[$key], "User")
                            $results[$key] = @{
                                Message = "User Path updated successfully."
                                Error = $null
                                NewValue = $changes[$key]
                            }
                        }
                        'Temp' {
                            $env:Path = $changes[$key]
                            $results[$key] = @{
                                Message = "Temporary Path updated successfully."
                                Error = $null
                                NewValue = $env:Path
                            }
                        }
                    }
                }
                catch {
                    $results[$key] = @{
                        Message = "Failed to update $key Path:"
                        Error = $_.Exception.Message
                        NewValue = $null
                    }
                }
            }
            Show-ChangeResult -Results $results
        }
    }
    $objForm.Close()
})

$buttonObjects['Cancel'].Add_Click({ $objForm.Close() })

# Load initial PATH values and show form
$LoadEnvironmentVariables.Invoke()

$objForm.Add_Shown({$objForm.Activate()})
[void]$objForm.ShowDialog()