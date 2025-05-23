Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Funktion zur Überprüfung und Anforderung von Administratorrechten
function Ensure-Admin {
    # Prüfen, ob das Skript bereits mit Administratorrechten ausgeführt wird
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        # Dialogfenster anzeigen, um den Benutzer zu fragen, ob das Skript als Administrator neu gestartet werden soll
        $dialogResult = [System.Windows.Forms.MessageBox]::Show("Dieses Skript erfordert Administratorrechte. Möchten Sie es als Administrator neu starten?", "Administratorrechte erforderlich", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)

        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Versuch, das Skript mit Administratorrechten neu zu starten
            try {
                # Parameter für den Neustart vorbereiten
                $powershellPath = (Get-Command PowerShell).Source
                $argumentList = "-File `"$($MyInvocation.MyCommand.Path)`""
                
                # Prozess als Administrator starten
                Start-Process -FilePath $powershellPath -Verb RunAs -ArgumentList $argumentList -ErrorAction Stop
                
                # Aktuelle (nicht-administrative) Instanz beenden
                Write-Host "Die nicht-administrative Instanz wird beendet, da eine neue administrative Instanz gestartet wurde."
                exit
            } catch {
                # Fehlermeldung anzeigen, wenn der Neustart fehlschlägt, und Skript beenden
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Neustart des Skripts als Administrator:`n$($_.Exception.Message)", "Fehler beim Neustart", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                Write-Host "Fehler beim Neustart als Administrator. Das Skript wird beendet."
                exit
            }
        } else {
            # Informationsmeldung anzeigen, wenn der Benutzer "Nein" wählt, und Skript beenden
            [System.Windows.Forms.MessageBox]::Show("Das Skript kann ohne Administratorrechte nicht fortfahren und wird nun beendet.", "Administratorrechte verweigert", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            Write-Host "Benutzer hat den Neustart als Administrator abgelehnt. Das Skript wird beendet."
            exit
        }
    }
    # Wenn das Skript bereits als Administrator ausgeführt wird, tut die Funktion nichts
    # und die Skriptausführung wird normal fortgesetzt.
}
Ensure-Admin
# GUI Creation
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Execution Policy Manager"
$mainForm.Size = New-Object System.Drawing.Size(780, 400)
$mainForm.StartPosition = "CenterScreen"
$mainForm.FormBorderStyle = 'FixedDialog'
$mainForm.MaximizeBox = $false
$mainForm.MinimizeBox = $false

$global:currentPolicies = @{}
$policyOptions = @("Restricted", "AllSigned", "RemoteSigned", "Unrestricted", "Bypass", "Undefined")

# Function to get current Execution Policies
function Get-CurrentExecutionPolicies {
    $policies = Get-ExecutionPolicy -List | Select-Object Scope, ExecutionPolicy
    $global:currentPolicies.Clear()
    foreach ($policy in $policies) {
        $global:currentPolicies[$policy.Scope.ToString()] = $policy.ExecutionPolicy.ToString()
    }
    foreach ($scopeName in @("MachinePolicy", "UserPolicy", "Process", "CurrentUser", "LocalMachine")) {
        if (-not $global:currentPolicies.ContainsKey($scopeName)) {
            $global:currentPolicies[$scopeName] = "Undefined"
        }
    }
}

# Function to update GUI display
function Update-PolicyDisplay {
    Get-CurrentExecutionPolicies
    $scopes = @("MachinePolicy", "UserPolicy", "Process", "CurrentUser", "LocalMachine")
    for ($i = 0; $i -lt $scopes.Length; $i++) {
        $scopeName = $scopes[$i]
        $label = $mainForm.Controls.Find("lblCurrent$scopeName", $true)[0]
        $comboBox = $mainForm.Controls.Find("cb$scopeName", $true)[0]

        if ($label) {
            $label.Text = $global:currentPolicies[$scopeName]
        }
        if ($comboBox) {
            $currentIndex = $policyOptions.IndexOf($global:currentPolicies[$scopeName])
            if ($currentIndex -eq -1) { $currentIndex = $policyOptions.IndexOf("Undefined") } 
            if ($currentIndex -ne -1) {
                 $comboBox.SelectedIndex = $currentIndex
            } else {
                 $comboBox.SelectedIndex = $policyOptions.IndexOf("Undefined") 
            }
        }
    }
    try {
        $effectivePolicy = Get-ExecutionPolicy
        $lblEffectivePolicyValue.Text = $effectivePolicy.ToString()
        $lblEffectivePolicyValue.ForeColor = [System.Drawing.Color]::Black
    } catch {
        $lblEffectivePolicyValue.Text = "Error retrieving"
        $lblEffectivePolicyValue.ForeColor = [System.Drawing.Color]::Red
    }
}

# Create GUI elements (Labels and ComboBoxes for each Scope)
$yPos = 25
$defaultFontSize = 9.5
$labelFont = New-Object System.Drawing.Font("Segoe UI", $defaultFontSize)
$boldLabelFont = New-Object System.Drawing.Font("Segoe UI", $defaultFontSize, [System.Drawing.FontStyle]::Bold)

$scopesToManage = @(
    @{Name="MachinePolicy"; Display="MachinePolicy (GPO - All users of this computer)"; Settable=$false},
    @{Name="UserPolicy"; Display="UserPolicy (GPO - Current user)"; Settable=$false},
    @{Name="Process"; Display="Process (Only for current PowerShell session)"; Settable=$true},
    @{Name="CurrentUser"; Display="CurrentUser (Settings for current user)"; Settable=$true},
    @{Name="LocalMachine"; Display="LocalMachine (Default for all users of this computer)"; Settable=$true}
)

foreach ($scopeInfo in $scopesToManage) {
    $scopeName = $scopeInfo.Name
    $scopeDisplay = $scopeInfo.Display
    $isSettable = $scopeInfo.Settable

    # Label for Scope name (Column 1)
    $labelScope = New-Object System.Windows.Forms.Label
    $labelScope.Text = "${scopeDisplay}:"
    $labelScope.Location = New-Object System.Drawing.Point(20, $yPos)
    $labelScope.AutoSize = $true
    $labelScope.Font = $labelFont
    $mainForm.Controls.Add($labelScope)

    # Label for current Policy (Column 2)
    $lblCurrentPolicy = New-Object System.Windows.Forms.Label
    $lblCurrentPolicy.Name = "lblCurrent$scopeName"
    $lblCurrentPolicy.Text = "loading..."
    $lblCurrentPolicy.Location = New-Object System.Drawing.Point(360, $yPos)
    $lblCurrentPolicy.AutoSize = $true
    $lblCurrentPolicy.Font = $boldLabelFont
    $mainForm.Controls.Add($lblCurrentPolicy)

    # ComboBox to set Policy or placeholder (Column 3)
    if ($isSettable) {
        $cbPolicy = New-Object System.Windows.Forms.ComboBox
        $cbPolicy.Name = "cb$scopeName"
        $cbPolicy.Location = New-Object System.Drawing.Point(540, ($yPos - 2))
        $cbPolicy.Size = New-Object System.Drawing.Size(180, 28)
        $cbPolicy.DropDownStyle = "DropDownList"
        $cbPolicy.Font = $labelFont
        $policyOptions | ForEach-Object { $cbPolicy.Items.Add($_) } | Out-Null
        $mainForm.Controls.Add($cbPolicy)
    } else {
        $lblNotSettable = New-Object System.Windows.Forms.Label
        $lblNotSettable.Text = "(Controlled by Group Policy)"
        $lblNotSettable.Location = New-Object System.Drawing.Point(540, $yPos)
        $lblNotSettable.AutoSize = $true
        $lblNotSettable.Font = $labelFont
        $lblNotSettable.ForeColor = [System.Drawing.Color]::DarkSlateGray
        $mainForm.Controls.Add($lblNotSettable)
    }
    $yPos += 40
}

# Label for effective Execution Policy
$lblEffectivePolicy = New-Object System.Windows.Forms.Label
$lblEffectivePolicy.Text = "Effective Policy (current session):"
$lblEffectivePolicy.Location = New-Object System.Drawing.Point(20, ($yPos + 10))
$lblEffectivePolicy.AutoSize = $true
$lblEffectivePolicy.Font = $labelFont
$mainForm.Controls.Add($lblEffectivePolicy)

$lblEffectivePolicyValue = New-Object System.Windows.Forms.Label
$lblEffectivePolicyValue.Name = "lblEffectivePolicyValue"
$lblEffectivePolicyValue.Text = "loading..."
$lblEffectivePolicyValue.Location = New-Object System.Drawing.Point(360, ($yPos + 10))
$lblEffectivePolicyValue.AutoSize = $true
$lblEffectivePolicyValue.Font = $boldLabelFont
$mainForm.Controls.Add($lblEffectivePolicyValue)

# Status Label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Name = "statusLabel"
$statusLabel.Text = "Ready."
$statusLabel.Location = New-Object System.Drawing.Point(20, ($yPos + 50))
$statusLabel.AutoSize = $true
$statusLabel.Font = $labelFont
$statusLabel.ForeColor = [System.Drawing.Color]::Green
$mainForm.Controls.Add($statusLabel)

# Button to apply changes
$applyButton = New-Object System.Windows.Forms.Button
$applyButton.Text = "APPLY Selected Policies"
$applyButton.Location = New-Object System.Drawing.Point(360, ($yPos + 85))
$applyButton.Size = New-Object System.Drawing.Size(260, 35)
$applyButton.Font = $boldLabelFont
$applyButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$applyButton.FlatAppearance.BorderSize = 1
$applyButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
$applyButton.BackColor = [System.Drawing.Color]::LightSteelBlue
$mainForm.Controls.Add($applyButton)

# Button to refresh
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "Refresh"
$refreshButton.Location = New-Object System.Drawing.Point(20, ($yPos + 85))
$refreshButton.Size = New-Object System.Drawing.Size(150, 35)
$refreshButton.Font = $labelFont
$refreshButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$refreshButton.FlatAppearance.BorderSize = 1
$refreshButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
$refreshButton.BackColor = [System.Drawing.Color]::LightGray
$mainForm.Controls.Add($refreshButton)

# Event Handler for Refresh button
$refreshButton.Add_Click({
    $statusLabel.Text = "Refreshing policies..."
    $statusLabel.ForeColor = [System.Drawing.Color]::Orange
    Update-PolicyDisplay
    $statusLabel.Text = "Policies refreshed."
    $statusLabel.ForeColor = [System.Drawing.Color]::Green
})

# Event Handler for Apply button
$applyButton.Add_Click(
    
    $statusLabel.Text = "Setting policies..."
    $statusLabel.ForeColor = [System.Drawing.Color]::Orange
    $changesMade = $false
    $errorOccurred = $false

    $scopesToManage | Where-Object {$_.Settable} | ForEach-Object {
        $scopeName = $_.Name
        $comboBox = $mainForm.Controls.Find("cb$scopeName", $true)[0]
        if ($comboBox -and $comboBox.SelectedItem) {
            $selectedPolicy = $comboBox.SelectedItem.ToString()
            $currentPolicy = $global:currentPolicies[$scopeName]

            if ($selectedPolicy -ne $currentPolicy) {
                try {
                    Write-Host "Setting ExecutionPolicy for Scope '$scopeName' to '$selectedPolicy'..."
                    Set-ExecutionPolicy -ExecutionPolicy $selectedPolicy -Scope $scopeName -Force -ErrorAction Stop
                    $changesMade = $true
                    Write-Host "Successfully set for Scope '$scopeName'."
                } catch {
                    $statusLabel.Text = "Error setting for Scope ${scopeName}: $($_.Exception.Message)"
                    $statusLabel.ForeColor = [System.Drawing.Color]::Red
                    $errorOccurred = $true
                    Write-Warning "Error setting ExecutionPolicy for Scope '$scopeName': $($_.Exception.Message)"
                }
            }
        }
    }

    Update-PolicyDisplay

    if ($errorOccurred) {
        # Error message is already set in the catch block
    } elseif ($changesMade) {
        $statusLabel.Text = "Execution policies successfully updated."
        $statusLabel.ForeColor = [System.Drawing.Color]::Green
    } else {
        $statusLabel.Text = "No changes made."
        $statusLabel.ForeColor = [System.Drawing.Color]::Blue
    }
})

# On Form Load, get current policies and display
$mainForm.Add_Load({
    Update-PolicyDisplay
})

# Show GUI
$mainForm.ShowDialog() | Out-Null

# Clean up
$mainForm.Dispose()
