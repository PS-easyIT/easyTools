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

# ToolTip Object for providing hints
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 10000 # Keep tooltip visible for 10 seconds
$toolTip.InitialDelay = 500   # Show tooltip after 0.5 seconds
$toolTip.ReshowDelay = 500

# Define Tooltip Texts
$scopeTooltips = @{
    "MachinePolicy" = "MachinePolicy: Wird durch Gruppenrichtlinien für alle Benutzer dieses Computers festgelegt. Diese Einstellung hat Vorrang vor allen anderen Bereichen und kann hier nicht geändert werden."
    "UserPolicy"    = "UserPolicy: Wird durch Gruppenrichtlinien für den aktuellen Benutzer festgelegt. Diese Einstellung hat Vorrang vor LocalMachine und CurrentUser und kann hier nicht geändert werden."
    "Process"       = "Process: Gilt nur für die aktuelle PowerShell-Sitzung. Die Einstellung geht beim Schließen der Sitzung verloren und hat Vorrang vor CurrentUser und LocalMachine (außer wenn diese durch GPO gesperrt sind)."
    "CurrentUser"   = "CurrentUser: Gilt nur für den aktuellen Benutzer und wird in dessen Registrierungsabschnitt gespeichert. Hat Vorrang vor LocalMachine (außer wenn diese durch GPO gesperrt sind)."
    "LocalMachine"  = "LocalMachine: Standardrichtlinie für alle Benutzer dieses Computers, wenn keine spezifischere Richtlinie (Process, CurrentUser) oder eine GPO-Richtlinie (MachinePolicy, UserPolicy) gilt. Wird im lokalen Maschinen-Registrierungsabschnitt gespeichert."
}

$policyExplanationTooltips = @{
    "Restricted"    = "Restricted: Keine Skripte dürfen ausgeführt werden. PowerShell kann nur im interaktiven Modus verwendet werden."
    "AllSigned"     = "AllSigned: Nur Skripte, die von einem vertrauenswürdigen Herausgeber digital signiert wurden, können ausgeführt werden."
    "RemoteSigned"  = "RemoteSigned: Lokal erstellte Skripte können ausgeführt werden. Skripte, die aus dem Internet heruntergeladen wurden (als 'Zone:Internet' markiert), müssen von einem vertrauenswürdigen Herausgeber signiert sein."
    "Unrestricted"  = "Unrestricted: Alle Skripte können ausgeführt werden. Eine Warnung wird für nicht signierte Skripte aus dem Internet angezeigt. Vorsicht ist geboten."
    "Bypass"        = "Bypass: Nichts wird blockiert und es gibt keine Warnungen oder Eingabeaufforderungen. Diese Richtlinie ist für Situationen gedacht, in denen ein PowerShell-Skript in eine größere Anwendung integriert ist oder PowerShell die Grundlage für ein Programm ist, das über ein eigenes Sicherheitsmodell verfügt."
    "Undefined"     = "Undefined: Es ist keine Ausführungsrichtlinie für diesen Bereich explizit festgelegt. Die effektive Richtlinie wird durch die Rangfolge der anderen Bereiche bestimmt. Wenn alle Bereiche 'Undefined' sind, ist die Standardrichtlinie 'Restricted'."
}

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
            $currentPolicyValue = $global:currentPolicies[$scopeName]
            $label.Text = $currentPolicyValue
            # Set tooltip for the current policy value label
            if ($policyExplanationTooltips.ContainsKey($currentPolicyValue)) {
                $toolTip.SetToolTip($label, $policyExplanationTooltips[$currentPolicyValue])
            } else {
                $toolTip.SetToolTip($label, "Keine spezifische Erklärung für diese Richtlinie verfügbar.")
            }
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
    # Set tooltip for the scope display label
    if ($scopeTooltips.ContainsKey($scopeName)) {
        $toolTip.SetToolTip($labelScope, $scopeTooltips[$scopeName])
    }

    # Label for current Policy (Column 2)
    $lblCurrentPolicy = New-Object System.Windows.Forms.Label
    $lblCurrentPolicy.Name = "lblCurrent$scopeName"
    $lblCurrentPolicy.Text = "loading..."
    $lblCurrentPolicy.Location = New-Object System.Drawing.Point(360, $yPos)
    $lblCurrentPolicy.AutoSize = $true
    $lblCurrentPolicy.Font = $boldLabelFont
    $mainForm.Controls.Add($lblCurrentPolicy)
    # Tooltip for lblCurrentPolicy will be set in Update-PolicyDisplay based on its content

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
        $toolTip.SetToolTip($cbPolicy, "Wählen Sie hier die neue Richtlinie für den Bereich '$scopeDisplay'.")
    } else {
        $lblNotSettable = New-Object System.Windows.Forms.Label
        $lblNotSettable.Text = "(Controlled by Group Policy)"
        $lblNotSettable.Location = New-Object System.Drawing.Point(540, $yPos)
        $lblNotSettable.AutoSize = $true
        $lblNotSettable.Font = $labelFont
        $lblNotSettable.ForeColor = [System.Drawing.Color]::DarkSlateGray
        $mainForm.Controls.Add($lblNotSettable)
        $toolTip.SetToolTip($lblNotSettable, "Diese Richtlinie wird durch eine Gruppenrichtlinie (GPO) verwaltet und kann mit diesem Tool nicht geändert werden.")
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
$toolTip.SetToolTip($lblEffectivePolicy, "Die effektive Ausführungsrichtlinie ist das Ergebnis der Richtlinien aller Bereiche, basierend auf ihrer Rangfolge. Dies ist die Richtlinie, die für die aktuelle PowerShell-Sitzung tatsächlich gilt.")

$lblEffectivePolicyValue = New-Object System.Windows.Forms.Label
$lblEffectivePolicyValue.Name = "lblEffectivePolicyValue"
$lblEffectivePolicyValue.Text = "loading..."
$lblEffectivePolicyValue.Location = New-Object System.Drawing.Point(360, ($yPos + 10))
$lblEffectivePolicyValue.AutoSize = $true
$lblEffectivePolicyValue.Font = $boldLabelFont
$mainForm.Controls.Add($lblEffectivePolicyValue)
$toolTip.SetToolTip($lblEffectivePolicyValue, "Der Wert der aktuell wirksamen Ausführungsrichtlinie.")

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
$applyButton.Add_Click({
    
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

# SIG # Begin signature block
# MIIbywYJKoZIhvcNAQcCoIIbvDCCG7gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC6Ndt12pzvWN8M
# 3DPVeuxO3bVvyHN/v3WAPkpi/Cb1u6CCFhcwggMQMIIB+KADAgECAhB3jzsyX9Cg
# jEi+sBC2rBMTMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFVBoaW5JVC1QU3Nj
# cmlwdHNfU2lnbjAeFw0yNTA3MDUwODI4MTZaFw0yNzA3MDUwODM4MTZaMCAxHjAc
# BgNVBAMMFVBoaW5JVC1QU3NjcmlwdHNfU2lnbjCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBALmz3o//iDA5MvAndTjGX7/AvzTSACClfuUR9WYK0f6Ut2dI
# mPxn+Y9pZlLjXIpZT0H2Lvxq5aSI+aYeFtuJ8/0lULYNCVT31Bf+HxervRBKsUyi
# W9+4PH6STxo3Pl4l56UNQMcWLPNjDORWRPWHn0f99iNtjI+L4tUC/LoWSs3obzxN
# 3uTypzlaPBxis2qFSTR5SWqFdZdRkcuI5LNsJjyc/QWdTYRrfmVqp0QrvcxzCv8u
# EiVuni6jkXfiE6wz+oeI3L2iR+ywmU6CUX4tPWoS9VTtmm7AhEpasRTmrrnSg20Q
# jiBa1eH5TyLAH3TcYMxhfMbN9a2xDX5pzM65EJUCAwEAAaNGMEQwDgYDVR0PAQH/
# BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQO7XOqiE/EYi+n
# IaR6YO5M2MUuVTANBgkqhkiG9w0BAQsFAAOCAQEAjYOKIwBu1pfbdvEFFaR/uY88
# peKPk0NnvNEc3dpGdOv+Fsgbz27JPvItITFd6AKMoN1W48YjQLaU22M2jdhjGN5i
# FSobznP5KgQCDkRsuoDKiIOTiKAAknjhoBaCCEZGw8SZgKJtWzbST36Thsdd/won
# ihLsuoLxfcFnmBfrXh3rTIvTwvfujob68s0Sf5derHP/F+nphTymlg+y4VTEAijk
# g2dhy8RAsbS2JYZT7K5aEJpPXMiOLBqd7oTGfM7y5sLk2LIM4cT8hzgz3v5yPMkF
# H2MdR//K403e1EKH9MsGuGAJZddVN8ppaiESoPLoXrgnw2SY5KCmhYw1xRFdjTCC
# BY0wggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290
# IENBMB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkq
# hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE9
# 8orYWcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9S
# H8aeFaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g
# 1ckgHWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RY
# jgwrt0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgD
# EI3Y1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNA
# vwjXWkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDg
# ohIbZpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQA
# zH0clcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOk
# GLimdwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHF
# ynIWIgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gd
# LfXZqbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYE
# FOzX44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUH
# MAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDov
# L2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# dDBFBgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGln
# aUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkq
# hkiG9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7
# IviHGmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/5
# 9PesMHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0
# POz3A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISf
# b8rbII01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhU
# LSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBq4wggSWoAMCAQICEAc2
# N7ckVHzYR6z9KGYqXlswDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTIyMDMyMzAwMDAw
# MFoXDTM3MDMyMjIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYg
# U0hBMjU2IFRpbWVTdGFtcGluZyBDQTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBAMaGNQZJs8E9cklRVcclA8TykTepl1Gh1tKD0Z5Mom2gsMyD+Vr2EaFE
# FUJfpIjzaPp985yJC3+dH54PMx9QEwsmc5Zt+FeoAn39Q7SE2hHxc7Gz7iuAhIoi
# GN/r2j3EF3+rGSs+QtxnjupRPfDWVtTnKC3r07G1decfBmWNlCnT2exp39mQh0YA
# e9tEQYncfGpXevA3eZ9drMvohGS0UvJ2R/dhgxndX7RUCyFobjchu0CsX7LeSn3O
# 9TkSZ+8OpWNs5KbFHc02DVzV5huowWR0QKfAcsW6Th+xtVhNef7Xj3OTrCw54qVI
# 1vCwMROpVymWJy71h6aPTnYVVSZwmCZ/oBpHIEPjQ2OAe3VuJyWQmDo4EbP29p7m
# O1vsgd4iFNmCKseSv6De4z6ic/rnH1pslPJSlRErWHRAKKtzQ87fSqEcazjFKfPK
# qpZzQmiftkaznTqj1QPgv/CiPMpC3BhIfxQ0z9JMq++bPf4OuGQq+nUoJEHtQr8F
# nGZJUlD0UfM2SU2LINIsVzV5K6jzRWC8I41Y99xh3pP+OcD5sjClTNfpmEpYPtMD
# iP6zj9NeS3YSUZPJjAw7W4oiqMEmCPkUEBIDfV8ju2TjY+Cm4T72wnSyPx4Jduyr
# XUZ14mCjWAkBKAAOhFTuzuldyF4wEr1GnrXTdrnSDmuZDNIztM2xAgMBAAGjggFd
# MIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBS6FtltTYUvcyl2mi91
# jGogj57IbzAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8B
# Af8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBpMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKG
# NWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290
# RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQC
# MAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAfVmOwJO2b5ipRCIBfmbW
# 2CFC4bAYLhBNE88wU86/GPvHUF3iSyn7cIoNqilp/GnBzx0H6T5gyNgL5Vxb122H
# +oQgJTQxZ822EpZvxFBMYh0MCIKoFr2pVs8Vc40BIiXOlWk/R3f7cnQU1/+rT4os
# equFzUNf7WC2qk+RZp4snuCKrOX9jLxkJodskr2dfNBwCnzvqLx1T7pa96kQsl3p
# /yhUifDVinF2ZdrM8HKjI/rAJ4JErpknG6skHibBt94q6/aesXmZgaNWhqsKRcnf
# xI2g55j7+6adcq/Ex8HBanHZxhOACcS2n82HhyS7T6NJuXdmkfFynOlLAlKnN36T
# U6w7HQhJD5TNOXrd/yVjmScsPT9rp/Fmw0HNT7ZAmyEhQNC3EyTN3B14OuSereU0
# cZLXJmvkOHOrpgFPvT87eK1MrfvElXvtCl8zOYdBeHo46Zzh3SP9HSjTx/no8Zhf
# +yvYfvJGnXUsHicsJttvFXseGYs2uJPU5vIXmVnKcPA3v5gA3yAWTyf7YGcWoWa6
# 3VXAOimGsJigK+2VQbc61RWYMbRiCQ8KvYHZE/6/pNHzV9m8BPqC3jLfBInwAM1d
# wvnQI38AC+R2AibZ8GV2QqYphwlHK+Z/GqSFD/yYlvZVVCsfgPrA8g4r5db7qS9E
# FUrnEw4d2zc4GqEr9u3WfPwwgga8MIIEpKADAgECAhALrma8Wrp/lYfG+ekE4zME
# MA0GCSqGSIb3DQEBCwUAMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2Vy
# dCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNI
# QTI1NiBUaW1lU3RhbXBpbmcgQ0EwHhcNMjQwOTI2MDAwMDAwWhcNMzUxMTI1MjM1
# OTU5WjBCMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQxIDAeBgNVBAMT
# F0RpZ2lDZXJ0IFRpbWVzdGFtcCAyMDI0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8A
# MIICCgKCAgEAvmpzn/aVIauWMLpbbeZZo7Xo/ZEfGMSIO2qZ46XB/QowIEMSvgjE
# dEZ3v4vrrTHleW1JWGErrjOL0J4L0HqVR1czSzvUQ5xF7z4IQmn7dHY7yijvoQ7u
# jm0u6yXF2v1CrzZopykD07/9fpAT4BxpT9vJoJqAsP8YuhRvflJ9YeHjes4fduks
# THulntq9WelRWY++TFPxzZrbILRYynyEy7rS1lHQKFpXvo2GePfsMRhNf1F41nyE
# g5h7iOXv+vjX0K8RhUisfqw3TTLHj1uhS66YX2LZPxS4oaf33rp9HlfqSBePejlY
# eEdU740GKQM7SaVSH3TbBL8R6HwX9QVpGnXPlKdE4fBIn5BBFnV+KwPxRNUNK6lY
# k2y1WSKour4hJN0SMkoaNV8hyyADiX1xuTxKaXN12HgR+8WulU2d6zhzXomJ2Ple
# I9V2yfmfXSPGYanGgxzqI+ShoOGLomMd3mJt92nm7Mheng/TBeSA2z4I78JpwGpT
# RHiT7yHqBiV2ngUIyCtd0pZ8zg3S7bk4QC4RrcnKJ3FbjyPAGogmoiZ33c1HG93V
# p6lJ415ERcC7bFQMRbxqrMVANiav1k425zYyFMyLNyE1QulQSgDpW9rtvVcIH7Wv
# G9sqYup9j8z9J1XqbBZPJ5XLln8mS8wWmdDLnBHXgYly/p1DhoQo5fkCAwEAAaOC
# AYswggGHMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQM
# MAoGCCsGAQUFBwMIMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATAf
# BgNVHSMEGDAWgBS6FtltTYUvcyl2mi91jGogj57IbzAdBgNVHQ4EFgQUn1csA3cO
# KBWQZqVjXu5Pkh92oFswWgYDVR0fBFMwUTBPoE2gS4ZJaHR0cDovL2NybDMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVTdGFt
# cGluZ0NBLmNybDCBkAYIKwYBBQUHAQEEgYMwgYAwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBYBggrBgEFBQcwAoZMaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVT
# dGFtcGluZ0NBLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAPa0eH3aZW+M4hBJH2UOR
# 9hHbm04IHdEoT8/T3HuBSyZeq3jSi5GXeWP7xCKhVireKCnCs+8GZl2uVYFvQe+p
# PTScVJeCZSsMo1JCoZN2mMew/L4tpqVNbSpWO9QGFwfMEy60HofN6V51sMLMXNTL
# fhVqs+e8haupWiArSozyAmGH/6oMQAh078qRh6wvJNU6gnh5OruCP1QUAvVSu4kq
# VOcJVozZR5RRb/zPd++PGE3qF1P3xWvYViUJLsxtvge/mzA75oBfFZSbdakHJe2B
# VDGIGVNVjOp8sNt70+kEoMF+T6tptMUNlehSR7vM+C13v9+9ZOUKzfRUAYSyyEmY
# tsnpltD/GWX8eM70ls1V6QG/ZOB6b6Yum1HvIiulqJ1Elesj5TMHq8CWT/xrW7tw
# ipXTJ5/i5pkU5E16RSBAdOp12aw8IQhhA/vEbFkEiF2abhuFixUDobZaA0VhqAsM
# HOmaT3XThZDNi5U2zHKhUs5uHHdG6BoQau75KiNbh0c+hatSF+02kULkftARjsyE
# pHKsF7u5zKRbt5oK5YGwFvgc4pEVUNytmB3BpIiowOIIuDgP5M9WArHYSAR16gc0
# dP2XdkMEP5eBsX7bf/MGN4K3HP50v/01ZHo/Z5lGLvNwQ7XHBx1yomzLP8lx4Q1z
# ZKDyHcp4VQJLu2kWTsKsOqQxggUKMIIFBgIBATA0MCAxHjAcBgNVBAMMFVBoaW5J
# VC1QU3NjcmlwdHNfU2lnbgIQd487Ml/QoIxIvrAQtqwTEzANBglghkgBZQMEAgEF
# AKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgor
# BgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3
# DQEJBDEiBCAKsh5spB56nQNNxNxOrpGo3JOCs5CBSdU67cKq0eRa4DANBgkqhkiG
# 9w0BAQEFAASCAQAtE0umUmgZVZ0NQo4KbUtZc8eZaQvcDWrLUsukaYs2SPMyklG2
# J+UP3XqysFDPAmfx0xyYJsWCwZAO1qDRFlsjShIni6mlUFyImEkEne/O7ckNRHyS
# M3k6f0coH8aMSWaz80vnCJdN/QPwjbKYtWumF/64yVBwoza9FyW4jP6wjM2RUm5a
# dZnibwX8gxgX46PJU7d+015zWyXKzKR8duuVnX5dxFWMce845EPToXhUPZYvPoVr
# cEMdK7l37nKuV7hT3WZDPyaVp48cabKYI2sJM+0mejI1/80h3cI41drdQsfbUbcX
# JhPxneaev2+k/IghfzNNE6aZVGBryuseArUioYIDIDCCAxwGCSqGSIb3DQEJBjGC
# Aw0wggMJAgEBMHcwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2
# IFRpbWVTdGFtcGluZyBDQQIQC65mvFq6f5WHxvnpBOMzBDANBglghkgBZQMEAgEF
# AKBpMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI1
# MDcwNTEwMTIwOFowLwYJKoZIhvcNAQkEMSIEIDSkwKqoWb1dvsp2Q3KhORvTLuNs
# 9mw9VQcRqeBKhg+dMA0GCSqGSIb3DQEBAQUABIICAH7/anSIuHXhD5z+FP0/bmeL
# 9s/gxhB1GMNRweX8z73fGQ2cHunpF9c6iEkauxfmxU1jHOdbZhFyC+94U+YWuTU7
# ozO1zNlKWQQJwdHWWvBlDWQ+ANh+FiIAT+lddJEyisIwCCYqH6rI4YXMUAgSnnsX
# U+j6s7SmWp9e1TX1LUCvU2is/9wfzf2+tN5Eybg9EmHU4M1Vhe9FP7+2WwMzVrq3
# kmI8gklPzISxRzNSGp7OpdcZxE2Pr/bjqWG2Nz7ql8wRQPK8bq+pWnpdbx6qH5UG
# lFCfdS5K4R+r5R+0a4g8gsNKi+qiPtrBJpc6QAliWPazniaqkp7DzZa54SBsKEXy
# 9EA+KnXIpbMIqATiOE1ZkQFZ43arCHUgKH6ssY08hnAn+pounzJ6doQ2WXnBZPKA
# iHEgq4zA3hQr/kEe5zy2RH2BtLAV96P7ECxebye3Mc3gYwIP7dHcvRK/X9GPf0Ko
# fEjr680s+0dkUfl+r/n6lQtynujUJIeF9giPsg2IEkk/TWtWJtnBDIFHAaDtUjEm
# iErhHAam0ZvNXfW/IvEZ7m7vL/p/oD56qhVbeV12zNSRWpF9rHAvl/VklZMuzkuG
# zxzP1XaODvEkP8HoJrWInYpbcQ8pPWXs0M6hL2etMuX0zkC3U48AszccOPZWiZTg
# 7eayuqYfItxmiUG++3DU
# SIG # End signature block
