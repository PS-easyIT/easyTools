#############################################
# Teil 1: Konfiguration, Logging und Utility-Funktionen
#############################################

# Benötigte Assemblies laden (für WPF und Windows Forms)
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Logging initialisieren
$script:logFolder = Join-Path -Path $PSScriptRoot -ChildPath "Logs"
$script:logFile = Join-Path -Path $script:logFolder -ChildPath "easySTARTUP_Setup.log"

if (-not (Test-Path -Path $script:logFolder)) {
    New-Item -Path $script:logFolder -ItemType Directory -Force | Out-Null
}

# Zentrale Logging-Funktion
function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Type = "Info"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Type] $Message"
    Add-Content -Path $script:logFile -Value $logEntry -Encoding UTF8
    
    switch ($Type) {
        "Info"    { Write-Host $logEntry -ForegroundColor Gray }
        "Warning" { Write-Host $logEntry -ForegroundColor Yellow }
        "Error"   { Write-Host $logEntry -ForegroundColor Red }
        "Success" { Write-Host $logEntry -ForegroundColor Green }
    }
}

# Konsolidierte Funktion zur Aktualisierung der Log-Anzeige in der GUI
function Update-LogDisplay {
    param (
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Type = "Info"
    )
    Write-Log -Message $Message -Type $Type

    $timestamp = Get-Date -Format "HH:mm:ss"
    $logEntry = "[$timestamp] "
    switch ($Type) {
        "Info"    { $logEntry += "[INFO] " }
        "Warning" { $logEntry += "[WARNING] " }
        "Error"   { $logEntry += "[ERROR] " }
        "Success" { $logEntry += "[SUCCESS] " }
    }
    $logEntry += $Message

    if ($Global:txtLog) {
        $Global:txtLog.AppendText("$logEntry`n")
        $Global:txtLog.ScrollToEnd()
    }
}

# Definition der Module
$requiredModules = @(
    @{
        Name         = "ExchangeOnlineManagement"
        MinVersion   = "3.8.0"
        Description  = "Exchange Online Verwaltung"
        Webseite     = "https://www.powershellgallery.com/packages/ExchangeOnlineManagement/"
        Install      = "Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser"
    },
    @{
        Name         = "Microsoft.Graph"
        MinVersion   = "2.28.0"
        Description  = "Microsoft Graph API"
        Webseite     = "https://www.powershellgallery.com/packages/Microsoft.Graph/"
        Install      = "Install-Module -Name Microsoft.Graph -Scope CurrentUser"
    },
    @{
        Name         = "PowerShellGet"
        MinVersion   = "2.2.5"
        Description  = "PS Module Management"
        Webseite     = "https://www.powershellgallery.com/packages/PowerShellGet/"
        Install      = "Install-Module -Name PowerShellGet -Scope CurrentUser"
    },
    @{
        Name         = "Microsoft.Graph.Beta"
        MinVersion   = "2.28.0"
        Description  = "Microsoft Graph Beta"
        Webseite     = "https://www.powershellgallery.com/packages/Microsoft.Graph.Beta/"
        Install      = "Install-Module -Name Microsoft.Graph.Beta -Scope CurrentUser"
    },
    @{
        Name         = "Microsoft.Entra"
        MinVersion   = "1.0.7"
        Description  = "Entra ID PowerShell"
        Webseite     = "https://www.powershellgallery.com/packages/Microsoft.Entra/"
        Install      = "Install-Module -Name Microsoft.Entra -Scope CurrentUser"
    },
    @{
        Name         = "Microsoft.Entra.Beta"
        MinVersion   = "1.0.7"
        Description  = "Entra ID PowerShell (Beta)"
        Webseite     = "https://www.powershellgallery.com/packages/Microsoft.Entra.Beta/"
        Install      = "Install-Module -Name Microsoft.Entra.Beta -Scope CurrentUser"
    },
    @{
        Name         = "PnP.PowerShell"
        MinVersion   = "3.1.0"
        Description  = "SharePoint PnP"
        Webseite     = "https://www.powershellgallery.com/packages/PnP.PowerShell/"
        Install      = "Install-Module -Name PnP.PowerShell -Scope CurrentUser"
    },
    @{
        Name         = "AzureAD"
        MinVersion   = "2.0.2.182"
        Description  = "Azure AD (legacy)"
        Webseite     = "https://www.powershellgallery.com/packages/AzureAD/"
        Install      = "Install-Module -Name AzureAD -Scope CurrentUser"
    },
    @{
        Name         = "MSOnline"
        MinVersion   = "1.1.183.81"
        Description  = "Azure AD V1 (legacy)"
        Webseite     = "https://www.powershellgallery.com/packages/MSOnline/"
        Install      = "Install-Module -Name MSOnline -Scope CurrentUser"
    },
    @{
        Name         = "Microsoft.Online.SharePoint.PowerShell"
        MinVersion   = "16.0.26017.12000"
        Description  = "SharePoint Online"
        Webseite     = "https://www.powershellgallery.com/packages/Microsoft.Online.SharePoint.PowerShell/"
        Install      = "Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser"
    }
)

# Funktion zur Prüfung, ob ein Modul installiert ist
function Test-ModuleInstalled {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModuleName,
        [Parameter(Mandatory = $false)]
        [version]$MinVersion
    )
    $module = Get-Module -ListAvailable -Name $ModuleName
    if (-not $module) {
        return $false
    }
    if ($MinVersion) {
        $latestVersion = $module | Sort-Object Version -Descending | Select-Object -First 1
        if ($latestVersion.Version -lt $MinVersion) {
            return $false
        }
    }
    return $true
}

# Funktion zur Installation eines erforderlichen Moduls
function Install-RequiredModule {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModuleName,
        [Parameter(Mandatory = $false)]
        [version]$MinVersion
    )
    try {
        if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
            Write-Log "Installing NuGet Provider..." -Type "Info"
            Install-PackageProvider -Name NuGet -Force -Scope CurrentUser | Out-Null
        }
        if ((Get-PSRepository -Name PSGallery).InstallationPolicy -ne "Trusted") {
            Write-Log "Setting PSGallery as trusted source..." -Type "Info"
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
        }
        $installParams = @{
            Name         = $ModuleName
            Force        = $true
            AllowClobber = $true
            Scope        = "CurrentUser"
            Repository   = "PSGallery"
        }
        if ($MinVersion) {
            $installParams.MinimumVersion = $MinVersion
            Write-Log "Installing $ModuleName (minimum version $MinVersion)..." -Type "Info"
        }
        else {
            Write-Log "Installing $ModuleName..." -Type "Info"
        }
        Install-Module @installParams -ErrorAction Stop
        Write-Log "Module $ModuleName successfully installed." -Type "Success"
        return $true
    }
    catch {
        Write-Log "Error installing $ModuleName - $($_.Exception.Message)" -Type "Error"
        return $false
    }
}

# Funktion zur Ermittlung von Moduldaten
function Get-ModuleData {
    $moduleData = @()
    foreach ($module in $requiredModules) {
        $moduleName    = $module.Name
        $minVersion    = $module.MinVersion
        $description   = $module.Description
        $webseite      = $module.Webseite
        $installCmd    = $module.Install

        $installed       = Test-ModuleInstalled -ModuleName $moduleName -MinVersion $minVersion
        $installedModule = Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1
        $installedVersion = if ($installedModule) { $installedModule.Version.ToString() } else { "Not installed" }
        
        $installEnabled = -not $installed
        
        $moduleData += [PSCustomObject]@{
            Name             = $moduleName
            Description      = $description
            MinVersion       = $minVersion
            InstalledVersion = $installedVersion
            InstallEnabled   = $installEnabled
            Webseite         = $webseite
            InstallCmd       = $installCmd
        }
    }
    return $moduleData
}

# Funktion zur Aktualisierung der Module
function Check-AllModules {
    Update-LogDisplay -Message "Checking modules..." -Type "Info"
    $moduleData = Get-ModuleData
    if ($Global:lvModules) {
        $Global:lvModules.ItemsSource = $moduleData
        Update-LogDisplay -Message "Modules loaded in ListView." -Type "Info"
    }
    else {
        Update-LogDisplay -Message "ListView element not found!" -Type "Error"
    }
}

# ExecutionPolicy-Funktionen
function Get-CurrentExecutionPolicies {
    $policies = Get-ExecutionPolicy -List | Select-Object Scope, ExecutionPolicy
    $policyData = @()
    
    $scopeDescriptions = @{
        "MachinePolicy" = "MachinePolicy (GPO - All users of this computer)"
        "UserPolicy"    = "UserPolicy (GPO - Current user)"
        "Process"       = "Process (Only for current PowerShell session)"
        "CurrentUser"   = "CurrentUser (Settings for current user)"
        "LocalMachine"  = "LocalMachine (Default for all users of this computer)"
    }
    
    foreach ($policy in $policies) {
        $scopeName = $policy.Scope.ToString()
        $canModify = $scopeName -notin @("MachinePolicy", "UserPolicy")
        
        $policyData += [PSCustomObject]@{
            Scope = $scopeName
            Description = $scopeDescriptions[$scopeName]
            CurrentPolicy = $policy.ExecutionPolicy.ToString()
            CanModify = $canModify
            PolicyOptions = @("Restricted", "AllSigned", "RemoteSigned", "Unrestricted", "Bypass", "Undefined")
        }
    }
    
    # Add effective policy
    try {
        $effectivePolicy = Get-ExecutionPolicy
        $policyData += [PSCustomObject]@{
            Scope = "Effective"
            Description = "Effective Policy (current session)"
            CurrentPolicy = $effectivePolicy.ToString()
            CanModify = $false
            PolicyOptions = @()
        }
    }
    catch {
        Update-LogDisplay -Message "Error getting effective policy: $($_.Exception.Message)" -Type "Warning"
    }
    
    return $policyData
}

function Update-ExecutionPolicyDisplay {
    $policyData = Get-CurrentExecutionPolicies
    if ($Global:lvExecutionPolicies) {
        $Global:lvExecutionPolicies.ItemsSource = $policyData
        Update-LogDisplay -Message "Execution policies updated." -Type "Info"
    }
}

function Set-ExecutionPolicyHandler {
    param($scope, $policy)
    
    # Check for admin rights if needed
    if ($scope -eq "LocalMachine") {
        if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            $result = [System.Windows.MessageBox]::Show(
                "Admin rights are required to modify LocalMachine scope. Restart as administrator?",
                "Admin Rights Required",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning)
            
            if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                try {
                    Start-Process PowerShell -Verb RunAs -ArgumentList ("-File `"{0}`"" -f $MyInvocation.MyCommand.Path)
                    $Global:window.Close()
                }
                catch {
                    Update-LogDisplay -Message "Error restarting as admin: $($_.Exception.Message)" -Type "Error"
                }
            }
            return
        }
    }
    
    try {
        Update-LogDisplay -Message "Setting ExecutionPolicy for scope '$scope' to '$policy'..." -Type "Info"
        Set-ExecutionPolicy -ExecutionPolicy $policy -Scope $scope -Force -ErrorAction Stop
        Update-LogDisplay -Message "ExecutionPolicy successfully set for scope '$scope'." -Type "Success"
        Update-ExecutionPolicyDisplay
    }
    catch {
        Update-LogDisplay -Message "Error setting ExecutionPolicy for scope '$scope': $($_.Exception.Message)" -Type "Error"
    }
}

# Funktion zum Öffnen von Web Links
function Open-WebLink {
    param($url)
    try {
        if (-not [string]::IsNullOrEmpty($url)) {
            Start-Process $url
            Update-LogDisplay -Message "Opening web link: $url" -Type "Info"
        } else {
            Update-LogDisplay -Message "No web link available for this module." -Type "Warning"
        }
    }
    catch {
        Update-LogDisplay -Message "Error opening web link: $($_.Exception.Message)" -Type "Error"
    }
}

function Install-ModuleHandler {
    param($buttonSender, $e)
    $button = [System.Windows.Controls.Button]$buttonSender
    $moduleName = $button.Tag.ToString()
    
    $moduleInfo = $requiredModules | Where-Object { $_.Name -eq $moduleName } | Select-Object -First 1
    
    if ($moduleInfo) {
        Update-LogDisplay -Message "Starting installation of $moduleName..." -Type "Info"
        $button.IsEnabled = $false
        $button.Content = "Installing..."
        
        $success = Install-RequiredModule -ModuleName $moduleName -MinVersion $moduleInfo.MinVersion
        
        if ($success) {
            Update-LogDisplay -Message "Module $moduleName successfully installed." -Type "Success"
        } else {
            Update-LogDisplay -Message "Error installing module $moduleName." -Type "Error"
        }
        
        Check-AllModules
    }
}

# XAML für die GUI (ohne Connect-Funktionen, mit ExecutionPolicy)
$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="easyIT PowerShell Setup &amp; Module Manager" 
    Height="1000" 
    Width="1400"
    WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style TargetType="Button" x:Key="InstallButtonStyle">
            <Setter Property="Padding" Value="5,2"/>
            <Setter Property="Background" Value="#EFEFEF"/>
        </Style>
        <Style TargetType="Button" x:Key="PolicyButtonStyle">
            <Setter Property="Padding" Value="5,2"/>
            <Setter Property="Background" Value="#E3F2FD"/>
            <Setter Property="FontSize" Value="11"/>
        </Style>
        <Style TargetType="Button" x:Key="ScriptButtonStyle">
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Margin" Value="0,2"/>
            <Setter Property="Background" Value="#E3F2FD"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <!-- Header -->
        <Border Grid.Row="0" Background="#0078D4" Padding="15,12">
            <Grid>
                <StackPanel>
                    <TextBlock Text="easyIT - PowerShell Module Manager &amp; Script Launcher" FontSize="20" FontWeight="Bold" Foreground="White"/>
                    <TextBlock Text="Manage PowerShell modules and execution policies | Launch scripts from subfolders" FontSize="12" Foreground="White" Margin="0,3,0,0"/>
                </StackPanel>
            </Grid>
        </Border>
        <!-- Content with Split Layout -->
        <Grid Grid.Row="1" Margin="15">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="2.2*"/> <!-- Status area: slightly larger -->
                <ColumnDefinition Width="1*"/> <!-- Script selection: 1/3 -->
            </Grid.ColumnDefinitions>
            <!-- Left side - Status area -->
            <Grid Grid.Column="0" Margin="0,0,8,0">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <!-- Module Check Section -->
                <GroupBox Grid.Row="0" Header="PowerShell Modules" Margin="0,0,0,8" Padding="8">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <ListView Grid.Row="0" x:Name="lvModules" Margin="0,5,0,5" BorderThickness="1" Height="260">
                            <ListView.View>
                                <GridView>
                                    <GridViewColumn Header="Module" Width="430"> <!-- Adjusted width -->
                                        <GridViewColumn.CellTemplate>
                                            <DataTemplate>
                                                <TextBlock Text="{Binding Name}" ToolTip="{Binding Description}"/>
                                            </DataTemplate>
                                        </GridViewColumn.CellTemplate>
                                    </GridViewColumn>
                                    <GridViewColumn Header="Act. Ver." DisplayMemberBinding="{Binding MinVersion}" Width="120"/>
                                    <GridViewColumn Header="Inst. Ver." DisplayMemberBinding="{Binding InstalledVersion}" Width="120"/>
                                    <GridViewColumn Header="Install" Width="100">
                                        <GridViewColumn.CellTemplate>
                                            <DataTemplate>
                                                <Button Content="Install/Update" 
                                                        Style="{StaticResource InstallButtonStyle}"
                                                        IsEnabled="{Binding InstallEnabled}"
                                                        Tag="{Binding Name}"
                                                        FontSize="9"/>
                                            </DataTemplate>
                                        </GridViewColumn.CellTemplate>
                                    </GridViewColumn>
                                    <GridViewColumn Header="Web Link" Width="80">
                                        <GridViewColumn.CellTemplate>
                                            <DataTemplate>
                                                <Button Content="WEB LINK" 
                                                        Style="{StaticResource InstallButtonStyle}"
                                                        Tag="{Binding Webseite}"
                                                        FontSize="9"
                                                        Background="#E8F4FD"/>
                                            </DataTemplate>
                                        </GridViewColumn.CellTemplate>
                                    </GridViewColumn>
                                </GridView>
                            </ListView.View>
                        </ListView>
                        <Button Grid.Row="1" x:Name="btnRefreshModules" Content="Refresh Modules" HorizontalAlignment="Left" Width="180" Height="22" Margin="0,5,0,0" Background="#FFE4FFE1" FontSize="11"/>
                    </Grid>
                </GroupBox>
                <!-- ExecutionPolicy Section -->
                <GroupBox Grid.Row="1" Header="PowerShell Execution Policies" Margin="0,8,0,8" Padding="8">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <ListView Grid.Row="0" x:Name="lvExecutionPolicies" Margin="0,0,0,8" BorderThickness="1" Height="165">
                            <ListView.View>
                                <GridView>
                                    <GridViewColumn Header="Scope" DisplayMemberBinding="{Binding Description}" Width="350"/> <!-- Adjusted width -->
                                    <GridViewColumn Header="Current Policy" DisplayMemberBinding="{Binding CurrentPolicy}" Width="150"/>
                                    <GridViewColumn Header="Modify" Width="375">
                                        <GridViewColumn.CellTemplate>
                                            <DataTemplate>
                                                <StackPanel Orientation="Horizontal">
                                                    <Button Content="Restricted" Style="{StaticResource PolicyButtonStyle}" Margin="1,0" Tag="{Binding Scope}" Width="60" FontSize="9"/>
                                                    <Button Content="AllSigned" Style="{StaticResource PolicyButtonStyle}" Margin="1,0" Tag="{Binding Scope}" Width="60" FontSize="9"/>
                                                    <Button Content="RemoteSigned" Style="{StaticResource PolicyButtonStyle}" Margin="1,0" Tag="{Binding Scope}" Width="75" FontSize="9"/>
                                                    <Button Content="Unrestricted" Style="{StaticResource PolicyButtonStyle}" Margin="1,0" Tag="{Binding Scope}" Width="70" FontSize="9"/>
                                                    <Button Content="Bypass" Style="{StaticResource PolicyButtonStyle}" Margin="1,0" Tag="{Binding Scope}" Width="50" FontSize="9"/>
                                                </StackPanel>
                                            </DataTemplate>
                                        </GridViewColumn.CellTemplate>
                                    </GridViewColumn>
                                </GridView>
                            </ListView.View>
                        </ListView>
                        <Button Grid.Row="1" x:Name="btnRefreshPolicies" Content="Refresh Execution Policies" HorizontalAlignment="Left" Width="180" Height="22" Background="#FFE4FFE1" FontSize="11"/>
                    </Grid>
                </GroupBox>
                <!-- Status Log -->
                <GroupBox Grid.Row="2" Header="Status and Log" Margin="0,8,0,0" Padding="8">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <TextBox x:Name="txtLog" IsReadOnly="True" TextWrapping="Wrap" 
                                 FontFamily="Consolas" FontSize="10" Background="#f5f5f5" BorderThickness="0"/>
                    </ScrollViewer>
                </GroupBox>
            </Grid>
            <!-- Right side - Script selection -->
            <GroupBox Grid.Column="1" Header="Available PowerShell Scripts" Margin="8,0,0,0" Padding="8">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <!-- Script list -->
                    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                        <StackPanel x:Name="spScriptList" Orientation="Vertical" Margin="0,3,0,3">
                            <!-- Script buttons will be added dynamically -->
                        </StackPanel>
                    </ScrollViewer>
                    <!-- Refresh button -->
                    <Button Grid.Row="2" x:Name="btnRefreshScripts" 
                            Content="Refresh Script List" Margin="0,3,0,0" 
                            Height="25" Background="#F0F0F0" FontSize="11"/>
                </Grid>
            </GroupBox>
        </Grid>
        <!-- Footer with Buttons and Info -->
        <Border Grid.Row="2" Background="#f0f0f0" Padding="15,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Linker Bereich: Versionsinformationen -->
                <StackPanel Grid.Column="0" Orientation="Vertical" HorizontalAlignment="Left" VerticalAlignment="Center">
                    <TextBlock Text="Version: 0.0.3" FontSize="9" Foreground="#555555"/>
                    <TextBlock Text="PSscripts.de | Andreas Hepp" FontSize="9" Foreground="#555555" Margin="0,2,0,0"/>
                    <TextBlock Text="© MIT License" FontSize="9" Foreground="#555555" Margin="0,2,0,0"/>
                </StackPanel>

                <!-- Rechter Bereich: Exit Button -->
                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                    <Button x:Name="btnExit" Content="Exit" Width="120" Height="28" Background="#FFFF7272" Margin="8,0,0,0" FontSize="11"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

#############################################
# Teil 2: GUI-Funktionen, Event-Handler und Main Block
#############################################

# Funktion zum Laden der XAML-GUI
function Import-XamlGui {
    try {
        if ([string]::IsNullOrEmpty($xaml)) {
            throw "XAML content is empty"
        }
        
        $xmlReaderSettings = New-Object System.Xml.XmlReaderSettings
        $xmlReaderSettings.IgnoreWhitespace = $true
        $xmlReaderSettings.IgnoreComments = $true
        $xmlReaderSettings.IgnoreProcessingInstructions = $true

        $reader = $null
        $stringReader = $null
        
        try {
            $stringReader = New-Object System.IO.StringReader($xaml)
            $reader = [System.Xml.XmlReader]::Create($stringReader, $xmlReaderSettings)
            
            $window = [System.Windows.Markup.XamlReader]::Load($reader)
            
            $regex = New-Object System.Text.RegularExpressions.Regex('x:Name="([^"]*)"')
            $matches = $regex.Matches($xaml)
            
            foreach ($match in $matches) {
                if ($match.Groups.Count -gt 1) {
                    $name = $match.Groups[1].Value
                    $element = $window.FindName($name)
                    
                    if ($null -ne $element) {
                        Set-Variable -Name $name -Value $element -Scope Global
                    }
                }
            }
            
            return $window
        }
        finally {
            if ($null -ne $reader) { $reader.Close() }
            if ($null -ne $stringReader) { $stringReader.Close() }
        }
    }
    catch {
        Write-Log "Error loading XAML: $($_.Exception.Message)" -Type "Error"
        throw
    }
}

# Funktion zur Aktualisierung der Skriptliste
function Update-ScriptList {
    try {
        Update-LogDisplay -Message "Loading available PowerShell scripts..." -Type "Info"
        
        $spScriptList = $Global:window.FindName("spScriptList")
        if ($null -eq $spScriptList) {
            Update-LogDisplay -Message "Script list element not found!" -Type "Error"
            return
        }
        
        $spScriptList.Children.Clear()
        
        $scriptFolders = @()
        $baseFolder = $PSScriptRoot
        if ([string]::IsNullOrEmpty($baseFolder)) {
            Update-LogDisplay -Message "PSScriptRoot not available, trying alternative path..." -Type "Warning"
            if ($MyInvocation.MyCommand.Path) {
                $baseFolder = Split-Path -Parent $MyInvocation.MyCommand.Path
            }
            else {
                $potentialPaths = @(
                    "D:\2-OneDrive\OneDrive - PhinIT\05-Administratives\10-Scripte\PowerShell\# EASYit",
                    "$env:USERPROFILE\Documents\PowerShell\Scripts",
                    (Get-Location).Path
                )
                foreach ($path in $potentialPaths) {
                    if (Test-Path -Path $path -ErrorAction SilentlyContinue) {
                        $baseFolder = $path
                        Update-LogDisplay -Message "Using found path: $baseFolder" -Type "Info"
                        break
                    }
                }
            }
        }
        
        if ([string]::IsNullOrEmpty($baseFolder)) {
            Update-LogDisplay -Message "No valid script directory found." -Type "Error"
            return
        }
        
        # Ausschlussmuster für Ordnernamen (case-insensitive durch -like)
        $excludeFolderPatterns = @("*old*", "*#old*", "*# old*")

        if (Test-Path -Path $baseFolder -ErrorAction SilentlyContinue) {
            $baseFolderLeafName = (Split-Path -Leaf $baseFolder)
            $skipBaseFolder = $false
            foreach ($pattern in $excludeFolderPatterns) {
                if ($baseFolderLeafName -like $pattern) {
                    $skipBaseFolder = $true
                    Update-LogDisplay -Message "Base folder '$baseFolder' is excluded because its name matches pattern '$pattern'." -Type "Warning"
                    break
                }
            }

            if (-not $skipBaseFolder) {
                Update-LogDisplay -Message "Base path for script search: $baseFolder" -Type "Info"
                $scriptFolders += $baseFolder # Den Basisordner selbst hinzufügen

                # Alle Unterordner rekursiv hinzufügen, die nicht den Ausschlussmustern entsprechen
                try {
                    # Hole ALLE Unterverzeichnisse zuerst
                    $allSubDirectories = Get-ChildItem -Path $baseFolder -Directory -Recurse -ErrorAction SilentlyContinue
                    
                    $validSubDirectories = @()
                    $trimmedBaseFolder = $baseFolder.TrimEnd([System.IO.Path]::DirectorySeparatorChar)

                    foreach ($dirInfoObject in $allSubDirectories) {
                        $currentDirFullName = $dirInfoObject.FullName
                        $isAnySegmentExcluded = $false
                        
                        # Erzeuge den relativen Pfad zum $baseFolder
                        $relativePathString = $null
                        if ($currentDirFullName.Length -gt $trimmedBaseFolder.Length) {
                            # Sicherstellen, dass es ein echter Unterpfad ist
                            if ($currentDirFullName.StartsWith($trimmedBaseFolder + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
                                $relativePathString = $currentDirFullName.Substring($trimmedBaseFolder.Length).TrimStart([System.IO.Path]::DirectorySeparatorChar)
                            } else {
                                # Nicht wirklich ein Unterpfad, überspringen oder Fehler loggen
                                # Update-LogDisplay -Message "Skipping '$currentDirFullName' as it's not a direct subpath of '$trimmedBaseFolder'." -Type "Debug"
                                continue
                            }
                        } else {
                             # $currentDirFullName ist $baseFolder selbst oder kürzer, sollte nicht von Get-ChildItem -Recurse kommen für Unterordner
                            continue
                        }
                        
                        if (-not ([string]::IsNullOrEmpty($relativePathString))) {
                            $pathSegments = $relativePathString.Split([System.IO.Path]::DirectorySeparatorChar)
                            foreach ($segment in $pathSegments) {
                                if ([string]::IsNullOrEmpty($segment)) { continue } # Leere Segmente überspringen
                                foreach ($pattern in $excludeFolderPatterns) {
                                    if ($segment -like $pattern) {
                                        $isAnySegmentExcluded = $true
                                        # Update-LogDisplay -Message "Excluding directory '$currentDirFullName' because path segment '$segment' matches exclusion pattern '$pattern'." -Type "Debug"
                                        break # pattern loop
                                    }
                                }
                                if ($isAnySegmentExcluded) { break } # segment loop
                            }
                        }

                        if (-not $isAnySegmentExcluded) {
                            $validSubDirectories += $currentDirFullName
                        }
                    }
                    
                    $scriptFolders += $validSubDirectories
                    
                    if ($validSubDirectories.Count -gt 0) {
                        Update-LogDisplay -Message "Added $($validSubDirectories.Count) subfolders to search (after manual path segment filtering based on patterns like '*old*', '*#old*', etc.)." -Type "Info"
                    } else {
                         Update-LogDisplay -Message "No subfolders added after manual path segment filtering." -Type "Info"
                    }
                }
                catch {
                    Update-LogDisplay -Message "Error processing subfolders from ${baseFolder}: $($_.Exception.Message)" -Type "Warning"
                }
            }
            # else: $baseFolder wurde übersprungen, $scriptFolders wird nicht mit $baseFolder gefüllt und keine Unterordner werden gesucht.
        }
        else {
            Update-LogDisplay -Message "Base folder $baseFolder does not exist or is not accessible." -Type "Error"
            return 
        }
        
        if ($scriptFolders.Count -gt 0) {
            Update-LogDisplay -Message "Searching for scripts in $($scriptFolders.Count) locations (excluding paths matching patterns like '*old*', '*#old*', etc.)." -Type "Info"
        } else {
            Update-LogDisplay -Message "No script folders to search (base folder might be excluded or contain no valid subfolders after exclusions)." -Type "Warning"
        }
        
        $scripts = @()
        foreach ($folder in $scriptFolders) { # $scriptFolders enthält bereits gefilterte Ordner
            if (-not [string]::IsNullOrEmpty($folder) -and (Test-Path -Path $folder -ErrorAction SilentlyContinue)) {
                try {
                    $currentScriptName = $null
                    if ($MyInvocation.MyCommand.Path) {
                        $currentScriptName = Split-Path -Leaf $MyInvocation.MyCommand.Path
                    }
                    $folderScripts = Get-ChildItem -Path $folder -Filter "*.ps1" -File -ErrorAction Stop | 
                                     Where-Object { $currentScriptName -eq $null -or $_.Name -ne $currentScriptName }
                    if ($folderScripts -and $folderScripts.Count -gt 0) {
                        $scripts += $folderScripts
                        Update-LogDisplay -Message "Found $($folderScripts.Count) script(s) in $folder" -Type "Info"
                    }
                }
                catch {
                    Update-LogDisplay -Message "Error searching $folder - $($_.Exception.Message)" -Type "Error"
                }
            }
            # Ein else-Zweig für nicht existierende Ordner in $scriptFolders ist hier weniger wahrscheinlich,
            # da die Liste bereits aus validierten und gefilterten Pfaden bestehen sollte.
        }
        
        if ($scripts.Count -eq 0) {
            Update-LogDisplay -Message "No PowerShell scripts found in any searched location (after applying exclusions like '*old*', '*#old*', etc.)." -Type "Warning"
            $textBlock = New-Object System.Windows.Controls.TextBlock
            $textBlock.Text = "No scripts found."
            $textBlock.Margin = "10,5"
            $textBlock.Foreground = "Gray"
            $textBlock.FontStyle = "Italic"
            $spScriptList.Children.Add($textBlock)
            return
        }
        
        $scripts = $scripts | Sort-Object -Property Name
        $scriptsByCategory = @{}
        foreach ($script in $scripts) {
            $category = "General" # Standardkategorie
            
            # Kategorie basierend auf dem übergeordneten Ordnernamen, relativ zum Basisordner
            # $baseFolder wird hier benötigt, um relative Pfade korrekt zu bestimmen.
            # Wenn der ursprüngliche $baseFolder übersprungen wurde, muss dies berücksichtigt werden.
            # Für die Kategorisierung verwenden wir den $PSScriptRoot oder den ermittelten $baseFolder als Referenz.
            # Die Variable $baseFolder hier ist der ursprünglich ermittelte, nicht notwendigerweise der, der zu $scriptFolders hinzugefügt wurde.
            # Dies ist für die relative Pfadbestimmung zur Kategorisierung wichtig.
            
            $categorizationBase = $PSScriptRoot # Fallback, falls $baseFolder komplex ermittelt wurde
            if (-not [string]::IsNullOrEmpty($baseFolder) -and (Test-Path $baseFolder)) {
                 $categorizationBase = $baseFolder
            }

            $scriptParentDirFullPath = Split-Path -Parent $script.FullName
            if ($scriptParentDirFullPath.StartsWith($categorizationBase, [System.StringComparison]::OrdinalIgnoreCase) `
                -and $scriptParentDirFullPath.Length -gt $categorizationBase.Length) {
                
                $relativeDirPath = $scriptParentDirFullPath.Substring($categorizationBase.Length)
                $relativeDirPath = $relativeDirPath.TrimStart('\/')
                
                $category = Split-Path -Leaf $scriptParentDirFullPath
            }
            
            if ($script.Name -match "^(easy\w+)_") {
                $category = $matches[1]
            }
            
            if (-not $scriptsByCategory.ContainsKey($category)) {
                $scriptsByCategory[$category] = @()
            }
            $scriptsByCategory[$category] += $script
        }
        
        foreach ($category in ($scriptsByCategory.Keys | Sort-Object)) {
            $headerBorder = New-Object System.Windows.Controls.Border
            $headerBorder.Background = "#E3F2FD"
            $headerBorder.Margin = "0,10,0,5"
            $headerBorder.Padding = "5"
            $headerBorder.CornerRadius = "3"
            $headerText = New-Object System.Windows.Controls.TextBlock
            $headerText.Text = $category
            $headerText.FontWeight = "Bold"
            $headerBorder.Child = $headerText
            $spScriptList.Children.Add($headerBorder)
            
            foreach ($script in $scriptsByCategory[$category]) {
                $button = New-Object System.Windows.Controls.Button
                $button.Style = $Global:window.FindResource("ScriptButtonStyle")
                $button.Tag = $script.FullName
                
                $buttonContent = New-Object System.Windows.Controls.StackPanel
                $buttonContent.Orientation = "Vertical"
                
                $nameText = New-Object System.Windows.Controls.TextBlock
                $nameText.Text = $script.Name
                $nameText.FontWeight = "SemiBold"
                $buttonContent.Children.Add($nameText)
                
                try {
                    $scriptContent = Get-Content -Path $script.FullName -TotalCount 20 -ErrorAction SilentlyContinue
                    $synopsis = $scriptContent | Where-Object { $_ -match "\.SYNOPSIS" }
                    if ($synopsis) {
                        $descIndex = $scriptContent.IndexOf($synopsis) + 1
                        if ($descIndex -lt $scriptContent.Count) {
                            $scriptDescription = $scriptContent[$descIndex].Trim()
                            if (-not [string]::IsNullOrWhiteSpace($scriptDescription)) {
                                $descText = New-Object System.Windows.Controls.TextBlock
                                $descText.Text = $scriptDescription
                                $descText.TextTrimming = "CharacterEllipsis"
                                $descText.Opacity = 0.7
                                $buttonContent.Children.Add($descText)
                            }
                        }
                    }
                }
                catch { }
                
                $button.Content = $buttonContent
                
                $button.Add_Click({
                    param($buttonSender, $e)
                    $scriptPath = $buttonSender.Tag
                    $scriptName = Split-Path -Leaf $scriptPath
                    $result = [System.Windows.MessageBox]::Show(
                        "Do you want to execute the script '$scriptName' in the current session?",
                        "Execute Script",
                        [System.Windows.MessageBoxButton]::YesNo,
                        [System.Windows.MessageBoxImage]::Question)
                    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                        try {
                            Update-LogDisplay -Message "Executing script: $scriptPath" -Type "Info"
                            Start-MainScript -ScriptPath $scriptPath -KeepSetupWindowOpen
                        }
                        catch {
                            Update-LogDisplay -Message "Error executing script: $($_.Exception.Message)" -Type "Error"
                            [System.Windows.MessageBox]::Show(
                                "Error executing script: $($_.Exception.Message)",
                                "Script Error",
                                [System.Windows.MessageBoxButton]::OK,
                                [System.Windows.MessageBoxImage]::Error)
                        }
                    }
                })
                $spScriptList.Children.Add($button)
            }
        }
        Update-LogDisplay -Message "Script list loaded - $($scripts.Count) scripts found." -Type "Success"
    }
    catch {
        Update-LogDisplay -Message "Error loading script list: $($_.Exception.Message)" -Type "Error"
    }
}
# Vereinfachte Funktion zum Starten des Hauptskripts
function Start-MainScript {
    param (
        [string]$ScriptPath,
        [switch]$KeepSetupWindowOpen
    )
    try {
        if (-not (Test-Path -Path $ScriptPath)) {
            Write-Log "Main script not found: $ScriptPath" -Type "Error"
            return $false
        }
        Write-Log "Loading main script: $ScriptPath" -Type "Info"
        if ($KeepSetupWindowOpen -and $Global:window) {
            Update-LogDisplay -Message "Main window stays open to maintain sessions" -Type "Info"
            $Global:window.WindowState = [System.Windows.WindowState]::Minimized
            $Global:window.Title = "easyIT Setup (Script running...)"
        }
        
        $scriptTimer = [System.Diagnostics.Stopwatch]::StartNew()
        Update-LogDisplay -Message "Executing script: $ScriptPath" -Type "Info"
        & $ScriptPath
        
        return $true
    }
    catch {
        Write-Log "Error executing main script: $($_.Exception.Message)" -Type "Error"
        return $false
    }
    finally {
        if ($scriptTimer) {
            $scriptTimer.Stop()
            Update-LogDisplay -Message "Script execution time: $($scriptTimer.Elapsed.ToString())" -Type "Info"
        }
        
        if ($KeepSetupWindowOpen -and $Global:window) {
            $Global:window.WindowState = [System.Windows.WindowState]::Normal
            $Global:window.Title = "easyIT PowerShell Setup & Module Manager"
        }
    }
}

# Funktion zum Registrieren der Event-Handler der GUI
function Register-EventHandlers {
    try {
        Update-LogDisplay -Message "Registering GUI event handlers..." -Type "Info"
        
        # Button events
        $btnRefreshScripts = $Global:window.FindName("btnRefreshScripts")
        if ($btnRefreshScripts) {
            $btnRefreshScripts.Add_Click({ Update-ScriptList })
            Update-LogDisplay -Message "Script list refresh button event registered." -Type "Info"
        }
        
        $btnRefresh = $Global:window.FindName("btnRefresh")
        if ($btnRefresh) {
            $btnRefresh.Add_Click({ Check-AllModules })
            Update-LogDisplay -Message "Module list refresh button event registered." -Type "Info"
        }
        
        $btnRefreshPolicies = $Global:window.FindName("btnRefreshPolicies")
        if ($btnRefreshPolicies) {
            $btnRefreshPolicies.Add_Click({ Update-ExecutionPolicyDisplay })
            Update-LogDisplay -Message "Execution policy refresh button event registered." -Type "Info"
        }
        
        $btnExit = $Global:window.FindName("btnExit")
        if ($btnExit) {
            $btnExit.Add_Click({ $Global:window.Close() })
            Update-LogDisplay -Message "Exit button event registered." -Type "Info"
        }
        
        # Module ListView events
        $lvModules = $Global:window.FindName("lvModules")
        if ($lvModules) {
            $lvModules.AddHandler(
                [System.Windows.Controls.Button]::ClickEvent, 
                [System.Windows.RoutedEventHandler]{
                    $button = $_.OriginalSource -as [System.Windows.Controls.Button]
                    if ($button) {
                        if ($button.Content -eq "Install/Update") {
                            Install-ModuleHandler -buttonSender $button -e $_
                        }
                        elseif ($button.Content -eq "WEB LINK") {
                            $webUrl = $button.Tag.ToString()
                            Open-WebLink -url $webUrl
                        }
                    }
                }
            )
            Update-LogDisplay -Message "Module ListView button events registered." -Type "Info"
        }
        
        # ExecutionPolicy ListView events
        $lvExecutionPolicies = $Global:window.FindName("lvExecutionPolicies")
        if ($lvExecutionPolicies) {
            $lvExecutionPolicies.AddHandler(
                [System.Windows.Controls.Button]::ClickEvent, 
                [System.Windows.RoutedEventHandler]{
                    $button = $_.OriginalSource -as [System.Windows.Controls.Button]
                    if ($button) {
                        $scope = $button.Tag.ToString()
                        $policy = $button.Content.ToString()
                        
                        # Only allow modification for certain scopes
                        if ($scope -in @("Process", "CurrentUser", "LocalMachine")) {
                            Set-ExecutionPolicyHandler -scope $scope -policy $policy
                        }
                        else {
                            Update-LogDisplay -Message "Cannot modify $scope - controlled by Group Policy." -Type "Warning"
                        }
                    }
                }
            )
            Update-LogDisplay -Message "ExecutionPolicy ListView button events registered." -Type "Info"
        }
        
        Update-LogDisplay -Message "All event handlers successfully registered." -Type "Success"
        return $true
    }
    catch {
        Update-LogDisplay -Message "Error registering event handlers: $($_.Exception.Message)" -Type "Error"
        return $false
    }
}

# Hauptblock zur Initialisierung und Anzeige der GUI
try {
    Write-Log "Starting easyIT PowerShell Setup Tool..." -Type "Info"
    
    # GUI laden
    $Global:window = Import-XamlGui
    if (-not $Global:window) { throw "GUI konnte nicht geladen werden." }
    
    # Module prüfen
    Check-AllModules
    
    # ExecutionPolicy anzeigen
    Update-ExecutionPolicyDisplay
    
    # Skriptliste laden
    Update-ScriptList
    
    # Event-Handler registrieren
    Register-EventHandlers

    Write-Log "GUI erfolgreich initialisiert und Event-Handler registriert." -Type "Success"
    
    # GUI anzeigen
    [void]$Global:window.ShowDialog()
}
catch {
    $exception = $_
    $errorMessage = $exception.Exception.Message
    $fullStackTrace = $exception.ToString() # Für detaillierte Protokollierung

    # Allgemeine Fehlerbehandlung
    $lineNumber = $exception.InvocationInfo.ScriptLineNumber
    Write-Log -Message "Fehler bei der Ausführung des Setup-Tools (Zeile $lineNumber): $errorMessage. FullError: $fullStackTrace" -Type "Error"
    [System.Windows.MessageBox]::Show("Ein unerwarteter Fehler ist aufgetreten: $errorMessage`nDetails finden Sie in den Protokollen.", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
}
