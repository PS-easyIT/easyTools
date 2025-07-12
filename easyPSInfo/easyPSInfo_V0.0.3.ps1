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
    Title="easyIT | PowerShell Module &amp; Policy Manager - with Script Launcher" 
    Height="1000" 
    Width="1450"
    WindowStartupLocation="CenterScreen"
    Background="#FFFFFF"
    FontFamily="Segoe UI"
    FontSize="12">
    <Window.Resources>
        <!-- Modern Button Style -->
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#0078D4" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Padding" Value="12,6" />
            <Setter Property="Margin" Value="4" />
            <Setter Property="FontWeight" Value="SemiBold" />
            <Setter Property="Cursor" Value="Hand" />
            <Setter Property="MinWidth" Value="80" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter Margin="{TemplateBinding Padding}" HorizontalAlignment="Center" VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#106EBE" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#005A9E" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.5" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="SmallModernButton" TargetType="Button" BasedOn="{StaticResource ModernButton}">
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="FontSize" Value="10"/>
            <Setter Property="MinWidth" Value="60"/>
        </Style>

        <Style x:Key="ExitButton" TargetType="Button" BasedOn="{StaticResource ModernButton}">
            <Setter Property="Background" Value="#D32F2F"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter Margin="{TemplateBinding Padding}" HorizontalAlignment="Center" VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#E57373" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#C62828" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.5" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="LinkButton" TargetType="Button" BasedOn="{StaticResource SmallModernButton}">
            <Setter Property="Background" Value="#E8F4FD"/>
            <Setter Property="Foreground" Value="#0078D4"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter Margin="{TemplateBinding Padding}" HorizontalAlignment="Center" VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#D0E0F0" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#B0C0D0" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Card Style for GroupBox -->
        <Style x:Key="CardGroupBox" TargetType="GroupBox">
            <Setter Property="BorderBrush" Value="#E0E0E0" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="Padding" Value="12,8,12,12" />
            <Setter Property="Margin" Value="0,0,0,10" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="GroupBox">
                        <Border Background="White"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="8">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <Border Background="#F7F7F7" CornerRadius="8,8,0,0" Padding="10,6" BorderThickness="0,0,0,1" BorderBrush="#E0E0E0">
                                    <ContentPresenter ContentSource="Header"/>
                                </Border>
                                <ContentPresenter Grid.Row="1" Margin="{TemplateBinding Padding}"/>
                            </Grid>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Script Button Style -->
        <Style x:Key="ScriptButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="#F9F9F9" />
            <Setter Property="Foreground" Value="#333333" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="BorderBrush" Value="#E0E0E0"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="Margin" Value="0,3"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                            <ContentPresenter Margin="{TemplateBinding Padding}" HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#E3F2FD" />
                                <Setter TargetName="border" Property="BorderBrush" Value="#0078D4"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#BBDEFB" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.6" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="ScriptListHeaderBorderStyle" TargetType="Border">
            <Setter Property="Background" Value="#E3F2FD"/>
            <Setter Property="Margin" Value="0,10,0,5"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="CornerRadius" Value="3"/>
        </Style>

    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <!-- Header -->
        <Border Grid.Row="0" Background="#FF1C323C" Padding="20,15">
            <StackPanel>
                <TextBlock Text="easyIT | PowerShell Module &amp; Policy Manager - with Script Launcher" FontSize="20" FontWeight="SemiBold" Foreground="#e2e2e2"/>
                <TextBlock Text="Manage PowerShell modules and execution policies | Launch scripts from subfolders" FontSize="12" Foreground="#d0e8ff" Margin="0,5,0,0"/>
            </StackPanel>
        </Border>
        <!-- Content with Split Layout -->
        <Grid Grid.Row="1" Margin="15">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="2.2*"/>
                <!-- Status area: slightly larger -->
                <ColumnDefinition Width="1*"/>
                <!-- Script selection: 1/3 -->
            </Grid.ColumnDefinitions>
            <!-- Left side - Status area -->
            <Grid Grid.Column="0" Margin="0,0,10,0">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <!-- Module Check Section -->
                <GroupBox Grid.Row="0" Header="PowerShell Modules" Style="{StaticResource CardGroupBox}">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <ListView Grid.Row="0" x:Name="lvModules" Margin="0,5,0,10" BorderThickness="1" BorderBrush="#E0E0E0" Height="260">
                            <ListView.View>
                                <GridView>
                                    <GridViewColumn Header="Module" Width="410">
                                        <GridViewColumn.CellTemplate>
                                            <DataTemplate>
                                                <TextBlock Text="{Binding Name}" ToolTip="{Binding Description}" VerticalAlignment="Center"/>
                                            </DataTemplate>
                                        </GridViewColumn.CellTemplate>
                                    </GridViewColumn>
                                    <GridViewColumn Header="Required" DisplayMemberBinding="{Binding MinVersion}" Width="100"/>
                                    <GridViewColumn Header="Installed" DisplayMemberBinding="{Binding InstalledVersion}" Width="100"/>
                                    <GridViewColumn Header="Action" Width="120">
                                        <GridViewColumn.CellTemplate>
                                            <DataTemplate>
                                                <Button Content="Install/Update" 
                                                        Style="{StaticResource SmallModernButton}"
                                                        IsEnabled="{Binding InstallEnabled}"
                                                        Tag="{Binding Name}"/>
                                            </DataTemplate>
                                        </GridViewColumn.CellTemplate>
                                    </GridViewColumn>
                                    <GridViewColumn Header="Link" Width="90">
                                        <GridViewColumn.CellTemplate>
                                            <DataTemplate>
                                                <Button Content="Web Link" 
                                                        Style="{StaticResource LinkButton}"
                                                        Tag="{Binding Webseite}"/>
                                            </DataTemplate>
                                        </GridViewColumn.CellTemplate>
                                    </GridViewColumn>
                                </GridView>
                            </ListView.View>
                        </ListView>
                        <Button Grid.Row="1" x:Name="btnRefreshModules" Content="Refresh Modules" Style="{StaticResource ModernButton}" HorizontalAlignment="Left" Width="180" Height="25" Margin="0,5,0,0"/>
                    </Grid>
                </GroupBox>
                <!-- ExecutionPolicy Section -->
                <GroupBox Grid.Row="1" Header="PowerShell Execution Policies" Style="{StaticResource CardGroupBox}">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <ListView Grid.Row="0" x:Name="lvExecutionPolicies" Margin="0,0,0,10" BorderThickness="1" BorderBrush="#E0E0E0" Height="165">
                            <ListView.View>
                                <GridView>
                                    <GridViewColumn Header="Scope" DisplayMemberBinding="{Binding Description}" Width="330"/>
                                    <GridViewColumn Header="Current Policy" DisplayMemberBinding="{Binding CurrentPolicy}" Width="150"/>
                                    <GridViewColumn Header="Modify" Width="400">
                                        <GridViewColumn.CellTemplate>
                                            <DataTemplate>
                                                <StackPanel Orientation="Horizontal">
                                                    <Button Content="Restricted" Style="{StaticResource SmallModernButton}" Margin="2,0" Tag="{Binding Scope}" Width="65"/>
                                                    <Button Content="AllSigned" Style="{StaticResource SmallModernButton}" Margin="2,0" Tag="{Binding Scope}" Width="65"/>
                                                    <Button Content="RemoteSigned" Style="{StaticResource SmallModernButton}" Margin="2,0" Tag="{Binding Scope}" Width="90"/>
                                                    <Button Content="Unrestricted" Style="{StaticResource SmallModernButton}" Margin="2,0" Tag="{Binding Scope}" Width="80"/>
                                                    <Button Content="Bypass" Style="{StaticResource SmallModernButton}" Margin="2,0" Tag="{Binding Scope}" Width="55"/>
                                                </StackPanel>
                                            </DataTemplate>
                                        </GridViewColumn.CellTemplate>
                                    </GridViewColumn>
                                </GridView>
                            </ListView.View>
                        </ListView>
                        <Button Grid.Row="1" x:Name="btnRefreshPolicies" Content="Refresh Execution Policies" Style="{StaticResource ModernButton}" HorizontalAlignment="Left" Width="200" Height="25"/>
                    </Grid>
                </GroupBox>
                <!-- Status Log -->
                <GroupBox Grid.Row="2" Header="Status and Log" Style="{StaticResource CardGroupBox}" Margin="0,10,0,0">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <TextBox x:Name="txtLog" IsReadOnly="True" TextWrapping="Wrap" 
                                 FontFamily="Consolas" FontSize="11" Background="#1E1E1E" Foreground="#A0D2A0" BorderThickness="0" Padding="5"/>
                    </ScrollViewer>
                </GroupBox>
            </Grid>
            <!-- Right side - Script selection -->
            <GroupBox Grid.Column="1" Header="Available PowerShell Scripts" Style="{StaticResource CardGroupBox}" Margin="10,0,0,0">
                <Grid>
                    <Grid.RowDefinitions>
                        <!-- Suchfeld (optional, hier nicht implementiert) -->
                        <!-- <RowDefinition Height="Auto"/> -->
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <!-- Script list -->
                    <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto" Margin="0,0,0,10">
                        <StackPanel x:Name="spScriptList" Orientation="Vertical" Margin="0,3,0,3">
                            <!-- Script buttons will be added dynamically -->
                        </StackPanel>
                    </ScrollViewer>
                    <!-- Refresh button -->
                    <Button Grid.Row="1" x:Name="btnRefreshScripts" 
                            Content="Refresh Script List" Style="{StaticResource ModernButton}"
                            Height="32"/>
                </Grid>
            </GroupBox>
        </Grid>
        <!-- Footer with Buttons and Info -->
        <Border Grid.Row="2" Background="#FF1C323C" Padding="20,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" Orientation="Vertical" HorizontalAlignment="Left" VerticalAlignment="Center">
                    <TextBlock Text="PSscripts.de | Andreas Hepp  | easyPSinfo  -  Version: 0.0.3  -  Update: 23.05.2025" FontSize="10" Foreground="#d0e8ff"/>
                </StackPanel>

                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                    <Button x:Name="btnExit" Content="Exit" Style="{StaticResource ExitButton}" Width="120" Height="32"/>
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
        # $initialBaseFolder wird als Referenz für die hierarchische Anzeige verwendet
        $initialBaseFolder = $PSScriptRoot
        if ([string]::IsNullOrEmpty($initialBaseFolder)) {
            Update-LogDisplay -Message "PSScriptRoot nicht verfügbar, versuche alternativen Pfad..." -Type "Warning"
            if ($MyInvocation.MyCommand.Path) {
                $initialBaseFolder = Split-Path -Parent $MyInvocation.MyCommand.Path
            }
            else {
                # Potenzielle Pfade, die dynamisch oder benutzerbezogen sind
                $potentialPaths = @(
                    Join-Path -Path $env:USERPROFILE -ChildPath "Documents\PowerShell\Scripts", # Standard-Benutzer-Skriptverzeichnis
                    (Get-Location).Path # Aktuelles Arbeitsverzeichnis
                )
                # Versuche, einen Standard-Modulpfad zu finden, falls vorhanden
                $defaultUserModulePath = Join-Path -Path $env:USERPROFILE -ChildPath "Documents\WindowsPowerShell\Modules"
                if (Test-Path -Path $defaultUserModulePath) {
                    $potentialPaths += $defaultUserModulePath
                }

                foreach ($path in $potentialPaths) {
                    if (Test-Path -Path $path -ErrorAction SilentlyContinue) {
                        $initialBaseFolder = $path
                        Update-LogDisplay -Message "Verwende gefundenen Pfad als Basis für die Skriptsuche: $initialBaseFolder" -Type "Info"
                        break
                    }
                }
            }
        }
        
        if ([string]::IsNullOrEmpty($initialBaseFolder)) {
            Update-LogDisplay -Message "No valid script directory base found." -Type "Error"
            return
        }
        
        # Ausschlussmuster für Ordnernamen (case-insensitive durch -like)
        $excludeFolderPatterns = @("*old*", "*#old*", "*# old*")

        if (Test-Path -Path $initialBaseFolder -ErrorAction SilentlyContinue) {
            $baseFolderLeafName = (Split-Path -Leaf $initialBaseFolder)
            $skipBaseFolderSearch = $false
            foreach ($pattern in $excludeFolderPatterns) {
                if ($baseFolderLeafName -like $pattern) {
                    $skipBaseFolderSearch = $true
                    Update-LogDisplay -Message "Base folder '$initialBaseFolder' is excluded from direct script search because its name matches pattern '$pattern'." -Type "Warning"
                    break
                }
            }

            if (-not $skipBaseFolderSearch) {
                Update-LogDisplay -Message "Adding base path for script search: $initialBaseFolder" -Type "Info"
                $scriptFolders += $initialBaseFolder 
            }
            # else: $initialBaseFolder is skipped for direct script search, but still acts as categorization root.

            # Alle Unterordner rekursiv hinzufügen, die nicht den Ausschlussmustern entsprechen
            try {
                # Hole ALLE Unterverzeichnisse zuerst
                $allSubDirectories = Get-ChildItem -Path $initialBaseFolder -Directory -Recurse -ErrorAction SilentlyContinue
                
                $validSubDirectories = @()
                $trimmedInitialBaseFolder = $initialBaseFolder.TrimEnd([System.IO.Path]::DirectorySeparatorChar)

                foreach ($dirInfoObject in $allSubDirectories) {
                    $currentDirFullName = $dirInfoObject.FullName
                    $isAnySegmentExcluded = $false
                    
                    $relativePathString = $null
                    if ($currentDirFullName.Length -gt $trimmedInitialBaseFolder.Length) {
                        if ($currentDirFullName.StartsWith($trimmedInitialBaseFolder + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
                            $relativePathString = $currentDirFullName.Substring($trimmedInitialBaseFolder.Length).TrimStart([System.IO.Path]::DirectorySeparatorChar)
                        } else {
                            # Nicht wirklich ein Unterpfad im Sinne der Hierarchie vom initialBaseFolder, überspringen
                            continue
                        }
                    } else {
                         # $currentDirFullName ist $initialBaseFolder selbst oder kürzer
                        continue
                    }
                    
                    if (-not ([string]::IsNullOrEmpty($relativePathString))) {
                        $pathSegments = $relativePathString.Split([System.IO.Path]::DirectorySeparatorChar)
                        foreach ($segment in $pathSegments) {
                            if ([string]::IsNullOrEmpty($segment)) { continue } 
                            foreach ($pattern in $excludeFolderPatterns) {
                                if ($segment -like $pattern) {
                                    $isAnySegmentExcluded = $true
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
                    Update-LogDisplay -Message "Added $($validSubDirectories.Count) subfolders to script search (after filtering path segments)." -Type "Info"
                } else {
                     Update-LogDisplay -Message "No subfolders added to script search after path segment filtering." -Type "Info"
                }
            }
            catch {
                Update-LogDisplay -Message "Error processing subfolders from ${initialBaseFolder}: $($_.Exception.Message)" -Type "Warning"
            }
        }
        else {
            Update-LogDisplay -Message "Initial base folder for scripts $initialBaseFolder does not exist or is not accessible." -Type "Error"
            return 
        }
        
        if ($scriptFolders.Count -gt 0) {
            Update-LogDisplay -Message "Searching for scripts in $($scriptFolders.Count) locations." -Type "Info"
        } else {
            Update-LogDisplay -Message "No script folders to search (base folder might be excluded or contain no valid subfolders after exclusions)." -Type "Warning"
        }
        
        $scripts = @()
        foreach ($folder in $scriptFolders) { 
            if (-not [string]::IsNullOrEmpty($folder) -and (Test-Path -Path $folder -ErrorAction SilentlyContinue)) {
                try {
                    $currentScriptName = $null
                    if ($MyInvocation.MyCommand.Path) {
                        $currentScriptName = Split-Path -Leaf $MyInvocation.MyCommand.Path
                    }
                    # Get scripts only from the current $folder, not its subdirectories (Depth 0 implicitly)
                    $folderScripts = Get-ChildItem -Path $folder -Filter "*.ps1" -File -ErrorAction Stop | 
                                     Where-Object { $currentScriptName -eq $null -or $_.Name -ne $currentScriptName }
                    if ($folderScripts -and $folderScripts.Count -gt 0) {
                        $scripts += $folderScripts
                    }
                }
                catch {
                    Update-LogDisplay -Message "Error searching for scripts in $folder - $($_.Exception.Message)" -Type "Error"
                }
            }
        }
        
        if ($scripts.Count -eq 0) {
            Update-LogDisplay -Message "No PowerShell scripts found in any searched location (after applying exclusions)." -Type "Warning"
            $textBlock = New-Object System.Windows.Controls.TextBlock
            $textBlock.Text = "No scripts found."
            $textBlock.Margin = "9"
            $textBlock.Foreground = "Gray"
            $textBlock.FontStyle = "Italic"
            $spScriptList.Children.Add($textBlock)
            return
        }
        
        $categorizationBase = $initialBaseFolder 

        $scriptsByExactDirectory = $scripts | Group-Object -Property DirectoryName

        # Erstelle eine korrekte hierarchische Sortierung
        $sortedDirectoryGroups = $scriptsByExactDirectory | Sort-Object {
            $dirPath = $_.Name # Full path of the directory being grouped
            if ($dirPath -eq $categorizationBase) {
                "000000_ROOT" # Root folder always first
            } else {
                # Berechne den relativen Pfad für hierarchische Sortierung
                $relativePath = ""
                if ($dirPath.StartsWith($categorizationBase, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $relativePath = $dirPath.Substring($categorizationBase.Length).TrimStart([System.IO.Path]::DirectorySeparatorChar)
                } else {
                    $relativePath = $dirPath
                }
                
                # Ersetze Backslashes durch Forward-Slashes für konsistente Sortierung
                $normalizedPath = $relativePath.Replace('\', '/')
                
                # Teile den Pfad in Segmente und sortiere hierarchisch
                $segments = $normalizedPath -split '/' | Where-Object { $_ -ne '' }
                $depth = $segments.Count
                
                # Erstelle Sortier-Key: Tiefe + normalisierter Pfad für korrekte Hierarchie
                "{0:D3}_{1}" -f $depth, $normalizedPath.ToLower()
            }
        }
        
        foreach ($directoryGroup in $sortedDirectoryGroups) {
            $currentDirectoryPath = $directoryGroup.Name 
            $scriptsInDirectory = $directoryGroup.Group | Sort-Object -Property Name

            $headerTextString = ""
            $headerMarginValueString = "0,10,0,5"
            $buttonLeftMarginValue = 10
            $currentDepth = 0

            if ($currentDirectoryPath -eq $categorizationBase) {
                $headerTextString = "Folder Root" 
            }
            else {
                # Berechne die korrekte Tiefe und den Display-Namen
                $relativePath = ""
                if ($currentDirectoryPath.StartsWith($categorizationBase, [System.StringComparison]::OrdinalIgnoreCase) `
                    -and $currentDirectoryPath.Length -gt $categorizationBase.Length) {
                    
                    $relativePath = $currentDirectoryPath.Substring($categorizationBase.Length).TrimStart([System.IO.Path]::DirectorySeparatorChar)
                    $pathSegments = $relativePath -split '[\\/]' | Where-Object { $_ -ne '' }
                    $currentDepth = $pathSegments.Count
                    
                    # Zeige den vollen relativen Pfad für bessere Klarheit
                    $displayPath = $pathSegments -join ' -> '
                    $indentPrefix = "  " * $currentDepth
                    $headerTextString = "$indentPrefix$displayPath"
                } else {
                    # Fallback für Pfade außerhalb der Basis
                    $headerTextString = (Split-Path -Leaf $currentDirectoryPath)
                }
                
                # Anpassung der Button-Einrückung basierend auf der Tiefe
                $buttonLeftMarginValue = 10 + ($currentDepth * 15)
            }

            $headerBorder = New-Object System.Windows.Controls.Border
            $headerBorder.Background = "#E3F2FD"
            $headerBorder.Margin = $headerMarginValueString
            $headerBorder.Padding = "8,4"
            $headerBorder.CornerRadius = "3"
            
            $headerTextElement = New-Object System.Windows.Controls.TextBlock
            $headerTextElement.Text = $headerTextString
            $headerTextElement.FontWeight = "Bold"
            $headerTextElement.FontSize = "10"
            $headerBorder.Child = $headerTextElement
            $spScriptList.Children.Add($headerBorder)
            
            foreach ($script in $scriptsInDirectory) {
                $button = New-Object System.Windows.Controls.Button
                $button.Style = $Global:window.FindResource("ScriptButtonStyle")
                $button.Tag = $script.FullName
                $button.Margin = "$($buttonLeftMarginValue),2,5,2"
                
                $buttonContent = New-Object System.Windows.Controls.StackPanel
                $buttonContent.Orientation = "Vertical"
                
                $nameText = New-Object System.Windows.Controls.TextBlock
                $nameText.Text = $script.Name
                $nameText.FontWeight = "SemiBold"
                $buttonContent.Children.Add($nameText)
                
                try {
                    $scriptFileContent = Get-Content -Path $script.FullName -TotalCount 20 -ErrorAction SilentlyContinue
                    $synopsisDirectiveLine = $scriptFileContent | Where-Object { $_ -match "\.SYNOPSIS" } | Select-Object -First 1
                    
                    if ($synopsisDirectiveLine) {
                        $synopsisDirectiveIndex = -1
                        for ($i = 0; $i -lt $scriptFileContent.Count; $i++) {
                            if ($scriptFileContent[$i] -eq $synopsisDirectiveLine) {
                                $synopsisDirectiveIndex = $i
                                break
                            }
                        }

                        if ($synopsisDirectiveIndex -ne -1) {
                            for ($i = $synopsisDirectiveIndex + 1; $i -lt $scriptFileContent.Count; $i++) {
                                $potentialDescription = $scriptFileContent[$i].Trim()
                                if (-not [string]::IsNullOrWhiteSpace($potentialDescription) -and -not $potentialDescription.StartsWith("#")) {
                                    $descText = New-Object System.Windows.Controls.TextBlock
                                    $descText.Text = $potentialDescription
                                    $descText.TextTrimming = "CharacterEllipsis"
                                    $descText.Opacity = 0.7
                                    $descText.FontSize = "10"
                                    $buttonContent.Children.Add($descText)
                                    break
                                }
                            }
                        }
                    }
                }
                catch { }
                
                $button.Content = $buttonContent
                
                $button.Add_Click({
                    param($buttonSender, $e)
                    $scriptPathToExecute = $buttonSender.Tag
                    $scriptFileName = Split-Path -Leaf $scriptPathToExecute
                    $result = [System.Windows.MessageBox]::Show(
                        "Do you want to execute the script '$scriptFileName' in the current session?",
                        "Execute Script",
                        [System.Windows.MessageBoxButton]::YesNo,
                        [System.Windows.MessageBoxImage]::Question)
                    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                        try {
                            Update-LogDisplay -Message "Executing script: $scriptPathToExecute" -Type "Info"
                            Start-MainScript -ScriptPath $scriptPathToExecute -KeepSetupWindowOpen
                        }
                        catch {
                            Update-LogDisplay -Message "Error executing script '$scriptFileName': $($_.Exception.Message)" -Type "Error"
                            [System.Windows.MessageBox]::Show(
                                "Error executing script '$scriptFileName': $($_.Exception.Message)",
                                "Script Execution Error",
                                [System.Windows.MessageBoxButton]::OK,
                                [System.Windows.MessageBoxImage]::Error)
                        }
                    }
                })
                $spScriptList.Children.Add($button)
            }
        }
        Update-LogDisplay -Message "Script list loaded - $($scripts.Count) scripts found and displayed hierarchically." -Type "Success"
    }
    catch {
        Update-LogDisplay -Message "Critical error while updating script list: $($_.Exception.Message)" -Type "Error"
        # Consider if re-throwing is appropriate: throw $_
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

# SIG # Begin signature block
# MIIcCAYJKoZIhvcNAQcCoIIb+TCCG/UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA2bGZ8oI/XYLuu
# RWnT/xJYqmVYF8ItqNqQ1sP4pOZH/6CCFk4wggMQMIIB+KADAgECAhB3jzsyX9Cg
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
# LSd+2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3H
# rFcF/yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAw
# MFoXDTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFt
# cGluZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQAD
# ggIPADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU
# 7UNqEY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR
# +2fkHUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwE
# u7EEbkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Za
# zch8NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW3
# 5xUUFREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gd
# FpBP9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rq
# BvKWxdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vH
# espYMQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QE
# PHrPV6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1
# Wd4+zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMB
# AAGjggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQG
# fHrK4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAO
# BgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEE
# azBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYB
# BQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0
# ZWRSb290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYG
# Z4EMAQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9
# EXZxML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk
# 97frPBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2
# UwM+NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71
# WPYAgwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQf
# jXQA1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noD
# js6+BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxi
# Df06VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/
# D284NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8Ml
# uDezooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG
# 2XlM9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8
# hcpSM9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLR
# xHanlXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5E
# aWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1l
# U3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAw
# WhcNMzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNl
# cnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVz
# dGFtcCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr
# 0cVMF3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBb
# ZHqRK71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQK
# WXymOtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wD
# cKgH+JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25
# CFyD23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6l
# vJukx7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dV
# mVzix4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuh
# KuAeNIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7C
# e7vMRHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTR
# ofOsNyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUw
# ggGRMAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzo
# MB8GA1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIH
# gDAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZR
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGlt
# ZVN0YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBS
# oFCGTmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRp
# bWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgG
# BmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5
# rmBB7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZE
# N/FPsLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwB
# D9R0oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QA
# GB9lctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBV
# N4ueLaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW6
# 0OiMEgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQ
# TwtSSpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC
# 3yXZi/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmA
# p/jsJ3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9T
# HFVkT+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84
# ZPvmpovq90K8eWyG2N01c4IhSOxqt81nMYIFEDCCBQwCAQEwNDAgMR4wHAYDVQQD
# DBVQaGluSVQtUFNzY3JpcHRzX1NpZ24CEHePOzJf0KCMSL6wELasExMwDQYJYIZI
# AWUDBAIBBQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0B
# CQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAv
# BgkqhkiG9w0BCQQxIgQgy6lQhLv4c0HaYyIdQg5wumICxX0sFNOVATz3yOH7D+ww
# DQYJKoZIhvcNAQEBBQAEggEAbcCq19XhVhPnHYzIaLNV6Q0GKIvJ0JZgoYDicEk4
# /dpQAc0HTUA8T0T1EBvTlRnh0M0yJDUrBH3WNmU3RNkv+edyg3wUQZCmol0ZQp8u
# A+ffuH6mhH29DZiL479Q8FFfN9ZNHOqWD/l4/CqJpm07m1T4iw/U4EpU6KyNKGpA
# mg+trpLAoSyszUMhVKPzc76FT6DzlsRrkmTUrylldoph3ukBQz1ZTo98vB4aYL8C
# UE/AuTdMwpFx8IfGYaVHfRU7g7saIG06zQw/i0CdSjj+hhNiVt2BNC1f5hKESG1j
# 3DnEkfA1kdTnZ6b3Zwa5Hg95WNOo9Bsdtci36xfaiJy5vqGCAyYwggMiBgkqhkiG
# 9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgw
# DQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0yNTA3MTIwODAwMjZaMC8GCSqGSIb3DQEJBDEiBCA/oHCv6qKr
# zfofI0sjICZvvcyPPQWb7KOg7u1LRiOqjjANBgkqhkiG9w0BAQEFAASCAgCQeYKB
# Lhe4kkO1rH4X9ss4xiwzIrK7ASz0f9Twp3zzRv5gh80fdkvuhXVHGajMKTDiiD3X
# T9K94L1qLDH/1lPIj30HY+G78WsbiYVt5znkArM0DPluF6n37221QzjrEHvshWvr
# fQx/ioJ/MaF0hz3HzWKjmmsBKa5m8g+LI9SfKpLpacaWuQdnzaLFXctWzkKKvLTS
# gVsH8R5J/5Lc1NgfTuDWz2SmO8SMHfvaQrz//F+SQHNhftRK3K0l6wzeUjt0u7it
# 49Y42nGACxoOnc1brnthwpNnYXll1K/drWQmeC+4xU0ufzb1Q+DyO4iv8Qc0qLty
# WqJz4XHxZJ1VSVveeQ8H6cFP6IZZeIo9mWBAPBPR57xnwKYr0K56/wf3qshz9+aU
# D28Djv/IFtVCelS/wHd1AT5JeJdbRkPGiDCrhTPtoVWxGK8rRXcjkZJb+gxamAx6
# xA/wWMkzKv6EnF7rkpF11VwOhQtdLUllsDNrIg6Cl8csoC1E04j8cKGk0w9Gq6yH
# 4JDsSHEb6SGOR1vyAVFW8Pp5vGG7Q+eVxLzilIM2t80+aFD/2hRy+bamGMOan6mN
# 1j0qfLz0rr5ofZy4Sfwzu4foBAGbK8wS5T+EOcR18u1buWkG/ioJJXBP7Ejs1thH
# bFrnae7BEjoU13PwhmbUNrDGnp3e4fgdjKqM+A==
# SIG # End signature block
