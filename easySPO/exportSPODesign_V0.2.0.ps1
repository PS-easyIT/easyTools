# WPF statt Windows.Forms verwenden
Write-Host "[DEBUG] Skript gestartet - Lade benötigte Assemblies..." -ForegroundColor Cyan
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Write-Host "[DEBUG] Assemblies geladen" -ForegroundColor Cyan

# Get current execution policy for use in the form
Write-Host "[DEBUG] Prüfe aktuelle Execution Policy..." -ForegroundColor Cyan
$epCurrent = Get-ExecutionPolicy -Scope Process -ErrorAction SilentlyContinue
Write-Host "[DEBUG] Execution Policy: $epCurrent" -ForegroundColor Cyan

# Check if modules are installed for use in the form
Write-Host "[DEBUG] Prüfe, ob erforderliche Module installiert sind..." -ForegroundColor Cyan
$pnpModuleName = "PnP.PowerShell"
$spoModuleName = "Microsoft.Online.SharePoint.PowerShell"

$pnpModuleInstalled = Get-Module -ListAvailable -Name $pnpModuleName
$spoModuleInstalled = Get-Module -ListAvailable -Name $spoModuleName

if ($pnpModuleInstalled) {
    Write-Host "[DEBUG] Modul $pnpModuleName ist installiert: Version $($pnpModuleInstalled.Version)" -ForegroundColor Green
} else {
    Write-Host "[DEBUG] Modul $pnpModuleName ist NICHT installiert" -ForegroundColor Yellow
}

if ($spoModuleInstalled) {
    Write-Host "[DEBUG] Modul $spoModuleName ist installiert: Version $($spoModuleInstalled.Version)" -ForegroundColor Green
} else {
    Write-Host "[DEBUG] Modul $spoModuleName ist NICHT installiert" -ForegroundColor Yellow
}

# Variable für exportiertes Site Script
$global:exportedSiteScript = $null
$global:createdSiteScriptId = $null
$global:createdSiteDesignId = $null

# Die XAML-Datei für die WPF-GUI laden
function Load-XamlGui {
    Write-Host "[DEBUG] Lade XAML GUI..." -ForegroundColor Cyan
    $xamlContent = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="SharePoint Design Export"
    Width="1300"
    Height="915"
    Background="#F3F3F3"
    WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <!--  Styles für Windows 11 / Metro Look  -->
        <Style x:Key="HeaderTextBlockStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#202020" />
            <Setter Property="FontSize" Value="24" />
            <Setter Property="FontWeight" Value="SemiBold" />
            <Setter Property="Margin" Value="10,0,0,0" />
            <Setter Property="VerticalAlignment" Value="Center" />
        </Style>

        <Style x:Key="StatusOkTextBlockStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#4CAF50" />
            <Setter Property="FontSize" Value="13" />
            <Setter Property="FontWeight" Value="SemiBold" />
        </Style>

        <Style x:Key="StatusErrorTextBlockStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#F44336" />
            <Setter Property="FontSize" Value="13" />
            <Setter Property="FontWeight" Value="SemiBold" />
        </Style>

        <Style x:Key="StatusInfoTextBlockStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#2196F3" />
            <Setter Property="FontSize" Value="13" />
            <Setter Property="FontWeight" Value="Normal" />
        </Style>

        <Style x:Key="LabelStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#505050" />
            <Setter Property="FontSize" Value="13" />
            <Setter Property="Margin" Value="0,5,0,2" />
        </Style>

        <Style x:Key="GroupHeaderStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#303030" />
            <Setter Property="FontSize" Value="14" />
            <Setter Property="FontWeight" Value="SemiBold" />
            <Setter Property="Margin" Value="0,0,0,10" />
        </Style>

        <Style x:Key="FooterTextBlockStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#505050" />
            <Setter Property="FontSize" Value="12" />
            <Setter Property="HorizontalAlignment" Value="Center" />
            <Setter Property="VerticalAlignment" Value="Center" />
        </Style>

        <Style x:Key="ModernGroupBoxStyle" TargetType="GroupBox">
            <Setter Property="BorderBrush" Value="#DDDDDD" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="Padding" Value="10" />
            <Setter Property="Margin" Value="0,5,0,10" />
            <Setter Property="Background" Value="#FFFFFF" />
        </Style>

        <Style x:Key="ModernTextBoxStyle" TargetType="TextBox">
            <Setter Property="Height" Value="30" />
            <Setter Property="Padding" Value="5,0" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
            <Setter Property="Background" Value="White" />
            <Setter Property="BorderBrush" Value="#CCCCCC" />
            <Setter Property="BorderThickness" Value="1" />
        </Style>

        <Style
            x:Key="RequiredTextBoxStyle"
            BasedOn="{StaticResource ModernTextBoxStyle}"
            TargetType="TextBox">
            <Setter Property="BorderBrush" Value="#F44336" />
            <Style.Triggers>
                <Trigger Property="Text" Value="">
                    <Setter Property="ToolTip" Value="Dieses Feld ist erforderlich" />
                </Trigger>
                <Trigger Property="Text" Value="{x:Null}">
                    <Setter Property="ToolTip" Value="Dieses Feld ist erforderlich" />
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="ModernButtonStyle" TargetType="Button">
            <Setter Property="Height" Value="30" />
            <Setter Property="Padding" Value="15,0" />
            <Setter Property="Background" Value="#0078D7" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Cursor" Value="Hand" />
            <Style.Triggers>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#CCCCCC" />
                    <Setter Property="Foreground" Value="#888888" />
                </Trigger>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#106EBE" />
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="SecondaryButtonStyle" TargetType="Button">
            <Setter Property="Height" Value="30" />
            <Setter Property="Padding" Value="15,0" />
            <Setter Property="Background" Value="#EFEFEF" />
            <Setter Property="Foreground" Value="#505050" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Cursor" Value="Hand" />
            <Style.Triggers>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#CCCCCC" />
                    <Setter Property="Foreground" Value="#888888" />
                </Trigger>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#E0E0E0" />
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="ModernCheckBoxStyle" TargetType="CheckBox">
            <Setter Property="Margin" Value="0,5" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
        </Style>

        <Style x:Key="ModernRadioButtonStyle" TargetType="RadioButton">
            <Setter Property="Margin" Value="0,5" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
        </Style>

        <Style x:Key="ModernProgressBarStyle" TargetType="ProgressBar">
            <Setter Property="Height" Value="10" />
            <Setter Property="Foreground" Value="#4CAF50" />
            <Setter Property="Background" Value="#E0E0E0" />
        </Style>

        <!--  Neue Workflow Styles  -->
        <Style x:Key="WorkflowIconContainer" TargetType="Border">
            <Setter Property="Width" Value="40" />
            <Setter Property="Height" Value="40" />
            <Setter Property="Background" Value="#F0F0F0" />
            <Setter Property="BorderBrush" Value="#DDDDDD" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="CornerRadius" Value="20" />
            <Setter Property="Margin" Value="10,5,10,5" />
            <Setter Property="Padding" Value="8" />
            <Setter Property="HorizontalAlignment" Value="Center" />
            <Setter Property="VerticalAlignment" Value="Center" />
        </Style>

        <Style x:Key="WorkflowLabelStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#505050" />
            <Setter Property="FontSize" Value="12" />
            <Setter Property="FontWeight" Value="Medium" />
            <Setter Property="TextAlignment" Value="Center" />
            <Setter Property="Margin" Value="0,3,0,2" />
        </Style>

        <!--  Workflow-Schritt-Status Styles  -->
        <Style
            x:Key="WorkflowIconCompleted"
            BasedOn="{StaticResource WorkflowIconContainer}"
            TargetType="Border">
            <Setter Property="Background" Value="#E3F2FD" />
            <Setter Property="BorderBrush" Value="#2196F3" />
        </Style>

        <Style
            x:Key="WorkflowIconActive"
            BasedOn="{StaticResource WorkflowIconContainer}"
            TargetType="Border">
            <Setter Property="Background" Value="#E8F5E9" />
            <Setter Property="BorderBrush" Value="#4CAF50" />
        </Style>
    </Window.Resources>

    <Grid Margin="0,0,0,0">
        <Grid.RowDefinitions>
            <RowDefinition Height="75" />
            <!--  Header  -->
            <RowDefinition Height="*" />
            <!--  Content  -->
            <RowDefinition Height="60" />
            <!--  Footer  -->
        </Grid.RowDefinitions>

        <!--  Header  -->
        <Border Grid.Row="0" Background="#0078D7">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto" />
                    <ColumnDefinition Width="*" />
                </Grid.ColumnDefinitions>

                <!--  Logo / App Icon  -->
                <Border
                    Grid.Column="0"
                    Width="50"
                    Height="50"
                    Margin="15,0,10,0"
                    Background="White"
                    CornerRadius="5">
                    <Path
                        Margin="10"
                        Data="M12,2C6.47,2 2,6.47 2,12C2,17.53 6.47,22 12,22C17.53,22 22,17.53 22,12C22,6.47 17.53,2 12,2M15.1,7.07C15.24,7.07 15.38,7.12 15.5,7.23L16.77,8.5C17,8.72 17,9.07 16.77,9.28L15.77,10.28L13.72,8.23L14.72,7.23C14.82,7.12 14.96,7.07 15.1,7.07M13.13,8.81L15.19,10.87L9.13,16.93H7.07V14.87L13.13,8.81Z"
                        Fill="#0078D7"
                        Stretch="Uniform" />
                </Border>

                <!--  App Title  -->
                <TextBlock
                    Grid.Column="1"
                    Foreground="White"
                    Style="{StaticResource HeaderTextBlockStyle}"
                    Text="SharePoint Design Export Tool" />
            </Grid>
        </Border>

        <!--  Content Area - 3 Columns  -->
        <Grid Grid.Row="1" Margin="20">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto" />
                <!--  Workflow Overview  -->
                <RowDefinition Height="*" />
                <!--  Main Content  -->
            </Grid.RowDefinitions>

            <!--  Workflow Overview  -->
            <GroupBox
                Grid.Row="0"
                Margin="0,0,0,30"
                Header="Ablauf">
                <StackPanel HorizontalAlignment="Center" Orientation="Horizontal">
                    <!--  Schritt 1: Module installieren  -->
                    <StackPanel>
                        <Border x:Name="workflowIcon1" Style="{StaticResource WorkflowIconContainer}">
                            <Viewbox>
                                <Canvas Width="16" Height="16">
                                    <!--  Down Arrow für Download  -->
                                    <Path Data="M2,0 L14,0 L8,10 Z" Fill="Gray" />
                                </Canvas>
                            </Viewbox>
                        </Border>
                        <TextBlock Style="{StaticResource WorkflowLabelStyle}" Text="Module" />
                        <!--  Erklärung für Schritt 1  -->
                        <TextBlock
                            FontSize="9"
                            Foreground="Gray"
                            Text="1. SPO + PnP Module installieren"
                            TextAlignment="Center" />
                    </StackPanel>
                    <!--  Schritt 2: Login  -->
                    <StackPanel>
                        <Border x:Name="workflowIcon2" Style="{StaticResource WorkflowIconContainer}">
                            <Viewbox>
                                <Canvas Width="16" Height="16">
                                    <!--  Shield Icon  -->
                                    <Path Data="M8,0 L16,4 L12,16 L4,16 L0,4 Z" Fill="Gray" />
                                </Canvas>
                            </Viewbox>
                        </Border>
                        <TextBlock Style="{StaticResource WorkflowLabelStyle}" Text="Login" />
                        <!--  Erklärung für Schritt 2  -->
                        <TextBlock
                            FontSize="9"
                            Foreground="Gray"
                            Text="  2. Login zu SPO   "
                            TextAlignment="Center" />
                    </StackPanel>
                    <!--  Schritt 3: Source Verbindung  -->
                    <StackPanel>
                        <Border x:Name="workflowIcon3" Style="{StaticResource WorkflowIconContainer}">
                            <Viewbox>
                                <Canvas Width="16" Height="16">
                                    <!--  Kettenglied  -->
                                    <Path
                                        Data="M4,4 A2,2 0 1,1 4,12 A2,2 0 1,1 4,4 M12,4 A2,2 0 1,1 12,12 A2,2 0 1,1 12,4"
                                        Fill="Transparent"
                                        Stroke="Gray"
                                        StrokeThickness="2" />
                                    <Line
                                        Stroke="Gray"
                                        StrokeThickness="2"
                                        X1="4"
                                        X2="12"
                                        Y1="8"
                                        Y2="8" />
                                </Canvas>
                            </Viewbox>
                        </Border>
                        <TextBlock Style="{StaticResource WorkflowLabelStyle}" Text="Source" />
                        <!--  Erklärung für Schritt 3  -->
                        <TextBlock
                            FontSize="9"
                            Foreground="Gray"
                            Text="   3. Quell-Site abrufen   "
                            TextAlignment="Center" />
                    </StackPanel>
                    <!--  Schritt 4: Site Script exportieren  -->
                    <StackPanel>
                        <Border x:Name="workflowIcon4" Style="{StaticResource WorkflowIconContainer}">
                            <Viewbox>
                                <Canvas Width="16" Height="16">
                                    <!--  Pfeil nach oben  -->
                                    <Path
                                        Data="M8,16 L8,4 M8,4 L4,8 M8,4 L12,8"
                                        Stroke="Gray"
                                        StrokeEndLineCap="Round"
                                        StrokeLineJoin="Round"
                                        StrokeStartLineCap="Round"
                                        StrokeThickness="2" />
                                </Canvas>
                            </Viewbox>
                        </Border>
                        <TextBlock Style="{StaticResource WorkflowLabelStyle}" Text="Export" />
                        <!--  Erklärung für Schritt 4  -->
                        <TextBlock
                            FontSize="9"
                            Foreground="Gray"
                            Text="   4. Exportieren   "
                            TextAlignment="Center" />
                    </StackPanel>
                    <!--  Schritt 5: Site Design erstellen  -->
                    <StackPanel>
                        <Border x:Name="workflowIcon5" Style="{StaticResource WorkflowIconContainer}">
                            <Viewbox>
                                <Canvas Width="16" Height="16">
                                    <!--  Palette Icon (vereinfacht)  -->
                                    <Ellipse
                                        Width="16"
                                        Height="16"
                                        Fill="Transparent"
                                        Stroke="Gray"
                                        StrokeThickness="2" />
                                    <Ellipse
                                        Canvas.Left="2"
                                        Canvas.Top="2"
                                        Width="4"
                                        Height="4"
                                        Fill="Gray" />
                                </Canvas>
                            </Viewbox>
                        </Border>
                        <TextBlock Style="{StaticResource WorkflowLabelStyle}" Text="Design" />
                        <!--  Erklärung für Schritt 5  -->
                        <TextBlock
                            FontSize="9"
                            Foreground="Gray"
                            Text="   5. Site Design erstellen   "
                            TextAlignment="Center" />
                    </StackPanel>
                    <!--  Schritt 6: Design anwenden  -->
                    <StackPanel>
                        <Border x:Name="workflowIcon6" Style="{StaticResource WorkflowIconContainer}">
                            <Viewbox>
                                <Canvas Width="16" Height="16">
                                    <!--  Häkchen  -->
                                    <Path
                                        Data="M2,8 L6,12 L14,2"
                                        Fill="Transparent"
                                        Stroke="Gray"
                                        StrokeEndLineCap="Round"
                                        StrokeLineJoin="Round"
                                        StrokeStartLineCap="Round"
                                        StrokeThickness="2" />
                                </Canvas>
                            </Viewbox>
                        </Border>
                        <TextBlock Style="{StaticResource WorkflowLabelStyle}" Text="Anwenden" />
                        <!--  Erklärung für Schritt 6  -->
                        <TextBlock
                            FontSize="9"
                            Foreground="Gray"
                            Text="    6. Design anwenden   "
                            TextAlignment="Center" />
                    </StackPanel>
                </StackPanel>
            </GroupBox>

            <!--  Main Content Area - 3 Columns  -->
            <Grid Grid.Row="1">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="1*" />
                    <!--  Left Column: Form Fields  -->
                    <ColumnDefinition Width="1*" />
                    <!--  Middle Column: Configuration  -->
                    <ColumnDefinition Width="1*" />
                    <!--  Right Column: Status & Buttons  -->
                </Grid.ColumnDefinitions>

                <!--  Left Column: Form Fields  -->
                <StackPanel Grid.Column="0" Margin="0,0,45,0">
                    <TextBlock Style="{StaticResource GroupHeaderStyle}" Text="Policy -&gt; Module -&gt; Verbindung -&gt; " />

                    <!--  Execution Policy Group  -->
                    <GroupBox
                        Height="65"
                        Header="Execution Policy"
                        Style="{StaticResource ModernGroupBoxStyle}">
                        <StackPanel>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto" />
                                    <ColumnDefinition Width="*" />
                                </Grid.ColumnDefinitions>
                                <TextBlock
                                    Grid.Column="0"
                                    VerticalAlignment="Center"
                                    Style="{StaticResource LabelStyle}"
                                    Text="Aktuelle Policy:" />
                                <TextBlock
                                    x:Name="txtCurrentPolicy"
                                    Grid.Column="1"
                                    Margin="10,0,0,0"
                                    VerticalAlignment="Center"
                                    Text="{Binding Policy}" />
                            </Grid>
                            <CheckBox
                                x:Name="chkRemoteSigned"
                                Margin="0,10,0,0"
                                Content="RemoteSigned setzen"
                                Style="{StaticResource ModernCheckBoxStyle}" Width="163" />
                        </StackPanel>
                    </GroupBox>

                    <!--  Module Installation  -->
                    <GroupBox Header="Erforderliche Module" Style="{StaticResource ModernGroupBoxStyle}" Height="170">
                        <StackPanel>
                            <TextBlock
                                Margin="0,0,0,3"
                                Style="{StaticResource LabelStyle}"
                                Text="SPO Management Shell Modul:" />
                            <TextBlock
                                x:Name="txtSPOModuleStatus"
                                Margin="0,0,0,12"
                                Style="{StaticResource StatusInfoTextBlockStyle}"
                                Text="STATUS ÜBERPRÜFEN..." FontSize="10" />

                            <TextBlock
                                Style="{StaticResource LabelStyle}"
                                Text="PnP PowerShell Modul:" />
                            <TextBlock
                                x:Name="txtPnPModuleStatus"
                                Margin="0,0,0,12"
                                Style="{StaticResource StatusInfoTextBlockStyle}"
                                Text="STATUS ÜBERPRÜFEN..." FontSize="10" />

                            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                                <Button
                                    x:Name="btnInstallSPOModule"
                                    Width="165"
                                    Margin="0,0,10,0"
                                    Content="SPO Modul installieren"
                                    Style="{StaticResource ModernButtonStyle}" />

                                <Button
                                    x:Name="btnInstallPnPModule"
                                    Width="165"
                                    Content="PnP Modul installieren"
                                    Style="{StaticResource ModernButtonStyle}" />
                            </StackPanel>
                        </StackPanel>
                    </GroupBox>
                    <TextBlock
                        Style="{StaticResource LabelStyle}"
                        Text="Admin Center URL (z.B. https://IhrTenant-admin.sharepoint.com)" />
                    <TextBox x:Name="txtAdminUrl" Style="{StaticResource RequiredTextBoxStyle}" />
                    <GroupBox
                        Height="160"
                        Header="Login zu SharePoint Online"
                        Style="{StaticResource ModernGroupBoxStyle}">
                        <StackPanel>
                            <RadioButton 
                                x:Name="radioSPOLogin" 
                                IsChecked="True" 
                                Content="SPO Management Shell (einfach)" 
                                Style="{StaticResource ModernRadioButtonStyle}" />

                            <RadioButton 
                                x:Name="radioPnPLogin" 
                                Content="PnP PowerShell (Device Code)" 
                                Style="{StaticResource ModernRadioButtonStyle}" 
                                Margin="0,8,0,0" />

                            <TextBlock
                                x:Name="txtLoginStatus"
                                Margin="0,8,0,0"
                                Style="{StaticResource StatusErrorTextBlockStyle}"
                                Text="Nicht angemeldet" />

                            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                                <Button
                                    x:Name="btnLoginMFA"
                                    Width="165"
                                    Margin="0,0,10,0"
                                    Content="SPO mit MFA anmelden"
                                    Style="{StaticResource ModernButtonStyle}" />

                                <Button
                                    x:Name="btnLogin"
                                    Width="165"
                                    Content="Bei SPO anmelden"
                                    Style="{StaticResource ModernButtonStyle}" />
                            </StackPanel>
                        </StackPanel>
                    </GroupBox>

                    <!--  Login Panel - Now for SPO Management Shell  -->

                    <!--  Site URLs  -->
                </StackPanel>

                <!--  Middle Column: Configuration  -->
                <StackPanel Grid.ColumnSpan="2" Margin="380,0,31,0">
                    <TextBlock Style="{StaticResource GroupHeaderStyle}" Text="-&gt; Export -&gt; Einzel oder Mehrere Sites" />
                    <TextBlock Style="{StaticResource LabelStyle}" ><Run Text="Quell"/><Run Language="de-de" Text="e"/><Run Text=" (z.B. https://IhrTenant.sharepoint.com/sites/Quelle)"/></TextBlock>
                    <TextBox x:Name="txtSourceSite" Style="{StaticResource RequiredTextBoxStyle}" />

                    <!--  Target Options  -->
                    <GroupBox
                        Height="172"
                        Header="Export/Transfer Optionen"
                        Style="{StaticResource ModernGroupBoxStyle}">
                        <StackPanel>
                            <RadioButton
                                x:Name="radioSingleSite"
                                Content="Export zu einer Ziel-Website"
                                IsChecked="True"
                                Style="{StaticResource ModernRadioButtonStyle}" />
                            <Grid Margin="20,5,0,0">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto" />
                                    <ColumnDefinition Width="*" />
                                </Grid.ColumnDefinitions>
                                <TextBlock
                                    Grid.Column="0"
                                    VerticalAlignment="Center"
                                    Style="{StaticResource LabelStyle}"
                                    Text="Ziel:" />
                                <TextBox
                                    x:Name="txtTargetSite"
                                    Grid.Column="1"
                                    Margin="30,0,0,0"
                                    Style="{StaticResource ModernTextBoxStyle}" />
                            </Grid>

                            <RadioButton
                                x:Name="radioMultiSite"
                                Margin="0,15,0,0"
                                Content="Export zu mehreren Websites mit Praefix"
                                Style="{StaticResource ModernRadioButtonStyle}" />
                            <Grid Margin="20,5,0,0">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto" />
                                    <ColumnDefinition Width="*" />
                                </Grid.ColumnDefinitions>
                                <TextBlock
                                    Grid.Column="0"
                                    VerticalAlignment="Center"
                                    Style="{StaticResource LabelStyle}"
                                    Text="Präfix:" />
                                <TextBox
                                    x:Name="txtPrefix"
                                    Grid.Column="1"
                                    Margin="19,0,0,0"
                                    IsEnabled="False"
                                    Style="{StaticResource ModernTextBoxStyle}" />
                            </Grid>
                        </StackPanel>
                    </GroupBox>

                    <!--  Advanced Options  -->
                    <CheckBox
                        x:Name="chkAdvancedOptions"
                        Margin="0,10,0,0"
                        Content="Erweiterte Export-Optionen aktivieren"
                        Style="{StaticResource ModernCheckBoxStyle}" />

                    <GroupBox
                        x:Name="grpAdvancedOptions"
                        Height="75"
                        Header="Erweiterte Export-Optionen"
                        IsEnabled="False"
                        Style="{StaticResource ModernGroupBoxStyle}">
                        <WrapPanel>
                            <CheckBox
                                x:Name="chkIncludeBranding"
                                Margin="5"
                                Content="Branding"
                                IsChecked="True"
                                Style="{StaticResource ModernCheckBoxStyle}" />
                            <CheckBox
                                x:Name="chkIncludeTheme"
                                Margin="5"
                                Content="Theme"
                                IsChecked="True"
                                Style="{StaticResource ModernCheckBoxStyle}" />
                            <CheckBox
                                x:Name="chkIncludeNavigation"
                                Margin="5"
                                Content="Navigation"
                                IsChecked="True"
                                Style="{StaticResource ModernCheckBoxStyle}" />
                            <CheckBox
                                x:Name="chkIncludePages"
                                Margin="5"
                                Content="Seiten-Layouts"
                                IsChecked="True"
                                Style="{StaticResource ModernCheckBoxStyle}" />
                        </WrapPanel>
                    </GroupBox>

                    <!--  Site Design/Script Titles  -->
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*" />
                            <ColumnDefinition Width="*" />
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0" Grid.ColumnSpan="2">
                            <TextBlock
                                Margin="0,10,0,0"
                                Style="{StaticResource LabelStyle}"
                                Text="Titel für Site Design" />
                            <TextBox
                                x:Name="txtSiteDesign"
                                Style="{StaticResource ModernTextBoxStyle}"
                                ToolTip="Standard: ExportedDesign" />
                        </StackPanel>
                    </Grid>
                    <TextBlock Style="{StaticResource LabelStyle}" Text="Titel für Site Script" />
                    <TextBox
                        x:Name="txtSiteScript"
                        Style="{StaticResource ModernTextBoxStyle}"
                        ToolTip="Standard: ExportedScript" />
                </StackPanel>

                <!--  Right Column: Status & Buttons  -->
                <StackPanel
                    Grid.Column="1"
                    Grid.ColumnSpan="2"
                    Margin="394,0,0,0">
                    <TextBlock Style="{StaticResource GroupHeaderStyle}" Text="-&gt; Status der Ausführung" />
                    <Border
                        Padding="10"
                        Background="#E3F2FD"
                        BorderBrush="#2196F3"
                        BorderThickness="1" Height="55" Width="429">
                        <TextBlock
                            Foreground="#0D47A1"
                            TextWrapping="Wrap" ><Run Text="HINWEIS: Dieses Tool exportiert nur Site-Design-Elemente "/><LineBreak/><Run Text="                 - "/><Run Text="Navigation, Themes, Layouts"/></TextBlock>
                    </Border>

                    <!--  Status Panel  -->
                    <GroupBox
                        Height="240"
                        Header="Status Meldungen"
                        Style="{StaticResource ModernGroupBoxStyle}">
                        <StackPanel>
                            <ScrollViewer Height="150" VerticalScrollBarVisibility="Auto">
                                <TextBlock
                                    x:Name="txtStatus"
                                    Style="{StaticResource StatusInfoTextBlockStyle}"
                                    Text="Bereit"
                                    TextWrapping="Wrap" Height="140" />
                            </ScrollViewer>

                            <!-- Device Code Information Panel -->
                            <TextBlock 
                                x:Name="txtDeviceCodeInfo" 
                                Margin="0,10,0,5"
                                Text="Bitte besuchen Sie microsoft.com/devicelogin und geben Sie den folgenden Code ein:" 
                                Visibility="Collapsed"
                                Style="{StaticResource StatusInfoTextBlockStyle}" />

                            <Grid x:Name="deviceCodeGrid" Visibility="Collapsed">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*" />
                                    <ColumnDefinition Width="Auto" />
                                </Grid.ColumnDefinitions>

                                <Border 
                                    Grid.Column="0"
                                    Background="#E3F2FD" 
                                    BorderBrush="#2196F3" 
                                    BorderThickness="1" 
                                    Padding="10,5">
                                    <TextBlock 
                                        x:Name="txtDeviceCode" 
                                        FontSize="16"
                                        FontWeight="Bold"
                                        HorizontalAlignment="Center" />
                                </Border>

                                <Button 
                                    x:Name="btnCopyCode" 
                                    Grid.Column="1"
                                    Content="Kopieren" 
                                    Margin="5,0,0,0"
                                    Style="{StaticResource SecondaryButtonStyle}" />
                            </Grid>
                        </StackPanel>
                    </GroupBox>

                    <!--  Progress Bar  -->
                    <TextBlock
                        Margin="0,10,0,0"
                        Style="{StaticResource LabelStyle}"
                        Text="Fortschritt" />
                    <ProgressBar
                        x:Name="progressOperation"
                        Margin="0,0,0,20"
                        Style="{StaticResource ModernProgressBarStyle}"
                        Value="0" />

                    <!--  Operations Buttons  -->
                    <StackPanel Margin="0,10,0,0" HorizontalAlignment="Center">
                        <Button
                            x:Name="btnValidateInputs"
                            Width="150"
                            Margin="0,5"
                            HorizontalAlignment="Right"
                            Content="Eingaben pruefen"
                            Style="{StaticResource ModernButtonStyle}" />
                        <Button
                            x:Name="btnExportDesign"
                            Width="150"
                            Margin="0,5"
                            HorizontalAlignment="Right"
                            Content="Design exportieren"
                            Style="{StaticResource ModernButtonStyle}" />
                        <Button
                            x:Name="btnCreateScriptDesign"
                            Width="150"
                            Margin="0,5"
                            HorizontalAlignment="Right"
                            Content="Script/Design erstellen"
                            IsEnabled="False"
                            Style="{StaticResource ModernButtonStyle}" />
                        <Button
                            x:Name="btnApplyDesign"
                            Width="150"
                            Margin="0,5"
                            HorizontalAlignment="Right"
                            Content="Design anwenden"
                            IsEnabled="False"
                            Style="{StaticResource ModernButtonStyle}" />
                    </StackPanel>
                </StackPanel>
            </Grid>
        </Grid>

        <!--  Footer  -->
        <Border
            Grid.Row="2"
            Background="#F0F0F0"
            BorderBrush="#E0E0E0"
            BorderThickness="0,1,0,0">
            <Grid>
                <TextBlock Style="{StaticResource FooterTextBlockStyle}" Text="SharePoint Design Export Tool v0.2.0" />
                <TextBlock
                    Margin="0,0,20,0"
                    HorizontalAlignment="Right"
                    Style="{StaticResource FooterTextBlockStyle}"
                    Text="© 2025 Andreas Hepp | www.psscripts.de">
                    <TextBlock.ToolTip>https://psscripts.de</TextBlock.ToolTip>
                </TextBlock>
                <Button
                            x:Name="btnCancel"
                            Content="Abbrechen"
                            Style="{StaticResource SecondaryButtonStyle}" Margin="22,14,1128,15" Background="#FFCA4747" BorderBrush="Red" Foreground="White" FontWeight="Bold" FontSize="16" />
            </Grid>
        </Border>
    </Grid>
</Window>
"@

    # Die XAML in ein XML-Objekt konvertieren
    Write-Host "[DEBUG] Konvertiere XAML in XML-Objekt..." -ForegroundColor Cyan
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader] $xamlContent)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    Write-Host "[DEBUG] XAML wurde erfolgreich in Window-Objekt konvertiert" -ForegroundColor Green
    
    # GUI-Elemente als Hashtable speichern
    $gui = @{}
    
    # Verbesserte Methode zum Finden aller benannten Elemente
    Write-Host "[DEBUG] Suche benannte Elemente im GUI..." -ForegroundColor Cyan
    try {
        # Dummy-Aufruf, um FindName zu aktivieren
        $window.FindName('nothing')
        
        # Extrahiere Namen mit Regex (bewährte Methode)
        $elementNames = @()
        $regex = [regex]'x:Name="([^"]*)"'
        $matches = $regex.Matches($xamlContent)
        
        foreach ($match in $matches) {
            $elementNames += $match.Groups[1].Value
        }
        
        # Finde jedes Element und füge es zum GUI-Dictionary hinzu
        foreach ($elementName in $elementNames) {
            try {
                $element = $window.FindName($elementName)
                if ($null -ne $element) {
                    $gui[$elementName] = $element
                    # Reduziere Debug-Output für bessere Performance
                    #Write-Host "[DEBUG] Element gefunden: $elementName" -ForegroundColor Cyan
                }
                else {
                    Write-Host "[DEBUG] Element nicht gefunden: $elementName" -ForegroundColor Yellow
                }
            }
            catch {
                Write-Host "[DEBUG] Fehler beim Finden des Elements: $elementName - $($_.Exception.Message)" -ForegroundColor Red
                Write-Warning "Element not found: $elementName"
            }
        }
        
        Write-Host "[DEBUG] GUI-Elemente geladen: $($gui.Count) Elemente gefunden" -ForegroundColor Green
    }
    catch {
        Write-Host "[ERROR] Kritischer Fehler beim Laden der GUI-Elemente: $($_.Exception.Message)" -ForegroundColor Red
        throw "Fehler beim Laden der GUI: $($_.Exception.Message)"
    }
    
    # Fenster und GUI-Elemente zurückgeben
    return @{
        Window = $window
        GUI = $gui
    }
}

function Update-WorkflowIconStatus {
    param (
        [Parameter(Mandatory=$true)]
        [int]$StepNumber,
        
        [Parameter(Mandatory=$true)]
        [ValidateSet("Default", "Active", "Completed")]
        [string]$Status,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$GuiElements,
        
        [Parameter(Mandatory=$false)]
        [System.Windows.Window]$Window
    )
    
    Write-Host "[DEBUG] Update-WorkflowIconStatus: Schritt $StepNumber auf Status '$Status' setzen" -ForegroundColor Cyan
    
    if ($null -eq $GuiElements) {
        Write-Host "[DEBUG] GUI-Elemente fehlen - Status kann nicht aktualisiert werden" -ForegroundColor Red
        return
    }
    
    $iconName = "workflowIcon$StepNumber"
    $icon = $GuiElements[$iconName]
    
    if ($null -ne $icon) {
        Write-Host "[DEBUG] Icon '$iconName' gefunden, ändere Style zu $Status" -ForegroundColor Cyan
        switch ($Status) {
            "Default" {
                $icon.Style = $Window.FindResource("WorkflowIconContainer")
                $pathElement = FindPathInBorder $icon
                if ($pathElement) {
                    if ($pathElement.Fill -is [System.Windows.Media.SolidColorBrush]) {
                        $pathElement.Fill = [System.Windows.Media.Brushes]::Gray
                    }
                    if ($pathElement.Stroke -is [System.Windows.Media.SolidColorBrush]) {
                        $pathElement.Stroke = [System.Windows.Media.Brushes]::Gray
                    }
                }
            }
            "Active" {
                $icon.Style = $Window.FindResource("WorkflowIconActive")
                $pathElement = FindPathInBorder $icon
                if ($pathElement) {
                    if ($pathElement.Fill -is [System.Windows.Media.SolidColorBrush]) {
                        $pathElement.Fill = [System.Windows.Media.Brushes]::Green
                    }
                    if ($pathElement.Stroke -is [System.Windows.Media.SolidColorBrush]) {
                        $pathElement.Stroke = [System.Windows.Media.Brushes]::Green
                    }
                }
            }
            "Completed" {
                $icon.Style = $Window.FindResource("WorkflowIconCompleted")
                $pathElement = FindPathInBorder $icon
                if ($pathElement) {
                    if ($pathElement.Fill -is [System.Windows.Media.SolidColorBrush]) {
                        $pathElement.Fill = [System.Windows.Media.Brushes]::Blue
                    }
                    if ($pathElement.Stroke -is [System.Windows.Media.SolidColorBrush]) {
                        $pathElement.Stroke = [System.Windows.Media.Brushes]::Blue
                    }
                }
            }
        }
        Write-Host "[DEBUG] Icon-Status wurde auf $Status gesetzt" -ForegroundColor Green
    } else {
        Write-Host "[DEBUG] Icon '$iconName' wurde nicht gefunden" -ForegroundColor Red
    }
}

function FindPathInBorder {
    param (
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.Border]$Border
    )
    
    Write-Host "[DEBUG] Suche Path-Element in Border" -ForegroundColor Cyan
    
    # Get the child of the border (should be a Viewbox)
    $viewbox = $Border.Child
    if ($viewbox -is [System.Windows.Controls.Viewbox]) {
        # Get the child of the Viewbox (should be a Canvas)
        $canvas = $viewbox.Child
        if ($canvas -is [System.Windows.Controls.Canvas]) {
            # Look for Path or other drawable elements in the Canvas children
            foreach ($child in $canvas.Children) {
                if ($child -is [System.Windows.Shapes.Path]) {
                    Write-Host "[DEBUG] Path-Element in Border gefunden" -ForegroundColor Green
                    return $child
                }
                elseif ($child -is [System.Windows.Shapes.Line]) {
                    Write-Host "[DEBUG] Line-Element in Border gefunden" -ForegroundColor Green
                    return $child
                }
                elseif ($child -is [System.Windows.Shapes.Ellipse]) {
                    Write-Host "[DEBUG] Ellipse-Element in Border gefunden" -ForegroundColor Green
                    return $child
                }
            }
        }
    }
    
    Write-Host "[DEBUG] Kein zeichenbares Element in Border gefunden" -ForegroundColor Yellow
    return $null
}

# Event handler für "Copy Code" Button
function Copy-DeviceCodeToClipboard {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$GuiElements
    )
    
    Write-Host "[DEBUG] Kopiere Device Code in die Zwischenablage" -ForegroundColor Cyan
    
    if ($null -eq $GuiElements -or $null -eq $GuiElements.txtDeviceCode) {
        Write-Host "[DEBUG] GUI-Elemente oder txtDeviceCode fehlen" -ForegroundColor Red
        return
    }
    
    $code = $GuiElements.txtDeviceCode.Text
    
    if (-not [string]::IsNullOrEmpty($code) -and $code -ne "Wird generiert...") {
        # Code in die Zwischenablage kopieren
        Write-Host "[DEBUG] Code gefunden: $code - kopiere in Zwischenablage" -ForegroundColor Cyan
        $code | Set-Clipboard
        
        # Bestätigung anzeigen
        $GuiElements.txtStatus.Text = "Device Code wurde in die Zwischenablage kopiert. Bitte bei Microsoft.com/devicelogin eingeben."
        Write-Host "[DEBUG] Code wurde in die Zwischenablage kopiert" -ForegroundColor Green
    } else {
        Write-Host "[DEBUG] Kein Code zum Kopieren vorhanden oder Code wird noch generiert" -ForegroundColor Yellow
        $GuiElements.txtStatus.Text = "Device Code wird noch generiert oder ist nicht verfügbar."
    }
}

# Neuer Custom Credential Dialog für SPO Login ohne MFA
function Show-CredentialDialog {
    param (
        [string]$Message = "Bitte geben Sie Ihre SharePoint Online-Anmeldedaten ein",
        [string]$WindowTitle = "SharePoint Online Anmeldung"
    )
    
    Write-Host "[DEBUG] Zeige separaten Anmeldedialog" -ForegroundColor Cyan
    
    # Create the WPF window for credentials
    $xamlContent = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="$WindowTitle" 
    Width="450" 
    Height="380" 
    WindowStartupLocation="CenterScreen" 
    ResizeMode="NoResize"
    Background="#FFD8F984">
    <Window.Resources>
        <Style x:Key="HeaderTextBlockStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#202020" />
            <Setter Property="FontSize" Value="16" />
            <Setter Property="FontWeight" Value="SemiBold" />
            <Setter Property="Margin" Value="0,0,0,15" />
            <Setter Property="TextWrapping" Value="Wrap" />
        </Style>

        <Style x:Key="LabelStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#505050" />
            <Setter Property="FontSize" Value="13" />
            <Setter Property="Margin" Value="0,10,0,5" />
        </Style>

        <Style x:Key="ModernTextBoxStyle" TargetType="TextBox">
            <Setter Property="Height" Value="36" />
            <Setter Property="Padding" Value="10,0" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
            <Setter Property="Background" Value="White" />
            <Setter Property="BorderBrush" Value="#CCCCCC" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="Margin" Value="0,0,0,10" />
        </Style>

        <Style x:Key="ModernPasswordBoxStyle" TargetType="PasswordBox">
            <Setter Property="Height" Value="36" />
            <Setter Property="Padding" Value="10,0" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
            <Setter Property="Background" Value="White" />
            <Setter Property="BorderBrush" Value="#CCCCCC" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="Margin" Value="0,0,0,20" />
        </Style>

        <Style x:Key="ModernButtonStyle" TargetType="Button">
            <Setter Property="Height" Value="36" />
            <Setter Property="Width" Value="120" />
            <Setter Property="Padding" Value="15,0" />
            <Setter Property="Background" Value="#0078D7" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Cursor" Value="Hand" />
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#106EBE" />
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="CancelButtonStyle" TargetType="Button">
            <Setter Property="Height" Value="36" />
            <Setter Property="Width" Value="120" />
            <Setter Property="Padding" Value="15,0" />
            <Setter Property="Background" Value="#EFEFEF" />
            <Setter Property="Foreground" Value="#505050" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Cursor" Value="Hand" />
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#E0E0E0" />
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <Border Background="White" Margin="10" CornerRadius="5">
        <Border.Effect>
            <DropShadowEffect BlurRadius="10" ShadowDepth="1" Opacity="0.2" />
        </Border.Effect>

        <Grid Margin="20">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <!-- Header area with icon and text -->
            <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,15">
                <Border Width="40" Height="40" Background="#E3F2FD" CornerRadius="20" Margin="0,0,15,0">
                    <Path Data="M12,4A4,4 0 0,1 16,8A4,4 0 0,1 12,12A4,4 0 0,1 8,8A4,4 0 0,1 12,4M12,14C16.42,14 20,15.79 20,18V20H4V18C4,15.79 7.58,14 12,14Z" 
                          Fill="#2196F3" Stretch="Uniform" Margin="8"/>
                </Border>
                <TextBlock Style="{StaticResource HeaderTextBlockStyle}" Text="$Message" VerticalAlignment="Center"/>
            </StackPanel>

            <TextBlock Grid.Row="1" Text="Benutzername:" Style="{StaticResource LabelStyle}"/>
            <TextBox x:Name="txtUsername" Grid.Row="2" Style="{StaticResource ModernTextBoxStyle}"/>

            <TextBlock Grid.Row="3" Text="Passwort:" Style="{StaticResource LabelStyle}"/>
            <PasswordBox x:Name="txtPassword" Grid.Row="4" Style="{StaticResource ModernPasswordBoxStyle}" Margin="0,0,0,0" VerticalAlignment="Top"/>

            <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Bottom">
                <Button x:Name="btnCancel" Content="Abbrechen" Style="{StaticResource CancelButtonStyle}" Margin="0,0,15,0" IsCancel="True" FontSize="16" FontWeight="Bold"/>
                <Button x:Name="btnOK" Content="Anmelden" Style="{StaticResource ModernButtonStyle}" IsDefault="True" FontWeight="Bold" FontSize="16"/>
            </StackPanel>
        </Grid>
    </Border>
</Window>
"@
    
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader] $xamlContent)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    
    # Get controls
    $txtUsername = $window.FindName("txtUsername")
    $txtPassword = $window.FindName("txtPassword")
    $btnOK = $window.FindName("btnOK")
    $btnCancel = $window.FindName("btnCancel")
    
    # WICHTIG: Initialisiere Credential im äußeren Scope
    # damit der Event-Handler die Variable aktualisieren kann
    $script:dialogCredential = $null
    
    # OK button click event
    $btnOK.Add_Click({
        if (-not [string]::IsNullOrEmpty($txtUsername.Text) -and -not [string]::IsNullOrEmpty($txtPassword.Password)) {
            $securePassword = ConvertTo-SecureString $txtPassword.Password -AsPlainText -Force
            # Wichtig: Aktualisiere die Variable im äußeren Scope
            $script:dialogCredential = New-Object System.Management.Automation.PSCredential ($txtUsername.Text, $securePassword)
            Write-Host "[DEBUG] Anmeldedaten wurden eingegeben: $($txtUsername.Text)" -ForegroundColor Green
            $window.DialogResult = $true
            $window.Close()
        } else {
            [System.Windows.MessageBox]::Show("Bitte geben Sie Benutzername und Passwort ein.", "Fehlende Anmeldedaten", 
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        }
    })
    
    # Cancel button click event - explizit hinzugefügt
    $btnCancel.Add_Click({
        Write-Host "[DEBUG] Anmeldedialog wurde vom Benutzer abgebrochen" -ForegroundColor Yellow
        $script:dialogCredential = $null
        $window.DialogResult = $false
        $window.Close()
    })
    
    # Show dialog and wait for result
    $result = $window.ShowDialog()
    
    # Return credential if OK was clicked
    if ($result) {
        Write-Host "[DEBUG] Dialog wurde mit OK geschlossen, gebe Anmeldedaten zurück" -ForegroundColor Green
        return $script:dialogCredential
    } else {
        Write-Host "[DEBUG] Dialog wurde abgebrochen, gebe null zurück" -ForegroundColor Yellow
        return $null
    }
}

# Tenant-Info-Funktionen
function Get-TenantInfoFromUrl {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Url
    )

    Write-Host "[DEBUG] Extrahiere Tenant-Informationen aus URL: $Url" -ForegroundColor Cyan
    
    # URL säubern und normalisieren
    $Url = $Url.Trim().TrimEnd('/')
    
    # Prüfen ob URL gültige SharePoint Online URL ist
    if (-not ($Url -like "https://*.sharepoint.com*")) {
        Write-Host "[ERROR] Die URL ist keine gültige SharePoint Online URL: $Url" -ForegroundColor Red
        return $null
    }
    
    # Standard-Ergebnisobjekt
    $result = [PSCustomObject]@{
        TenantName = $null
        TenantId = $null
        AdminUrl = $null
        RootUrl = $null
        IsAdminSite = $false
        AuthorityUrl = $null
    }
    
    # Extrahiere Tenant-Namen
    if ($Url -match "https://([^.]+)-admin\.sharepoint\.com") {
        # Admin-URL
        $result.TenantName = $matches[1]
        $result.IsAdminSite = $true
    }
    elseif ($Url -match "https://([^.]+)\.sharepoint\.com") {
        # Standard SharePoint URL
        $result.TenantName = $matches[1]
    }
    else {
        # Unbekanntes Format
        Write-Host "[ERROR] Konnte Tenant-Namen nicht aus der URL extrahieren: $Url" -ForegroundColor Red
        return $null
    }
    
    # Konstruiere die verschiedenen URLs basierend auf dem Tenant-Namen
    $result.RootUrl = "https://$($result.TenantName).sharepoint.com"
    $result.AdminUrl = "https://$($result.TenantName)-admin.sharepoint.com"
    
    # Für die modernen authentifizierungsmethoden, setze die Authority URL 
    # (tenant-spezifischer Endpunkt für Azure AD)
    $result.AuthorityUrl = "https://login.microsoftonline.com/$($result.TenantName).onmicrosoft.com"
    
    # Tenant ID ist in diesem Fall nicht verfügbar ohne Azure AD Abfrage
    # Bei Bedarf kann hier eine Funktion ergänzt werden, um die TenantId zu ermitteln
    $result.TenantId = "$($result.TenantName).onmicrosoft.com"
    
    Write-Host "[DEBUG] Erfolgreich extrahierte Tenant-Infos:" -ForegroundColor Green
    Write-Host "[DEBUG] - Tenant: $($result.TenantName)" -ForegroundColor Green
    Write-Host "[DEBUG] - Admin URL: $($result.AdminUrl)" -ForegroundColor Green
    Write-Host "[DEBUG] - Root URL: $($result.RootUrl)" -ForegroundColor Green
    Write-Host "[DEBUG] - Authority: $($result.AuthorityUrl)" -ForegroundColor Green
    
    return $result
}

function Get-SPOSiteUrl {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Url,
        
        [Parameter(Mandatory = $false)]
        [switch]$AsAdminUrl
    )
    
    # Hole Tenant-Infos
    $tenantInfo = Get-TenantInfoFromUrl -Url $Url
    
    if ($null -eq $tenantInfo) {
        return $null
    }
    
    if ($AsAdminUrl) {
        return $tenantInfo.AdminUrl
    }
    else {
        # Falls es eine Admin-URL ist, gib die Root-URL zurück
        if ($tenantInfo.IsAdminSite) {
            return $tenantInfo.RootUrl
        }
        # Andernfalls gib die ursprüngliche URL zurück
        return $Url
    }
}

function Get-SPOAdminCenterUrl {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TenantName
    )
    
    # Bereinige Tenant-Namen
    $TenantName = $TenantName.Replace(".onmicrosoft.com", "").Replace("-admin", "")
    
    # Konstruiere Admin-URL
    $adminUrl = "https://$TenantName-admin.sharepoint.com"
    
    return $adminUrl
}

function Parse-SPOUrl {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Url
    )
    
    $urlParts = @{
        TenantName = $null
        SiteName = $null
        IsAdminSite = $false
        IsSubsite = $false
        SubsitePath = $null
        FullSitePath = $null
    }
    
    # Säubere die URL
    $Url = $Url.Trim().TrimEnd('/')
    
    # Admin Site
    if ($Url -match "https://([^.]+)-admin\.sharepoint\.com(/.+)?") {
        $urlParts.TenantName = $matches[1]
        $urlParts.IsAdminSite = $true
        $urlParts.FullSitePath = $Url
    }
    # Root-Site
    elseif ($Url -match "https://([^.]+)\.sharepoint\.com/?$") {
        $urlParts.TenantName = $matches[1]
        $urlParts.SiteName = ""  # Root-Site
        $urlParts.FullSitePath = $Url
    }
    # Site Collection
    elseif ($Url -match "https://([^.]+)\.sharepoint\.com/sites/([^/]+)(/.*)?") {
        $urlParts.TenantName = $matches[1]
        $urlParts.SiteName = $matches[2]
        
        if (($matches.Count > 3) -and ($matches[3])) {
            $urlParts.IsSubsite = $true
            $urlParts.SubsitePath = $matches[3]
        }
        
        $urlParts.FullSitePath = $Url
    }
    # Teams/Gruppen Site Collection
    elseif ($Url -match "https://([^.]+)\.sharepoint\.com/teams/([^/]+)(/.*)?") {
        $urlParts.TenantName = $matches[1]
        $urlParts.SiteName = $matches[2]
        
        if (($matches.Count > 3) -and ($matches[3])) {
            $urlParts.IsSubsite = $true
            $urlParts.SubsitePath = $matches[3]
        }
        
        $urlParts.FullSitePath = $Url
    }
    else {
        Write-Host "[ERROR] Konnte URL nicht parsen: $Url" -ForegroundColor Red
        return $null
    }
    
    return $urlParts
}

# SharePoint-Verbindung
function Connect-ToSharePoint {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Url,
        
        [Parameter(Mandatory=$true)]
        [ValidateSet("SPO-Standard", "SPO-MFA", "PnP-DeviceCode", "SPO-AppPassword", "PnP-Interactive")]
        [string]$Method,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$GuiElements,
        
        [Parameter(Mandatory=$false)]
        [System.Windows.Window]$Window,
        
        [Parameter(Mandatory=$false)]
        [System.Management.Automation.PSCredential]$AppCredential
    )
    
    Write-Host "[DEBUG] Connect-ToSharePoint: Verbinde zu $Url mit Methode $Method" -ForegroundColor Cyan
    
    try {
        # Extrahiere Tenant-Informationen aus der URL
        $tenantInfo = Get-TenantInfoFromUrl -Url $Url
        if ($null -eq $tenantInfo) {
            throw "Tenant-Informationen konnten nicht aus der URL extrahiert werden."
        }
        
        # Verwende normalisierte Admin-URL, wenn es eine Admin-URL sein sollte
        if ($Url -like "*-admin.sharepoint.com*") {
            $normalizedUrl = $tenantInfo.AdminUrl
            Write-Host "[DEBUG] Normalisierte Admin-URL: $normalizedUrl" -ForegroundColor Cyan
        } else {
            $normalizedUrl = $Url
        }
        
        # Setze Authority-URL für moderne Authentifizierung
        $authorityUrl = $tenantInfo.AuthorityUrl
        Write-Host "[DEBUG] Authority-URL für Authentifizierung: $authorityUrl" -ForegroundColor Cyan
        
        switch ($Method) {
            "SPO-Standard" {
                # Import SPO Management Module
                Import-Module -Name $script:spoModuleName -ErrorAction Stop
                
                # SPO-Verbindung ohne MFA - Separater Anmeldedialog
                $cred = Show-CredentialDialog -Message "Bitte geben Sie Ihre SharePoint Online-Anmeldedaten ein" -WindowTitle "SPO Login"
                
                if ($null -ne $cred) {
                    # Zusätzliches Debug-Logging
                    Write-Host "[DEBUG] Anmeldedaten wurden zurückgegeben, starte Verbindung..." -ForegroundColor Cyan
                    Write-Host "[DEBUG] Benutzername: $($cred.UserName)" -ForegroundColor Cyan
                    
                    try {
                        # Versuche zunächst mit tenant-spezifischer URL
                        Write-Host "[DEBUG] Versuche Verbindung mit tenant-spezifischem Endpunkt" -ForegroundColor Cyan
                        if ($GuiElements) {
                            $GuiElements.txtStatus.Text = "Verbinde mit tenant-spezifischem Endpunkt..."
                        }
                        
                        # Verwende TenantId für die Authentifizierung
                        Connect-SPOService -Url $normalizedUrl -Credential $cred -ErrorAction Stop
                        
                        # Nach erfolgreicher Verbindung
                        Write-Host "[DEBUG] SPO-Verbindung erfolgreich hergestellt (Standard mit tenant-spezifischem Endpunkt)" -ForegroundColor Green
                        if ($GuiElements) {
                            $GuiElements.txtLoginStatus.Text = "Angemeldet via SPO Management Shell (Standard)"
                            $GuiElements.txtLoginStatus.Style = $Window.FindResource("StatusOkTextBlockStyle")
                            $GuiElements.txtStatus.Text = "Login zu SharePoint Online erfolgreich. Sie können nun den Design-Export starten."
                        }
                        return $true
                    }
                    catch [System.Management.Automation.CommandNotFoundException] {
                        throw  # Modul-Fehler weiterleiten
                    }
                    catch {
                        # Bei Azure AD-spezifischen Fehlern
                        if ($_.Exception.Message -match "AADSTS9001023") {
                            Write-Host "[DEBUG] Tenant-spezifischer Endpunkt erforderlich: $($_.Exception.Message)" -ForegroundColor Yellow
                            if ($GuiElements) {
                                $GuiElements.txtStatus.Text = "Authentifizierungsfehler: Ihr Tenant erfordert moderne Authentifizierungsmethoden."
                                
                                [System.Windows.MessageBox]::Show(
                                    "Authentifizierungsfehler: Ihr Tenant erfordert moderne Authentifizierungsmethoden.`n`n" +
                                    "Der Fehler weist darauf hin, dass Sie eine der folgenden Methoden verwenden sollten:`n" +
                                    "- SPO mit MFA-Unterstützung`n" +
                                    "- PnP PowerShell mit Device Code oder Interactive Login`n`n" +
                                    "Möchten Sie eine dieser Methoden jetzt ausprobieren?", 
                                    "Tenant erfordert moderne Authentifizierung", 
                                    [System.Windows.MessageBoxButton]::OK, 
                                    [System.Windows.MessageBoxImage]::Information
                                )
                            }
                            throw "Tenant erfordert moderne Authentifizierung. Verwenden Sie SPO-MFA oder PnP-Methoden."
                        }
                        else {
                            # Andere Fehler weiterleiten
                            throw
                        }
                    }
                } else {
                    Write-Host "[DEBUG] Anmeldung wurde vom Benutzer abgebrochen oder Anmeldedaten waren ungültig" -ForegroundColor Yellow
                    if ($GuiElements) {
                        $GuiElements.txtStatus.Text = "Anmeldung wurde abgebrochen oder Anmeldedaten waren ungültig."
                    }
                    return $false
                }
            }
            "SPO-MFA" {
                # Import SPO Management Module
                Import-Module -Name $script:spoModuleName -ErrorAction Stop
                
                # Überprüfen verfügbarer MFA-Parameter für unterschiedliche Modulversionen
                $commandInfo = Get-Command Connect-SPOService -ErrorAction Stop
                $hasModernAuth = $commandInfo.Parameters.ContainsKey("ModernAuth")
                $hasUseWebLogin = $commandInfo.Parameters.ContainsKey("UseWebLogin")
                $hasDeviceAuth = $commandInfo.Parameters.ContainsKey("UseDeviceAuthentication") 
                $hasMfaAuth = $commandInfo.Parameters.ContainsKey("MfaAuthentication")
                
                Write-Host "[DEBUG] SPO Modul Parameter: ModernAuth=$hasModernAuth, UseWebLogin=$hasUseWebLogin, UseDeviceAuthentication=$hasDeviceAuth, MfaAuthentication=$hasMfaAuth" -ForegroundColor Cyan
                
                if ($GuiElements) {
                    $GuiElements.txtStatus.Text = "Verbinde mit MFA-Unterstützung zu SharePoint Online..."
                }
                
                # Verbesserte MFA-Implementierung mit tenant-spezifischen Endpunkten
                try {
                    # Versuche zuerst mit normalisierten Endpunkten
                    Write-Host "[DEBUG] Versuche direkte Verbindung mit tenant-spezifischem Endpunkt" -ForegroundColor Cyan
                    if ($GuiElements) {
                        $GuiElements.txtStatus.Text = "Verbinde mit tenant-spezifischem Endpunkt (MFA)..."
                    }
                    
                    # Verwende die AuthorityUrl als AzureEnvironment, falls möglich
                    if ($commandInfo.Parameters.ContainsKey("AzureEnvironment")) {
                        Write-Host "[DEBUG] Verwende AzureEnvironment-Parameter für tenant-spezifischen Endpunkt" -ForegroundColor Cyan
                        Connect-SPOService -Url $normalizedUrl -AzureEnvironment $tenantInfo.TenantId -ErrorAction Stop
                    } else {
                        # Standard-Verbindung ohne zusätzliche Parameter
                        Connect-SPOService -Url $normalizedUrl -ErrorAction Stop
                    }
                    
                    # Bei Erfolg
                    Write-Host "[DEBUG] SPO-Verbindung mit tenant-spezifischem Endpunkt erfolgreich hergestellt (MFA)" -ForegroundColor Green
                    if ($GuiElements) {
                        $GuiElements.txtLoginStatus.Text = "Angemeldet via SPO Management Shell (MFA)"
                        $GuiElements.txtLoginStatus.Style = $Window.FindResource("StatusOkTextBlockStyle")
                        $GuiElements.txtStatus.Text = "Login zu SharePoint Online erfolgreich. Sie können nun den Design-Export starten."
                    }
                    return $true
                }
                catch {
                    # Bei Fehlern spezifische MFA-Methoden ausprobieren
                    Write-Host "[DEBUG] Direkte tenant-spezifische Verbindung fehlgeschlagen: $($_.Exception.Message)" -ForegroundColor Yellow
                    
                    if ($_.Exception.Message -match "(400) Bad Request") {
                        Write-Host "[DEBUG] 400 Bad Request - Dies deutet auf ein Problem mit der Authentifizierungsmethode hin" -ForegroundColor Yellow
                        if ($GuiElements) {
                            $GuiElements.txtStatus.Text = "SPO-MFA Login mit 400-Fehler fehlgeschlagen. Empfehlung: Verwenden Sie PnP PowerShell."
                            
                            [System.Windows.MessageBox]::Show(
                                "Die SPO Management Shell MFA-Authentifizierung ist fehlgeschlagen (400 Bad Request).`n`n" +
                                "Dieser Fehler tritt oft auf, wenn der Tenant bestimmte moderne Authentifizierungsanforderungen hat.`n`n" +
                                "Empfehlung: Verwenden Sie die PnP PowerShell-Option mit Device Code oder interaktivem Login.", 
                                "SPO MFA Authentifizierung nicht möglich", 
                                [System.Windows.MessageBoxButton]::OK, 
                                [System.Windows.MessageBoxImage]::Warning
                            )
                        }
                        # Schalte um auf PnP als Alternative
                        if ($GuiElements) {
                            $GuiElements.radioPnPLogin.IsChecked = $true
                        }
                        return $false
                    }
                    
                    # Alternative Methoden probieren wie zuvor
                    # ...existing code...
                }
            }
            "PnP-DeviceCode" {
                # Import PnP Module
                Import-Module -Name $script:pnpModuleName -ErrorAction Stop
                
                # Vorbereitung der UI für Device Code-Anzeige
                if ($GuiElements) {
                    $GuiElements.txtDeviceCodeInfo.Visibility = "Visible"
                    $GuiElements.deviceCodeGrid.Visibility = "Visible"
                    $GuiElements.txtDeviceCode.Text = "Wird generiert..."
                    $GuiElements.txtStatus.Text = "Warte auf Device Code von Microsoft..."
                }
                
                try {
                    # Verbesserte Device Login mit PnP
                    $connection = Connect-PnPOnline -Url $normalizedUrl -DeviceLogin -ReturnConnection -TenantAdminUrl $tenantInfo.AdminUrl -ErrorAction Stop
                    
                    # Nach erfolgreicher Verbindung
                    if ($connection) {
                        Write-Host "[DEBUG] PnP-Verbindung mit Device Code erfolgreich hergestellt" -ForegroundColor Green
                        if ($GuiElements) {
                            $GuiElements.txtDeviceCode.Text = "[Authentifizierung erfolgreich]"
                            $GuiElements.txtLoginStatus.Text = "Angemeldet via PnP (Device Code)"
                            $GuiElements.txtLoginStatus.Style = $Window.FindResource("StatusOkTextBlockStyle")
                            $GuiElements.txtStatus.Text = "Device Code Login erfolgreich. Sie können nun den Design-Export starten."
                            
                            # Device Code UI ausblenden nach erfolgreichem Login
                            $GuiElements.deviceCodeGrid.Visibility = "Collapsed"
                            $GuiElements.txtDeviceCodeInfo.Visibility = "Collapsed"
                        }
                        return $true
                    } else {
                        throw "Verbindung wurde nicht hergestellt"
                    }
                }
                catch {
                    if ($GuiElements) {
                        # Device Code UI ausblenden bei Fehler
                        $GuiElements.deviceCodeGrid.Visibility = "Collapsed"
                        $GuiElements.txtDeviceCodeInfo.Visibility = "Collapsed"
                    }
                    throw
                }
            }
            "PnP-Interactive" {
                # Import PnP Module
                Import-Module -Name $script:pnpModuleName -ErrorAction Stop
                
                # Interactive Login mit PnP (MFA-freundliches WebLogin)
                Write-Host "[DEBUG] Versuche interaktiven PnP-Login" -ForegroundColor Cyan
                if ($GuiElements) {
                    $GuiElements.txtStatus.Text = "Öffne interaktiven Anmeldedialog..."
                }
                
                # Verwende Interactive Parameter für modern auth mit Tenant-Information
                $connection = Connect-PnPOnline -Url $normalizedUrl -Interactive -ReturnConnection -TenantAdminUrl $tenantInfo.AdminUrl -ErrorAction Stop
                
                if ($connection) {
                    Write-Host "[DEBUG] PnP-Verbindung mit Interactive Login erfolgreich hergestellt" -ForegroundColor Green
                    if ($GuiElements) {
                        $GuiElements.txtLoginStatus.Text = "Angemeldet via PnP (Interaktiv)"
                        $GuiElements.txtLoginStatus.Style = $Window.FindResource("StatusOkTextBlockStyle")
                        $GuiElements.txtStatus.Text = "Interaktiver Login erfolgreich. Sie können nun den Design-Export starten."
                    }
                    return $true
                } else {
                    throw "PnP interaktive Verbindung wurde nicht hergestellt"
                }
            }
            "SPO-AppPassword" {
                # Import SPO Management Module
                Import-Module -Name $script:spoModuleName -ErrorAction Stop
                
                # App-Passwort Methode für unbeaufsichtigte Skripts
                if ($null -eq $AppCredential) {
                    # Anmeldedialog für App-Passwort mit Hinweis
                    $message = "Bitte geben Sie Ihre SharePoint Online-Anmeldedaten ein`n" +
                               "Hinweis: Verwenden Sie ein App-Passwort, das Sie unter https://aka.ms/createapppassword erstellt haben"
                    $cred = Show-CredentialDialog -Message $message -WindowTitle "SPO App-Passwort Login"
                    
                    if ($null -eq $cred) {
                        if ($GuiElements) {
                            $GuiElements.txtStatus.Text = "App-Passwort Login abgebrochen."
                        }
                        return $false
                    }
                    
                    $AppCredential = $cred
                }
                
                # Verbindung mit App-Passwort herstellen
                Write-Host "[DEBUG] Versuche Anmeldung mit App-Passwort" -ForegroundColor Cyan
                if ($GuiElements) {
                    $GuiElements.txtStatus.Text = "Verbinde mit App-Passwort..."
                }
                
                Connect-SPOService -url $Url -Credential $AppCredential -ErrorAction Stop
                
                # Nach erfolgreicher Verbindung
                Write-Host "[DEBUG] SPO-Verbindung mit App-Passwort erfolgreich hergestellt" -ForegroundColor Green
                if ($GuiElements) {
                    $GuiElements.txtLoginStatus.Text = "Angemeldet via App-Passwort"
                    $GuiElements.txtLoginStatus.Style = $Window.FindResource("StatusOkTextBlockStyle")
                    $GuiElements.txtStatus.Text = "Login zu SharePoint Online mit App-Passwort erfolgreich. Sie können nun den Design-Export starten."
                }
                return $true
            }
        }
    }
    catch [System.Management.Automation.CommandNotFoundException] {
        Write-Host "[ERROR] Befehl nicht gefunden. Stellen Sie sicher, dass die Module installiert sind." -ForegroundColor Red
        if ($GuiElements) {
            $GuiElements.txtStatus.Text = "Fehler: Benötigter Befehl nicht gefunden. Bitte stellen Sie sicher, dass alle Module installiert und importiert sind."
            $GuiElements.txtStatus.Style = $Window.FindResource("StatusErrorTextBlockStyle")
        }
        return $false
    }
    catch [System.Net.WebException] {
        $errorMsg = $_.Exception.Message
        
        # Spezielle Fehlerbehandlung für typische MFA-bezogene Fehler
        if ($errorMsg -match "The sign-in name or password does not match") {
            Write-Host "[ERROR] Authentifizierungsfehler - Dies ist typisch bei aktivierter MFA ohne passende Login-Methode." -ForegroundColor Red
            if ($GuiElements) {
                $GuiElements.txtStatus.Text = "MFA-Fehler: Die gewählte Login-Methode unterstützt keine MFA. Bitte versuchen Sie eine MFA-kompatible Option."
                $GuiElements.txtStatus.Style = $Window.FindResource("StatusErrorTextBlockStyle")
            }
        }
        else {
            Write-Host "[ERROR] Netzwerk-Fehler: $errorMsg" -ForegroundColor Red
            if ($GuiElements) {
                $GuiElements.txtStatus.Text = "Netzwerkfehler: $errorMsg. Bitte überprüfen Sie Ihre Verbindung."
                $GuiElements.txtStatus.Style = $Window.FindResource("StatusErrorTextBlockStyle")
            }
        }
        
        return $false
    }
    catch [Microsoft.SharePoint.Client.IdcrlException] {
        # Spezifische Fehlerbehandlung für IdcrlException (häufig bei MFA-Problemen)
        Write-Host "[ERROR] SharePoint Authentifizierungsfehler: $($_.Exception.Message) - Möglicherweise MFA-Problem." -ForegroundColor Red
        if ($GuiElements) {
            $GuiElements.txtStatus.Text = "SharePoint Authentifizierungsfehler: Bitte verwenden Sie eine MFA-kompatible Anmeldemethode."
            $GuiElements.txtStatus.Style = $Window.FindResource("StatusErrorTextBlockStyle")
        }
        return $false
    }
    catch {
        # Verbesserte Fehlerbehandlung für Azure AD-Authentifizierungsfehler
        if ($_.Exception.Message -match "AADSTS") {
            Write-Host "[ERROR] Azure AD Authentifizierungsfehler: $($_.Exception.Message)" -ForegroundColor Red
            
            if ($_.Exception.Message -match "AADSTS9001023") {
                if ($GuiElements) {
                    $GuiElements.txtStatus.Text = "Azure AD Authentifizierungsfehler: Ihr Tenant erfordert moderne Authentifizierungsmethoden. Bitte verwenden Sie PnP PowerShell."
                    $GuiElements.txtStatus.Style = $Window.FindResource("StatusErrorTextBlockStyle")
                }
            } 
            else {
                if ($GuiElements) {
                    $GuiElements.txtStatus.Text = "Azure AD Authentifizierungsfehler: $($_.Exception.Message)"
                    $GuiElements.txtStatus.Style = $Window.FindResource("StatusErrorTextBlockStyle")
                }
            }
        }
        else {
            Write-Host "[ERROR] Fehler bei der SharePoint-Verbindung: $($_.Exception.Message)" -ForegroundColor Red
            if ($GuiElements) {
                $GuiElements.txtStatus.Text = "Fehler bei der SharePoint-Verbindung: $($_.Exception.Message)"
                $GuiElements.txtStatus.Style = $Window.FindResource("StatusErrorTextBlockStyle")
            }
        }
        return $false
    }
}

# Funktion zum Exportieren des Site Designs
function Export-SPOSiteDesign {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory=$false)]
        [hashtable]$ExportOptions = @{
            IncludeBranding = $true
            IncludeTheme = $true
            IncludeNavigationSettings = $true
            IncludePages = $true
        },
        
        [Parameter(Mandatory=$false)]
        [hashtable]$GuiElements,
        
        [Parameter(Mandatory=$false)]
        [System.Windows.Window]$Window
    )
    
    Write-Host "[DEBUG] Export-SPOSiteDesign: Exportiere von $SiteUrl" -ForegroundColor Cyan
    
    try {
        # PnP PowerShell für den Export verwenden
        Import-Module -Name $script:pnpModuleName -ErrorAction Stop
        
        # Verbindung zur Site herstellen, falls noch nicht verbunden
        Write-Host "[DEBUG] Prüfe, ob Verbindung zur Site $SiteUrl besteht" -ForegroundColor Cyan
        try {
            $ctx = Get-PnPContext -ErrorAction SilentlyContinue
            $connected = ($null -ne $ctx)
            
            if (-not $connected) {
                Write-Host "[DEBUG] Keine aktive Verbindung gefunden, stelle neue Verbindung her" -ForegroundColor Yellow
                if ($GuiElements) {
                    $GuiElements.txtStatus.Text = "Verbinde zur Quell-Website..."
                }
                
                Connect-PnPOnline -Url $SiteUrl -DeviceLogin -ErrorAction Stop
                Write-Host "[DEBUG] Verbindung zur Quell-Website hergestellt" -ForegroundColor Green
                
                if ($GuiElements) {
                    $GuiElements.txtStatus.Text = "Verbindung zur Quell-Website hergestellt. Starte Export..."
                }
            }
        }
        catch {
            Write-Host "[ERROR] Verbindung zur Site fehlgeschlagen: $($_.Exception.Message)" -ForegroundColor Red
            if ($GuiElements) {
                $GuiElements.txtStatus.Text = "Fehler bei der Verbindung zur Quell-Website: $($_.Exception.Message)"
                $GuiElements.txtStatus.Style = $Window.FindResource("StatusErrorTextBlockStyle")
            }
            return $null
        }
        
        # Site Design exportieren mit PnP
        Write-Host "[DEBUG] Exportiere Site Design mit den Optionen: $($ExportOptions | ConvertTo-Json -Compress)" -ForegroundColor Cyan
        if ($GuiElements) {
            $GuiElements.txtStatus.Text = "Exportiere Site Design..."
        }
        
        # Tatsächlicher Export mit PnP anstatt Simulation
        $script:exportedSiteScript = Get-PnPSiteScriptFromWeb -IncludeBranding $ExportOptions.IncludeBranding `
                                                         -IncludeTheme $ExportOptions.IncludeTheme `
                                                         -IncludeNavigationSettings $ExportOptions.IncludeNavigationSettings `
                                                         -IncludeRegionalSettings $true `
                                                         -IncludeLinksToExportedItems $true `
                                                         -IncludePages $ExportOptions.IncludePages `
                                                         -ErrorAction Stop
        
        Write-Host "[DEBUG] Site Design erfolgreich exportiert" -ForegroundColor Green
        if ($GuiElements) {
            $GuiElements.txtStatus.Text = "Site Design wurde erfolgreich exportiert. Sie können nun das Script/Design erstellen."
        }
        
        return $script:exportedSiteScript
    }
    catch [System.Management.Automation.CommandNotFoundException] {
        Write-Host "[ERROR] PnP PowerShell Befehl nicht gefunden. Stellen Sie sicher, dass das Modul installiert ist." -ForegroundColor Red
        if ($GuiElements) {
            $GuiElements.txtStatus.Text = "Fehler: PnP PowerShell Befehl nicht gefunden. Bitte installieren Sie das PnP PowerShell Modul."
            $GuiElements.txtStatus.Style = $Window.FindResource("StatusErrorTextBlockStyle")
        }
        return $null
    }
    catch {
        Write-Host "[ERROR] Fehler beim Exportieren des Site Designs: $($_.Exception.Message)" -ForegroundColor Red
        if ($GuiElements) {
            $GuiElements.txtStatus.Text = "Fehler beim Exportieren des Site Designs: $($_.Exception.Message)"
            $GuiElements.txtStatus.Style = $Window.FindResource("StatusErrorTextBlockStyle")
        }
        return $null
    }
}

# Funktion zur Erstellung von Site Script und Site Design
function Create-SPOSiteScriptAndDesign {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SiteScriptContent,
        
        [Parameter(Mandatory=$true)]
        [string]$SiteScriptTitle,
        
        [Parameter(Mandatory=$true)]
        [string]$SiteDesignTitle,
        
        [Parameter(Mandatory=$false)]
        [string]$Description = "Erstellt durch SPO Design Export Tool",
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("CommunicationSite", "TeamSite", "Both")]
        [string]$WebTemplate = "Both",
        
        [Parameter(Mandatory=$false)]
        [hashtable]$GuiElements,
        
        [Parameter(Mandatory=$false)]
        [System.Windows.Window]$Window,
        
        [Parameter(Mandatory=$false)]
        [bool]$UsePnP = $false
    )
    
    # ...existing code...
}

# ...existing code...