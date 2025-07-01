[xml]$Global:XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    Title="easyADReport v0.4.2"
    Width="1775"
    Height="1000"
    Background="#F8F9FA"
    FontFamily="Segoe UI"
    ResizeMode="CanResizeWithGrip"
    WindowStartupLocation="CenterScreen">

    <!-- Window Resources for Modern Styling -->
    <Window.Resources>
        <!-- Modern Card Style -->
        <Style x:Key="ModernCard" TargetType="Border">
            <Setter Property="Background" Value="White"/>
            <Setter Property="CornerRadius" Value="12"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>

        <!-- Modern Button Style -->
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#E5E7EB"/>
            <Setter Property="Padding" Value="16,10"/>
            <Setter Property="FontWeight" Value="Medium"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" 
                                Background="{TemplateBinding Background}" 
                                BorderBrush="{TemplateBinding BorderBrush}" 
                                BorderThickness="{TemplateBinding BorderThickness}" 
                                CornerRadius="8">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"
                                              Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#F9FAFB"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="#3B82F6"/>
                                <Setter TargetName="border" Property="Effect">
                                    <Setter.Value>
                                        <DropShadowEffect BlurRadius="6" ShadowDepth="1" Color="#103B82F6" Opacity="0.1"/>
                                    </Setter.Value>
                                </Setter>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#F3F4F6"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="#1D4ED8"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Primary Button Style -->
        <Style x:Key="PrimaryButton" TargetType="Button">
            <Setter Property="Background" Value="#3B82F6"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="24,14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" 
                                Background="{TemplateBinding Background}" 
                                CornerRadius="8">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"
                                              Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#2563EB"/>
                                <Setter TargetName="border" Property="Effect">
                                    <Setter.Value>
                                        <DropShadowEffect BlurRadius="8" ShadowDepth="2" Color="#103B82F6" Opacity="0.3"/>
                                    </Setter.Value>
                                </Setter>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#1D4ED8"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Category Header Style -->
        <Style x:Key="CategoryHeader" TargetType="TextBlock">
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="#2D3748"/>
            <Setter Property="Margin" Value="0,8,0,4"/>
        </Style>

        <!-- Sidebar Menu Button Style -->
        <Style x:Key="SidebarButton" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="FontWeight" Value="Normal"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="HorizontalAlignment" Value="Stretch"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Margin" Value="0,1"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" 
                                Background="{TemplateBinding Background}" 
                                CornerRadius="6"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" 
                                              VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#F1F5F9"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#E2E8F0"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Active Sidebar Button Style -->
        <Style x:Key="SidebarButtonActive" TargetType="Button" BasedOn="{StaticResource SidebarButton}">
            <Setter Property="Background" Value="#EBF4FF"/>
            <Setter Property="Foreground" Value="#1E40AF"/>
            <Setter Property="FontWeight" Value="Medium"/>
        </Style>

        <!-- Expandable Section Style -->
        <Style x:Key="ExpanderStyle" TargetType="Expander">
            <Setter Property="IsExpanded" Value="True"/>
            <Setter Property="Margin" Value="0,4,0,8"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Expander">
                        <Border Background="#FAFBFC" CornerRadius="8" Margin="0,2">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <ToggleButton x:Name="HeaderSite" 
                                              Grid.Row="0" 
                                              IsChecked="{Binding IsExpanded, RelativeSource={RelativeSource TemplatedParent}}"
                                              Background="Transparent"
                                              BorderThickness="0"
                                              Padding="12,10"
                                              HorizontalAlignment="Stretch"
                                              HorizontalContentAlignment="Left">
                                    <StackPanel Orientation="Horizontal">
                                        <Path x:Name="arrow" 
                                              Fill="#6B7280" 
                                              Margin="0,0,8,0"
                                              VerticalAlignment="Center"
                                              Data="M4,6 L8,10 L4,14" 
                                              Stroke="#6B7280" 
                                              StrokeThickness="1.2"
                                              Width="12"
                                              Height="12">
                                            <Path.RenderTransform>
                                                <RotateTransform x:Name="arrowTransform" Angle="0" CenterX="6" CenterY="10"/>
                                            </Path.RenderTransform>
                                        </Path>
                                        <ContentPresenter Content="{TemplateBinding Header}" 
                                                          TextBlock.FontWeight="SemiBold" 
                                                          TextBlock.FontSize="13"
                                                          TextBlock.Foreground="#374151"/>
                                    </StackPanel>
                                </ToggleButton>
                                <ContentPresenter x:Name="ExpandSite" 
                                                  Grid.Row="1" 
                                                  Visibility="Collapsed" 
                                                  Margin="8,0,8,8"/>
                            </Grid>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsExpanded" Value="True">
                                <Setter TargetName="ExpandSite" Property="Visibility" Value="Visible"/>
                                <Trigger.EnterActions>
                                    <BeginStoryboard>
                                        <Storyboard>
                                            <DoubleAnimation Storyboard.TargetName="arrowTransform" 
                                                             Storyboard.TargetProperty="Angle" 
                                                             To="90" Duration="0:0:0.2"/>
                                        </Storyboard>
                                    </BeginStoryboard>
                                </Trigger.EnterActions>
                                <Trigger.ExitActions>
                                    <BeginStoryboard>
                                        <Storyboard>
                                            <DoubleAnimation Storyboard.TargetName="arrowTransform" 
                                                             Storyboard.TargetProperty="Angle" 
                                                             To="0" Duration="0:0:0.2"/>
                                        </Storyboard>
                                    </BeginStoryboard>
                                </Trigger.ExitActions>
                            </Trigger>
                            <Trigger SourceName="HeaderSite" Property="IsMouseOver" Value="True">
                                <Setter TargetName="HeaderSite" Property="Background" Value="#F1F5F9"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="70"/>
            <!-- Modern Header -->
            <RowDefinition Height="*"/>
            <!-- Content Area -->
            <RowDefinition Height="60"/>
            <!-- Enhanced Footer -->
        </Grid.RowDefinitions>

        <!-- Modern Header (Grid.Row="0") -->
        <Border Grid.Row="0" Background="White" BorderThickness="0,0,0,1" BorderBrush="#E1E5E9">
            <Grid Margin="24,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- App Title -->
                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <Ellipse Width="36" Height="36" Fill="#0078D7" Margin="0,0,12,0">
                        <Ellipse.Effect>
                            <DropShadowEffect BlurRadius="6" ShadowDepth="1" Color="#0D000000" Opacity="0.15"/>
                        </Ellipse.Effect>
                    </Ellipse>
                    <StackPanel VerticalAlignment="Center">
                        <TextBlock Text="easyADReport" FontSize="20" FontWeight="SemiBold" Foreground="#1A202C"/>
                        <TextBlock Text="easy Active Directory Reporting" FontSize="11" Foreground="#718096" Margin="0,-2,0,0"/>
                    </StackPanel>
                </StackPanel>

                <!-- Statistics Card -->
                <Border Grid.Column="2" Style="{StaticResource ModernCard}" Padding="16,8" MinWidth="140">
                    <StackPanel>
                        <TextBlock Text="Total Results" FontSize="11" Foreground="#718096" HorizontalAlignment="Center"/>
                        <TextBlock x:Name="TotalResultCountText" Text="0" FontSize="24" FontWeight="Bold" 
                                   Foreground="#0078D7" HorizontalAlignment="Center" Margin="0,2,0,0"/>
                    </StackPanel>
                </Border>
            </Grid>
        </Border>

        <!-- Content Area (Grid.Row="1") -->
        <Grid Grid.Row="1" Margin="24,20,24,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="320"/>
                <!-- Enhanced Sidebar - erweitert für bessere Lesbarkeit -->
                <ColumnDefinition Width="24"/>
                <!-- Spacing - vergrößert -->
                <ColumnDefinition Width="*"/>
                <!-- Main Content -->
            </Grid.ColumnDefinitions>

            <!-- Modern Sidebar with Categorized Reports -->
            <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Grid.ColumnSpan="2" Margin="0,0,10,0">
                <Border Style="{StaticResource ModernCard}" Padding="16,20">
                    <StackPanel>
                        <!-- Sidebar Header -->
                        <Grid Margin="0,0,0,20">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <TextBlock Text="📊 Quick Reports" Style="{StaticResource CategoryHeader}" 
                                       FontSize="16" FontWeight="Bold" Grid.Row="0" Margin="0,0,0,8"/>
                            <TextBlock Text="Select a predefined report to execute instantly" 
                                       FontSize="12" Foreground="#6B7280" Grid.Row="1" TextWrapping="Wrap"/>
                        </Grid>

                        <!-- Users - General Overview -->
                        <Expander Header="👤 Users - Overview" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickAllUsers" Content="All Users" Style="{StaticResource SidebarButton}" ToolTip="Retrieve all user accounts in the domain"/>
                                <Button x:Name="ButtonQuickDisabledUsers" Content="Disabled Users" Style="{StaticResource SidebarButton}" ToolTip="Find all disabled user accounts"/>
                                <Button x:Name="ButtonQuickLockedUsers" Content="Locked Users" Style="{StaticResource SidebarButton}" ToolTip="Show accounts currently locked out"/>
                                <Button x:Name="ButtonQuickNeverLoggedOn" Content="Never Logged On" Style="{StaticResource SidebarButton}" ToolTip="Users who have never logged into the system"/>
                                <Button x:Name="ButtonQuickAdminUsers" Content="Administrators" Style="{StaticResource SidebarButton}" ToolTip="Show all administrative accounts"/>
                                <Button x:Name="ButtonQuickGuestAccountStatus" Content="Guest Account Status" Style="{StaticResource SidebarButton}" ToolTip="Check status of guest accounts"/>
                            </StackPanel>
                        </Expander>

                        <!-- Users - Activity and Time -->
                        <Expander Header="📅 Users - Activity" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickInactiveUsers" Content="Inactive Users (90+ days)" Style="{StaticResource SidebarButton}" ToolTip="Users inactive for more than 90 days"/>
                                <Button x:Name="ButtonQuickRecentlyCreatedUsers" Content="Recently Created (30d)" Style="{StaticResource SidebarButton}" ToolTip="New user accounts created within 30 days"/>
                                <Button x:Name="ButtonQuickRecentlyDeletedUsers" Content="Recently Deleted" Style="{StaticResource SidebarButton}" ToolTip="Recently deleted user accounts"/>
                                <Button x:Name="ButtonQuickRecentlyModifiedUsers" Content="Recently Modified" Style="{StaticResource SidebarButton}" ToolTip="User accounts with recent changes"/>
                                <Button x:Name="ButtonQuickExpiringAccounts" Content="Expiring Accounts" Style="{StaticResource SidebarButton}" ToolTip="Accounts approaching expiration date"/>
                            </StackPanel>
                        </Expander>

                        <!-- Users - Password and Security -->
                        <Expander Header="🔐 Users - Passwords" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickNeverExpire" Content="Password Never Expires" Style="{StaticResource SidebarButton}" ToolTip="Users with non-expiring passwords"/>
                                <Button x:Name="ButtonQuickPasswordExpiringSoon" Content="Password Expiring (7d)" Style="{StaticResource SidebarButton}" ToolTip="Passwords expiring within 7 days"/>
                                <Button x:Name="ButtonQuickExpiredPasswords" Content="Expired Passwords" Style="{StaticResource SidebarButton}" ToolTip="Users with expired passwords"/>
                                <Button x:Name="ButtonQuickStalePasswords" Content="Stale Passwords" Style="{StaticResource SidebarButton}" ToolTip="Users with very old passwords"/>
                                <Button x:Name="ButtonQuickNeverChangingPasswords" Content="Never Changing Passwords" Style="{StaticResource SidebarButton}" ToolTip="Users who never changed their password"/>
                                <Button x:Name="ButtonQuickReversibleEncryption" Content="Reversible Encryption" Style="{StaticResource SidebarButton}" ToolTip="Users with reversible password encryption"/>
                            </StackPanel>
                        </Expander>

                        <!-- Users - Organization and Attributes -->
                        <Expander Header="🏢 Users - Organization" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickUsersByDepartment" Content="Users by Department" Style="{StaticResource SidebarButton}" ToolTip="Group users by their department"/>
                                <Button x:Name="ButtonQuickUsersByManager" Content="Users by Manager" Style="{StaticResource SidebarButton}" ToolTip="Show users organized by manager"/>
                                <Button x:Name="ButtonQuickUsersWithoutManager" Content="Users without Manager" Style="{StaticResource SidebarButton}" ToolTip="Users with no assigned manager"/>
                                <Button x:Name="ButtonQuickUsersMissingRequiredAttributes" Content="Missing Required Attributes" Style="{StaticResource SidebarButton}" ToolTip="Users with incomplete profile information"/>
                            </StackPanel>
                        </Expander>

                        <!-- Users - Advanced Security -->
                        <Expander Header="🛡️ Users - Advanced Security" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickKerberosDES" Content="Kerberos DES Users" Style="{StaticResource SidebarButton}" ToolTip="Users using weak DES encryption"/>
                                <Button x:Name="ButtonQuickUsersWithSPN" Content="Users with SPN" Style="{StaticResource SidebarButton}" ToolTip="User accounts with Service Principal Names"/>
                                <Button x:Name="ButtonQuickUsersDuplicateLogonNames" Content="Duplicate Logon Names" Style="{StaticResource SidebarButton}" ToolTip="Find duplicate user logon names"/>
                                <Button x:Name="ButtonQuickOrphanedSIDsUsers" Content="Orphaned SIDs" Style="{StaticResource SidebarButton}" ToolTip="User accounts with orphaned SIDs"/>
                            </StackPanel>
                        </Expander>

                        <!-- Groups - Overview -->
                        <Expander Header="👥 Groups - Overview" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickGroups" Content="All Groups" Style="{StaticResource SidebarButton}" ToolTip="Retrieve all groups in the domain"/>
                                <Button x:Name="ButtonQuickSecurityGroups" Content="Security Groups" Style="{StaticResource SidebarButton}" ToolTip="Show only security groups"/>
                                <Button x:Name="ButtonQuickDistributionGroups" Content="Distribution Lists" Style="{StaticResource SidebarButton}" ToolTip="Show only distribution groups"/>
                                <Button x:Name="ButtonQuickGroupsByTypeScope" Content="Groups by Type/Scope" Style="{StaticResource SidebarButton}" ToolTip="Categorize groups by type and scope"/>
                                <Button x:Name="ButtonQuickMailEnabledGroups" Content="Mail Enabled Groups" Style="{StaticResource SidebarButton}" ToolTip="Groups enabled for email"/>
                                <Button x:Name="ButtonQuickDynamicDistGroups" Content="Dynamic Distribution Groups" Style="{StaticResource SidebarButton}" ToolTip="Dynamic distribution groups"/>
                            </StackPanel>
                        </Expander>

                        <!-- Groups - Structure and Issues -->
                        <Expander Header="⚠️ Groups - Structure Issues" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickEmptyGroups" Content="Empty Groups" Style="{StaticResource SidebarButton}" ToolTip="Groups with no members"/>
                                <Button x:Name="ButtonQuickNestedGroups" Content="Nested Groups" Style="{StaticResource SidebarButton}" ToolTip="Groups containing other groups"/>
                                <Button x:Name="ButtonQuickCircularGroups" Content="Circular References" Style="{StaticResource SidebarButton}" ToolTip="Groups with circular membership"/>
                                <Button x:Name="ButtonQuickGroupsWithoutOwners" Content="Groups without Owners" Style="{StaticResource SidebarButton}" ToolTip="Groups with no assigned owners"/>
                                <Button x:Name="ButtonQuickLargeGroups" Content="Large Groups" Style="{StaticResource SidebarButton}" ToolTip="Groups with many members"/>
                                <Button x:Name="ButtonQuickRecentlyModifiedGroups" Content="Recently Modified Groups" Style="{StaticResource SidebarButton}" ToolTip="Groups with recent changes"/>
                            </StackPanel>
                        </Expander>

                        <!-- Computers - Overview -->
                        <Expander Header="💻 Computers - Overview" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickComputers" Content="All Computers" Style="{StaticResource SidebarButton}" ToolTip="Retrieve all computer accounts"/>
                                <Button x:Name="ButtonQuickInactiveComputers" Content="Inactive Computers (90+ days)" Style="{StaticResource SidebarButton}" ToolTip="Computers inactive for more than 90 days"/>
                                <Button x:Name="ButtonQuickComputersNeverLoggedOn" Content="Computers Never Logged On" Style="{StaticResource SidebarButton}" ToolTip="Computer accounts that never logged on"/>
                                <Button x:Name="ButtonQuickComputersByLocation" Content="Computers by Location" Style="{StaticResource SidebarButton}" ToolTip="Group computers by location"/>
                                <Button x:Name="ButtonQuickDuplicateComputerNames" Content="Duplicate Computer Names" Style="{StaticResource SidebarButton}" ToolTip="Find duplicate computer names"/>
                            </StackPanel>
                        </Expander>

                        <!-- Computers - Operating System and Security -->
                        <Expander Header="🖥️ Computers - OS Security" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickOSSummary" Content="Operating System Summary" Style="{StaticResource SidebarButton}" ToolTip="Summary of OS versions in use"/>
                                <Button x:Name="ButtonQuickComputersByOSVersion" Content="Computers by OS Version" Style="{StaticResource SidebarButton}" ToolTip="Group computers by OS version"/>
                                <Button x:Name="ButtonQuickBitLockerStatus" Content="BitLocker Status" Style="{StaticResource SidebarButton}" ToolTip="Check BitLocker encryption status"/>
                            </StackPanel>
                        </Expander>

                        <!-- Service Accounts -->
                        <Expander Header="⚙️ Service Accounts" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickServiceAccountsOverview" Content="Service Accounts Overview" Style="{StaticResource SidebarButton}" ToolTip="Overview of all service accounts"/>
                                <Button x:Name="ButtonQuickManagedServiceAccounts" Content="Managed Service Accounts" Style="{StaticResource SidebarButton}" ToolTip="Show managed service accounts (MSA)"/>
                                <Button x:Name="ButtonQuickServiceAccountsSPN" Content="Service Accounts with SPN" Style="{StaticResource SidebarButton}" ToolTip="Service accounts with Service Principal Names"/>
                                <Button x:Name="ButtonQuickHighPrivServiceAccounts" Content="High Privilege Service Accounts" Style="{StaticResource SidebarButton}" ToolTip="Service accounts with high privileges"/>
                                <Button x:Name="ButtonQuickServiceAccountPasswordAge" Content="Service Account Password Age" Style="{StaticResource SidebarButton}" ToolTip="Check service account password ages"/>
                                <Button x:Name="ButtonQuickUnusedServiceAccounts" Content="Unused Service Accounts" Style="{StaticResource SidebarButton}" ToolTip="Unused or obsolete service accounts"/>
                            </StackPanel>
                        </Expander>

                        <!-- Security Analysis -->
                        <Expander Header="🔍 Security Analysis" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickWeakPasswordPolicy" Content="Weak Password Policies" Style="{StaticResource SidebarButton}" ToolTip="Analyze weak password policies"/>
                                <Button x:Name="ButtonQuickRiskyGroupMemberships" Content="Risky Memberships" Style="{StaticResource SidebarButton}" ToolTip="Find risky group memberships"/>
                                <Button x:Name="ButtonQuickPrivilegedAccounts" Content="Privileged Accounts" Style="{StaticResource SidebarButton}" ToolTip="Show all privileged accounts"/>
                                <Button x:Name="ButtonQuickKerberoastable" Content="Kerberoastable Accounts" Style="{StaticResource SidebarButton}" ToolTip="Accounts vulnerable to Kerberoasting"/>
                                <Button x:Name="ButtonQuickASREPRoastable" Content="ASREPRoastable Accounts" Style="{StaticResource SidebarButton}" ToolTip="Accounts vulnerable to ASREPRoasting"/>
                                <Button x:Name="ButtonQuickHoneyTokens" Content="Honey Tokens" Style="{StaticResource SidebarButton}" ToolTip="Identify honey token accounts"/>
                                <Button x:Name="ButtonQuickPrivilegeEscalation" Content="Privilege Escalation Paths" Style="{StaticResource SidebarButton}" ToolTip="Find privilege escalation paths"/>
                                <Button x:Name="ButtonQuickExposedCredentials" Content="Exposed Credentials" Style="{StaticResource SidebarButton}" ToolTip="Check for exposed credentials"/>
                                <Button x:Name="ButtonQuickSuspiciousLogons" Content="Suspicious Logons" Style="{StaticResource SidebarButton}" ToolTip="Detect suspicious logon patterns"/>
                                <Button x:Name="ButtonQuickForeignSecurityPrincipals" Content="Foreign Security Principals" Style="{StaticResource SidebarButton}" ToolTip="Find foreign security principals"/>
                                <Button x:Name="ButtonQuickSIDHistoryAbuse" Content="SID History Abuse" Style="{StaticResource SidebarButton}" ToolTip="Detect SID history abuse"/>
                            </StackPanel>
                        </Expander>

                        <!-- Permissions and ACL -->
                        <Expander Header="🔑 Permissions and ACL" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickACLAnalysis" Content="ACL Analysis" Style="{StaticResource SidebarButton}" ToolTip="Analyze Access Control Lists"/>
                                <Button x:Name="ButtonQuickInheritanceBreaks" Content="Inheritance Breaks" Style="{StaticResource SidebarButton}" ToolTip="Find ACL inheritance breaks"/>
                                <Button x:Name="ButtonQuickAdminSDHolderObjects" Content="AdminSDHolder Objects" Style="{StaticResource SidebarButton}" ToolTip="Show AdminSDHolder protected objects"/>
                                <Button x:Name="ButtonQuickAdvancedDelegation" Content="Advanced Delegation" Style="{StaticResource SidebarButton}" ToolTip="Advanced delegation analysis"/>
                                <Button x:Name="ButtonQuickDelegation" Content="Delegation Analysis" Style="{StaticResource SidebarButton}" ToolTip="General delegation analysis"/>
                                <Button x:Name="ButtonQuickSchemaPermissions" Content="Schema Permissions" Style="{StaticResource SidebarButton}" ToolTip="Check schema permissions"/>
                                <Button x:Name="ButtonQuickDCSyncRights" Content="DCSync Rights Analysis" Style="{StaticResource SidebarButton}" ToolTip="Analyze DCSync rights"/>
                                <Button x:Name="ButtonQuickSchemaAdmins" Content="Schema Admin Paths" Style="{StaticResource SidebarButton}" ToolTip="Show schema admin privilege paths"/>
                            </StackPanel>
                        </Expander>

                        <!-- Policies and GPOs -->
                        <Expander Header="📋 GPOs and Policies" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickGPOOverview" Content="GPO Overview" Style="{StaticResource SidebarButton}" ToolTip="General Group Policy Objects overview"/>
                                <Button x:Name="ButtonQuickUnlinkedGPOs" Content="Unlinked GPOs" Style="{StaticResource SidebarButton}" ToolTip="Find unlinked Group Policy Objects"/>
                                <Button x:Name="ButtonQuickEmptyGPOs" Content="Empty GPOs" Style="{StaticResource SidebarButton}" ToolTip="Show empty Group Policy Objects"/>
                                <Button x:Name="ButtonQuickGPOPermissions" Content="GPO Permissions" Style="{StaticResource SidebarButton}" ToolTip="Analyze GPO permissions"/>
                                <Button x:Name="ButtonQuickPasswordPolicySummary" Content="Password Policy Summary" Style="{StaticResource SidebarButton}" ToolTip="Summary of password policies"/>
                                <Button x:Name="ButtonQuickAccountLockoutPolicies" Content="Account Lockout Policies" Style="{StaticResource SidebarButton}" ToolTip="Review account lockout policies"/>
                                <Button x:Name="ButtonQuickFineGrainedPasswordPolicies" Content="Fine-Grained Password Policies" Style="{StaticResource SidebarButton}" ToolTip="Show fine-grained password policies"/>
                            </StackPanel>
                        </Expander>

                        <!-- AD Infrastructure and Health -->
                        <Expander Header="🏥 AD Infrastructure" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickFSMORoles" Content="FSMO Role Holders" Style="{StaticResource SidebarButton}" ToolTip="Show FSMO role holders"/>
                                <Button x:Name="ButtonQuickDCStatus" Content="Domain Controller Status" Style="{StaticResource SidebarButton}" ToolTip="Check domain controller status"/>
                                <Button x:Name="ButtonQuickReplicationStatus" Content="Replication Status" Style="{StaticResource SidebarButton}" ToolTip="Check replication status"/>
                                <Button x:Name="ButtonQuickSYSVOLHealth" Content="SYSVOL Health Check" Style="{StaticResource SidebarButton}" ToolTip="Analyze SYSVOL health"/>
                                <Button x:Name="ButtonQuickDNSHealth" Content="DNS Health Analysis" Style="{StaticResource SidebarButton}" ToolTip="Check DNS health and configuration"/>
                                <Button x:Name="ButtonQuickBackupStatus" Content="Backup Readiness" Style="{StaticResource SidebarButton}" ToolTip="Check backup readiness status"/>
                                <Button x:Name="ButtonQuickOUHierarchy" Content="OU Hierarchy" Style="{StaticResource SidebarButton}" ToolTip="Show Organizational Unit hierarchy"/>
                                <Button x:Name="ButtonQuickSitesSubnets" Content="Sites and Subnets" Style="{StaticResource SidebarButton}" ToolTip="Display sites and subnets"/>
                                <Button x:Name="ButtonQuickTrustRelationships" Content="Trust Relationships" Style="{StaticResource SidebarButton}" ToolTip="Show trust relationships"/>
                            </StackPanel>
                        </Expander>

                        <!-- Schema and Advanced Analysis -->
                        <Expander Header="🔬 Schema and Advanced Analysis" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickSchemaAnalysis" Content="Schema Extensions" Style="{StaticResource SidebarButton}" ToolTip="Analyze Active Directory schema extensions"/>
                                <Button x:Name="ButtonQuickCertificateAnalysis" Content="Certificate Security" Style="{StaticResource SidebarButton}" ToolTip="Analyze certificate security"/>
                                <Button x:Name="ButtonQuickQuotasLimits" Content="Quotas and Limits" Style="{StaticResource SidebarButton}" ToolTip="Check directory quotas and limits"/>
                            </StackPanel>
                        </Expander>

                        <!-- Statistics and Reports -->
                        <Expander Header="📈 Statistics and Reports" Style="{StaticResource ExpanderStyle}" IsExpanded="False">
                            <StackPanel>
                                <Button x:Name="ButtonQuickDepartmentStats" Content="Department Statistics" Style="{StaticResource SidebarButton}" ToolTip="Generate department statistics"/>
                                <Button x:Name="ButtonQuickDepartmentSecurity" Content="Department Security" Style="{StaticResource SidebarButton}" ToolTip="Analyze department security posture"/>
                            </StackPanel>
                        </Expander>
                    </StackPanel>
                </Border>
            </ScrollViewer>

            <!-- Main Content Area -->
            <Grid Grid.Column="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <!-- Filter Section -->
                    <RowDefinition Height="20"/>
                    <!-- Spacing -->
                    <RowDefinition Height="*"/>
                    <!-- Results -->
                </Grid.RowDefinitions>

                <!-- Enhanced Filter Section -->
                <Grid Grid.Row="0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                        <ColumnDefinition Width="25"/>
                        <ColumnDefinition Width="Auto" MinWidth="435"/>
                        <ColumnDefinition Width="20"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <!-- Advanced Filter Card -->
                    <Border Grid.Column="0" Style="{StaticResource ModernCard}" Padding="20,16">
                        <StackPanel>
                            <TextBlock Style="{StaticResource CategoryHeader}" Margin="0,0,0,12"><Run Language="de-de" Text="Search with "/><Run Text="Filter"/></TextBlock>

                            <!-- Object Type Selection -->
                            <StackPanel Orientation="Horizontal" Margin="0,0,0,12">
                                <TextBlock Text="Object Type:" VerticalAlignment="Center" Width="90" FontWeight="Medium"/>
                                <RadioButton x:Name="RadioButtonUser" Content="User" IsChecked="True" Margin="12,0" VerticalAlignment="Center"/>
                                <RadioButton x:Name="RadioButtonGroup" Content="Group" Margin="12,0" VerticalAlignment="Center"/>
                                <RadioButton x:Name="RadioButtonComputer" Content="Computer" Margin="12,0" VerticalAlignment="Center"/>
                                <RadioButton x:Name="RadioButtonGroupMemberships" Content="Memberships" Margin="12,0" VerticalAlignment="Center"/>
                            </StackPanel>

                            <!-- Filter 1 -->
                            <Grid Margin="0,0,0,8">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="80"/>
                                    <ColumnDefinition Width="140"/>
                                    <ColumnDefinition Width="100"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="Filter 1:" VerticalAlignment="Center" FontWeight="Medium"/>
                                <ComboBox x:Name="ComboBoxFilterAttribute1" Grid.Column="1" Margin="8,0" VerticalAlignment="Center"/>
                                <ComboBox x:Name="ComboBoxFilterOperator1" Grid.Column="2" Margin="8,0" VerticalAlignment="Center">
                                    <ComboBoxItem Content="Contains" IsSelected="True"/>
                                    <ComboBoxItem Content="Equals"/>
                                    <ComboBoxItem Content="StartsWith"/>
                                    <ComboBoxItem Content="EndsWith"/>
                                    <ComboBoxItem Content="NotEqual"/>
                                </ComboBox>
                                <TextBox x:Name="TextBoxFilterValue1" Grid.Column="3" Margin="8,0,0,0" VerticalAlignment="Center" Padding="8,6"/>
                            </Grid>

                            <!-- Logic Selector -->
                            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                                <TextBlock Text="Logic:" VerticalAlignment="Center" Width="80" FontWeight="Medium"/>
                                <RadioButton x:Name="RadioButtonAnd" Content="AND" IsChecked="True" Margin="8,0" VerticalAlignment="Center"/>
                                <RadioButton x:Name="RadioButtonOr" Content="OR" Margin="12,0" VerticalAlignment="Center"/>
                                <CheckBox x:Name="CheckBoxUseSecondFilter" Content="Use second filter" VerticalAlignment="Center" Margin="20,0"/>
                            </StackPanel>

                            <!-- Filter 2 -->
                            <Grid x:Name="SecondFilterPanel" IsEnabled="False">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="80"/>
                                    <ColumnDefinition Width="140"/>
                                    <ColumnDefinition Width="100"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="Filter 2:" VerticalAlignment="Center" FontWeight="Medium"/>
                                <ComboBox x:Name="ComboBoxFilterAttribute2" Grid.Column="1" Margin="8,0" VerticalAlignment="Center"/>
                                <ComboBox x:Name="ComboBoxFilterOperator2" Grid.Column="2" Margin="8,0" VerticalAlignment="Center">
                                    <ComboBoxItem Content="Contains" IsSelected="True"/>
                                    <ComboBoxItem Content="Equals"/>
                                    <ComboBoxItem Content="StartsWith"/>
                                    <ComboBoxItem Content="EndsWith"/>
                                    <ComboBoxItem Content="NotEqual"/>
                                </ComboBox>
                                <TextBox x:Name="TextBoxFilterValue2" Grid.Column="3" Margin="8,0,0,0" VerticalAlignment="Center" Padding="8,6"/>
                            </Grid>
                        </StackPanel>
                    </Border>

                    <!-- Attributes Selection Card -->
                    <Border Grid.Column="2" Style="{StaticResource ModernCard}" Padding="20,16">
                        <StackPanel>
                            <Grid Margin="0,0,0,12">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Text="Export Attributes" Style="{StaticResource CategoryHeader}" Grid.Column="0"/>
                                <StackPanel Grid.Column="1" Orientation="Horizontal">
                                    <Button x:Name="ButtonSelectAllAttributes" Content="All" Style="{StaticResource ModernButton}" 
                                            Margin="0,0,4,0" Padding="6,4" FontSize="10" ToolTip="Select all attributes"/>
                                    <Button x:Name="ButtonSelectNoneAttributes" Content="None" Style="{StaticResource ModernButton}" 
                                            Padding="6,4" FontSize="10" ToolTip="Deselect all attributes"/>
                                </StackPanel>
                            </Grid>

                            <!-- Attribute Categories -->
                            <TabControl x:Name="TabControlAttributes" Height="140" BorderThickness="0" Background="Transparent">
                                <TabItem Header="Basic" FontSize="11">
                                    <ListBox x:Name="ListBoxBasicAttributes" SelectionMode="Multiple" BorderThickness="0" Background="Transparent">
                                        <ListBoxItem Content="DisplayName" IsSelected="True"/>
                                        <ListBoxItem Content="SamAccountName" IsSelected="True"/>
                                        <ListBoxItem Content="GivenName"/>
                                        <ListBoxItem Content="Surname"/>
                                        <ListBoxItem Content="mail"/>
                                        <ListBoxItem Content="Department"/>
                                        <ListBoxItem Content="Title"/>
                                        <ListBoxItem Content="Enabled" IsSelected="True"/>
                                    </ListBox>
                                </TabItem>
                                <TabItem Header="Security" FontSize="11">
                                    <ListBox x:Name="ListBoxSecurityAttributes" SelectionMode="Multiple" BorderThickness="0" Background="Transparent">
                                        <ListBoxItem Content="LastLogonTimestamp"/>
                                        <ListBoxItem Content="PasswordExpired"/>
                                        <ListBoxItem Content="PasswordLastSet"/>
                                        <ListBoxItem Content="AccountExpirationDate"/>
                                        <ListBoxItem Content="badPwdCount"/>
                                        <ListBoxItem Content="lockoutTime"/>
                                        <ListBoxItem Content="UserAccountControl"/>
                                        <ListBoxItem Content="memberOf"/>
                                    </ListBox>
                                </TabItem>
                                <TabItem Header="Extended" FontSize="11">
                                    <ListBox x:Name="ListBoxExtendedAttributes" SelectionMode="Multiple" BorderThickness="0" Background="Transparent">
                                        <ListBoxItem Content="whenCreated"/>
                                        <ListBoxItem Content="whenChanged"/>
                                        <ListBoxItem Content="Manager"/>
                                        <ListBoxItem Content="Company"/>
                                        <ListBoxItem Content="physicalDeliveryOfficeName"/>
                                        <ListBoxItem Content="telephoneNumber"/>
                                        <ListBoxItem Content="homeDirectory"/>
                                        <ListBoxItem Content="scriptPath"/>
                                    </ListBox>
                                </TabItem>
                            </TabControl>
                        </StackPanel>
                    </Border>

                    <!-- Action Buttons -->
                    <StackPanel Grid.Column="4" VerticalAlignment="Center" MinWidth="180" Height="120">
                        <Button x:Name="ButtonQueryAD" Content="SEARCH" Style="{StaticResource PrimaryButton}" 
                                Height="55" FontSize="15" Margin="0,0,0,12" Width="175"/>
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                            <Button x:Name="ButtonExportCSV" Content="CSV" Style="{StaticResource ModernButton}" 
                                    Width="83" Height="45" Margin="0,0,8,0" Background="#FFCCD8FF"/>
                            <Button x:Name="ButtonExportHTML" Content="HTML" Style="{StaticResource ModernButton}" 
                                    Width="83" Height="45" Background="#FFCCD8FF"/>
                        </StackPanel>
                    </StackPanel>

                </Grid>

                <!-- Results Section -->
                <Border Grid.Row="2" Style="{StaticResource ModernCard}" Padding="20,16" Margin="0,0,0,10">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <!-- Results Header -->
                        <Grid Grid.Row="0" Margin="0,0,0,16">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Grid.Column="0" Text="Query Results" Style="{StaticResource CategoryHeader}"/>
                            <StackPanel Grid.Column="1" Orientation="Horizontal">
                                <Button x:Name="ButtonRefresh" Content="Reset Query Window" Style="{StaticResource ModernButton}" Margin="0,0,8,0" Padding="8,4"/>
                                <Button x:Name="ButtonCopy" Content="Copy Selected Rows" Style="{StaticResource ModernButton}" Padding="8,4"/>
                            </StackPanel>
                        </Grid>

                        <!-- Enhanced DataGrid -->
                        <DataGrid Grid.Row="1" x:Name="DataGridResults" 
                                  AutoGenerateColumns="True" 
                                  IsReadOnly="True" 
                                  BorderThickness="1" 
                                  BorderBrush="#E1E5E9"
                                  Background="White" 
                                  GridLinesVisibility="Horizontal" 
                                  RowBackground="White" 
                                  AlternatingRowBackground="#FAFBFC"
                                  HeadersVisibility="Column"
                                  CanUserSortColumns="True"
                                  CanUserReorderColumns="True"
                                  CanUserResizeColumns="True">
                            <DataGrid.ColumnHeaderStyle>
                                <Style TargetType="DataGridColumnHeader">
                                    <Setter Property="Background" Value="#F7FAFC"/>
                                    <Setter Property="Foreground" Value="#2D3748"/>
                                    <Setter Property="FontWeight" Value="SemiBold"/>
                                    <Setter Property="BorderBrush" Value="#E1E5E9"/>
                                    <Setter Property="BorderThickness" Value="0,0,1,1"/>
                                    <Setter Property="Padding" Value="12,8"/>
                                </Style>
                            </DataGrid.ColumnHeaderStyle>
                            <DataGrid.CellStyle>
                                <Style TargetType="DataGridCell">
                                    <Setter Property="Padding" Value="12,6"/>
                                    <Setter Property="BorderThickness" Value="0"/>
                                    <Style.Triggers>
                                        <Trigger Property="IsSelected" Value="True">
                                            <Setter Property="Background" Value="#E3F2FD"/>
                                            <Setter Property="Foreground" Value="#1565C0"/>
                                        </Trigger>
                                    </Style.Triggers>
                                </Style>
                            </DataGrid.CellStyle>
                        </DataGrid>
                    </Grid>
                </Border>
            </Grid>
        </Grid>

        <!-- Enhanced Footer (Grid.Row="2") -->
        <Border Grid.Row="2" Background="White" BorderThickness="0,1,0,0" BorderBrush="#E1E5E9">
            <Grid Margin="24,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Status and Info -->
                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <Ellipse x:Name="StatusIndicator" Width="8" Height="8" Fill="#10B981" Margin="0,0,8,0"/>
                    <TextBlock x:Name="TextBlockStatus" Text="Ready" FontWeight="Medium" Foreground="#374151" Margin="0,0,16,0"/>
                </StackPanel>

                <!-- Version Info -->
                <StackPanel Grid.Column="3" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="Last update: 01.07.2025" FontSize="12" Foreground="#6B7280"/>
                    <TextBlock Text="easyADReport v0.4.2" FontSize="12" Foreground="#9CA3AF" VerticalAlignment="Center"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Setze die Ausgabekodierung auf UTF-8, um Probleme mit Umlauten zu vermeiden
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)

# Assembly fÃ¼r WPF laden
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms # FÃ¼r SaveFileDialog

# --- Globale AD-Gruppennamen fÃ¼r Deutsch/Englisch KompatibilitÃ¤t ---
$Global:ADGroupNames = @{
    DomainAdmins = @("Domain Admins", "DomÃ¤nen-Admins")
    EnterpriseAdmins = @("Enterprise Admins", "Organisations-Admins")
    SchemaAdmins = @("Schema Admins", "Schema-Admins")
    Administrators = @("Administrators", "Administratoren")
    AccountOperators = @("Account Operators", "Konten-Operatoren")
    ServerOperators = @("Server Operators", "Server-Operatoren")
    BackupOperators = @("Backup Operators", "Sicherungs-Operatoren")
    PrintOperators = @("Print Operators", "Druck-Operatoren")
    Replicator = @("Replicator", "Replikations-Operator")
    RemoteDesktopUsers = @("Remote Desktop Users", "Remotedesktopbenutzer")
    PowerUsers = @("Power Users", "Hauptbenutzer")
    DomainControllers = @("Domain Controllers", "DomÃ¤nencontroller")
    EnterpriseDomainControllers = @("Enterprise Domain Controllers", "Organisations-DomÃ¤nencontroller")
}

# Hilfsfunktion zum Finden von AD-Gruppen in beiden Sprachen
Function Get-ADGroupByNames {
    param(
        [string[]]$GroupNames,
        [switch]$ReturnAll
    )
    
    $foundGroups = @()
    foreach ($name in $GroupNames) {
        try {
            $group = Get-ADGroup -Filter "Name -eq '$name'" -ErrorAction SilentlyContinue
            if ($group) {
                if ($ReturnAll) {
                    $foundGroups += $group
                } else {
                    return $group
                }
            }
        } catch {
            # Ignore errors for non-existent groups
        }
    }
    
    if ($ReturnAll) {
        return $foundGroups
    }
    return $null
}

# --- Log-Funktion fÃ¼r konsistente Fehlerausgabe ---
Function Write-ADReportLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Type = 'Info',
        
        [Parameter(Mandatory=$false)]
        [switch]$GUI,
        
        [Parameter(Mandatory=$false)]
        [switch]$Terminal
    )
    
    # StandardmÃ¤ÃŸig sowohl GUI als auch Terminal, wenn nicht explizit angegeben
    if (-not $GUI -and -not $Terminal) {
        $GUI = $true
        $Terminal = $true
    }
    
    # Ausgabe in der GUI
    if ($GUI -and $Global:TextBlockStatus) {
        $Global:TextBlockStatus.Text = $Message
    }
    
    # Ausgabe im Terminal
    if ($Terminal) {
        switch ($Type) {
            'Info'    { Write-Host $Message }
            'Warning' { Write-Warning $Message }
            'Error'   { Write-Error $Message }
        }
    }
}

# --- Debug-Log-Funktion fÃ¼r konsistente Debug-Ausgabe ---
Function Write-DebugLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [string]$Category = 'Debug'
    )
    
    # Debug-Ausgabe nur wenn $DebugPreference gesetzt ist
    if ($DebugPreference -ne 'SilentlyContinue') {
        Write-Debug "[$Category] $Message"
    }
}

# --- Funktion zum Abrufen von AD-Daten ---
Function Get-ADReportData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [string]$FilterAttribute,
        [Parameter(Mandatory=$false)]
        [string]$FilterValue,
        [Parameter(Mandatory=$false)]
        [string]$FilterOperator = "Contains",
        [Parameter(Mandatory=$false)]
        [string]$FilterAttribute2,
        [Parameter(Mandatory=$false)]
        [string]$FilterValue2,
        [Parameter(Mandatory=$false)]
        [string]$FilterOperator2 = "Contains",
        [Parameter(Mandatory=$false)]
        [string]$FilterLogic = "AND",
        [Parameter(Mandatory=$true)]
        [System.Collections.IList]$SelectedAttributes,
        [Parameter(Mandatory=$false)]
        [string]$CustomFilter,
        [Parameter(Mandatory=$false)]
        [string]$ObjectType = "User"
    )

    # ÃœberprÃ¼fen, ob das AD-Modul verfÃ¼gbar ist
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-ADReportLog -Message "Error: Active Directory module not found." -Type Error
        return $null
    }

    try {
        # Konvertiere SelectedAttributes zu String-Array
        $PropertiesToLoad = @(
            $SelectedAttributes | ForEach-Object {
                if ($null -ne $_) {
                    if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                        $_.Content.ToString()
                    } else {
                        $_.ToString()
                    }
                }
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        # Basis-Eigenschaften hinzufÃ¼gen
        if ('DistinguishedName' -notin $PropertiesToLoad) { $PropertiesToLoad += 'DistinguishedName' }
        if ('ObjectClass' -notin $PropertiesToLoad) { $PropertiesToLoad += 'ObjectClass' }
        $PropertiesToLoad = $PropertiesToLoad | Select-Object -Unique

        if ($PropertiesToLoad.Count -eq 0) {
            $PropertiesToLoad = @("DisplayName", "SamAccountName", "ObjectClass")
        }

        # Filter erstellen
        $Filter = "*"
        
        if ($CustomFilter) {
            $Filter = $CustomFilter
        } else {
            # Erstelle Filter basierend auf den Eingaben
            $FilterPart1 = ""
            $FilterPart2 = ""
            
            # Erster Filter
            if (-not [string]::IsNullOrWhiteSpace($FilterValue) -and -not [string]::IsNullOrWhiteSpace($FilterAttribute)) {
                $FilterPart1 = Build-FilterString -Attribute $FilterAttribute -Value $FilterValue -Operator $FilterOperator
            } elseif (-not [string]::IsNullOrWhiteSpace($FilterValue) -and [string]::IsNullOrWhiteSpace($FilterAttribute)) {
                $FilterPart1 = Build-FilterString -Attribute "DisplayName" -Value $FilterValue -Operator "Contains"
            }
            
            # Zweiter Filter (wenn aktiviert)
            if (-not [string]::IsNullOrWhiteSpace($FilterValue2) -and -not [string]::IsNullOrWhiteSpace($FilterAttribute2)) {
                $FilterPart2 = Build-FilterString -Attribute $FilterAttribute2 -Value $FilterValue2 -Operator $FilterOperator2
            }
            
            # Kombiniere Filter mit UND/ODER Logik
            if ($FilterPart1 -and $FilterPart2) {
                if ($FilterLogic -eq "AND") {
                    $Filter = "$FilterPart1 -and $FilterPart2"
                } else {
                    $Filter = "$FilterPart1 -or $FilterPart2"
                }
            } elseif ($FilterPart1) {
                $Filter = $FilterPart1
            } elseif ($FilterPart2) {
                $Filter = $FilterPart2
            }
        }

        Write-ADReportLog -Message "Executing AD query with filter: $Filter" -Type Info -Terminal

        # AD-Abfrage basierend auf Objekttyp
        $Results = $null
        switch ($ObjectType) {
            "User" {
                if ($Filter -and $Filter.Trim() -eq "LockedOut -eq `$true") {
                    $LockedOutAccounts = Search-ADAccount -LockedOut -UsersOnly -ErrorAction Stop
                    if ($LockedOutAccounts) {
                        $Results = foreach ($Account in $LockedOutAccounts) {
                            try {
                                Get-ADUser -Identity $Account.SamAccountName -Properties $PropertiesToLoad -ErrorAction SilentlyContinue
                            } catch {
                                Write-Warning "Could not get details for user $($Account.SamAccountName): $($_.Exception.Message)"
                                $null
                            }
                        }
                        $Results = $Results | Where-Object {$_ -ne $null}
                    }
                } else {
                    # Debug-Ausgabe fÃ¼r Filter
                    Write-DebugLog "Executing Get-ADUser with filter: $Filter" -Category "ADQuery"
                    $Results = @(Get-ADUser -Filter $Filter -Properties $PropertiesToLoad -ErrorAction Stop)
                }
            }
            "Group" {
                $Results = @(Get-ADGroup -Filter $Filter -Properties $PropertiesToLoad -ErrorAction Stop)
            }
            "Computer" {
                $Results = @(Get-ADComputer -Filter $Filter -Properties $PropertiesToLoad -ErrorAction Stop)
            }
        }

        if ($Results) {
            $Results = $Results | Select-Object -Property $PropertiesToLoad -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
            return $Results
        } else {
            Write-ADReportLog -Message "No objects found for the specified filter." -Type Warning
            return $null
        }
    } catch {
        $ErrorMessage = "Error in AD query: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# Hilfsfunktion zum Erstellen von Filter-Strings
Function Build-FilterString {
    param(
        [string]$Attribute,
        [string]$Value,
        [string]$Operator
    )
    
    switch ($Operator) {
        "Contains" { return "$Attribute -like '*$Value*'" }
        "Equals" { return "$Attribute -eq '$Value'" }
        "StartsWith" { return "$Attribute -like '$Value*'" }
        "EndsWith" { return "$Attribute -like '*$Value'" }
        "NotEqual" { return "$Attribute -ne '$Value'" }
        default { return "$Attribute -like '*$Value*'" }
    }
}

# --- Funktion zum Abrufen der Gruppenmitgliedschaften eines Benutzers ---
Function Get-UserGroupMemberships {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$SamAccountName
    )

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-ADReportLog -Message "Fehler: Active Directory Modul nicht gefunden." -Type Error
        return $null
    }

    try {
        $User = Get-ADUser -Identity $SamAccountName -Properties SamAccountName, Name -ErrorAction Stop | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount' # HinzugefÃ¼gt: Name fÃ¼r UserDisplayName
        if (-not $User) {
            Write-ADReportLog -Message "Benutzer $SamAccountName nicht gefunden." -Type Warning
            return $null
        }
        
        $Groups = Get-ADPrincipalGroupMembership -Identity $User -ErrorAction Stop | 
                  Get-ADGroup -Properties Name, SamAccountName, Description, GroupCategory, GroupScope -ErrorAction SilentlyContinue # HinzugefÃ¼gt: SamAccountName fÃ¼r GroupSamAccountName

        if ($Groups) {
            $GroupMemberships = $Groups | ForEach-Object {
                [PSCustomObject]@{
                    UserDisplayName = $User.Name
                    UserSamAccountName = $User.SamAccountName
                    GroupName = $_.Name
                    GroupSamAccountName = $_.SamAccountName
                    GroupDescription = $_.Description
                    GroupCategory = $_.GroupCategory
                    GroupScope = $_.GroupScope
                }
            }
            return $GroupMemberships
        } else {
            Write-ADReportLog -Message "Keine Gruppenmitgliedschaften fÃ¼r Benutzer $SamAccountName gefunden." -Type Info
            return [System.Collections.ArrayList]::new() # Leeres Array zurÃ¼ckgeben, um Fehler zu vermeiden
        }
    } catch {
        $ErrorMessage = "Fehler beim Abrufen der Gruppenmitgliedschaften fuer $($SamAccountName): $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Funktion zum Abrufen von AD-Gruppendaten ---
Function Get-ADGroupReportData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [System.Collections.IList]$SelectedAttributes,
        [Parameter(Mandatory=$false)]
        [string]$CustomFilter
    )

    # ÃœberprÃ¼fen, ob das AD-Modul verfÃ¼gbar ist
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-ADReportLog -Message "Fehler: Active Directory Modul nicht gefunden." -Type Error
        return $null
    }

    try {
        # Konvertiere SelectedAttributes zu String-Array und filtere leere/null Werte
        $PropertiesToLoad = @(
            $SelectedAttributes | ForEach-Object {
                if ($null -ne $_) {
                    # Wenn es sich um ListBoxItem handelt, Content verwenden
                    if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                        $_.Content.ToString()
                    } else {
                        $_.ToString()
                    }
                }
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        if ('ObjectClass' -notin $PropertiesToLoad) { # Wichtig fÃ¼r Visualisierung und Typbestimmung
            $PropertiesToLoad += 'ObjectClass'
        }

        if ($PropertiesToLoad.Count -eq 0) {
            $PropertiesToLoad = @("Name", "SamAccountName", "GroupCategory", "GroupScope", "ObjectClass") 
        }

        $FilterToUse = "*"
        if ($CustomFilter) {
            $FilterToUse = $CustomFilter
        }
        
        Write-Host "FÃ¼hre Get-ADGroup mit Filter '$FilterToUse' und Eigenschaften '$($PropertiesToLoad -join ', ')' aus"
        $Groups = Get-ADGroup -Filter $FilterToUse -Properties $PropertiesToLoad -ErrorAction Stop

        if ($Groups) {
            # Erstelle Array mit den bereinigten Attributnamen fÃ¼r Select-Object
            $SelectAttributes = @(
                $SelectedAttributes | ForEach-Object {
                    if ($null -ne $_) {
                        if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                            $_.Content.ToString()
                        } else {
                            $_.ToString()
                        }
                    }
                } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
            
            # Verwende Select-Object, um ein Array von Objekten zu erstellen
            # Dies stellt sicher, dass wir eine IEnumerable-Sammlung zurÃ¼ckgeben
            $Output = $Groups | Select-Object $SelectAttributes -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
            return $Output
        } else {
            Write-ADReportLog -Message "No groups found for the specified filter." -Type Warning
            return $null
        }
    } catch {
        $ErrorMessage = "Error in AD group query: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Funktion zum Abrufen von AD-Computerdaten ---
Function Get-ADComputerReportData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [System.Collections.IList]$SelectedAttributes,
        [Parameter(Mandatory=$false)]
        [string]$CustomFilter
    )

    # ÃœberprÃ¼fen, ob das AD-Modul verfÃ¼gbar ist
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-ADReportLog -Message "Fehler: Active Directory Modul nicht gefunden." -Type Error
        return $null
    }

    try {
        # Konvertiere SelectedAttributes zu String-Array und filtere leere/null Werte
        $PropertiesToLoad = @(
            $SelectedAttributes | ForEach-Object {
                if ($null -ne $_) {
                    # Wenn es sich um ListBoxItem handelt, Content verwenden
                    if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                        $_.Content.ToString()
                    } else {
                        $_.ToString()
                    }
                }
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        if ('ObjectClass' -notin $PropertiesToLoad) { # Wichtig fÃ¼r Visualisierung und Typbestimmung
            $PropertiesToLoad += 'ObjectClass'
        }

        if ($PropertiesToLoad.Count -eq 0) {
            $PropertiesToLoad = @("Name", "DNSHostName", "OperatingSystem", "Enabled", "ObjectClass") 
        }

        $FilterToUse = "*"
        if ($CustomFilter) {
            $FilterToUse = $CustomFilter
        }
        
        Write-Host "FÃ¼hre Get-ADComputer mit Filter '$FilterToUse' und Eigenschaften '$($PropertiesToLoad -join ', ')' aus"
        $Computers = Get-ADComputer -Filter $FilterToUse -Properties $PropertiesToLoad -ErrorAction Stop

        if ($Computers) {
            # Erstelle Array mit den bereinigten Attributnamen fÃ¼r Select-Object
            $SelectAttributes = @(
                $SelectedAttributes | ForEach-Object {
                    if ($null -ne $_) {
                        if ($_ -is [System.Windows.Controls.ListBoxItem]) {
                            $_.Content.ToString()
                        } else {
                            $_.ToString()
                        }
                    }
                } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
            
            # Verwende Select-Object, um ein Array von Objekten zu erstellen
            $Output = $Computers | Select-Object $SelectAttributes -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
            return $Output
        } else {
            Write-ADReportLog -Message "No computers found for the specified filter." -Type Warning
            return $null
        }
    } catch {
        $ErrorMessage = "Error in AD computer query: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Funktion zum Abrufen von Gruppenmitgliedschaftsberichten ---
Function Get-ADGroupMembershipsReportData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$FilterAttribute,

        [Parameter(Mandatory=$true)]
        [string]$FilterValue
    )

    Write-ADReportLog -Message "Fetching group membership data for filter: $FilterAttribute = $FilterValue" -Type Info -Terminal
    $ReportOutput = @()

    try {
        $TargetObjects = Get-ADObject -Filter "$FilterAttribute -like '*$FilterValue*'" -Properties DisplayName, SamAccountName, MemberOf, Member, ObjectClass -ErrorAction SilentlyContinue

        if (-not $TargetObjects) {
            Write-ADReportLog -Message "No objects found for the specified filter '$FilterAttribute -like *$FilterValue*'." -Type Warning
            return $ReportOutput
        }

        # Gruppiere Objekte nach Typ für bessere Strukturierung
        $GroupedObjects = $TargetObjects | Group-Object {
            if ($_.ObjectClass -is [array]) { $_.ObjectClass[-1] } else { $_.ObjectClass }
        }

        foreach ($ObjectGroup in $GroupedObjects) {
            foreach ($TargetObject in $ObjectGroup.Group) {
                $objectClassSimple = if ($TargetObject.ObjectClass -is [array]) {
                    $TargetObject.ObjectClass[-1]
                } else {
                    $TargetObject.ObjectClass
                }
                
                $objDisplayName = if ($TargetObject.DisplayName) { $TargetObject.DisplayName } else { $TargetObject.Name }
                $objSamAccountName = if ($TargetObject.SamAccountName) { $TargetObject.SamAccountName } else { "N/A" }

                Write-ADReportLog -Message "Processing object: $objDisplayName (Type: $objectClassSimple)" -Type Info

                # Header-Eintrag für das aktuelle Objekt
                $ReportOutput += [PSCustomObject]@{
                    ObjectName = "$objDisplayName ($objSamAccountName)"
                    ObjectSAM = ""
                    ObjectType = $objectClassSimple.ToUpper()
                    Relationship = ""
                    RelatedObject = ""
                    RelatedObjectSAM = ""
                    RelatedObjectType = ""
                }

                if ($objectClassSimple -eq 'user' -or $objectClassSimple -eq 'computer') {
                    if ($TargetObject.MemberOf) {
                        foreach ($groupDN in $TargetObject.MemberOf) {
                            try {
                                $groupObject = Get-ADGroup -Identity $groupDN -Properties DisplayName, SamAccountName -ErrorAction Stop
                                $groupDisplayName = if ($groupObject.DisplayName) { $groupObject.DisplayName } else { $groupObject.Name }
                                $ReportOutput += [PSCustomObject]@{
                                    ObjectName = ""
                                    ObjectSAM = ""
                                    ObjectType = ""
                                    Relationship = "└─ IST MITGLIED VON"
                                    RelatedObject = $groupDisplayName
                                    RelatedObjectSAM = $groupObject.SamAccountName
                                    RelatedObjectType = "Group"
                                }
                            } catch {
                                Write-ADReportLog -Message "Error resolving group DN '$groupDN' for object '$objDisplayName': $($_.Exception.Message)" -Type Warning
                                $ReportOutput += [PSCustomObject]@{
                                    ObjectName = ""
                                    ObjectSAM = ""
                                    ObjectType = ""
                                    Relationship = "└─ IST MITGLIED VON (FEHLER)"
                                    RelatedObject = $groupDN
                                    RelatedObjectSAM = "N/A"
                                    RelatedObjectType = "Group (Fehler)"
                                }
                            }
                        }
                    } else {
                        $ReportOutput += [PSCustomObject]@{
                            ObjectName = ""
                            ObjectSAM = ""
                            ObjectType = ""
                            Relationship = "└─ KEINE GRUPPENMITGLIEDSCHAFTEN"
                            RelatedObject = ""
                            RelatedObjectSAM = ""
                            RelatedObjectType = ""
                        }
                    }
                } elseif ($objectClassSimple -eq 'group') {
                    if ($TargetObject.Member) {
                        foreach ($memberDN in $TargetObject.Member) {
                            try {
                                $memberObject = Get-ADObject -Identity $memberDN -Properties DisplayName, SamAccountName, ObjectClass -ErrorAction Stop
                                $memberObjectClassSimple = if ($memberObject.ObjectClass -is [array]) {
                                    $memberObject.ObjectClass[-1]
                                } else {
                                    $memberObject.ObjectClass
                                }
                                $memberName = if ($memberObject.DisplayName) { $memberObject.DisplayName } else { $memberObject.Name }
                                $memberSam = if ($memberObject.SamAccountName) { $memberObject.SamAccountName } else { "N/A" }

                                $ReportOutput += [PSCustomObject]@{
                                    ObjectName = ""
                                    ObjectSAM = ""
                                    ObjectType = ""
                                    Relationship = "└─ HAT MITGLIED"
                                    RelatedObject = $memberName
                                    RelatedObjectSAM = $memberSam
                                    RelatedObjectType = $memberObjectClassSimple
                                }
                            } catch {
                                Write-ADReportLog -Message "Error resolving member DN '$memberDN' for group '$objDisplayName': $($_.Exception.Message)" -Type Warning
                                $ReportOutput += [PSCustomObject]@{
                                    ObjectName = ""
                                    ObjectSAM = ""
                                    ObjectType = ""
                                    Relationship = "└─ HAT MITGLIED (FEHLER)"
                                    RelatedObject = $memberDN
                                    RelatedObjectSAM = "N/A"
                                    RelatedObjectType = "Unbekannt (Fehler)"
                                }
                            }
                        }
                    } else {
                        $ReportOutput += [PSCustomObject]@{
                            ObjectName = ""
                            ObjectSAM = ""
                            ObjectType = ""
                            Relationship = "└─ KEINE MITGLIEDER"
                            RelatedObject = ""
                            RelatedObjectSAM = ""
                            RelatedObjectType = ""
                        }
                    }
                }

                # Leerzeile nach jedem Objekt für bessere Lesbarkeit
                $ReportOutput += [PSCustomObject]@{
                    ObjectName = ""
                    ObjectSAM = ""
                    ObjectType = ""
                    Relationship = ""
                    RelatedObject = ""
                    RelatedObjectSAM = ""
                    RelatedObjectType = ""
                }
            }
        }
    } catch {
        $ErrorMessage = "Error in Get-ADGroupMembershipsReportData: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
    }
    return $ReportOutput
}

# --- Security Audit Functions ---
Function Get-WeakPasswordPolicyUsers {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing users with weak password policies..." -Type Info -Terminal
        
        # Properties relevant for comprehensive password policy analysis
        $Properties = @(
            "DisplayName", "SamAccountName", "Enabled", "PasswordNeverExpires", 
            "PasswordNotRequired", "PasswordLastSet", "LastLogonDate", "AdminCount",
            "CannotChangePassword", "SmartcardLogonRequired", "TrustedForDelegation",
            "DoesNotRequirePreAuth", "UseDESKeyOnly", "AccountExpirationDate",
            "LastBadPasswordAttempt", "BadLogonCount", "LogonCount", "PrimaryGroup",
            "MemberOf", "ServicePrincipalNames", "UserAccountControl", "LockedOut",
            "TrustedToAuthForDelegation", "AllowReversiblePasswordEncryption",
            "whenCreated", "Description", "UserPrincipalName", "DistinguishedName"
        )
        
        # Retrieve Domain Password Policy for comparisons
        $DomainPasswordPolicy = Get-ADDefaultDomainPasswordPolicy
        $MinPasswordAge = $DomainPasswordPolicy.MinPasswordAge.Days
        $MaxPasswordAge = $DomainPasswordPolicy.MaxPasswordAge.Days
        $MinPasswordLength = $DomainPasswordPolicy.MinPasswordLength
        
        Write-ADReportLog -Message "Domain Password Policy - Min Age: $MinPasswordAge days, Max Age: $MaxPasswordAge days, Min Length: $MinPasswordLength chars" -Type Info -Terminal
        
        # Load all users with relevant properties
        $AllUsers = Get-ADUser -Filter * -Properties $Properties | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
        Write-ADReportLog -Message "$($AllUsers.Count) users loaded for analysis..." -Type Info -Terminal
        
        # Define high-risk service account patterns
        $ServiceAccountPatterns = @("svc", "service", "app", "sql", "iis", "web", "backup", "sync", "admin", "sa")
        $TestAccountPatterns = @("test", "temp", "demo", "guest", "anonymous", "trial")
        
        # Enhanced analysis for weak password policies
        $WeakPasswordUsers = foreach ($user in $AllUsers) {
            $issues = @()
            $riskLevel = 0
            $recommendations = @()
            $securityFlags = @()
            
            # 1. Password never expires
            if ($user.PasswordNeverExpires -eq $true) {
                $issues += "Password never expires"
                $riskLevel += 3
                $recommendations += "Enable password expiration"
                $securityFlags += "NO_EXPIRY"
            }
            
            # 2. Password not required
            if ($user.PasswordNotRequired -eq $true) {
                $issues += "Password not required"
                $riskLevel += 5  # Critical risk
                $recommendations += "Enforce password requirement"
                $securityFlags += "NO_PASSWORD_REQ"
            }
            
            # 3. No password set or extremely old password
            if ($user.PasswordLastSet -eq $null) {
                $issues += "No password set"
                $riskLevel += 5
                $recommendations += "Set password immediately"
                $securityFlags += "NO_PASSWORD"
            } elseif ($user.PasswordLastSet -lt (Get-Date).AddDays(-($MaxPasswordAge * 3))) {
                $issues += "Password extremely outdated (>$($MaxPasswordAge * 3) days)"
                $riskLevel += 4
                $recommendations += "Force password reset immediately"
                $securityFlags += "ANCIENT_PASSWORD"
            } elseif ($user.PasswordLastSet -lt (Get-Date).AddDays(-($MaxPasswordAge * 2))) {
                $issues += "Password very old (>$($MaxPasswordAge * 2) days)"
                $riskLevel += 3
                $recommendations += "Schedule password reset"
                $securityFlags += "OLD_PASSWORD"
            } elseif ($user.PasswordLastSet -lt (Get-Date).AddDays(-365)) {
                $issues += "Password older than 1 year"
                $riskLevel += 2
                $recommendations += "Password reset recommended"
                $securityFlags += "STALE_PASSWORD"
            }
            
            # 4. User cannot change password
            if ($user.CannotChangePassword -eq $true) {
                $issues += "Cannot change password"
                $riskLevel += 2
                $recommendations += "Allow password changes (except for service accounts)"
                $securityFlags += "NO_CHANGE_ALLOWED"
            }
            
            # 5. Kerberos Pre-Authentication disabled (ASREPRoast attack possible)
            if ($user.DoesNotRequirePreAuth -eq $true) {
                $issues += "Kerberos Pre-Auth disabled (ASREPRoast vulnerability)"
                $riskLevel += 4
                $recommendations += "Enable Kerberos Pre-Authentication"
                $securityFlags += "ASREP_ROASTABLE"
            }
            
            # 6. Weak encryption (DES)
            if ($user.UseDESKeyOnly -eq $true) {
                $issues += "Uses weak DES encryption"
                $riskLevel += 3
                $recommendations += "Disable DES encryption"
                $securityFlags += "WEAK_ENCRYPTION"
            }
            
            # 7. Reversible password encryption
            if ($user.AllowReversiblePasswordEncryption -eq $true) {
                $issues += "Reversible password encryption enabled"
                $riskLevel += 4
                $recommendations += "Disable reversible password encryption"
                $securityFlags += "REVERSIBLE_ENCRYPTION"
            }
            
            # 8. Smartcard authentication not used for privileged accounts
            if ($user.AdminCount -eq 1 -and $user.SmartcardLogonRequired -eq $false) {
                $issues += "Privileged account without smartcard requirement"
                $riskLevel += 3
                $recommendations += "Enable smartcard authentication for admin accounts"
                $securityFlags += "ADMIN_NO_SMARTCARD"
            }
            
            # 9. Delegation for normal user accounts
            if (($user.TrustedForDelegation -eq $true -or $user.TrustedToAuthForDelegation -eq $true) -and $user.AdminCount -ne 1) {
                $issues += "Delegation enabled for standard user"
                $riskLevel += 3
                $recommendations += "Restrict delegation to service accounts only"
                $securityFlags += "UNEXPECTED_DELEGATION"
            }
            
            # 10. Excessive failed logon attempts
            if ($user.BadLogonCount -gt 10) {
                $issues += "Excessive failed logon attempts ($($user.BadLogonCount))"
                $riskLevel += 2
                $recommendations += "Investigate account for potential compromise"
                $securityFlags += "HIGH_FAILED_LOGONS"
            } elseif ($user.BadLogonCount -gt 5) {
                $issues += "Multiple failed logon attempts ($($user.BadLogonCount))"
                $riskLevel += 1
                $recommendations += "Monitor account activity"
                $securityFlags += "FAILED_LOGONS"
            }
            
            # 11. Service account without SPN (potentially misconfigured)
            $isServiceAccount = $false
            foreach ($pattern in $ServiceAccountPatterns) {
                if ($user.SamAccountName -like "*$pattern*" -or $user.DisplayName -like "*$pattern*") {
                    $isServiceAccount = $true
                    break
                }
            }
            
            if ($isServiceAccount) {
                if (-not $user.ServicePrincipalNames) {
                    $issues += "Service account without SPN"
                    $riskLevel += 1
                    $recommendations += "Configure SPN for service account"
                    $securityFlags += "SERVICE_NO_SPN"
                }
                
                # Service accounts should not be interactive
                if ($user.LastLogonDate -and $user.LastLogonDate -gt (Get-Date).AddDays(-30)) {
                    $issues += "Interactive logons detected for service account"
                    $riskLevel += 2
                    $recommendations += "Review service account usage"
                    $securityFlags += "SERVICE_INTERACTIVE"
                }
            }
            
            # 12. Never logged in accounts with high privileges
            if ($user.LogonCount -eq 0 -and $user.AdminCount -eq 1) {
                $issues += "Admin account never used"
                $riskLevel += 3
                $recommendations += "Disable unused admin account"
                $securityFlags += "UNUSED_ADMIN"
            } elseif ($user.LogonCount -eq 0 -and $user.whenCreated -lt (Get-Date).AddDays(-30)) {
                $issues += "Account never used (created >30 days ago)"
                $riskLevel += 1
                $recommendations += "Consider disabling unused account"
                $securityFlags += "NEVER_USED"
            }
            
            # 13. Inactive privileged accounts
            if ($user.AdminCount -eq 1 -and $user.LastLogonDate -and $user.LastLogonDate -lt (Get-Date).AddDays(-60)) {
                $issues += "Privileged account inactive for >60 days"
                $riskLevel += 3
                $recommendations += "Review inactive admin account"
                $securityFlags += "INACTIVE_ADMIN"
            } elseif ($user.LastLogonDate -and $user.LastLogonDate -lt (Get-Date).AddDays(-180)) {
                $issues += "Account inactive for >180 days"
                $riskLevel += 2
                $recommendations += "Consider disabling inactive account"
                $securityFlags += "LONG_INACTIVE"
            }
            
            # 14. Test/temp accounts without expiration
            $isTestAccount = $false
            foreach ($pattern in $TestAccountPatterns) {
                if ($user.SamAccountName -like "*$pattern*" -or $user.DisplayName -like "*$pattern*") {
                    $isTestAccount = $true
                    break
                }
            }
            
            if ($isTestAccount) {
                if ($user.AccountExpirationDate -eq $null) {
                    $issues += "Test/temp account without expiration date"
                    $riskLevel += 2
                    $recommendations += "Set expiration date for temporary accounts"
                    $securityFlags += "TEST_NO_EXPIRY"
                }
                
                if ($user.whenCreated -lt (Get-Date).AddDays(-90)) {
                    $issues += "Old test account (>90 days)"
                    $riskLevel += 1
                    $recommendations += "Review necessity of old test account"
                    $securityFlags += "OLD_TEST_ACCOUNT"
                }
            }
            
            # 15. Locked accounts with admin privileges
            if ($user.LockedOut -eq $true -and $user.AdminCount -eq 1) {
                $issues += "Locked privileged account"
                $riskLevel += 2
                $recommendations += "Investigate locked admin account"
                $securityFlags += "LOCKED_ADMIN"
            }
            
            # 16. Accounts with suspicious creation patterns
            if ($user.whenCreated -gt (Get-Date).AddDays(-7) -and $user.AdminCount -eq 1) {
                $issues += "Recently created admin account"
                $riskLevel += 2
                $recommendations += "Verify legitimacy of new admin account"
                $securityFlags += "NEW_ADMIN"
            }
            
            # 17. Accounts with generic or weak naming
            $weakNames = @("admin", "administrator", "user", "guest", "test", "temp", "service", "default")
            foreach ($weakName in $weakNames) {
                if ($user.SamAccountName -eq $weakName -or $user.SamAccountName -like "$weakName*") {
                    $issues += "Generic/predictable account name"
                    $riskLevel += 1
                    $recommendations += "Use non-predictable account names"
                    $securityFlags += "WEAK_NAMING"
                    break
                }
            }
            
            # Only return users with identified vulnerabilities
            if ($issues.Count -gt 0) {
                # Categorize risk level
                $riskCategory = switch ([int]$riskLevel) {
                    {$_ -ge 10} { [string]"Critical" }
                    {$_ -ge 7} { [string]"High" }
                    {$_ -ge 4} { [string]"Medium" }
                    {$_ -ge 2} { [string]"Low" }
                    default    { [string]"Minimal" }
                }
                

                # Add context information
                $contextInfo = @()
                if ($user.AdminCount -eq 1) { $contextInfo += "Privileged Account" }
                if ($user.Enabled -eq $false) { $contextInfo += "Disabled" }
                if ($user.LockedOut -eq $true) { $contextInfo += "Locked" }
                if ($user.ServicePrincipalNames) { $contextInfo += "Service Account" }
                if ($isServiceAccount) { $contextInfo += "Service Pattern" }
                if ($isTestAccount) { $contextInfo += "Test/Temp Account" }
                if ($user.SamAccountName -match "^(admin|administrator|root|sa|service|svc)") { $contextInfo += "System Account" }
                
                # Calculate compliance status
                $complianceIssues = 0
                if ($user.PasswordNeverExpires) { $complianceIssues++ }
                if ($user.PasswordNotRequired) { $complianceIssues++ }
                if ($user.DoesNotRequirePreAuth) { $complianceIssues++ }
                if ($user.UseDESKeyOnly) { $complianceIssues++ }
                if ($user.AllowReversiblePasswordEncryption) { $complianceIssues++ }
                
                $complianceStatus = if ($complianceIssues -eq 0) { "Compliant" } 
                                   elseif ($complianceIssues -le 2) { "Partially Compliant" } 
                                   else { "Non-Compliant" }
                
                # Calculate password age in days
                $passwordAge = if ($user.PasswordLastSet) { 
                    [math]::Round((New-TimeSpan -Start $user.PasswordLastSet -End (Get-Date)).TotalDays) 
                } else { "Never Set" }
                
                # Calculate account age
                $accountAge = [math]::Round((New-TimeSpan -Start $user.whenCreated -End (Get-Date)).TotalDays)
                
                # Determine urgency level
                $urgencyLevel = if ($user.AdminCount -eq 1 -and $riskLevel -ge 7) { "Immediate Action Required" }
                               elseif ($riskLevel -ge 10) { "Critical" }
                               elseif ($riskLevel -ge 7) { "Urgent" }
                               elseif ($riskLevel -ge 4) { "High Priority" }
                               elseif ($riskLevel -ge 2) { "Medium Priority" }
                               else { "Low Priority" }
                
                # Enhanced output object with comprehensive analysis
                [PSCustomObject]@{
                    DisplayName = $user.DisplayName
                    SamAccountName = $user.SamAccountName
                    UserPrincipalName = $user.UserPrincipalName
                    Enabled = $user.Enabled
                    LockedOut = $user.LockedOut
                    Context = if ($contextInfo) { $contextInfo -join ", " } else { "Standard User" }
                    PasswordLastSet = $user.PasswordLastSet
                    PasswordAge = $passwordAge
                    AccountCreated = $user.whenCreated
                    AccountAge = $accountAge
                    LastLogonDate = $user.LastLogonDate
                    LogonCount = $user.LogonCount
                    BadLogonCount = $user.BadLogonCount
                    LastBadPasswordAttempt = $user.LastBadPasswordAttempt
                    Vulnerabilities = $issues -join "; "
                    SecurityFlags = $securityFlags -join ", "
                    RiskLevel = $riskCategory
                    RiskScore = $riskLevel
                    ComplianceStatus = $complianceStatus
                    ComplianceIssues = $complianceIssues
                    UrgencyLevel = $urgencyLevel
                    Recommendations = $recommendations -join "; "
                    Description = $user.Description
                    LastAssessment = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    TotalIssuesFound = $issues.Count
                    RequiresImmediateAction = ($urgencyLevel -eq "Immediate Action Required" -or $urgencyLevel -eq "Critical")
                }
            }
        }
        
        # Enhanced statistics for logging
        $totalIssues = $WeakPasswordUsers.Count
        $criticalIssues = ($WeakPasswordUsers | Where-Object { $_.RiskLevel -eq "Critical" }).Count
        $highIssues = ($WeakPasswordUsers | Where-Object { $_.RiskLevel -eq "High" }).Count
        $mediumIssues = ($WeakPasswordUsers | Where-Object { $_.RiskLevel -eq "Medium" }).Count
        $adminIssues = ($WeakPasswordUsers | Where-Object { $_.Context -like "*Privileged Account*" }).Count
        $nonCompliant = ($WeakPasswordUsers | Where-Object { $_.ComplianceStatus -eq "Non-Compliant" }).Count
        $immediateAction = ($WeakPasswordUsers | Where-Object { $_.RequiresImmediateAction -eq $true }).Count
        $serviceAccountIssues = ($WeakPasswordUsers | Where-Object { $_.Context -like "*Service*" }).Count
        $testAccountIssues = ($WeakPasswordUsers | Where-Object { $_.Context -like "*Test*" }).Count
        
        Write-ADReportLog -Message "Password Policy Analysis completed:" -Type Info -Terminal
        Write-ADReportLog -Message "  Total: $totalIssues users with vulnerabilities" -Type Info -Terminal
        Write-ADReportLog -Message "  Risk Distribution - Critical: $criticalIssues, High: $highIssues, Medium: $mediumIssues" -Type Info -Terminal
        Write-ADReportLog -Message "  Privileged accounts affected: $adminIssues" -Type Info -Terminal
        Write-ADReportLog -Message "  Service accounts with issues: $serviceAccountIssues" -Type Info -Terminal
        Write-ADReportLog -Message "  Test/temp accounts with issues: $testAccountIssues" -Type Info -Terminal
        Write-ADReportLog -Message "  Compliance violations: $nonCompliant" -Type Info -Terminal
        Write-ADReportLog -Message "  Requiring immediate action: $immediateAction" -Type Info -Terminal
        
        return $WeakPasswordUsers
        
    } catch {
        $ErrorMessage = "Error analyzing password policies: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-RiskyGroupMemberships {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analysiere riskante Gruppenmitgliedschaften..." -Type Info -Terminal
        
        # Definiere hochprivilegierte Gruppen mit deutschen und englischen Namen
        $RiskyGroups = @(
            'Domain Controllers', 'Domänencontroller',
            'Enterprise Admins',
            'Domain Admins', 'Domänen-Admins', 
            'Account Operators',
            'Remote Desktop Users',
            'Enterprise Domain Controllers', 'Organisations-Domänencontroller',
            'Schema Admins',
            'Backup Operators',
            'Replicator',
            'Server Operators', 
            'Print Operators',
            'Power Users', 'Hauptbenutzer',
            'Administrators'
        )

        # Füge zusätzliche Gruppen aus der Konfiguration hinzu
        $RiskyGroups += foreach ($groupType in $Global:ADGroupNames.GetEnumerator()) {
            if ([string]::IsNullOrEmpty($groupType.Value)) {
                Write-ADReportLog -Message "Warnung: Leerer Gruppenname für $($groupType.Key)" -Type Warning
                continue
            }
            $groupType.Value
        }
        
        $RiskyUsers = [System.Collections.Generic.List[PSObject]]::new()
        
        # Gruppenanalyse mit erweiterter Fehlerbehandlung
        foreach ($groupName in $RiskyGroups) {
            try {
                if ([string]::IsNullOrEmpty($groupName)) {
                    Write-ADReportLog -Message "Überspringe leeren Gruppennamen" -Type Warning
                    continue
                }

                # Versuche beide Namensformate (deutsch/englisch)
                $group = Get-ADGroup -Filter "Name -eq '$groupName' -or SamAccountName -eq '$groupName'" -ErrorAction SilentlyContinue
                if (-not $group) {
                    Write-ADReportLog -Message "Gruppe '$groupName' nicht gefunden" -Type Warning
                    continue
                }
                
                $members = Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction Stop |
                    Where-Object { $_.objectClass -eq "user" } |
                    ForEach-Object {
                        try {
                            Get-ADObject -Identity $_.DistinguishedName -Properties DisplayName, SamAccountName, ObjectClass -ErrorAction Stop
                        } catch {
                            Write-ADReportLog -Message "Fehler beim Abrufen des Benutzerobjekts: $($_.Exception.Message)" -Type Warning
                            $null
                        }
                    } | Where-Object { $_ -ne $null }
                
                foreach ($member in $members) {
                    try {
                        $userDetails = Get-ADUser -Identity $member.DistinguishedName -Properties DisplayName, Enabled, LastLogonDate, PasswordLastSet -ErrorAction Stop |
                            Select-Object DisplayName, SamAccountName, Enabled, LastLogonDate, PasswordLastSet
                        
                        # Dynamische Risikobewertung basierend auf Gruppennamen
                        $riskLevel = switch -Wildcard ($group.Name) {
                            { $_ -match '(Domain Admins|Domänen-Admins|Enterprise Admins|Schema Admins)' } { "Kritisch" }
                            { $_ -match '(Administrators|Domain Controllers|Domänencontroller)' } { "Hoch" }
                            { $_ -match '(Account Operators|Server Operators|Backup Operators|Print Operators)' } { "Mittel" }
                            default { "Niedrig" }
                        }
                        
                        # Empfehlungsgenerierung als String
                        $recommendation = if (-not $userDetails.Enabled) {
                            "Deaktiviertes Konto aus Gruppe entfernen"
                        } elseif ($null -eq $userDetails.LastLogonDate) {
                            "Konto wurde nie angemeldet - überprüfen"
                        } elseif ($userDetails.LastLogonDate -lt (Get-Date).AddDays(-90)) {
                            "Inaktives Konto überprüfen"
                        } else {
                            "Berechtigungen regelmäßig überwachen"
                        }
                        
                        # Objekterstellung mit Typisierung
                        $riskUser = [PSCustomObject]@{
                            DisplayName = $userDetails.DisplayName
                            SamAccountName = $userDetails.SamAccountName
                            Enabled = [bool]$userDetails.Enabled
                            LastLogonDate = $userDetails.LastLogonDate
                            PasswordLastSet = $userDetails.PasswordLastSet
                            RisikoGruppe = $group.Name
                            Risikostufe = $riskLevel
                            Empfehlung = $recommendation
                        }
                        
                        $RiskyUsers.Add($riskUser)
                    }
                    catch {
                        Write-ADReportLog -Message "Fehler bei Benutzer $($member.SamAccountName): $($_.Exception.Message)" -Type Warning
                    }
                }
            }
            catch {
                Write-ADReportLog -Message "Gruppenanalyse fehlgeschlagen für '$groupName': $($_.Exception.Message)" -Type Warning
            }
        }
        
        # Deduplizierung und Ausgabe
        $UniqueRiskyUsers = $RiskyUsers | Sort-Object SamAccountName -Unique
        
        Write-ADReportLog -Message "Gruppenanalyse abgeschlossen. Gefundene Risikofälle: $($UniqueRiskyUsers.Count)" -Type Info -Terminal
        return $UniqueRiskyUsers
    }
    catch {
        $ErrorMessage = "Gesamtfehler Gruppenanalyse: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error -Terminal
        return $null
    }
}

Function Get-PrivilegedAccounts {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analysiere Konten mit erhÃ¶hten Rechten..." -Type Info -Terminal
        
        # Eigenschaften fÃ¼r privilegierte Konten
        $Properties = @(
            "DisplayName", "SamAccountName", "Enabled", "AdminCount", 
            "LastLogonDate", "PasswordLastSet", "PasswordNeverExpires",
            "ServicePrincipalNames", "TrustedForDelegation", "TrustedToAuthForDelegation"
        )
        
        # Alle Benutzer mit AdminCount = 1 (historisch privilegiert)
        $AdminCountUsers = Get-ADUser -Filter "AdminCount -eq 1" -Properties $Properties | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
        
        # Service-Konten (Konten mit SPNs)
        $ServiceAccounts = Get-ADUser -Filter "ServicePrincipalNames -like '*'" -Properties $Properties | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
        
        # Konten mit Delegierungsrechten
        $DelegationAccounts = Get-ADUser -Filter "TrustedForDelegation -eq `$true -or TrustedToAuthForDelegation -eq `$true" -Properties $Properties | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
        
        # Alle privilegierten Konten zusammenfÃ¼hren
        $AllPrivilegedAccounts = @()
        $AllPrivilegedAccounts += $AdminCountUsers
        $AllPrivilegedAccounts += $ServiceAccounts
        $AllPrivilegedAccounts += $DelegationAccounts
        
        # Duplikate entfernen und analysieren
        $UniquePrivilegedAccounts = $AllPrivilegedAccounts | Sort-Object SamAccountName -Unique | ForEach-Object {
            $account = $_
            
            # Risikofaktoren analysieren
            $riskFactors = @()
            $riskLevel = 0
            
            if ($account.AdminCount -eq 1) {
                $riskFactors += "AdminCount set"
                $riskLevel += 2
            }
            
            if ($account.ServicePrincipalNames) {
                $riskFactors += "Service-Account (SPN)"
                $riskLevel += 1
            }
            
            if ($account.TrustedForDelegation) {
                $riskFactors += "Delegation activated"
                $riskLevel += 2
            }
            
            if ($account.TrustedToAuthForDelegation) {
                $riskFactors += "Constrained Delegation"
                $riskLevel += 1
            }
            
            if ($account.PasswordNeverExpires) {
                $riskFactors += "Password never expires"
                $riskLevel += 1
            }
            
            if ($account.Enabled -eq $false) {
                $riskFactors += "Account disabled"
                $riskLevel += 3  # Deactivated privileged accounts are a high risk
            }

            if ($account.LastLogonDate -and $account.LastLogonDate -lt (Get-Date).AddDays(-90)) {
                $riskFactors += "Inactive (>90 days)"
                $riskLevel += 2
            }
            
            # Privilegien-Level bestimmen
            $privilegeLevel = "Standard"
            if ($account.AdminCount -eq 1 -and ($account.TrustedForDelegation -or $account.TrustedToAuthForDelegation)) {
                $privilegeLevel = "Critical"
            } elseif ($account.AdminCount -eq 1) {
                $privilegeLevel = "High"
            } elseif ($account.ServicePrincipalNames -or $account.TrustedForDelegation) {
                $privilegeLevel = "Medium"
            }
            
            # Risikostufe bestimmen
            $overallRisk = switch ($riskLevel) {
                {$_ -ge 5} { "Critical" }
                {$_ -ge 3} { "High" }
                {$_ -ge 2} { "Medium" }
                default { "Low" }
            }
            
            # Empfehlungen generieren
            $recommendations = @()
            if ($account.Enabled -eq $false -and $account.AdminCount -eq 1) {
                $recommendations += "Reset AdminCount for deactivated account"
            }
            if ($account.LastLogonDate -and $account.LastLogonDate -lt (Get-Date).AddDays(-90)) {
                $recommendations += "Check account usage"
            }
            if ($account.PasswordNeverExpires -and $account.AdminCount -eq 1) {
                $recommendations += "Enable password expiration"
            }
            if ($account.TrustedForDelegation) {
                $recommendations += "Review delegation rights"
            }
            
            [PSCustomObject]@{
                AdminCount = $account.AdminCount
                ServiceAccount = [bool]$account.ServicePrincipalNames
                DisplayName = $account.DisplayName
                SamAccountName = $account.SamAccountName
                Enabled = $account.Enabled
                LastLogonDate = $account.LastLogonDate
                PasswordLastSet = $account.PasswordLastSet
                PrivilegeLevel = $privilegeLevel
                RiskFactors = $riskFactors -join "; "
                Recommendations = if ($recommendations) { $recommendations -join "; " } else { "Regular review" }
                Delegation = $account.TrustedForDelegation -or $account.TrustedToAuthForDelegation
            }
        }
        
        Write-ADReportLog -Message "$($UniquePrivilegedAccounts.Count) Konten mit erhÃ¶hten Rechten gefunden." -Type Info -Terminal
        return $UniquePrivilegedAccounts
        
    } catch {
        $ErrorMessage = "Fehler beim Analysieren der privilegierten Konten: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Organisationsstruktur-Reports ---
Function Get-DepartmentStatistics {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing department statistics..." -Type Info -Terminal
        
        # Lade alle Benutzer mit Abteilungsinformationen
        $Users = Get-ADUser -Filter * -Properties Department, Enabled, LastLogonDate, whenCreated, PasswordLastSet, PasswordNeverExpires, LockedOut, Title -ErrorAction Stop
        
        # Gruppiere nach Abteilung und erstelle leeres Array für Ergebnisse
        $DepartmentStats = [System.Collections.ArrayList]@()
        $DepartmentGroups = $Users | Group-Object Department
        
        foreach ($dept in $DepartmentGroups) {
            $deptName = if ([string]::IsNullOrWhiteSpace($dept.Name)) { "(No Department)" } else { $dept.Name }
            $deptUsers = $dept.Group
            
            # Statistiken berechnen
            $enabledCount = ($deptUsers | Where-Object { $_.Enabled -eq $true } | Measure-Object).Count
            $disabledCount = ($deptUsers | Where-Object { $_.Enabled -eq $false } | Measure-Object).Count
            $lockedCount = ($deptUsers | Where-Object { $_.LockedOut -eq $true } | Measure-Object).Count
            $neverExpireCount = ($deptUsers | Where-Object { $_.PasswordNeverExpires -eq $true } | Measure-Object).Count
            $inactiveCount = ($deptUsers | Where-Object { 
                $_.LastLogonDate -and 
                $_.LastLogonDate -lt (Get-Date).AddDays(-90) 
            } | Measure-Object).Count
            
            # Durchschnittliches Kontoalter berechnen
            $accountAges = @()
            foreach ($user in $deptUsers) {
                if ($user.whenCreated) {
                    $age = [int]((Get-Date) - $user.whenCreated).Days
                    $accountAges += $age
                }
            }
            
            $avgAccountAge = if ($accountAges.Count -gt 0) {
                [math]::Round(($accountAges | Measure-Object -Average).Average, 1)
            } else {
                0
            }
            
            # Titel zählen
            $uniqueTitles = @($deptUsers | Where-Object { -not [string]::IsNullOrEmpty($_.Title) } | Select-Object -ExpandProperty Title -Unique).Count
            
            # Sicherheits-Score berechnen
            $totalUsers = ($deptUsers | Measure-Object).Count
            $securityScore = 100
            if ($totalUsers -gt 0) {
                $issueCount = [int]($disabledCount + $lockedCount + $neverExpireCount + $inactiveCount)
                $securityScore = [math]::Round(100 - (($issueCount / $totalUsers) * 100), 1)
            }
            
            # Füge Statistiken zum Array hinzu
            $null = $DepartmentStats.Add([PSCustomObject]@{
                Department = $deptName
                TotalUsers = $totalUsers
                EnabledUsers = $enabledCount
                DisabledUsers = $disabledCount
                LockedUsers = $lockedCount
                InactiveUsers = $inactiveCount
                PasswordNeverExpires = $neverExpireCount
                AvgAccountAgeDays = $avgAccountAge
                UniqueTitles = $uniqueTitles
                SecurityScore = $securityScore
            })
        }
        
        Write-ADReportLog -Message "Department statistics analysis completed. $($DepartmentStats.Count) departments found." -Type Info -Terminal
        return $DepartmentStats | Sort-Object TotalUsers -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing department statistics: $($_.Exception.Message)" -Type Error
        return $null
    }
}

Function Get-DepartmentSecurityRisks {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing security risks by department..." -Type Info -Terminal
        
        # Lade alle Benutzer mit erweiterten Eigenschaften
        $Users = Get-ADUser -Filter * -Properties Department, Enabled, LastLogonDate, PasswordLastSet, PasswordNeverExpires, 
                                                  LockedOut, AdminCount, ServicePrincipalNames, DoesNotRequirePreAuth,
                                                  TrustedForDelegation, AllowReversiblePasswordEncryption -ErrorAction Stop
        
        # Gruppiere nach Abteilung und erstelle leeres Array für Ergebnisse
        $DepartmentRisks = [System.Collections.ArrayList]@()
        $DepartmentGroups = $Users | Group-Object Department
        
        foreach ($dept in $DepartmentGroups) {
            $deptName = if ([string]::IsNullOrWhiteSpace($dept.Name)) { "(No Department)" } else { $dept.Name }
            $deptUsers = $dept.Group
            
            # Risiko-Metriken berechnen
            $adminUsers = ($deptUsers | Where-Object { $_.AdminCount -eq 1 } | Measure-Object).Count
            $serviceAccounts = ($deptUsers | Where-Object { $null -ne $_.ServicePrincipalNames -and @($_.ServicePrincipalNames).Count -gt 0 } | Measure-Object).Count
            $kerberoastable = ($deptUsers | Where-Object { $null -ne $_.ServicePrincipalNames -and @($_.ServicePrincipalNames).Count -gt 0 -and $_.Enabled -eq $true } | Measure-Object).Count
            $asrepRoastable = ($deptUsers | Where-Object { $_.DoesNotRequirePreAuth -eq $true } | Measure-Object).Count
            $delegationEnabled = ($deptUsers | Where-Object { $_.TrustedForDelegation -eq $true } | Measure-Object).Count
            $reversiblePwd = ($deptUsers | Where-Object { $_.AllowReversiblePasswordEncryption -eq $true } | Measure-Object).Count
            $neverExpire = ($deptUsers | Where-Object { $_.PasswordNeverExpires -eq $true } | Measure-Object).Count
            $oldPasswords = ($deptUsers | Where-Object { $_.PasswordLastSet -and $_.PasswordLastSet -lt (Get-Date).AddDays(-180) } | Measure-Object).Count
            
            # Risiko-Score berechnen
            [int]$riskScore = 0
            $riskScore += [int]($adminUsers * 3)
            $riskScore += [int]($kerberoastable * 2)
            $riskScore += [int]($asrepRoastable * 3)
            $riskScore += [int]($delegationEnabled * 2)
            $riskScore += [int]($reversiblePwd * 4)
            $riskScore += [int]($neverExpire * 0.5)
            $riskScore += [int]($oldPasswords * 0.3)
            
            # Risiko-Level bestimmen
            $riskLevel = switch ($riskScore) {
                {$_ -ge 20} { "Critical" }
                {$_ -ge 10} { "High" }
                {$_ -ge 5} { "Medium" }
                {$_ -gt 0} { "Low" }
                default { "Minimal" }
            }
            
            # Füge Risiken zum Array hinzu
            $null = $DepartmentRisks.Add([PSCustomObject]@{
                Department = $deptName
                TotalUsers = ($deptUsers | Measure-Object).Count
                AdminUsers = $adminUsers
                ServiceAccounts = $serviceAccounts
                Kerberoastable = $kerberoastable
                ASREPRoastable = $asrepRoastable
                DelegationEnabled = $delegationEnabled
                ReversiblePasswords = $reversiblePwd
                PasswordNeverExpires = $neverExpire
                OldPasswords = $oldPasswords
                RiskScore = $riskScore
                RiskLevel = $riskLevel
            })
        }
        
        Write-ADReportLog -Message "Department security risk analysis completed." -Type Info -Terminal
        return $DepartmentRisks | Sort-Object RiskScore -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing department security risks: $($_.Exception.Message)" -Type Error
        return $null
    }
}

Function Get-ASREPRoastableAccounts {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing ASREPRoastable accounts (pre-authentication disabled)..." -Type Info -Terminal
        
        # Lade alle Benutzer mit deaktivierter Kerberos-Vorauthentifizierung
        $ASREPUsers = Get-ADUser -Filter "DoesNotRequirePreAuth -eq `$true" -Properties DoesNotRequirePreAuth, DisplayName,
                                                                                        SamAccountName, Enabled, PasswordLastSet,
                                                                                        PasswordNeverExpires, LastLogonDate, AdminCount,
                                                                                        Description, whenCreated, Department -ErrorAction Stop
        
        $ASREPRoastableAccounts = [System.Collections.ArrayList]@()
        foreach ($user in $ASREPUsers) {
            # Risiko-Bewertung
            [System.Collections.ArrayList]$riskFactors = @()
            [int]$riskLevel = 5  # Basis-Risiko für ASREP Roastable ist hoch
            
            if ($user.AdminCount -eq 1) {
                $riskFactors.Add("Privileged Account") | Out-Null
                $riskLevel += 3
            }
            
            if ($user.Enabled -eq $true) {
                $riskFactors.Add("Account Enabled") | Out-Null
                $riskLevel += 1
            }
            
            if ($user.PasswordNeverExpires) {
                $riskFactors.Add("Password Never Expires") | Out-Null
                $riskLevel += 2
            }
            
            if ($user.LastLogonDate -and $user.LastLogonDate -gt (Get-Date).AddDays(-30)) {
                $riskFactors.Add("Recently Active") | Out-Null
                $riskLevel += 1
            }
            
            $overallRisk = switch ($riskLevel) {
                {$_ -ge 8} { "Critical" }
                {$_ -ge 6} { "High" }
                {$_ -ge 4} { "Medium" }
                default { "Low" }
            }
            
            $ASREPRoastableAccounts.Add([PSCustomObject]@{
                DisplayName = $user.DisplayName
                SamAccountName = $user.SamAccountName
                Department = $user.Department
                Enabled = $user.Enabled
                DoesNotRequirePreAuth = $true
                PasswordLastSet = $user.PasswordLastSet
                PasswordNeverExpires = $user.PasswordNeverExpires
                LastLogonDate = $user.LastLogonDate
                AdminAccount = ($user.AdminCount -eq 1)
                AccountAge = if ($user.whenCreated) { [math]::Round((New-TimeSpan -Start $user.whenCreated -End (Get-Date)).TotalDays) } else { "Unknown" }
                RiskFactors = $riskFactors -join "; "
                RiskLevel = $overallRisk
                Remediation = "Enable Kerberos pre-authentication immediately"
                Description = $user.Description
            }) | Out-Null
        }
        
        Write-ADReportLog -Message "ASREPRoastable accounts analysis completed. $($ASREPRoastableAccounts.Count) accounts found." -Type Info -Terminal
        return $ASREPRoastableAccounts | Sort-Object RiskLevel, AdminAccount -Descending
        
    } catch {
        $ErrorMessage = "Error analyzing ASREPRoastable accounts: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-DelegationAnalysis {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing delegation settings..." -Type Info -Terminal
        
        # Lade alle Objekte mit Delegierung
        $DelegatedObjects = @()
        
        # Unconstrained Delegation (sehr gefÃ¤hrlich)
        $UnconstrainedDelegation = Get-ADObject -Filter "TrustedForDelegation -eq `$true -and TrustedToAuthForDelegation -eq `$false" `
                                                -Properties Name, ObjectClass, DistinguishedName, whenCreated, whenChanged -ErrorAction SilentlyContinue
        
        foreach ($obj in $UnconstrainedDelegation) {
            $DelegatedObjects += [PSCustomObject]@{
                Name = $obj.Name
                ObjectType = $obj.ObjectClass
                DelegationType = "Unconstrained Delegation"
                DistinguishedName = $obj.DistinguishedName
                Risk = "Critical"
                Created = $obj.whenCreated
                LastModified = $obj.whenChanged
                Remediation = "Remove unconstrained delegation or convert to constrained delegation"
            }
        }
        
        # Constrained Delegation
        $ConstrainedDelegation = Get-ADObject -Filter "TrustedToAuthForDelegation -eq `$true" `
                                             -Properties Name, ObjectClass, DistinguishedName, msDS-AllowedToDelegateTo, whenCreated, whenChanged -ErrorAction SilentlyContinue
        
        foreach ($obj in $ConstrainedDelegation) {
            $allowedServices = if ($obj.'msDS-AllowedToDelegateTo') { $obj.'msDS-AllowedToDelegateTo' -join "; " } else { "None specified" }
            
            $DelegatedObjects += [PSCustomObject]@{
                Name = $obj.Name
                ObjectType = $obj.ObjectClass
                DelegationType = "Constrained Delegation"
                AllowedServices = $allowedServices
                DistinguishedName = $obj.DistinguishedName
                Risk = "Medium"
                Created = $obj.whenCreated
                LastModified = $obj.whenChanged
                Remediation = "Review allowed services and minimize delegation scope"
            }
        }
        
        # Resource-based Constrained Delegation
        $ResourceBasedDelegation = Get-ADObject -Filter "msDS-AllowedToActOnBehalfOfOtherIdentity -like '*'" `
                                               -Properties Name, ObjectClass, DistinguishedName, msDS-AllowedToActOnBehalfOfOtherIdentity, whenCreated, whenChanged -ErrorAction SilentlyContinue
        
        foreach ($obj in $ResourceBasedDelegation) {
            $DelegatedObjects += [PSCustomObject]@{
                Name = $obj.Name
                ObjectType = $obj.ObjectClass
                DelegationType = "Resource-based Constrained Delegation"
                DistinguishedName = $obj.DistinguishedName
                Risk = "High"
                Created = $obj.whenCreated
                LastModified = $obj.whenChanged
                Remediation = "Audit resource-based delegation permissions"
            }
        }
        
        Write-ADReportLog -Message "Delegation analysis completed. $($DelegatedObjects.Count) delegated objects found." -Type Info -Terminal
        return $DelegatedObjects | Sort-Object Risk, DelegationType
        
    } catch {
        $ErrorMessage = "Error analyzing delegation settings: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Erweiterte Privilegien-Eskalation Analyse ---
Function Get-DCSyncRights {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing DCSync rights..." -Type Info -Terminal
        
        $Domain = Get-ADDomain
        $DomainDN = $Domain.DistinguishedName
        $DCSyncRights = @()
        
        # Get ACL for domain object
        try {
            $DomainACL = Get-Acl "AD:\$DomainDN" -ErrorAction Stop
        } catch {
            Write-ADReportLog -Message "Could not access domain ACL. Requires appropriate permissions." -Type Warning -Terminal
            return @([PSCustomObject]@{
                Identity = "N/A"
                ObjectType = "Error"
                Permissions = "Could not read ACL"
                Risk = "Unknown"
                Remediation = "Run with Domain Admin privileges"
            })
        }
        
        # DCSync requires DS-Replication-Get-Changes and DS-Replication-Get-Changes-All
        $ReplicationGUIDs = @{
            "DS-Replication-Get-Changes" = "1131f6aa-9c07-11d1-f79f-00c04fc2dcd2"
            "DS-Replication-Get-Changes-All" = "1131f6ad-9c07-11d1-f79f-00c04fc2dcd2"
            "DS-Replication-Get-Changes-In-Filtered-Set" = "89e95b76-444d-4c62-991a-0facbeda640c"
        }
        
        # Analyze each ACE
        foreach ($ace in $DomainACL.Access) {
            if ($ace.AccessControlType -eq "Allow" -and $ace.ObjectType) {
                $aceObjectType = $ace.ObjectType.ToString()
                
                foreach ($guid in $ReplicationGUIDs.GetEnumerator()) {
                    if ($aceObjectType -eq $guid.Value) {
                        # Resolve the identity
                        $identity = $ace.IdentityReference.Value
                        try {
                            $adObject = Get-ADObject -Filter "SamAccountName -eq '$($identity.Split('\')[1])'" -Properties ObjectClass, AdminCount -ErrorAction SilentlyContinue
                            $objectType = if ($adObject) { $adObject.ObjectClass } else { "Unknown" }
                            $isAdmin = if ($adObject -and $adObject.AdminCount -eq 1) { $true } else { $false }
                        } catch {
                            $objectType = "Unknown"
                            $isAdmin = $false
                        }
                        
                        # Check if identity is a known privileged group
                        $isExpectedGroup = $false
                        $allExpectedGroups = @()
                        $allExpectedGroups += $Global:ADGroupNames.DomainControllers
                        $allExpectedGroups += $Global:ADGroupNames.EnterpriseDomainControllers  
                        $allExpectedGroups += $Global:ADGroupNames.Administrators
                        
                        foreach ($expectedGroup in $allExpectedGroups) {
                            if ($identity -like "*$expectedGroup*") {
                                $isExpectedGroup = $true
                                break
                            }
                        }
                        
                        $DCSyncRights += [PSCustomObject]@{
                            Identity = $identity
                            Permission = $guid.Key
                            ObjectType = $objectType
                            IsPrivileged = $isAdmin
                            Risk = if (-not $isExpectedGroup) { "High" } else { "Expected" }
                            Remediation = if (-not $isExpectedGroup) { 
                                "Review and remove DCSync rights from this identity" 
                            } else { 
                                "Standard DCSync permission - monitor for changes" 
                            }
                        }
                    }
                }
            }
        }
        
        # Group by identity to show complete DCSync capability
        $DCSyncCapable = $DCSyncRights | Group-Object Identity | Where-Object {
            ($_.Group.Permission -contains "DS-Replication-Get-Changes") -and 
            ($_.Group.Permission -contains "DS-Replication-Get-Changes-All")
        }
        
        $FinalReport = foreach ($capable in $DCSyncCapable) {
            $permissions = $capable.Group.Permission -join ", "
            $firstEntry = $capable.Group[0]
            
            [PSCustomObject]@{
                Identity = $capable.Name
                ObjectType = $firstEntry.ObjectType
                IsPrivileged = $firstEntry.IsPrivileged
                Permissions = $permissions
                HasFullDCSync = $true
                Risk = $firstEntry.Risk
                Remediation = $firstEntry.Remediation
            }
        }
        
        if ($FinalReport.Count -eq 0) {
            Write-ADReportLog -Message "No explicit DCSync rights found. This might be normal if only default permissions exist." -Type Info -Terminal
        } else {
            Write-ADReportLog -Message "DCSync rights analysis completed. $($FinalReport.Count) identities with DCSync capabilities found." -Type Info -Terminal
        }
        
        return $FinalReport
        
    } catch {
        $ErrorMessage = "Error analyzing DCSync rights: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-SchemaAdminPaths {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing Schema Admin access paths..." -Type Info -Terminal
        
        $SchemaAdminPaths = @()
        
        # Get Schema Admins group members - try both English and German names
        $SchemaAdmins = $null
        $SchemaAdminsGroup = Get-ADGroupByNames -GroupNames $Global:ADGroupNames.SchemaAdmins
        
        if ($SchemaAdminsGroup) {
            Write-ADReportLog -Message "Found Schema Admins group as: $($SchemaAdminsGroup.Name)" -Type Info -Terminal
            $SchemaAdmins = Get-ADGroupMember -Identity $SchemaAdminsGroup -Recursive -ErrorAction SilentlyContinue
        }
        
        if (-not $SchemaAdmins) {
            Write-ADReportLog -Message "Schema Admins group not found. This might be normal if the group doesn't exist." -Type Info -Terminal
            return @([PSCustomObject]@{
                Name = "N/A"
                SamAccountName = "N/A"
                ObjectType = "Information"
                Status = "Schema Admins group not found"
                Details = "The Schema Admins group might not exist or has a different name"
                Remediation = "Verify Schema Admins group existence"
            })
        }
        
        foreach ($admin in $SchemaAdmins) {
            $userDetails = Get-ADUser -Identity $admin.SamAccountName -Properties Enabled, LastLogonDate, PasswordLastSet, AdminCount, whenCreated -ErrorAction SilentlyContinue
            
            if ($userDetails) {
                $SchemaAdminPaths += [PSCustomObject]@{
                    Name = $admin.Name
                    SamAccountName = $admin.SamAccountName
                    ObjectType = $admin.ObjectClass
                    Enabled = $userDetails.Enabled
                    LastLogonDate = $userDetails.LastLogonDate
                    PasswordLastSet = $userDetails.PasswordLastSet
                    AccountAge = if ($userDetails.whenCreated) { [math]::Round((New-TimeSpan -Start $userDetails.whenCreated -End (Get-Date)).TotalDays) } else { "Unknown" }
                    Risk = if ($userDetails.Enabled -and $userDetails.LastLogonDate -gt (Get-Date).AddDays(-90)) { "Active Schema Admin" } else { "Inactive Schema Admin" }
                    Remediation = "Schema Admins should be empty except during schema modifications"
                }
            }
        }
        
        # Check for users who can add themselves to Schema Admins
        $EnterpriseAdmins = $null
        $EnterpriseAdminsGroup = Get-ADGroupByNames -GroupNames $Global:ADGroupNames.EnterpriseAdmins
        
        if ($EnterpriseAdminsGroup) {
            $EnterpriseAdmins = Get-ADGroupMember -Identity $EnterpriseAdminsGroup -Recursive -ErrorAction SilentlyContinue
        }
        
        if ($EnterpriseAdmins) {
            foreach ($ea in $EnterpriseAdmins) {
                # Check if not already in Schema Admins (if Schema Admins exists)
                $isAlreadySchemaAdmin = $false
                if ($SchemaAdmins) {
                    $isAlreadySchemaAdmin = $SchemaAdmins.SamAccountName -contains $ea.SamAccountName
                }
                
                if (-not $isAlreadySchemaAdmin) {
                    $userDetails = Get-ADUser -Identity $ea.SamAccountName -Properties Enabled, LastLogonDate -ErrorAction SilentlyContinue
                    
                    if ($userDetails) {
                        $SchemaAdminPaths += [PSCustomObject]@{
                            Name = $ea.Name
                            SamAccountName = $ea.SamAccountName
                            ObjectType = "Potential Path"
                            Enabled = $userDetails.Enabled
                            LastLogonDate = $userDetails.LastLogonDate
                            Risk = "Can elevate to Schema Admin"
                            Remediation = "Monitor Enterprise Admin membership"
                        }
                    }
                }
            }
        }
        
        Write-ADReportLog -Message "Schema Admin paths analysis completed. $($SchemaAdminPaths.Count) paths found." -Type Info -Terminal
        return $SchemaAdminPaths
        
    } catch {
        $ErrorMessage = "Error analyzing Schema Admin paths: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-CertificateSecurityAnalysis {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing Certificate Services security..." -Type Info -Terminal
        
        $CertAnalysis = @()
        
        # Check if AD CS is installed
        try {
            $Domain = Get-ADDomain
            $ConfigDN = "CN=Configuration,$($Domain.DistinguishedName)"
            $CertConfig = Get-ADObject -Filter "objectClass -eq 'pKIEnrollmentService'" -SearchBase $ConfigDN -Properties * -ErrorAction SilentlyContinue
        } catch {
            $CertConfig = $null
        }
        
        if (-not $CertConfig) {
            Write-ADReportLog -Message "No Certificate Services found in this domain." -Type Info -Terminal
            return @([PSCustomObject]@{
                Component = "AD CS Status"
                Finding = "Not Installed"
                Risk = "N/A"
                Details = "Active Directory Certificate Services not detected in this domain"
            })
        }
        
        # Analyze certificate templates
        $CertTemplates = Get-ADObject -Filter "objectClass -eq 'pKICertificateTemplate'" -Properties * -SearchBase "CN=Certificate Templates,CN=Public Key Services,CN=Services,CN=Configuration,$((Get-ADDomain).DistinguishedName)" -ErrorAction SilentlyContinue
        
        foreach ($template in $CertTemplates) {
            # Check for dangerous configurations
            $risks = @()
            $riskLevel = "Low"
            
            # Check if template allows requestor to specify SAN
            if ($template.'msPKI-Certificate-Name-Flag' -band 1) {
                $risks += "Allows SAN specification"
                $riskLevel = "High"
            }
            
            # Check for client authentication
            if ($template.'msPKI-Certificate-Application-Policy' -contains "1.3.6.1.5.5.7.3.2") {
                $risks += "Allows client authentication"
                if ($riskLevel -ne "High") { $riskLevel = "Medium" }
            }
            
            # Check enrollment permissions
            $templateACL = Get-Acl "AD:$($template.DistinguishedName)" -ErrorAction SilentlyContinue
            $enrollmentRights = $templateACL.Access | Where-Object { $_.ActiveDirectoryRights -match "ExtendedRight" }
            
            if ($enrollmentRights) {
                $enrollers = $enrollmentRights.IdentityReference.Value -join ", "
                if ($enrollers -match "Authenticated Users|Domain Users") {
                    $risks += "Wide enrollment permissions"
                    $riskLevel = "Critical"
                }
            }
            
            if ($risks.Count -gt 0) {
                $CertAnalysis += [PSCustomObject]@{
                    TemplateName = $template.Name
                    DisplayName = $template.DisplayName
                    Risks = $risks -join "; "
                    RiskLevel = $riskLevel
                    EnrollmentRights = $enrollers
                    Remediation = "Review template permissions and configuration"
                }
            }
        }
        
        # Check for web enrollment
        $WebEnrollment = Get-ADObject -Filter "objectClass -eq 'certificationAuthority'" -Properties * -ErrorAction SilentlyContinue
        
        if ($WebEnrollment) {
            $CertAnalysis += [PSCustomObject]@{
                Component = "Web Enrollment"
                Status = "Detected"
                Risk = "Medium"
                Details = "Web enrollment increases attack surface"
                Remediation = "Ensure web enrollment is properly secured with HTTPS and authentication"
            }
        }
        
        Write-ADReportLog -Message "Certificate security analysis completed. $($CertAnalysis.Count) findings." -Type Info -Terminal
        return $CertAnalysis
        
    } catch {
        $ErrorMessage = "Error analyzing Certificate Services: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Performance & Health Monitoring ---
Function Get-SYSVOLHealthCheck {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Performing SYSVOL health check..." -Type Info -Terminal
        
        $SYSVOLHealth = @()
        $DomainControllers = Get-ADDomainController -Filter * -ErrorAction Stop
        
        foreach ($DC in $DomainControllers) {
            try {
                # Test SYSVOL share accessibility
                $SYSVOLPath = "\\$($DC.HostName)\SYSVOL"
                $SYSVOLAccessible = Test-Path $SYSVOLPath -ErrorAction SilentlyContinue
                
                # Check DFSR or FRS
                $ReplicationMethod = "Unknown"
                try {
                    $DFSRCheck = Get-WmiObject -ComputerName $DC.HostName -Namespace "root\microsoftdfs" -Class "dfsrreplicatedfolderinfo" -ErrorAction SilentlyContinue
                    if ($DFSRCheck) {
                        $ReplicationMethod = "DFSR"
                    } else {
                        $FRSCheck = Get-WmiObject -ComputerName $DC.HostName -Class "Win32_Service" -Filter "Name='NtFrs'" -ErrorAction SilentlyContinue
                        if ($FRSCheck -and $FRSCheck.State -eq "Running") {
                            $ReplicationMethod = "FRS (Legacy)"
                        }
                    }
                } catch {
                    $ReplicationMethod = "Unable to determine"
                }
                
                # Check for orphaned GPO folders
                $orphanedGPOs = 0
                if ($SYSVOLAccessible) {
                    try {
                        $GPOFolders = Get-ChildItem "$SYSVOLPath\$($DC.Domain)\Policies" -Directory -ErrorAction SilentlyContinue
                        $ADGPOs = Get-GPO -All -Domain $DC.Domain -ErrorAction SilentlyContinue
                        $ADGPOIds = $ADGPOs | ForEach-Object { "{$($_.Id)}" }
                        
                        foreach ($folder in $GPOFolders) {
                            if ($folder.Name -ne "PolicyDefinitions" -and $ADGPOIds -notcontains $folder.Name) {
                                $orphanedGPOs++
                            }
                        }
                    } catch {
                        $orphanedGPOs = "Unable to check"
                    }
                }
                
                $SYSVOLHealth += [PSCustomObject]@{
                    DomainController = $DC.Name
                    SYSVOLAccessible = $SYSVOLAccessible
                    ReplicationMethod = $ReplicationMethod
                    OrphanedGPOFolders = $orphanedGPOs
                    NetlogonShare = Test-Path "\\$($DC.HostName)\NETLOGON" -ErrorAction SilentlyContinue
                    Status = if ($SYSVOLAccessible -and $ReplicationMethod -ne "Unknown") { "Healthy" } else { "Issues Detected" }
                    Remediation = if ($ReplicationMethod -eq "FRS (Legacy)") { "Migrate from FRS to DFSR" } 
                                 elseif ($orphanedGPOs -gt 0) { "Clean up orphaned GPO folders" }
                                 else { "None required" }
                }
                
            } catch {
                $SYSVOLHealth += [PSCustomObject]@{
                    DomainController = $DC.Name
                    SYSVOLAccessible = $false
                    Status = "Error"
                    Error = $_.Exception.Message
                }
            }
        }
        
        Write-ADReportLog -Message "SYSVOL health check completed for $($SYSVOLHealth.Count) domain controllers." -Type Info -Terminal
        return $SYSVOLHealth
        
    } catch {
        $ErrorMessage = "Error performing SYSVOL health check: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-DNSHealthAnalysis {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing DNS health..." -Type Info -Terminal
        
        $DNSHealth = @()
        $Domain = Get-ADDomain
        
        # Check DNS zones - only if DNS cmdlets are available
        $DNSZones = $null
        try {
            if (Get-Command Get-DnsServerZone -ErrorAction SilentlyContinue) {
                $DNSZones = Get-DnsServerZone -ComputerName $Domain.PDCEmulator -ErrorAction SilentlyContinue
            } else {
                Write-ADReportLog -Message "DNS Server cmdlets not available. Skipping DNS zone analysis." -Type Info -Terminal
            }
        } catch {
            Write-ADReportLog -Message "Could not query DNS zones: $($_.Exception.Message)" -Type Warning -Terminal
        }
        
        if ($DNSZones) {
            # Group zones by type for tree structure
            $forwardZones = $DNSZones | Where-Object { -not $_.IsReverseLookupZone }
            $reverseZones = $DNSZones | Where-Object { $_.IsReverseLookupZone }
            
            # Process Forward Zones
            foreach ($zone in $forwardZones) {
                $issuesList = [System.Collections.ArrayList]::new()
                
                # Security check
                if ($zone.DynamicUpdate -eq "NonsecureAndSecure") {
                    $issuesList.Add("Allows non-secure dynamic updates") | Out-Null
                    $status = "Security Risk"
                } else {
                    $status = "Healthy"
                }
                
                # Aging/Scavenging check
                if ($zone.AgingEnabled) {
                    $scavengingEnabled = $true
                    $noRefreshInterval = $zone.NoRefreshInterval
                    $refreshInterval = $zone.RefreshInterval
                } else {
                    $issuesList.Add("Scavenging not enabled") | Out-Null
                    if ($status -eq "Healthy") { $status = "Warning" }
                    $scavengingEnabled = $false
                }
                
                $DNSHealth += [PSCustomObject]@{
                    Category = "Forward Zones"
                    ZoneName = $zone.ZoneName
                    Type = "Forward Lookup"
                    ZoneType = $zone.ZoneType
                    IsDsIntegrated = $zone.IsDsIntegrated
                    DynamicUpdate = $zone.DynamicUpdate
                    Status = $status
                    Issues = if ($issuesList.Count -gt 0) { $issuesList -join "; " } else { "None" }
                    ScavengingEnabled = $scavengingEnabled
                    NoRefreshInterval = if ($scavengingEnabled) { $noRefreshInterval } else { "N/A" }
                    RefreshInterval = if ($scavengingEnabled) { $refreshInterval } else { "N/A" }
                }
            }
            
            # Process Reverse Zones
            foreach ($zone in $reverseZones) {
                $issuesList = [System.Collections.ArrayList]::new()
                
                if ($zone.DynamicUpdate -eq "NonsecureAndSecure") {
                    $issuesList.Add("Allows non-secure dynamic updates") | Out-Null
                    $status = "Security Risk"
                } else {
                    $status = "Healthy"
                }
                
                if ($zone.AgingEnabled) {
                    $scavengingEnabled = $true
                    $noRefreshInterval = $zone.NoRefreshInterval
                    $refreshInterval = $zone.RefreshInterval
                } else {
                    $issuesList.Add("Scavenging not enabled") | Out-Null
                    if ($status -eq "Healthy") { $status = "Warning" }
                    $scavengingEnabled = $false
                }
                
                $DNSHealth += [PSCustomObject]@{
                    Category = "Reverse Zones"
                    ZoneName = $zone.ZoneName
                    Type = "Reverse Lookup"
                    ZoneType = $zone.ZoneType
                    IsDsIntegrated = $zone.IsDsIntegrated
                    DynamicUpdate = $zone.DynamicUpdate
                    Status = $status
                    Issues = if ($issuesList.Count -gt 0) { $issuesList -join "; " } else { "None" }
                    ScavengingEnabled = $scavengingEnabled
                    NoRefreshInterval = if ($scavengingEnabled) { $noRefreshInterval } else { "N/A" }
                    RefreshInterval = if ($scavengingEnabled) { $refreshInterval } else { "N/A" }
                }
            }
        }
        
        # Add DNS Health Summary
        $DNSHealth += [PSCustomObject]@{
            Category = "DNS Health Summary"
            ZoneName = "Overall Status"
            Type = "System Check"
            ZoneType = "N/A"
            IsDsIntegrated = "N/A" 
            DynamicUpdate = "N/A"
            Status = if ($DNSHealth.Status -contains "Security Risk") { "Critical" } 
                    elseif ($DNSHealth.Status -contains "Warning") { "Warning" }
                    else { "Healthy" }
            Issues = "Stale records check required"
            ScavengingEnabled = "N/A"
            NoRefreshInterval = "N/A"
            RefreshInterval = "N/A"
        }
        
        Write-ADReportLog -Message "DNS health analysis completed. $($DNSHealth.Count) items analyzed." -Type Info -Terminal
        return $DNSHealth
        
    } catch {
        $ErrorMessage = "Error analyzing DNS health: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Backup & Recovery Readiness ---
Function Get-BackupReadinessStatus {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Checking AD backup readiness..." -Type Info -Terminal
        
        $BackupStatus = @()
        $Domain = Get-ADDomain
        
        # Check System State backup on PDC
        $PDC = $Domain.PDCEmulator
        
        # Tombstone lifetime
        $ConfigNC = "CN=Configuration,$($Domain.DistinguishedName)"
        $TombstoneObject = Get-ADObject -Identity "CN=Directory Service,CN=Windows NT,CN=Services,$ConfigNC" -Properties tombstoneLifetime -ErrorAction SilentlyContinue
        $TombstoneLifetime = if ($TombstoneObject.tombstoneLifetime) { $TombstoneObject.tombstoneLifetime } else { 180 }
        
        $BackupStatus += [PSCustomObject]@{
            Component = "Tombstone Lifetime"
            Value = "$TombstoneLifetime days"
            Status = if ($TombstoneLifetime -ge 180) { "Good" } else { "Warning" }
            Details = "Backups older than tombstone lifetime cannot be restored"
            Remediation = if ($TombstoneLifetime -lt 180) { "Consider increasing tombstone lifetime" } else { "None required" }
        }
        
        # Check Deleted Objects container size
        try {
            $DeletedObjects = Get-ADObject -Filter * -IncludeDeletedObjects -SearchBase "CN=Deleted Objects,$($Domain.DistinguishedName)" -ErrorAction SilentlyContinue
            $DeletedCount = if ($DeletedObjects) { @($DeletedObjects).Count } else { 0 }
            
            $BackupStatus += [PSCustomObject]@{
                Component = "Deleted Objects"
                Value = "$DeletedCount objects"
                Status = if ($DeletedCount -gt 10000) { "Warning" } else { "Good" }
                Details = "Large number of deleted objects can impact restore performance"
                Remediation = if ($DeletedCount -gt 10000) { "Consider garbage collection" } else { "None required" }
            }
        } catch {
            $BackupStatus += [PSCustomObject]@{
                Component = "Deleted Objects"
                Value = "Unable to access"
                Status = "Unknown"
                Details = "Requires elevated permissions"
            }
        }
        
        # DSRM password age check
        $DomainControllers = Get-ADDomainController -Filter *
        
        foreach ($DC in $DomainControllers) {
            $BackupStatus += [PSCustomObject]@{
                Component = "DSRM Password"
                DomainController = $DC.Name
                Status = "Manual Check Required"
                Details = "DSRM password age cannot be checked remotely"
                Remediation = "Ensure DSRM passwords are documented and regularly updated"
            }
        }
        
        # Backup GPO existence
        $BackupGPOs = Get-GPO -All | Where-Object { $_.DisplayName -like "*Backup*" -or $_.DisplayName -like "*Restore*" }
        
        $BackupStatus += [PSCustomObject]@{
            Component = "Backup Procedures"
            Value = if ($BackupGPOs) { "$($BackupGPOs.Count) backup-related GPOs found" } else { "No backup GPOs found" }
            Status = if ($BackupGPOs) { "Good" } else { "Warning" }
            Details = "GPOs can help standardize backup procedures"
            Remediation = if (-not $BackupGPOs) { "Consider creating backup procedure GPOs" } else { "Review existing backup GPOs" }
        }
        
        Write-ADReportLog -Message "Backup readiness check completed." -Type Info -Terminal
        return $BackupStatus
        
    } catch {
        $ErrorMessage = "Error checking backup readiness: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- Schema & Trusts Analysis ---
Function Get-SchemaAnalysis {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing AD schema..." -Type Info -Terminal
        
        $SchemaAnalysis = @()
        $RootDSE = Get-ADRootDSE
        $SchemaNC = $RootDSE.schemaNamingContext
        
        # Get schema version and details
        $Schema = Get-ADObject -Identity $SchemaNC -Properties objectVersion, msDS-Behavior-Version
        
        # Alle möglichen Schema-Versionen definieren
        $AllSchemaVersions = @{
            30 = @{
                Name = "Windows Server 2003"
                Details = "Base Windows Server 2003 Schema"
            }
            31 = @{
                Name = "Windows Server 2003 R2" 
                Details = "Supports Read-Only Domain Controllers"
            }
            44 = @{
                Name = "Windows Server 2008"
                Details = "Supports Fine-Grained Password Policies"
            }
            47 = @{
                Name = "Windows Server 2008 R2"
                Details = "Supports Managed Service Accounts, Authentication Mechanism Assurance"
            }
            56 = @{
                Name = "Windows Server 2012"
                Details = "Supports Group Managed Service Accounts, Kerberos Armoring"
            }
            69 = @{
                Name = "Windows Server 2012 R2"
                Details = "Supports Dynamic Access Control, Kerberos KDC Support for Claims"
            }
            87 = @{
                Name = "Windows Server 2016"
                Details = "Supports Privileged Access Management, Microsoft Passport for Work"
            }
            88 = @{
                Name = "Windows Server 2019/2022"
                Details = "Supports FIDO2 Authentication, Group MSA Enhancements"
            }
            91 = @{
                Name = "Windows Server 2025"
                Details = "Supports Enhanced Security Features, Latest AD Improvements"
            }
        }

        # Nur die aktive Schema-Version anzeigen
        if ($AllSchemaVersions.ContainsKey($Schema.objectVersion)) {
            $SchemaAnalysis += [PSCustomObject]@{
                Name = $AllSchemaVersions[$Schema.objectVersion].Name
                Details = $AllSchemaVersions[$Schema.objectVersion].Details
                Status = "Aktiv"
                IsCurrent = $true
            }
        }
        
        # Count custom schema extensions
        $AllSchemaObjects = Get-ADObject -Filter * -SearchBase $SchemaNC -Properties whenCreated, adminDescription
        $CustomSchema = $AllSchemaObjects | Where-Object { 
            $_.whenCreated -and 
            $_.adminDescription -notlike "Microsoft*" -and 
            $_.Name -notlike "ms-*" 
        }
        
        $SchemaAnalysis += [PSCustomObject]@{
            Name = "Benutzerdefinierte Erweiterungen"
            Details = "$($CustomSchema.Count) benutzerdefinierte Schema-Objekte gefunden"
            Status = "Info"
            IsCurrent = $true
        }
        
        # Recent schema changes
        $RecentChanges = $AllSchemaObjects | Where-Object { 
            $_.whenCreated -gt (Get-Date).AddDays(-90) 
        }
        
        if ($RecentChanges) {
            foreach ($change in $RecentChanges) {
                $SchemaAnalysis += [PSCustomObject]@{
                    Name = "Kürzliche Änderung: $($change.Name)"
                    Details = "Schema-Änderung am $($change.whenCreated.ToString('dd.MM.yyyy'))"
                    Status = "Info"
                    IsCurrent = $true
                }
            }
        }
        
        Write-ADReportLog -Message "Schema analysis completed." -Type Info -Terminal
        return $SchemaAnalysis
        
    } catch {
        $ErrorMessage = "Error in schema analysis: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-TrustRelationshipAnalysis {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing trust relationships..." -Type Info -Terminal
        
        $TrustAnalysis = @()
        
        # Get all trusts
        $Trusts = Get-ADTrust -Filter * -Properties * -ErrorAction Stop
        
        foreach ($trust in $Trusts) {
            # Analyze trust properties
            $trustHealth = "Unknown"
            $issues = @()
            
            # Check trust type and transitivity
            if ($trust.TrustType -eq "External" -and $trust.Transitivity -eq "Transitive") {
                $issues += "External trust is transitive"
                $trustHealth = "Security Risk"
            }
            
            # Check SID filtering
            if ($trust.SIDFilteringForestAware -eq $false -or $trust.SIDFilteringQuarantined -eq $false) {
                $issues += "SID filtering may be disabled"
                if ($trustHealth -ne "Security Risk") { $trustHealth = "Warning" }
            }
            
            # Check selective authentication
            if ($trust.SelectiveAuthentication -eq $false -and $trust.TrustType -eq "External") {
                $issues += "Selective authentication not enabled"
                if ($trustHealth -eq "Unknown") { $trustHealth = "Review Needed" }
            }
            
            if ($trustHealth -eq "Unknown" -and $issues.Count -eq 0) { $trustHealth = "Healthy" }
            
            $TrustAnalysis += [PSCustomObject]@{
                TrustPartner = $trust.Target
                Direction = $trust.Direction
                TrustType = $trust.TrustType
                Transitivity = if ($trust.Transitivity) { "Transitive" } else { "Non-Transitive" }
                SIDFiltering = if ($trust.SIDFilteringQuarantined) { "Enabled" } else { "Check Required" }
                SelectiveAuth = if ($trust.SelectiveAuthentication) { "Enabled" } else { "Disabled" }
                Created = $trust.Created
                Status = $trustHealth
                Issues = $issues -join "; "
                Remediation = if ($issues.Count -gt 0) { "Review trust configuration for security best practices" } else { "None required" }
            }
        }
        
        # Forest trust insights
        $Forest = Get-ADForest
        if ($Forest.Domains.Count -gt 1) {
            $TrustAnalysis += [PSCustomObject]@{
                Component = "Forest Structure"
                Domains = $Forest.Domains.Count
                GlobalCatalogs = $Forest.GlobalCatalogs.Count
                Sites = $Forest.Sites.Count
                Status = "Multi-Domain Forest"
                Details = "Complex forest structure requires careful trust management"
            }
        }
        
        Write-ADReportLog -Message "Trust relationship analysis completed. $($Trusts.Count) trusts analyzed." -Type Info -Terminal
        return $TrustAnalysis
        
    } catch {
        $ErrorMessage = "Error analyzing trust relationships: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-QuotasAndLimits {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing AD quotas and limits..." -Type Info -Terminal
        
        $QuotaAnalysis = @()
        $Domain = Get-ADDomain
        $Forest = Get-ADForest
        
        # RID Pool check
        $DomainControllers = Get-ADDomainController -Filter *
        foreach ($DC in $DomainControllers) {
            try {
                $RIDInfo = Get-ADObject -Identity "CN=RID Manager$,CN=System,$($Domain.DistinguishedName)" -Properties rIDAvailablePool -Server $DC.HostName -ErrorAction SilentlyContinue
                
                if ($RIDInfo -and $RIDInfo.rIDAvailablePool) {
                    try {
                        if ($RIDInfo.rIDAvailablePool -is [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
                            $ridPoolValue = [int64]$RIDInfo.rIDAvailablePool[0]
                        } else {
                            $ridPoolValue = [int64]$RIDInfo.rIDAvailablePool
                        }
                        
                        [int64]$pow32 = [math]::Pow(2, 32)
                        [int64]$pow30 = [math]::Pow(2, 30)
                        
                        [int64]$totalSIDS = [math]::Floor($ridPoolValue / $pow32)
                        [int64]$totalRIDS = $totalSIDS * $pow30
                        [int64]$currentRIDPoolCount = $ridPoolValue % $pow32
                        
                        $percentUsed = if ($totalRIDS -gt 0) {
                            [math]::Round((($totalRIDS - $currentRIDPoolCount) / $totalRIDS * 100), 2)
                        } else { 0 }
                        
                        $QuotaAnalysis += [PSCustomObject]@{
                            Component = "RID Pool"
                            DomainController = $DC.Name
                            TotalRIDs = $totalRIDS
                            RemainingRIDs = $currentRIDPoolCount
                            PercentUsed = "$percentUsed%"
                            Status = if ($percentUsed -gt 80) { "Critical" } elseif ($percentUsed -gt 60) { "Warning" } else { "Healthy" }
                            Remediation = if ($percentUsed -gt 80) { "RID pool nearly exhausted - plan RID recovery" } else { "Monitor RID usage" }
                        }
                    } catch {
                        Write-ADReportLog -Message "Error calculating RID pool for $($DC.Name): $($_.Exception.Message)" -Type Warning
                        continue
                    }
                }
            } catch {
                Write-ADReportLog -Message "Could not retrieve RID pool information from $($DC.Name)" -Type Warning
            }
        }

        # LDAP Query Limits
        $DefaultPolicy = Get-ADObject -Identity "CN=Default Query Policy,CN=Query-Policies,CN=Directory Service,CN=Windows NT,CN=Services,CN=Configuration,$($Domain.DistinguishedName)" -Properties * -ErrorAction SilentlyContinue
        
        if ($DefaultPolicy) {
            $MaxPageSize = ($DefaultPolicy.lDAPAdminLimits | Where-Object { $_ -like "MaxPageSize=*" } | ForEach-Object { $_.Split('=')[1] }) -as [int]
            $MaxQueryDuration = ($DefaultPolicy.lDAPAdminLimits | Where-Object { $_ -like "MaxQueryDuration=*" } | ForEach-Object { $_.Split('=')[1] }) -as [int]
            $MaxResults = ($DefaultPolicy.lDAPAdminLimits | Where-Object { $_ -like "MaxResultSetSize=*" } | ForEach-Object { $_.Split('=')[1] }) -as [int]
            
            $QuotaAnalysis += [PSCustomObject]@{
                Component = "LDAP Query Limits"
                MaxPageSize = $MaxPageSize
                MaxQueryDuration = "$MaxQueryDuration seconds"
                MaxResults = $MaxResults
                Status = if ($MaxPageSize -lt 1000 -or $MaxQueryDuration -lt 120) { "Review" } else { "OK" }
                Remediation = "Adjust LDAP limits if performance issues occur"
            }
        }

        # Kerberos Token Size
        $LargeTokenUsers = Get-ADUser -Filter * -Properties MemberOf | Where-Object { $_.MemberOf.Count -gt 100 }
        $MaxTokenSize = 65535 # Maximum token size in bytes
        
        if ($LargeTokenUsers) {
            $QuotaAnalysis += [PSCustomObject]@{
                Component = "Kerberos Token Size"
                UsersAtRisk = $LargeTokenUsers.Count
                MaxTokenSize = "$MaxTokenSize bytes"
                Status = if ($LargeTokenUsers.Count -gt 10) { "Warning" } else { "Review" }
                Remediation = "Review group memberships for users with many groups"
            }
        }

        # Tombstone Lifetime
        $TombstoneLifetime = Get-ADObject -Identity "CN=Directory Service,CN=Windows NT,CN=Services,CN=Configuration,$($Domain.DistinguishedName)" -Properties tombstoneLifetime
        $TombstoneValue = if ($TombstoneLifetime.tombstoneLifetime) { $TombstoneLifetime.tombstoneLifetime } else { 60 }
        
        $QuotaAnalysis += [PSCustomObject]@{
            Component = "Tombstone Lifetime"
            Value = "$TombstoneValue days"
            Status = if ($TombstoneValue -lt 60) { "Warning" } else { "OK" }
            Remediation = if ($TombstoneValue -lt 60) { "Increase tombstone lifetime to at least 60 days" } else { "No action required" }
        }

        # Replication Limits
        $SiteLinks = Get-ADReplicationSiteLink -Filter *
        $AverageInterval = 0
        
        if ($SiteLinks.Count -gt 0) {
            $AverageInterval = ($SiteLinks | Measure-Object -Property ReplicationInterval -Average).Average
        }
        
        $QuotaAnalysis += [PSCustomObject]@{
            Component = "Replication Configuration"
            TotalLinks = $SiteLinks.Count
            AvgReplicationInterval = "$([math]::Round($AverageInterval, 2)) minutes"
            Status = if ($AverageInterval -gt 180) { "Review" } else { "OK" }
            Remediation = if ($AverageInterval -gt 180) { "Long replication intervals may cause inconsistencies" } else { "No action required" }
        }

        # Domain Limits
        $QuotaAnalysis += [PSCustomObject]@{
            Component = "Domain Structure"
            ForestDomains = $Forest.Domains.Count
            GlobalCatalogs = $Forest.GlobalCatalogs.Count
            Sites = $Forest.Sites.Count
            Status = if ($Forest.Domains.Count -gt 5) { "Complex" } else { "Standard" }
            Remediation = if ($Forest.Domains.Count -gt 5) { "Complex forest structure requires careful management" } else { "No action required" }
        }

        # Password Policies
        $PasswordPolicy = Get-ADDefaultDomainPasswordPolicy
        
        $QuotaAnalysis += [PSCustomObject]@{
            Component = "Password Policies"
            MaxPasswordAge = "$($PasswordPolicy.MaxPasswordAge.Days) days"
            MinPasswordLength = $PasswordPolicy.MinPasswordLength
            PasswordHistory = $PasswordPolicy.PasswordHistoryCount
            Status = if ($PasswordPolicy.MinPasswordLength -lt 12) { "Review" } else { "OK" }
            Remediation = if ($PasswordPolicy.MinPasswordLength -lt 12) { "Increase minimum password length for better security" } else { "No action required" }
        }

        Write-ADReportLog -Message "Quota and limit analysis completed." -Type Info -Terminal
        return $QuotaAnalysis
        
    } catch {
        $ErrorMessage = "Error analyzing quotas and limits: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

# --- AD-Health Funktionen ---
Function Get-FSMORoles {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Retrieving FSMO role holders..." -Type Info -Terminal

        # Load localized messages
        $msgTable = data {
            #culture="en-US" 
            ConvertFrom-StringData @'
            ForestWide = Forest-wide
            DomainSpecific = Domain-specific
            Online = Online
            Offline = Offline
            SchemaDesc = Manages the Active Directory schema
            DomainNamingDesc = Manages adding and removing domains
            PDCDesc = Time synchronization and password changes
            RIDDesc = Distributes RID pools to domain controllers
            InfraDesc = Manages cross-domain references
'@
        }

        # Import localized data if available
        try {
            Import-LocalizedData -BindingVariable msgTable
        } catch {
            Write-ADReportLog -Message "Using default English messages - no localization found" -Type Info
        }
        
        $Forest = Get-ADForest
        $Domain = Get-ADDomain
        
        $FSMORoles = @()
        
        # Forest-wide FSMO roles
        $FSMORoles += [PSCustomObject]@{
            Role = "Schema Master"
            Type = $msgTable.ForestWide
            Server = $Forest.SchemaMaster
            Domain = $Forest.Name
            Status = if (Test-Connection -ComputerName $Forest.SchemaMaster -Count 1 -Quiet) { $msgTable.Online } else { $msgTable.Offline }
            Description = $msgTable.SchemaDesc
        }

        $FSMORoles += [PSCustomObject]@{
            Role = "Domain Naming Master"
            Type = $msgTable.ForestWide
            Server = $Forest.DomainNamingMaster
            Domain = $Forest.Name
            Status = if (Test-Connection -ComputerName $Forest.DomainNamingMaster -Count 1 -Quiet) { $msgTable.Online } else { $msgTable.Offline }
            Description = $msgTable.DomainNamingDesc
        }
        
        # Domain-specific FSMO roles
        $FSMORoles += [PSCustomObject]@{
            Role = "PDC Emulator"
            Type = $msgTable.DomainSpecific
            Server = $Domain.PDCEmulator
            Domain = $Domain.Name
            Status = if (Test-Connection -ComputerName $Domain.PDCEmulator -Count 1 -Quiet) { $msgTable.Online } else { $msgTable.Offline }
            Description = $msgTable.PDCDesc
        }

        $FSMORoles += [PSCustomObject]@{
            Role = "RID Master"
            Type = $msgTable.DomainSpecific
            Server = $Domain.RIDMaster
            Domain = $Domain.Name
            Status = if (Test-Connection -ComputerName $Domain.RIDMaster -Count 1 -Quiet) { $msgTable.Online } else { $msgTable.Offline }
            Description = $msgTable.RIDDesc
        }

        $FSMORoles += [PSCustomObject]@{
            Role = "Infrastructure Master"
            Type = $msgTable.DomainSpecific
            Server = $Domain.InfrastructureMaster
            Domain = $Domain.Name
            Status = if (Test-Connection -ComputerName $Domain.InfrastructureMaster -Count 1 -Quiet) { $msgTable.Online } else { $msgTable.Offline }
            Description = $msgTable.InfraDesc
        }
        
        Write-ADReportLog -Message "$($FSMORoles.Count) FSMO roles found." -Type Info -Terminal
        return $FSMORoles
        
    } catch {
        $ErrorMessage = "Error retrieving FSMO roles: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return $null
    }
}

Function Get-DomainControllerStatus {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Collecting AD Health Data..." -Type Info -Terminal
        
        # Initialize AD Health Report array
        $ADHealthReport = @()
        
        # 1. Forest Information
        try {
            $Forest = Get-ADForest
            $ADHealthReport += [PSCustomObject]@{
                Category = "Forest Information" 
                Parameter = "Forest Name"
                Value = $Forest.Name
                Status = "OK"
                Details = "Forest Functional Level: $($Forest.ForestMode)"
            }

            # Get Schema Version
            $schemaVersion = "Unknown"
            if ($Forest.SchemaVersion) {
                $schemaVersion = if ($Forest.SchemaVersion -is [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
                    $Forest.SchemaVersion[0]
                } else {
                    $Forest.SchemaVersion.ToString()
                }
            }
            
            $ADHealthReport += [PSCustomObject]@{
                Category = "Forest Information"
                Parameter = "Schema Version"
                Value = $schemaVersion
                Status = "OK" 
                Details = "Current schema version"
            }
        }
        catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "Forest Information"
                Parameter = "Forest Access"
                Value = "Error"
                Status = "Critical"
                Details = $_.Exception.Message
            }
        }

        # 2. Domain Information
        try {
            $Domain = Get-ADDomain
            $ADHealthReport += [PSCustomObject]@{
                Category = "Domain Information"
                Parameter = "Domain Name"
                Value = $Domain.NetBIOSName
                Status = "OK"
                Details = "FQDN: $($Domain.DNSRoot), Level: $($Domain.DomainMode)"
            }

            # PDC Emulator Check
            $pdcStatus = Test-Connection -ComputerName $Domain.PDCEmulator -Count 1 -Quiet -ErrorAction SilentlyContinue
            $ADHealthReport += [PSCustomObject]@{
                Category = "Domain Information"
                Parameter = "PDC Emulator"
                Value = $Domain.PDCEmulator
                Status = if ($pdcStatus) { "OK" } else { "Warning" }
                Details = if ($pdcStatus) { "PDC Emulator is reachable" } else { "PDC Emulator not reachable" }
            }
        }
        catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "Domain Information"
                Parameter = "Domain Access"
                Value = "Error"
                Status = "Critical"
                Details = $_.Exception.Message
            }
        }

        # 3. Domain Controller Health Check
        try {
            $DCs = Get-ADDomainController -Filter * -ErrorAction Stop
            
            $DCCount = @($DCs).Count
            $ADHealthReport += [PSCustomObject]@{
                Category = "Domain Controllers"
                Parameter = "Total DCs"
                Value = $DCCount
                Status = if ($DCCount -ge 2) { "OK" } else { "Warning" }
                Details = "Minimum recommended: 2 DCs for redundancy"
            }

            foreach ($DC in $DCs) {
                # Connectivity Test
                $pingTest = Test-Connection -ComputerName $DC.HostName -Count 1 -Quiet -ErrorAction SilentlyContinue
                $ADHealthReport += [PSCustomObject]@{
                    Category = "Domain Controllers"
                    Parameter = "DC Connectivity"
                    Value = $DC.Name
                    Status = if ($pingTest) { "OK" } else { "Critical" }
                    Details = if ($pingTest) { "DC is responding" } else { "DC not responding" }
                }

                # LDAP Service Check
                $ldapTest = Test-NetConnection -ComputerName $DC.HostName -Port 389 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                $ADHealthReport += [PSCustomObject]@{
                    Category = "Domain Controllers"
                    Parameter = "LDAP Service"
                    Value = "$($DC.Name):389"
                    Status = if ($ldapTest) { "OK" } else { "Critical" }
                    Details = if ($ldapTest) { "LDAP service available" } else { "LDAP service unavailable" }
                }

                # Global Catalog Check
                if ($DC.IsGlobalCatalog) {
                    $gcTest = Test-NetConnection -ComputerName $DC.HostName -Port 3268 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                    $ADHealthReport += [PSCustomObject]@{
                        Category = "Domain Controllers"
                        Parameter = "Global Catalog"
                        Value = "$($DC.Name):3268"
                        Status = if ($gcTest) { "OK" } else { "Warning" }
                        Details = if ($gcTest) { "Global Catalog service available" } else { "Global Catalog service unavailable" }
                    }
                }
            }
        }
        catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "Domain Controllers"
                Parameter = "DC Health Check"
                Value = "Error"
                Status = "Critical"
                Details = $_.Exception.Message
            }
        }

        # 4. FSMO Roles Health Check
        try {
            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO Roles"
                Parameter = "Schema Master"
                Value = $Forest.SchemaMaster
                Status = if (Test-Connection -ComputerName $Forest.SchemaMaster -Count 1 -Quiet -ErrorAction SilentlyContinue) { "OK" } else { "Critical" }
                Details = "Forest-wide FSMO role"
            }

            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO Roles"
                Parameter = "Domain Naming Master"
                Value = $Forest.DomainNamingMaster
                Status = if (Test-Connection -ComputerName $Forest.DomainNamingMaster -Count 1 -Quiet -ErrorAction SilentlyContinue) { "OK" } else { "Critical" }
                Details = "Forest-wide FSMO role"
            }

            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO Roles"
                Parameter = "PDC Emulator"
                Value = $Domain.PDCEmulator
                Status = if (Test-Connection -ComputerName $Domain.PDCEmulator -Count 1 -Quiet -ErrorAction SilentlyContinue) { "OK" } else { "Critical" }
                Details = "Domain-wide role for time sync and password changes"
            }

            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO Roles"
                Parameter = "RID Master"
                Value = $Domain.RIDMaster
                Status = if (Test-Connection -ComputerName $Domain.RIDMaster -Count 1 -Quiet -ErrorAction SilentlyContinue) { "OK" } else { "Critical" }
                Details = "Manages RID pool assignments"
            }

            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO Roles"
                Parameter = "Infrastructure Master"
                Value = $Domain.InfrastructureMaster
                Status = if (Test-Connection -ComputerName $Domain.InfrastructureMaster -Count 1 -Quiet -ErrorAction SilentlyContinue) { "OK" } else { "Critical" }
                Details = "Manages cross-domain object references"
            }
        }
        catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "FSMO Roles"
                Parameter = "FSMO Check"
                Value = "Error"
                Status = "Critical"
                Details = $_.Exception.Message
            }
        }

        # 5. DNS Health Check
        try {
            $dnsCheck = Resolve-DnsName -Name $Domain.DNSRoot -Type A -ErrorAction SilentlyContinue
            $ADHealthReport += [PSCustomObject]@{
                Category = "DNS Health"
                Parameter = "Domain DNS Resolution"
                Value = $Domain.DNSRoot
                Status = if ($dnsCheck) { "OK" } else { "Warning" }
                Details = if ($dnsCheck) { "DNS resolution successful" } else { "DNS resolution failed" }
            }

            # Check SRV Records
            $srvCheck = Resolve-DnsName -Name "_ldap._tcp.$($Domain.DNSRoot)" -Type SRV -ErrorAction SilentlyContinue
            $ADHealthReport += [PSCustomObject]@{
                Category = "DNS Health"
                Parameter = "LDAP SRV Records"
                Value = "_ldap._tcp.$($Domain.DNSRoot)"
                Status = if ($srvCheck) { "OK" } else { "Warning" }
                Details = if ($srvCheck) { "$($srvCheck.Count) SRV records found" } else { "No SRV records found" }
            }
        }
        catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "DNS Health"
                Parameter = "DNS Check"
                Value = "Error"
                Status = "Warning"
                Details = $_.Exception.Message
            }
        }

        # 6. Replication Health
        try {
            $replData = Get-ADReplicationPartnerMetadata -Target (Get-ADDomainController).HostName -Partition (Get-ADDomain).DistinguishedName -ErrorAction SilentlyContinue | Select-Object -First 3
            
            if ($replData) {
                $recentReplCount = ($replData | Where-Object { $_.LastReplicationSuccess -gt (Get-Date).AddHours(-24) }).Count
                $ADHealthReport += [PSCustomObject]@{
                    Category = "Replication"
                    Parameter = "Replication Status"
                    Value = "$recentReplCount/$($replData.Count) Partners"
                    Status = if ($recentReplCount -eq $replData.Count) { "OK" } elseif ($recentReplCount -gt 0) { "Warning" } else { "Critical" }
                    Details = "Successful replication in last 24 hours"
                }
            }
            else {
                $ADHealthReport += [PSCustomObject]@{
                    Category = "Replication"
                    Parameter = "Replication Status"
                    Value = "No Partners"
                    Status = "Warning"
                    Details = "No replication partners found"
                }
            }
        }
        catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "Replication"
                Parameter = "Replication Check"
                Value = "Error"
                Status = "Warning"
                Details = $_.Exception.Message
            }
        }

        # 7. System Information
        try {
            $ADHealthReport += [PSCustomObject]@{
                Category = "System Information"
                Parameter = "Server Name"
                Value = $env:COMPUTERNAME
                Status = "Info"
                Details = "Context: $($env:USERDOMAIN)\$($env:USERNAME)"
            }

            $ADHealthReport += [PSCustomObject]@{
                Category = "System Information"
                Parameter = "PowerShell Version"
                Value = $PSVersionTable.PSVersion.ToString()
                Status = "Info"
                Details = "AD Module: $(if (Get-Module -ListAvailable -Name ActiveDirectory) { 'Available' } else { 'Not Available' })"
            }

            $ADHealthReport += [PSCustomObject]@{
                Category = "System Information"
                Parameter = "System Time"
                Value = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                Status = "Info"
                Details = "Time Zone: $((Get-TimeZone).DisplayName)"
            }
        }
        catch {
            $ADHealthReport += [PSCustomObject]@{
                Category = "System Information"
                Parameter = "System Check"
                Value = "Partial Error"
                Status = "Info"
                Details = $_.Exception.Message
            }
        }

        # Generate Summary
        $criticalCount = ($ADHealthReport | Where-Object { $_.Status -eq "Critical" }).Count
        $warningCount = ($ADHealthReport | Where-Object { $_.Status -eq "Warning" }).Count
        $okCount = ($ADHealthReport | Where-Object { $_.Status -eq "OK" }).Count

        $summary = [PSCustomObject]@{
            Category = "=== SUMMARY ==="
            Parameter = "AD Health Status"
            Value = if ($criticalCount -eq 0 -and $warningCount -eq 0) { "Healthy" } 
                   elseif ($criticalCount -eq 0) { "Minor Issues" } 
                   else { "Critical Issues" }
            Status = if ($criticalCount -eq 0 -and $warningCount -eq 0) { "OK" }
                    elseif ($criticalCount -eq 0) { "Warning" }
                    else { "Critical" }
            Details = "OK: $okCount, Warnings: $warningCount, Critical: $criticalCount"
        }

        # Add summary to beginning of report
        $FinalReport = @($summary) + $ADHealthReport

        Write-ADReportLog -Message "AD Health Check completed: $($FinalReport.Count) checks performed." -Type Info -Terminal
        return $FinalReport

    }
    catch {
        $ErrorMessage = "Critical error during AD Health Check: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return @([PSCustomObject]@{
            Category = "ERROR"
            Parameter = "AD Health Check"
            Value = "Failed"
            Status = "Critical"
            Details = $ErrorMessage
        })
    }
}

Function Get-ReplicationStatus {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Checking AD replication status..." -Type Info -Terminal
        
        $ReplicationData = @()
        $DomainControllers = Get-ADDomainController -Filter *
        
        # 1. Collect replication failures for all DCs
        Write-ADReportLog -Message "Checking replication failures..." -Type Info -Terminal
        foreach ($DC in $DomainControllers) {
            try {
                $ReplicationFailures = Get-ADReplicationFailure -Target $DC.HostName -ErrorAction SilentlyContinue
                
                if ($ReplicationFailures) {
                    foreach ($failure in $ReplicationFailures) {
                        $ReplicationData += [PSCustomObject]@{
                            Category = "Replication Failures"
                            SourceDC = $DC.Name
                            TargetDC = $failure.Partner
                            Partition = $failure.PartitionDN
                            Status = "ERROR"
                            ErrorType = $failure.FailureType
                            ErrorCount = $failure.FailureCount
                            FirstFailure = $failure.FirstFailureTime
                            LastFailure = $failure.LastFailureTime
                            Details = $failure.LastError
                        }
                    }
                } else {
                    # No errors for this DC
                    $ReplicationData += [PSCustomObject]@{
                        Category = "Replication Status"
                        SourceDC = $DC.Name
                        TargetDC = "All Partners"
                        Partition = "All Partitions"
                        Status = "OK"
                        ErrorType = "None"
                        ErrorCount = 0
                        FirstFailure = $null
                        LastFailure = $null
                        Details = "No replication failures found"
                    }
                }
            } catch {
                $ReplicationData += [PSCustomObject]@{
                    Category = "System Error"
                    SourceDC = $DC.Name
                    TargetDC = "N/A"
                    Partition = "N/A"
                    Status = "CRITICAL"
                    ErrorType = "Query Error"
                    ErrorCount = 1
                    FirstFailure = Get-Date
                    LastFailure = Get-Date
                    Details = "Error retrieving replication data: $($_.Exception.Message)"
                }
            }
        }
        
        # 2. Collect replication partner metadata
        Write-ADReportLog -Message "Collecting replication partner metadata..." -Type Info -Terminal
        foreach ($DC in $DomainControllers) {
            try {
                $PartnerMetadata = Get-ADReplicationPartnerMetadata -Target $DC.HostName -ErrorAction SilentlyContinue
                
                if ($PartnerMetadata) {
                    foreach ($partner in $PartnerMetadata) {
                        $timeSinceLastSync = "Unknown"
                        $syncStatus = "Unknown"
                        
                        if ($partner.LastReplicationSuccess) {
                            $timeSince = (Get-Date) - $partner.LastReplicationSuccess
                            $timeSinceLastSync = "$([math]::Round($timeSince.TotalHours, 1)) hours"
                            
                            # Status based on time since last sync
                            if ($timeSince.TotalHours -lt 1) {
                                $syncStatus = "EXCELLENT"
                            } elseif ($timeSince.TotalHours -lt 6) {
                                $syncStatus = "GOOD"
                            } elseif ($timeSince.TotalHours -lt 24) {
                                $syncStatus = "WARNING"
                            } else {
                                $syncStatus = "CRITICAL"
                            }
                        }
                        
                        $ReplicationData += [PSCustomObject]@{
                            Category = "Partner Metadata"
                            SourceDC = $DC.Name
                            TargetDC = $partner.Partner -replace ".*CN=NTDS Settings,CN=([^,]+),.*", '$1'
                            Partition = $partner.Partition
                            Status = $syncStatus
                            ErrorType = if ($partner.ConsecutiveReplicationFailures -gt 0) { "Consecutive Failures" } else { "None" }
                            ErrorCount = $partner.ConsecutiveReplicationFailures
                            FirstFailure = $partner.LastReplicationAttempt
                            LastFailure = $partner.LastReplicationSuccess
                            Details = "Last sync: $timeSinceLastSync ago | USN: $($partner.LastReplicationSuccess)"
                        }
                    }
                }
            } catch {
                Write-Warning "Error retrieving partner metadata for $($DC.Name): $($_.Exception.Message)"
            }
        }
        
        # 3. Additional replication health checks
        Write-ADReportLog -Message "Performing extended replication diagnostics..." -Type Info -Terminal
        
        # DC-to-DC connectivity tests
        foreach ($SourceDC in $DomainControllers) {
            foreach ($TargetDC in $DomainControllers) {
                if ($SourceDC.Name -ne $TargetDC.Name) {
                    try {
                        # Test LDAP connectivity
                        $ldapTest = Test-NetConnection -ComputerName $TargetDC.HostName -Port 389 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                        
                        # Test RPC connectivity (for replication)
                        $rpcTest = Test-NetConnection -ComputerName $TargetDC.HostName -Port 135 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                        
                        $connectionStatus = "OK"
                        $connectionDetails = "LDAP: OK, RPC: OK"
                        
                        if (-not $ldapTest -and -not $rpcTest) {
                            $connectionStatus = "CRITICAL"
                            $connectionDetails = "LDAP: FAILED, RPC: FAILED"
                        } elseif (-not $ldapTest) {
                            $connectionStatus = "WARNING"
                            $connectionDetails = "LDAP: FAILED, RPC: OK"
                        } elseif (-not $rpcTest) {
                            $connectionStatus = "WARNING"
                            $connectionDetails = "LDAP: OK, RPC: FAILED"
                        }
                        
                        $ReplicationData += [PSCustomObject]@{
                            Category = "DC Connectivity"
                            SourceDC = $SourceDC.Name
                            TargetDC = $TargetDC.Name
                            Partition = "Network Connectivity"
                            Status = $connectionStatus
                            ErrorType = if ($connectionStatus -ne "OK") { "Network Issue" } else { "None" }
                            ErrorCount = if ($connectionStatus -eq "CRITICAL") { 2 } elseif ($connectionStatus -eq "WARNING") { 1 } else { 0 }
                            FirstFailure = if ($connectionStatus -ne "OK") { Get-Date } else { $null }
                            LastFailure = if ($connectionStatus -ne "OK") { Get-Date } else { $null }
                            Details = $connectionDetails
                        }
                    } catch {
                        Write-Warning "Error testing connectivity between $($SourceDC.Name) and $($TargetDC.Name)"
                    }
                }
            }
        }
        
        # 4. Create replication summary
        $totalErrors = ($ReplicationData | Where-Object { $_.Status -eq "CRITICAL" -or $_.Status -eq "ERROR" }).Count
        $warnings = ($ReplicationData | Where-Object { $_.Status -eq "WARNING" }).Count
        $okCount = ($ReplicationData | Where-Object { $_.Status -eq "OK" -or $_.Status -eq "GOOD" -or $_.Status -eq "EXCELLENT" }).Count
        
        # Add summary to the beginning
        $summary = [PSCustomObject]@{
            Category = "SUMMARY"
            SourceDC = "All DCs"
            TargetDC = "All Partners"
            Partition = "N/A"
            Status = if ($totalErrors -eq 0 -and $warnings -eq 0) { "HEALTHY" } elseif ($totalErrors -eq 0) { "MINOR ISSUES" } else { "CRITICAL ISSUES" }
            ErrorType = "Overview"
            ErrorCount = $totalErrors
            FirstFailure = $null
            LastFailure = $null
            Details = "OK/Good: $okCount | Warnings: $warnings | Critical/Error: $totalErrors | DCs: $($DomainControllers.Count)"
        }
        
        # Final report with summary
        $FinalReplicationData = @($summary) + $ReplicationData
        
        Write-ADReportLog -Message "Replication analysis completed: $($FinalReplicationData.Count) entries found." -Type Info -Terminal
        return $FinalReplicationData
        
    } catch {
        $ErrorMessage = "Critical error during replication analysis: $($_.Exception.Message)"
        Write-ADReportLog -Message $ErrorMessage -Type Error
        return @([PSCustomObject]@{
            Category = "SYSTEM ERROR"
            SourceDC = "Unknown"
            TargetDC = "Unknown"
            Partition = "N/A"
            Status = "CRITICAL"
            ErrorType = "System Error"
            ErrorCount = 1
            FirstFailure = Get-Date
            LastFailure = Get-Date
            Details = $ErrorMessage
        })
    }
}

# --- Funktion zum Abrufen von OU-Hierarchie-Berichten als Baumstruktur ---
Function Get-ADOUHierarchyReport {
    [CmdletBinding()]
    param()

    Write-ADReportLog -Message "Generating OU hierarchy report as tree structure..." -Type Info -Terminal
    try {
        # Get OUs only, no containers
        $AllOUs = Get-ADOrganizationalUnit -Filter * `
            -Properties DistinguishedName, Name, Description, whenCreated, whenChanged, 
                       ProtectedFromAccidentalDeletion, LinkedGroupPolicyObjects, ManagedBy -ErrorAction Stop
        
        if (-not $AllOUs) {
            Write-ADReportLog -Message "No organizational units found in AD." -Type Warning -Terminal
            return $null
        }

        # Get domain information for root level
        $Domain = Get-ADDomain -ErrorAction Stop
        $DomainDN = $Domain.DistinguishedName
        
        # Create hierarchy dictionary and root OUs list
        $HierarchyDict = @{}
        $RootOUs = [System.Collections.ArrayList]@()
        
        foreach ($OU in $AllOUs) {
            $Level = 0
            $ParentDN = $OU.DistinguishedName.Substring($OU.DistinguishedName.IndexOf(',') + 1)
            
            # Calculate indentation depth based on OU path
            $Level = ($OU.DistinguishedName.Split(',').Count - $DomainDN.Split(',').Count)
            $Indent = "-- " * $Level

            # Get linked GPOs count
            $GPOCount = if ($OU.LinkedGroupPolicyObjects) { 
                @($OU.LinkedGroupPolicyObjects).Count 
            } else { 
                0 
            }

            # Get manager info if exists
            $Manager = if ($OU.ManagedBy) {
                try {
                    $mgr = Get-ADObject $OU.ManagedBy -Properties DisplayName -ErrorAction SilentlyContinue
                    $mgr.DisplayName
                } catch {
                    "Unknown"
                }
            } else {
                "Not Set"
            }
            
            $Entry = [PSCustomObject]@{
                Level = $Level
                Name = $OU.Name
                FullPath = $OU.DistinguishedName
                DisplayName = "$Indent$($OU.Name)"
                Description = $OU.Description
                Created = $OU.whenCreated
                Modified = $OU.whenChanged
                ParentDN = if ($ParentDN -eq $DomainDN) { "Root" } else { $ParentDN }
                Protected = $OU.ProtectedFromAccidentalDeletion
                LinkedGPOs = $GPOCount
                ManagedBy = $Manager
                Status = if ($OU.ProtectedFromAccidentalDeletion) { "Protected" } else { "Not Protected" }
            }
            
            if ($ParentDN -eq $DomainDN) {
                [void]$RootOUs.Add($Entry)
            } else {
                if (-not $HierarchyDict.ContainsKey($ParentDN)) {
                    $HierarchyDict[$ParentDN] = [System.Collections.ArrayList]@()
                }
                [void]$HierarchyDict[$ParentDN].Add($Entry)
            }
        }
        
        # Create sorted output list
        $Results = [System.Collections.ArrayList]@()
        
        # Add domain root as first entry
        [void]$Results.Add([PSCustomObject]@{
            Level = 0
            DisplayName = $Domain.NetBIOSName
            Description = "Domain Root"
            Created = $null
            Modified = $null
            FullPath = $DomainDN
            ParentDN = $null
            Protected = $true
            LinkedGPOs = 0
            ManagedBy = "System"
            Status = "Domain Root"
        })
        
        # Recursive function to build tree structure
        function Add-ChildOUs {
            param($ParentDN)
            
            $Children = $HierarchyDict[$ParentDN] | Sort-Object Name
            foreach ($Child in $Children) {
                [void]$Results.Add($Child)
                if ($HierarchyDict.ContainsKey($Child.FullPath)) {
                    Add-ChildOUs -ParentDN $Child.FullPath
                }
            }
        }
        
        # Process root level OUs
        foreach ($RootOU in ($RootOUs | Sort-Object Name)) {
            [void]$Results.Add($RootOU)
            if ($HierarchyDict.ContainsKey($RootOU.FullPath)) {
                Add-ChildOUs -ParentDN $RootOU.FullPath
            }
        }
        
        Write-ADReportLog -Message "OU hierarchy report successfully created. Found $($Results.Count) entries." -Type Info -Terminal
        return $Results
    }
    catch {
        Write-ADReportLog -Message "Error creating OU hierarchy report: $($_.Exception.Message)" -Type Error -Terminal
        return $null
    }
}

# --- Function to get AD Sites and Subnets Report ---
Function Get-ADSitesAndSubnetsReport {
    [CmdletBinding()]
    param()

    # Initialize message table for localization
    $msgTable = data {
        #culture="en-US" 
        ConvertFrom-StringData @'
        GatheringInfo = Gathering AD Sites and Subnets information...
        NoServersFound = Could not get servers for site {0}: {1}
        NoSiteLinksFound = Could not get site links for site {0}: {1}
        SitesRetrieved = Successfully retrieved {0} AD Replication Sites.
        NoSitesFound = No AD Replication Sites found.
        SubnetsRetrieved = Successfully retrieved {0} AD Replication Subnets.
        NoSubnetsFound = No AD Replication Subnets found.
        ReportGenerated = Successfully generated Sites and Subnets Report for {0} entries.
        NoDataFound = No data found for Sites and Subnets Report.
        ErrorGenerating = Error generating Sites and Subnets Report: {0}
'@
    }

    # Import localized data if available
    try {
        Import-LocalizedData -BindingVariable msgTable
    } catch {
        Write-ADReportLog -Message "Using default English messages - localization file not found." -Type Info
    }

    Write-ADReportLog -Message $msgTable.GatheringInfo -Type Info -Terminal
    Initialize-ResultCounters

    try {
        $Report = @()

        # Get Sites
        $Sites = Get-ADReplicationSite -Filter * -Properties Description, Options, InterSiteTopologyGenerator -ErrorAction SilentlyContinue
        if ($Sites) {
            foreach ($Site in $Sites) {
                # Get servers for this site separately
                $ServersCount = 0
                try {
                    $Servers = Get-ADDomainController -Filter { Site -eq $Site.Name } -ErrorAction SilentlyContinue
                    $ServersCount = if ($Servers) { @($Servers).Count } else { 0 }
                } catch {
                    Write-ADReportLog -Message ($msgTable.NoServersFound -f $Site.Name, $_.Exception.Message) -Type Warning
                }

                # Get Site Links separately
                $SiteLinksText = ""
                try {
                    $SiteLinks = Get-ADReplicationSiteLink -Filter "(Sites -eq '$($Site.DistinguishedName)')" -ErrorAction SilentlyContinue
                    $SiteLinksText = if ($SiteLinks) { ($SiteLinks | ForEach-Object { $_.Name }) -join ", " }
                } catch {
                    Write-ADReportLog -Message ($msgTable.NoSiteLinksFound -f $Site.Name, $_.Exception.Message) -Type Warning
                }

                $Report += [PSCustomObject]@{
                    Type = "Site"
                    Name = $Site.Name
                    DistinguishedName = $Site.DistinguishedName
                    Description = $Site.Description
                    ServersInSiteCount = $ServersCount
                    InterSiteTopologyGenerator = $Site.InterSiteTopologyGenerator
                    Options = $Site.Options
                    SiteLinks = $SiteLinksText
                    Location = $null
                    AssociatedSite = $null
                }
            }
            Write-ADReportLog -Message ($msgTable.SitesRetrieved -f $Sites.Count) -Type Info -Terminal
        } else {
            Write-ADReportLog -Message $msgTable.NoSitesFound -Type Warning -Terminal
        }

        # Get Subnets
        $Subnets = Get-ADReplicationSubnet -Filter * -Properties Description, Location, Site -ErrorAction SilentlyContinue
        if ($Subnets) {
            foreach ($Subnet in $Subnets) {
                $Report += [PSCustomObject]@{
                    Type = "Subnet"
                    Name = $Subnet.Name
                    DistinguishedName = $Subnet.DistinguishedName
                    Description = $Subnet.Description
                    ServersInSiteCount = $null 
                    InterSiteTopologyGenerator = $null
                    Options = $null
                    SiteLinks = $null
                    Location = $Subnet.Location
                    AssociatedSite = try { 
                        (Get-ADReplicationSite -Identity $Subnet.Site -ErrorAction SilentlyContinue).Name 
                    } catch { 
                        $Subnet.Site 
                    }
                }
            }
            Write-ADReportLog -Message ($msgTable.SubnetsRetrieved -f $Subnets.Count) -Type Info -Terminal
        } else {
            Write-ADReportLog -Message $msgTable.NoSubnetsFound -Type Warning -Terminal
        }

        if ($Report.Count -gt 0) {
            Write-ADReportLog -Message ($msgTable.ReportGenerated -f $Report.Count) -Type Info
            return $Report | Sort-Object Type, Name
        } else {
            Write-ADReportLog -Message $msgTable.NoDataFound -Type Info
            return @(
                [PSCustomObject]@{
                    Type = "Information"
                    Name = "No Data"
                    DistinguishedName = "N/A"
                    Description = "No Sites or Subnets found in Active Directory"
                    ServersInSiteCount = 0
                    InterSiteTopologyGenerator = "N/A"
                    Options = "N/A"
                    SiteLinks = "N/A"
                    Location = "N/A"
                    AssociatedSite = "N/A"
                }
            )
        }
    }
    catch {
        $ErrorMessage = $msgTable.ErrorGenerating -f $_.Exception.Message
        Write-ADReportLog -Message $ErrorMessage -Type Error -Terminal
        return @(
            [PSCustomObject]@{
                Type = "Error"
                Name = "System Error"
                DistinguishedName = "N/A" 
                Description = $ErrorMessage
                ServersInSiteCount = 0
                InterSiteTopologyGenerator = "N/A"
                Options = "N/A"
                SiteLinks = "N/A"
                Location = "N/A"
                AssociatedSite = "N/A"
            }
        )
    }
}


# --- Neue Roadmap Features - Benutzer-Reports ---
Function Get-StalePasswords {
    [CmdletBinding()]
    param([int]$Days = 90)
    
    try {
        Write-ADReportLog -Message "Analyzing users with stale passwords (older than $Days days)..." -Type Info -Terminal
        $CutoffDate = (Get-Date).AddDays(-$Days)
        $Users = Get-ADUser -Filter "Enabled -eq `$true" -Properties PasswordLastSet, DisplayName, SamAccountName, Department, Title -ErrorAction Stop
        $StalePasswordUsers = $Users | Where-Object { $_.PasswordLastSet -lt $CutoffDate }
        
        $Results = $StalePasswordUsers | Select-Object DisplayName, SamAccountName, Department, Title, 
            @{Name="PasswordLastSet";Expression={$_.PasswordLastSet}},
            @{Name="DaysSinceLastChange";Expression={(New-TimeSpan -Start $_.PasswordLastSet -End (Get-Date)).Days}}
        
        return $Results
    } catch {
        Write-ADReportLog -Message "Error analyzing stale passwords: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-NeverChangingPasswords {
    [CmdletBinding()]
    param([int]$Days = 365)
    
    try {
        Write-ADReportLog -Message "Analyzing users with passwords that never change or are very old (older than $Days days)..." -Type Info -Terminal
        $CutoffDate = (Get-Date).AddDays(-$Days)
        $Users = Get-ADUser -Filter "Enabled -eq `$true" -Properties PasswordLastSet, WhenCreated, DisplayName, SamAccountName, Department, Title, PasswordNeverExpires -ErrorAction Stop
        
        $NeverChangedUsers = $Users | Where-Object { 
            # Passwort nie geÃ¤ndert seit Erstellung
            ($_.PasswordLastSet -eq $_.WhenCreated) -or 
            # Passwort ist null
            ($_.PasswordLastSet -eq $null) -or
            # Passwort ist sehr alt
            ($_.PasswordLastSet -lt $CutoffDate) -or
            # Passwort lÃ¤uft nie ab und ist alt
            ($_.PasswordNeverExpires -eq $true -and $_.PasswordLastSet -lt $CutoffDate)
        }
        
        $Results = foreach ($user in $NeverChangedUsers) {
            $daysSinceChange = if ($user.PasswordLastSet) { 
                (New-TimeSpan -Start $user.PasswordLastSet -End (Get-Date)).Days 
            } else { 
                if ($user.WhenCreated) { (New-TimeSpan -Start $user.WhenCreated -End (Get-Date)).Days } else { 9999 }
            }
            
            [PSCustomObject]@{
                DisplayName = $user.DisplayName
                SamAccountName = $user.SamAccountName
                Department = $user.Department
                Title = $user.Title
                WhenCreated = $user.WhenCreated
                PasswordLastSet = if ($user.PasswordLastSet) { $user.PasswordLastSet } else { "Never" }
                PasswordNeverExpires = $user.PasswordNeverExpires
                DaysSinceLastChange = $daysSinceChange
                RiskLevel = if ($daysSinceChange -gt 730) { "Critical" } elseif ($daysSinceChange -gt 365) { "High" } else { "Medium" }
            }
        }
        
        Write-ADReportLog -Message "Never changing passwords analysis completed. $($Results.Count) accounts found." -Type Info -Terminal
        return $Results | Sort-Object DaysSinceLastChange -Descending
    } catch {
        Write-ADReportLog -Message "Error analyzing never changing passwords: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-ExpiringAccounts {
    [CmdletBinding()]
    param([int]$Days = 30)
    
    try {
        Write-ADReportLog -Message "Analyzing accounts expiring within $Days days..." -Type Info -Terminal
        $FutureDate = (Get-Date).AddDays($Days)
        $Users = Get-ADUser -Filter "Enabled -eq `$true -and AccountExpirationDate -like '*'" -Properties AccountExpirationDate, DisplayName, SamAccountName, Department -ErrorAction Stop
        $ExpiringUsers = $Users | Where-Object { $_.AccountExpirationDate -le $FutureDate -and $_.AccountExpirationDate -gt (Get-Date) }
        
        $Results = $ExpiringUsers | Select-Object DisplayName, SamAccountName, Department,
            @{Name="AccountExpirationDate";Expression={$_.AccountExpirationDate}},
            @{Name="DaysUntilExpiration";Expression={(New-TimeSpan -Start (Get-Date) -End $_.AccountExpirationDate).Days}}
        
        return $Results
    } catch {
        Write-ADReportLog -Message "Error analyzing expiring accounts: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-ReversibleEncryptionUsers {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing users with reversible encryption..." -Type Info -Terminal
        $Users = Get-ADUser -Filter "Enabled -eq `$true" -Properties AllowReversiblePasswordEncryption, DisplayName, SamAccountName, Department -ErrorAction Stop
        $ReversibleUsers = $Users | Where-Object { $_.AllowReversiblePasswordEncryption -eq $true }
        
        $Results = $ReversibleUsers | Select-Object DisplayName, SamAccountName, Department,
            @{Name="AllowReversiblePasswordEncryption";Expression={$_.AllowReversiblePasswordEncryption}}
        
        return $Results
    } catch {
        Write-ADReportLog -Message "Error analyzing reversible encryption users: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-KerberosDESUsers {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing users with Kerberos DES encryption..." -Type Info -Terminal
        $Users = Get-ADUser -Filter "Enabled -eq `$true" -Properties UserAccountControl, DisplayName, SamAccountName, Department -ErrorAction Stop
        $DESUsers = $Users | Where-Object { ($_.UserAccountControl -band 0x200000) -ne 0 }
        
        $Results = $DESUsers | Select-Object DisplayName, SamAccountName, Department,
            @{Name="UsesDESEncryption";Expression={"True"}}
        
        return $Results
    } catch {
        Write-ADReportLog -Message "Error analyzing Kerberos DES users: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-UsersWithSPN {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing users with Service Principal Names..." -Type Info -Terminal
        $Users = Get-ADUser -Filter "Enabled -eq `$true -and ServicePrincipalName -like '*'" -Properties ServicePrincipalName, DisplayName, SamAccountName, Department -ErrorAction Stop
        
        $Results = $Users | Select-Object DisplayName, SamAccountName, Department,
            @{Name="ServicePrincipalNames";Expression={$_.ServicePrincipalName -join "; "}}
        
        return $Results
    } catch {
        Write-ADReportLog -Message "Error analyzing users with SPN: $($_.Exception.Message)" -Type Error
        return @()
    }
}

# --- Weitere User Report Funktionen ---
Function Get-GuestAccountStatus {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing Guest account status..." -Type Info -Terminal
        
        # Suche nach dem Guest Account (kann verschiedene Namen haben)
        $GuestAccounts = @()
        
        # Standard Guest Account Namen in verschiedenen Sprachen
        $GuestNames = @("Guest", "Gast", "InvitÃ©", "Invitado", "Ospite")
        
        foreach ($name in $GuestNames) {
            $account = Get-ADUser -Filter "SamAccountName -eq '$name'" -Properties * -ErrorAction SilentlyContinue
            if ($account) {
                $GuestAccounts += $account
            }
        }
        
        # Auch nach SID suchen (Well-known Guest SID endet mit -501)
        $DomainSID = (Get-ADDomain).DomainSID.Value
        $GuestSID = "$DomainSID-501"
        
        try {
            $GuestBySID = Get-ADUser -Identity $GuestSID -Properties * -ErrorAction SilentlyContinue
            if ($GuestBySID -and -not ($GuestAccounts | Where-Object { $_.SID -eq $GuestBySID.SID })) {
                $GuestAccounts += $GuestBySID
            }
        } catch {
            # SID nicht gefunden ist OK
        }
        
        if ($GuestAccounts.Count -eq 0) {
            Write-ADReportLog -Message "No Guest account found in domain." -Type Info -Terminal
            return @([PSCustomObject]@{
                AccountName = "No Guest Account"
                Status = "Not Found"
                SecurityRisk = "Low"
                Recommendation = "Guest account not present in domain (good security practice)"
            })
        }
        
        $Results = foreach ($guest in $GuestAccounts) {
            # Analysiere Sicherheitsrisiken
            $risks = @()
            $riskLevel = "Low"
            $recommendations = @()
            
            if ($guest.Enabled) {
                $risks += "Account is enabled"
                $riskLevel = "High"
                $recommendations += "Disable Guest account immediately"
            }
            
            if ($guest.PasswordNeverExpires) {
                $risks += "Password never expires"
                if ($riskLevel -ne "High") { $riskLevel = "Medium" }
                $recommendations += "Set password expiration"
            }
            
            if ($guest.PasswordNotRequired) {
                $risks += "Password not required"
                $riskLevel = "Critical"
                $recommendations += "Enforce password requirement"
            }
            
            if ($guest.LastLogonDate -and $guest.LastLogonDate -gt (Get-Date).AddDays(-30)) {
                $risks += "Recently used (within 30 days)"
                if ($riskLevel -eq "Low") { $riskLevel = "Medium" }
                $recommendations += "Investigate recent usage"
            }
            
            # PrÃ¼fe Gruppenmitgliedschaften
            $groups = Get-ADPrincipalGroupMembership -Identity $guest -ErrorAction SilentlyContinue
            $privilegedGroups = @()
            
            foreach ($group in $groups) {
                if ($group.Name -ne "Domain Guests" -and $group.Name -ne "Guests") {
                    $privilegedGroups += $group.Name
                }
            }
            
            if ($privilegedGroups.Count -gt 0) {
                $risks += "Member of additional groups"
                $riskLevel = "Critical"
                $recommendations += "Remove from all groups except Domain Guests"
            }
            
            [PSCustomObject]@{
                AccountName = $guest.SamAccountName
                DisplayName = $guest.DisplayName
                Enabled = $guest.Enabled
                PasswordLastSet = $guest.PasswordLastSet
                LastLogonDate = $guest.LastLogonDate
                PasswordNeverExpires = $guest.PasswordNeverExpires
                PasswordNotRequired = $guest.PasswordNotRequired
                AccountLockedOut = $guest.LockedOut
                SID = $guest.SID
                Groups = if ($privilegedGroups) { $privilegedGroups -join ", " } else { "Domain Guests only" }
                RiskLevel = $riskLevel
                SecurityIssues = if ($risks) { $risks -join "; " } else { "None" }
                Recommendations = if ($recommendations) { $recommendations -join "; " } else { "Account properly secured" }
                WhenCreated = $guest.WhenCreated
                WhenChanged = $guest.WhenChanged
            }
        }
        
        Write-ADReportLog -Message "Guest account analysis completed. $($Results.Count) account(s) found." -Type Info -Terminal
        return $Results
        
    } catch {
        Write-ADReportLog -Message "Error analyzing Guest account status: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-UsersByDepartment {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing users by department..." -Type Info -Terminal
        
        # Lade alle Benutzer mit Department-Attribut
        $Users = Get-ADUser -Filter * -Properties Department, DisplayName, SamAccountName, Title, Enabled, LastLogonDate, Manager -ErrorAction Stop
        
        # Gruppiere nach Department und sortiere Departments alphabetisch
        $DepartmentGroups = $Users | Group-Object Department | Sort-Object Name
        
        $Results = foreach ($deptGroup in $DepartmentGroups) {
            $deptName = if ([string]::IsNullOrWhiteSpace($deptGroup.Name)) { "(No Department)" } else { $deptGroup.Name }
            
            # Department-Header erstellen
            [PSCustomObject]@{
                Type = "Department"
                Department = $deptName
                DisplayName = ""
                SamAccountName = ""
                Title = ""
                Manager = ""
                Enabled = $null
                LastLogonDate = $null
                DeptUserCount = $deptGroup.Count
                DeptEnabledCount = if ($deptName -eq "(No Department)") { $null } else { ($deptGroup.Group | Where-Object { $_.Enabled }).Count }
                DeptActiveCount = if ($deptName -eq "(No Department)") { $null } else { ($deptGroup.Group | Where-Object { $_.LastLogonDate -gt (Get-Date).AddDays(-30) }).Count }
                ActivityStatus = ""
            }
            
            # Manager Dictionary erstellen
            $managersDict = @{}
            foreach ($user in $deptGroup.Group) {
                if ($user.Manager) {
                    if (-not $managersDict.ContainsKey($user.Manager)) {
                        try {
                            $managerObj = Get-ADUser -Identity $user.Manager -Properties DisplayName -ErrorAction SilentlyContinue
                            if ($managerObj) {
                                $managersDict[$user.Manager] = $managerObj.DisplayName
                            }
                        } catch {
                            # Manager nicht gefunden
                        }
                    }
                }
            }
            
            # Benutzer innerhalb der Abteilung alphabetisch sortieren und ausgeben
            $sortedUsers = $deptGroup.Group | Sort-Object DisplayName
            foreach ($user in $sortedUsers) {
                $managerName = "None"
                if ($user.Manager -and $managersDict.ContainsKey($user.Manager)) {
                    $managerName = $managersDict[$user.Manager]
                }
                
                [PSCustomObject]@{
                    Type = "User"
                    Department = $deptName
                    DisplayName = "    ├─ $($user.DisplayName)" # Einrückung für Baumstruktur
                    SamAccountName = $user.SamAccountName
                    Title = $user.Title
                    Manager = $managerName
                    Enabled = $user.Enabled
                    LastLogonDate = $user.LastLogonDate
                    DeptUserCount = $null # Nur im Department-Header anzeigen
                    DeptEnabledCount = $null
                    DeptActiveCount = $null
                    ActivityStatus = if ($user.LastLogonDate -gt (Get-Date).AddDays(-30)) { "Active" } 
                                   elseif ($user.LastLogonDate -gt (Get-Date).AddDays(-90)) { "Inactive" }
                                   else { "Very Inactive" }
                }
            }
        }
        
        Write-ADReportLog -Message "Users by department analysis completed. Found users in $($DepartmentGroups.Count) departments." -Type Info -Terminal
        return $Results
        
    } catch {
        Write-ADReportLog -Message "Error analyzing users by department: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-UsersByManager {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing users by manager..." -Type Info -Terminal
        
        # Lade alle Benutzer mit Manager-Attribut
        $Users = Get-ADUser -Filter * -Properties Manager, DisplayName, SamAccountName, Department, Title, Enabled, LastLogonDate -ErrorAction Stop
        
        # Erstelle Dictionary für Manager-Namen
        $managersDict = @{}
        $usersWithManager = $Users | Where-Object { $_.Manager }
        
        foreach ($user in $usersWithManager) {
            if (-not $managersDict.ContainsKey($user.Manager)) {
                try {
                    $managerObj = Get-ADUser -Identity $user.Manager -Properties DisplayName, Department, Title -ErrorAction SilentlyContinue
                    if ($managerObj) {
                        $managersDict[$user.Manager] = $managerObj
                    }
                } catch {
                    # Manager nicht gefunden
                }
            }
        }
        
        # Gruppiere nach Manager
        $ManagerGroups = $usersWithManager | Group-Object Manager
        
        $Results = @()
        
        # Verarbeite Benutzer mit Manager
        foreach ($mgrGroup in $ManagerGroups) {
            $managerInfo = $managersDict[$mgrGroup.Name]
            $managerName = if ($managerInfo) { $managerInfo.DisplayName } else { "Unknown Manager" }
            $managerDept = if ($managerInfo) { $managerInfo.Department } else { "Unknown" }
            $managerTitle = if ($managerInfo) { $managerInfo.Title } else { "Unknown" }
            
            # Manager-Header
            $Results += [PSCustomObject]@{
                ManagerName = $managerName
                ManagerDepartment = $managerDept
                ManagerTitle = $managerTitle
                DirectReports = $mgrGroup.Count
                UserDisplayName = "└─ Manager: $managerName"
                UserSamAccountName = ""
                UserDepartment = $managerDept
                UserTitle = $managerTitle
                UserEnabled = $true
                UserLastLogon = $null
            }
            
            # Sortierte Mitarbeiter unter dem Manager
            $sortedUsers = $mgrGroup.Group | Sort-Object DisplayName
            foreach ($user in $sortedUsers) {
                $Results += [PSCustomObject]@{
                    ManagerName = $managerName
                    ManagerDepartment = $managerDept
                    ManagerTitle = $managerTitle
                    DirectReports = $null
                    UserDisplayName = "    ├─ $($user.DisplayName)"
                    UserSamAccountName = $user.SamAccountName
                    UserDepartment = $user.Department
                    UserTitle = $user.Title
                    UserEnabled = $user.Enabled
                    UserLastLogon = $user.LastLogonDate
                }
            }
        }
        
        # Benutzer ohne Manager als separate Gruppe
        $usersWithoutManager = $Users | Where-Object { -not $_.Manager } | Sort-Object DisplayName
        if ($usersWithoutManager) {
            # Header für Benutzer ohne Manager
            $Results += [PSCustomObject]@{
                ManagerName = "(No Manager)"
                ManagerDepartment = "N/A" 
                ManagerTitle = "N/A"
                DirectReports = $usersWithoutManager.Count
                UserDisplayName = "└─ Users without Manager"
                UserSamAccountName = ""
                UserDepartment = "N/A"
                UserTitle = "N/A"
                UserEnabled = $true
                UserLastLogon = $null
            }
            
            foreach ($user in $usersWithoutManager) {
                $Results += [PSCustomObject]@{
                    ManagerName = "(No Manager)"
                    ManagerDepartment = "N/A"
                    ManagerTitle = "N/A"
                    DirectReports = $null
                    ActiveReports = $null
                    UserDisplayName = "    ├─ $($user.DisplayName)"
                    UserSamAccountName = $user.SamAccountName
                    UserDepartment = $user.Department
                    UserTitle = $user.Title
                    UserEnabled = $user.Enabled
                    UserLastLogon = $user.LastLogonDate
                }
            }
        }
        
        Write-ADReportLog -Message "Users by manager analysis completed. $($Results.Count) users analyzed." -Type Info -Terminal
        return $Results
        
    } catch {
        Write-ADReportLog -Message "Error analyzing users by manager: $($_.Exception.Message)" -Type Error
        return @()
    }
}
Function Get-EmptyGroups {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing empty groups..." -Type Info -Terminal
        $Groups = Get-ADGroup -Filter * -Properties Members, Name, GroupCategory, GroupScope, whenCreated, whenChanged -ErrorAction Stop
        
        $EmptyGroups = $Groups | Where-Object { 
            # PrÃ¼fe sowohl direkte Members als auch Ã¼ber Get-ADGroupMember
            $_.Members.Count -eq 0 -and 
            @(Get-ADGroupMember -Identity $_.DistinguishedName -ErrorAction SilentlyContinue).Count -eq 0
        }
        
        $Results = foreach ($group in $EmptyGroups) {
            $ageInDays = if ($group.whenCreated) { (New-TimeSpan -Start $group.whenCreated -End (Get-Date)).Days } else { 0 }
            
            [PSCustomObject]@{
                Name = $group.Name
                GroupCategory = $group.GroupCategory
                GroupScope = $group.GroupScope
                Description = $group.Description
                WhenCreated = $group.whenCreated
                WhenChanged = $group.whenChanged
                AgeInDays = $ageInDays
                MemberCount = 0
                CleanupRecommended = if ($ageInDays -gt 90 -and [string]::IsNullOrWhiteSpace($group.Description)) { "Yes" } else { "Review" }
            }
        }
        
        Write-ADReportLog -Message "Empty groups analysis completed. $($Results.Count) groups found." -Type Info -Terminal
        return $Results | Sort-Object AgeInDays -Descending
    } catch {
        Write-ADReportLog -Message "Error analyzing empty groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-CircularNestedGroups {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing nested group memberships..." -Type Info -Terminal
        $Groups = Get-ADGroup -Filter * -Properties MemberOf, Name, GroupCategory -ErrorAction Stop
        $NestedGroups = $Groups | Where-Object { $_.MemberOf.Count -gt 0 }
        
        $Results = $NestedGroups | Select-Object Name, GroupCategory,
            @{Name="MemberOfGroups";Expression={$_.MemberOf.Count}},
            @{Name="ParentGroups";Expression={($_.MemberOf | ForEach-Object { (Get-ADGroup $_).Name }) -join "; "}}
        
        return $Results
    } catch {
        Write-ADReportLog -Message "Error analyzing nested groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}
Function Get-GroupsByTypeAndScope {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing groups by type and scope..." -Type Info -Terminal
        
        $Groups = Get-ADGroup -Filter * -Properties Name, GroupCategory, GroupScope, whenCreated, whenChanged, MemberOf, Members -ErrorAction Stop
        
        $Results = foreach ($group in $Groups) {
            [PSCustomObject]@{
                Name = $group.Name
                Description = $group.Description
                GroupCategory = $group.GroupCategory # Security oder Distribution
                GroupScope = $group.GroupScope # DomainLocal, Global oder Universal
                WhenCreated = $group.whenCreated
                WhenChanged = $group.whenChanged
                MemberCount = if ($group.Members) { $group.Members.Count } else { 0 }
                MemberOfCount = if ($group.MemberOf) { $group.MemberOf.Count } else { 0 }
                AgeInDays = if ($group.whenCreated) { 
                    [math]::Round((New-TimeSpan -Start $group.whenCreated -End (Get-Date)).TotalDays, 0)
                } else { 0 }
            }
        }
        
        Write-ADReportLog -Message "Groups by type and scope analysis completed. $($Results.Count) groups found." -Type Info -Terminal
        return $Results | Sort-Object GroupCategory, GroupScope
    } catch {
        Write-ADReportLog -Message "Error analyzing groups by type and scope: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-DynamicDistributionGroups {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing dynamic distribution groups..." -Type Info -Terminal
        
        # Suche nach dynamischen Verteilerlisten anhand typischer Attribute
        # Exchange-Attribute nur verwenden, wenn Exchange installiert ist
        # Prüfe, ob Exchange-Attribute verfügbar sind
        $exchangeInstalled = $false
        try {
            $testGroup = Get-ADGroup -Filter * -Properties msExchRecipientDisplayType -ErrorAction Stop | Select-Object -First 1
            if ($null -ne $testGroup.PSObject.Properties['msExchRecipientDisplayType']) {
                $exchangeInstalled = $true
            }
        } catch {
            $exchangeInstalled = $false
        }
        
        if ($exchangeInstalled) {
            $DynamicGroups = Get-ADGroup -Filter {
                (GroupCategory -eq 'Distribution') -and 
                ((msExchRecipientDisplayType -eq '7') -or 
                 (msExchRecipientTypeDetails -eq '8589934592'))
            } -Properties Name, whenCreated, whenChanged, msExchRecipientDisplayType, msExchRecipientTypeDetails -ErrorAction Stop
        } else {
            # Ohne Exchange können wir keine echten Dynamic Distribution Groups identifizieren
            Write-ADReportLog -Message "Exchange attributes not available. Cannot identify dynamic distribution groups." -Type Warning
            $DynamicGroups = @()
        }
        
        $Results = foreach ($group in $DynamicGroups) {
            [PSCustomObject]@{
                Name = $group.Name
                Description = $group.Description
                WhenCreated = $group.whenCreated
                WhenChanged = $group.whenChanged
                AgeInDays = if ($group.whenCreated) {
                    [math]::Round((New-TimeSpan -Start $group.whenCreated -End (Get-Date)).TotalDays, 0)
                } else { 0 }
                RecipientType = "Dynamic Distribution Group"
                Status = "Active"
            }
        }
        
        Write-ADReportLog -Message "Dynamic distribution groups analysis completed. $($Results.Count) groups found." -Type Info -Terminal
        return $Results | Sort-Object Name
    } catch {
        Write-ADReportLog -Message "Error analyzing dynamic distribution groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-NestedGroups {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing nested group memberships..." -Type Info -Terminal
        $Groups = Get-ADGroup -Filter * -Properties MemberOf, Name, GroupCategory -ErrorAction Stop
        $NestedGroups = $Groups | Where-Object { $_.MemberOf.Count -gt 0 }
        
        $Results = $NestedGroups | Select-Object Name, GroupCategory,
            @{Name="MemberOfGroups";Expression={$_.MemberOf.Count}},
            @{Name="ParentGroups";Expression={($_.MemberOf | ForEach-Object { (Get-ADGroup $_).Name }) -join "; "}}
        
        return $Results
    } catch {
        Write-ADReportLog -Message "Error analyzing nested groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}

# --- Neue Roadmap Features - Computer-Reports ---
Function Get-OSSummary {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Generating Operating System summary..." -Type Info -Terminal
        $Computers = Get-ADComputer -Filter "Enabled -eq `$true" -Properties OperatingSystem, OperatingSystemVersion -ErrorAction Stop
        
        $OSGroups = $Computers | Group-Object OperatingSystem | Sort-Object Count -Descending
        $Results = $OSGroups | Select-Object @{Name="OperatingSystem";Expression={$_.Name}},
            @{Name="Count";Expression={$_.Count}},
            @{Name="Percentage";Expression={[math]::Round(($_.Count / $Computers.Count) * 100, 2)}}
        
        return $Results
    } catch {
        Write-ADReportLog -Message "Error generating OS summary: $($_.Exception.Message)" -Type Error
        return @()
    }
}


# --- Neue Roadmap Features - Service Account Reports ---
Function Get-ServiceAccountsOverview {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Generating comprehensive Service Accounts overview..." -Type Info -Terminal
        
        # Service Accounts identifizieren durch verschiedene Kriterien
        $PotentialServiceAccounts = @()
        
        # 1. Accounts mit Service-Ã¤hnlichen Namen
        $NameBasedSvcAccounts = Get-ADUser -Filter "(Name -like '*svc*') -or (Name -like '*service*') -or (Name -like '*sql*') -or (Name -like '*iis*') -or (Name -like '*app*') -or (Name -like '*web*')" -Properties Description, LastLogonDate, PasswordLastSet, ServicePrincipalName, PasswordNeverExpires, Enabled, whenCreated, Department -ErrorAction SilentlyContinue
        $PotentialServiceAccounts += $NameBasedSvcAccounts
        
        # 2. Accounts mit SPNs (Service Principal Names)
        $SPNAccounts = Get-ADUser -Filter "ServicePrincipalName -like '*'" -Properties Description, LastLogonDate, PasswordLastSet, ServicePrincipalName, PasswordNeverExpires, Enabled, whenCreated, Department -ErrorAction SilentlyContinue
        $PotentialServiceAccounts += $SPNAccounts
        
        # 3. Accounts mit service-bezogenen Beschreibungen
        $DescBasedSvcAccounts = Get-ADUser -Filter "Description -like '*service*' -or Description -like '*application*' -or Description -like '*sql*' -or Description -like '*database*'" -Properties Description, LastLogonDate, PasswordLastSet, ServicePrincipalName, PasswordNeverExpires, Enabled, whenCreated, Department -ErrorAction SilentlyContinue
        $PotentialServiceAccounts += $DescBasedSvcAccounts
        
        # Duplikate entfernen
        $UniqueServiceAccounts = $PotentialServiceAccounts | Sort-Object DistinguishedName | Get-Unique -AsString
        
        $Results = foreach ($account in $UniqueServiceAccounts) {
            $daysSincePasswordChange = if ($account.PasswordLastSet) { 
                (New-TimeSpan -Start $account.PasswordLastSet -End (Get-Date)).Days 
            } else { 9999 }
            
            $daysSinceLastLogon = if ($account.LastLogonDate) { 
                (New-TimeSpan -Start $account.LastLogonDate -End (Get-Date)).Days 
            } else { 9999 }
            
            # Risk Assessment
            $riskScore = 0
            $riskFactors = @()
            
            if ($daysSincePasswordChange -gt 365) { $riskScore += 3; $riskFactors += "Old Password" }
            if ($account.PasswordNeverExpires) { $riskScore += 2; $riskFactors += "Password Never Expires" }
            if ($daysSinceLastLogon -gt 90) { $riskScore += 2; $riskFactors += "Inactive" }
            if ($account.ServicePrincipalName.Count -gt 0) { $riskScore += 1; $riskFactors += "Has SPN" }
            if ($account.Enabled -eq $false) { $riskScore -= 2; $riskFactors += "Disabled" }
            
            $riskLevel = switch ($riskScore) {
                {$_ -ge 5} { "High" }
                {$_ -ge 3} { "Medium" }
                {$_ -ge 1} { "Low" }
                default { "Minimal" }
            }
            
            [PSCustomObject]@{
                Name = $account.Name
                SamAccountName = $account.SamAccountName
                Description = $account.Description
                Department = $account.Department
                Enabled = $account.Enabled
                LastLogonDate = if ($account.LastLogonDate) { $account.LastLogonDate } else { "Never" }
                DaysSinceLastLogon = $daysSinceLastLogon
                PasswordLastSet = if ($account.PasswordLastSet) { $account.PasswordLastSet } else { "Never" }
                DaysSincePasswordChange = $daysSincePasswordChange
                PasswordNeverExpires = $account.PasswordNeverExpires
                HasSPN = $account.ServicePrincipalName.Count -gt 0
                SPNCount = $account.ServicePrincipalName.Count
                RiskLevel = $riskLevel
                RiskFactors = $riskFactors -join ", "
                WhenCreated = $account.whenCreated
            }
        }
        
        Write-ADReportLog -Message "Service Accounts overview completed. $($Results.Count) accounts found." -Type Info -Terminal
        return $Results | Sort-Object RiskLevel, DaysSincePasswordChange -Descending
    } catch {
        Write-ADReportLog -Message "Error analyzing service accounts: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-ManagedServiceAccounts {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing Managed Service Accounts..." -Type Info -Terminal
        
        # gMSA und sMSA haben spezielle ObjectClasses
        $MSAs = Get-ADObject -Filter "ObjectClass -eq 'msDS-ManagedServiceAccount' -or ObjectClass -eq 'msDS-GroupManagedServiceAccount'" -Properties Name, ObjectClass, Created, Modified, DistinguishedName -ErrorAction Stop
        
        $Results = foreach ($msa in $MSAs) {
            [PSCustomObject]@{
                Name = $msa.Name
                AccountType = if($msa.ObjectClass -eq "msDS-GroupManagedServiceAccount") {"Group MSA"} else {"Standalone MSA"}
                Created = $msa.Created
                Modified = $msa.Modified
                DistinguishedName = $msa.DistinguishedName
                Status = "Active" # MSAs sind immer aktiv
            }
        }
        
        Write-ADReportLog -Message "Found $($Results.Count) managed service accounts" -Type Info -Terminal
        return $Results
        
    } catch {
        Write-ADReportLog -Message "Error analyzing managed service accounts: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-HoneyTokens {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analysiere potenzielle Honey Token Accounts und verdächtige Aktivitäten..." -Type Info -Terminal
        
        # Honey Tokens sind oft spezielle Accounts die zur Angriffserkennung verwendet werden
        # Wir suchen nach Accounts mit verdächtigen Eigenschaften oder Namensmustern
        
        $SuspiciousPatterns = @(
            "*honey*", "*canary*", "*trap*", "*decoy*", "*bait*", 
            "*test*", "*dummy*", "*fake*", "*monitor*", "*audit*"
        )
        
        [System.Collections.Generic.List[PSObject]]$PotentialHoneyTokens = @()
        
        # 1. Accounts mit verdächtigen Namen
        foreach ($pattern in $SuspiciousPatterns) {
            $accounts = Get-ADUser -Filter "Name -like '$pattern'" -Properties Description, LastLogonDate, PasswordLastSet, Enabled, whenCreated, Department -ErrorAction SilentlyContinue
            if ($null -ne $accounts) {
                $PotentialHoneyTokens.AddRange($accounts)
            }
        }
        
        # 2. Accounts mit verdächtigen Beschreibungen
        $DescriptionBasedAccounts = Get-ADUser -Filter "Description -like '*honey*' -or Description -like '*canary*' -or Description -like '*monitoring*' -or Description -like '*security*'" -Properties Description, LastLogonDate, PasswordLastSet, Enabled, whenCreated, Department -ErrorAction SilentlyContinue
        if ($null -ne $DescriptionBasedAccounts) {
            $PotentialHoneyTokens.AddRange($DescriptionBasedAccounts)
        }
        
        # 3. Niemals verwendete Admin-ähnliche Accounts (verdächtig)
        $UnusedAdminAccounts = Get-ADUser -Filter "Name -like '*admin*' -and Enabled -eq 'True'" -Properties Description, LastLogonDate, PasswordLastSet, Enabled, whenCreated, Department -ErrorAction SilentlyContinue | Where-Object { $_.LastLogonDate -eq $null }
        if ($null -ne $UnusedAdminAccounts) {
            $PotentialHoneyTokens.AddRange($UnusedAdminAccounts)
        }
        
        # Duplikate entfernen
        $UniqueAccounts = $PotentialHoneyTokens | Sort-Object DistinguishedName -Unique
        
        if ($null -eq $UniqueAccounts -or $UniqueAccounts.Count -eq 0) {
            Write-ADReportLog -Message "Keine potenziellen Honey Token Accounts gefunden." -Type Info -Terminal
            return @([PSCustomObject]@{
                Name = "Keine Ergebnisse"
                SamAccountName = "N/A"
                Description = "Keine potenziellen Honey Token Accounts in der aktuellen Umgebung gefunden"
                PotentialHoneyToken = "Keine gefunden"
                SuspicionLevel = "N/A"
                Indicators = "Analyse abgeschlossen - keine verdächtigen Muster erkannt"
            })
        }
        
        $Results = foreach ($account in $UniqueAccounts) {
            $ageInDays = if ($null -ne $account.whenCreated) { 
                [math]::Round((New-TimeSpan -Start $account.whenCreated -End (Get-Date)).TotalDays, 0)
            } else { 
                0 
            }
            
            $daysSinceLastLogon = if ($null -ne $account.LastLogonDate) { 
                [math]::Round((New-TimeSpan -Start $account.LastLogonDate -End (Get-Date)).TotalDays, 0)
            } else { 
                9999 
            }
            
            # Verdächtigkeits-Score
            [int]$suspicionScore = 0
            [System.Collections.Generic.List[string]]$indicators = @()
            
            # Name-basierte Indikatoren
            if ($account.Name -match "(honey|canary|trap|decoy|bait)") { $suspicionScore += 5; $indicators.Add("Verdächtiges Namensmuster") }
            if ($account.Name -match "(test|dummy|fake)") { $suspicionScore += 3; $indicators.Add("Test-Account Muster") }
            if ($account.Name -match "(monitor|audit)") { $suspicionScore += 4; $indicators.Add("Monitoring-Muster") }
            
            # Verhalten-basierte Indikatoren
            if ($null -eq $account.LastLogonDate -and $ageInDays -gt 30) { $suspicionScore += 3; $indicators.Add("Nie verwendet") }
            if ($account.Enabled -and $daysSinceLastLogon -gt 180) { $suspicionScore += 2; $indicators.Add("Lang inaktiv") }
            if ([string]::IsNullOrWhiteSpace($account.Department)) { $suspicionScore += 1; $indicators.Add("Keine Abteilung") }
            
            # Beschreibung-basierte Indikatoren
            if ($null -ne $account.Description -and $account.Description -match "(honey|canary|monitoring|security)") { 
                $suspicionScore += 4
                $indicators.Add("Sicherheits-Beschreibung") 
            }
            
            $suspicionLevel = switch ($suspicionScore) {
                {$_ -ge 7} { "Sehr Hoch" }
                {$_ -ge 5} { "Hoch" }
                {$_ -ge 3} { "Mittel" }
                {$_ -ge 1} { "Niedrig" }
                default { "Minimal" }
            }
            
            [PSCustomObject]@{
                Name = $account.Name
                SamAccountName = $account.SamAccountName
                Description = $account.Description
                Department = $account.Department
                Enabled = $account.Enabled
                WhenCreated = $account.whenCreated
                LastLogonDate = if ($null -ne $account.LastLogonDate) { $account.LastLogonDate } else { "Nie" }
                DaysSinceLastLogon = $daysSinceLastLogon
                AgeInDays = $ageInDays
                SuspicionLevel = $suspicionLevel
                SuspicionScore = $suspicionScore
                Indicators = $indicators -join ", "
                PotentialHoneyToken = if ($suspicionScore -ge 5) { "Wahrscheinlich" } elseif ($suspicionScore -ge 3) { "Möglich" } else { "Unwahrscheinlich" }
            }
        }
        
        Write-ADReportLog -Message "Honey Token Analyse abgeschlossen. $($Results.Count) potenzielle Accounts gefunden." -Type Info -Terminal
        return $Results | Sort-Object SuspicionScore, SuspicionLevel -Descending
        
    } catch {
        Write-ADReportLog -Message "Fehler bei der Honey Token Erkennung: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-GPOOverview {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Generating comprehensive GPO overview..." -Type Info -Terminal
        
        # PrÃ¼fe ob GroupPolicy Module verfÃ¼gbar ist
        if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
            Write-ADReportLog -Message "GroupPolicy PowerShell module not available. Attempting alternative approach..." -Type Warning
            
            # Alternative: Ãœber AD direkt abfragen
            $GPOs = Get-ADObject -SearchBase "CN=Policies,CN=System,$((Get-ADDomain).DistinguishedName)" -Filter "ObjectClass -eq 'groupPolicyContainer'" -Properties DisplayName, whenCreated, whenChanged, gPCFileSysPath -ErrorAction Stop
            
            $Results = foreach ($gpo in $GPOs) {
                [PSCustomObject]@{
                    DisplayName = $gpo.DisplayName
                    Id = $gpo.Name
                    CreationTime = $gpo.whenCreated
                    ModificationTime = $gpo.whenChanged
                    GPOStatus = "Unknown (Manual Check Required)"
                    FileSysPath = $gpo.gPCFileSysPath
                    AgeInDays = if ($gpo.whenCreated) { (New-TimeSpan -Start $gpo.whenCreated -End (Get-Date)).Days } else { 0 }
                    DaysSinceModification = if ($gpo.whenChanged) { (New-TimeSpan -Start $gpo.whenChanged -End (Get-Date)).Days } else { 0 }
                }
            }
            
            return $Results | Sort-Object DaysSinceModification
        }
        
        Import-Module GroupPolicy -ErrorAction Stop
        $GPOs = Get-GPO -All -ErrorAction Stop
        
        $Results = foreach ($gpo in $GPOs) {
            # GPO Links analysieren
            $gpoLinks = @()
            try {
                $gpoReport = Get-GPOReport -Guid $gpo.Id -ReportType Xml -ErrorAction SilentlyContinue
                if ($gpoReport) {
                    # Vereinfachte Link-Analyse
                    $linkCount = ([xml]$gpoReport).GPO.LinksTo.SOMPath.Count
                } else {
                    $linkCount = 0
                }
            } catch {
                $linkCount = 0
            }
            
            $ageInDays = if ($gpo.CreationTime) { (New-TimeSpan -Start $gpo.CreationTime -End (Get-Date)).Days } else { 0 }
            $daysSinceModification = if ($gpo.ModificationTime) { (New-TimeSpan -Start $gpo.ModificationTime -End (Get-Date)).Days } else { 0 }
            
            # Status-Bewertung
            $healthStatus = "Good"
            $issues = @()
            
            if ($daysSinceModification -gt 365) { $issues += "Not modified for over 1 year"; $healthStatus = "Review" }
            if ($linkCount -eq 0) { $issues += "No links found"; $healthStatus = "Orphaned" }
            if ($gpo.GpoStatus -eq "AllSettingsDisabled") { $issues += "All settings disabled"; $healthStatus = "Disabled" }
            
            [PSCustomObject]@{
                DisplayName = $gpo.DisplayName
                Id = $gpo.Id
                CreationTime = $gpo.CreationTime
                ModificationTime = $gpo.ModificationTime
                GPOStatus = $gpo.GpoStatus
                AgeInDays = $ageInDays
                DaysSinceModification = $daysSinceModification
                LinkCount = $linkCount
                HealthStatus = $healthStatus
                Issues = $issues -join "; "
                Domain = $gpo.DomainName
                Owner = $gpo.Owner
            }
        }
        
        Write-ADReportLog -Message "GPO overview completed. $($Results.Count) GPOs found." -Type Info -Terminal
        return $Results | Sort-Object HealthStatus, DaysSinceModification -Descending
    } catch {
        Write-ADReportLog -Message "Error generating GPO overview: $($_.Exception.Message)" -Type Error
        return @()
    }
}

# --- Weitere Gruppen-Report Funktionen ---
Function Get-CircularGroups {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing circular group references..." -Type Info -Terminal
        
        # Dictionary zur Speicherung der Gruppenmitgliedschaften
        $groupMemberships = @{}
        $circularReferences = @()
        
        # Lade alle Gruppen mit ihren Mitgliedschaften
        $allGroups = Get-ADGroup -Filter * -Properties MemberOf, Members, GroupCategory, GroupScope -ErrorAction Stop
        
        # Baue Dictionary auf
        foreach ($group in $allGroups) {
            $groupMemberships[$group.DistinguishedName] = @{
                Group = $group
                MemberOf = $group.MemberOf
                Visited = $false
                InStack = $false
            }
        }
        
        # Funktion zur rekursiven Suche nach zirkulären Referenzen
        function Find-CircularReference {
            param(
                [string]$GroupDN,
                [System.Collections.ArrayList]$Path
            )
            
            if (-not $groupMemberships.ContainsKey($GroupDN)) { return }
            
            $groupData = $groupMemberships[$GroupDN]
            
            # Wenn bereits im Stack, haben wir eine zirkuläre Referenz gefunden
            if ($groupData.InStack) {
                $circularPath = $Path.Clone()
                $circularPath.Add($GroupDN) | Out-Null
                
                # Finde den Start des Zyklus
                $startIndex = $circularPath.IndexOf($GroupDN)
                $cycle = $circularPath[$startIndex..($circularPath.Count - 1)]
                
                # Erstelle eine eindeutige ID für den Zyklus
                $cycleId = ($cycle | Sort-Object) -join '|'
                
                if (-not ($circularReferences | Where-Object { $_.CycleId -eq $cycleId })) {
                    $groupNames = $cycle | ForEach-Object {
                        if ($groupMemberships.ContainsKey($_)) {
                            $groupMemberships[$_].Group.Name
                        } else {
                            "Unknown"
                        }
                    }
                    
                    $circularReferences += [PSCustomObject]@{
                        CycleId = $cycleId
                        CircularPath = $groupNames -join " → "
                        NumberOfGroups = $cycle.Count
                        FirstGroup = $groupNames[0]
                        GroupsInvolved = $groupNames -join ", "
                        Risk = if ($cycle.Count -gt 3) { "High" } elseif ($cycle.Count -eq 3) { "Medium" } else { "Low" }
                        Recommendation = "Break circular reference by removing one membership link"
                    }
                }
                return
            }
            
            # Markiere als besucht und im Stack
            $groupData.Visited = $true
            $groupData.InStack = $true
            $Path.Add($GroupDN) | Out-Null
            
            # Durchlaufe alle Parent-Gruppen
            foreach ($parentDN in $groupData.MemberOf) {
                Find-CircularReference -GroupDN $parentDN -Path $Path
            }
            
            # Entferne aus Stack
            $groupData.InStack = $false
            $Path.RemoveAt($Path.Count - 1)
        }
        
        # Suche nach zirkulären Referenzen für alle Gruppen
        foreach ($groupDN in $groupMemberships.Keys) {
            if (-not $groupMemberships[$groupDN].Visited) {
                $path = New-Object System.Collections.ArrayList
                Find-CircularReference -GroupDN $groupDN -Path $path
            }
        }
        
        if ($circularReferences.Count -eq 0) {
            Write-ADReportLog -Message "No circular group references found." -Type Info -Terminal
            return @([PSCustomObject]@{
                CircularPath = "None Found"
                NumberOfGroups = 0
                FirstGroup = "N/A"
                GroupsInvolved = "No circular references detected"
                Risk = "None"
                Recommendation = "Environment is healthy - no circular references"
            })
        }
        
        Write-ADReportLog -Message "Circular group reference analysis completed. $($circularReferences.Count) circular reference(s) found." -Type Info -Terminal
        return $circularReferences | Sort-Object NumberOfGroups, FirstGroup
        
    } catch {
        Write-ADReportLog -Message "Error analyzing circular group references: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-GroupsByTypeScope {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing groups by type and scope..." -Type Info -Terminal
        
        $Groups = Get-ADGroup -Filter * -Properties GroupCategory, GroupScope, ManagedBy, whenCreated, whenChanged, Members -ErrorAction Stop
        
        # Gruppiere nach Type und Scope
        $GroupStats = $Groups | Group-Object -Property GroupCategory, GroupScope
        
        $Results = foreach ($statGroup in $GroupStats) {
            $category = $statGroup.Group[0].GroupCategory
            $scope = $statGroup.Group[0].GroupScope
            
            # Berechne durchschnittliche Mitgliederzahl
            $memberCounts = @()
            foreach ($group in $statGroup.Group) {
                $memberCount = @(Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue).Count
                $memberCounts += $memberCount
            }
            
            $avgMembers = if ($memberCounts.Count -gt 0) { 
                [math]::Round(($memberCounts | Measure-Object -Average).Average, 1) 
            } else { 0 }
            
            $maxMembers = if ($memberCounts.Count -gt 0) { 
                ($memberCounts | Measure-Object -Maximum).Maximum 
            } else { 0 }
            
            # Best Practice Empfehlungen
            $recommendation = switch ("$category-$scope") {
                "Security-Global" { "Standard for most security groups" }
                "Security-DomainLocal" { "Good for resource permissions" }
                "Security-Universal" { "Use for cross-forest scenarios" }
                "Distribution-Global" { "Standard for email distribution" }
                "Distribution-DomainLocal" { "Rarely used - review necessity" }
                "Distribution-Universal" { "Good for cross-forest distribution" }
                default { "Review group configuration" }
            }
            
            [PSCustomObject]@{
                GroupCategory = $category
                GroupScope = $scope
                Count = $statGroup.Count
                Percentage = [math]::Round(($statGroup.Count / $Groups.Count) * 100, 2)
                AverageMemberCount = $avgMembers
                MaxMemberCount = $maxMembers
                Recommendation = $recommendation
                Examples = ($statGroup.Group | Select-Object -First 3 | ForEach-Object { $_.Name }) -join ", "
            }
        }
        
        Write-ADReportLog -Message "Groups by type/scope analysis completed. $($Results.Count) combinations found." -Type Info -Terminal
        return $Results | Sort-Object GroupCategory, GroupScope
        
    } catch {
        Write-ADReportLog -Message "Error analyzing groups by type and scope: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-DynamicDistGroups {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing dynamic distribution groups..." -Type Info -Terminal
        
        # Hinweis: Dynamic Distribution Groups sind ein Exchange-Feature
        # Diese Funktion würde Exchange-Cmdlets benötigen
        
        Write-ADReportLog -Message "Dynamic Distribution Groups require Exchange PowerShell module." -Type Warning
        
        return @([PSCustomObject]@{
            Name = "Not Available"
            Status = "Exchange Required"
            Description = "Dynamic Distribution Groups are an Exchange feature"
            Recommendation = "Use Exchange Management Shell for this report"
        })
        
    } catch {
        Write-ADReportLog -Message "Error analyzing dynamic distribution groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-MailEnabledGroups {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing mail-enabled groups..." -Type Info -Terminal
        
        # Suche nach Gruppen mit Mail-Attributen
        # Prüfe, ob Exchange-Attribute verfügbar sind
        $exchangeInstalled = $false
        try {
            $testGroup = Get-ADGroup -Filter * -Properties msExchRecipientTypeDetails -ErrorAction Stop | Select-Object -First 1
            if ($null -ne $testGroup.PSObject.Properties['msExchRecipientTypeDetails']) {
                $exchangeInstalled = $true
            }
        } catch {
            $exchangeInstalled = $false
        }
        
        # Basis-Properties, die immer verfügbar sind
        $baseProperties = @('mail', 'proxyAddresses', 'GroupCategory', 'GroupScope', 'ManagedBy')
        
        # Exchange-spezifische Properties nur wenn Exchange installiert ist
        if ($exchangeInstalled) {
            $properties = $baseProperties + @('legacyExchangeDN', 'msExchRecipientTypeDetails')
        } else {
            $properties = $baseProperties
        }
        
        $MailGroups = Get-ADGroup -Filter "mail -like '*'" -Properties $properties -ErrorAction Stop
        
        if ($MailGroups.Count -eq 0) {
            Write-ADReportLog -Message "No mail-enabled groups found." -Type Info -Terminal
            return @([PSCustomObject]@{
                Name = "No Results"
                Mail = "N/A"
                GroupCategory = "N/A"
                Status = "No mail-enabled groups found"
                Recommendation = "Mail-enabled groups may be managed through Exchange"
            })
        }
        
        $Results = foreach ($group in $MailGroups) {
            # Prüfe Gruppenmitgliederzahl
            $memberCount = @(Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue).Count
            
            # Analysiere Mail-Eigenschaften
            $hasMultipleProxies = ($group.proxyAddresses -and $group.proxyAddresses.Count -gt 1)
            $primarySMTP = $group.proxyAddresses | Where-Object { $_ -clike "SMTP:*" } | Select-Object -First 1
            
            # Manager auflösen
            $managerName = "None"
            if ($group.ManagedBy) {
                try {
                    $manager = Get-ADUser -Identity $group.ManagedBy -Properties DisplayName -ErrorAction SilentlyContinue
                    if ($manager) {
                        $managerName = $manager.DisplayName
                    }
                } catch {
                    $managerName = "Unknown"
                }
            }
            
            # Status und Empfehlungen
            $status = "Active"
            $recommendations = @()
            
            if ([string]::IsNullOrWhiteSpace($group.mail)) {
                $status = "Incomplete"
                $recommendations += "Mail attribute is empty"
            }
            
            if ($memberCount -eq 0) {
                $recommendations += "Group has no members"
            }
            
            if ($memberCount -gt 1000) {
                $recommendations += "Large distribution list - consider breaking up"
            }
            
            if ([string]::IsNullOrWhiteSpace($group.ManagedBy)) {
                $recommendations += "No manager assigned"
            }
            
            [PSCustomObject]@{
                Name = $group.Name
                Mail = $group.mail
                PrimarySMTP = if ($primarySMTP) { $primarySMTP -replace "SMTP:", "" } else { $group.mail }
                GroupCategory = $group.GroupCategory
                GroupScope = $group.GroupScope
                MemberCount = $memberCount
                ManagedBy = $managerName
                ProxyAddressCount = if ($group.proxyAddresses) { $group.proxyAddresses.Count } else { 0 }
                Status = $status
                Recommendations = if ($recommendations) { $recommendations -join "; " } else { "Properly configured" }
            }
        }
        
        Write-ADReportLog -Message "Mail-enabled groups analysis completed. $($Results.Count) groups found." -Type Info -Terminal
        return $Results | Sort-Object MemberCount -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing mail-enabled groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-GroupsWithoutOwners {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing groups without owners/managers..." -Type Info -Terminal
        
        $Groups = Get-ADGroup -Filter * -Properties ManagedBy, GroupCategory, GroupScope, whenCreated, whenChanged, info -ErrorAction Stop
        
        # Filtere Gruppen ohne ManagedBy
        $UnmanagedGroups = $Groups | Where-Object { [string]::IsNullOrWhiteSpace($_.ManagedBy) }
        
        $Results = foreach ($group in $UnmanagedGroups) {
            # Prüfe Gruppengröße
            $memberCount = @(Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue).Count
            
            # Alter der Gruppe
            $ageInDays = if ($group.whenCreated) { 
                (New-TimeSpan -Start $group.whenCreated -End (Get-Date)).Days 
            } else { 0 }
            
            # Risikobewertung
            $riskLevel = "Low"
            $riskFactors = @()
            
            if ($group.GroupCategory -eq "Security") {
                $riskLevel = "Medium"
                $riskFactors += "Security group"
            }
            
            if ($memberCount -gt 50) {
                if ($riskLevel -eq "Low") { $riskLevel = "Medium" }
                $riskFactors += "Large membership ($memberCount members)"
            }
            
            if ($memberCount -gt 100 -and $group.GroupCategory -eq "Security") {
                $riskLevel = "High"
            }
            
            if ($ageInDays -gt 365) {
                $riskFactors += "Old group (>1 year)"
            }
            
            # Empfehlungen
            $recommendations = @("Assign a group manager/owner")
            
            if ($memberCount -eq 0 -and $ageInDays -gt 90) {
                $recommendations += "Consider deleting empty old group"
            }
            
            if ($group.GroupCategory -eq "Security" -and $memberCount -gt 20) {
                $recommendations += "Security groups should have designated owners"
            }
            
            [PSCustomObject]@{
                Name = $group.Name
                GroupCategory = $group.GroupCategory
                GroupScope = $group.GroupScope
                MemberCount = $memberCount
                Description = if ($group.Description) { $group.Description } else { "(No description)" }
                WhenCreated = $group.whenCreated
                AgeInDays = $ageInDays
                LastModified = $group.whenChanged
                RiskLevel = $riskLevel
                RiskFactors = if ($riskFactors) { $riskFactors -join "; " } else { "None" }
                Recommendations = $recommendations -join "; "
            }
        }
        
        Write-ADReportLog -Message "Groups without owners analysis completed. $($Results.Count) groups found." -Type Info -Terminal
        return $Results | Sort-Object RiskLevel, MemberCount -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing groups without owners: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-LargeGroups {
    [CmdletBinding()]
    param([int]$Threshold = 100)
    
    try {
        Write-ADReportLog -Message "Analyzing large groups (threshold: $Threshold members)..." -Type Info -Terminal
        
        $Groups = Get-ADGroup -Filter * -Properties GroupCategory, GroupScope, ManagedBy, whenCreated -ErrorAction Stop
        
        $LargeGroups = @()
        
        foreach ($group in $Groups) {
            $memberCount = @(Get-ADGroupMember -Identity $group.DistinguishedName -Recursive -ErrorAction SilentlyContinue).Count
            
            if ($memberCount -ge $Threshold) {
                # Nested group analysis
                $directMembers = @(Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue)
                $nestedGroups = @($directMembers | Where-Object { $_.objectClass -eq "group" })
                
                # Manager info
                $managerName = "None"
                if ($group.ManagedBy) {
                    try {
                        $manager = Get-ADUser -Identity $group.ManagedBy -Properties DisplayName -ErrorAction SilentlyContinue
                        if ($manager) {
                            $managerName = $manager.DisplayName
                        }
                    } catch {
                        $managerName = "Unknown"
                    }
                }
                
                # Performance impact assessment
                $performanceImpact = "Low"
                if ($memberCount -gt 1000) {
                    $performanceImpact = "High"
                } elseif ($memberCount -gt 500) {
                    $performanceImpact = "Medium"
                }
                
                # Recommendations
                $recommendations = @()
                if ($memberCount -gt 5000) {
                    $recommendations += "Consider breaking into smaller groups"
                }
                if ($nestedGroups.Count -gt 10) {
                    $recommendations += "High number of nested groups may impact performance"
                }
                if ($group.GroupScope -eq "Global" -and $memberCount -gt 5000) {
                    $recommendations += "Consider using Universal scope for large groups"
                }
                
                $LargeGroups += [PSCustomObject]@{
                    Name = $group.Name
                    GroupCategory = $group.GroupCategory
                    GroupScope = $group.GroupScope
                    TotalMembers = $memberCount
                    DirectMembers = $directMembers.Count
                    NestedGroups = $nestedGroups.Count
                    ManagedBy = $managerName
                    WhenCreated = $group.whenCreated
                    PerformanceImpact = $performanceImpact
                    Recommendations = if ($recommendations) { $recommendations -join "; " } else { "Monitor group size" }
                }
            }
        }
        
        if ($LargeGroups.Count -eq 0) {
            Write-ADReportLog -Message "No large groups found (threshold: $Threshold members)." -Type Info -Terminal
            return @([PSCustomObject]@{
                Name = "No Results"
                TotalMembers = 0
                Status = "No groups with $Threshold or more members found"
                Recommendations = "All groups are within size limits"
            })
        }
        
        Write-ADReportLog -Message "Large groups analysis completed. $($LargeGroups.Count) groups found." -Type Info -Terminal
        return $LargeGroups | Sort-Object TotalMembers -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing large groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-RecentlyModifiedGroups {
    [CmdletBinding()]
    param([int]$Days = 7)
    
    try {
        Write-ADReportLog -Message "Analyzing recently modified groups (last $Days days)..." -Type Info -Terminal
        
        $CutoffDate = (Get-Date).AddDays(-$Days)
        $Groups = Get-ADGroup -Filter "whenChanged -gt '$CutoffDate'" -Properties whenChanged, whenCreated, ManagedBy, GroupCategory, GroupScope, modifyTimeStamp -ErrorAction Stop
        
        if ($Groups.Count -eq 0) {
            Write-ADReportLog -Message "No groups modified in the last $Days days." -Type Info -Terminal
            return @([PSCustomObject]@{
                Name = "No Results"
                Status = "No groups modified in the last $Days days"
                LastModified = "N/A"
                Recommendations = "Normal - no recent group modifications"
            })
        }
        
        $Results = foreach ($group in $Groups) {
            # Versuche die Art der Änderung zu ermitteln
            $changeType = "Modified"
            if ($group.whenCreated -and $group.whenCreated -gt $CutoffDate) {
                $changeType = "Created"
            }
            
            # Manager info
            $managerName = "None"
            if ($group.ManagedBy) {
                try {
                    $manager = Get-ADUser -Identity $group.ManagedBy -Properties DisplayName -ErrorAction SilentlyContinue
                    if ($manager) {
                        $managerName = $manager.DisplayName
                    }
                } catch {
                    $managerName = "Unknown"
                }
            }
            
            # Mitgliederzahl für Kontext
            $memberCount = @(Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction SilentlyContinue).Count
            
            # Tage seit Änderung
            $daysSinceChange = [math]::Round((New-TimeSpan -Start $group.whenChanged -End (Get-Date)).TotalDays, 1)
            
            [PSCustomObject]@{
                Name = $group.Name
                GroupCategory = $group.GroupCategory
                GroupScope = $group.GroupScope
                ChangeType = $changeType
                LastModified = $group.whenChanged
                DaysSinceChange = $daysSinceChange
                CreatedDate = $group.whenCreated
                MemberCount = $memberCount
                ManagedBy = $managerName
                ModificationTime = $group.whenChanged.ToString("HH:mm:ss")
            }
        }
        
        Write-ADReportLog -Message "Recently modified groups analysis completed. $($Results.Count) groups found." -Type Info -Terminal
        return $Results | Sort-Object LastModified -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing recently modified groups: $($_.Exception.Message)" -Type Error
        return @()
    }
}

# --- Weitere Computer-Report Funktionen ---
Function Get-ComputersByOSVersion {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing computers by OS version..." -Type Info -Terminal
        
        $Computers = Get-ADComputer -Filter * -Properties OperatingSystem, OperatingSystemVersion, OperatingSystemServicePack, Name -ErrorAction Stop
        
        $Results = foreach ($computer in $Computers) {
            # Support-Status ermitteln
            $supportStatus = "Unknown"
            $eolDate = "Unknown"
            
            if ($computer.OperatingSystem -like "*Windows 7*") {
                $supportStatus = "End of Life"
                $eolDate = "2020-01-14"
            } elseif ($computer.OperatingSystem -like "*Windows 8*" -and $computer.OperatingSystem -notlike "*8.1*") {
                $supportStatus = "End of Life"
                $eolDate = "2016-01-12"
            } elseif ($computer.OperatingSystem -like "*Windows 8.1*") {
                $supportStatus = "End of Life"
                $eolDate = "2023-01-10"
            } elseif ($computer.OperatingSystem -like "*Windows 10*") {
                $supportStatus = "Check Version"
                $eolDate = "Version-dependent"
            } elseif ($computer.OperatingSystem -like "*Windows 11*") {
                $supportStatus = "Supported"
                $eolDate = "Active"
            } elseif ($computer.OperatingSystem -like "*Server 2008*" -and $computer.OperatingSystem -notlike "*R2*") {
                $supportStatus = "End of Life"
                $eolDate = "2015-07-14"
            } elseif ($computer.OperatingSystem -like "*Server 2008 R2*") {
                $supportStatus = "End of Life"
                $eolDate = "2020-01-14"
            } elseif ($computer.OperatingSystem -like "*Server 2012*" -and $computer.OperatingSystem -notlike "*R2*") {
                $supportStatus = "End of Life"
                $eolDate = "2023-10-10"
            } elseif ($computer.OperatingSystem -like "*Server 2012 R2*") {
                $supportStatus = "End of Life"
                $eolDate = "2023-10-10"
            } elseif ($computer.OperatingSystem -like "*Server 2016*") {
                $supportStatus = "Supported"
                $eolDate = "2027-01-12"
            } elseif ($computer.OperatingSystem -like "*Server 2019*") {
                $supportStatus = "Supported"
                $eolDate = "2029-01-09"
            } elseif ($computer.OperatingSystem -like "*Server 2022*") {
                $supportStatus = "Supported"
                $eolDate = "2031-10-14"
            }
            
            [PSCustomObject]@{
                ComputerName = $computer.Name
                OperatingSystem = if ($computer.OperatingSystem) { $computer.OperatingSystem } else { "Unknown" }
                Version = if ($computer.OperatingSystemVersion) { $computer.OperatingSystemVersion } else { "Unknown" }
                ServicePack = if ($computer.OperatingSystemServicePack) { $computer.OperatingSystemServicePack } else { "None" }
                SupportStatus = $supportStatus
                EndOfLifeDate = $eolDate
            }
        }
        
        Write-ADReportLog -Message "Computers by OS version analysis completed. $($Results.Count) computers found." -Type Info -Terminal
        return $Results | Sort-Object OperatingSystem, Version
        
    } catch {
        Write-ADReportLog -Message "Error analyzing computers by OS version: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-OSSummary {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing OS summary..." -Type Info -Terminal
        
        $Computers = Get-ADComputer -Filter * -Properties OperatingSystem, OperatingSystemVersion, OperatingSystemServicePack, LastLogonDate, Enabled -ErrorAction Stop
        
        # Gruppiere nach OS
        $OSGroups = $Computers | Group-Object OperatingSystem
        
        $Results = foreach ($osGroup in $OSGroups) {
            $os = if ([string]::IsNullOrWhiteSpace($osGroup.Name)) { "Unknown" } else { $osGroup.Name }
            
            # Statistiken für dieses OS
            $enabledCount = @($osGroup.Group | Where-Object { $_.Enabled }).Count
            $activeCount = @($osGroup.Group | Where-Object { 
                $_.LastLogonDate -and $_.LastLogonDate -gt (Get-Date).AddDays(-30) 
            }).Count
            
            # Support-Status ermitteln
            $supportStatus = "Unknown"
            $eolDate = "Unknown"
            
            if ($os -like "*Windows 7*") {
                $supportStatus = "End of Life"
                $eolDate = "2020-01-14"
            } elseif ($os -like "*Windows 8*" -and $os -notlike "*8.1*") {
                $supportStatus = "End of Life"
                $eolDate = "2016-01-12"
            } elseif ($os -like "*Windows 8.1*") {
                $supportStatus = "End of Life"
                $eolDate = "2023-01-10"
            } elseif ($os -like "*Windows 10*") {
                $supportStatus = "Check Version"
                $eolDate = "Version-dependent"
            } elseif ($os -like "*Windows 11*") {
                $supportStatus = "Supported"
                $eolDate = "Active"
            } elseif ($os -like "*Server 2008*" -and $os -notlike "*R2*") {
                $supportStatus = "End of Life"
                $eolDate = "2015-07-14"
            } elseif ($os -like "*Server 2008 R2*") {
                $supportStatus = "End of Life"
                $eolDate = "2020-01-14"
            } elseif ($os -like "*Server 2012*" -and $os -notlike "*R2*") {
                $supportStatus = "End of Life"
                $eolDate = "2023-10-10"
            } elseif ($os -like "*Server 2012 R2*") {
                $supportStatus = "End of Life"
                $eolDate = "2023-10-10"
            } elseif ($os -like "*Server 2016*") {
                $supportStatus = "Supported"
                $eolDate = "2027-01-12"
            } elseif ($os -like "*Server 2019*") {
                $supportStatus = "Supported"
                $eolDate = "2029-01-09"
            } elseif ($os -like "*Server 2022*") {
                $supportStatus = "Supported"
                $eolDate = "2031-10-14"
            }
            
            # Risikobewertung
            $riskLevel = "Low"
            if ($supportStatus -eq "End of Life") {
                $riskLevel = "Critical"
            } elseif ($supportStatus -eq "Check Version") {
                $riskLevel = "Medium"
            }
            
            [PSCustomObject]@{
                OperatingSystem = $os
                Count = $osGroup.Count
                EnabledCount = $enabledCount
                ActiveCount = $activeCount
                InactiveCount = $osGroup.Count - $activeCount
                Percentage = [math]::Round(($osGroup.Count / $Computers.Count) * 100, 2)
                SupportStatus = $supportStatus
                EndOfLifeDate = $eolDate
                RiskLevel = $riskLevel
                Recommendation = if ($supportStatus -eq "End of Life") { 
                    "Urgent: Upgrade or replace systems" 
                } elseif ($supportStatus -eq "Check Version") { 
                    "Verify specific version support status" 
                } else { 
                    "Keep systems updated" 
                }
            }
        }
        
        Write-ADReportLog -Message "OS summary analysis completed. $($Results.Count) unique OS versions found." -Type Info -Terminal
        return $Results | Sort-Object RiskLevel, Count -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing OS summary: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-BitLockerStatus {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing BitLocker status..." -Type Info -Terminal
        
        # Nur die relevanten BitLocker-Attribute abfragen
        $Computers = Get-ADComputer -Filter * -Properties Name, DistinguishedName -ErrorAction Stop
        
        $Results = foreach ($computer in $Computers) {
            $bitlockerStatus = "Unknown"
            $recoveryKey = "Not Found"
            
            # BitLocker-Recovery-Informationen suchen
            try {
                $recoveryObjects = Get-ADObject -Filter {objectClass -eq 'msFVE-RecoveryInformation'} -SearchBase $computer.DistinguishedName -Properties "msFVE-RecoveryPassword" -ErrorAction Stop
                
                if ($recoveryObjects) {
                    $recoveryKey = $recoveryObjects."msFVE-RecoveryPassword"
                    $bitlockerStatus = "Enabled"
                }
            }
            catch {
                Write-ADReportLog -Message "Warning: Could not retrieve BitLocker recovery info for $($computer.Name): $($_.Exception.Message)" -Type Warning
            }
            
            [PSCustomObject]@{
                ComputerName = $computer.Name
                BitLockerStatus = $bitlockerStatus
                RecoveryKey = $recoveryKey
                Recommendation = if ($bitlockerStatus -eq "Unknown") {
                    "Enable BitLocker encryption"
                } elseif ($recoveryKey -eq "Not Found") {
                    "BitLocker enabled but recovery key not stored in AD"
                } else {
                    "BitLocker properly configured with recovery key"
                }
            }
        }
        
        Write-ADReportLog -Message "BitLocker status analysis completed. $($Results.Count) computers analyzed." -Type Info -Terminal
        return $Results
        
    } catch {
        Write-ADReportLog -Message "Error analyzing BitLocker status: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-StaleComputerPasswords {
    [CmdletBinding()]
    param([int]$Days = 90)
    
    try {
        Write-ADReportLog -Message "Analyzing stale computer passwords..." -Type Info -Terminal
        
        $CutoffDate = (Get-Date).AddDays(-$Days)
        $Computers = Get-ADComputer -Filter * -Properties PasswordLastSet, LastLogonDate, OperatingSystem -ErrorAction Stop
        
        $Results = foreach ($computer in $Computers) {
            if ($computer.PasswordLastSet -lt $CutoffDate) {
                [PSCustomObject]@{
                    ComputerName = $computer.Name
                    PasswordLastSet = $computer.PasswordLastSet
                    DaysSinceLastChange = [math]::Round((New-TimeSpan -Start $computer.PasswordLastSet -End (Get-Date)).TotalDays, 1)
                    LastLogon = $computer.LastLogonDate
                    OperatingSystem = $computer.OperatingSystem
                    RiskLevel = if ((New-TimeSpan -Start $computer.PasswordLastSet -End (Get-Date)).TotalDays -gt 180) { "High" } else { "Medium" }
                    Recommendation = "Reset computer password and verify network connectivity"
                }
            }
        }
        
        Write-ADReportLog -Message "Stale computer passwords analysis completed. $($Results.Count) computers found." -Type Info -Terminal
        return $Results | Sort-Object DaysSinceLastChange -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing stale computer passwords: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-ComputersNeverLoggedOn {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing computers that never logged on..." -Type Info -Terminal
        
        $Computers = Get-ADComputer -Filter "LastLogonDate -notlike '*'" -Properties LastLogonDate, whenCreated, OperatingSystem, Enabled, Description, DistinguishedName -ErrorAction Stop
        
        if ($Computers.Count -eq 0) {
            Write-ADReportLog -Message "No computers found that never logged on." -Type Info -Terminal
            return @([PSCustomObject]@{
                Name = "No Results"
                Status = "All computers have logged on at least once"
                Recommendation = "Environment is healthy"
            })
        }
        
        $Results = foreach ($computer in $Computers) {
            # Alter des Computer-Objekts
            $ageInDays = if ($computer.whenCreated) {
                (New-TimeSpan -Start $computer.whenCreated -End (Get-Date)).Days
            } else { "Unknown" }
            
            # OU-Pfad extrahieren
            $ouPath = if ($computer.DistinguishedName -match 'CN=[^,]+,(.+)$') { $matches[1] } else { "Unknown" }
            
            # Risikobewertung
            $riskLevel = "Low"
            $recommendations = @()
            
            if ($computer.Enabled) {
                $riskLevel = "Medium"
                $recommendations += "Consider disabling unused computer"
            }
            
            if ($ageInDays -is [int] -and $ageInDays -gt 90) {
                if ($riskLevel -eq "Low") { $riskLevel = "Medium" }
                if ($computer.Enabled) { $riskLevel = "High" }
                $recommendations += "Old unused computer account - consider deletion"
            }
            
            if ($ageInDays -is [int] -and $ageInDays -gt 180) {
                $riskLevel = "High"
                $recommendations = @("Delete unused computer account")
            }
            
            [PSCustomObject]@{
                Name = $computer.Name
                OperatingSystem = if ($computer.OperatingSystem) { $computer.OperatingSystem } else { "Unknown" }
                Enabled = $computer.Enabled
                WhenCreated = $computer.whenCreated
                AgeInDays = $ageInDays
                Description = if ($computer.Description) { $computer.Description } else { "(No description)" }
                OrganizationalUnit = $ouPath
                RiskLevel = $riskLevel
                Recommendations = if ($recommendations) { $recommendations -join "; " } else { "Monitor account" }
                Status = if ($computer.Enabled) { "Enabled but never used" } else { "Disabled and never used" }
            }
        }
        
        Write-ADReportLog -Message "Computers never logged on analysis completed. $($Results.Count) computers found." -Type Info -Terminal
        return $Results | Sort-Object RiskLevel, AgeInDays -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing computers never logged on: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-DuplicateComputerNames {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing duplicate computer names..." -Type Info -Terminal
        
        $Computers = Get-ADComputer -Filter * -Properties DNSHostName, Enabled, OperatingSystem, LastLogonDate, whenCreated, DistinguishedName, IPv4Address -ErrorAction Stop
        
        # Suche nach Duplikaten basierend auf Name (ohne $)
        $ComputerNames = $Computers | ForEach-Object { 
            [PSCustomObject]@{
                Computer = $_
                CleanName = $_.Name.TrimEnd('$').ToUpper()
            }
        }
        
        # Gruppiere nach bereinigtem Namen
        $DuplicateGroups = $ComputerNames | Group-Object CleanName | Where-Object { $_.Count -gt 1 }
        
        if ($DuplicateGroups.Count -eq 0) {
            Write-ADReportLog -Message "No duplicate computer names found." -Type Info -Terminal
            return @([PSCustomObject]@{
                Name = "No Duplicates"
                Status = "No duplicate computer names detected"
                Count = 0
                Recommendation = "Environment is healthy"
            })
        }
        
        $Results = foreach ($dupGroup in $DuplicateGroups) {
            foreach ($item in $dupGroup.Group) {
                $computer = $item.Computer
                
                # OU-Pfad extrahieren
                $ouPath = if ($computer.DistinguishedName -match 'CN=[^,]+,(.+)$') { $matches[1] } else { "Unknown" }
                
                # Bestimme welcher Computer der "aktive" ist
                $isActive = $computer.Enabled -and $computer.LastLogonDate -and 
                           $computer.LastLogonDate -gt (Get-Date).AddDays(-30)
                
                [PSCustomObject]@{
                    Name = $computer.Name
                    CleanName = $item.CleanName
                    DNSHostName = $computer.DNSHostName
                    IPv4Address = if ($computer.IPv4Address) { $computer.IPv4Address } else { "No IP" }
                    Enabled = $computer.Enabled
                    OperatingSystem = if ($computer.OperatingSystem) { $computer.OperatingSystem } else { "Unknown" }
                    LastLogonDate = $computer.LastLogonDate
                    WhenCreated = $computer.whenCreated
                    OrganizationalUnit = $ouPath
                    DuplicateCount = $dupGroup.Count
                    Status = if ($isActive) { "Active Duplicate" } 
                            elseif ($computer.Enabled) { "Enabled but Inactive" } 
                            else { "Disabled" }
                    Recommendation = if (-not $computer.Enabled -and (-not $computer.LastLogonDate -or 
                                      $computer.LastLogonDate -lt (Get-Date).AddDays(-90))) {
                                        "Delete disabled/inactive duplicate"
                                    } elseif ($dupGroup.Count -gt 2) {
                                        "Multiple duplicates - immediate investigation required"
                                    } else {
                                        "Investigate and resolve duplicate"
                                    }
                }
            }
        }
        
        Write-ADReportLog -Message "Duplicate computer names analysis completed. $($DuplicateGroups.Count) duplicate groups found." -Type Info -Terminal
        return $Results | Sort-Object CleanName, Status
        
    } catch {
        Write-ADReportLog -Message "Error analyzing duplicate computer names: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-ComputersByLocation {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing computers by location..." -Type Info -Terminal
        
        $Computers = Get-ADComputer -Filter * -Properties Location, Enabled, OperatingSystem, LastLogonDate, Description, DistinguishedName -ErrorAction Stop
        
        # Gruppiere nach Location
        $LocationGroups = $Computers | Group-Object Location
        
        $Results = foreach ($locGroup in $LocationGroups) {
            $location = if ([string]::IsNullOrWhiteSpace($locGroup.Name)) { "(No Location)" } else { $locGroup.Name }
            
            # Statistiken für diese Location
            $enabledCount = @($locGroup.Group | Where-Object { $_.Enabled }).Count
            $activeCount = @($locGroup.Group | Where-Object { 
                $_.LastLogonDate -and $_.LastLogonDate -gt (Get-Date).AddDays(-30) 
            }).Count
            
            # OS-Verteilung in dieser Location
            $osDistribution = $locGroup.Group | Group-Object OperatingSystem | 
                             Sort-Object Count -Descending | 
                             Select-Object -First 3 | 
                             ForEach-Object { "$($_.Name): $($_.Count)" }
            
            # Durchschnittliche Inaktivität
            $inactiveDays = @()
            foreach ($comp in $locGroup.Group) {
                if ($comp.LastLogonDate) {
                    $inactiveDays += (New-TimeSpan -Start $comp.LastLogonDate -End (Get-Date)).Days
                }
            }
            
            $avgInactiveDays = if ($inactiveDays.Count -gt 0) {
                [math]::Round(($inactiveDays | Measure-Object -Average).Average, 1)
            } else { "N/A" }
            
            [PSCustomObject]@{
                Location = $location
                TotalComputers = $locGroup.Count
                EnabledComputers = $enabledCount
                ActiveComputers = $activeCount
                InactiveComputers = $locGroup.Count - $activeCount
                Percentage = [math]::Round(($locGroup.Count / $Computers.Count) * 100, 2)
                TopOperatingSystems = if ($osDistribution) { $osDistribution -join "; " } else { "Unknown" }
                AverageInactiveDays = $avgInactiveDays
                HealthStatus = if ($activeCount -lt ($enabledCount * 0.5)) { "Poor" } 
                              elseif ($activeCount -lt ($enabledCount * 0.8)) { "Fair" } 
                              else { "Good" }
                Recommendation = if ($enabledCount -gt 0 -and $activeCount -lt ($enabledCount * 0.5)) {
                                   "Many inactive computers - review and cleanup"
                               } elseif ([string]::IsNullOrWhiteSpace($locGroup.Name)) {
                                   "Update location information for better tracking"
                               } else {
                                   "Location properly maintained"
                               }
            }
        }
        
        Write-ADReportLog -Message "Computers by location analysis completed. $($Results.Count) locations found." -Type Info -Terminal
        return $Results | Sort-Object TotalComputers -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing computers by location: $($_.Exception.Message)" -Type Error
        return @()
    }
}
Function Get-ServiceAccountsSPN {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing service accounts with SPNs..." -Type Info -Terminal
        
        # Service Accounts mit SPNs finden
        $SPNAccounts = Get-ADUser -Filter {ServicePrincipalName -like "*"} -Properties ServicePrincipalName, Description, LastLogonDate, PasswordLastSet, Enabled, whenCreated, Department -ErrorAction Stop
        
        $Results = foreach ($account in $SPNAccounts) {
            $daysSincePasswordChange = if ($account.PasswordLastSet) { 
                (New-TimeSpan -Start $account.PasswordLastSet -End (Get-Date)).Days 
            } else { 9999 }
            
            $daysSinceLastLogon = if ($account.LastLogonDate) { 
                (New-TimeSpan -Start $account.LastLogonDate -End (Get-Date)).Days 
            } else { 9999 }
            
            # Risikobewertung
            $riskScore = 0
            $riskFactors = @()
            
            if ($daysSincePasswordChange -gt 180) { $riskScore += 2; $riskFactors += "Old Password" }
            if ($daysSinceLastLogon -gt 90) { $riskScore += 2; $riskFactors += "Inactive" }
            if ($account.ServicePrincipalName.Count -gt 5) { $riskScore += 1; $riskFactors += "Multiple SPNs" }
            if ($account.Enabled -eq $false) { $riskScore -= 1; $riskFactors += "Disabled" }
            
            $riskLevel = switch ($riskScore) {
                {$_ -ge 4} { "High" }
                {$_ -ge 2} { "Medium" }
                default { "Low" }
            }
            
            [PSCustomObject]@{
                Name = $account.Name
                SamAccountName = $account.SamAccountName
                Description = $account.Description
                Department = $account.Department
                Enabled = $account.Enabled
                LastLogonDate = if ($account.LastLogonDate) { $account.LastLogonDate } else { "Never" }
                PasswordLastSet = if ($account.PasswordLastSet) { $account.PasswordLastSet } else { "Never" }
                SPNCount = $account.ServicePrincipalName.Count
                SPNs = $account.ServicePrincipalName -join "; "
                RiskLevel = $riskLevel
                RiskFactors = $riskFactors -join ", "
                WhenCreated = $account.whenCreated
            }
        }
        
        Write-ADReportLog -Message "Service accounts with SPNs analysis completed. $($Results.Count) accounts found." -Type Info -Terminal
        return $Results | Sort-Object RiskLevel, SPNCount -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing service accounts with SPNs: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-HighPrivServiceAccounts {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing high privileged service accounts..." -Type Info -Terminal
        
        # Define privileged groups in English and German
        $PrivGroups = @(
            # English groups
            "Domain Admins",
            "Enterprise Admins", 
            "Schema Admins",
            "Administrators",
            "Account Operators",
            "Backup Operators",
            "Server Operators",
            # German groups
            "Domänen-Admins",
            "Organisations-Admins",
            "Schema-Admins",
            "Administratoren",
            "Konten-Operatoren", 
            "Sicherungs-Operatoren",
            "Server-Operatoren"
        )
        
        $Results = @()
        foreach ($group in $PrivGroups) {
            try {
                $groupMembers = Get-ADGroupMember -Identity $group -Recursive -ErrorAction SilentlyContinue |
                    Where-Object {$_.objectClass -eq "user"} |
                    Get-ADUser -Properties Description, LastLogonDate, PasswordLastSet, Enabled, whenCreated, Department, ServicePrincipalName
                
                foreach ($account in $groupMembers) {
                    # Check for service accounts in both languages
                    if ($account.SamAccountName -like "*svc*" -or 
                        $account.SamAccountName -like "*service*" -or
                        $account.SamAccountName -like "*dienst*" -or
                        $account.Description -like "*service*" -or
                        $account.Description -like "*dienst*") {
                        
                        $daysSincePasswordChange = if ($account.PasswordLastSet) { 
                            (New-TimeSpan -Start $account.PasswordLastSet -End (Get-Date)).Days 
                        } else { 9999 }
                        
                        # Map German group names to English for consistent output
                        $groupNameEn = switch ($group) {
                            "Domänen-Admins" { "Domain Admins" }
                            "Organisations-Admins" { "Enterprise Admins" }
                            "Schema-Admins" { "Schema Admins" }
                            "Administratoren" { "Administrators" }
                            "Konten-Operatoren" { "Account Operators" }
                            "Sicherungs-Operatoren" { "Backup Operators" }
                            "Server-Operatoren" { "Server Operators" }
                            default { $group }
                        }
                        
                        $Results += [PSCustomObject]@{
                            Name = $account.Name
                            SamAccountName = $account.SamAccountName
                            Description = $account.Description
                            Department = $account.Department
                            PrivilegedGroup = $groupNameEn
                            Enabled = $account.Enabled
                            LastLogonDate = if ($account.LastLogonDate) { $account.LastLogonDate } else { "Never" }
                            PasswordLastSet = if ($account.PasswordLastSet) { $account.PasswordLastSet } else { "Never" }
                            DaysSincePasswordChange = $daysSincePasswordChange
                            HasSPN = $account.ServicePrincipalName.Count -gt 0
                            RiskLevel = if ($daysSincePasswordChange -gt 180) { "High" } 
                                      elseif ($daysSincePasswordChange -gt 90) { "Medium" }
                                      else { "Low" }
                            WhenCreated = $account.whenCreated
                        }
                    }
                }
            } catch {
                Write-ADReportLog -Message "Error processing group $group : $($_.Exception.Message)" -Type Warning
                continue
            }
        }
        
        # Remove duplicates that might occur from checking both English and German groups
        $Results = $Results | Sort-Object SamAccountName -Unique
        
        Write-ADReportLog -Message "High privileged service accounts analysis completed. $($Results.Count) accounts found." -Type Info -Terminal
        return $Results | Sort-Object RiskLevel, DaysSincePasswordChange -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing high privileged service accounts: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-ServiceAccountPasswordAge {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing service account password ages..." -Type Info -Terminal
        
        # Service Accounts anhand von Namenskonventionen und Beschreibungen identifizieren
        $ServiceAccounts = Get-ADUser -Filter {
            (SamAccountName -like "*svc*") -or 
            (SamAccountName -like "*service*") -or 
            (Description -like "*service*")
        } -Properties Description, LastLogonDate, PasswordLastSet, Enabled, whenCreated, Department -ErrorAction Stop
        
        $Results = foreach ($account in $ServiceAccounts) {
            $daysSincePasswordChange = if ($account.PasswordLastSet) { 
                (New-TimeSpan -Start $account.PasswordLastSet -End (Get-Date)).Days 
            } else { 9999 }
            
            $passwordStatus = switch ($daysSincePasswordChange) {
                {$_ -gt 365} { "Critical" }
                {$_ -gt 180} { "Warning" }
                {$_ -gt 90} { "Review" }
                default { "Good" }
            }
            
            [PSCustomObject]@{
                Name = $account.Name
                SamAccountName = $account.SamAccountName
                Description = $account.Description
                Department = $account.Department
                Enabled = $account.Enabled
                PasswordLastSet = if ($account.PasswordLastSet) { $account.PasswordLastSet } else { "Never" }
                DaysSincePasswordChange = $daysSincePasswordChange
                PasswordStatus = $passwordStatus
                LastLogonDate = if ($account.LastLogonDate) { $account.LastLogonDate } else { "Never" }
                WhenCreated = $account.whenCreated
                Recommendation = switch ($passwordStatus) {
                    "Critical" { "Immediate password change required" }
                    "Warning" { "Plan password change soon" }
                    "Review" { "Review password policy compliance" }
                    default { "Password age within policy" }
                }
            }
        }
        
        Write-ADReportLog -Message "Service account password age analysis completed. $($Results.Count) accounts found." -Type Info -Terminal
        return $Results | Sort-Object DaysSincePasswordChange -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing service account password ages: $($_.Exception.Message)" -Type Error
        return @()
    }
}

Function Get-UnusedServiceAccounts {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing unused service accounts..." -Type Info -Terminal
        
        # Service Accounts mit langer Inaktivität finden
        $ServiceAccounts = Get-ADUser -Filter {
            (SamAccountName -like "*svc*") -or 
            (SamAccountName -like "*service*") -or 
            (Description -like "*service*")
        } -Properties Description, LastLogonDate, PasswordLastSet, Enabled, whenCreated, Department -ErrorAction Stop
        
        $Results = foreach ($account in $ServiceAccounts) {
            $daysSinceLastLogon = if ($account.LastLogonDate) { 
                (New-TimeSpan -Start $account.LastLogonDate -End (Get-Date)).Days 
            } else { 9999 }
            
            if ($daysSinceLastLogon -gt 90 -or $null -eq $account.LastLogonDate) {
                $unusedStatus = switch ($daysSinceLastLogon) {
                    {$_ -gt 365} { "Critical" }
                    {$_ -gt 180} { "Warning" }
                    default { "Review" }
                }
                
                [PSCustomObject]@{
                    Name = $account.Name
                    SamAccountName = $account.SamAccountName
                    Description = $account.Description
                    Department = $account.Department
                    Enabled = $account.Enabled
                    LastLogonDate = if ($account.LastLogonDate) { $account.LastLogonDate } else { "Never" }
                    DaysSinceLastLogon = $daysSinceLastLogon
                    UnusedStatus = $unusedStatus
                    PasswordLastSet = if ($account.PasswordLastSet) { $account.PasswordLastSet } else { "Never" }
                    WhenCreated = $account.whenCreated
                    Recommendation = switch ($unusedStatus) {
                        "Critical" { "Consider immediate deactivation" }
                        "Warning" { "Verify if still needed" }
                        default { "Review account usage" }
                    }
                }
            }
        }
        
        Write-ADReportLog -Message "Unused service accounts analysis completed. $($Results.Count) accounts found." -Type Info -Terminal
        return $Results | Sort-Object DaysSinceLastLogon -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing unused service accounts: $($_.Exception.Message)" -Type Error
        return @()
    }
}
function Get-UnlinkedGPOs {
    try {
        Write-ADReportLog -Message "Analyzing unlinked GPOs..." -Type Info -Terminal
        
        # Alle GPOs abrufen
        $AllGPOs = Get-GPO -All -ErrorAction Stop
        
        $Results = foreach ($gpo in $AllGPOs) {
            # XML-Report generieren um Links zu prüfen
            $report = Get-GPOReport -Guid $gpo.Id -ReportType XML
            if ($report -notmatch "<LinksTo>") {
                [PSCustomObject]@{
                    Name = $gpo.DisplayName
                    ID = $gpo.Id
                    CreationTime = $gpo.CreationTime
                    ModificationTime = $gpo.ModificationTime
                    Status = $gpo.GpoStatus
                    WMIFilter = if ($gpo.WmiFilter) { $gpo.WmiFilter.Name } else { "None" }
                    Owner = $gpo.Owner
                    Description = $gpo.Description
                    Recommendation = "Review and consider deletion if not needed"
                }
            }
        }
        
        Write-ADReportLog -Message "Unlinked GPOs analysis completed. $($Results.Count) GPOs found." -Type Info -Terminal
        return $Results | Sort-Object ModificationTime -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing unlinked GPOs: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-EmptyGPOs {
    try {
        Write-ADReportLog -Message "Analyzing empty GPOs..." -Type Info -Terminal
        
        # Alle GPOs abrufen
        $AllGPOs = Get-GPO -All -ErrorAction Stop
        
        $Results = foreach ($gpo in $AllGPOs) {
            $report = [xml](Get-GPOReport -Guid $gpo.Id -ReportType Xml)
            
            # Prüfen ob Computer- und User-Einstellungen leer sind
            if (-not $report.GPO.Computer.ExtensionData -and -not $report.GPO.User.ExtensionData) {
                [PSCustomObject]@{
                    Name = $gpo.DisplayName
                    ID = $gpo.Id
                    CreationTime = $gpo.CreationTime
                    ModificationTime = $gpo.ModificationTime
                    Status = $gpo.GpoStatus
                    WMIFilter = if ($gpo.WmiFilter) { $gpo.WmiFilter.Name } else { "None" }
                    Owner = $gpo.Owner
                    Description = $gpo.Description
                    Recommendation = "Review and consider deletion"
                }
            }
        }
        
        Write-ADReportLog -Message "Empty GPOs analysis completed. $($Results.Count) GPOs found." -Type Info -Terminal
        return $Results | Sort-Object ModificationTime -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing empty GPOs: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-GPOPermissions {
    try {
        Write-ADReportLog -Message "Analyzing GPO permissions..." -Type Info -Terminal
        
        # Alle GPOs abrufen
        $AllGPOs = Get-GPO -All -ErrorAction Stop
        
        $Results = foreach ($gpo in $AllGPOs) {
            $permissions = Get-GPPermission -Guid $gpo.Id -All
            
            foreach ($perm in $permissions) {
                [PSCustomObject]@{
                    GPOName = $gpo.DisplayName
                    GPOID = $gpo.Id
                    Trustee = $perm.Trustee.Name
                    TrusteeType = $perm.Trustee.SidType
                    Permission = $perm.Permission
                    Inherited = $perm.Inherited
                    RiskLevel = switch ($perm.Permission) {
                        "GpoEditDeleteModifySecurity" { "High" }
                        "GpoEdit" { "Medium" }
                        default { "Low" }
                    }
                    Recommendation = switch ($perm.Permission) {
                        "GpoEditDeleteModifySecurity" { "Review full control permissions" }
                        "GpoEdit" { "Verify edit permissions are required" }
                        default { "Standard permission level" }
                    }
                }
            }
        }
        
        Write-ADReportLog -Message "GPO permissions analysis completed. $($Results.Count) permissions found." -Type Info -Terminal
        return $Results | Sort-Object GPOName,RiskLevel
        
    } catch {
        Write-ADReportLog -Message "Error analyzing GPO permissions: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-PasswordPolicySummary {
    try {
        Write-ADReportLog -Message "Analyzing password policies..." -Type Info -Terminal
        
        # Default Domain Policy abrufen
        $defaultPolicy = Get-ADDefaultDomainPasswordPolicy -ErrorAction Stop
        
        $Results = @([PSCustomObject]@{
            PolicyType = "Default Domain Policy"
            MinPasswordLength = $defaultPolicy.MinPasswordLength
            MaxPasswordAge = $defaultPolicy.MaxPasswordAge.Days
            MinPasswordAge = $defaultPolicy.MinPasswordAge.Days
            PasswordHistoryCount = $defaultPolicy.PasswordHistoryCount
            ComplexityEnabled = $defaultPolicy.ComplexityEnabled
            ReversibleEncryption = $defaultPolicy.ReversibleEncryptionEnabled
            LockoutThreshold = $defaultPolicy.LockoutThreshold
            LockoutDuration = if ($defaultPolicy.LockoutDuration) { $defaultPolicy.LockoutDuration.Minutes } else { 0 }
            ResetCounterAfter = if ($defaultPolicy.LockoutObservationWindow) { $defaultPolicy.LockoutObservationWindow.Minutes } else { 0 }
            RiskLevel = switch ($defaultPolicy.MinPasswordLength) {
                {$_ -lt 8} { "High" }
                {$_ -lt 12} { "Medium" }
                default { "Low" }
            }
            Recommendation = switch ($defaultPolicy.MinPasswordLength) {
                {$_ -lt 8} { "Increase minimum password length to at least 12 characters" }
                {$_ -lt 12} { "Consider increasing minimum password length" }
                default { "Policy meets basic security requirements" }
            }
        })
        
        Write-ADReportLog -Message "Password policy analysis completed." -Type Info -Terminal
        return $Results
        
    } catch {
        Write-ADReportLog -Message "Error analyzing password policies: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-AccountLockoutPolicies {
    try {
        Write-ADReportLog -Message "Analyzing account lockout policies..." -Type Info -Terminal
        
        # Default Domain Policy abrufen
        $defaultPolicy = Get-ADDefaultDomainPasswordPolicy -ErrorAction Stop
        
        $Results = @([PSCustomObject]@{
            PolicyName = "Default Domain Policy"
            LockoutThreshold = $defaultPolicy.LockoutThreshold
            LockoutDuration = if ($defaultPolicy.LockoutDuration) { $defaultPolicy.LockoutDuration.Minutes } else { 0 }
            ResetCounterAfter = if ($defaultPolicy.LockoutObservationWindow) { $defaultPolicy.LockoutObservationWindow.Minutes } else { 0 }
            RiskLevel = switch ($defaultPolicy.LockoutThreshold) {
                0 { "High" }
                {$_ -gt 10} { "Medium" }
                default { "Low" }
            }
            Recommendation = switch ($defaultPolicy.LockoutThreshold) {
                0 { "Enable account lockout policy" }
                {$_ -gt 10} { "Consider reducing lockout threshold" }
                default { "Policy meets security recommendations" }
            }
        })
        
        Write-ADReportLog -Message "Account lockout policy analysis completed." -Type Info -Terminal
        return $Results
        
    } catch {
        Write-ADReportLog -Message "Error analyzing account lockout policies: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-FineGrainedPasswordPolicies {
    try {
        Write-ADReportLog -Message "Analyzing fine-grained password policies..." -Type Info -Terminal
        
        # Fine-Grained Password Policies abrufen
        $FGPPs = Get-ADFineGrainedPasswordPolicy -Filter * -Properties * -ErrorAction Stop
        
        $Results = foreach ($fgpp in $FGPPs) {
            [PSCustomObject]@{
                Name = $fgpp.Name
                Precedence = $fgpp.Precedence
                MinPasswordLength = $fgpp.MinPasswordLength
                MaxPasswordAge = $fgpp.MaxPasswordAge.Days
                MinPasswordAge = $fgpp.MinPasswordAge.Days
                PasswordHistoryCount = $fgpp.PasswordHistoryCount
                ComplexityEnabled = $fgpp.ComplexityEnabled
                ReversibleEncryption = $fgpp.ReversibleEncryptionEnabled
                LockoutThreshold = $fgpp.LockoutThreshold
                LockoutDuration = if ($fgpp.LockoutDuration) { $fgpp.LockoutDuration.Minutes } else { 0 }
                ResetCounterAfter = if ($fgpp.LockoutObservationWindow) { $fgpp.LockoutObservationWindow.Minutes } else { 0 }
                AppliesTo = ($fgpp.AppliesTo | ForEach-Object { (Get-ADObject $_).Name }) -join "; "
                RiskLevel = switch ($fgpp.MinPasswordLength) {
                    {$_ -lt 8} { "High" }
                    {$_ -lt 12} { "Medium" }
                    default { "Low" }
                }
                Recommendation = switch ($fgpp.MinPasswordLength) {
                    {$_ -lt 8} { "Increase minimum password length to at least 12 characters" }
                    {$_ -lt 12} { "Consider increasing minimum password length" }
                    default { "Policy meets basic security requirements" }
                }
            }
        }
        
        Write-ADReportLog -Message "Fine-grained password policies analysis completed. $($Results.Count) policies found." -Type Info -Terminal
        return $Results | Sort-Object Precedence
        
    } catch {
        Write-ADReportLog -Message "Error analyzing fine-grained password policies: $($_.Exception.Message)" -Type Error
        return @()
    }
}
function Get-PrivilegeEscalationPaths {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing potential privilege escalation paths..." -Type Info -Terminal
        
        # Define privileged groups in German and English to support both languages
        $PrivGroups = @(
            # German Groups
            "Domänen-Admins",
            "Organisations-Admins", 
            "Schema-Admins",
            "Administratoren",
            "Konten-Operatoren",
            "Sicherungs-Operatoren",
            "Server-Operatoren",
            "Druck-Operatoren",
            # English Groups
            "Domain Admins",
            "Enterprise Admins",
            "Schema Admins", 
            "Administrators",
            "Account Operators",
            "Backup Operators",
            "Server Operators",
            "Print Operators"
        )

        # System accounts to exclude in both languages
        $ExcludedAccounts = @(
            # German
            "SYSTEM", "Administratoren", "Domänen-Admins", "Organisations-Admins", 
            "Jeder", "SELBST", "Terminalserver-Lizenzserver", "Zertifikatherausgeber",
            # English
            "Everyone", "SELF", "Terminal Server License Servers", "Certificate Publishers"
        ) -join '|'
        
        $Results = @()
        
        # Check all users with dangerous rights on privileged groups
        foreach ($group in $PrivGroups) {
            try {
                $groupObj = Get-ADGroup -Filter "Name -eq '$group'" -ErrorAction Stop
                if ($groupObj) {
                    $groupDN = $groupObj.DistinguishedName
                    
                    $acl = Get-Acl -Path "AD:\$groupDN" -ErrorAction Stop
                    $dangerousRights = $acl.Access | Where-Object {
                        ($_.ActiveDirectoryRights -match "WriteProperty|GenericWrite|WriteDacl|WriteOwner|GenericAll|ExtendedRight") -and
                        ($_.AccessControlType -eq "Allow") -and
                        ($_.IdentityReference.Value -notmatch $ExcludedAccounts)
                    }
                    
                    foreach ($right in $dangerousRights) {
                        $identity = $right.IdentityReference.Value.Split('\')[-1]
                        try {
                            $user = Get-ADUser -Filter "SamAccountName -eq '$identity' -or Name -eq '$identity'" -Properties Department,Description,LastLogonDate,whenCreated,memberOf -ErrorAction Stop
                            if ($user) {
                                $Results += [PSCustomObject]@{
                                    Name = $user.Name
                                    SamAccountName = $user.SamAccountName
                                    Department = $user.Department
                                    TargetGroup = $group
                                    AccessRight = $right.ActiveDirectoryRights
                                    RiskLevel = "High"
                                    Description = $user.Description
                                    LastLogon = if ($user.LastLogonDate) { $user.LastLogonDate } else { "Never" }
                                    Created = $user.whenCreated
                                    Recommendation = "Überprüfen und entfernen Sie gefährliche Zugriffsrechte / Review and remove dangerous access rights"
                                    AdditionalInfo = "Benutzer hat direkte Schreibrechte auf privilegierte Gruppe / User has direct write permissions on privileged group"
                                }
                            }
                        } catch {
                            # Skip warning for filtered system accounts
                            if ($identity -notmatch $ExcludedAccounts) {
                                Write-ADReportLog -Message "Error processing user $identity : $($_.Exception.Message)" -Type Warning
                            }
                        }
                    }
                }
            } catch {
                Write-ADReportLog -Message "Error accessing group $group : $($_.Exception.Message)" -Type Warning
            }
        }
        
        Write-ADReportLog -Message "Privilege escalation analysis completed. $($Results.Count) paths found." -Type Info -Terminal
        if ($Results.Count -eq 0) {
            Write-ADReportLog -Message "No privilege escalation paths detected." -Type Info
        }
        return $Results | Sort-Object RiskLevel,TargetGroup
        
    } catch {
        Write-ADReportLog -Message "Error analyzing privilege escalation paths: $($_.Exception.Message)" -Type Error
        return @()
    }
}
# End of Selection

function Get-ExposedCredentials {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing potentially exposed credentials..." -Type Info -Terminal
        
        $Results = @()
        
        # Benutzer mit reversible Verschlüsselung
        $reversibleUsers = Get-ADUser -Filter {AllowReversiblePasswordEncryption -eq $true} -Properties Department,Description,LastLogonDate,whenCreated
        foreach ($user in $reversibleUsers) {
            $Results += [PSCustomObject]@{
                Name = $user.Name
                SamAccountName = $user.SamAccountName
                Department = $user.Department
                ExposureType = "Reversible Encryption"
                RiskLevel = "High"
                Description = $user.Description
                LastLogon = if ($user.LastLogonDate) { $user.LastLogonDate } else { "Never" }
                Created = $user.whenCreated
                Recommendation = "Disable reversible password encryption"
            }
        }
        
        # Benutzer mit Kerberos DES
        $desUsers = Get-ADUser -Filter {UseDESKeyOnly -eq $true} -Properties Department,Description,LastLogonDate,whenCreated
        foreach ($user in $desUsers) {
            $Results += [PSCustomObject]@{
                Name = $user.Name
                SamAccountName = $user.SamAccountName
                Department = $user.Department
                ExposureType = "DES Encryption"
                RiskLevel = "High"
                Description = $user.Description
                LastLogon = if ($user.LastLogonDate) { $user.LastLogonDate } else { "Never" }
                Created = $user.whenCreated
                Recommendation = "Disable DES encryption usage"
            }
        }
        
        Write-ADReportLog -Message "Exposed credentials analysis completed. $($Results.Count) accounts found." -Type Info -Terminal
        return $Results | Sort-Object RiskLevel,ExposureType
        
    } catch {
        Write-ADReportLog -Message "Error analyzing exposed credentials: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-SuspiciousLogons {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing suspicious logon patterns..." -Type Info -Terminal
        
        $Results = @()
        
        # Benutzer mit ungewöhnlichen Anmeldezeiten (außerhalb 6-20 Uhr)
        $users = Get-ADUser -Filter * -Properties Department,Description,LastLogonDate,logonHours,whenCreated
        foreach ($user in $users) {
            if ($user.logonHours) {
                $logonHours = [System.BitConverter]::ToString($user.logonHours)
                if ($logonHours -match "FF") {
                    $Results += [PSCustomObject]@{
                        Name = $user.Name
                        SamAccountName = $user.SamAccountName
                        Department = $user.Department
                        Pattern = "24/7 Logon Hours"
                        RiskLevel = "Medium"
                        Description = $user.Description
                        LastLogon = if ($user.LastLogonDate) { $user.LastLogonDate } else { "Never" }
                        Created = $user.whenCreated
                        Recommendation = "Review logon hour restrictions"
                    }
                }
            }
        }
        
        Write-ADReportLog -Message "Suspicious logon analysis completed. $($Results.Count) patterns found." -Type Info -Terminal
        return $Results | Sort-Object RiskLevel,Pattern
        
    } catch {
        Write-ADReportLog -Message "Error analyzing suspicious logons: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-ForeignSecurityPrincipals {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing foreign security principals..." -Type Info -Terminal
        
        $FSPs = Get-ADObject -Filter {objectClass -eq "foreignSecurityPrincipal"} -Properties memberOf,whenCreated
        
        $Results = foreach ($fsp in $FSPs) {
            $sid = $fsp.Name
            $memberOfGroups = $fsp.memberOf | ForEach-Object { (Get-ADObject $_).Name }
            
            # Versuche den Kontonamen aus der SID zu ermitteln
            try {
                $account = New-Object System.Security.Principal.SecurityIdentifier($sid)
                $accountName = $account.Translate([System.Security.Principal.NTAccount]).Value
            }
            catch {
                $accountName = "Unknown Account"
            }
            
            [PSCustomObject]@{
                Name = $accountName 
                ObjectSID = $sid
                MemberOfGroups = $memberOfGroups -join "; "
                Created = $fsp.whenCreated
                RiskLevel = if ($memberOfGroups -match "Admin|Schema|Enterprise") { "High" } else { "Medium" }
                Recommendation = "Review foreign security principal memberships"
            }
        }
        
        Write-ADReportLog -Message "Foreign security principals analysis completed. $($Results.Count) principals found." -Type Info -Terminal
        return $Results | Sort-Object RiskLevel
        
    } catch {
        Write-ADReportLog -Message "Error analyzing foreign security principals: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-SIDHistoryAbuse {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing potential SID history abuse..." -Type Info -Terminal
        
        $Results = @()
        
        # Benutzer mit SID-History
        $users = Get-ADUser -Filter * -Properties SIDHistory,Department,Description,LastLogonDate,whenCreated
        foreach ($user in $users) {
            if ($user.SIDHistory) {
                $Results += [PSCustomObject]@{
                    Name = $user.Name
                    SamAccountName = $user.SamAccountName
                    Department = $user.Department
                    SIDHistoryCount = $user.SIDHistory.Count
                    RiskLevel = "High"
                    Description = $user.Description
                    LastLogon = if ($user.LastLogonDate) { $user.LastLogonDate } else { "Never" }
                    Created = $user.whenCreated
                    Recommendation = "Review and clean up SID history"
                }
            }
        }
        
        Write-ADReportLog -Message "SID history abuse analysis completed. $($Results.Count) accounts found." -Type Info -Terminal
        return $Results | Sort-Object SIDHistoryCount -Descending
        
    } catch {
        Write-ADReportLog -Message "Error analyzing SID history abuse: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-ACLAnalysis {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing ACL permissions..." -Type Info -Terminal
        
        $Results = @()
        
        # Alle AD-Objekte mit ACLs abrufen
        $objects = Get-ADObject -Filter * -Properties nTSecurityDescriptor
        foreach ($object in $objects) {
            $acl = $object.nTSecurityDescriptor
            
            # Prüfe auf ungewöhnliche ACL-Einträge
            foreach ($ace in $acl.Access) {
                if ($ace.AccessControlType -eq "Allow" -and 
                    ($ace.ActiveDirectoryRights -match "GenericAll|WriteDacl|WriteOwner")) {
                    
                    $Results += [PSCustomObject]@{
                        ObjectName = $object.Name
                        ObjectClass = $object.ObjectClass
                        IdentityReference = $ace.IdentityReference
                        AccessRights = $ace.ActiveDirectoryRights
                        RiskLevel = "High"
                        Recommendation = "Review high-risk permissions"
                    }
                }
            }
        }
        
        Write-ADReportLog -Message "ACL analysis completed. $($Results.Count) suspicious entries found." -Type Info -Terminal
        return $Results | Sort-Object RiskLevel
        
    } catch {
        Write-ADReportLog -Message "Error analyzing ACLs: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-InheritanceBreaks {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing inheritance breaks..." -Type Info -Terminal
        
        $Results = @()
        
        # Objekte mit deaktivierter Vererbung finden
        $objects = Get-ADObject -Filter * -Properties nTSecurityDescriptor
        foreach ($object in $objects) {
            $acl = $object.nTSecurityDescriptor
            
            if (-not $acl.AreAccessRulesProtected) {
                continue
            }
            
            $Results += [PSCustomObject]@{
                ObjectName = $object.Name
                ObjectClass = $object.ObjectClass
                Path = $object.DistinguishedName
                RiskLevel = "Medium"
                Recommendation = "Review inheritance settings"
            }
        }
        
        Write-ADReportLog -Message "Inheritance break analysis completed. $($Results.Count) objects found." -Type Info -Terminal
        return $Results
        
    } catch {
        Write-ADReportLog -Message "Error analyzing inheritance breaks: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-AdminSDHolderObjects {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing AdminSDHolder objects..." -Type Info -Terminal
        
        $Results = @()
        
        # AdminSDHolder geschützte Objekte finden
        $protectedObjects = Get-ADObject -LDAPFilter "(adminCount=1)" -Properties adminCount,whenCreated
        foreach ($object in $protectedObjects) {
            $Results += [PSCustomObject]@{
                Name = $object.Name
                ObjectClass = $object.ObjectClass
                Created = $object.whenCreated
                RiskLevel = "Medium"
                Recommendation = "Verify AdminSDHolder protection"
            }
        }
        
        Write-ADReportLog -Message "AdminSDHolder analysis completed. $($Results.Count) objects found." -Type Info -Terminal
        return $Results
        
    } catch {
        Write-ADReportLog -Message "Error analyzing AdminSDHolder objects: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-AdvancedDelegations {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing advanced delegations..." -Type Info -Terminal
        
        $Results = @()
        
        # Delegierte Berechtigungen analysieren
        $objects = Get-ADObject -Filter * -Properties nTSecurityDescriptor
        foreach ($object in $objects) {
            $acl = $object.nTSecurityDescriptor
            
            foreach ($ace in $acl.Access) {
                if ($ace.IsInherited -eq $false -and 
                    $ace.IdentityReference -notmatch "BUILTIN|NT AUTHORITY") {
                    
                    $Results += [PSCustomObject]@{
                        ObjectName = $object.Name
                        DelegatedTo = $ace.IdentityReference
                        Permissions = $ace.ActiveDirectoryRights
                        RiskLevel = "Medium"
                        Recommendation = "Review custom delegations"
                    }
                }
            }
        }
        
        Write-ADReportLog -Message "Advanced delegation analysis completed. $($Results.Count) delegations found." -Type Info -Terminal
        return $Results
        
    } catch {
        Write-ADReportLog -Message "Error analyzing advanced delegations: $($_.Exception.Message)" -Type Error
        return @()
    }
}

function Get-SchemaPermissions {
    [CmdletBinding()]
    param()
    
    try {
        Write-ADReportLog -Message "Analyzing schema permissions..." -Type Info -Terminal
        
        $Results = @()
        
        # Schema-Berechtigungen prüfen
        $schemaNC = (Get-ADRootDSE).schemaNamingContext
        
        # Standard-Filter "*" verwenden
        $filter = "*"
        
        $schemaObjects = Get-ADObject -SearchBase $schemaNC -SearchScope OneLevel -Filter "name -like '$filter'" -Properties nTSecurityDescriptor
        
        foreach ($object in $schemaObjects) {
            $acl = $object.nTSecurityDescriptor
            
            foreach ($ace in $acl.Access) {
                if ($ace.IsInherited -eq $false -and 
                    $ace.ActiveDirectoryRights -match "WriteProperty|WriteDacl|WriteOwner") {
                    
                    $Results += [PSCustomObject]@{
                        SchemaObject = $object.Name
                        IdentityReference = $ace.IdentityReference
                        Permissions = $ace.ActiveDirectoryRights
                        RiskLevel = "High"
                        Recommendation = "Review schema modifications"
                    }
                }
            }
        }
        
        Write-ADReportLog -Message "Schema permissions analysis completed. Found $($Results.Count) entries for filter '$filter'" -Type Info -Terminal

        if ($Results.Count -gt 0) {
            # Ergebnisse in DataGrid anzeigen
            $Global:DataGridResults.ItemsSource = $Results | Sort-Object RiskLevel
            
            # Status aktualisieren
            $Global:TextBlockSelectedRows.Text = "Gefundene Einträge: $($Results.Count)" 
            $Global:TextBlockLastUpdate.Text = "Letzte Aktualisierung: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')"
            $Global:StatusIndicator.Fill = "#FF00C800"
        } else {
            # Keine Ergebnisse gefunden
            $Global:DataGridResults.ItemsSource = @()
            $Global:TextBlockSelectedRows.Text = "Keine Einträge gefunden"
            $Global:TextBlockLastUpdate.Text = "Letzte Aktualisierung: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')"
            $Global:StatusIndicator.Fill = "#FFFFFF00" # Gelb für "keine Ergebnisse"
        }
        
        return $Results
        
    } catch {
        Write-ADReportLog -Message "Error analyzing schema permissions: $($_.Exception.Message)" -Type Error
        $Global:StatusIndicator.Fill = "#FFFF0000"
        $Global:DataGridResults.ItemsSource = @()
        $Global:TextBlockSelectedRows.Text = "Fehler bei der Analyse"
        return @()
    }
}

# ===================================
# NETWORK TOPOLOGY FUNCTIONS
# ===================================

Function Get-ADNetworkTopology {
    [CmdletBinding()]
    param(
        [ValidateSet("DomainControllers", "Sites", "OUHierarchy", "Trusts")]
        [string]$ViewType = "DomainControllers"
    )
    
    $topology = @{
        nodes = @()
        links = @()
    }
    
    try {
        switch ($ViewType) {
            "DomainControllers" {
                # Hole Domain und DCs
                $domain = Get-ADDomain
                $dcs = Get-ADDomainController -Filter *
                
                # Zentraler Domain-Knoten
                $topology.nodes += @{
                    id = "domain"
                    label = $domain.Name
                    type = "domain"
                    x = 400
                    y = 300
                    color = "#0078D7"
                    size = 40
                }
                
                # DC Knoten im Kreis anordnen
                $angleStep = 360 / $dcs.Count
                $radius = 200
                $centerX = 400
                $centerY = 300
                
                for ($i = 0; $i -lt $dcs.Count; $i++) {
                    $dc = $dcs[$i]
                    $angle = $i * $angleStep * [Math]::PI / 180
                    $x = $centerX + $radius * [Math]::Cos($angle)
                    $y = $centerY + $radius * [Math]::Sin($angle)
                    
                    $topology.nodes += @{
                        id = $dc.Name
                        label = $dc.Name
                        type = "dc"
                        x = $x
                        y = $y
                        color = if ($dc.IsGlobalCatalog) { "#10B981" } else { "#3B82F6" }
                        size = 30
                        properties = @{
                            Site = $dc.Site
                            IP = $dc.IPv4Address
                            OS = $dc.OperatingSystem
                            IsGC = $dc.IsGlobalCatalog
                            IsRODC = $dc.IsReadOnly
                        }
                    }
                    
                    # Link zum Domain-Knoten
                    $topology.links += @{
                        source = "domain"
                        target = $dc.Name
                        type = "contains"
                        color = "#9CA3AF"
                    }
                }
            }
            
            "Sites" {
                # AD Sites und Replikationslinks
                $sites = Get-ADReplicationSite -Filter *
                $siteLinks = Get-ADReplicationSiteLink -Filter *
                
                # Erstelle Site-Knoten
                $gridSize = [Math]::Ceiling([Math]::Sqrt($sites.Count))
                $spacing = 150
                
                for ($i = 0; $i -lt $sites.Count; $i++) {
                    $site = $sites[$i]
                    $row = [Math]::Floor($i / $gridSize)
                    $col = $i % $gridSize
                    
                    $topology.nodes += @{
                        id = $site.Name
                        label = $site.Name
                        type = "site"
                        x = 100 + ($col * $spacing)
                        y = 100 + ($row * $spacing)
                        color = "#10B981"
                        size = 35
                        properties = @{
                            Description = $site.Description
                            Location = $site.Location
                        }
                    }
                }
                
                # Erstelle Site-Links
                foreach ($link in $siteLinks) {
                    $sitesInLink = $link.SitesIncluded | ForEach-Object {
                        ($_ -split ',')[0] -replace 'CN=', ''
                    }
                    
                    for ($i = 0; $i -lt $sitesInLink.Count - 1; $i++) {
                        for ($j = $i + 1; $j -lt $sitesInLink.Count; $j++) {
                            $topology.links += @{
                                source = $sitesInLink[$i]
                                target = $sitesInLink[$j]
                                type = "replication"
                                color = "#3B82F6"
                                cost = $link.Cost
                            }
                        }
                    }
                }
            }
        }
        
        return $topology
    }
    catch {
        Write-Error "Fehler beim Erstellen der Netzwerktopologie: $_"
        return $topology
    }
}

# Funktion zum Zeichnen auf Canvas
Function Draw-TopologyOnCanvas {
    param(
        [System.Windows.Controls.Canvas]$Canvas,
        [hashtable]$Topology
    )
    
    $Canvas.Children.Clear()
    
    # Zeichne Links
    foreach ($link in $Topology.links) {
        $sourceNode = $Topology.nodes | Where-Object { $_.id -eq $link.source }
        $targetNode = $Topology.nodes | Where-Object { $_.id -eq $link.target }
        
        if ($sourceNode -and $targetNode) {
            $line = New-Object System.Windows.Shapes.Line
            $line.X1 = $sourceNode.x
            $line.Y1 = $sourceNode.y
            $line.X2 = $targetNode.x
            $line.Y2 = $targetNode.y
            $line.Stroke = $link.color
            $line.StrokeThickness = 2
            $line.Opacity = 0.6
            $Canvas.Children.Add($line)
        }
    }
    
    # Zeichne Knoten
    foreach ($node in $Topology.nodes) {
        # Knoten-Container
        $nodeGroup = New-Object System.Windows.Controls.Canvas
        
        # Kreis für Knoten
        $ellipse = New-Object System.Windows.Shapes.Ellipse
        $ellipse.Width = $node.size
        $ellipse.Height = $node.size
        $ellipse.Fill = $node.color
        $ellipse.Stroke = "#E5E7EB"
        $ellipse.StrokeThickness = 2
        
        # Schatten
        $shadow = New-Object System.Windows.Media.Effects.DropShadowEffect
        $shadow.BlurRadius = 8
        $shadow.ShadowDepth = 2
        $shadow.Opacity = 0.3
        $ellipse.Effect = $shadow
        
        # Text-Label
        $label = New-Object System.Windows.Controls.TextBlock
        $label.Text = $node.label
        $label.Foreground = "White"
        $label.FontWeight = "SemiBold"
        $label.FontSize = 10
        $label.TextAlignment = "Center"
        
        # Positionierung
        [System.Windows.Controls.Canvas]::SetLeft($ellipse, $node.x - $node.size/2)
        [System.Windows.Controls.Canvas]::SetTop($ellipse, $node.y - $node.size/2)
        [System.Windows.Controls.Canvas]::SetLeft($label, $node.x - 30)
        [System.Windows.Controls.Canvas]::SetTop($label, $node.y - 5)
        
        # Tooltip
        if ($node.properties) {
            $tooltip = ""
            foreach ($prop in $node.properties.GetEnumerator()) {
                $tooltip += "$($prop.Key): $($prop.Value)`n"
            }
            $ellipse.ToolTip = $tooltip.TrimEnd()
        }
        
        $Canvas.Children.Add($ellipse)
        $Canvas.Children.Add($label)
    }
}

# ===================================
# SECURITY HEATMAP FUNCTIONS
# ===================================

Function Get-SecurityHeatMapData {
    [CmdletBinding()]
    param(
        [ValidateSet("PasswordSecurity", "AccountActivity", "PrivilegeLevel", "Compliance")]
        [string]$MetricType = "PasswordSecurity",
        
        [ValidateSet("Department", "OU", "Manager")]
        [string]$GroupBy = "Department"
    )
    
    $heatMapData = @()
    
    try {
        # Hole Benutzer mit relevanten Attributen
        $users = Get-ADUser -Filter * -Properties *
        
        # Gruppiere Benutzer
        $groups = switch ($GroupBy) {
            "Department" { $users | Group-Object Department | Where-Object { $_.Name } }
            "OU" { $users | Group-Object { ($_.DistinguishedName -split ',')[1] } }
            "Manager" { $users | Group-Object Manager | Where-Object { $_.Name } }
        }
        
        foreach ($group in $groups) {
            $score = 0
            $details = @{}
            
            switch ($MetricType) {
                "PasswordSecurity" {
                    # Berechne Passwort-Sicherheitsscore
                    $pwdNeverExpires = ($group.Group | Where-Object { $_.PasswordNeverExpires }).Count
                    $pwdNotRequired = ($group.Group | Where-Object { $_.PasswordNotRequired }).Count
                    $oldPasswords = ($group.Group | Where-Object { 
                        $_.PasswordLastSet -and $_.PasswordLastSet -lt (Get-Date).AddDays(-180) 
                    }).Count
                    
                    # Score: 0-100 (100 = sehr sicher)
                    $totalUsers = $group.Count
                    $score = 100
                    $score -= ($pwdNeverExpires / $totalUsers) * 30
                    $score -= ($pwdNotRequired / $totalUsers) * 40
                    $score -= ($oldPasswords / $totalUsers) * 30
                    
                    $details = @{
                        "Total Users" = $totalUsers
                        "Password Never Expires" = $pwdNeverExpires
                        "Password Not Required" = $pwdNotRequired
                        "Old Passwords (>180d)" = $oldPasswords
                    }
                }
                
                "AccountActivity" {
                    # Berechne Account-Aktivitäts-Risiko
                    $inactiveUsers = ($group.Group | Where-Object { 
                        $_.LastLogonDate -and $_.LastLogonDate -lt (Get-Date).AddDays(-90) 
                    }).Count
                    $neverLoggedOn = ($group.Group | Where-Object { -not $_.LastLogonDate }).Count
                    $disabledAccounts = ($group.Group | Where-Object { -not $_.Enabled }).Count
                    
                    $totalUsers = $group.Count
                    $score = 100
                    $score -= ($inactiveUsers / $totalUsers) * 40
                    $score -= ($neverLoggedOn / $totalUsers) * 30
                    $score -= ($disabledAccounts / $totalUsers) * 30
                    
                    $details = @{
                        "Total Users" = $totalUsers
                        "Inactive (>90d)" = $inactiveUsers
                        "Never Logged On" = $neverLoggedOn
                        "Disabled Accounts" = $disabledAccounts
                    }
                }
            }
            
            $heatMapData += @{
                Group = $group.Name
                Score = [Math]::Round($score, 1)
                Count = $group.Count
                Details = $details
                Color = Get-HeatMapColor -Score $score
            }
        }
        
        return $heatMapData | Sort-Object Score
    }
    catch {
        Write-Error "Fehler beim Erstellen der HeatMap-Daten: $_"
        return @()
    }
}

Function Get-HeatMapColor {
    param([double]$Score)
    
    # Farbverlauf von Rot (0) über Gelb (50) zu Grün (100)
    if ($Score -ge 80) { return "#10B981" }  # Grün
    elseif ($Score -ge 60) { return "#34D399" }  # Hellgrün
    elseif ($Score -ge 40) { return "#FDE047" }  # Gelb
    elseif ($Score -ge 20) { return "#FB923C" }  # Orange
    else { return "#EF4444" }  # Rot
}

Function Draw-SecurityHeatMap {
    param(
        [System.Windows.Controls.Grid]$Container,
        [array]$HeatMapData
    )
    
    $Container.Children.Clear()
    $Container.RowDefinitions.Clear()
    $Container.ColumnDefinitions.Clear()
    
    # Berechne Grid-Layout
    $itemCount = $HeatMapData.Count
    $cols = [Math]::Ceiling([Math]::Sqrt($itemCount))
    $rows = [Math]::Ceiling($itemCount / $cols)
    
    # Erstelle Grid-Definitionen
    for ($i = 0; $i -lt $rows; $i++) {
        $rowDef = New-Object System.Windows.RowDefinition
        $rowDef.Height = "1*"
        $Container.RowDefinitions.Add($rowDef)
    }
    
    for ($i = 0; $i -lt $cols; $i++) {
        $colDef = New-Object System.Windows.ColumnDefinition
        $colDef.Width = "1*"
        $Container.ColumnDefinitions.Add($colDef)
    }
    
    # Füge HeatMap-Zellen hinzu
    for ($i = 0; $i -lt $HeatMapData.Count; $i++) {
        $data = $HeatMapData[$i]
        $row = [Math]::Floor($i / $cols)
        $col = $i % $cols
        
        # Zellen-Container
        $border = New-Object System.Windows.Controls.Border
        $border.Margin = "5"
        $border.CornerRadius = "8"
        $border.Background = $data.Color
        $border.MinHeight = 100
        
        # Inhalt
        $stack = New-Object System.Windows.Controls.StackPanel
        $stack.VerticalAlignment = "Center"
        $stack.Margin = "10"
        
        # Gruppenname
        $groupLabel = New-Object System.Windows.Controls.TextBlock
        $groupLabel.Text = $data.Group
        $groupLabel.FontWeight = "Bold"
        $groupLabel.FontSize = 14
        $groupLabel.Foreground = "White"
        $groupLabel.TextWrapping = "Wrap"
        $groupLabel.HorizontalAlignment = "Center"
        $stack.Children.Add($groupLabel)
        
        # Score
        $scoreLabel = New-Object System.Windows.Controls.TextBlock
        $scoreLabel.Text = "$($data.Score)%"
        $scoreLabel.FontSize = 24
        $scoreLabel.FontWeight = "Bold"
        $scoreLabel.Foreground = "White"
        $scoreLabel.HorizontalAlignment = "Center"
        $scoreLabel.Margin = "0,5,0,0"
        $stack.Children.Add($scoreLabel)
        
        # Benutzeranzahl
        $countLabel = New-Object System.Windows.Controls.TextBlock
        $countLabel.Text = "$($data.Count) users"
        $countLabel.FontSize = 11
        $countLabel.Foreground = "White"
        $countLabel.Opacity = 0.8
        $countLabel.HorizontalAlignment = "Center"
        $stack.Children.Add($countLabel)
        
        $border.Child = $stack
        
        # Tooltip mit Details
        $tooltip = New-Object System.Windows.Controls.ToolTip
        $tooltipStack = New-Object System.Windows.Controls.StackPanel
        
        foreach ($detail in $data.Details.GetEnumerator()) {
            $detailText = New-Object System.Windows.Controls.TextBlock
            $detailText.Text = "$($detail.Key): $($detail.Value)"
            $detailText.Margin = "0,2"
            $tooltipStack.Children.Add($detailText)
        }
        
        $tooltip.Content = $tooltipStack
        $border.ToolTip = $tooltip
        
        # Position im Grid
        [System.Windows.Controls.Grid]::SetRow($border, $row)
        [System.Windows.Controls.Grid]::SetColumn($border, $col)
        
        $Container.Children.Add($border)
    }
}

# ===================================
# CUSTOM REPORT BUILDER FUNCTIONS
# ===================================

# Report Template Klasse
class ReportTemplate {
    [string]$Name
    [string]$Description
    [array]$Fields
    [hashtable]$Filters
    [string]$Layout
    
    ReportTemplate([string]$name) {
        $this.Name = $name
        $this.Fields = @()
        $this.Filters = @{}
        $this.Layout = "Table"
    }
}

Function Initialize-ReportBuilder {
    param(
        [System.Windows.Controls.ListBox]$AvailableFieldsList,
        [System.Windows.Controls.StackPanel]$ReportCanvas
    )
    
    # Verfügbare Felder laden
    $availableFields = @(
        @{Name="DisplayName"; Category="User"; Type="String"}
        @{Name="SamAccountName"; Category="User"; Type="String"}
        @{Name="Department"; Category="User"; Type="String"}
        @{Name="Title"; Category="User"; Type="String"}
        @{Name="Manager"; Category="User"; Type="String"}
        @{Name="LastLogonDate"; Category="User"; Type="DateTime"}
        @{Name="PasswordLastSet"; Category="User"; Type="DateTime"}
        @{Name="Enabled"; Category="User"; Type="Boolean"}
        @{Name="Name"; Category="Computer"; Type="String"}
        @{Name="OperatingSystem"; Category="Computer"; Type="String"}
        @{Name="LastLogonDate"; Category="Computer"; Type="DateTime"}
        @{Name="GroupName"; Category="Group"; Type="String"}
        @{Name="GroupCategory"; Category="Group"; Type="String"}
        @{Name="GroupScope"; Category="Group"; Type="String"}
        @{Name="Members"; Category="Group"; Type="Array"}
    )
    
    foreach ($field in $availableFields) {
        $item = New-Object System.Windows.Controls.ListBoxItem
        $item.Content = "$($field.Category).$($field.Name)"
        $item.Tag = $field
        
        # Drag & Drop aktivieren
        $item.AllowDrop = $true
        $item.Add_MouseMove({
            param($sender, $e)
            if ($e.LeftButton -eq 'Pressed') {
                [System.Windows.DragDrop]::DoDragDrop($sender, $sender.Tag, 'Copy')
            }
        })
        
        $AvailableFieldsList.Items.Add($item)
    }
    
    # Drop-Handler für Canvas
    $ReportCanvas.AllowDrop = $true
    $ReportCanvas.Add_Drop({
        param($sender, $e)
        $field = $e.Data.GetData([hashtable])
        if ($field) {
            Add-FieldToReport -Field $field -Canvas $sender
        }
    })
    
    $ReportCanvas.Add_DragOver({
        param($sender, $e)
        $e.Effects = 'Copy'
        $e.Handled = $true
    })
}

Function Add-FieldToReport {
    param(
        [hashtable]$Field,
        [System.Windows.Controls.StackPanel]$Canvas
    )
    
    # Erstelle Feld-Container
    $fieldContainer = New-Object System.Windows.Controls.Border
    $fieldContainer.Background = "White"
    $fieldContainer.BorderBrush = "#E5E7EB"
    $fieldContainer.BorderThickness = "1"
    $fieldContainer.CornerRadius = "6"
    $fieldContainer.Margin = "0,5"
    $fieldContainer.Padding = "10"
    
    $grid = New-Object System.Windows.Controls.Grid
    $col1 = New-Object System.Windows.ColumnDefinition
    $col1.Width = "*"
    $col2 = New-Object System.Windows.ColumnDefinition
    $col2.Width = "Auto"
    $grid.ColumnDefinitions.Add($col1)
    $grid.ColumnDefinitions.Add($col2)
    
    # Feldname
    $nameBlock = New-Object System.Windows.Controls.TextBlock
    $nameBlock.Text = "$($Field.Category).$($Field.Name)"
    $nameBlock.FontWeight = "Medium"
    $nameBlock.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetColumn($nameBlock, 0)
    $grid.Children.Add($nameBlock)
    
    # Entfernen-Button
    $removeBtn = New-Object System.Windows.Controls.Button
    $removeBtn.Content = "✕"
    $removeBtn.Width = 20
    $removeBtn.Height = 20
    $removeBtn.Background = "Transparent"
    $removeBtn.BorderThickness = 0
    $removeBtn.Tag = $fieldContainer
    $removeBtn.Add_Click({
        param($s, $e)
        $Canvas.Children.Remove($s.Tag)
    })
    [System.Windows.Controls.Grid]::SetColumn($removeBtn, 1)
    $grid.Children.Add($removeBtn)
    
    $fieldContainer.Child = $grid
    $Canvas.Children.Remove($Canvas.Children[0]) # Entferne Platzhalter-Text
    $Canvas.Children.Add($fieldContainer)
}

Function Build-CustomReport {
    param(
        [System.Windows.Controls.StackPanel]$ReportCanvas
    )
    
    $selectedFields = @()
    foreach ($child in $ReportCanvas.Children) {
        if ($child -is [System.Windows.Controls.Border]) {
            $textBlock = $child.Child.Children[0]
            $selectedFields += $textBlock.Text
        }
    }
    
    if ($selectedFields.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Bitte wählen Sie mindestens ein Feld aus.", "Report Builder", "OK", "Warning")
        return
    }
    
    # Generiere LDAP-Properties aus Feldnamen
    $properties = $selectedFields | ForEach-Object {
        ($_ -split '\.')[-1]
    } | Select-Object -Unique
    
    # Bestimme Objekttyp
    $categories = $selectedFields | ForEach-Object {
        ($_ -split '\.')[0]
    } | Select-Object -Unique
    
    $results = @()
    
    try {
        if ($categories -contains "User") {
            $users = Get-ADUser -Filter * -Properties $properties
            foreach ($user in $users) {
                $obj = [PSCustomObject]@{}
                foreach ($field in $selectedFields) {
                    if ($field -like "User.*") {
                        $propName = ($field -split '\.')[-1]
                        $obj | Add-Member -NotePropertyName $field -NotePropertyValue $user.$propName
                    }
                }
                $results += $obj
            }
        }
        
        # Ähnliche Logik für Computer und Groups...
        
        return $results
    }
    catch {
        Write-Error "Fehler beim Erstellen des benutzerdefinierten Berichts: $_"
        return @()
    }
}

Function Save-ReportTemplate {
    param(
        [System.Windows.Controls.StackPanel]$ReportCanvas,
        [string]$TemplateName,
        [string]$Description
    )
    
    $template = [ReportTemplate]::new($TemplateName)
    $template.Description = $Description
    
    foreach ($child in $ReportCanvas.Children) {
        if ($child -is [System.Windows.Controls.Border]) {
            $textBlock = $child.Child.Children[0]
            $template.Fields += $textBlock.Text
        }
    }
    
    # Speichere Template als JSON
    $templatesPath = "$env:APPDATA\easyADReport\Templates"
    if (-not (Test-Path $templatesPath)) {
        New-Item -ItemType Directory -Path $templatesPath -Force
    }
    
    $templateFile = Join-Path $templatesPath "$($TemplateName).json"
    $template | ConvertTo-Json | Out-File $templateFile -Encoding UTF8
    
    return $templateFile
}

# Funktion zum Aktualisieren der ErgebniszÃ¤hler im Header
Function Initialize-ResultCounters {
    # GesamtergebniszÃ¤hler auf 0 setzen
    if ($null -ne $Global:TotalResultCountText) {
        $Global:TotalResultCountText.Text = "0"
    }
    
    # Sicherstellen, dass alle ZÃ¤hler zurÃ¼ckgesetzt werden
    if ($null -ne $Global:UserCountText) {
        $Global:UserCountText.Text = "0"
    }
    
    if ($null -ne $Global:ComputerCountText) {
        $Global:ComputerCountText.Text = "0"
    }
    
    if ($null -ne $Global:GroupCountText) {
        $Global:GroupCountText.Text = "0"
    }
    
    # Status zurÃ¼cksetzen
    if ($null -ne $Global:TextBlockStatus) {
        $Global:TextBlockStatus.Text = "Ready for query..."
    }
    
    # DataGrid leeren
    if ($null -ne $Global:DataGridResults) {
        $Global:DataGridResults.ItemsSource = $null
    }
}

# Funktion zum Aktualisieren der ErgebniszÃ¤hler im Header
Function Update-ResultCounters {
    param (
        [Parameter(Mandatory=$true)]
        [System.Collections.IList]$Results
    )

    if ($null -eq $Results) {
        if ($null -ne $Global:TotalResultCountText) { $Global:TotalResultCountText.Text = "0" }
        if ($null -ne $Global:UserCountText) { $Global:UserCountText.Text = "0" }
        if ($null -ne $Global:ComputerCountText) { $Global:ComputerCountText.Text = "0" }
        if ($null -ne $Global:GroupCountText) { $Global:GroupCountText.Text = "0" }
        return
    }

    $totalCount = $Results.Count
    $userCount = 0
    $computerCount = 0
    $groupCount = 0

    if ($totalCount -gt 0) {
        foreach ($item in $Results) {
            if ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [array] -and $item.ObjectClass -contains "user") {
                $userCount++
            } elseif ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [string] -and $item.ObjectClass -eq "user") {
                $userCount++
            } elseif ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [array] -and $item.ObjectClass -contains "computer") {
                $computerCount++
            } elseif ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [string] -and $item.ObjectClass -eq "computer") {
                $computerCount++
            } elseif ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [array] -and $item.ObjectClass -contains "group") {
                $groupCount++
            } elseif ($item.PSObject.Properties["ObjectClass"] -and $item.ObjectClass -is [string] -and $item.ObjectClass -eq "group") {
                $groupCount++
            } elseif ($item.PSObject.Properties["ObjectCategory"] -and $item.ObjectCategory -like "*CN=Person,CN=Schema,CN=Configuration,*") {
                $userCount++
            } elseif ($item.PSObject.Properties["ObjectCategory"] -and $item.ObjectCategory -like "*CN=Computer,CN=Schema,CN=Configuration,*") {
                $computerCount++
            } elseif ($item.PSObject.Properties["ObjectCategory"] -and $item.ObjectCategory -like "*CN=Group,CN=Schema,CN=Configuration,*") {
                $groupCount++
            }
        }
    }

    if ($null -ne $Global:TotalResultCountText) { $Global:TotalResultCountText.Text = $totalCount.ToString() }
    if ($null -ne $Global:UserCountText) { $Global:UserCountText.Text = $userCount.ToString() }
    if ($null -ne $Global:ComputerCountText) { $Global:ComputerCountText.Text = $computerCount.ToString() }
    if ($null -ne $Global:GroupCountText) { $Global:GroupCountText.Text = $groupCount.ToString() }
}

# Funktion zur universellen Aktualisierung der Ergebnisanzeige und DataGrid
Function Update-ADReportResults {
    param (
        [Parameter(Mandatory=$false)]
        $Results = $null,
        
        [Parameter(Mandatory=$false)]
        [string]$StatusMessage
    )

    # Stelle sicher, dass Results immer ein Array ist, auch bei NULL
    if ($null -eq $Results) {
        $Results = @()
    } elseif ($Results -isnot [array]) {
        $Results = @($Results)
    }

    if ($null -ne $Global:DataGridResults) {
        $Global:DataGridResults.ItemsSource = $Results
    }

    if (-not [string]::IsNullOrWhiteSpace($StatusMessage) -and $null -ne $Global:TextBlockStatus) {
        $Global:TextBlockStatus.Text = $StatusMessage
    }
    
    Update-ResultCounters -Results $Results
    Update-ResultVisualization -Results $Results
}

# Hilfsfunktion zum sicheren HinzufÃ¼gen von Event-Handlern
Function Add-SafeEventHandler {
    param(
        [Parameter(Mandatory=$true)]
        $Button,
        [Parameter(Mandatory=$true)]
        [scriptblock]$Handler
    )
    
    if ($null -ne $Button) {
        try {
            $Button.add_Click($Handler)
        } catch {
            Write-ADReportLog -Message "Warnung: Konnte Event-Handler fÃ¼r Button nicht hinzufÃ¼gen: $($_.Exception.Message)" -Type Warning -Terminal
        }
   }
}

# --- Globale Variablen fÃ¼r UI Elemente --- 
Function Initialize-ADReportForm {
    param($XAMLContent)
    # ÃœberprÃ¼fen, ob das Window-Objekt bereits existiert und zurÃ¼cksetzen
    if ($Global:Window) {
        Remove-Variable -Name Window -Scope Global -ErrorAction SilentlyContinue
    }
    
    $reader = New-Object System.Xml.XmlNodeReader $XAMLContent
    $Global:Window = [Windows.Markup.XamlReader]::Load( $reader )

    # --- UI Elemente zu globalen Variablen zuweisen --- 
    # Objekttyp Radio Buttons
    $Global:RadioButtonUser = $Window.FindName("RadioButtonUser")
    $Global:RadioButtonGroup = $Window.FindName("RadioButtonGroup")
    $Global:RadioButtonComputer = $Window.FindName("RadioButtonComputer")
    $Global:RadioButtonGroupMemberships = $Window.FindName("RadioButtonGroupMemberships")

    # Filter und Attribute
    $Global:ComboBoxFilterAttribute1 = $Window.FindName("ComboBoxFilterAttribute1")
    $Global:ComboBoxFilterOperator1 = $Window.FindName("ComboBoxFilterOperator1")
    $Global:TextBoxFilterValue1 = $Window.FindName("TextBoxFilterValue1")
    $Global:ComboBoxFilterAttribute2 = $Window.FindName("ComboBoxFilterAttribute2")  
    $Global:ComboBoxFilterOperator2 = $Window.FindName("ComboBoxFilterOperator2")
    $Global:TextBoxFilterValue2 = $Window.FindName("TextBoxFilterValue2")
    $Global:RadioButtonAnd = $Window.FindName("RadioButtonAnd")
    $Global:RadioButtonOr = $Window.FindName("RadioButtonOr")
    $Global:CheckBoxUseSecondFilter = $Window.FindName("CheckBoxUseSecondFilter")
    $Global:SecondFilterPanel = $Window.FindName("SecondFilterPanel")

    # Bestehende UI-Elemente
    $Global:RadioButtonUser = $Window.FindName("RadioButtonUser")
    $Global:RadioButtonGroup = $Window.FindName("RadioButtonGroup")
    $Global:RadioButtonComputer = $Window.FindName("RadioButtonComputer")
    $Global:RadioButtonGroupMemberships = $Window.FindName("RadioButtonGroupMemberships")
    $Global:ListBoxSelectAttributes = $Window.FindName("ListBoxSelectAttributes")
    $Global:ButtonQueryAD = $Window.FindName("ButtonQueryAD")
    $Global:DataGridResults = $Window.FindName("DataGridResults")
    $Global:TextBlockStatus = $Window.FindName("TextBlockStatus")
    $Global:TotalResultCountText = $Window.FindName("TotalResultCountText")
    $Global:ButtonExportCSV = $Window.FindName("ButtonExportCSV")
    $Global:ButtonExportHTML = $Window.FindName("ButtonExportHTML")
    $Global:ButtonRefresh = $Window.FindName("ButtonRefresh")
    $Global:ButtonCopy = $Window.FindName("ButtonCopy")

    $Global:ButtonQuickAllUsers = $Window.FindName("ButtonQuickAllUsers")
    $Global:ButtonQuickDisabledUsers = $Window.FindName("ButtonQuickDisabledUsers")
    $Global:ButtonQuickLockedUsers = $Window.FindName("ButtonQuickLockedUsers")
    $Global:ButtonQuickNeverExpire = $Window.FindName("ButtonQuickNeverExpire")
    $Global:ButtonQuickInactiveUsers = $Window.FindName("ButtonQuickInactiveUsers")
    $Global:ButtonQuickAdminUsers = $Window.FindName("ButtonQuickAdminUsers")
    $Global:ButtonQuickRecentlyCreatedUsers = $Window.FindName("ButtonQuickRecentlyCreatedUsers")
    $Global:ButtonQuickPasswordExpiringSoon = $Window.FindName("ButtonQuickPasswordExpiringSoon")
    $Global:ButtonQuickExpiredPasswords = $Window.FindName("ButtonQuickExpiredPasswords")
    $Global:ButtonQuickNeverLoggedOn = $Window.FindName("ButtonQuickNeverLoggedOn")
    $Global:ButtonQuickRecentlyDeletedUsers = $Window.FindName("ButtonQuickRecentlyDeletedUsers")
    $Global:ButtonQuickRecentlyModifiedUsers = $Window.FindName("ButtonQuickRecentlyModifiedUsers")
    $Global:ButtonQuickUsersWithoutManager = $Window.FindName("ButtonQuickUsersWithoutManager")
    $Global:ButtonQuickUsersMissingRequiredAttributes = $Window.FindName("ButtonQuickUsersMissingRequiredAttributes")
    $Global:ButtonQuickUsersDuplicateLogonNames = $Window.FindName("ButtonQuickUsersDuplicateLogonNames")
    $Global:ButtonQuickOrphanedSIDsUsers = $Window.FindName("ButtonQuickOrphanedSIDsUsers")
    
    $Global:ButtonQuickGroups = $Window.FindName("ButtonQuickGroups")
    $Global:ButtonQuickSecurityGroups = $Window.FindName("ButtonQuickSecurityGroups")
    $Global:ButtonQuickDistributionGroups = $Window.FindName("ButtonQuickDistributionGroups")
    $Global:ButtonQuickComputers = $Window.FindName("ButtonQuickComputers")
    $Global:ButtonQuickInactiveComputers = $Window.FindName("ButtonQuickInactiveComputers")

    $Global:ButtonQuickWeakPasswordPolicy = $Window.FindName("ButtonQuickWeakPasswordPolicy")
    $Global:ButtonQuickRiskyGroupMemberships = $Window.FindName("ButtonQuickRiskyGroupMemberships")
    $Global:ButtonQuickPrivilegedAccounts = $Window.FindName("ButtonQuickPrivilegedAccounts")

    $Global:ButtonQuickFSMORoles = $Window.FindName("ButtonQuickFSMORoles")
    $Global:ButtonQuickDCStatus = $Window.FindName("ButtonQuickDCStatus")
    $Global:ButtonQuickReplicationStatus = $Window.FindName("ButtonQuickReplicationStatus")
    $Global:ButtonQuickOUHierarchy = $Window.FindName("ButtonQuickOUHierarchy")
    $Global:ButtonQuickSitesSubnets = $Window.FindName("ButtonQuickSitesSubnets")
    
    $Global:ButtonQuickDepartmentStats = $Window.FindName("ButtonQuickDepartmentStats")
    $Global:ButtonQuickDepartmentSecurity = $Window.FindName("ButtonQuickDepartmentSecurity")
    $Global:ButtonQuickKerberoastable = $Window.FindName("ButtonQuickKerberoastable")
    $Global:ButtonQuickASREPRoastable = $Window.FindName("ButtonQuickASREPRoastable")
    $Global:ButtonQuickDelegation = $Window.FindName("ButtonQuickDelegation")
    $Global:ButtonQuickDCSyncRights = $Window.FindName("ButtonQuickDCSyncRights")
    $Global:ButtonQuickSchemaAdmins = $Window.FindName("ButtonQuickSchemaAdmins")
    $Global:ButtonQuickCertificateAnalysis = $Window.FindName("ButtonQuickCertificateAnalysis")
    $Global:ButtonQuickSYSVOLHealth = $Window.FindName("ButtonQuickSYSVOLHealth")
    $Global:ButtonQuickDNSHealth = $Window.FindName("ButtonQuickDNSHealth")
    $Global:ButtonQuickBackupStatus = $Window.FindName("ButtonQuickBackupStatus")
    $Global:ButtonQuickSchemaAnalysis = $Window.FindName("ButtonQuickSchemaAnalysis")
    $Global:ButtonQuickTrustRelationships = $Window.FindName("ButtonQuickTrustRelationships")
    $Global:ButtonQuickQuotasLimits = $Window.FindName("ButtonQuickQuotasLimits")

    $Global:ButtonQuickStalePasswords = $Window.FindName("ButtonQuickStalePasswords")
    $Global:ButtonQuickNeverChangingPasswords = $Window.FindName("ButtonQuickNeverChangingPasswords")
    $Global:ButtonQuickExpiringAccounts = $Window.FindName("ButtonQuickExpiringAccounts")
    $Global:ButtonQuickReversibleEncryption = $Window.FindName("ButtonQuickReversibleEncryption")
    $Global:ButtonQuickKerberosDES = $Window.FindName("ButtonQuickKerberosDES")
    $Global:ButtonQuickUsersWithSPN = $Window.FindName("ButtonQuickUsersWithSPN")
    $Global:ButtonQuickGuestAccountStatus = $Window.FindName("ButtonQuickGuestAccountStatus")
    $Global:ButtonQuickUsersByDepartment = $Window.FindName("ButtonQuickUsersByDepartment")
    $Global:ButtonQuickUsersByManager = $Window.FindName("ButtonQuickUsersByManager")
    $Global:ButtonQuickMobileDeviceUsers = $Window.FindName("ButtonQuickMobileDeviceUsers")
    $Global:ButtonQuickEmptyGroups = $Window.FindName("ButtonQuickEmptyGroups")
    $Global:ButtonQuickNestedGroups = $Window.FindName("ButtonQuickNestedGroups")
    $Global:ButtonQuickCircularGroups = $Window.FindName("ButtonQuickCircularGroups")
    $Global:ButtonQuickGroupsByTypeScope = $Window.FindName("ButtonQuickGroupsByTypeScope")
    $Global:ButtonQuickDynamicDistGroups = $Window.FindName("ButtonQuickDynamicDistGroups")
    $Global:ButtonQuickMailEnabledGroups = $Window.FindName("ButtonQuickMailEnabledGroups")
    $Global:ButtonQuickGroupsWithoutOwners = $Window.FindName("ButtonQuickGroupsWithoutOwners")
    $Global:ButtonQuickLargeGroups = $Window.FindName("ButtonQuickLargeGroups")
    $Global:ButtonQuickRecentlyModifiedGroups = $Window.FindName("ButtonQuickRecentlyModifiedGroups")
    

    $Global:ButtonQuickOSSummary = $Window.FindName("ButtonQuickOSSummary")
    $Global:ButtonQuickComputersByOSVersion = $Window.FindName("ButtonQuickComputersByOSVersion")
    $Global:ButtonQuickBitLockerStatus = $Window.FindName("ButtonQuickBitLockerStatus")
    $Global:ButtonQuickStaleComputerPasswords = $Window.FindName("ButtonQuickStaleComputerPasswords")
    $Global:ButtonQuickComputersNeverLoggedOn = $Window.FindName("ButtonQuickComputersNeverLoggedOn")
    $Global:ButtonQuickDuplicateComputerNames = $Window.FindName("ButtonQuickDuplicateComputerNames")
    $Global:ButtonQuickComputersByLocation = $Window.FindName("ButtonQuickComputersByLocation")
    
    # Neue Security Audit Buttons aus Roadmap
    $Global:ButtonQuickHoneyTokens = $Window.FindName("ButtonQuickHoneyTokens")
    $Global:ButtonQuickPrivilegeEscalation = $Window.FindName("ButtonQuickPrivilegeEscalation")
    $Global:ButtonQuickExposedCredentials = $Window.FindName("ButtonQuickExposedCredentials")
    $Global:ButtonQuickSuspiciousLogons = $Window.FindName("ButtonQuickSuspiciousLogons")
    $Global:ButtonQuickForeignSecurityPrincipals = $Window.FindName("ButtonQuickForeignSecurityPrincipals")
    $Global:ButtonQuickSIDHistoryAbuse = $Window.FindName("ButtonQuickSIDHistoryAbuse")
    
    # Service Account Buttons
    $Global:ButtonQuickServiceAccountsOverview = $Window.FindName("ButtonQuickServiceAccountsOverview")
    $Global:ButtonQuickManagedServiceAccounts = $Window.FindName("ButtonQuickManagedServiceAccounts")
    $Global:ButtonQuickServiceAccountsSPN = $Window.FindName("ButtonQuickServiceAccountsSPN")
    $Global:ButtonQuickHighPrivServiceAccounts = $Window.FindName("ButtonQuickHighPrivServiceAccounts")
    $Global:ButtonQuickServiceAccountPasswordAge = $Window.FindName("ButtonQuickServiceAccountPasswordAge")
    $Global:ButtonQuickUnusedServiceAccounts = $Window.FindName("ButtonQuickUnusedServiceAccounts")
    
    # GPO & Policy Buttons
    $Global:ButtonQuickGPOOverview = $Window.FindName("ButtonQuickGPOOverview")
    $Global:ButtonQuickUnlinkedGPOs = $Window.FindName("ButtonQuickUnlinkedGPOs")
    $Global:ButtonQuickEmptyGPOs = $Window.FindName("ButtonQuickEmptyGPOs")
    $Global:ButtonQuickGPOPermissions = $Window.FindName("ButtonQuickGPOPermissions")
    $Global:ButtonQuickPasswordPolicySummary = $Window.FindName("ButtonQuickPasswordPolicySummary")
    $Global:ButtonQuickAccountLockoutPolicies = $Window.FindName("ButtonQuickAccountLockoutPolicies")
    $Global:ButtonQuickFineGrainedPasswordPolicies = $Window.FindName("ButtonQuickFineGrainedPasswordPolicies")
        
    # Advanced Permissions Buttons
    $Global:ButtonQuickACLAnalysis = $Window.FindName("ButtonQuickACLAnalysis")
    $Global:ButtonQuickInheritanceBreaks = $Window.FindName("ButtonQuickInheritanceBreaks")
    $Global:ButtonQuickAdminSDHolderObjects = $Window.FindName("ButtonQuickAdminSDHolderObjects")
    $Global:ButtonQuickAdvancedDelegation = $Window.FindName("ButtonQuickAdvancedDelegation")
    $Global:ButtonQuickSchemaPermissions = $Window.FindName("ButtonQuickSchemaPermissions")
    
    # Attribute selection buttons
    $Global:ButtonSelectAllAttributes = $Window.FindName("ButtonSelectAllAttributes")
    $Global:ButtonSelectNoneAttributes = $Window.FindName("ButtonSelectNoneAttributes")
    $Global:TabControlAttributes = $Window.FindName("TabControlAttributes")
    $Global:ListBoxBasicAttributes = $Window.FindName("ListBoxBasicAttributes")
    $Global:ListBoxSecurityAttributes = $Window.FindName("ListBoxSecurityAttributes")
    $Global:ListBoxExtendedAttributes = $Window.FindName("ListBoxExtendedAttributes")
    
    # Help and About buttons
    $Global:ButtonHelp = $Window.FindName("ButtonHelp")
    $Global:ButtonAbout = $Window.FindName("ButtonAbout")
    
    # Footer elements
    $Global:StatusIndicator = $Window.FindName("StatusIndicator")
    $Global:TextBlockSelectedRows = $Window.FindName("TextBlockSelectedRows")
    $Global:TextBlockLastUpdate = $Window.FindName("TextBlockLastUpdate")

    # Event Handler fÃ¼r erweiterte Filter
    $Global:CheckBoxUseSecondFilter.add_Checked({
        $Global:SecondFilterPanel.IsEnabled = $true
    })
    
    $Global:CheckBoxUseSecondFilter.add_Unchecked({
        $Global:SecondFilterPanel.IsEnabled = $false
    })

    # Helper function to get all selected attributes from all ListBoxes
    function Get-AllSelectedAttributes {
        $selectedAttributes = @()
        
        if ($Global:ListBoxBasicAttributes) {
            foreach ($item in $Global:ListBoxBasicAttributes.SelectedItems) {
                $selectedAttributes += $item.Content
            }
        }
        if ($Global:ListBoxSecurityAttributes) {
            foreach ($item in $Global:ListBoxSecurityAttributes.SelectedItems) {
                $selectedAttributes += $item.Content
            }
        }
        if ($Global:ListBoxExtendedAttributes) {
            foreach ($item in $Global:ListBoxExtendedAttributes.SelectedItems) {
                $selectedAttributes += $item.Content
            }
        }
        
        return $selectedAttributes | Select-Object -Unique
    }

    # Helper function to select specific attributes in the tabbed ListBoxes
    function Select-AttributesInListBoxes {
        param (
            [string[]]$Attributes
        )
        
        # Clear all selections
        if ($Global:ListBoxBasicAttributes) {
            foreach ($item in $Global:ListBoxBasicAttributes.Items) {
                $item.IsSelected = $false
            }
        }
        if ($Global:ListBoxSecurityAttributes) {
            foreach ($item in $Global:ListBoxSecurityAttributes.Items) {
                $item.IsSelected = $false
            }
        }
        if ($Global:ListBoxExtendedAttributes) {
            foreach ($item in $Global:ListBoxExtendedAttributes.Items) {
                $item.IsSelected = $false
            }
        }
        
        # Select requested attributes
        foreach ($attr in $Attributes) {
            # Check in Basic Attributes
            if ($Global:ListBoxBasicAttributes) {
                foreach ($item in $Global:ListBoxBasicAttributes.Items) {
                    if ($item.Content -eq $attr) {
                        $item.IsSelected = $true
                        break
                    }
                }
            }
            # Check in Security Attributes
            if ($Global:ListBoxSecurityAttributes) {
                foreach ($item in $Global:ListBoxSecurityAttributes.Items) {
                    if ($item.Content -eq $attr) {
                        $item.IsSelected = $true
                        break
                    }
                }
            }
            # Check in Extended Attributes
            if ($Global:ListBoxExtendedAttributes) {
                foreach ($item in $Global:ListBoxExtendedAttributes.Items) {
                    if ($item.Content -eq $attr) {
                        $item.IsSelected = $true
                        break
                    }
                }
            }
        }
    }

    # Helper function to populate attributes in the tabbed ListBoxes
    function Set-AttributesForObjectType {
        param (
            [string]$ObjectType
        )
        
        # Clear all ListBoxes
        if ($Global:ListBoxBasicAttributes) { $Global:ListBoxBasicAttributes.Items.Clear() }
        if ($Global:ListBoxSecurityAttributes) { $Global:ListBoxSecurityAttributes.Items.Clear() }
        if ($Global:ListBoxExtendedAttributes) { $Global:ListBoxExtendedAttributes.Items.Clear() }
        
        switch ($ObjectType) {
            "User" {
                # Basic Attributes
                @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", "Department", "Title", "Enabled") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    if ($_ -in @("DisplayName", "SamAccountName", "Enabled")) { $item.IsSelected = $true }
                    $Global:ListBoxBasicAttributes.Items.Add($item)
                }
                
                # Security Attributes
                @("LastLogonTimestamp", "PasswordExpired", "PasswordLastSet", "AccountExpirationDate", "badPwdCount", "lockoutTime", "UserAccountControl", "memberOf") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    $Global:ListBoxSecurityAttributes.Items.Add($item)
                }
                
                # Extended Attributes
                @("whenCreated", "whenChanged", "Manager", "Company", "physicalDeliveryOfficeName", "telephoneNumber", "homeDirectory", "scriptPath") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    $Global:ListBoxExtendedAttributes.Items.Add($item)
                }
            }
            "Group" {
                # Basic Attributes
                @("Name", "SamAccountName", "Description", "GroupCategory", "GroupScope", "mail") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    if ($_ -in @("Name", "SamAccountName")) { $item.IsSelected = $true }
                    $Global:ListBoxBasicAttributes.Items.Add($item)
                }
                
                # Security Attributes
                @("ManagedBy", "info", "memberOf", "members") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    $Global:ListBoxSecurityAttributes.Items.Add($item)
                }
                
                # Extended Attributes
                @("whenCreated", "whenChanged", "distinguishedName", "objectGUID") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    $Global:ListBoxExtendedAttributes.Items.Add($item)
                }
            }
            "Computer" {
                # Basic Attributes
                @("Name", "DNSHostName", "OperatingSystem", "OperatingSystemVersion", "Enabled", "IPv4Address") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    if ($_ -in @("Name", "OperatingSystem", "Enabled")) { $item.IsSelected = $true }
                    $Global:ListBoxBasicAttributes.Items.Add($item)
                }
                
                # Security Attributes
                @("LastLogonDate", "PasswordLastSet", "userAccountControl") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    $Global:ListBoxSecurityAttributes.Items.Add($item)
                }
                
                # Extended Attributes
                @("whenCreated", "Description", "Location", "ManagedBy", "servicePrincipalName") | ForEach-Object {
                    $item = New-Object System.Windows.Controls.ListBoxItem
                    $item.Content = $_
                    $Global:ListBoxExtendedAttributes.Items.Add($item)
                }
            }
        }
    }

    # RadioButton Event Handler
    $RadioButtonUser.add_Checked({
        Write-ADReportLog -Message "Object type changed to User" -Type Info -Terminal
        
        # Filter-Attribute für Benutzer
        $UserFilterAttributes = @("SamAccountName", "DisplayName", "GivenName", "Surname", "mail", "Department", "Title", "EmployeeID", "UserPrincipalName")
        $Global:ComboBoxFilterAttribute1.Items.Clear()
        $Global:ComboBoxFilterAttribute2.Items.Clear()
        $UserFilterAttributes | ForEach-Object { 
            $Global:ComboBoxFilterAttribute1.Items.Add($_)
            $Global:ComboBoxFilterAttribute2.Items.Add($_)
        }
        
        # Vorauswahl auf SamAccountName setzen
        $Global:ComboBoxFilterAttribute1.SelectedItem = "SamAccountName"
        $Global:ComboBoxFilterAttribute2.SelectedItem = "DisplayName"
        
        # Populate attributes in tabbed ListBoxes
        Set-AttributesForObjectType -ObjectType "User"
        
        $Global:TextBlockStatus.Text = "Ready for user query"
    })

    $RadioButtonGroup.add_Checked({
        Write-ADReportLog -Message "Object type changed to Group" -Type Info -Terminal
        
        # Filter-Attribute für Gruppen
        $GroupFilterAttributes = @("SamAccountName", "Name", "Description", "GroupCategory", "GroupScope")
        $Global:ComboBoxFilterAttribute1.Items.Clear()
        $Global:ComboBoxFilterAttribute2.Items.Clear()
        $GroupFilterAttributes | ForEach-Object { 
            $Global:ComboBoxFilterAttribute1.Items.Add($_)
            $Global:ComboBoxFilterAttribute2.Items.Add($_)
        }
        
        # Vorauswahl auf SamAccountName setzen
        $Global:ComboBoxFilterAttribute1.SelectedItem = "SamAccountName"
        $Global:ComboBoxFilterAttribute2.SelectedItem = "Name"
        
        # Populate attributes in tabbed ListBoxes
        Set-AttributesForObjectType -ObjectType "Group"
        
        $Global:TextBlockStatus.Text = "Ready for group query"
    })

    $RadioButtonComputer.add_Checked({
        Write-ADReportLog -Message "Object type changed to Computer" -Type Info -Terminal
        
        # Filter-Attribute für Computer
        $ComputerFilterAttributes = @("Name", "DNSHostName", "OperatingSystem", "Description")
        $Global:ComboBoxFilterAttribute1.Items.Clear()
        $Global:ComboBoxFilterAttribute2.Items.Clear()
        $ComputerFilterAttributes | ForEach-Object { 
            $Global:ComboBoxFilterAttribute1.Items.Add($_)
            $Global:ComboBoxFilterAttribute2.Items.Add($_)
        }
        
        # Vorauswahl auf Name setzen (da Computer meist über Namen gesucht werden)
        $Global:ComboBoxFilterAttribute1.SelectedItem = "Name"
        $Global:ComboBoxFilterAttribute2.SelectedItem = "DNSHostName"
        
        # Populate attributes in tabbed ListBoxes
        Set-AttributesForObjectType -ObjectType "Computer"
        
        $Global:TextBlockStatus.Text = "Ready for computer query"
    })

    $RadioButtonGroupMemberships.add_Checked({
        Write-ADReportLog -Message "Object type changed to GroupMemberships" -Type Info -Terminal
        
        # Filter-Attribute für Gruppenmitgliedschaften
        $MembershipFilterAttributes = @("SamAccountName", "Name", "DisplayName")
        $Global:ComboBoxFilterAttribute1.Items.Clear()
        $Global:ComboBoxFilterAttribute2.Items.Clear()
        $MembershipFilterAttributes | ForEach-Object { 
            $Global:ComboBoxFilterAttribute1.Items.Add($_)
            $Global:ComboBoxFilterAttribute2.Items.Add($_)
        }
        
        # Vorauswahl auf SamAccountName setzen
        $Global:ComboBoxFilterAttribute1.SelectedItem = "SamAccountName"
        $Global:ComboBoxFilterAttribute2.SelectedItem = "Name"
        
        # Clear all ListBoxes for Group Memberships (no attribute selection needed)
        if ($Global:ListBoxBasicAttributes) { $Global:ListBoxBasicAttributes.Items.Clear() }
        if ($Global:ListBoxSecurityAttributes) { $Global:ListBoxSecurityAttributes.Items.Clear() }
        if ($Global:ListBoxExtendedAttributes) { $Global:ListBoxExtendedAttributes.Items.Clear() }
        Write-ADReportLog -Message "Attribute selection disabled for GroupMemberships query." -Type Info
    })

    # Event Handler fÃ¼r ButtonQueryAD anpassen, um Objekttyp zu berÃ¼cksichtigen
    $ButtonQueryAD.add_Click({
        Write-ADReportLog -Message "Executing query..." -Type Info
        try {
            # Hole Filter-Werte
            $SelectedFilterAttribute = if ($Global:ComboBoxFilterAttribute1.SelectedItem) { $Global:ComboBoxFilterAttribute1.SelectedItem.ToString() } else { "" }
            $FilterValue = $Global:TextBoxFilterValue1.Text
            $FilterOperator = if ($Global:ComboBoxFilterOperator1.SelectedItem) { $Global:ComboBoxFilterOperator1.SelectedItem.Content.ToString() } else { "Contains" }
            
            # Zweiter Filter (wenn aktiviert)
            $UseSecondFilter = $Global:CheckBoxUseSecondFilter.IsChecked
            $SelectedFilterAttribute2 = ""
            $FilterValue2 = ""
            $FilterOperator2 = "Contains"
            $FilterLogic = if ($Global:RadioButtonAnd.IsChecked) { "AND" } else { "OR" }
            
            if ($UseSecondFilter) {
                $SelectedFilterAttribute2 = if ($Global:ComboBoxFilterAttribute2.SelectedItem) { $Global:ComboBoxFilterAttribute2.SelectedItem.ToString() } else { "" }
                $FilterValue2 = $Global:TextBoxFilterValue2.Text
                $FilterOperator2 = if ($Global:ComboBoxFilterOperator2.SelectedItem) { $Global:ComboBoxFilterOperator2.SelectedItem.Content.ToString() } else { "Contains" }
            }
            
            # Get selected attributes from all three ListBoxes
            $SelectedAttributes = Get-AllSelectedAttributes
            Write-Host "DEBUG: Selected attributes: $($SelectedAttributes -join '; ')"
            $isUserSearch = $Global:RadioButtonUser.IsChecked

            if ($SelectedAttributes.Count -eq 0 -and $Global:RadioButtonGroupMemberships.IsChecked -eq $false) {
                [System.Windows.MessageBox]::Show("Please select at least one attribute for export.", "Warnung", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
                return
            }

            # Bestimme den aktuell ausgewÃ¤hlten Objekttyp
            $ObjectType = if ($Global:RadioButtonUser.IsChecked) { "User" } 
                        elseif ($Global:RadioButtonGroup.IsChecked) { "Group" } 
                        elseif ($Global:RadioButtonGroupMemberships.IsChecked) { "GroupMemberships" } 
                        else { "Computer" }
            
            # AD-Abfrage basierend auf Objekttyp durchfÃ¼hren
            $ReportData = $null
            switch ($ObjectType) {
                "User" {
                    $ReportData = Get-ADReportData -FilterAttribute $SelectedFilterAttribute -FilterValue $FilterValue -FilterOperator $FilterOperator `
                                                  -FilterAttribute2 $SelectedFilterAttribute2 -FilterValue2 $FilterValue2 -FilterOperator2 $FilterOperator2 `
                                                  -FilterLogic $FilterLogic -SelectedAttributes $SelectedAttributes -ObjectType "User"
                }
                "Group" {
                    $ReportData = Get-ADReportData -FilterAttribute $SelectedFilterAttribute -FilterValue $FilterValue -FilterOperator $FilterOperator `
                                                  -FilterAttribute2 $SelectedFilterAttribute2 -FilterValue2 $FilterValue2 -FilterOperator2 $FilterOperator2 `
                                                  -FilterLogic $FilterLogic -SelectedAttributes $SelectedAttributes -ObjectType "Group"
                }
                "Computer" {
                    $ReportData = Get-ADReportData -FilterAttribute $SelectedFilterAttribute -FilterValue $FilterValue -FilterOperator $FilterOperator `
                                                  -FilterAttribute2 $SelectedFilterAttribute2 -FilterValue2 $FilterValue2 -FilterOperator2 $FilterOperator2 `
                                                  -FilterLogic $FilterLogic -SelectedAttributes $SelectedAttributes -ObjectType "Computer"
                }
                "GroupMemberships" {
                    if (-not ([string]::IsNullOrWhiteSpace($FilterValue)) -and -not ([string]::IsNullOrWhiteSpace($SelectedFilterAttribute))) {
                        $ReportData = Get-ADGroupMembershipsReportData -FilterAttribute $SelectedFilterAttribute -FilterValue $FilterValue
                    } else {
                        Write-ADReportLog -Message "Filter attribute or value is empty for GroupMemberships query. Please specify a filter." -Type Warning
                        [System.Windows.Forms.MessageBox]::Show("Bitte geben Sie einen Filter (Attribut und Wert) fÃ¼r die Mitgliedschaftsabfrage an.", "Hinweis", "OK", "Information")
                        $ReportData = @()
                    }
                }
            }
            
            if ($ReportData) {
                try { # Inner try for processing $ReportData
                    if ($Global:DataGridResults) {
                        $Global:DataGridResults.ItemsSource = $null # Clear items first
                        if ($Global:DataGridResults.Columns) { $Global:DataGridResults.Columns.Clear() } # Clear columns
                    }
                    # Debug-Informationen
                    Write-ADReportLog -Message "ReportData Typ: $($ReportData.GetType().FullName)" -Type Info -Terminal
                    
                    # Wir brauchen sicherzustellen, dass wir immer eine Liste/Sammlung haben, auch bei einzelnen Objekten
                    # Verwende @() um es als Array zu erzwingen
                    $ReportCollection = @($ReportData)
                    
                    # Direkte Zuweisung an DataGrid
                    $Global:DataGridResults.ItemsSource = $ReportCollection
                    
                    # ZÃ¤hle die Anzahl der Ergebnisse
                    $Count = $ReportCollection.Count
                    Write-ADReportLog -Message "Query completed. $Count result(s) found for $ObjectType." -Type Info
                    
                    # ErgebniszÃ¤hler im Header aktualisieren
                    Update-ResultCounters -Results $ReportCollection
                    
                    if ($isUserSearch -and $ReportCollection.Count -eq 1 -and $ReportCollection[0].PSObject.Properties['SamAccountName']) {
                        $userSamAccountName = $ReportCollection[0].SamAccountName
                        
                        # Check if the "Mitgliedschaften" RadioButton is checked
                        if ($Global:RadioButtonGroupMemberships.IsChecked -eq $true) {
                            Write-ADReportLog -Message "Rufe Gruppenmitgliedschaften fÃ¼r Benutzer $($userSamAccountName) ab (RadioButton 'Mitgliedschaften' ist aktiv)..." -Type Info
                            $GroupMemberships = Get-UserGroupMemberships -SamAccountName $userSamAccountName
                            
                            if ($GroupMemberships -and $GroupMemberships.Count -gt 0) {
                                Write-ADReportLog -Message "$($GroupMemberships.Count) Gruppenmitgliedschaften gefunden." -Type Info
                                $DisplayData = $GroupMemberships | Select-Object @{
                                    Name = 'Benutzer';
                                    Expression = {$_.UserDisplayName}
                                }, @{
                                    Name = 'Benutzer (SAM)';
                                    Expression = {$_.UserSamAccountName}
                                }, @{
                                    Name = 'Gruppe';
                                    Expression = {$_.GroupName}
                                }, @{
                                    Name = 'Gruppe (SAM)';
                                    Expression = {$_.GroupSamAccountName}
                                }, @{
                                    Name = 'Gruppentyp';
                                    Expression = {"$($_.GroupCategory) / $($_.GroupScope)"}
                                }, @{
                                    Name = 'Beschreibung';
                                }
                            } else {
                                # RadioButton "Mitgliedschaften" is NOT checked. Display user data as usual.
                                Write-ADReportLog -Message "RadioButton 'Mitgliedschaften' ist nicht aktiv. Zeige Benutzerdaten fÃ¼r $userSamAccountName." -Type Info
                                $Global:DataGridResults.ItemsSource = $ReportCollection
                                $Global:TextBlockStatus.Text = "Benutzer $userSamAccountName gefunden. Mitgliedschaften nicht abgefragt."
                                Update-ResultCounters -Results $ReportCollection
                                Update-ResultVisualization -Results $ReportCollection
                            }
                        } else {
                            # RadioButton "Mitgliedschaften" is NOT checked. Display user data as usual.
                            Write-ADReportLog -Message "RadioButton 'Mitgliedschaften' ist nicht aktiv. Zeige Benutzerdaten fÃ¼r $userSamAccountName." -Type Info
                            $Global:DataGridResults.ItemsSource = $ReportCollection
                            $Global:TextBlockStatus.Text = "Benutzer $userSamAccountName gefunden. Mitgliedschaften nicht abgefragt."
                            Update-ResultCounters -Results $ReportCollection
                            Update-ResultVisualization -Results $ReportCollection
                        }
                    } else {
                        # This is the original 'else' for cases other than single user search.
                        # It should remain as is, displaying the $ReportCollection.
                        $Global:DataGridResults.ItemsSource = $ReportCollection
                        Update-ResultCounters -Results $ReportCollection
                        Update-ResultVisualization -Results $ReportCollection # Ensure visualization is updated here too.
                    }
                } catch {
                    $ErrorMessage = "Error in query process: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    if ($Global:DataGridResults) {
                        $Global:DataGridResults.ItemsSource = $null
                        if ($Global:DataGridResults.Columns) { $Global:DataGridResults.Columns.Clear() }
                    }
                    Update-ResultCounters -Results @() # Leeres Array fÃ¼r die ZÃ¤hler
                } 
            } else {
                Write-ADReportLog -Message "No data returned from query for $ObjectType." -Type Info
                if ($Global:DataGridResults) {
                    $Global:DataGridResults.ItemsSource = $null
                    if ($Global:DataGridResults.Columns) { $Global:DataGridResults.Columns.Clear() }
                }
                Update-ResultCounters -Results @()
                if ($Global:TextBlockStatus) { $Global:TextBlockStatus.Text = "No results found for $ObjectType." }
            }
        } catch { # Outer catch for the whole ButtonQueryAD.add_Click
            $OuterCatchErrorMessage = "An unexpected error occurred during query execution: $($_.Exception.Message)"
            Write-ADReportLog -Message $OuterCatchErrorMessage -Type Error
            try {
                if ($Global:DataGridResults) {
                    $Global:DataGridResults.ItemsSource = $null
                    if ($Global:DataGridResults.Columns) { $Global:DataGridResults.Columns.Clear() }
                }
                Update-ResultCounters -Results @()
                if ($Global:TextBlockStatus) { $Global:TextBlockStatus.Text = "Error: Query failed." } 
            } catch {
                Write-ADReportLog -Message "CRITICAL: Error within the main query button's outer catch block: $($_.Exception.Message)" -Type Error
            }
        }
    }) # End of .add_Click({
            $ButtonQuickAllUsers.add_Click({
                Write-ADReportLog -Message "Loading all users..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""
                    
                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "GivenName", "Surname", "mail", "Enabled", "LastLogonDate", "whenCreated", "LockedOut")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "*" -SelectedAttributes $QuickReportAttributes
                    Update-ADReportResults -Results $ReportData -StatusMessage "All users loaded."
                    $Global:TextBlockStatus.Text = "All users loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "All users loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading all users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickDisabledUsers.add_Click({
                Write-ADReportLog -Message "Loading disabled users..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "Enabled", "LastLogonDate")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes
                    
                    $ReportData = Get-ADReportData -CustomFilter "Enabled -eq `$false" -SelectedAttributes $QuickReportAttributes
                    Update-ADReportResults -Results $ReportData -StatusMessage "Disabled users loaded."
                    $Global:TextBlockStatus.Text = "Disabled users loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "Disabled users loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading disabled users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickLockedUsers.add_Click({
                Write-ADReportLog -Message "Loading locked out users..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "LockedOut", "LastLogonDate", "BadLogonCount")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Verwende Search-ADAccount statt Get-ADUser fÃ¼r gesperrte Konten (basierend auf Microsoft-Dokumentation)
                    $LockedOutAccounts = Search-ADAccount -LockedOut -UsersOnly -ErrorAction SilentlyContinue
                    
                    if ($LockedOutAccounts) {
                        # Hole detaillierte Informationen fÃ¼r jeden gesperrten Benutzer
                        $DetailedLockedAccounts = @()
                        foreach ($account in $LockedOutAccounts) {
                            $userDetails = Get-ADUser -Identity $account.SamAccountName -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                            if ($userDetails) {
                                $DetailedLockedAccounts += $userDetails
                            }
                        }
                        
                        if ($DetailedLockedAccounts.Count -gt 0) {
                            Update-ADReportResults -Results $DetailedLockedAccounts -StatusMessage "Locked out users loaded."
                            $Global:TextBlockStatus.Text = "Locked out users loaded. $($DetailedLockedAccounts.Count) record(s) found."
                            Write-ADReportLog -Message "Locked out users loaded. $($DetailedLockedAccounts.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No locked out users found."
                            $Global:TextBlockStatus.Text = "No locked out users found."
                            Write-ADReportLog -Message "No locked out users found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "No locked out users found."
                        $Global:TextBlockStatus.Text = "No locked out users found."
                        Write-ADReportLog -Message "No locked out users found." -Type Info
                    }
                } catch {
                    $ErrorMessage = "Error loading locked out users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickNeverExpire.add_Click({
                Write-ADReportLog -Message "Loading users with password never expires..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "PasswordNeverExpires", "Enabled")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "PasswordNeverExpires -eq `$true" -SelectedAttributes $QuickReportAttributes
                    Update-ADReportResults -Results $ReportData -StatusMessage "Users with password never expires loaded."
                    $Global:TextBlockStatus.Text = "Users with password never expires loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "Users with password never expires loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading users with password never expires: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickInactiveUsers.add_Click({
                Write-ADReportLog -Message "Loading inactive users (90 days)..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "LastLogonDate", "Enabled")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Verwende FileTime-Format fÃ¼r AD-Datumsvergleiche
                    $inactiveThreshold = (Get-Date).AddDays(-90).ToFileTime()
                    # Alternative: Lade alle Benutzer und filtere mit Where-Object
                    $AllUsers = Get-ADUser -Filter * -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                    
                    if ($AllUsers) {
                        $InactiveDate = (Get-Date).AddDays(-90)
                        $InactiveUsers = $AllUsers | Where-Object { 
                            $_.LastLogonDate -and $_.LastLogonDate -lt $InactiveDate 
                        }
                        
                        if ($InactiveUsers -and $InactiveUsers.Count -gt 0) {
                            Update-ADReportResults -Results $InactiveUsers -StatusMessage "Inactive users loaded."
                            $Global:TextBlockStatus.Text = "Inactive users loaded. $($InactiveUsers.Count) record(s) found."
                            Write-ADReportLog -Message "Inactive users loaded. $($InactiveUsers.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No inactive users found."
                            $Global:TextBlockStatus.Text = "No users inactive for more than 90 days found."
                            Write-ADReportLog -Message "No inactive users found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "Failed to query users."
                        $Global:TextBlockStatus.Text = "Failed to query users."
                        Write-ADReportLog -Message "Failed to query users." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading inactive users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickAdminUsers.add_Click({
                Write-ADReportLog -Message "Loading admin users..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    # Verbesserte Methode zum Finden von Admin-Benutzern
                    # Zuerst alle Benutzer laden
                    $AllUsers = Get-ADUser -Filter * -Properties DisplayName, SamAccountName, Enabled, LastLogonDate | Select-Object -ExcludeProperty 'PropertyNames','AddedProperties','RemovedProperties','ModifiedProperties','PropertiesCount'
                    
                    # Bekannte Admin-Gruppenbezeichnungen (deutsch und englisch)
                    $AdminGroups = @()
                    $AdminGroups += $Global:ADGroupNames.DomainAdmins
                    $AdminGroups += $Global:ADGroupNames.EnterpriseAdmins
                    $AdminGroups += $Global:ADGroupNames.Administrators
                    $AdminGroups += $Global:ADGroupNames.SchemaAdmins
                    $AdminGroups += @("AAD DC Administrators", "Azure AD-DC-Administratoren")
                    $AdminGroups += $Global:ADGroupNames.ServerOperators
                    $AdminGroups += $Global:ADGroupNames.AccountOperators
                    $AdminGroups += $Global:ADGroupNames.BackupOperators
                    
                    $AdminUsers = @()
                    # PrÃ¼fe Benutzer auf Admin-Rechte - erst SIDs der Admin-Gruppen ermitteln
                    $AdminGroupSIDs = @()
                    foreach ($groupName in $AdminGroups) {
                        try {
                            $group = Get-ADGroup -Filter "Name -eq '$groupName'" -ErrorAction SilentlyContinue
                            if ($group) {
                                $AdminGroupSIDs += $group.SID.Value
                            }
                        } catch {
                            # Ignoriere Fehler bei nicht existierenden Gruppen
                        }
                    }
                    
                    # Nur fortfahren, wenn wir Admin-Gruppen gefunden haben
                    if ($AdminGroupSIDs.Count -gt 0) {
                        # FÃ¼r jeden Benutzer die Gruppenmitgliedschaften prÃ¼fen
                        foreach ($user in $AllUsers) {
                            $memberOf = Get-ADPrincipalGroupMembership -Identity $user.SamAccountName -ErrorAction SilentlyContinue
                            foreach ($group in $memberOf) {
                                if ($AdminGroupSIDs -contains $group.SID.Value) {
                                    # Benutzer ist Mitglied einer Admin-Gruppe
                                    $adminUser = $user | Select-Object DisplayName, SamAccountName, Enabled, LastLogonDate, 
                                        @{Name='AdminGroups'; Expression={
                                            ($memberOf | Where-Object { $AdminGroupSIDs -contains $_.SID.Value } | Select-Object -ExpandProperty Name) -join ", "
                                        }}
                                    $AdminUsers += $adminUser
                                    break # Genug, wenn wir eine Admin-Gruppe gefunden haben
                                }
                            }
                        }
                    }
                    
                    if ($AdminUsers.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $null
                        $Global:DataGridResults.Columns.Clear()
                        $Global:DataGridResults.ItemsSource = $AdminUsers
                        Write-ADReportLog -Message "Administrators loaded. $($AdminUsers.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $AdminUsers
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        $Global:DataGridResults.Columns.Clear()
                        Write-ADReportLog -Message "No administrator accounts found." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading admin users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickRecentlyCreatedUsers.add_Click({
                Write-ADReportLog -Message "Loading recently created users (30 days)..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1 
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "whenCreated", "Enabled", "mail")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Lade alle Benutzer und filtere mit Where-Object für Datumsvergleiche
                    $AllUsers = Get-ADUser -Filter * -Properties $QuickReportAttributes -ErrorAction SilentlyContinue |
                        Select-Object DisplayName, SamAccountName, whenCreated, Enabled, mail |
                        Sort-Object whenCreated -Descending # Sortiere nach Erstellungsdatum absteigend
                    
                    if ($AllUsers) {
                        $CreatedDate = (Get-Date).AddDays(-30)
                        $RecentUsers = $AllUsers | Where-Object {
                            $_.whenCreated -and $_.whenCreated -gt $CreatedDate
                        }
                        
                        if ($RecentUsers -and $RecentUsers.Count -gt 0) {
                            Update-ADReportResults -Results $RecentUsers -StatusMessage "Recently created users loaded."
                            $Global:TextBlockStatus.Text = "Recently created users loaded. $($RecentUsers.Count) record(s) found."
                            Write-ADReportLog -Message "Recently created users loaded. $($RecentUsers.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No recently created users found."
                            $Global:TextBlockStatus.Text = "No users created in the last 30 days."
                            Write-ADReportLog -Message "No recently created users found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "Failed to query users."
                        $Global:TextBlockStatus.Text = "Failed to query users."
                        Write-ADReportLog -Message "Failed to query users." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading recently created users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickPasswordExpiringSoon.add_Click({
                Write-ADReportLog -Message "Loading users with password expiring soon (7 days)..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "PasswordLastSet", "Enabled", "PasswordNeverExpires", "AccountExpirationDate")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Hole die Domain Password Policy
                    $DomainPasswordPolicy = Get-ADDefaultDomainPasswordPolicy
                    $MaxPasswordAge = $DomainPasswordPolicy.MaxPasswordAge.Days
                    
                    # Lade aktivierte Benutzer deren Passwort ablaufen kann
                    $AllActiveUsers = Get-ADUser -Filter "PasswordNeverExpires -eq `$false -and Enabled -eq `$true" -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                    
                    if ($AllActiveUsers) {
                        $ExpiryThreshold = (Get-Date).AddDays(-($MaxPasswordAge - 7))
                        $UsersPasswordExpiring = $AllActiveUsers | Where-Object { 
                            $_.PasswordLastSet -and $_.PasswordLastSet -lt $ExpiryThreshold 
                        }
                        
                        if ($UsersPasswordExpiring -and $UsersPasswordExpiring.Count -gt 0) {
                            Update-ADReportResults -Results $UsersPasswordExpiring -StatusMessage "Users with password expiring soon loaded."
                            $Global:TextBlockStatus.Text = "Users with password expiring soon loaded. $($UsersPasswordExpiring.Count) record(s) found."
                            Write-ADReportLog -Message "Users with password expiring soon loaded. $($UsersPasswordExpiring.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No users with password expiring soon found."
                            $Global:TextBlockStatus.Text = "No users with passwords expiring in the next 7 days."
                            Write-ADReportLog -Message "No users with password expiring soon found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "Failed to query users."
                        $Global:TextBlockStatus.Text = "Failed to query users."
                        Write-ADReportLog -Message "Failed to query users." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading users with password expiring soon: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickGroups.add_Click({
                Write-ADReportLog -Message "Loading all groups..." -Type Info
                try {
                    $Global:RadioButtonGroup.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("Name", "SamAccountName", "Description", "GroupCategory", "GroupScope", "whenCreated")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "*" -SelectedAttributes $QuickReportAttributes -ObjectType "Group"
                    Update-ADReportResults -Results $ReportData -StatusMessage "All groups loaded."
                    $Global:TextBlockStatus.Text = "All groups loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "All groups loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading all groups: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickSecurityGroups.add_Click({
                Write-ADReportLog -Message "Loading security groups..." -Type Info
                try {
                    $Global:RadioButtonGroup.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("Name", "SamAccountName", "GroupCategory", "GroupScope")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "GroupCategory -eq 'Security'" -SelectedAttributes $QuickReportAttributes -ObjectType "Group"
                    Update-ADReportResults -Results $ReportData -StatusMessage "Security groups loaded."
                    $Global:TextBlockStatus.Text = "Security groups loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "Security groups loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading security groups: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickComputers.add_Click({
                Write-ADReportLog -Message "Loading all computers..." -Type Info
                try {
                    $Global:RadioButtonComputer.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("Name", "DNSHostName", "OperatingSystem", "Enabled", "LastLogonDate", "whenCreated")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "*" -SelectedAttributes $QuickReportAttributes -ObjectType "Computer"
                    Update-ADReportResults -Results $ReportData -StatusMessage "All computers loaded."
                    $Global:TextBlockStatus.Text = "All computers loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "All computers loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading all computers: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickInactiveComputers.add_Click({
                Write-ADReportLog -Message "Loading inactive computers (90 days)..." -Type Info
                try {
                    $Global:RadioButtonComputer.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("Name", "DNSHostName", "LastLogonDate", "Enabled", "OperatingSystem")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Lade alle Computer und filtere mit Where-Object fÃ¼r Datumsvergleiche
                    $AllComputers = Get-ADComputer -Filter * -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                    
                    if ($AllComputers) {
                        $InactiveDate = (Get-Date).AddDays(-90)
                        $InactiveComputers = $AllComputers | Where-Object { 
                            $_.LastLogonDate -and $_.LastLogonDate -lt $InactiveDate 
                        }
                        
                        if ($InactiveComputers -and $InactiveComputers.Count -gt 0) {
                            Update-ADReportResults -Results $InactiveComputers -StatusMessage "Inactive computers loaded."
                            $Global:TextBlockStatus.Text = "Inactive computers loaded. $($InactiveComputers.Count) record(s) found."
                            Write-ADReportLog -Message "Inactive computers loaded. $($InactiveComputers.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No inactive computers found."
                            $Global:TextBlockStatus.Text = "No computers inactive for more than 90 days found."
                            Write-ADReportLog -Message "No inactive computers found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "Failed to query computers."
                        $Global:TextBlockStatus.Text = "Failed to query computers."
                        Write-ADReportLog -Message "Failed to query computers." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading inactive computers: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickExpiredPasswords.add_Click({
                Write-ADReportLog -Message "Loading users with expired passwords..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "PasswordLastSet", "Enabled", "PasswordNeverExpires")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Hole die Domain Password Policy
                    $DomainPasswordPolicy = Get-ADDefaultDomainPasswordPolicy
                    $MaxPasswordAge = $DomainPasswordPolicy.MaxPasswordAge.Days
                    
                    # Lade aktivierte Benutzer deren Passwort ablaufen kann
                    $AllActiveUsers = Get-ADUser -Filter "PasswordNeverExpires -eq `$false -and Enabled -eq `$true" -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                    
                    if ($AllActiveUsers) {
                        $ExpiredDate = (Get-Date).AddDays(-$MaxPasswordAge)
                        $UsersPasswordExpired = $AllActiveUsers | Where-Object { 
                            $_.PasswordLastSet -and $_.PasswordLastSet -lt $ExpiredDate 
                        }
                        
                        if ($UsersPasswordExpired -and $UsersPasswordExpired.Count -gt 0) {
                            Update-ADReportResults -Results $UsersPasswordExpired -StatusMessage "Users with expired passwords loaded."
                            $Global:TextBlockStatus.Text = "Users with expired passwords loaded. $($UsersPasswordExpired.Count) record(s) found."
                            Write-ADReportLog -Message "Users with expired passwords loaded. $($UsersPasswordExpired.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No users with expired passwords found."
                            $Global:TextBlockStatus.Text = "No users with expired passwords found."
                            Write-ADReportLog -Message "No users with expired passwords found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "Failed to query users."
                        $Global:TextBlockStatus.Text = "Failed to query users."
                        Write-ADReportLog -Message "Failed to query users." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading users with expired passwords: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickNeverLoggedOn.add_Click({
                Write-ADReportLog -Message "Loading users who never logged on..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "LastLogonDate", "whenCreated", "Enabled")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "LastLogonDate -notlike '*'" -SelectedAttributes $QuickReportAttributes -ObjectType "User"
                    Update-ADReportResults -Results $ReportData -StatusMessage "Users who never logged on loaded."
                    $Global:TextBlockStatus.Text = "Users who never logged on loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "Users who never logged on loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading users who never logged on: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickRecentlyDeletedUsers.add_Click({
                Write-ADReportLog -Message "Loading recently deleted users..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "whenDeleted", "isDeleted", "lastKnownParent", "objectClass")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # FÃ¼r gelÃ¶schte Objekte mÃ¼ssen wir den Deleted Objects Container abfragen
                    try {
                        $deletedAfter = (Get-Date).AddDays(-30)
                        $Domain = Get-ADDomain
                        $DeletedObjectsContainer = "CN=Deleted Objects,$($Domain.DistinguishedName)"
                        
                        # Verwende Get-ADObject mit IncludeDeletedObjects
                        $DeletedUsers = Get-ADObject -Filter {(ObjectClass -eq "user") -and (whenDeleted -gt $deletedAfter)} `
                                                    -IncludeDeletedObjects `
                                                    -SearchBase $DeletedObjectsContainer `
                                                    -Properties DisplayName, SamAccountName, whenDeleted, isDeleted, lastKnownParent `
                                                    -ErrorAction SilentlyContinue | 
                                       Select-Object DisplayName, SamAccountName, whenDeleted, isDeleted, lastKnownParent, ObjectClass

                        if ($DeletedUsers) {
                            $Global:DataGridResults.ItemsSource = $DeletedUsers
                            Write-ADReportLog -Message "Recently deleted users loaded. $($DeletedUsers.Count) result(s) found." -Type Info
                            Update-ResultCounters -Results $DeletedUsers
                        } else {
                            $Global:DataGridResults.ItemsSource = $null
                            Write-ADReportLog -Message "No recently deleted users found." -Type Info
                            $Global:TextBlockStatus.Text = "No recently deleted users found in the last 30 days."
                        }
                    } catch {
                        # Fallback wenn keine Berechtigung fÃ¼r Deleted Objects
                        Write-ADReportLog -Message "Cannot access deleted objects. Requires appropriate permissions." -Type Warning
                        $Global:DataGridResults.ItemsSource = $null
                        $Global:TextBlockStatus.Text = "Access to deleted objects denied. Administrator permissions required."
                    }
                } catch {
                    $ErrorMessage = "Error loading recently deleted users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickRecentlyModifiedUsers.add_Click({
                Write-ADReportLog -Message "Loading recently modified users..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "whenChanged", "whenCreated", "Enabled", "modifyTimeStamp")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Lade alle Benutzer und filtere mit Where-Object fÃ¼r Datumsvergleiche
                    $AllUsers = Get-ADUser -Filter * -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                    
                    if ($AllUsers) {
                        $ModifiedDate = (Get-Date).AddDays(-7)
                        $ModifiedUsers = $AllUsers | Where-Object { 
                            $_.whenChanged -and $_.whenChanged -gt $ModifiedDate 
                        }
                        
                        if ($ModifiedUsers -and $ModifiedUsers.Count -gt 0) {
                            Update-ADReportResults -Results $ModifiedUsers -StatusMessage "Recently modified users loaded."
                            $Global:TextBlockStatus.Text = "Recently modified users loaded. $($ModifiedUsers.Count) record(s) found."
                            Write-ADReportLog -Message "Recently modified users loaded. $($ModifiedUsers.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "No recently modified users found."
                            $Global:TextBlockStatus.Text = "No users modified in the last 7 days."
                            Write-ADReportLog -Message "No recently modified users found." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "Failed to query users."
                        $Global:TextBlockStatus.Text = "Failed to query users."
                        Write-ADReportLog -Message "Failed to query users." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading recently modified users: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickUsersWithoutManager.add_Click({
                Write-ADReportLog -Message "Loading users without manager..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("DisplayName", "SamAccountName", "Department", "Title", "Enabled", "manager")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    # Manager-Attribut unterstÃ¼tzt nur -eq und -ne Operatoren
                    # Lade alle aktivierten Benutzer und filtere dann nach fehlenden Managern
                    $AllEnabledUsers = Get-ADUser -Filter "Enabled -eq `$true" -Properties $QuickReportAttributes -ErrorAction SilentlyContinue
                    
                    if ($AllEnabledUsers) {
                        # Filtere Benutzer ohne Manager
                        $UsersWithoutManager = $AllEnabledUsers | Where-Object { [string]::IsNullOrWhiteSpace($_.manager) }
                        
                        if ($UsersWithoutManager.Count -gt 0) {
                            Update-ADReportResults -Results $UsersWithoutManager -StatusMessage "Users without manager loaded."
                            $Global:TextBlockStatus.Text = "Users without manager loaded. $($UsersWithoutManager.Count) record(s) found."
                            Write-ADReportLog -Message "Users without manager loaded. $($UsersWithoutManager.Count) result(s) found." -Type Info
                        } else {
                            Update-ADReportResults -Results @() -StatusMessage "All users have a manager assigned."
                            $Global:TextBlockStatus.Text = "All users have a manager assigned."
                            Write-ADReportLog -Message "All users have a manager assigned." -Type Info
                        }
                    } else {
                        Update-ADReportResults -Results @() -StatusMessage "No enabled users found."
                        $Global:TextBlockStatus.Text = "No enabled users found."
                        Write-ADReportLog -Message "No enabled users found." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading users without manager: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickUsersMissingRequiredAttributes.add_Click({
                Write-ADReportLog -Message "Loading users missing required attributes..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    # Lade alle Benutzer mit wichtigen Attributen
                    $AllUsers = Get-ADUser -Filter "Enabled -eq `$true" -Properties DisplayName, SamAccountName, mail, telephoneNumber, Department, Title, manager, Enabled
                    
                    # Filtere Benutzer mit fehlenden Attributen
                    $UsersWithMissingAttributes = @()
                    foreach ($user in $AllUsers) {
                        $missingAttributes = @()
                        
                        if ([string]::IsNullOrWhiteSpace($user.DisplayName)) { $missingAttributes += "DisplayName" }
                        if ([string]::IsNullOrWhiteSpace($user.mail)) { $missingAttributes += "Email" }
                        if ([string]::IsNullOrWhiteSpace($user.telephoneNumber)) { $missingAttributes += "Phone" }
                        if ([string]::IsNullOrWhiteSpace($user.Department)) { $missingAttributes += "Department" }
                        if ([string]::IsNullOrWhiteSpace($user.Title)) { $missingAttributes += "Title" }
                        if ([string]::IsNullOrWhiteSpace($user.manager)) { $missingAttributes += "Manager" }
                        
                        if ($missingAttributes.Count -gt 0) {
                            $UsersWithMissingAttributes += [PSCustomObject]@{
                                DisplayName = $user.DisplayName
                                SamAccountName = $user.SamAccountName
                                mail = $user.mail
                                Department = $user.Department
                                Title = $user.Title
                                Enabled = $user.Enabled
                                MissingAttributes = $missingAttributes -join ", "
                                MissingCount = $missingAttributes.Count
                            }
                        }
                    }
                    
                    if ($UsersWithMissingAttributes.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $UsersWithMissingAttributes | Sort-Object MissingCount -Descending
                        Write-ADReportLog -Message "Users missing required attributes loaded. $($UsersWithMissingAttributes.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $UsersWithMissingAttributes
                        $Global:TextBlockStatus.Text = "Users missing required attributes loaded. $($UsersWithMissingAttributes.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No users with missing required attributes found." -Type Info
                        $Global:TextBlockStatus.Text = "All active users have required attributes filled."
                    }
                } catch {
                    $ErrorMessage = "Error loading users missing required attributes: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickUsersDuplicateLogonNames.add_Click({
                Write-ADReportLog -Message "Loading users with duplicate logon names..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    # Lade alle Benutzer
                    $AllUsers = Get-ADUser -Filter * -Properties DisplayName, SamAccountName, UserPrincipalName, Enabled, whenCreated
                    
                    # Gruppiere nach SamAccountName um Duplikate zu finden
                    $DuplicateSamAccounts = $AllUsers | Group-Object SamAccountName | Where-Object { $_.Count -gt 1 }
                    
                    # Gruppiere auch nach UserPrincipalName
                    $DuplicateUPNs = $AllUsers | Where-Object { $_.UserPrincipalName } | Group-Object UserPrincipalName | Where-Object { $_.Count -gt 1 }
                    
                    $DuplicateUsers = @()
                    
                    # Verarbeite SamAccountName Duplikate
                    foreach ($group in $DuplicateSamAccounts) {
                        foreach ($user in $group.Group) {
                            $DuplicateUsers += [PSCustomObject]@{
                                DisplayName = $user.DisplayName
                                SamAccountName = $user.SamAccountName
                                UserPrincipalName = $user.UserPrincipalName
                                Enabled = $user.Enabled
                                whenCreated = $user.whenCreated
                                DuplicateType = "SamAccountName"
                                DuplicateCount = $group.Count
                            }
                        }
                    }
                    
                    # Verarbeite UPN Duplikate
                    foreach ($group in $DuplicateUPNs) {
                        foreach ($user in $group.Group) {
                            # PrÃ¼fe ob dieser User nicht schon als SamAccountName Duplikat erfasst wurde
                            if (-not ($DuplicateUsers | Where-Object { $_.SamAccountName -eq $user.SamAccountName })) {
                                $DuplicateUsers += [PSCustomObject]@{
                                    DisplayName = $user.DisplayName
                                    SamAccountName = $user.SamAccountName
                                    UserPrincipalName = $user.UserPrincipalName
                                    Enabled = $user.Enabled
                                    whenCreated = $user.whenCreated
                                    DuplicateType = "UserPrincipalName"
                                    DuplicateCount = $group.Count
                                }
                            }
                        }
                    }
                    
                    if ($DuplicateUsers.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $DuplicateUsers | Sort-Object SamAccountName
                        Write-ADReportLog -Message "Users with duplicate logon names loaded. $($DuplicateUsers.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $DuplicateUsers
                        $Global:TextBlockStatus.Text = "Users with duplicate logon names loaded. $($DuplicateUsers.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No users with duplicate logon names found." -Type Info
                        $Global:TextBlockStatus.Text = "No duplicate logon names found."
                    }
                } catch {
                    $ErrorMessage = "Error loading users with duplicate logon names: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickOrphanedSIDsUsers.add_Click({
                Write-ADReportLog -Message "Loading orphaned SIDs (Foreign Security Principals)..." -Type Info
                try {
                    # Lade Foreign Security Principals
                    $Domain = Get-ADDomain
                    $ForeignSecurityPrincipals = Get-ADObject -Filter * -SearchBase "CN=ForeignSecurityPrincipals,$($Domain.DistinguishedName)" -Properties Name, ObjectSID, whenCreated, whenChanged
                    
                    $OrphanedSIDs = @()
                    
                    foreach ($fsp in $ForeignSecurityPrincipals) {
                        # Versuche den SID aufzulÃ¶sen
                        $resolved = $false
                        $resolvedName = "Unknown"
                        $sidString = $fsp.Name
                        
                        try {
                            # Versuche SID zu einem Namen aufzulÃ¶sen
                            $sid = New-Object System.Security.Principal.SecurityIdentifier($sidString)
                            $resolvedName = $sid.Translate([System.Security.Principal.NTAccount]).Value
                            $resolved = $true
                        } catch {
                            # SID konnte nicht aufgelÃ¶st werden - wahrscheinlich verwaist
                            $resolved = $false
                        }
                        
                        if (-not $resolved) {
                            $OrphanedSIDs += [PSCustomObject]@{
                                Name = $fsp.Name
                                DistinguishedName = $fsp.DistinguishedName
                                ObjectSID = $sidString
                                Status = "Orphaned"
                                ResolvedName = "Cannot resolve"
                                whenCreated = $fsp.whenCreated
                                whenChanged = $fsp.whenChanged
                            }
                        } else {
                            # Optional: Auch aufgelÃ¶ste SIDs anzeigen
                            $OrphanedSIDs += [PSCustomObject]@{
                                Name = $fsp.Name
                                DistinguishedName = $fsp.DistinguishedName
                                ObjectSID = $sidString
                                Status = "Active"
                                ResolvedName = $resolvedName
                                whenCreated = $fsp.whenCreated
                                whenChanged = $fsp.whenChanged
                            }
                        }
                    }
                    
                    # Filtere nur die verwaisten SIDs
                    $OrphanedOnly = $OrphanedSIDs | Where-Object { $_.Status -eq "Orphaned" }
                    
                    if ($OrphanedOnly.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $OrphanedOnly
                        Write-ADReportLog -Message "Orphaned SIDs loaded. $($OrphanedOnly.Count) orphaned out of $($OrphanedSIDs.Count) total FSPs found." -Type Info
                        Update-ResultCounters -Results $OrphanedOnly
                        $Global:TextBlockStatus.Text = "Orphaned SIDs loaded. $($OrphanedOnly.Count) orphaned out of $($OrphanedSIDs.Count) total FSPs."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No orphaned SIDs found. All $($OrphanedSIDs.Count) FSPs can be resolved." -Type Info
                        $Global:TextBlockStatus.Text = "No orphaned SIDs found. All Foreign Security Principals are valid."
                    }
                } catch {
                    $ErrorMessage = "Error loading orphaned SIDs: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickReversibleEncryption.add_Click({
                Write-ADReportLog -Message "Loading users with reversible encryption..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $ReversibleUsers = Get-ReversibleEncryptionUsers
                    if ($ReversibleUsers -and $ReversibleUsers.Count -gt 0) {
                        Update-ADReportResults -Results $ReversibleUsers
                        Write-ADReportLog -Message "Users with reversible encryption loaded. $($ReversibleUsers.Count) result(s) found." -Type Info
                        $Global:TextBlockStatus.Text = "Users with reversible encryption loaded. $($ReversibleUsers.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No users with reversible encryption found." -Type Info
                        $Global:TextBlockStatus.Text = "No users with reversible encryption found."
                    }
                } catch {
                    $ErrorMessage = "Error loading users with reversible encryption: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickKerberosDES.add_Click({
                Write-ADReportLog -Message "Loading users with Kerberos DES encryption..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $DESUsers = Get-KerberosDESUsers
                    if ($DESUsers -and $DESUsers.Count -gt 0) {
                        Update-ADReportResults -Results $DESUsers
                        Write-ADReportLog -Message "Users with Kerberos DES loaded. $($DESUsers.Count) result(s) found." -Type Info
                        $Global:TextBlockStatus.Text = "Users with Kerberos DES loaded. $($DESUsers.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No users with Kerberos DES found." -Type Info
                        $Global:TextBlockStatus.Text = "No users with Kerberos DES encryption found."
                    }
                } catch {
                    $ErrorMessage = "Error loading users with Kerberos DES: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickUsersWithSPN.add_Click({
                Write-ADReportLog -Message "Loading users with Service Principal Names..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $SPNUsers = Get-UsersWithSPN
                    if ($SPNUsers -and $SPNUsers.Count -gt 0) {
                        Update-ADReportResults -Results $SPNUsers
                        Write-ADReportLog -Message "Users with SPN loaded. $($SPNUsers.Count) result(s) found." -Type Info
                        $Global:TextBlockStatus.Text = "Users with SPN loaded. $($SPNUsers.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No users with SPN found." -Type Info
                        $Global:TextBlockStatus.Text = "No users with Service Principal Names found."
                    }
                } catch {
                    $ErrorMessage = "Error loading users with SPN: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickGuestAccountStatus.add_Click({
                Write-ADReportLog -Message "Analyzing Guest account status..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $GuestStatus = Get-GuestAccountStatus
                    if ($GuestStatus -and $GuestStatus.Count -gt 0) {
                        Update-ADReportResults -Results $GuestStatus
                        Write-ADReportLog -Message "Guest account analysis completed. $($GuestStatus.Count) result(s) found." -Type Info
                        $Global:TextBlockStatus.Text = "Guest account analysis completed. $($GuestStatus.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No Guest account information found." -Type Info
                        $Global:TextBlockStatus.Text = "No Guest account found."
                    }
                } catch {
                    $ErrorMessage = "Error analyzing Guest account: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickUsersByDepartment.add_Click({
                Write-ADReportLog -Message "Loading users by department..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $UsersByDept = Get-UsersByDepartment
                    if ($UsersByDept -and $UsersByDept.Count -gt 0) {
                        Update-ADReportResults -Results $UsersByDept
                        Write-ADReportLog -Message "Users by department loaded. $($UsersByDept.Count) result(s) found." -Type Info
                        $Global:TextBlockStatus.Text = "Users by department loaded. $($UsersByDept.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No department data found." -Type Info
                        $Global:TextBlockStatus.Text = "No users with department information found."
                    }
                } catch {
                    $ErrorMessage = "Error loading users by department: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickUsersByManager.add_Click({
                Write-ADReportLog -Message "Loading users by manager..." -Type Info
                try {
                    $Global:RadioButtonUser.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $UsersByMgr = Get-UsersByManager
                    if ($UsersByMgr -and $UsersByMgr.Count -gt 0) {
                        Update-ADReportResults -Results $UsersByMgr
                        Write-ADReportLog -Message "Users by manager loaded. $($UsersByMgr.Count) result(s) found." -Type Info
                        $Global:TextBlockStatus.Text = "Users by manager loaded. $($UsersByMgr.Count) record(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No manager data found." -Type Info
                        $Global:TextBlockStatus.Text = "No users with manager information found."
                    }
                } catch {
                    $ErrorMessage = "Error loading users by manager: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickDistributionGroups.add_Click({
                Write-ADReportLog -Message "Loading distribution groups..." -Type Info
                try {
                    $Global:RadioButtonGroup.IsChecked = $true
                    $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                    $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                    $Global:TextBoxFilterValue1.Text = ""
                    $Global:TextBoxFilterValue2.Text = ""

                    $QuickReportAttributes = @("Name", "SamAccountName", "GroupCategory", "GroupScope", "mail")
                    Select-AttributesInListBoxes -Attributes $QuickReportAttributes

                    $ReportData = Get-ADReportData -CustomFilter "GroupCategory -eq 'Distribution'" -SelectedAttributes $QuickReportAttributes -ObjectType "Group"
                    Update-ADReportResults -Results $ReportData -StatusMessage "Distribution groups loaded."
                    $Global:TextBlockStatus.Text = "Distribution groups loaded. $($ReportData.Count) record(s) found."
                    Write-ADReportLog -Message "Distribution groups loaded. $($ReportData.Count) result(s) found." -Type Info
                } catch {
                    $ErrorMessage = "Error loading distribution groups: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    Update-ADReportResults -Results @()
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

            $ButtonQuickWeakPasswordPolicy.add_Click({
                Write-ADReportLog -Message "Loading users with weak password policies..." -Type Info
                try {
                    $WeakPasswordUsers = Get-WeakPasswordPolicyUsers
                    if ($WeakPasswordUsers -and $WeakPasswordUsers.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $WeakPasswordUsers
                        Write-ADReportLog -Message "Users with weak password policies loaded. $($WeakPasswordUsers.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $WeakPasswordUsers
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No users with weak password policies found." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading users with weak password policies: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    $Global:DataGridResults.ItemsSource = $null
                }
            })

            $ButtonQuickRiskyGroupMemberships.add_Click({
                Write-ADReportLog -Message "Loading risky group memberships..." -Type Info
                try {
                    $RiskyMemberships = Get-RiskyGroupMemberships
                    if ($RiskyMemberships -and $RiskyMemberships.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $RiskyMemberships
                        Write-ADReportLog -Message "Risky group memberships loaded. $($RiskyMemberships.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $RiskyMemberships
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No risky group memberships found." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading risky group memberships: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    $Global:DataGridResults.ItemsSource = $null
                }
            })

            $ButtonQuickPrivilegedAccounts.add_Click({
                Write-ADReportLog -Message "Loading privileged accounts..." -Type Info
                try {
                    $PrivilegedAccounts = Get-PrivilegedAccounts
                    if ($PrivilegedAccounts -and $PrivilegedAccounts.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $PrivilegedAccounts
                        Write-ADReportLog -Message "Privileged accounts loaded. $($PrivilegedAccounts.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $PrivilegedAccounts
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No privileged accounts found." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading privileged accounts: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    $Global:DataGridResults.ItemsSource = $null
                }
            })

            $ButtonQuickOUHierarchy.add_Click({
                Write-ADReportLog -Message "Loading OU hierarchy..." -Type Info
                try {
                    $OUHierarchy = Get-ADOUHierarchyReport
                    if ($OUHierarchy -and $OUHierarchy.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $OUHierarchy
                        Write-ADReportLog -Message "OU hierarchy loaded. $($OUHierarchy.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $OUHierarchy
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No OU hierarchy information found." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading OU hierarchy: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    $Global:DataGridResults.ItemsSource = $null
                }
            })

            $ButtonQuickSitesSubnets.add_Click({
                Write-ADReportLog -Message "Loading AD sites and subnets..." -Type Info
                try {
                    $SitesSubnets = @(Get-ADSitesAndSubnetsReport)
                    if ($SitesSubnets -and $SitesSubnets.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $SitesSubnets
                        Write-ADReportLog -Message "AD sites and subnets loaded. $($SitesSubnets.Count) result(s) found." -Type Info
                        Update-ResultCounters -Results $SitesSubnets
                        $Global:TextBlockStatus.Text = "AD sites and subnets loaded. $($SitesSubnets.Count) result(s) found."
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No AD sites and subnets information found." -Type Warning
                        $Global:TextBlockStatus.Text = "No AD sites and subnets information found."
                    }
                } catch {
                    $ErrorMessage = "Error loading AD sites and subnets: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    $Global:DataGridResults.ItemsSource = $null
                    $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
                }
            })

    # --- AD-Health Event Handler ---
    $ButtonQuickFSMORoles.add_Click({
        Write-ADReportLog -Message "Loading FSMO role holders..." -Type Info
        try {
            $FSMORoles = Get-FSMORoles
            if ($FSMORoles -and $FSMORoles.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $FSMORoles
                Write-ADReportLog -Message "FSMO role holders loaded. $($FSMORoles.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $FSMORoles
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No FSMO role information found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading FSMO role holders: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickDCStatus.add_Click({
        Write-ADReportLog -Message "Loading domain controller status..." -Type Info
        try {
            $DomainControllers = Get-DomainControllerStatus
            if ($DomainControllers -and $DomainControllers.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $DomainControllers
                Write-ADReportLog -Message "Domain controller status loaded. $($DomainControllers.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $DomainControllers
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No domain controllers found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading domain controller status: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    $ButtonQuickReplicationStatus.add_Click({
        Write-ADReportLog -Message "Loading AD replication status..." -Type Info
        try {
            $ReplicationStatus = Get-ReplicationStatus
            if ($ReplicationStatus -and $ReplicationStatus.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $ReplicationStatus
                Write-ADReportLog -Message "AD replication status loaded. $($ReplicationStatus.Count) result(s) found." -Type Info
                Update-ResultCounters -Results $ReplicationStatus
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No replication status information found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading AD replication status: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })

    # --- Event Handler fÃ¼r neue erweiterte Reports ---
    
    # Organisationsstruktur-Reports
    if ($null -ne $Global:ButtonQuickDepartmentStats) {
        $ButtonQuickDepartmentStats.add_Click({
        Write-ADReportLog -Message "Loading department statistics..." -Type Info
        try {
            $DeptStats = Get-DepartmentStatistics
            if ($DeptStats -and $DeptStats.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $DeptStats
                Write-ADReportLog -Message "Department statistics loaded. $($DeptStats.Count) department(s) analyzed." -Type Info
                Update-ResultCounters -Results $DeptStats
                $Global:TextBlockStatus.Text = "Department statistics loaded. $($DeptStats.Count) department(s) analyzed."
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No department statistics found." -Type Warning
                $Global:TextBlockStatus.Text = "No department statistics found."
            }
        } catch {
            $ErrorMessage = "Error loading department statistics: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })
    }

    if ($null -ne $Global:ButtonQuickDepartmentSecurity) {
        $ButtonQuickDepartmentSecurity.add_Click({
        Write-ADReportLog -Message "Loading department security analysis..." -Type Info
        try {
            $DeptSecurity = Get-DepartmentSecurityRisks
            if ($DeptSecurity -and $DeptSecurity.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $DeptSecurity
                Write-ADReportLog -Message "Department security analysis loaded. $($DeptSecurity.Count) department(s) analyzed." -Type Info
                Update-ResultCounters -Results $DeptSecurity
                $Global:TextBlockStatus.Text = "Department security analysis loaded. $($DeptSecurity.Count) department(s) analyzed."
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No department security data found." -Type Warning
                $Global:TextBlockStatus.Text = "No department security data found."
            }
        } catch {
            $ErrorMessage = "Error loading department security analysis: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })
    }

    # Kerberos Security
    if ($null -ne $Global:ButtonQuickKerberoastable) {
        $ButtonQuickKerberoastable.add_Click({
        Write-ADReportLog -Message "Loading Kerberoastable accounts..." -Type Info
        try {
            $Kerberoastable = Get-KerberoastableAccounts
            if ($Kerberoastable -and $Kerberoastable.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $Kerberoastable
                Write-ADReportLog -Message "Kerberoastable accounts loaded. $($Kerberoastable.Count) account(s) found." -Type Info
                Update-ResultCounters -Results $Kerberoastable
                $Global:TextBlockStatus.Text = "Kerberoastable accounts loaded. $($Kerberoastable.Count) account(s) with SPNs found."
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No Kerberoastable accounts found." -Type Info
                $Global:TextBlockStatus.Text = "No Kerberoastable accounts (users with SPNs) found."
            }
        } catch {
            $ErrorMessage = "Error loading Kerberoastable accounts: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })
    }

    if ($null -ne $Global:ButtonQuickASREPRoastable) {
        $ButtonQuickASREPRoastable.add_Click({
        Write-ADReportLog -Message "Loading ASREPRoastable accounts..." -Type Info
        try {
            $ASREPRoastable = Get-ASREPRoastableAccounts
            if ($ASREPRoastable -and $ASREPRoastable.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $ASREPRoastable
                Write-ADReportLog -Message "ASREPRoastable accounts loaded. $($ASREPRoastable.Count) account(s) found." -Type Info
                Update-ResultCounters -Results $ASREPRoastable
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No ASREPRoastable accounts found." -Type Info
            }
        } catch {
            $ErrorMessage = "Error loading ASREPRoastable accounts: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
    }

    if ($null -ne $Global:ButtonQuickDelegation) {
        $ButtonQuickDelegation.add_Click({
        Write-ADReportLog -Message "Loading delegation analysis..." -Type Info
        try {
            $Delegation = Get-DelegationAnalysis
            if ($Delegation -and $Delegation.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $Delegation
                Write-ADReportLog -Message "Delegation analysis loaded. $($Delegation.Count) delegated object(s) found." -Type Info
                Update-ResultCounters -Results $Delegation
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No delegation settings found." -Type Info
            }
        } catch {
            $ErrorMessage = "Error loading delegation analysis: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
    }

    # Advanced Security
    if ($null -ne $Global:ButtonQuickDCSyncRights) {
        $ButtonQuickDCSyncRights.add_Click({
        Write-ADReportLog -Message "Loading DCSync rights analysis..." -Type Info
        try {
            $DCSyncRights = Get-DCSyncRights
            if ($DCSyncRights -and $DCSyncRights.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $DCSyncRights
                Write-ADReportLog -Message "DCSync rights analysis loaded. $($DCSyncRights.Count) identities with DCSync found." -Type Info
                Update-ResultCounters -Results $DCSyncRights
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No DCSync rights found." -Type Info
            }
        } catch {
            $ErrorMessage = "Error loading DCSync rights: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
    } else {
        Write-ADReportLog -Message "ButtonQuickDCSyncRights nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
    }

    # Advanced Security
    try {
        if ($null -ne $Global:ButtonQuickSchemaAdmins) {
            $ButtonQuickSchemaAdmins.add_Click({
        Write-ADReportLog -Message "Loading Schema Admin paths..." -Type Info
        try {
            $SchemaAdmins = Get-SchemaAdminPaths
            if ($SchemaAdmins -and $SchemaAdmins.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $SchemaAdmins
                Write-ADReportLog -Message "Schema Admin paths loaded. $($SchemaAdmins.Count) path(s) found." -Type Info
                Update-ResultCounters -Results $SchemaAdmins
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No Schema Admin paths found." -Type Info
            }
        } catch {
            $ErrorMessage = "Error loading Schema Admin paths: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickSchemaAdmins nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickSchemaAdmins: $($_.Exception.Message)" -Type Warning -Terminal
    }

    try {
        if ($null -ne $Global:ButtonQuickCertificateAnalysis) {
            $ButtonQuickCertificateAnalysis.add_Click({
        Write-ADReportLog -Message "Loading certificate security analysis..." -Type Info
        try {
            $CertAnalysis = Get-CertificateSecurityAnalysis
            if ($CertAnalysis -and $CertAnalysis.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $CertAnalysis
                Write-ADReportLog -Message "Certificate security analysis loaded. $($CertAnalysis.Count) finding(s)." -Type Info
                Update-ResultCounters -Results $CertAnalysis
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No certificate security findings." -Type Info
            }
        } catch {
            $ErrorMessage = "Error loading certificate analysis: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickCertificateAnalysis nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickCertificateAnalysis: $($_.Exception.Message)" -Type Warning -Terminal
    }

    # Advanced Monitoring
    try {
        if ($null -ne $Global:ButtonQuickSYSVOLHealth) {
            $ButtonQuickSYSVOLHealth.add_Click({
        Write-ADReportLog -Message "Loading SYSVOL health check..." -Type Info
        try {
            $SYSVOLHealth = @(Get-SYSVOLHealthCheck)
            if ($SYSVOLHealth -and $SYSVOLHealth.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $SYSVOLHealth
                Write-ADReportLog -Message "SYSVOL health check loaded. $($SYSVOLHealth.Count) DC(s) checked." -Type Info
                Update-ResultCounters -Results $SYSVOLHealth
                $Global:TextBlockStatus.Text = "SYSVOL health check loaded. $($SYSVOLHealth.Count) domain controller(s) checked."
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No SYSVOL health data found." -Type Warning
                $Global:TextBlockStatus.Text = "No SYSVOL health data found."
            }
        } catch {
            $ErrorMessage = "Error loading SYSVOL health: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
            $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickSYSVOLHealth nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickSYSVOLHealth: $($_.Exception.Message)" -Type Warning -Terminal
    }

    try {
        if ($null -ne $Global:ButtonQuickDNSHealth) {
            $ButtonQuickDNSHealth.add_Click({
        Write-ADReportLog -Message "Loading DNS health analysis..." -Type Info
        try {
            $DNSHealth = Get-DNSHealthAnalysis
            if ($DNSHealth -and $DNSHealth.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $DNSHealth
                Write-ADReportLog -Message "DNS health analysis loaded. $($DNSHealth.Count) item(s) analyzed." -Type Info
                Update-ResultCounters -Results $DNSHealth
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No DNS health data found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading DNS health: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickDNSHealth nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickDNSHealth: $($_.Exception.Message)" -Type Warning -Terminal
    }

    try {
        if ($null -ne $Global:ButtonQuickBackupStatus) {
            $ButtonQuickBackupStatus.add_Click({
        Write-ADReportLog -Message "Loading backup readiness status..." -Type Info
        try {
            $BackupStatus = Get-BackupReadinessStatus
            if ($BackupStatus -and $BackupStatus.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $BackupStatus
                Write-ADReportLog -Message "Backup readiness status loaded. $($BackupStatus.Count) check(s) performed." -Type Info
                Update-ResultCounters -Results $BackupStatus
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No backup status data found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading backup status: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickBackupStatus nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickBackupStatus: $($_.Exception.Message)" -Type Warning -Terminal
    }

    # Schema & Trusts
    try {
        if ($null -ne $Global:ButtonQuickSchemaAnalysis) {
            $ButtonQuickSchemaAnalysis.add_Click({
                Write-ADReportLog -Message "Loading schema analysis..." -Type Info
                try {
                    $SchemaAnalysis = Get-SchemaAnalysis
                    if ($SchemaAnalysis -and $SchemaAnalysis.Count -gt 0) {
                        $Global:DataGridResults.ItemsSource = $SchemaAnalysis
                        Write-ADReportLog -Message "Schema analysis loaded. $($SchemaAnalysis.Count) item(s) found." -Type Info
                        Update-ResultCounters -Results $SchemaAnalysis
                    } else {
                        $Global:DataGridResults.ItemsSource = $null
                        Write-ADReportLog -Message "No schema data found." -Type Warning
                    }
                } catch {
                    $ErrorMessage = "Error loading schema analysis: $($_.Exception.Message)"
                    Write-ADReportLog -Message $ErrorMessage -Type Error
                    $Global:DataGridResults.ItemsSource = $null
                }
            })
        } else {
            Write-ADReportLog -Message "ButtonQuickSchemaAnalysis nicht gefunden - Funktion wird übersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickSchemaAnalysis: $($_.Exception.Message)" -Type Warning -Terminal
    }

    try {
        if ($null -ne $Global:ButtonQuickTrustRelationships) {
            $ButtonQuickTrustRelationships.add_Click({
        Write-ADReportLog -Message "Loading trust relationship analysis..." -Type Info
        try {
            $TrustAnalysis = Get-TrustRelationshipAnalysis
            if ($TrustAnalysis -and $TrustAnalysis.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $TrustAnalysis
                Write-ADReportLog -Message "Trust relationship analysis loaded. $($TrustAnalysis.Count) trust(s) analyzed." -Type Info
                Update-ResultCounters -Results $TrustAnalysis
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No trust relationships found." -Type Info
            }
        } catch {
            $ErrorMessage = "Error loading trust analysis: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickTrustRelationships nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickTrustRelationships: $($_.Exception.Message)" -Type Warning -Terminal
    }

    try {
        if ($null -ne $Global:ButtonQuickQuotasLimits) {
            $ButtonQuickQuotasLimits.add_Click({
        Write-ADReportLog -Message "Loading quotas and limits analysis..." -Type Info
        try {
            $QuotasLimits = Get-QuotasAndLimits
            if ($QuotasLimits -and $QuotasLimits.Count -gt 0) {
                $Global:DataGridResults.ItemsSource = $QuotasLimits
                Write-ADReportLog -Message "Quotas and limits analysis loaded. $($QuotasLimits.Count) item(s) analyzed." -Type Info
                Update-ResultCounters -Results $QuotasLimits
            } else {
                $Global:DataGridResults.ItemsSource = $null
                Write-ADReportLog -Message "No quota data found." -Type Warning
            }
        } catch {
            $ErrorMessage = "Error loading quotas and limits: $($_.Exception.Message)"
            Write-ADReportLog -Message $ErrorMessage -Type Error
            $Global:DataGridResults.ItemsSource = $null
        }
    })
        } else {
            Write-ADReportLog -Message "ButtonQuickQuotasLimits nicht gefunden - Funktion wird Ã¼bersprungen." -Type Warning -Terminal
        }
    } catch {
        Write-ADReportLog -Message "Fehler beim Initialisieren von ButtonQuickQuotasLimits: $($_.Exception.Message)" -Type Warning -Terminal
    }

    $ButtonExportCSV.add_Click({
        Write-ADReportLog -Message "Preparing CSV export..." -Type Info
        if ($null -eq $Global:DataGridResults.ItemsSource -or $Global:DataGridResults.Items.Count -eq 0) {
            Write-ADReportLog -Message "No data available for export." -Type Warning
            [System.Windows.Forms.MessageBox]::Show("Es sind keine Daten zum Exportieren vorhanden.", "Hinweis", "OK", "Information") | Out-Null
            return
        }

        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "CSV (Comma delimited) (*.csv)|*.csv"
        $SaveFileDialog.Title = "Save CSV file as"
        $SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")

        if ($SaveFileDialog.ShowDialog() -eq "OK") {
            $FilePath = $SaveFileDialog.FileName
            try {
                $Global:DataGridResults.ItemsSource | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
                $Global:TextBlockStatus.Text = "Daten erfolgreich nach $FilePath exportiert."
                [System.Windows.Forms.MessageBox]::Show("Data successfully exported!", "CSV Export", "OK", "Information")
            } catch {
                $ErrorMessage = "Error in CSV export: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren der Daten: $($_.Exception.Message)", "Export Fehler", "OK", "Error")
            }
        } else {
            Write-ADReportLog -Message "CSV export canceled by user." -Type Info
        }
    })

    $ButtonExportHTML.add_Click({
        Write-ADReportLog -Message "Preparing HTML export..." -Type Info
        if ($null -eq $Global:DataGridResults.ItemsSource -or $Global:DataGridResults.Items.Count -eq 0) {
            $Global:TextBlockStatus.Text = "No data available for export."
            [System.Windows.Forms.MessageBox]::Show("Es sind keine Daten zum Exportieren vorhanden.", "Hinweis", "OK", "Information")
            return
        }

        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "HTML File (*.html;*.htm)|*.html;*.htm"
        $SaveFileDialog.Title = "Save HTML file as"
        $SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")

        if ($SaveFileDialog.ShowDialog() -eq "OK") {
            $FilePath = $SaveFileDialog.FileName
            try {
                # Create a more attractive HTML header
                $HtmlHead = @"
<meta http-equiv='Content-Type' content='text/html; charset=utf-8' />
<title>Active Directory Report</title>
<style>
  body { font-family: Segoe UI, Arial, sans-serif; margin: 20px; }
  table { border-collapse: collapse; width: 90%; margin: 20px auto; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
  th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
  th { background-color: #0078D4; color: white; }
  tr:nth-child(even) { background-color: #f2f2f2; }
  h1 { text-align: center; color: #333; }
</style>
"@
                $DateTimeNow = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
                $ReportTitle = "Active Directory Report - Created on $DateTimeNow"

                $Global:DataGridResults.ItemsSource | ConvertTo-Html -Head $HtmlHead -Body "<h1>$ReportTitle</h1>" | Out-File -FilePath $FilePath -Encoding UTF8
                $Global:TextBlockStatus.Text = "Daten erfolgreich nach $FilePath exportiert."
                [System.Windows.Forms.MessageBox]::Show("Data successfully exported!", "HTML Export", "OK", "Information")
            } catch {
                $ErrorMessage = "Error in HTML export: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren der Daten: $($_.Exception.Message)", "Export Fehler", "OK", "Error")
            }
        } else {
            Write-ADReportLog -Message "HTML export canceled by user." -Type Info
        }
    })

    # --- Event Handler fÃ¼r Attribute Selection Buttons ---
    if ($null -ne $Global:ButtonSelectAllAttributes) {
        $ButtonSelectAllAttributes.add_Click({
            Write-ADReportLog -Message "Selecting all attributes..." -Type Info
            try {
                # Select all items in all three ListBoxes
                if ($Global:ListBoxBasicAttributes) {
                    foreach ($item in $Global:ListBoxBasicAttributes.Items) {
                        $item.IsSelected = $true
                    }
                }
                if ($Global:ListBoxSecurityAttributes) {
                    foreach ($item in $Global:ListBoxSecurityAttributes.Items) {
                        $item.IsSelected = $true
                    }
                }
                if ($Global:ListBoxExtendedAttributes) {
                    foreach ($item in $Global:ListBoxExtendedAttributes.Items) {
                        $item.IsSelected = $true
                    }
                }
                Write-ADReportLog -Message "All attributes selected." -Type Info
            } catch {
                Write-ADReportLog -Message "Error selecting all attributes: $($_.Exception.Message)" -Type Error
            }
        })
    }

    if ($null -ne $Global:ButtonSelectNoneAttributes) {
        $ButtonSelectNoneAttributes.add_Click({
            Write-ADReportLog -Message "Deselecting all attributes..." -Type Info
            try {
                # Deselect all items in all three ListBoxes
                if ($Global:ListBoxBasicAttributes) {
                    foreach ($item in $Global:ListBoxBasicAttributes.Items) {
                        $item.IsSelected = $false
                    }
                }
                if ($Global:ListBoxSecurityAttributes) {
                    foreach ($item in $Global:ListBoxSecurityAttributes.Items) {
                        $item.IsSelected = $false
                    }
                }
                if ($Global:ListBoxExtendedAttributes) {
                    foreach ($item in $Global:ListBoxExtendedAttributes.Items) {
                        $item.IsSelected = $false
                    }
                }
                Write-ADReportLog -Message "All attributes deselected." -Type Info
            } catch {
                Write-ADReportLog -Message "Error deselecting all attributes: $($_.Exception.Message)" -Type Error
            }
        })
    }

    # --- Event Handler für Refresh und Copy Buttons ---
    if ($null -ne $Global:ButtonRefresh) {
        $ButtonRefresh.add_Click({
            Write-ADReportLog -Message "Setze Query-Fenster zurück..." -Type Info
            try {
                # Setze RadioButtons zurück
                if ($Global:RadioButtonUser -and $Global:RadioButtonUser.GetType().GetProperty('IsChecked')) {
                    $Global:RadioButtonUser.IsChecked = $true
                }
                if ($Global:RadioButtonGroup -and $Global:RadioButtonGroup.GetType().GetProperty('IsChecked')) {
                    $Global:RadioButtonGroup.IsChecked = $false
                }
                if ($Global:RadioButtonComputer -and $Global:RadioButtonComputer.GetType().GetProperty('IsChecked')) {
                    $Global:RadioButtonComputer.IsChecked = $false
                }
                if ($Global:RadioButtonGroupMembership -and $Global:RadioButtonGroupMembership.GetType().GetProperty('IsChecked')) {
                    $Global:RadioButtonGroupMembership.IsChecked = $false
                }

                # Leere alle ListBoxen
                if ($Global:ListBoxBasicAttributes) {
                    foreach ($item in $Global:ListBoxBasicAttributes.Items) {
                        $item.IsSelected = $false
                    }
                }
                if ($Global:ListBoxSecurityAttributes) {
                    foreach ($item in $Global:ListBoxSecurityAttributes.Items) {
                        $item.IsSelected = $false
                    }
                }
                if ($Global:ListBoxExtendedAttributes) {
                    foreach ($item in $Global:ListBoxExtendedAttributes.Items) {
                        $item.IsSelected = $false
                    }
                }

                # Leere das DataGrid
                if ($Global:DataGridResults) {
                    $Global:DataGridResults.ItemsSource = $null
                }

                Write-ADReportLog -Message "Query-Fenster erfolgreich zurückgesetzt." -Type Info
            }
            catch {
                Write-ADReportLog -Message "Fehler beim Zurücksetzen des Query-Fensters: $($_.Exception.Message)" -Type Error
            }
        })
    }

    if ($null -ne $Global:ButtonCopy) {
        $ButtonCopy.add_Click({
            Write-ADReportLog -Message "Prüfe auf markierte Zeilen..." -Type Info
            try {
                # Prüfe ob Zeilen markiert sind
                $selectedItems = $Global:DataGridResults.SelectedItems
                
                if ($selectedItems.Count -gt 0) {
                    Write-ADReportLog -Message "Kopiere markierte Zeilen in die Zwischenablage..." -Type Info
                    
                    # Konvertiere ausgewählte Zeilen in tabulierte Textform
                    $clipboardText = ""
                    
                    # Header
                    $columns = $Global:DataGridResults.Columns | Where-Object {$_.Visibility -eq 'Visible'}
                    $clipboardText += $columns.Header -join "`t"
                    $clipboardText += "`r`n"
                    
                    # Nur markierte Zeilen
                    foreach ($item in $selectedItems) {
                        $rowData = @()
                        foreach ($column in $columns) {
                            $cellValue = $item.$($column.Header)
                            if ($null -eq $cellValue) { $cellValue = "" }
                            $rowData += $cellValue.ToString()
                        }
                        $clipboardText += $rowData -join "`t"
                        $clipboardText += "`r`n"
                    }
                    
                    # In Zwischenablage kopieren
                    [System.Windows.Clipboard]::SetText($clipboardText)
                    Write-ADReportLog -Message "Markierte Zeilen erfolgreich in die Zwischenablage kopiert." -Type Info
                }
                else {
                    [System.Windows.MessageBox]::Show(
                        "Bitte markieren Sie mindestens eine Zeile zum Kopieren.",
                        "Keine Auswahl",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Information
                    )
                    Write-ADReportLog -Message "Keine Zeilen markiert zum Kopieren." -Type Warning
                }
            }
            catch {
                Write-ADReportLog -Message "Fehler beim Kopieren in die Zwischenablage: $($_.Exception.Message)" -Type Error
            }
        })
    }


    # --- Event Handler fÃ¼r Help and About Buttons ---
    if ($null -ne $Global:ButtonHelp) {
        $ButtonHelp.add_Click({
            Write-ADReportLog -Message "Showing help dialog..." -Type Info
            [System.Windows.MessageBox]::Show(
                "easyADReport Help:`n`n" +
                "1. Select Object Type: Choose between Users, Groups, Computers, or Group Memberships`n`n" +
                "2. Set Filters: Configure search filters using attribute, operator, and value`n`n" +
                "3. Select Attributes: Choose which attributes to include in the report`n`n" +
                "4. Quick Reports: Use predefined reports for common AD queries`n`n" +
                "5. Export Options: Export results to CSV or HTML format`n`n" +
                "For more information, visit the documentation.",
                "Help",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            ) | Out-Null
        })
    }

    if ($null -ne $Global:ButtonAbout) {
        $ButtonAbout.add_Click({
            Write-ADReportLog -Message "Showing about dialog..." -Type Info
            [System.Windows.MessageBox]::Show(
                "easyADReport v0.3.3`n`n" +
                "A comprehensive Active Directory reporting tool`n`n" +
                "Features:`n" +
                "â€¢ Advanced filtering and search capabilities`n" +
                "â€¢ Multiple export formats (CSV, HTML)`n" +
                "â€¢ Quick report templates for common scenarios`n" +
                "â€¢ Security and compliance reports`n" +
                "â€¢ Multi-language support`n`n" +
                "Â© 2024 Your Organization",
                "About easyADReport",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            ) | Out-Null
        })
    }

    # --- Event Handler fÃ¼r neue Quick Report Buttons ---
    
    # Event Handler fÃ¼r ButtonQuickStalePasswords
    if ($null -ne $Global:ButtonQuickStalePasswords) {
        $ButtonQuickStalePasswords.add_Click({
            Write-ADReportLog -Message "Loading users with stale passwords..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                $Global:TextBoxFilterValue1.Text = ""
                $Global:TextBoxFilterValue2.Text = ""

                $StalePasswords = Get-StalePasswords -Days 90
                if ($StalePasswords -and $StalePasswords.Count -gt 0) {
                    Update-ADReportResults -Results $StalePasswords
                    Write-ADReportLog -Message "Stale passwords loaded. $($StalePasswords.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $StalePasswords
                    $Global:TextBlockStatus.Text = "Stale passwords loaded. $($StalePasswords.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No users with stale passwords found." -Type Info
                    $Global:TextBlockStatus.Text = "No users with stale passwords found."
                }
            } catch {
                $ErrorMessage = "Error loading stale passwords: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler fÃ¼r ButtonQuickNeverChangingPasswords
    if ($null -ne $Global:ButtonQuickNeverChangingPasswords) {
        $ButtonQuickNeverChangingPasswords.add_Click({
            Write-ADReportLog -Message "Loading users with never changing passwords..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                $Global:ComboBoxFilterAttribute1.SelectedIndex = -1
                $Global:ComboBoxFilterAttribute2.SelectedIndex = -1
                $Global:TextBoxFilterValue1.Text = ""
                $Global:TextBoxFilterValue2.Text = ""

                $NeverChangingPasswords = Get-NeverChangingPasswords -Days 365
                if ($NeverChangingPasswords -and $NeverChangingPasswords.Count -gt 0) {
                    Update-ADReportResults -Results $NeverChangingPasswords
                    Write-ADReportLog -Message "Never changing passwords loaded. $($NeverChangingPasswords.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $NeverChangingPasswords
                    $Global:TextBlockStatus.Text = "Never changing passwords loaded. $($NeverChangingPasswords.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No users with never changing passwords found." -Type Info
                    $Global:TextBlockStatus.Text = "No users with never changing passwords found."
                }
            } catch {
                $ErrorMessage = "Error loading never changing passwords: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler fÃ¼r ButtonQuickEmptyGroups
    if ($null -ne $Global:ButtonQuickEmptyGroups) {
        $ButtonQuickEmptyGroups.add_Click({
            Write-ADReportLog -Message "Loading empty groups..." -Type Info
            try {
                $Global:RadioButtonGroup.IsChecked = $true
                
                $EmptyGroups = Get-EmptyGroups
                if ($EmptyGroups -and $EmptyGroups.Count -gt 0) {
                    Update-ADReportResults -Results $EmptyGroups
                    Write-ADReportLog -Message "Empty groups loaded. $($EmptyGroups.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $EmptyGroups
                    $Global:TextBlockStatus.Text = "Empty groups loaded. $($EmptyGroups.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No empty groups found." -Type Info
                    $Global:TextBlockStatus.Text = "No empty groups found."
                }
            } catch {
                $ErrorMessage = "Error loading empty groups: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickNestedGroups
    if ($null -ne $Global:ButtonQuickNestedGroups) {
        $ButtonQuickNestedGroups.add_Click({
            Write-ADReportLog -Message "Loading nested groups..." -Type Info
            try {
                $Global:RadioButtonGroup.IsChecked = $true
                
                $NestedGroups = Get-NestedGroups
                if ($NestedGroups -and $NestedGroups.Count -gt 0) {
                    Update-ADReportResults -Results $NestedGroups
                    Write-ADReportLog -Message "Nested groups loaded. $($NestedGroups.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $NestedGroups
                    $Global:TextBlockStatus.Text = "Nested groups loaded. $($NestedGroups.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No nested groups found." -Type Info
                    $Global:TextBlockStatus.Text = "No nested groups found."
                }
            } catch {
                $ErrorMessage = "Error loading nested groups: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickCircularGroups
    if ($null -ne $Global:ButtonQuickCircularGroups) {
        $ButtonQuickCircularGroups.add_Click({
            Write-ADReportLog -Message "Loading circular nested groups..." -Type Info
            try {
                $Global:RadioButtonGroup.IsChecked = $true
                
                $CircularGroups = Get-CircularNestedGroups
                if ($CircularGroups -and $CircularGroups.Count -gt 0) {
                    Update-ADReportResults -Results $CircularGroups
                    Write-ADReportLog -Message "Circular nested groups loaded. $($CircularGroups.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $CircularGroups
                    $Global:TextBlockStatus.Text = "Circular nested groups loaded. $($CircularGroups.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No circular nested groups found." -Type Info
                    $Global:TextBlockStatus.Text = "No circular nested groups found."
                }
            } catch {
                $ErrorMessage = "Error loading circular nested groups: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickGroupsByTypeScope
    if ($null -ne $Global:ButtonQuickGroupsByTypeScope) {
        $ButtonQuickGroupsByTypeScope.add_Click({
            Write-ADReportLog -Message "Loading groups by type and scope..." -Type Info
            try {
                $Global:RadioButtonGroup.IsChecked = $true
                
                $GroupsByTypeScope = Get-GroupsByTypeAndScope
                if ($GroupsByTypeScope -and $GroupsByTypeScope.Count -gt 0) {
                    Update-ADReportResults -Results $GroupsByTypeScope
                    Write-ADReportLog -Message "Groups by type and scope loaded. $($GroupsByTypeScope.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $GroupsByTypeScope
                    $Global:TextBlockStatus.Text = "Groups by type and scope loaded. $($GroupsByTypeScope.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No groups found." -Type Info
                    $Global:TextBlockStatus.Text = "No groups found."
                }
            } catch {
                $ErrorMessage = "Error loading groups by type and scope: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickDynamicDistGroups
    if ($null -ne $Global:ButtonQuickDynamicDistGroups) {
        $ButtonQuickDynamicDistGroups.add_Click({
            Write-ADReportLog -Message "Loading dynamic distribution groups..." -Type Info
            try {
                $Global:RadioButtonGroup.IsChecked = $true
                
                $DynamicGroups = Get-DynamicDistributionGroups
                if ($DynamicGroups -and $DynamicGroups.Count -gt 0) {
                    Update-ADReportResults -Results $DynamicGroups
                    Write-ADReportLog -Message "Dynamic distribution groups loaded. $($DynamicGroups.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $DynamicGroups
                    $Global:TextBlockStatus.Text = "Dynamic distribution groups loaded. $($DynamicGroups.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No dynamic distribution groups found." -Type Info
                    $Global:TextBlockStatus.Text = "No dynamic distribution groups found."
                }
            } catch {
                $ErrorMessage = "Error loading dynamic distribution groups: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickMailEnabledGroups
    if ($null -ne $Global:ButtonQuickMailEnabledGroups) {
        $ButtonQuickMailEnabledGroups.add_Click({
            Write-ADReportLog -Message "Loading mail-enabled groups..." -Type Info
            try {
                $Global:RadioButtonGroup.IsChecked = $true
                
                $MailEnabledGroups = Get-MailEnabledGroups
                if ($MailEnabledGroups -and $MailEnabledGroups.Count -gt 0) {
                    Update-ADReportResults -Results $MailEnabledGroups
                    Write-ADReportLog -Message "Mail-enabled groups loaded. $($MailEnabledGroups.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $MailEnabledGroups
                    $Global:TextBlockStatus.Text = "Mail-enabled groups loaded. $($MailEnabledGroups.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No mail-enabled groups found." -Type Info
                    $Global:TextBlockStatus.Text = "No mail-enabled groups found."
                }
            } catch {
                $ErrorMessage = "Error loading mail-enabled groups: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickGroupsWithoutOwners
    if ($null -ne $Global:ButtonQuickGroupsWithoutOwners) {
        $ButtonQuickGroupsWithoutOwners.add_Click({
            Write-ADReportLog -Message "Loading groups without owners..." -Type Info
            try {
                $Global:RadioButtonGroup.IsChecked = $true
                
                $GroupsWithoutOwners = Get-GroupsWithoutOwners
                if ($GroupsWithoutOwners -and $GroupsWithoutOwners.Count -gt 0) {
                    Update-ADReportResults -Results $GroupsWithoutOwners
                    Write-ADReportLog -Message "Groups without owners loaded. $($GroupsWithoutOwners.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $GroupsWithoutOwners
                    $Global:TextBlockStatus.Text = "Groups without owners loaded. $($GroupsWithoutOwners.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No groups without owners found." -Type Info
                    $Global:TextBlockStatus.Text = "No groups without owners found."
                }
            } catch {
                $ErrorMessage = "Error loading groups without owners: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickLargeGroups
    if ($null -ne $Global:ButtonQuickLargeGroups) {
        $ButtonQuickLargeGroups.add_Click({
            Write-ADReportLog -Message "Loading large groups..." -Type Info
            try {
                $Global:RadioButtonGroup.IsChecked = $true
                
                $LargeGroups = Get-LargeGroups -Threshold 100
                if ($LargeGroups -and $LargeGroups.Count -gt 0) {
                    Update-ADReportResults -Results $LargeGroups
                    Write-ADReportLog -Message "Large groups loaded. $($LargeGroups.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $LargeGroups
                    $Global:TextBlockStatus.Text = "Large groups loaded. $($LargeGroups.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No large groups found." -Type Info
                    $Global:TextBlockStatus.Text = "No large groups found."
                }
            } catch {
                $ErrorMessage = "Error loading large groups: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickRecentlyModifiedGroups
    if ($null -ne $Global:ButtonQuickRecentlyModifiedGroups) {
        $ButtonQuickRecentlyModifiedGroups.add_Click({
            Write-ADReportLog -Message "Loading recently modified groups..." -Type Info
            try {
                $Global:RadioButtonGroup.IsChecked = $true
                
                $RecentlyModifiedGroups = Get-RecentlyModifiedGroups -Days 30
                if ($RecentlyModifiedGroups -and $RecentlyModifiedGroups.Count -gt 0) {
                    Update-ADReportResults -Results $RecentlyModifiedGroups
                    Write-ADReportLog -Message "Recently modified groups loaded. $($RecentlyModifiedGroups.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $RecentlyModifiedGroups
                    $Global:TextBlockStatus.Text = "Recently modified groups loaded. $($RecentlyModifiedGroups.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No recently modified groups found." -Type Info
                    $Global:TextBlockStatus.Text = "No recently modified groups found."
                }
            } catch {
                $ErrorMessage = "Error loading recently modified groups: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickOSSummary
    if ($null -ne $Global:ButtonQuickOSSummary) {
        $ButtonQuickOSSummary.add_Click({
            Write-ADReportLog -Message "Analyzing OS summary..." -Type Info
            try {
                $Global:RadioButtonComputer.IsChecked = $true
                
                $OSSummary = Get-OSSummary
                if ($OSSummary -and $OSSummary.Count -gt 0) {
                    Update-ADReportResults -Results $OSSummary
                    Write-ADReportLog -Message "OS summary analysis completed. $($OSSummary.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $OSSummary
                    $Global:TextBlockStatus.Text = "OS summary analysis completed. $($OSSummary.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No OS summary data found." -Type Info
                    $Global:TextBlockStatus.Text = "No OS summary data found."
                }
            } catch {
                $ErrorMessage = "Error analyzing OS summary: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickComputersByOSVersion
    if ($null -ne $Global:ButtonQuickComputersByOSVersion) {
        $ButtonQuickComputersByOSVersion.add_Click({
            Write-ADReportLog -Message "Analyzing computers by OS version..." -Type Info
            try {
                $Global:RadioButtonComputer.IsChecked = $true
                
                $ComputersByOS = Get-ComputersByOSVersion
                if ($ComputersByOS -and $ComputersByOS.Count -gt 0) {
                    Update-ADReportResults -Results $ComputersByOS
                    Write-ADReportLog -Message "OS version analysis completed. $($ComputersByOS.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $ComputersByOS
                    $Global:TextBlockStatus.Text = "OS version analysis completed. $($ComputersByOS.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No computers found." -Type Info
                    $Global:TextBlockStatus.Text = "No computers found."
                }
            } catch {
                $ErrorMessage = "Error analyzing computers by OS version: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickBitLockerStatus
    if ($null -ne $Global:ButtonQuickBitLockerStatus) {
        $ButtonQuickBitLockerStatus.add_Click({
            Write-ADReportLog -Message "Analyzing BitLocker status..." -Type Info
            try {
                $Global:RadioButtonComputer.IsChecked = $true
                
                $BitLockerStatus = Get-BitLockerStatus
                if ($BitLockerStatus -and $BitLockerStatus.Count -gt 0) {
                    Update-ADReportResults -Results $BitLockerStatus
                    Write-ADReportLog -Message "BitLocker status analysis completed. $($BitLockerStatus.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $BitLockerStatus
                    $Global:TextBlockStatus.Text = "BitLocker status analysis completed. $($BitLockerStatus.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No BitLocker information found." -Type Info
                    $Global:TextBlockStatus.Text = "No BitLocker information found."
                }
            } catch {
                $ErrorMessage = "Error analyzing BitLocker status: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickStaleComputerPasswords
    if ($null -ne $Global:ButtonQuickStaleComputerPasswords) {
        $ButtonQuickStaleComputerPasswords.add_Click({
            Write-ADReportLog -Message "Analyzing stale computer passwords..." -Type Info
            try {
                $Global:RadioButtonComputer.IsChecked = $true
                
                $StalePasswords = Get-StaleComputerPasswords
                if ($StalePasswords -and $StalePasswords.Count -gt 0) {
                    Update-ADReportResults -Results $StalePasswords
                    Write-ADReportLog -Message "Stale password analysis completed. $($StalePasswords.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $StalePasswords
                    $Global:TextBlockStatus.Text = "Stale password analysis completed. $($StalePasswords.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No stale computer passwords found." -Type Info
                    $Global:TextBlockStatus.Text = "No stale computer passwords found."
                }
            } catch {
                $ErrorMessage = "Error analyzing stale computer passwords: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickComputersNeverLoggedOn
    if ($null -ne $Global:ButtonQuickComputersNeverLoggedOn) {
        $ButtonQuickComputersNeverLoggedOn.add_Click({
            Write-ADReportLog -Message "Analyzing computers that never logged on..." -Type Info
            try {
                $Global:RadioButtonComputer.IsChecked = $true
                
                $NeverLoggedOn = Get-ComputersNeverLoggedOn
                if ($NeverLoggedOn -and $NeverLoggedOn.Count -gt 0) {
                    Update-ADReportResults -Results $NeverLoggedOn
                    Write-ADReportLog -Message "Never logged on analysis completed. $($NeverLoggedOn.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $NeverLoggedOn
                    $Global:TextBlockStatus.Text = "Never logged on analysis completed. $($NeverLoggedOn.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No computers found that never logged on." -Type Info
                    $Global:TextBlockStatus.Text = "No computers found that never logged on."
                }
            } catch {
                $ErrorMessage = "Error analyzing computers that never logged on: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickDuplicateComputerNames
    if ($null -ne $Global:ButtonQuickDuplicateComputerNames) {
        $ButtonQuickDuplicateComputerNames.add_Click({
            Write-ADReportLog -Message "Analyzing duplicate computer names..." -Type Info
            try {
                $Global:RadioButtonComputer.IsChecked = $true
                
                $DuplicateNames = Get-DuplicateComputerNames
                if ($DuplicateNames -and $DuplicateNames.Count -gt 0) {
                    Update-ADReportResults -Results $DuplicateNames
                    Write-ADReportLog -Message "Duplicate names analysis completed. $($DuplicateNames.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $DuplicateNames
                    $Global:TextBlockStatus.Text = "Duplicate names analysis completed. $($DuplicateNames.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No duplicate computer names found." -Type Info
                    $Global:TextBlockStatus.Text = "No duplicate computer names found."
                }
            } catch {
                $ErrorMessage = "Error analyzing duplicate computer names: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickComputersByLocation
    if ($null -ne $Global:ButtonQuickComputersByLocation) {
        $ButtonQuickComputersByLocation.add_Click({
            Write-ADReportLog -Message "Analyzing computers by location..." -Type Info
            try {
                $Global:RadioButtonComputer.IsChecked = $true
                
                $ComputersByLocation = Get-ComputersByLocation
                if ($ComputersByLocation -and $ComputersByLocation.Count -gt 0) {
                    Update-ADReportResults -Results $ComputersByLocation
                    Write-ADReportLog -Message "Location analysis completed. $($ComputersByLocation.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $ComputersByLocation
                    $Global:TextBlockStatus.Text = "Location analysis completed. $($ComputersByLocation.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No computer location data found." -Type Info
                    $Global:TextBlockStatus.Text = "No computer location data found."
                }
            } catch {
                $ErrorMessage = "Error analyzing computers by location: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }
    
    # Event Handler fÃ¼r ButtonQuickServiceAccountsOverview
    if ($null -ne $Global:ButtonQuickServiceAccountsOverview) {
        $ButtonQuickServiceAccountsOverview.add_Click({
            Write-ADReportLog -Message "Loading service accounts overview..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $ServiceAccounts = Get-ServiceAccountsOverview
                if ($ServiceAccounts -and $ServiceAccounts.Count -gt 0) {
                    Update-ADReportResults -Results $ServiceAccounts
                    Write-ADReportLog -Message "Service accounts loaded. $($ServiceAccounts.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $ServiceAccounts
                    $Global:TextBlockStatus.Text = "Service accounts loaded. $($ServiceAccounts.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No service accounts found." -Type Info
                    $Global:TextBlockStatus.Text = "No service accounts found."
                }
            } catch {
                $ErrorMessage = "Error loading service accounts: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickManagedServiceAccounts
    if ($null -ne $Global:ButtonQuickManagedServiceAccounts) {
        $ButtonQuickManagedServiceAccounts.add_Click({
            Write-ADReportLog -Message "Analyzing managed service accounts..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $ManagedAccounts = Get-ManagedServiceAccounts
                if ($ManagedAccounts -and $ManagedAccounts.Count -gt 0) {
                    Update-ADReportResults -Results $ManagedAccounts
                    Write-ADReportLog -Message "Managed service accounts loaded. $($ManagedAccounts.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $ManagedAccounts
                    $Global:TextBlockStatus.Text = "Managed service accounts loaded. $($ManagedAccounts.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No managed service accounts found." -Type Info
                    $Global:TextBlockStatus.Text = "No managed service accounts found."
                }
            } catch {
                $ErrorMessage = "Error loading managed service accounts: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickServiceAccountsSPN 
    if ($null -ne $Global:ButtonQuickServiceAccountsSPN) {
        $ButtonQuickServiceAccountsSPN.add_Click({
            Write-ADReportLog -Message "Analyzing service accounts with SPNs..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $SPNAccounts = Get-ServiceAccountsSPN
                if ($SPNAccounts -and $SPNAccounts.Count -gt 0) {
                    Update-ADReportResults -Results $SPNAccounts
                    Write-ADReportLog -Message "Service accounts with SPNs loaded. $($SPNAccounts.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $SPNAccounts
                    $Global:TextBlockStatus.Text = "Service accounts with SPNs loaded. $($SPNAccounts.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No service accounts with SPNs found." -Type Info
                    $Global:TextBlockStatus.Text = "No service accounts with SPNs found."
                }
            } catch {
                $ErrorMessage = "Error loading service accounts with SPNs: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickHighPrivServiceAccounts
    if ($null -ne $Global:ButtonQuickHighPrivServiceAccounts) {
        $ButtonQuickHighPrivServiceAccounts.add_Click({
            Write-ADReportLog -Message "Analyzing high privileged service accounts..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $HighPrivAccounts = Get-HighPrivServiceAccounts
                if ($HighPrivAccounts -and $HighPrivAccounts.Count -gt 0) {
                    Update-ADReportResults -Results $HighPrivAccounts
                    Write-ADReportLog -Message "High privileged service accounts loaded. $($HighPrivAccounts.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $HighPrivAccounts
                    $Global:TextBlockStatus.Text = "High privileged service accounts loaded. $($HighPrivAccounts.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No high privileged service accounts found." -Type Info
                    $Global:TextBlockStatus.Text = "No high privileged service accounts found."
                }
            } catch {
                $ErrorMessage = "Error loading high privileged service accounts: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickServiceAccountPasswordAge
    if ($null -ne $Global:ButtonQuickServiceAccountPasswordAge) {
        $ButtonQuickServiceAccountPasswordAge.add_Click({
            Write-ADReportLog -Message "Analyzing service account password ages..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $PasswordAgeAccounts = Get-ServiceAccountPasswordAge
                if ($PasswordAgeAccounts -and $PasswordAgeAccounts.Count -gt 0) {
                    Update-ADReportResults -Results $PasswordAgeAccounts
                    Write-ADReportLog -Message "Service account password ages loaded. $($PasswordAgeAccounts.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $PasswordAgeAccounts
                    $Global:TextBlockStatus.Text = "Service account password ages loaded. $($PasswordAgeAccounts.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No service accounts found for password age analysis." -Type Info
                    $Global:TextBlockStatus.Text = "No service accounts found for password age analysis."
                }
            } catch {
                $ErrorMessage = "Error analyzing service account password ages: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickUnusedServiceAccounts
    if ($null -ne $Global:ButtonQuickUnusedServiceAccounts) {
        $ButtonQuickUnusedServiceAccounts.add_Click({
            Write-ADReportLog -Message "Analyzing unused service accounts..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $UnusedAccounts = Get-UnusedServiceAccounts
                if ($UnusedAccounts -and $UnusedAccounts.Count -gt 0) {
                    Update-ADReportResults -Results $UnusedAccounts
                    Write-ADReportLog -Message "Unused service accounts loaded. $($UnusedAccounts.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $UnusedAccounts
                    $Global:TextBlockStatus.Text = "Unused service accounts loaded. $($UnusedAccounts.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No unused service accounts found." -Type Info
                    $Global:TextBlockStatus.Text = "No unused service accounts found."
                }
            } catch {
                $ErrorMessage = "Error analyzing unused service accounts: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickUnlinkedGPOs
    if ($null -ne $Global:ButtonQuickUnlinkedGPOs) {
        $ButtonQuickUnlinkedGPOs.add_Click({
            Write-ADReportLog -Message "Analyzing unlinked GPOs..." -Type Info
            try {
                $UnlinkedGPOs = Get-UnlinkedGPOs
                if ($UnlinkedGPOs -and $UnlinkedGPOs.Count -gt 0) {
                    Update-ADReportResults -Results $UnlinkedGPOs
                    Write-ADReportLog -Message "Unlinked GPOs loaded. $($UnlinkedGPOs.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $UnlinkedGPOs
                    $Global:TextBlockStatus.Text = "Unlinked GPOs loaded. $($UnlinkedGPOs.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No unlinked GPOs found." -Type Info
                    $Global:TextBlockStatus.Text = "No unlinked GPOs found."
                }
            } catch {
                $ErrorMessage = "Error analyzing unlinked GPOs: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickEmptyGPOs
    if ($null -ne $Global:ButtonQuickEmptyGPOs) {
        $ButtonQuickEmptyGPOs.add_Click({
            Write-ADReportLog -Message "Analyzing empty GPOs..." -Type Info
            try {
                $EmptyGPOs = Get-EmptyGPOs
                if ($EmptyGPOs -and $EmptyGPOs.Count -gt 0) {
                    Update-ADReportResults -Results $EmptyGPOs
                    Write-ADReportLog -Message "Empty GPOs loaded. $($EmptyGPOs.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $EmptyGPOs
                    $Global:TextBlockStatus.Text = "Empty GPOs loaded. $($EmptyGPOs.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No empty GPOs found." -Type Info
                    $Global:TextBlockStatus.Text = "No empty GPOs found."
                }
            } catch {
                $ErrorMessage = "Error analyzing empty GPOs: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickGPOPermissions
    if ($null -ne $Global:ButtonQuickGPOPermissions) {
        $ButtonQuickGPOPermissions.add_Click({
            Write-ADReportLog -Message "Analyzing GPO permissions..." -Type Info
            try {
                $GPOPermissions = Get-GPOPermissions
                if ($GPOPermissions -and $GPOPermissions.Count -gt 0) {
                    Update-ADReportResults -Results $GPOPermissions
                    Write-ADReportLog -Message "GPO permissions loaded. $($GPOPermissions.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $GPOPermissions
                    $Global:TextBlockStatus.Text = "GPO permissions loaded. $($GPOPermissions.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No GPO permissions found." -Type Info
                    $Global:TextBlockStatus.Text = "No GPO permissions found."
                }
            } catch {
                $ErrorMessage = "Error analyzing GPO permissions: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickPasswordPolicySummary
    if ($null -ne $Global:ButtonQuickPasswordPolicySummary) {
        $ButtonQuickPasswordPolicySummary.add_Click({
            Write-ADReportLog -Message "Loading password policy summary..." -Type Info
            try {
                $PasswordPolicies = Get-PasswordPolicySummary
                if ($PasswordPolicies -and $PasswordPolicies.Count -gt 0) {
                    Update-ADReportResults -Results $PasswordPolicies
                    Write-ADReportLog -Message "Password policies loaded. $($PasswordPolicies.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $PasswordPolicies
                    $Global:TextBlockStatus.Text = "Password policies loaded. $($PasswordPolicies.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No password policies found." -Type Info
                    $Global:TextBlockStatus.Text = "No password policies found."
                }
            } catch {
                $ErrorMessage = "Error loading password policies: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickAccountLockoutPolicies
    if ($null -ne $Global:ButtonQuickAccountLockoutPolicies) {
        $ButtonQuickAccountLockoutPolicies.add_Click({
            Write-ADReportLog -Message "Loading account lockout policies..." -Type Info
            try {
                $LockoutPolicies = Get-AccountLockoutPolicies
                if ($LockoutPolicies -and $LockoutPolicies.Count -gt 0) {
                    Update-ADReportResults -Results $LockoutPolicies
                    Write-ADReportLog -Message "Account lockout policies loaded. $($LockoutPolicies.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $LockoutPolicies
                    $Global:TextBlockStatus.Text = "Account lockout policies loaded. $($LockoutPolicies.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No account lockout policies found." -Type Info
                    $Global:TextBlockStatus.Text = "No account lockout policies found."
                }
            } catch {
                $ErrorMessage = "Error loading account lockout policies: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickFineGrainedPasswordPolicies
    if ($null -ne $Global:ButtonQuickFineGrainedPasswordPolicies) {
        $ButtonQuickFineGrainedPasswordPolicies.add_Click({
            Write-ADReportLog -Message "Loading fine-grained password policies..." -Type Info
            try {
                $FGPPs = Get-FineGrainedPasswordPolicies
                if ($FGPPs -and $FGPPs.Count -gt 0) {
                    Update-ADReportResults -Results $FGPPs
                    Write-ADReportLog -Message "Fine-grained password policies loaded. $($FGPPs.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $FGPPs
                    $Global:TextBlockStatus.Text = "Fine-grained password policies loaded. $($FGPPs.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No fine-grained password policies found." -Type Info
                    $Global:TextBlockStatus.Text = "No fine-grained password policies found."
                }
            } catch {
                $ErrorMessage = "Error loading fine-grained password policies: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickACLAnalysis
    if ($null -ne $Global:ButtonQuickACLAnalysis) {
        $ButtonQuickACLAnalysis.add_Click({
            Write-ADReportLog -Message "Analyzing ACL permissions..." -Type Info
            try {
                $ACLAnalysis = Get-ACLAnalysis
                if ($ACLAnalysis -and $ACLAnalysis.Count -gt 0) {
                    Update-ADReportResults -Results $ACLAnalysis
                    Write-ADReportLog -Message "ACL analysis completed. $($ACLAnalysis.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $ACLAnalysis
                    $Global:TextBlockStatus.Text = "ACL analysis completed. $($ACLAnalysis.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No ACL issues found." -Type Info
                    $Global:TextBlockStatus.Text = "No ACL issues found."
                }
            } catch {
                $ErrorMessage = "Error analyzing ACLs: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickInheritanceBreaks 
    if ($null -ne $Global:ButtonQuickInheritanceBreaks) {
        $ButtonQuickInheritanceBreaks.add_Click({
            Write-ADReportLog -Message "Analyzing inheritance breaks..." -Type Info
            try {
                $InheritanceBreaks = Get-InheritanceBreaks
                if ($InheritanceBreaks -and $InheritanceBreaks.Count -gt 0) {
                    Update-ADReportResults -Results $InheritanceBreaks
                    Write-ADReportLog -Message "Inheritance break analysis completed. $($InheritanceBreaks.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $InheritanceBreaks
                    $Global:TextBlockStatus.Text = "Inheritance break analysis completed. $($InheritanceBreaks.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No inheritance breaks found." -Type Info
                    $Global:TextBlockStatus.Text = "No inheritance breaks found."
                }
            } catch {
                $ErrorMessage = "Error analyzing inheritance breaks: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickAdminSDHolderObjects
    if ($null -ne $Global:ButtonQuickAdminSDHolderObjects) {
        $ButtonQuickAdminSDHolderObjects.add_Click({
            Write-ADReportLog -Message "Analyzing AdminSDHolder objects..." -Type Info
            try {
                $AdminSDHolderObjects = Get-AdminSDHolderObjects
                if ($AdminSDHolderObjects -and $AdminSDHolderObjects.Count -gt 0) {
                    Update-ADReportResults -Results $AdminSDHolderObjects
                    Write-ADReportLog -Message "AdminSDHolder analysis completed. $($AdminSDHolderObjects.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $AdminSDHolderObjects
                    $Global:TextBlockStatus.Text = "AdminSDHolder analysis completed. $($AdminSDHolderObjects.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No AdminSDHolder objects found." -Type Info
                    $Global:TextBlockStatus.Text = "No AdminSDHolder objects found."
                }
            } catch {
                $ErrorMessage = "Error analyzing AdminSDHolder objects: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickAdvancedDelegation
    if ($null -ne $Global:ButtonQuickAdvancedDelegation) {
        $ButtonQuickAdvancedDelegation.add_Click({
            Write-ADReportLog -Message "Analyzing advanced delegations..." -Type Info
            try {
                $AdvancedDelegations = Get-AdvancedDelegations
                if ($AdvancedDelegations -and $AdvancedDelegations.Count -gt 0) {
                    Update-ADReportResults -Results $AdvancedDelegations
                    Write-ADReportLog -Message "Advanced delegation analysis completed. $($AdvancedDelegations.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $AdvancedDelegations
                    $Global:TextBlockStatus.Text = "Advanced delegation analysis completed. $($AdvancedDelegations.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No advanced delegations found." -Type Info
                    $Global:TextBlockStatus.Text = "No advanced delegations found."
                }
            } catch {
                $ErrorMessage = "Error analyzing advanced delegations: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickSchemaPermissions
    if ($null -ne $Global:ButtonQuickSchemaPermissions) {
        $ButtonQuickSchemaPermissions.add_Click({
            Write-ADReportLog -Message "Analyzing schema permissions..." -Type Info
            try {
                $SchemaPermissions = Get-SchemaPermissions
                if ($SchemaPermissions -and $SchemaPermissions.Count -gt 0) {
                    Update-ADReportResults -Results $SchemaPermissions
                    Write-ADReportLog -Message "Schema permissions analysis completed. $($SchemaPermissions.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $SchemaPermissions
                    $Global:TextBlockStatus.Text = "Schema permissions analysis completed. $($SchemaPermissions.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No schema permissions found." -Type Info
                    $Global:TextBlockStatus.Text = "No schema permissions found."
                }
            } catch {
                $ErrorMessage = "Error analyzing schema permissions: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler fÃ¼r ButtonQuickGPOOverview
    if ($null -ne $Global:ButtonQuickGPOOverview) {
        $ButtonQuickGPOOverview.add_Click({
            Write-ADReportLog -Message "Loading GPO overview..." -Type Info
            try {
                $GPOOverview = Get-GPOOverview
                if ($GPOOverview -and $GPOOverview.Count -gt 0) {
                    Update-ADReportResults -Results $GPOOverview
                    Write-ADReportLog -Message "GPO overview loaded. $($GPOOverview.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $GPOOverview
                    $Global:TextBlockStatus.Text = "GPO overview loaded. $($GPOOverview.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No GPOs found." -Type Info
                    $Global:TextBlockStatus.Text = "No GPOs found."
                }
            } catch {
                $ErrorMessage = "Error loading GPO overview: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler fÃ¼r ButtonQuickHoneyTokens
    if ($null -ne $Global:ButtonQuickHoneyTokens) {
        $ButtonQuickHoneyTokens.add_Click({
            Write-ADReportLog -Message "Analyzing potential honey tokens..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $HoneyTokens = Get-HoneyTokens
                if ($HoneyTokens -and $HoneyTokens.Count -gt 0) {
                    Update-ADReportResults -Results $HoneyTokens
                    Write-ADReportLog -Message "Honey token analysis completed. $($HoneyTokens.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $HoneyTokens
                    $Global:TextBlockStatus.Text = "Honey token analysis completed. $($HoneyTokens.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No honey tokens detected." -Type Info
                    $Global:TextBlockStatus.Text = "No honey tokens detected."
                }
            } catch {
                $ErrorMessage = "Error analyzing honey tokens: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickPrivilegeEscalation
    if ($null -ne $Global:ButtonQuickPrivilegeEscalation) {
        $ButtonQuickPrivilegeEscalation.add_Click({
            Write-ADReportLog -Message "Analyzing potential privilege escalation paths..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $PrivEscPaths = Get-PrivilegeEscalationPaths
                if ($PrivEscPaths -and $PrivEscPaths.Count -gt 0) {
                    Update-ADReportResults -Results $PrivEscPaths
                    Write-ADReportLog -Message "Privilege escalation analysis completed. $($PrivEscPaths.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $PrivEscPaths
                    $Global:TextBlockStatus.Text = "Privilege escalation analysis completed. $($PrivEscPaths.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No privilege escalation paths detected." -Type Info
                    $Global:TextBlockStatus.Text = "No privilege escalation paths detected."
                }
            } catch {
                $ErrorMessage = "Error analyzing privilege escalation paths: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickExposedCredentials
    if ($null -ne $Global:ButtonQuickExposedCredentials) {
        $ButtonQuickExposedCredentials.add_Click({
            Write-ADReportLog -Message "Analyzing exposed credentials..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $ExposedCreds = Get-ExposedCredentials
                if ($ExposedCreds -and $ExposedCreds.Count -gt 0) {
                    Update-ADReportResults -Results $ExposedCreds
                    Write-ADReportLog -Message "Exposed credentials analysis completed. $($ExposedCreds.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $ExposedCreds
                    $Global:TextBlockStatus.Text = "Exposed credentials analysis completed. $($ExposedCreds.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No exposed credentials detected." -Type Info
                    $Global:TextBlockStatus.Text = "No exposed credentials detected."
                }
            } catch {
                $ErrorMessage = "Error analyzing exposed credentials: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickSuspiciousLogons
    if ($null -ne $Global:ButtonQuickSuspiciousLogons) {
        $ButtonQuickSuspiciousLogons.add_Click({
            Write-ADReportLog -Message "Analyzing suspicious logon patterns..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $SuspiciousLogons = Get-SuspiciousLogons
                if ($SuspiciousLogons -and $SuspiciousLogons.Count -gt 0) {
                    Update-ADReportResults -Results $SuspiciousLogons
                    Write-ADReportLog -Message "Suspicious logon analysis completed. $($SuspiciousLogons.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $SuspiciousLogons
                    $Global:TextBlockStatus.Text = "Suspicious logon analysis completed. $($SuspiciousLogons.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No suspicious logon patterns detected." -Type Info
                    $Global:TextBlockStatus.Text = "No suspicious logon patterns detected."
                }
            } catch {
                $ErrorMessage = "Error analyzing suspicious logons: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickForeignSecurityPrincipals
    if ($null -ne $Global:ButtonQuickForeignSecurityPrincipals) {
        $ButtonQuickForeignSecurityPrincipals.add_Click({
            Write-ADReportLog -Message "Analyzing foreign security principals..." -Type Info
            try {
                $ForeignPrincipals = Get-ForeignSecurityPrincipals
                if ($ForeignPrincipals -and $ForeignPrincipals.Count -gt 0) {
                    Update-ADReportResults -Results $ForeignPrincipals
                    Write-ADReportLog -Message "Foreign security principals analysis completed. $($ForeignPrincipals.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $ForeignPrincipals
                    $Global:TextBlockStatus.Text = "Foreign security principals analysis completed. $($ForeignPrincipals.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No foreign security principals found." -Type Info
                    $Global:TextBlockStatus.Text = "No foreign security principals found."
                }
            } catch {
                $ErrorMessage = "Error analyzing foreign security principals: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Event Handler für ButtonQuickSIDHistoryAbuse
    if ($null -ne $Global:ButtonQuickSIDHistoryAbuse) {
        $ButtonQuickSIDHistoryAbuse.add_Click({
            Write-ADReportLog -Message "Analyzing potential SID history abuse..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $SIDHistoryAbuse = Get-SIDHistoryAbuse
                if ($SIDHistoryAbuse -and $SIDHistoryAbuse.Count -gt 0) {
                    Update-ADReportResults -Results $SIDHistoryAbuse
                    Write-ADReportLog -Message "SID history abuse analysis completed. $($SIDHistoryAbuse.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $SIDHistoryAbuse
                    $Global:TextBlockStatus.Text = "SID history abuse analysis completed. $($SIDHistoryAbuse.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No potential SID history abuse detected." -Type Info
                    $Global:TextBlockStatus.Text = "No potential SID history abuse detected."
                }
            } catch {
                $ErrorMessage = "Error analyzing SID history abuse: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }
    
    # Event Handler fÃ¼r weitere Quick Report Buttons
    if ($null -ne $Global:ButtonQuickExpiringAccounts) {
        $ButtonQuickExpiringAccounts.add_Click({
            Write-ADReportLog -Message "Analyzing expiring accounts..." -Type Info
            try {
                $Global:RadioButtonUser.IsChecked = $true
                
                $ExpiringAccounts = Get-ExpiringAccounts
                if ($ExpiringAccounts -and $ExpiringAccounts.Count -gt 0) {
                    Update-ADReportResults -Results $ExpiringAccounts
                    Write-ADReportLog -Message "Expiring accounts analysis completed. $($ExpiringAccounts.Count) result(s) found." -Type Info
                    Update-ResultCounters -Results $ExpiringAccounts
                    $Global:TextBlockStatus.Text = "Expiring accounts analysis completed. $($ExpiringAccounts.Count) record(s) found."
                } else {
                    $Global:DataGridResults.ItemsSource = $null
                    Write-ADReportLog -Message "No expiring accounts found." -Type Info
                    $Global:TextBlockStatus.Text = "No expiring accounts found."
                }
            } catch {
                $ErrorMessage = "Error analyzing expiring accounts: $($_.Exception.Message)"
                Write-ADReportLog -Message $ErrorMessage -Type Error
                Update-ADReportResults -Results @()
                $Global:TextBlockStatus.Text = "Error: $ErrorMessage"
            }
        })
    }

    # Definiere den ListBox Selection Changed Handler als Script Variable
    $script:ListBoxSelectionChangedHandler = $null
    
    # Fenster anzeigen
    $null = $Window.ShowDialog()
}

# Initialize Visualization Tabs wenn vorhanden
if ($null -ne $Global:ButtonRefreshTopology) {
    $ButtonRefreshTopology.Add_Click({
        $viewType = $Global:ComboBoxTopologyView.SelectedItem.Content
        $topology = Get-ADNetworkTopology -ViewType $viewType
        Draw-TopologyOnCanvas -Canvas $Global:CanvasNetworkTopology -Topology $topology
    })
}

if ($null -ne $Global:ButtonRefreshHeatMap) {
    $ButtonRefreshHeatMap.Add_Click({
        $metric = $Global:ComboBoxHeatMapMetric.SelectedItem.Content -replace ' Score', ''
        $grouping = $Global:ComboBoxHeatMapGrouping.SelectedItem.Content -replace 'By ', ''
        $heatMapData = Get-SecurityHeatMapData -MetricType $metric -GroupBy $grouping
        Draw-SecurityHeatMap -Container $Global:HeatMapContainer -HeatMapData $heatMapData
    })
}

if ($null -ne $Global:ListBoxAvailableFields) {
    Initialize-ReportBuilder -AvailableFieldsList $Global:ListBoxAvailableFields `
                            -ReportCanvas $Global:ReportBuilderCanvas
}

if ($null -ne $Global:ButtonPreviewReport) {
    $ButtonPreviewReport.Add_Click({
        $customReport = Build-CustomReport -ReportCanvas $Global:ReportBuilderCanvas
        Update-ADReportResults -Results $customReport
        $Global:MainTabControl.SelectedIndex = 0  # Wechsel zur Tabellenansicht
    })
}

# Funktion zur Visualisierung der Ergebnisse (Placeholder fÃ¼r zukÃ¼nftige Implementierung)
Function Update-ResultVisualization {
    param (
        [Parameter(Mandatory=$false)]
        $Results
    )
    
    # Diese Funktion ist ein Placeholder fÃ¼r zukÃ¼nftige Visualisierungen
    # Momentan wird nur ein Debug-Log-Eintrag erstellt
    Write-DebugLog "Update-ResultVisualization aufgerufen mit $($Results.Count) Ergebnissen" "Visualization"
}

# Platzhalter-Funktion fÃ¼r initiale Visualisierung beim Start
Function Initialize-ResultVisualization {
    [CmdletBinding()]
    param()
    
    # Diese Funktion ist ein Placeholder fÃ¼r zukÃ¼nftige Visualisierungen
    Write-DebugLog "Initialize-ResultVisualization aufgerufen" "Visualization"
}

# --- Hauptlogik ---
Function Start-ADReportGUI {
    # Bereinige alle alten globalen UI-Variablen vor dem Start
    $UiVariables = @("Window", "ComboBoxFilterAttribute", "TextBoxFilterValue", "ListBoxSelectAttributes",
                  "ButtonQueryAD", "ButtonQuickAllUsers", "ButtonQuickDisabledUsers", "ButtonQuickLockedUsers",
                  "ButtonQuickGroups", "ButtonQuickSecurityGroups", "ButtonQuickDistributionGroups",
                  "ButtonQuickNeverExpire", "ButtonQuickInactiveUsers", "ButtonQuickAdminUsers",
                  "ButtonQuickComputers", "ButtonQuickInactiveComputers", 
                  "ButtonQuickWeakPasswordPolicy", "ButtonQuickRiskyGroupMemberships", "ButtonQuickPrivilegedAccounts",
                  "DataGridResults", "TextBlockStatus", "ButtonExportCSV", "ButtonExportHTML",
                  "ResultCountGrid", "UserCountText", "ComputerCountText", "GroupCountText")
    
    foreach ($var in $UiVariables) {
        Remove-Variable -Name $var -Scope Global -ErrorAction SilentlyContinue
    }

    # Ruft Initialize-ADReportForm auf, welche die UI lÃ¤dt, Elemente zuweist und fÃ¼llt.
    Initialize-ADReportForm -XAMLContent $Global:XAML
}

# --- Skriptstart ---
Start-ADReportGUI