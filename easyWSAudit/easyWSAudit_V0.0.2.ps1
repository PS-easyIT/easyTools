# easyWSAudit - Windows Server Audit Tool
# Version: 0.0.1 - Vereinfacht ohne asynchrone Ausfuehrung

# Debug-Modus - Auf $true setzen fuer ausfuehrliche Logging-Informationen
$DEBUG = $false
$DebugLogPath = "$env:TEMP\easyWSAudit_Debug.log"

# Debug-Funktion
function Write-DebugLog {
    param(
        [string]$Message,
        [string]$Source = "General"
    )
    
    if ($DEBUG) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        $logMessage = "[$timestamp] [$Source] $Message"
        
        # Ausgabe in die Konsole
        Write-Host $logMessage -ForegroundColor Yellow
        
        # Ausgabe in die Log-Datei
        Add-Content -Path $DebugLogPath -Value $logMessage
        
        # Wenn das Debug-Display vorhanden ist, dort auch anzeigen
        if ($script:txtDebugOutput -and $window) {
            try {
                $script:txtDebugOutput.Dispatcher.Invoke([Action]{
                    $script:txtDebugOutput.Text += "$logMessage`r`n"
                    $script:txtDebugOutput.ScrollToEnd()
                }, "Normal")
            } catch {
                # Ignoriere Dispatcher-Fehler
            }
        }
    }
}

# Starte das Debug-Log
if ($DEBUG) {
    $startMessage = "=== easyWSAudit Debug Log gestartet $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") ==="
    Set-Content -Path $DebugLogPath -Value $startMessage -Force
    Write-Host $startMessage -ForegroundColor Cyan
}

# Importiere notwendige Module
Write-DebugLog "Importiere notwendige Module..." "Init"
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms # Fuer den SaveFileDialog
Write-DebugLog "Module importiert" "Init"

# Definiere die Commands fuer verschiedene Server-Rollen und Systeminformationen
$commands = @(
    # === SYSTEM INFORMATIONEN ===
    @{Name="Systeminformationen"; Command="Get-ComputerInfo"; Type="PowerShell"; Category="System"},
    @{Name="Betriebssystem Details"; Command="Get-CimInstance Win32_OperatingSystem"; Type="PowerShell"; Category="System"},
    @{Name="Hardware Informationen"; Command="Get-CimInstance Win32_ComputerSystem"; Type="PowerShell"; Category="Hardware"},
    @{Name="CPU Informationen"; Command="Get-CimInstance Win32_Processor"; Type="PowerShell"; Category="Hardware"},
    @{Name="Arbeitsspeicher Details"; Command="Get-CimInstance Win32_PhysicalMemory"; Type="PowerShell"; Category="Hardware"},
    @{Name="Festplatten Informationen"; Command="Get-CimInstance Win32_LogicalDisk"; Type="PowerShell"; Category="Storage"},
    @{Name="Volume Informationen"; Command="Get-Volume"; Type="PowerShell"; Category="Storage"},
    @{Name="Installierte Features und Rollen"; Command="Get-WindowsFeature | Where-Object { `$_.Installed -eq `$true }"; Type="PowerShell"; Category="Features"},
    @{Name="Installierte Programme"; Command="Get-CimInstance Win32_Product | Select-Object Name, Version, Vendor | Sort-Object Name"; Type="PowerShell"; Category="Software"},
    @{Name="Windows Updates"; Command="Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 20"; Type="PowerShell"; Category="Updates"},
    @{Name="Netzwerkkonfiguration"; Command="Get-NetIPConfiguration"; Type="PowerShell"; Category="Network"},
    @{Name="Netzwerkadapter"; Command="Get-NetAdapter"; Type="PowerShell"; Category="Network"},
    @{Name="Aktive Netzwerkverbindungen"; Command="Get-NetTCPConnection | Where-Object State -eq 'Listen' | Select-Object LocalAddress, LocalPort, OwningProcess"; Type="PowerShell"; Category="Network"},
    @{Name="Firewall Regeln"; Command="Get-NetFirewallRule | Where-Object Enabled -eq 'True' | Select-Object DisplayName, Direction, Action | Sort-Object DisplayName"; Type="PowerShell"; Category="Security"},
    @{Name="Services (Automatisch)"; Command="Get-Service | Where-Object StartType -eq 'Automatic' | Sort-Object Status, Name"; Type="PowerShell"; Category="Services"},
    @{Name="Services (Laufend)"; Command="Get-Service | Where-Object Status -eq 'Running' | Sort-Object Name"; Type="PowerShell"; Category="Services"},
    @{Name="Geplante Aufgaben"; Command="Get-ScheduledTask | Where-Object State -eq 'Ready' | Select-Object TaskName, TaskPath, State"; Type="PowerShell"; Category="Tasks"},
    @{Name="Event-Log System (Letzte 24h)"; Command="Get-WinEvent -FilterHashtable @{LogName='System'; StartTime=(Get-Date).AddDays(-1)} -MaxEvents 50 | Select-Object TimeCreated, Id, LevelDisplayName, Message"; Type="PowerShell"; Category="Events"},
    @{Name="Event-Log Application (Letzte 24h)"; Command="Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=(Get-Date).AddDays(-1)} -MaxEvents 50 | Select-Object TimeCreated, Id, LevelDisplayName, Message"; Type="PowerShell"; Category="Events"},
    @{Name="Lokale Benutzer"; Command="Get-LocalUser | Select-Object Name, Enabled, LastLogon, PasswordRequired"; Type="PowerShell"; Category="Security"},
    @{Name="Lokale Gruppen"; Command="Get-LocalGroup | Select-Object Name, Description"; Type="PowerShell"; Category="Security"},
    
    # === ACTIVE DIRECTORY UMFASSENDE AUDITS ===
    @{Name="AD Domain Controller Status"; Command="Get-ADDomainController -Filter * | Select-Object Name, Site, IPv4Address, OperatingSystem, IsGlobalCatalog, IsReadOnly"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Domain Informationen"; Command="Get-ADDomain | Select-Object Name, NetBIOSName, DomainMode, PDCEmulator, RIDMaster, InfrastructureMaster"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Forest Informationen"; Command="Get-ADForest | Select-Object Name, ForestMode, DomainNamingMaster, SchemaMaster, Sites, Domains"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Organizational Units"; Command="Get-ADOrganizationalUnit -Filter * | Select-Object Name, DistinguishedName, Description | Sort-Object Name"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Domaenen Administratoren"; Command="Get-ADGroupMember -Identity 'Domain Admins' | Get-ADUser -Properties LastLogonDate, PasswordLastSet, Enabled | Select-Object Name, SamAccountName, Enabled, LastLogonDate, PasswordLastSet"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Enterprise Administratoren"; Command="Get-ADGroupMember -Identity 'Enterprise Admins' | Get-ADUser -Properties LastLogonDate, PasswordLastSet, Enabled | Select-Object Name, SamAccountName, Enabled, LastLogonDate, PasswordLastSet"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Schema Administratoren"; Command="Get-ADGroupMember -Identity 'Schema Admins' | Get-ADUser -Properties LastLogonDate, PasswordLastSet, Enabled | Select-Object Name, SamAccountName, Enabled, LastLogonDate, PasswordLastSet"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Privilegierte Gruppen"; Command="@('Domain Admins','Enterprise Admins','Schema Admins','Administrators','Account Operators','Backup Operators','Print Operators','Server Operators') | ForEach-Object { Get-ADGroup -Identity `$_ -Properties Members | Select-Object Name, @{Name='MemberCount';Expression={`$_.Members.Count}} }"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Benutzer ohne Passwort-Ablauf"; Command="Get-ADUser -Filter {PasswordNeverExpires -eq `$true -and Enabled -eq `$true} -Properties PasswordNeverExpires, LastLogonDate | Select-Object Name, SamAccountName, LastLogonDate"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Deaktivierte Benutzer"; Command="Get-ADUser -Filter {Enabled -eq `$false} | Select-Object Name, SamAccountName, DistinguishedName | Sort-Object Name"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Computer Accounts"; Command="Get-ADComputer -Filter * -Properties OperatingSystem, LastLogonDate | Select-Object Name, OperatingSystem, LastLogonDate, Enabled | Sort-Object LastLogonDate -Descending"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Replikations-Status"; Command="repadmin /replsummary"; Type="CMD"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD FSMO Rollen"; Command="Get-ADDomain | Select-Object PDCEmulator, RIDMaster, InfrastructureMaster; Get-ADForest | Select-Object DomainNamingMaster, SchemaMaster"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Sites und Subnets"; Command="Get-ADReplicationSite | Select-Object Name, Description; Get-ADReplicationSubnet -Filter * | Select-Object Name, Site, Location"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Trust Relationships"; Command="Get-ADTrust -Filter * | Select-Object Name, Direction, TrustType, DisallowTransivity"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    
    # === DNS UMFASSENDE DIAGNOSEN ===
    @{Name="DNS Server Konfiguration"; Command="Get-DnsServer | Select-Object ComputerName, ZoneScavenging, EnableDnsSec, ServerSetting"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},
    @{Name="DNS Server Zonen"; Command="Get-DnsServerZone | Select-Object ZoneName, ZoneType, IsAutoCreated, IsDsIntegrated, IsReverseLookupZone"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},
    @{Name="DNS Forwarders"; Command="Get-DnsServerForwarder"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},
    @{Name="DNS Cache-Inhalt"; Command="Get-DnsServerCache | Select-Object -First 20"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},
    @{Name="DNS Scavenging Einstellungen"; Command="Get-DnsServerScavenging"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},
    @{Name="DNS Event-Log Errors"; Command="Get-WinEvent -FilterHashtable @{LogName='DNS Server'; Level=2,3} -MaxEvents 50 | Select-Object TimeCreated, Id, LevelDisplayName, Message"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},
    @{Name="DNS Root Hints"; Command="Get-DnsServerRootHint"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},
    @{Name="DNS Service Dependencies"; Command="`$Services='DNS','Netlogon','Kerberos Key Distribution Center'; ForEach (`$Service in `$Services) {Get-Service `$Service | Select-Object Name, Status, StartType}"; Type="PowerShell"; FeatureName="DNS"; Category="DNS"},
    
    # === DHCP UMFASSENDE AUDITS ===
    @{Name="DHCP Server Konfiguration"; Command="Get-DhcpServerInDC"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP Server Settings"; Command="Get-DhcpServerSetting"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP IPv4 Bereiche"; Command="Get-DhcpServerv4Scope | Select-Object ScopeId, Name, StartRange, EndRange, SubnetMask, State, LeaseDuration"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP IPv6 Bereiche"; Command="Get-DhcpServerv6Scope | Select-Object Prefix, Name, State, PreferredLifetime"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP Reservierungen"; Command="Get-DhcpServerv4Reservation | Select-Object ScopeId, IPAddress, ClientId, Name, Description"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP Lease-Statistiken"; Command="Get-DhcpServerv4ScopeStatistics | Select-Object ScopeId, AddressesFree, AddressesInUse, PercentageInUse"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP Optionen (Server)"; Command="Get-DhcpServerv4OptionValue | Select-Object OptionId, Name, Value"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP Failover Konfiguration"; Command="Get-DhcpServerv4Failover | Select-Object Name, PartnerServer, Mode, State"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    @{Name="DHCP Event-Log"; Command="Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Dhcp-Server/Operational'} -MaxEvents 50 | Select-Object TimeCreated, Id, LevelDisplayName, Message"; Type="PowerShell"; FeatureName="DHCP"; Category="DHCP"},
    
    # === IIS UMFASSENDE AUDITS ===
    @{Name="IIS Server Informationen"; Command="Get-IISServerManager | Select-Object *"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    @{Name="IIS Websites"; Command="Get-IISSite | Select-Object Name, Id, State, PhysicalPath, @{Name='Bindings';Expression={`$_.Bindings | ForEach-Object {`$_.Protocol + '://' + `$_.BindingInformation}}}"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    @{Name="IIS Application Pools"; Command="Get-IISAppPool | Select-Object Name, State, ProcessModel, Recycling"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    @{Name="IIS Anwendungen"; Command="Get-WebApplication | Select-Object Site, Path, PhysicalPath, ApplicationPool"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    @{Name="IIS Virtuelle Verzeichnisse"; Command="Get-WebVirtualDirectory | Select-Object Site, Application, Path, PhysicalPath"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    @{Name="IIS SSL Zertifikate"; Command="Get-ChildItem IIS:SslBindings | Select-Object IPAddress, Port, Host, Thumbprint, Subject"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    @{Name="IIS Modules"; Command="Get-WebGlobalModule | Select-Object Name, Image"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    @{Name="IIS Handler Mappings"; Command="Get-WebHandler | Select-Object Name, Path, Verb, Modules"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    @{Name="IIS Event-Log"; Command="Get-WinEvent -FilterHashtable @{LogName='System'; ProviderName='Microsoft-Windows-IIS*'} -MaxEvents 50 | Select-Object TimeCreated, Id, LevelDisplayName, Message"; Type="PowerShell"; FeatureName="Web-Server"; Category="IIS"},
    
    # === WDS (WINDOWS DEPLOYMENT SERVICES) ===
    @{Name="WDS Server Konfiguration"; Command="wdsutil /get-server /show:config"; Type="CMD"; FeatureName="WDS"; Category="WDS"},
    @{Name="WDS Boot Images"; Command="wdsutil /get-allimages /show:install"; Type="CMD"; FeatureName="WDS"; Category="WDS"},
    @{Name="WDS Install Images"; Command="wdsutil /get-allimages /show:boot"; Type="CMD"; FeatureName="WDS"; Category="WDS"},
    @{Name="WDS Transport Server"; Command="wdsutil /get-transportserver /show:config"; Type="CMD"; FeatureName="WDS"; Category="WDS"},
    @{Name="WDS Multicast"; Command="wdsutil /get-allmulticasttransmissions"; Type="CMD"; FeatureName="WDS"; Category="WDS"},
    @{Name="WDS Client Einstellungen"; Command="Get-WdsClient | Select-Object DeviceID, DeviceName, BootImagePath, ReferralServer"; Type="PowerShell"; FeatureName="WDS"; Category="WDS"},
    
    # === HYPER-V AUDITS ===
    @{Name="Hyper-V Host Informationen"; Command="Get-VMHost | Select-Object ComputerName, LogicalProcessorCount, MemoryCapacity, VirtualMachinePath, VirtualHardDiskPath"; Type="PowerShell"; FeatureName="Hyper-V"; Category="Hyper-V"},
    @{Name="Hyper-V Virtuelle Maschinen"; Command="Get-VM | Select-Object Name, State, CPUUsage, MemoryAssigned, Uptime, Version, Generation"; Type="PowerShell"; FeatureName="Hyper-V"; Category="Hyper-V"},
    @{Name="Hyper-V Switches"; Command="Get-VMSwitch | Select-Object Name, SwitchType, NetAdapterInterfaceDescription, AllowManagementOS"; Type="PowerShell"; FeatureName="Hyper-V"; Category="Hyper-V"},
    @{Name="Hyper-V Snapshots"; Command="Get-VMSnapshot * | Select-Object VMName, Name, SnapshotType, CreationTime, ParentSnapshotName"; Type="PowerShell"; FeatureName="Hyper-V"; Category="Hyper-V"},
    @{Name="Hyper-V Integration Services"; Command="Get-VM | Get-VMIntegrationService | Select-Object VMName, Name, Enabled, PrimaryStatusDescription"; Type="PowerShell"; FeatureName="Hyper-V"; Category="Hyper-V"},
    @{Name="Hyper-V Replikation"; Command="Get-VMReplication | Select-Object VMName, State, Mode, FrequencySec, PrimaryServer, ReplicaServer"; Type="PowerShell"; FeatureName="Hyper-V"; Category="Hyper-V"},
    
    # === FAILOVER CLUSTER AUDITS ===
    @{Name="Cluster Informationen"; Command="Get-Cluster | Select-Object Name, Domain, AddEvictDelay, BackupInProgress, BlockCacheSize"; Type="PowerShell"; FeatureName="Failover-Clustering"; Category="Cluster"},
    @{Name="Cluster Nodes"; Command="Get-ClusterNode | Select-Object Name, State, Type, ID"; Type="PowerShell"; FeatureName="Failover-Clustering"; Category="Cluster"},
    @{Name="Cluster Resources"; Command="Get-ClusterResource | Select-Object Name, ResourceType, State, OwnerNode, OwnerGroup"; Type="PowerShell"; FeatureName="Failover-Clustering"; Category="Cluster"},
    @{Name="Cluster Shared Volumes"; Command="Get-ClusterSharedVolume | Select-Object Name, State, Node, SharedVolumeInfo"; Type="PowerShell"; FeatureName="Failover-Clustering"; Category="Cluster"},
    @{Name="Cluster Networks"; Command="Get-ClusterNetwork | Select-Object Name, Role, State, Address, AddressMask"; Type="PowerShell"; FeatureName="Failover-Clustering"; Category="Cluster"},
    @{Name="Cluster Validation Report"; Command="Test-Cluster -ReportName 'C:\\temp\\ClusterValidation.htm' -Include 'Storage','Network','System Configuration','Inventory'"; Type="PowerShell"; FeatureName="Failover-Clustering"; Category="Cluster"},
    
    # === WINDOWS SERVER UPDATE SERVICES (WSUS) ===
    @{Name="WSUS Server Konfiguration"; Command="Get-WsusServer | Select-Object Name, PortNumber, ServerProtocolVersion, DatabasePath"; Type="PowerShell"; FeatureName="UpdateServices"; Category="WSUS"},
    @{Name="WSUS Update Kategorien"; Command="Get-WsusServer | Get-WsusClassification | Select-Object Classification, ID"; Type="PowerShell"; FeatureName="UpdateServices"; Category="WSUS"},
    @{Name="WSUS Computer Gruppen"; Command="Get-WsusServer | Get-WsusComputerTargetGroup | Select-Object Name, ID, ComputerTargets"; Type="PowerShell"; FeatureName="UpdateServices"; Category="WSUS"},
    @{Name="WSUS Synchronisation Status"; Command="Get-WsusServer | Get-WsusSubscription | Select-Object LastSynchronizationTime, SynchronizeAutomatically, NumberOfSynchronizations"; Type="PowerShell"; FeatureName="UpdateServices"; Category="WSUS"},
    
    # === FILE SERVICES ===
    @{Name="File Shares"; Command="Get-SmbShare | Select-Object Name, Path, Description, ShareType, ShareState"; Type="PowerShell"; FeatureName="FS-FileServer"; Category="FileServices"},
    @{Name="DFS Namespaces"; Command="Get-DfsnRoot | Select-Object Path, Type, State, Description"; Type="PowerShell"; FeatureName="FS-DFS-Namespace"; Category="FileServices"},
    @{Name="DFS Replication Groups"; Command="Get-DfsReplicationGroup | Select-Object GroupName, State, Description"; Type="PowerShell"; FeatureName="FS-DFS-Replication"; Category="FileServices"},
    @{Name="File Server Resource Manager Quotas"; Command="Get-FsrmQuota | Select-Object Path, Size, SoftLimit, Usage, Description"; Type="PowerShell"; FeatureName="FS-Resource-Manager"; Category="FileServices"},
    @{Name="Shadow Copies"; Command="Get-CimInstance -ClassName Win32_ShadowCopy | Select-Object VolumeName, InstallDate, Count"; Type="PowerShell"; FeatureName="FS-FileServer"; Category="FileServices"},
    
    # === PRINT SERVICES ===
    @{Name="Print Server Drucker"; Command="Get-Printer | Select-Object Name, DriverName, PortName, Shared, Published, PrinterStatus"; Type="PowerShell"; FeatureName="Print-Services"; Category="PrintServices"},
    @{Name="Print Server Treiber"; Command="Get-PrinterDriver | Select-Object Name, Manufacturer, DriverVersion, PrinterEnvironment"; Type="PowerShell"; FeatureName="Print-Services"; Category="PrintServices"},
    @{Name="Print Server Ports"; Command="Get-PrinterPort | Select-Object Name, Description, PortMonitor, PortType"; Type="PowerShell"; FeatureName="Print-Services"; Category="PrintServices"},
    
    # === REMOTE DESKTOP SERVICES ===
    @{Name="RDS Server Informationen"; Command="Get-RDServer | Select-Object Server, Roles"; Type="PowerShell"; FeatureName="RDS-RD-Server"; Category="RDS"},
    @{Name="RDS Session Collections"; Command="Get-RDSessionCollection | Select-Object CollectionName, CollectionDescription, Size"; Type="PowerShell"; FeatureName="RDS-RD-Server"; Category="RDS"},
    @{Name="RDS User Sessions"; Command="Get-RDUserSession | Select-Object CollectionName, DomainName, UserName, SessionState"; Type="PowerShell"; FeatureName="RDS-RD-Server"; Category="RDS"},
    @{Name="RDS Licensing"; Command="Get-RDLicenseConfiguration | Select-Object Mode, LicenseServer"; Type="PowerShell"; FeatureName="RDS-Licensing"; Category="RDS"},
    
    # === CERTIFICATE SERVICES ===
    @{Name="Certificate Authority Info"; Command="certutil -getconfig"; Type="CMD"; FeatureName="ADCS-Cert-Authority"; Category="PKI"},
    @{Name="CA Certificate Templates"; Command="certutil -template"; Type="CMD"; FeatureName="ADCS-Cert-Authority"; Category="PKI"},
    @{Name="Certificate Store - Personal"; Command="Get-ChildItem Cert:\\LocalMachine\\My | Select-Object Subject, Issuer, NotAfter, Thumbprint"; Type="PowerShell"; Category="PKI"},
    @{Name="Certificate Store - Root"; Command="Get-ChildItem Cert:\\LocalMachine\\Root | Select-Object Subject, Issuer, NotAfter, Thumbprint"; Type="PowerShell"; Category="PKI"},
    
    # === ERWEITERTE SICHERHEITS-AUDITS ===
    @{Name="Security Event Log (Letzte 100)"; Command="Get-WinEvent -FilterHashtable @{LogName='Security'} -MaxEvents 100 | Select-Object TimeCreated, Id, LevelDisplayName, UserId, Message"; Type="PowerShell"; Category="Security"},
    @{Name="Failed Logon Attempts"; Command="Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4625} -MaxEvents 50 | Select-Object TimeCreated, Message"; Type="PowerShell"; Category="Security"},
    @{Name="Account Lockouts"; Command="Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4740} -MaxEvents 50 | Select-Object TimeCreated, Message"; Type="PowerShell"; Category="Security"},
    @{Name="Privilege Use Auditing"; Command="Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4672,4673,4674} -MaxEvents 50 | Select-Object TimeCreated, Id, Message"; Type="PowerShell"; Category="Security"},
    @{Name="Unsecure LDAP Binds"; Command="Get-WinEvent -FilterHashtable @{LogName='Directory Service'; ID=2889} -MaxEvents 20 | Select-Object TimeCreated, Message"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Security"},
    
    # === ACTIVE DIRECTORY FEDERATION SERVICES (ADFS) ===
    @{Name="ADFS Server Konfiguration"; Command="Get-AdfsProperties | Select-Object DisplayName, HostName, HttpPort, HttpsPort, TlsClientPort"; Type="PowerShell"; FeatureName="ADFS-Federation"; Category="ADFS"},
    @{Name="ADFS Relying Party Trusts"; Command="Get-AdfsRelyingPartyTrust | Select-Object Name, Enabled, Identifier, IssuanceAuthorizationRules"; Type="PowerShell"; FeatureName="ADFS-Federation"; Category="ADFS"},
    @{Name="ADFS Claims Provider Trusts"; Command="Get-AdfsClaimsProviderTrust | Select-Object Name, Enabled, Identifier, AcceptanceTransformRules"; Type="PowerShell"; FeatureName="ADFS-Federation"; Category="ADFS"},
    @{Name="ADFS Certificates"; Command="Get-AdfsCertificate | Select-Object CertificateType, Thumbprint, Subject, NotAfter"; Type="PowerShell"; FeatureName="ADFS-Federation"; Category="ADFS"},
    @{Name="ADFS Endpoints"; Command="Get-AdfsEndpoint | Select-Object AddressPath, Enabled, Protocol, SecurityMode"; Type="PowerShell"; FeatureName="ADFS-Federation"; Category="ADFS"},
    
    # === ACTIVE DIRECTORY LIGHTWEIGHT DIRECTORY SERVICES (ADLDS) ===
    @{Name="ADLDS Instances"; Command="Get-CimInstance -ClassName Win32_Service | Where-Object {`$_.Name -like 'ADAM_*'} | Select-Object Name, State, StartMode, PathName"; Type="PowerShell"; FeatureName="ADLDS"; Category="ADLDS"},
    @{Name="ADLDS Configuration"; Command="dsdbutil -c 'activate instance `$instancename' quit | Out-String"; Type="CMD"; FeatureName="ADLDS"; Category="ADLDS"},
    
    # === ACTIVE DIRECTORY RIGHTS MANAGEMENT SERVICES (ADRMS) ===
    @{Name="ADRMS Cluster Info"; Command="Get-RmsCluster | Select-Object ClusterName, ClusterUrl, Version"; Type="PowerShell"; FeatureName="ADRMS"; Category="ADRMS"},
    @{Name="ADRMS Server Info"; Command="Get-RmsServer | Select-Object Name, ClusterName, Version, IsConnected"; Type="PowerShell"; FeatureName="ADRMS"; Category="ADRMS"},
    @{Name="ADRMS Templates"; Command="Get-RmsTemplate | Select-Object Name, Description, Validity, Created"; Type="PowerShell"; FeatureName="ADRMS"; Category="ADRMS"},
    
    # === DEVICE HEALTH ATTESTATION SERVICE ===
    @{Name="Device Health Attestation Service"; Command="Get-DHASActiveEncryptionCertificate; Get-DHASActiveSigningCertificate"; Type="PowerShell"; FeatureName="DeviceHealthAttestationService"; Category="DeviceAttestation"},
    
    # === VOLUME ACTIVATION SERVICES ===
    @{Name="KMS Server Konfiguration"; Command="slmgr /dlv; cscript C:\\Windows\\System32\\slmgr.vbs /dli"; Type="CMD"; FeatureName="VolumeActivation"; Category="VolumeActivation"},
    @{Name="KMS Client Status"; Command="Get-CimInstance -ClassName SoftwareLicensingProduct | Where-Object {`$_.LicenseStatus -eq 1} | Select-Object Name, Description, LicenseStatus"; Type="PowerShell"; FeatureName="VolumeActivation"; Category="VolumeActivation"},
    
    # === WINDOWS SERVER BACKUP ===
    @{Name="Windows Server Backup Policies"; Command="Get-WBPolicy | Select-Object PolicyState, BackupTargets, FilesSpecsToBackup"; Type="PowerShell"; FeatureName="Windows-Server-Backup"; Category="Backup"},
    @{Name="Windows Server Backup Jobs"; Command="Get-WBJob -Previous 10 | Select-Object JobType, JobState, StartTime, EndTime, HResult"; Type="PowerShell"; FeatureName="Windows-Server-Backup"; Category="Backup"},
    @{Name="Windows Server Backup Disks"; Command="Get-WBDisk | Select-Object DiskNumber, Label, InternalDiskNumber"; Type="PowerShell"; FeatureName="Windows-Server-Backup"; Category="Backup"},
    
    # === NETWORK POLICY AND ACCESS SERVICES (NPAS) ===
    @{Name="NPS Server Konfiguration"; Command="netsh nps show config"; Type="CMD"; FeatureName="NPAS"; Category="NPAS"},
    @{Name="NPS Network Policies"; Command="Get-NpsNetworkPolicy | Select-Object PolicyName, Enabled, ProcessingOrder, ConditionText"; Type="PowerShell"; FeatureName="NPAS"; Category="NPAS"},
    @{Name="NPS Connection Request Policies"; Command="Get-NpsConnectionRequestPolicy | Select-Object Name, Enabled, ProcessingOrder"; Type="PowerShell"; FeatureName="NPAS"; Category="NPAS"},
    @{Name="NPS RADIUS Clients"; Command="Get-NpsRadiusClient | Select-Object Name, Address, SharedSecret, VendorName"; Type="PowerShell"; FeatureName="NPAS"; Category="NPAS"},
    
    # === HOST GUARDIAN SERVICE ===
    @{Name="HGS Service Info"; Command="Get-HgsServer | Select-Object Name, State, Version"; Type="PowerShell"; FeatureName="HostGuardianServiceRole"; Category="HGS"},
    @{Name="HGS Attestation Policies"; Command="Get-HgsAttestationPolicy | Select-Object Name, PolicyVersion, Stage"; Type="PowerShell"; FeatureName="HostGuardianServiceRole"; Category="HGS"},
    
    # === REMOTE ACCESS SERVICES ===
    @{Name="DirectAccess Konfiguration"; Command="Get-DAServer | Select-Object ConnectToAddress, TunnelType, AuthenticationMethod"; Type="PowerShell"; FeatureName="RemoteAccess"; Category="RemoteAccess"},
    @{Name="VPN Server Konfiguration"; Command="Get-VpnServerConfiguration | Select-Object TunnelType, EncryptionLevel, IdleDisconnectSeconds"; Type="PowerShell"; FeatureName="RemoteAccess"; Category="RemoteAccess"},
    @{Name="Routing Table"; Command="Get-NetRoute | Select-Object DestinationPrefix, NextHop, RouteMetric, Protocol"; Type="PowerShell"; FeatureName="RemoteAccess"; Category="RemoteAccess"},
    
    # === WINDOWS INTERNAL DATABASE ===
    @{Name="Windows Internal Database Instanzen"; Command="Get-CimInstance -ClassName Win32_Service | Where-Object {`$_.Name -like '*MSSQL*MICROSOFT*'} | Select-Object Name, State, StartMode"; Type="PowerShell"; FeatureName="Windows-Internal-Database"; Category="InternalDB"},
    @{Name="SQL Server Express Instanzen"; Command="Get-ItemProperty 'HKLM:\\SOFTWARE\\Microsoft\\Microsoft SQL Server\\Instance Names\\SQL' -ErrorAction SilentlyContinue"; Type="PowerShell"; FeatureName="Windows-Internal-Database"; Category="InternalDB"},
    
    # === WINDOWS DEFENDER FEATURES ===
    @{Name="Windows Defender Status"; Command="Get-MpComputerStatus | Select-Object AntivirusEnabled, AMServiceEnabled, RealTimeProtectionEnabled, IoavProtectionEnabled"; Type="PowerShell"; FeatureName="Windows-Defender-Features"; Category="WindowsDefender"},
    @{Name="Windows Defender Preferences"; Command="Get-MpPreference | Select-Object DisableRealtimeMonitoring, DisableIntrusionPreventionSystem, DisableIOAVProtection"; Type="PowerShell"; FeatureName="Windows-Defender-Features"; Category="WindowsDefender"},
    @{Name="Windows Defender Threats"; Command="Get-MpThreatDetection | Select-Object -First 20 | Select-Object ThreatID, ActionSuccess, DetectionTime, ThreatName"; Type="PowerShell"; FeatureName="Windows-Defender-Features"; Category="WindowsDefender"},
    
    # === WINDOWS PROCESS ACTIVATION SERVICE (WAS) ===
    @{Name="WAS Service Status"; Command="Get-Service WAS | Select-Object Name, Status, StartType, ServiceType"; Type="PowerShell"; FeatureName="Windows-Process-Activation-Service"; Category="WAS"},
    @{Name="Application Pool WAS"; Command="Get-IISAppPool | Select-Object Name, State, ProcessModel, Enable32BitAppOnWin64"; Type="PowerShell"; FeatureName="Windows-Process-Activation-Service"; Category="WAS"},
    
    # === WINDOWS SEARCH SERVICE ===
    @{Name="Windows Search Service"; Command="Get-Service WSearch | Select-Object Name, Status, StartType"; Type="PowerShell"; FeatureName="Windows-Search-Service"; Category="SearchService"},
    @{Name="Search Indexer Status"; Command="Get-CimInstance -ClassName Win32_Service | Where-Object Name -eq 'WSearch' | Select-Object Name, State, ProcessId"; Type="PowerShell"; FeatureName="Windows-Search-Service"; Category="SearchService"},
    
    # === ERWEITERTE SYSTEM-AUDITS (basierend auf GitHub Best Practices) ===
    @{Name="Lokale Administratoren"; Command="net localgroup administrators"; Type="CMD"; Category="Security"},
    @{Name="Guest Account Status"; Command="net user guest"; Type="CMD"; Category="Security"},
    @{Name="Shared Folders"; Command="net share"; Type="CMD"; Category="FileSharing"},
    @{Name="Benutzer Profile"; Command="Get-ChildItem C:\\Users | Select-Object Name, CreationTime, LastWriteTime"; Type="PowerShell"; Category="UserProfiles"},
    @{Name="Windows Firewall Profile"; Command="netsh advfirewall show allprofiles"; Type="CMD"; Category="Firewall"},
    @{Name="Power Management"; Command="powercfg /a"; Type="CMD"; Category="PowerManagement"},
    @{Name="Credential Manager"; Command="vaultcmd /listschema; vaultcmd /list"; Type="CMD"; Category="CredentialManager"},
    @{Name="Audit Policy Settings"; Command="auditpol.exe /get /category:*"; Type="CMD"; Category="AuditPolicy"},
    @{Name="Group Policy Results"; Command="gpresult /r"; Type="CMD"; Category="GroupPolicy"},
    @{Name="Installed Software (Registry)"; Command="Get-ItemProperty HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate"; Type="PowerShell"; Category="InstalledSoftware"},
    @{Name="Environment Variables"; Command="Get-ChildItem Env: | Sort-Object Name"; Type="PowerShell"; Category="Environment"},
    
    # === AD HEALTH CHECK (basierend auf Microsoft DevBlogs) ===
    @{Name="AD DHCP Server in AD"; Command="Get-ADObject -SearchBase (`"cn=configuration,`" + (Get-ADDomain).DistinguishedName) -Filter `"objectclass -eq 'dhcpclass' -AND Name -ne 'dhcproot'`" | Select-Object Name"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Service Dependencies Health"; Command="`$Services='DNS','DFS Replication','Intersite Messaging','Kerberos Key Distribution Center','NetLogon','Active Directory Domain Services'; ForEach (`$Service in `$Services) {Get-Service `$Service -ErrorAction SilentlyContinue | Select-Object Name, Status, StartType}"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD DC Diagnostics"; Command="dcdiag /test:dns /e /v"; Type="CMD"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Time Sync Status"; Command="w32tm /query /status"; Type="CMD"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    @{Name="AD Sysvol Replication"; Command="dfsrdiag replicationstate /member:*"; Type="CMD"; FeatureName="AD-Domain-Services"; Category="Active-Directory"},
    
    # === WINDOWS SERVER ESSENTIALS ===
    @{Name="Server Essentials Dashboard"; Command="Get-WssUser | Select-Object UserName, FullName, Description, Enabled"; Type="PowerShell"; FeatureName="ServerEssentialsRole"; Category="ServerEssentials"},
    @{Name="Server Essentials Backup"; Command="Get-WssClientBackup | Select-Object ComputerName, LastBackupTime, LastSuccessfulBackupTime"; Type="PowerShell"; FeatureName="ServerEssentialsRole"; Category="ServerEssentials"},
    
    # === STORAGE MANAGEMENT ===
    @{Name="Storage Pools"; Command="Get-StoragePool | Select-Object FriendlyName, HealthStatus, OperationalStatus, TotalPhysicalCapacity"; Type="PowerShell"; FeatureName="Windows-Storage-Management"; Category="Storage"},
    @{Name="Virtual Disks"; Command="Get-VirtualDisk | Select-Object FriendlyName, HealthStatus, OperationalStatus, Size, AllocatedSize"; Type="PowerShell"; FeatureName="Windows-Storage-Management"; Category="Storage"},
    @{Name="Storage Spaces"; Command="Get-StorageSpace | Select-Object FriendlyName, HealthStatus, ProvisioningType, ResiliencySettingName"; Type="PowerShell"; FeatureName="Windows-Storage-Management"; Category="Storage"},
    @{Name="Physical Disks"; Command="Get-PhysicalDisk | Select-Object FriendlyName, HealthStatus, OperationalStatus, Size, MediaType"; Type="PowerShell"; FeatureName="Windows-Storage-Management"; Category="Storage"},
    
    # === MIGRATION SERVICES ===
    @{Name="Windows Server Migration Tools"; Command="Get-SmigServerFeature | Select-Object FeatureName, Status"; Type="PowerShell"; FeatureName="Windows-Server-Migration"; Category="Migration"},
    
    # === WINDOWS IDENTITY FOUNDATION ===
    @{Name="Windows Identity Foundation"; Command="Get-WindowsFeature Windows-Identity-Foundation | Select-Object Name, InstallState, FeatureType"; Type="PowerShell"; FeatureName="Windows-Identity-Foundation"; Category="Identity"}
)

# Erweiterte Verbindungsaudit-Befehle fuer umfassende Netzwerkanalyse
$connectionAuditCommands = @(
    # === NETZWERK-VERBINDUNGSBAUM ===
    @{Name="Verbindungsbaum (Aktive TCP)"; Command="Get-NetTCPConnection | Where-Object { `$_.State -eq 'Established' } | ForEach-Object { `$proc = Get-Process -Id `$_.OwningProcess -ErrorAction SilentlyContinue; [PSCustomObject]@{ LocalIP=`$_.LocalAddress; LocalPort=`$_.LocalPort; RemoteIP=`$_.RemoteAddress; RemotePort=`$_.RemotePort; Process=if(`$proc){`$proc.ProcessName}else{'N/A'}; PID=`$_.OwningProcess; User=if(`$proc){try{`$proc.StartInfo.UserName}catch{'System'}}else{'N/A'} } } | Sort-Object Process, RemoteIP"; Type="PowerShell"; Category="Verbindungsbaum"},
    @{Name="Netzwerk-Topologie-Map"; Command="`$adapters = Get-NetAdapter | Where-Object Status -eq 'Up'; `$routes = Get-NetRoute | Where-Object RouteMetric -lt 500; `$topology = @(); foreach(`$adapter in `$adapters) { `$ip = Get-NetIPAddress -InterfaceIndex `$adapter.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -First 1; if(`$ip) { `$gateway = `$routes | Where-Object { `$_.InterfaceIndex -eq `$adapter.InterfaceIndex -and `$_.DestinationPrefix -eq '0.0.0.0/0' } | Select-Object -First 1; `$topology += [PSCustomObject]@{ Interface=`$adapter.Name; MAC=`$adapter.MacAddress; IP=`$ip.IPAddress; Subnet=`$ip.PrefixLength; Gateway=if(`$gateway){`$gateway.NextHop}else{'N/A'}; Speed=`$adapter.LinkSpeed } } } `$topology | Format-Table -AutoSize"; Type="PowerShell"; Category="Verbindungsbaum"},
    @{Name="Prozess-Netzwerk-Zuordnung (Erweitert)"; Command="Get-NetTCPConnection | Group-Object OwningProcess | ForEach-Object { `$proc = Get-Process -Id `$_.Name -ErrorAction SilentlyContinue; `$connections = `$_.Group; [PSCustomObject]@{ PID=`$_.Name; ProcessName=if(`$proc){`$proc.ProcessName}else{'Unknown'}; ProcessPath=if(`$proc){try{`$proc.MainModule.FileName}catch{'N/A'}}else{'N/A'}; AnzahlVerbindungen=`$connections.Count; AktiveVerbindungen=(`$connections | Where-Object State -eq 'Established').Count; LauschtPorts=(`$connections | Where-Object State -eq 'Listen').Count; ExterneIPs=(`$connections | Where-Object { `$_.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.|^::1|^fe80:' -and `$_.RemoteAddress -ne '0.0.0.0' -and `$_.RemoteAddress -ne '::' } | Select-Object -ExpandProperty RemoteAddress -Unique | Measure-Object).Count; StartZeit=if(`$proc){`$proc.StartTime}else{'N/A'} } } | Sort-Object AnzahlVerbindungen -Descending"; Type="PowerShell"; Category="Verbindungsbaum"},

    # === AKTIVE NETZWERKVERBINDUNGEN (Erweitert) ===
    @{Name="Alle TCP-Verbindungen (Performance)"; Command="Get-NetTCPConnection | Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, State, OwningProcess, CreationTime | Sort-Object State, LocalPort"; Type="PowerShell"; Category="TCP-Connections"},
    @{Name="Etablierte TCP-Verbindungen"; Command="Get-NetTCPConnection -State Established | Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, OwningProcess, CreationTime"; Type="PowerShell"; Category="TCP-Connections"},
    @{Name="Lauschende Ports (Listen)"; Command="Get-NetTCPConnection -State Listen | Select-Object LocalAddress, LocalPort, OwningProcess | Sort-Object LocalPort"; Type="PowerShell"; Category="TCP-Connections"},
    @{Name="UDP-Endpunkte"; Command="Get-NetUDPEndpoint | Select-Object LocalAddress, LocalPort, OwningProcess | Sort-Object LocalPort"; Type="PowerShell"; Category="UDP-Connections"},
    @{Name="Externe Verbindungen (Internet)"; Command="Get-NetTCPConnection | Where-Object {`$_.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.|^::1|^fe80:' -and `$_.RemoteAddress -ne '0.0.0.0' -and `$_.RemoteAddress -ne '::'} | Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, State, OwningProcess"; Type="PowerShell"; Category="External-Connections"},
    
    # === ERWEITERTE PROZESS-NETZWERK ANALYSE ===
    @{Name="Top-Prozesse nach Verbindungen"; Command="Get-NetTCPConnection | Group-Object OwningProcess | Sort-Object Count -Descending | Select-Object -First 10 | ForEach-Object { `$proc = Get-Process -Id `$_.Name -ErrorAction SilentlyContinue; [PSCustomObject]@{ PID=`$_.Name; Process=if(`$proc){`$proc.ProcessName}else{'Unknown'}; Verbindungen=`$_.Count; Established=(`$_.Group | Where-Object State -eq 'Established').Count; Listen=(`$_.Group | Where-Object State -eq 'Listen').Count } }"; Type="PowerShell"; Category="Process-Network"},
    @{Name="Verdächtige Prozess-Verbindungen"; Command="Get-NetTCPConnection | Where-Object { `$_.State -eq 'Established' -and `$_.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.' } | ForEach-Object { `$proc = Get-Process -Id `$_.OwningProcess -ErrorAction SilentlyContinue; if(`$proc -and `$proc.Path -notmatch 'Windows|Program Files') { [PSCustomObject]@{ Process=`$proc.ProcessName; PID=`$_.OwningProcess; Path=try{`$proc.MainModule.FileName}catch{'N/A'}; RemoteIP=`$_.RemoteAddress; RemotePort=`$_.RemotePort; Company=try{`$proc.Company}catch{'N/A'} } } } | Sort-Object Process"; Type="PowerShell"; Category="Process-Network"},
    @{Name="Netzwerk-Statistiken pro Prozess"; Command="Get-Process | Where-Object { `$_.Id -ne 0 } | ForEach-Object { `$pid = `$_.Id; `$tcpConns = @(Get-NetTCPConnection | Where-Object OwningProcess -eq `$pid); `$udpConns = @(Get-NetUDPEndpoint | Where-Object OwningProcess -eq `$pid); if(`$tcpConns.Count -gt 0 -or `$udpConns.Count -gt 0) { [PSCustomObject]@{ ProcessName=`$_.ProcessName; PID=`$pid; TCP_Verbindungen=`$tcpConns.Count; UDP_Verbindungen=`$udpConns.Count; CPU_Percent=0; RAM_MB=[math]::Round(`$_.WorkingSet64/1MB,2) } } } | Sort-Object TCP_Verbindungen -Descending"; Type="PowerShell"; Category="Process-Network"},

    # === LOKALE GERÄTE UND NETZWERK-ERKENNUNG ===
    @{Name="ARP-Cache (Lokale Geräte)"; Command="Get-NetNeighbor | Where-Object State -ne 'Unreachable' | Select-Object IPAddress, MacAddress, State, InterfaceAlias | Sort-Object IPAddress"; Type="PowerShell"; Category="Local-Devices"},
    @{Name="MAC-Adressen-Analyse"; Command="Get-NetNeighbor | Where-Object { `$_.MacAddress -ne '00-00-00-00-00-00' -and `$_.State -ne 'Unreachable' } | ForEach-Object { `$vendor = switch -Regex (`$_.MacAddress.Substring(0,8)) { '^00-50-56' {'VMware'}; '^00-0C-29' {'VMware'}; '^08-00-27' {'VirtualBox'}; '^00-15-5D' {'Microsoft Hyper-V'}; '^00-1B-21' {'Dell'}; '^00-25-90' {'Dell'}; '^D4-BE-D9' {'Dell'}; '^B8-2A-72' {'Dell'}; '^70-B3-D5' {'HP'}; '^3C-D9-2B' {'HP'}; '^94-57-A5' {'HP'}; default {'Unknown'} }; [PSCustomObject]@{ IP=`$_.IPAddress; MAC=`$_.MacAddress; Vendor=`$vendor; State=`$_.State; Interface=`$_.InterfaceAlias } } | Sort-Object Vendor, IP"; Type="PowerShell"; Category="Local-Devices"},
    @{Name="DHCP-Lease-Informationen"; Command="try { Get-DhcpServerv4Lease -ComputerName localhost | Select-Object IPAddress, ClientId, HostName, AddressState, LeaseExpiryTime | Sort-Object IPAddress } catch { 'DHCP-Server nicht verfügbar oder keine Berechtigung' }"; Type="PowerShell"; Category="Local-Devices"},
    @{Name="Wireless-Netzwerke (Falls verfügbar)"; Command="try { netsh wlan show profiles } catch { 'Wireless-Adapter nicht verfügbar' }"; Type="CMD"; Category="Local-Devices"},

    # === DNS UND NETZWERK-AUFLÖSUNG ===
    @{Name="DNS-Cache-Analyse"; Command="Get-DnsClientCache | Where-Object { `$_.Type -eq 'A' } | Select-Object Name, Data, TTL, Section | Sort-Object Name"; Type="PowerShell"; Category="DNS-Info"},
    @{Name="DNS-Server-Konfiguration"; Command="Get-DnsClientServerAddress | Where-Object { `$_.ServerAddresses.Count -gt 0 } | Select-Object InterfaceAlias, @{Name='DNS_Server';Expression={`$_.ServerAddresses -join ', '}} | Sort-Object InterfaceAlias"; Type="PowerShell"; Category="DNS-Info"},
    @{Name="Reverse-DNS-Lookup (Top-IPs)"; Command="`$topIPs = Get-NetTCPConnection | Where-Object { `$_.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.|^::1|^fe80:' -and `$_.RemoteAddress -ne '0.0.0.0' -and `$_.RemoteAddress -ne '::' } | Group-Object RemoteAddress | Sort-Object Count -Descending | Select-Object -First 10; `$topIPs | ForEach-Object { `$ip = `$_.Name; try { `$hostname = [System.Net.Dns]::GetHostEntry(`$ip).HostName } catch { `$hostname = 'N/A' }; [PSCustomObject]@{ IP=`$ip; Hostname=`$hostname; Verbindungen=`$_.Count } } | Sort-Object Verbindungen -Descending"; Type="PowerShell"; Category="DNS-Info"},
    @{Name="Domänen-DNS-Informationen"; Command="nslookup `$env:USERDNSDOMAIN 2>null | Select-String -Pattern 'Address|Server'"; Type="CMD"; Category="DNS-Info"},

    # === GEO-IP UND EXTERNE ANALYSE ===
    @{Name="Geo-IP-Analyse (Externe IPs)"; Command="`$externalIPs = Get-NetTCPConnection | Where-Object { `$_.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.|^::1|^fe80:' -and `$_.RemoteAddress -ne '0.0.0.0' -and `$_.RemoteAddress -ne '::' } | Select-Object -ExpandProperty RemoteAddress -Unique | Select-Object -First 5; `$externalIPs | ForEach-Object { `$ip = `$_; [PSCustomObject]@{ IP=`$ip; Land='Online-Analyse erforderlich'; Region='API-Limit'; Stadt='Verfügbar'; ISP='ipinfo.io'; Hostname='Manual-Check' } }"; Type="PowerShell"; Category="Geo-IP"},
    @{Name="Bedrohungsanalyse (Blacklist-Check)"; Command="`$suspiciousIPs = Get-NetTCPConnection | Where-Object { `$_.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.|^::1|^fe80:' -and `$_.RemoteAddress -ne '0.0.0.0' -and `$_.RemoteAddress -ne '::' } | Select-Object -ExpandProperty RemoteAddress -Unique; `$suspiciousIPs | ForEach-Object { `$ip = `$_; `$isPrivate = `$ip -match '^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.'; `$isSuspicious = `$ip -match '^(185\.243\.|91\.234\.|77\.83\.)' -or `$ip -match '^(193\.0\.14\.|208\.67\.222\.)'; [PSCustomObject]@{ IP=`$ip; Type=if(`$isPrivate){'Privat'}elseif(`$isSuspicious){'⚠️ Verdächtig'}else{'Public'}; Port_Count=(Get-NetTCPConnection | Where-Object RemoteAddress -eq `$ip).Count } } | Sort-Object Type, IP"; Type="PowerShell"; Category="Geo-IP"},

    # === FIREWALL UND SICHERHEIT ===
    @{Name="Windows-Firewall-Status"; Command="Get-NetFirewallProfile | Select-Object Name, Enabled, DefaultInboundAction, DefaultOutboundAction, LogFileName"; Type="PowerShell"; Category="Firewall-Logs"},
    @{Name="Firewall-Regeln (Aktiv)"; Command="Get-NetFirewallRule | Where-Object { `$_.Enabled -eq 'True' -and `$_.Direction -eq 'Inbound' } | Select-Object DisplayName, Direction, Action, Protocol, LocalPort | Sort-Object Protocol, LocalPort | Select-Object -First 50"; Type="PowerShell"; Category="Firewall-Logs"},
    @{Name="Firewall-Verbindungslogs"; Command="try { Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Windows Firewall With Advanced Security/Firewall'} -MaxEvents 50 | Select-Object TimeCreated, Id, LevelDisplayName, Message | Sort-Object TimeCreated -Descending } catch { 'Firewall-Logs nicht verfügbar' }"; Type="PowerShell"; Category="Firewall-Logs"},
    @{Name="Blockierte Verbindungen"; Command="try { Get-WinEvent -FilterHashtable @{LogName='Security'; ID=5157} -MaxEvents 20 | ForEach-Object { `$xml = [xml]`$_.ToXml(); [PSCustomObject]@{ Zeit=`$_.TimeCreated; Prozess=`$xml.Event.EventData.Data | Where-Object Name -eq 'Application' | Select-Object -ExpandProperty '#text'; Quelle=`$xml.Event.EventData.Data | Where-Object Name -eq 'SourceAddress' | Select-Object -ExpandProperty '#text'; Ziel=`$xml.Event.EventData.Data | Where-Object Name -eq 'DestAddress' | Select-Object -ExpandProperty '#text'; Port=`$xml.Event.EventData.Data | Where-Object Name -eq 'DestPort' | Select-Object -ExpandProperty '#text' } } | Sort-Object Zeit -Descending } catch { 'Sicherheitslogs nicht verfügbar oder keine Berechtigung' }"; Type="PowerShell"; Category="Firewall-Logs"},

    # === NETZWERK-EVENTS UND MONITORING ===
    @{Name="Netzwerk-Sicherheitsereignisse"; Command="try { Get-WinEvent -FilterHashtable @{LogName='Security'; ID=5156} -MaxEvents 30 | Select-Object TimeCreated, Id, LevelDisplayName, Message | Sort-Object TimeCreated -Descending } catch { 'Sicherheitslogs nicht verfügbar' }"; Type="PowerShell"; Category="Network-Events"},
    @{Name="Netzwerk-Adapter-Events"; Command="try { Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Kernel-Network/Analytic'} -MaxEvents 20 | Select-Object TimeCreated, Id, LevelDisplayName, Message | Sort-Object TimeCreated -Descending } catch { 'Kernel-Network-Logs nicht verfügbar' }"; Type="PowerShell"; Category="Network-Events"},
    @{Name="Prozess-Netzwerk-Events"; Command="try { Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Kernel-Process/Analytic'} -MaxEvents 25 | Where-Object { `$_.Message -like '*network*' -or `$_.Message -like '*socket*' } | Select-Object TimeCreated, Id, ProcessId, Message | Sort-Object TimeCreated -Descending } catch { 'Process-Events nicht verfügbar' }"; Type="PowerShell"; Category="Network-Events"},

    # === ACTIVE DIRECTORY UND DOMÄNEN-INFORMATIONEN ===
    @{Name="Domänen-Controller-Informationen"; Command="try { Get-ADDomainController -Discover -Service ADWS,KDC,TimeService | Select-Object Name, IPv4Address, Site, OperatingSystem, Domain } catch { try { nltest /dclist:`$env:USERDNSDOMAIN } catch { 'AD-Modul nicht verfügbar' } }"; Type="PowerShell"; Category="Domain-Users"},
    @{Name="Privilegierte AD-Gruppen"; Command="try { @('Domain Admins', 'Enterprise Admins', 'Schema Admins', 'Administrators') | ForEach-Object { `$group = `$_; try { Get-ADGroupMember -Identity `$group | Select-Object @{Name='Group';Expression={`$group}}, Name, SamAccountName, objectClass } catch { [PSCustomObject]@{Group=`$group; Name='Gruppe nicht gefunden'; SamAccountName='N/A'; objectClass='N/A'} } } } catch { 'AD-PowerShell-Modul nicht verfügbar' }"; Type="PowerShell"; Category="Domain-Users"},
    @{Name="Kürzliche AD-Anmeldungen"; Command="try { Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4624} -MaxEvents 50 | Where-Object { `$_.Message -notlike '*ANONYMOUS*' } | ForEach-Object { `$msg = `$_.Message; `$user = if(`$msg -match 'Account Name:\\s+([^\\r\\n]+)') { `$matches[1] } else { 'Unknown' }; `$workstation = if(`$msg -match 'Workstation Name:\\s+([^\\r\\n]+)') { `$matches[1] } else { 'Unknown' }; [PSCustomObject]@{ Zeit=`$_.TimeCreated; Benutzer=`$user; Workstation=`$workstation; LogonType=if(`$msg -match 'Logon Type:\\s+(\\d+)') { `$matches[1] } else { 'Unknown' } } } | Where-Object { `$_.Benutzer -ne '-' -and `$_.Benutzer -ne 'ANONYMOUS LOGON' } | Sort-Object Zeit -Descending | Select-Object -First 20 } catch { 'Sicherheitslogs nicht verfügbar' }"; Type="PowerShell"; Category="Domain-Users"},
    @{Name="LDAP-Verbindungstests"; Command="try { `$domain = `$env:USERDNSDOMAIN; if(`$domain) { `$dcIP = (nslookup `$domain 2>null | Select-String -Pattern '\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}' | Select-Object -First 1).Matches.Value; if(`$dcIP) { Test-NetConnection -ComputerName `$dcIP -Port 389; Test-NetConnection -ComputerName `$dcIP -Port 636 } else { 'Domain-Controller-IP nicht ermittelbar' } } else { 'Nicht in einer Domäne' } } catch { 'LDAP-Test fehlgeschlagen' }"; Type="PowerShell"; Category="Domain-Users"},

    # === REMOTE-SESSIONS UND RDP ===
    @{Name="Remote-Desktop-Verbindungen"; Command="try { Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-TerminalServices-LocalSessionManager/Operational'} -MaxEvents 30 | Select-Object TimeCreated, Id, LevelDisplayName, Message | Sort-Object TimeCreated -Descending } catch { 'RDP-Logs nicht verfügbar' }"; Type="PowerShell"; Category="Remote-Sessions"},
    @{Name="SMB-Verbindungen"; Command="try { Get-SmbConnection | Select-Object ServerName, ShareName, UserName, Dialect } catch { 'SMB-Informationen nicht verfügbar' }"; Type="PowerShell"; Category="Remote-Sessions"},
    @{Name="SMB-Freigaben-Zugriffe"; Command="try { Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-SmbServer/Security'} -MaxEvents 50 | Select-Object TimeCreated, Id, Message | Sort-Object TimeCreated -Descending } catch { 'SMB-Security-Logs nicht verfügbar' }"; Type="PowerShell"; Category="Remote-Sessions"},
    @{Name="Aktive Terminal-Sessions"; Command="try { quser 2>null } catch { try { query session } catch { 'Terminal-Session-Abfrage nicht verfügbar' } }"; Type="CMD"; Category="Remote-Sessions"},

    # === ERWEITERTE NETZWERK-TOPOLOGIE ===
    @{Name="Routing-Tabelle (Detailliert)"; Command="Get-NetRoute | Select-Object DestinationPrefix, NextHop, InterfaceAlias, RouteMetric, Protocol, @{Name='NetworkCategory';Expression={if(`$_.DestinationPrefix -eq '0.0.0.0/0'){'Default Gateway'}elseif(`$_.DestinationPrefix -like '169.254.*'){'APIPA'}elseif(`$_.DestinationPrefix -like '224.*'){'Multicast'}else{'Network Route'}}} | Sort-Object RouteMetric, DestinationPrefix"; Type="PowerShell"; Category="Network-Topology"},
    @{Name="Netzwerk-Interface-Statistiken"; Command="Get-NetAdapterStatistics | Select-Object Name, @{Name='Empfangen_GB';Expression={[math]::Round(`$_.BytesReceived/1GB,3)}}, @{Name='Gesendet_GB';Expression={[math]::Round(`$_.BytesSent/1GB,3)}}, @{Name='Pakete_Empfangen';Expression={`$_.PacketsReceived}}, @{Name='Pakete_Gesendet';Expression={`$_.PacketsSent}}, @{Name='Fehler_Eingehend';Expression={`$_.InboundErrors}}, @{Name='Fehler_Ausgehend';Expression={`$_.OutboundErrors}} | Sort-Object Empfangen_GB -Descending"; Type="PowerShell"; Category="Network-Topology"},
    @{Name="Gateway- und DNS-Konfiguration"; Command="Get-NetIPConfiguration | Where-Object {`$_.IPv4DefaultGateway -or `$_.IPv6DefaultGateway} | Select-Object InterfaceAlias, @{Name='IPv4_Adresse';Expression={(`$_.IPv4Address | Select-Object -First 1).IPAddress}}, @{Name='IPv4_Gateway';Expression={(`$_.IPv4DefaultGateway | Select-Object -First 1).NextHop}}, @{Name='DNS_Server';Expression={`$_.DNSServer.ServerAddresses -join '; '}}, @{Name='DHCP_Aktiviert';Expression={`$_.NetProfile.NetworkCategory}} | Sort-Object InterfaceAlias"; Type="PowerShell"; Category="Network-Topology"},
    @{Name="Netzwerk-Troubleshooting-Infos"; Command="`$networkInfo = @(); Get-NetAdapter | Where-Object Status -eq 'Up' | ForEach-Object { `$adapter = `$_; `$tcpStats = Get-NetTCPConnection | Where-Object { try { (Get-NetAdapter -InterfaceIndex `$_.LocalAddress -ErrorAction SilentlyContinue).InterfaceIndex -eq `$adapter.InterfaceIndex } catch { `$false } }; `$networkInfo += [PSCustomObject]@{ Interface=`$adapter.Name; MAC=`$adapter.MacAddress; Status=`$adapter.Status; LinkSpeed=`$adapter.LinkSpeed; MediaType=`$adapter.MediaType; TCP_Verbindungen=(`$tcpStats | Measure-Object).Count; Typ=if(`$adapter.Virtual){'Virtual'}else{'Physical'} } }; `$networkInfo | Sort-Object Interface"; Type="PowerShell"; Category="Network-Topology"}
)

# Variable fuer die Verbindungsaudit-Ergebnisse
$global:connectionAuditResults = @{}

# Erweiterte Netzwerk-Verbindungsbaum-Analyse-Funktionen
function Get-NetworkConnectionTree {
    Write-DebugLog "Erstelle erweiterten Netzwerk-Verbindungsbaum" "ConnectionAudit"
    
    try {
        # Sammle alle aktiven Verbindungen
        $tcpConnections = Get-NetTCPConnection | Where-Object { $_.State -eq 'Established' }
        $udpConnections = Get-NetUDPEndpoint # Diese Variable wird aktuell nicht weiter verwendet, könnte für zukünftige Erweiterungen sein
        $processes = Get-Process
        
        # Erstelle Verbindungsbaum-Struktur
        $connectionTree = @{
            Timestamp = Get-Date
            ServerInfo = @{
                ComputerName = $env:COMPUTERNAME
                Domain = $env:USERDNSDOMAIN
                User = $env:USERNAME
                OS = (Get-CimInstance Win32_OperatingSystem).Caption
            }
            NetworkTopology = @{}
            ActiveConnections = @()
            ProcessMapping = @{} # Diese Struktur wird gefüllt, aber nicht explizit im Rückgabewert verwendet, könnte implizit durch $processInfo sein
            ExternalConnections = @()
            SecurityAnalysis = @{} # Diese Struktur wird initialisiert, aber nicht gefüllt
        }
        
        # Netzwerk-Topologie analysieren
        $adapters = Get-NetAdapter | Where-Object Status -eq 'Up'
        foreach ($adapter in $adapters) {
            $ipConfig = Get-NetIPConfiguration -InterfaceIndex $adapter.InterfaceIndex -ErrorAction SilentlyContinue
            if ($ipConfig) {
                $connectionTree.NetworkTopology[$adapter.Name] = @{
                    MAC = $adapter.MacAddress
                    IP = ($ipConfig.IPv4Address | Select-Object -First 1).IPAddress
                    Gateway = ($ipConfig.IPv4DefaultGateway | Select-Object -First 1).NextHop
                    DNS = $ipConfig.DNSServer.ServerAddresses -join ', '
                    Speed = $adapter.LinkSpeed
                    Type = if ($adapter.Virtual) { "Virtual" } else { "Physical" }
                }
            }
        }
        
        # Prozess-Mapping erstellen
        foreach ($connection in $tcpConnections) {
            $process = $processes | Where-Object Id -eq $connection.OwningProcess | Select-Object -First 1
            $processInfo = if ($process) {
                @{
                    Name = $process.ProcessName
                    Path = $(try { $process.MainModule.FileName } catch { "N/A" })
                    StartTime = $process.StartTime
                    Company = $(try { $process.Company } catch { "N/A" })
                }
            } else {
                @{ Name = "Unknown"; Path = "N/A"; StartTime = "N/A"; Company = "N/A" }
            }
            
            $connectionInfo = @{
                LocalAddress = $connection.LocalAddress
                LocalPort = $connection.LocalPort
                RemoteAddress = $connection.RemoteAddress
                RemotePort = $connection.RemotePort
                State = $connection.State
                ProcessInfo = $processInfo
                IsExternal = $connection.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.|^::1|^fe80:'
            }
            
            $connectionTree.ActiveConnections += $connectionInfo
            
            # Externe Verbindungen separat sammeln
            if ($connectionInfo.IsExternal -and $connection.RemoteAddress -ne '0.0.0.0' -and $connection.RemoteAddress -ne '::') { # Zusätzliche Prüfung für IPv6 unspecified
                $connectionTree.ExternalConnections += $connectionInfo
            }
        }
        
        return $connectionTree
    }
    catch {
        Write-DebugLog "FEHLER beim Erstellen des Verbindungsbaums: $($_.Exception.Message) $($_.ScriptStackTrace)" "ConnectionAudit"
        return $null
    }
}

function Format-ConnectionTreeHTML {
    param(
        [hashtable]$ConnectionTree,
        [hashtable]$Results
    )
    
    if (-not $ConnectionTree) {
        return "<p>Verbindungsbaum konnte nicht erstellt werden.</p>"
    }
    
    $html = @"
<div class="connection-tree-container">
    <h2>🌐 Netzwerk-Verbindungsbaum - $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</h2>
    
    <div class="server-summary">
        <h3>📊 Server-Übersicht</h3>
        <div class="info-grid">
            <div class="info-item"><strong>Server:</strong> $($ConnectionTree.ServerInfo.ComputerName)</div>
            <div class="info-item"><strong>Domäne:</strong> $($ConnectionTree.ServerInfo.Domain)</div>
            <div class="info-item"><strong>Benutzer:</strong> $($ConnectionTree.ServerInfo.User)</div>
            <div class="info-item"><strong>OS:</strong> $($ConnectionTree.ServerInfo.OS)</div>
        </div>
    </div>
    
    <div class="network-topology">
        <h3>🔗 Netzwerk-Topologie</h3>
        <table class="styled-table">
            <thead>
                <tr><th>Interface</th><th>IP-Adresse</th><th>Gateway</th><th>DNS</th><th>Typ</th><th>Geschwindigkeit</th></tr>
            </thead>
            <tbody>
"@
    
    foreach ($interface in $ConnectionTree.NetworkTopology.Keys) {
        $topo = $ConnectionTree.NetworkTopology[$interface]
        $html += @"
                <tr>
                    <td>$interface</td>
                    <td>$($topo.IP)</td>
                    <td>$($topo.Gateway)</td>
                    <td>$($topo.DNS)</td>
                    <td>$($topo.Type)</td>
                    <td>$($topo.Speed)</td>
                </tr>
"@
    }
    
    $html += @"
            </tbody>
        </table>
    </div>
    
    <div class="active-connections">
        <h3>⚡ Aktive Verbindungen ($(($ConnectionTree.ActiveConnections | Measure-Object).Count))</h3>
        <table class="styled-table">
            <thead>
                <tr><th>Prozess</th><th>Lokal</th><th>Remote</th><th>Status</th><th>Typ</th><th>Firma</th></tr>
            </thead>
            <tbody>
"@
    
    foreach ($conn in ($ConnectionTree.ActiveConnections | Sort-Object { $_.ProcessInfo.Name })) {
        $connectionType = if ($conn.IsExternal) { "🌍 Extern" } else { "🏠 Lokal" }
        $html += @"
                <tr class="$(if ($conn.IsExternal) { 'external-connection' } else { 'local-connection' })">
                    <td><strong>$($conn.ProcessInfo.Name)</strong></td>
                    <td>$($conn.LocalAddress):$($conn.LocalPort)</td>
                    <td>$($conn.RemoteAddress):$($conn.RemotePort)</td>
                    <td>$($conn.State)</td>
                    <td>$connectionType</td>
                    <td>$($conn.ProcessInfo.Company)</td>
                </tr>
"@
    }
    
    $html += @"
            </tbody>
        </table>
    </div>
    
    <div class="external-analysis">
        <h3>🌍 Externe Verbindungen - Sicherheitsanalyse</h3>
        <p><strong>Anzahl externer Verbindungen:</strong> $(($ConnectionTree.ExternalConnections | Measure-Object).Count)</p>
        <table class="styled-table">
            <thead>
                <tr><th>Remote-IP</th><th>Port</th><th>Prozess</th><th>Pfad</th><th>Bewertung</th></tr>
            </thead>
            <tbody>
"@
    
    foreach ($extConn in ($ConnectionTree.ExternalConnections | Sort-Object RemoteAddress)) {
        $risk = "✅ Normal"
        if ($extConn.ProcessInfo.Path -notmatch "Windows|Program Files") {
            $risk = "⚠️ Prüfen"
        }
        if ($extConn.RemoteAddress -match "^(185\.243\.|91\.234\.|77\.83\.)") {
            $risk = "🚨 Verdächtig"
        }
        
        $html += @"
                <tr class="$(if ($risk -eq '🚨 Verdächtig') { 'suspicious-connection' } elseif ($risk -eq '⚠️ Prüfen') { 'warning-connection' } else { 'normal-connection' })">
                    <td>$($extConn.RemoteAddress)</td>
                    <td>$($extConn.RemotePort)</td>
                    <td>$($extConn.ProcessInfo.Name)</td>
                    <td>$($extConn.ProcessInfo.Path)</td>
                    <td>$risk</td>
                </tr>
"@
    }
    
    $html += @"
            </tbody>
        </table>
    </div>
</div>

<style>
.connection-tree-container { margin: 20px 0; }
.server-summary { background: #f8f9fa; padding: 15px; border-radius: 8px; margin: 15px 0; }
.info-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 10px; }
.info-item { padding: 8px; background: white; border-radius: 4px; border-left: 4px solid #FD7E14; }
.network-topology, .active-connections, .external-analysis { margin: 20px 0; }
.styled-table { width: 100%; border-collapse: collapse; margin: 10px 0; }
.styled-table th { background: #FD7E14; color: white; padding: 12px; text-align: left; }
.styled-table td { padding: 10px; border-bottom: 1px solid #ddd; }
.external-connection { background-color: #fff3cd; }
.local-connection { background-color: #f8f9fa; }
.suspicious-connection { background-color: #f8d7da; }
.warning-connection { background-color: #fff3cd; }
.normal-connection { background-color: #d1f7d1; }
.styled-table tr:hover { background-color: #f5f5f5; }
</style>
"@
    
    return $html
}

# === VERBINDUNGSAUDIT SPEZIFISCHE KOMMANDOS ===
$connectionAuditCommands = @(
    # === AKTIVE NETZWERKVERBINDUNGEN ===
    @{Name="Alle TCP Verbindungen"; Command="Get-NetTCPConnection | Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, State, OwningProcess, CreationTime | Sort-Object State, LocalPort"; Type="PowerShell"; Category="TCP-Connections"},
    @{Name="Etablierte TCP Verbindungen"; Command="Get-NetTCPConnection -State Established | Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, OwningProcess, CreationTime"; Type="PowerShell"; Category="TCP-Connections"},
    @{Name="Lauschende Ports (Listen)"; Command="Get-NetTCPConnection -State Listen | Select-Object LocalAddress, LocalPort, OwningProcess | Sort-Object LocalPort"; Type="PowerShell"; Category="TCP-Connections"},
    @{Name="UDP Endpunkte"; Command="Get-NetUDPEndpoint | Select-Object LocalAddress, LocalPort, OwningProcess | Sort-Object LocalPort"; Type="PowerShell"; Category="UDP-Connections"},
    @{Name="Externe Verbindungen (nicht lokal)"; Command="Get-NetTCPConnection | Where-Object {`$_.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.|^::1|^fe80:' -and `$_.RemoteAddress -ne '0.0.0.0' -and `$_.RemoteAddress -ne '::'} | Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, State, OwningProcess"; Type="PowerShell"; Category="External-Connections"},
    
    # === PROZESS-NETZWERK ZUORDNUNG ===
    @{Name="Prozesse mit Netzwerkverbindungen"; Command="Get-NetTCPConnection | Where-Object {`$_.State -eq 'Established'} | ForEach-Object { `$conn = `$_; try { `$process = Get-Process -Id `$conn.OwningProcess -ErrorAction Stop; [PSCustomObject]@{ ProcessName = `$process.ProcessName; PID = `$process.Id; LocalAddress = `$conn.LocalAddress; LocalPort = `$conn.LocalPort; RemoteAddress = `$conn.RemoteAddress; RemotePort = `$conn.RemotePort; ProcessPath = `$process.Path } } catch { [PSCustomObject]@{ ProcessName = 'Unknown'; PID = `$conn.OwningProcess; LocalAddress = `$conn.LocalAddress; LocalPort = `$conn.LocalPort; RemoteAddress = `$conn.RemoteAddress; RemotePort = `$conn.RemotePort; ProcessPath = 'N/A' } } } | Sort-Object ProcessName"; Type="PowerShell"; Category="Process-Network"},
    @{Name="Top Prozesse nach Verbindungen"; Command="`$connections = Get-NetTCPConnection | Group-Object OwningProcess; `$connections | ForEach-Object { try { `$process = Get-Process -Id `$_.Name -ErrorAction Stop; [PSCustomObject]@{ ProcessName = `$process.ProcessName; PID = `$_.Name; ConnectionCount = `$_.Count; ProcessPath = `$process.Path } } catch { [PSCustomObject]@{ ProcessName = 'Unknown'; PID = `$_.Name; ConnectionCount = `$_.Count; ProcessPath = 'N/A' } } } | Sort-Object ConnectionCount -Descending | Select-Object -First 20"; Type="PowerShell"; Category="Process-Network"},
    @{Name="System-Prozesse mit Netzwerkzugriff"; Command="Get-NetTCPConnection | Where-Object {`$_.OwningProcess -lt 1000} | ForEach-Object { `$conn = `$_; try { `$process = Get-Process -Id `$conn.OwningProcess -ErrorAction Stop; [PSCustomObject]@{ ProcessName = `$process.ProcessName; PID = `$process.Id; LocalPort = `$conn.LocalPort; RemoteAddress = `$conn.RemoteAddress; State = `$conn.State } } catch { [PSCustomObject]@{ ProcessName = 'System/Unknown'; PID = `$conn.OwningProcess; LocalPort = `$conn.LocalPort; RemoteAddress = `$conn.RemoteAddress; State = `$conn.State } } } | Sort-Object PID"; Type="PowerShell"; Category="Process-Network"},
    
    # === LOKALE GERAETE (ARP-CACHE) ===
    @{Name="ARP Cache (alle Geraete)"; Command="Get-NetNeighbor | Select-Object IPAddress, MacAddress, State, InterfaceAlias | Sort-Object IPAddress"; Type="PowerShell"; Category="Local-Devices"},
    @{Name="ARP Cache (nur erreichbare)"; Command="Get-NetNeighbor -State Reachable | Select-Object IPAddress, MacAddress, InterfaceAlias"; Type="PowerShell"; Category="Local-Devices"},
    @{Name="MAC-Adressen im lokalen Netz"; Command="arp -a"; Type="CMD"; Category="Local-Devices"},
    @{Name="Netzwerk-Interfaces"; Command="Get-NetAdapter | Select-Object Name, InterfaceDescription, LinkSpeed, MacAddress, Status | Sort-Object Name"; Type="PowerShell"; Category="Local-Devices"},
    @{Name="DHCP-Leases (falls DHCP-Server)"; Command="if (Get-WindowsFeature -Name DHCP | Where-Object {`$_.Installed}) { Get-DhcpServerv4Lease -AllLeases | Select-Object IPAddress, ClientId, HostName, LeaseExpiryTime | Sort-Object IPAddress } else { 'DHCP-Server nicht installiert' }"; Type="PowerShell"; Category="Local-Devices"},
    
    # === DNS INFORMATIONEN ===
    @{Name="DNS Cache"; Command="Get-DnsClientCache | Select-Object Entry, Name, Data, TimeToLive | Sort-Object Name"; Type="PowerShell"; Category="DNS-Info"},
    @{Name="DNS Server Konfiguration"; Command="Get-DnsClientServerAddress | Select-Object InterfaceAlias, ServerAddresses"; Type="PowerShell"; Category="DNS-Info"},
    @{Name="Reverse DNS fuer externe IPs"; Command="`$extIPs = Get-NetTCPConnection | Where-Object {`$_.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.|^::1|^fe80:' -and `$_.RemoteAddress -ne '0.0.0.0' -and `$_.RemoteAddress -ne '::'} | Select-Object -ExpandProperty RemoteAddress -Unique; `$extIPs | ForEach-Object { try { `$hostname = [System.Net.Dns]::GetHostEntry(`$_).HostName; [PSCustomObject]@{ IPAddress = `$_; Hostname = `$hostname } } catch { [PSCustomObject]@{ IPAddress = `$_; Hostname = 'Aufloesung fehlgeschlagen' } } } | Sort-Object IPAddress"; Type="PowerShell"; Category="DNS-Info"},
    
    # === GEO-IP INFORMATIONEN ===
    @{Name="Geo-IP Analyse externer Verbindungen"; Command="`$extIPs = Get-NetTCPConnection | Where-Object {`$_.RemoteAddress -notmatch '^127\.|^10\.|^192\.168\.|^172\.(1[6-9]|2[0-9]|3[0-1])\.|^::1|^fe80:' -and `$_.RemoteAddress -ne '0.0.0.0' -and `$_.RemoteAddress -ne '::'} | Select-Object -ExpandProperty RemoteAddress -Unique | Select-Object -First 10; `$extIPs | ForEach-Object { try { `$geoInfo = Invoke-RestMethod `"http://ipinfo.io/`$_/json`" -TimeoutSec 5; [PSCustomObject]@{ IPAddress = `$_; Country = `$geoInfo.country; Region = `$geoInfo.region; City = `$geoInfo.city; Organization = `$geoInfo.org; ISP = `$geoInfo.isp } } catch { [PSCustomObject]@{ IPAddress = `$_; Country = 'N/A'; Region = 'N/A'; City = 'N/A'; Organization = 'Abfrage fehlgeschlagen'; ISP = 'N/A' } } }"; Type="PowerShell"; Category="Geo-IP"},
    
    # === FIREWALL UND LOGGING ===
    @{Name="Firewall Verbindungs-Logs"; Command="if (Test-Path `$env:SystemRoot\\system32\\LogFiles\\Firewall\\pfirewall.log) { Get-Content `$env:SystemRoot\\system32\\LogFiles\\Firewall\\pfirewall.log -Tail 50 | Where-Object {`$_ -match 'ALLOW|DROP'} } else { 'Firewall-Logging nicht aktiviert oder Log-Datei nicht gefunden' }"; Type="PowerShell"; Category="Firewall-Logs"},
    @{Name="Firewall Logging Status"; Command="Get-NetFirewallProfile | Select-Object Name, LogAllowed, LogBlocked, LogFileName, LogMaxSizeKilobytes"; Type="PowerShell"; Category="Firewall-Logs"},
    @{Name="Aktive Firewall Regeln"; Command="Get-NetFirewallRule | Where-Object {`$_.Enabled -eq 'True' -and `$_.Action -eq 'Allow'} | Select-Object DisplayName, Direction, Protocol, LocalPort, RemoteAddress | Sort-Object Direction, Protocol"; Type="PowerShell"; Category="Firewall-Logs"},
    
    # === EVENT LOGS FUER VERBINDUNGEN ===
    @{Name="Netzwerk-Events (Security Log)"; Command="Get-WinEvent -FilterHashtable @{LogName='Security'; ID=5156,5157,5158} -MaxEvents 50 | Select-Object TimeCreated, Id, LevelDisplayName, Message"; Type="PowerShell"; Category="Network-Events"},
    @{Name="Prozessstart-Events"; Command="Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4688} -MaxEvents 30 | Select-Object TimeCreated, Message"; Type="PowerShell"; Category="Network-Events"},
    @{Name="Windows Firewall Events"; Command="Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Windows Firewall With Advanced Security/Firewall'} -MaxEvents 30 | Select-Object TimeCreated, Id, LevelDisplayName, Message"; Type="PowerShell"; Category="Network-Events"},
    
    # === DOMAENEN-USER AUDIT (DE/EN) ===
    @{Name="AD-User (aktuelle Domaene)"; Command="if (Get-Module -ListAvailable -Name ActiveDirectory) { Import-Module ActiveDirectory; Get-ADUser -Filter * -Properties LastLogonDate, PasswordLastSet, Enabled | Select-Object Name, SamAccountName, Enabled, LastLogonDate, PasswordLastSet, DistinguishedName | Sort-Object Name } else { 'ActiveDirectory PowerShell-Modul nicht verfuegbar' }"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Domain-Users"},
    @{Name="Aktuell angemeldete Domaenen-User"; Command="if (Get-Module -ListAvailable -Name ActiveDirectory) { Import-Module ActiveDirectory; Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4624} -MaxEvents 100 | Where-Object {`$_.Message -match 'Logon Type:\\s*[23]' -and `$_.Message -notmatch 'ANONYMOUS|`$'} | ForEach-Object { if (`$_.Message -match 'Account Name:\\s*([^\\r\\n]+)' -and `$_.Message -match 'Account Domain:\\s*([^\\r\\n]+)') { [PSCustomObject]@{ TimeCreated = `$_.TimeCreated; AccountName = `$matches[1].Trim(); AccountDomain = `$matches[2].Trim(); LogonType = if (`$_.Message -match 'Logon Type:\\s*(\\d+)') { `$matches[1] } else { 'Unknown' } } } } | Where-Object {`$_.AccountName -ne '-' -and `$_.AccountName -ne 'ANONYMOUS LOGON'} | Sort-Object TimeCreated -Descending | Select-Object -First 20 } else { 'ActiveDirectory PowerShell-Modul nicht verfuegbar' }"; Type="PowerShell"; Category="Domain-Users"},
    @{Name="Privilegierte AD-Gruppen Mitglieder"; Command="if (Get-Module -ListAvailable -Name ActiveDirectory) { Import-Module ActiveDirectory; `$privGroups = @('Domain Admins', 'Enterprise Admins', 'Schema Admins', 'Administrators', 'Domänen-Admins', 'Organisations-Admins', 'Schema-Admins'); `$results = @(); foreach (`$group in `$privGroups) { try { `$members = Get-ADGroupMember -Identity `$group -ErrorAction SilentlyContinue | Get-ADUser -Properties LastLogonDate -ErrorAction SilentlyContinue; foreach (`$member in `$members) { `$results += [PSCustomObject]@{ GroupName = `$group; UserName = `$member.Name; SamAccountName = `$member.SamAccountName; Enabled = `$member.Enabled; LastLogonDate = `$member.LastLogonDate } } } catch { } }; `$results | Sort-Object GroupName, UserName } else { 'ActiveDirectory PowerShell-Modul nicht verfuegbar' }"; Type="PowerShell"; FeatureName="AD-Domain-Services"; Category="Domain-Users"},
    @{Name="Letzten Anmeldungen (Domaene)"; Command="if (Get-Module -ListAvailable -Name ActiveDirectory) { Import-Module ActiveDirectory; Get-WinEvent -FilterHashtable @{LogName='Security'; ID=4624,4625} -MaxEvents 100 | Where-Object {`$_.Message -notmatch 'ANONYMOUS|`$' -and `$_.Message -match '@|\\\\'} | ForEach-Object { if (`$_.Message -match 'Account Name:\\s*([^\\r\\n]+)' -and `$_.Message -match 'Account Domain:\\s*([^\\r\\n]+)') { [PSCustomObject]@{ TimeCreated = `$_.TimeCreated; EventID = `$_.Id; AccountName = `$matches[1].Trim(); AccountDomain = `$matches[2].Trim(); Status = if (`$_.Id -eq 4624) { 'Erfolg' } else { 'Fehlgeschlagen' } } } } | Where-Object {`$_.AccountName -ne '-'} | Sort-Object TimeCreated -Descending | Select-Object -First 30 } else { 'ActiveDirectory PowerShell-Modul nicht verfuegbar' }"; Type="PowerShell"; Category="Domain-Users"},
    
    # === REMOTE VERBINDUNGEN ===
    @{Name="RDP-Sitzungen"; Command="qwinsta"; Type="CMD"; Category="Remote-Sessions"},
    @{Name="Remote Desktop Events"; Command="Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-TerminalServices-LocalSessionManager/Operational'} -MaxEvents 30 | Select-Object TimeCreated, Id, LevelDisplayName, Message"; Type="PowerShell"; Category="Remote-Sessions"},
    @{Name="SMB-Verbindungen"; Command="Get-SmbConnection | Select-Object ServerName, ShareName, UserName, Dialect"; Type="PowerShell"; Category="Remote-Sessions"},
    @{Name="SMB-Freigaben Zugriffe"; Command="Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-SmbServer/Security'} -MaxEvents 50 | Select-Object TimeCreated, Id, Message"; Type="PowerShell"; Category="Remote-Sessions"},
    
    # === ROUTING UND NETZWERK-TOPOLOGIE ===
    @{Name="Routing Tabelle"; Command="Get-NetRoute | Select-Object DestinationPrefix, NextHop, InterfaceAlias, RouteMetric, Protocol | Sort-Object DestinationPrefix"; Type="PowerShell"; Category="Network-Topology"},
    @{Name="Netzwerk-Statistiken"; Command="Get-NetAdapterStatistics | Select-Object Name, BytesReceived, BytesSent, PacketsReceived, PacketsSent"; Type="PowerShell"; Category="Network-Topology"},
    @{Name="Gateway-Informationen"; Command="Get-NetIPConfiguration | Where-Object {`$_.IPv4DefaultGateway} | Select-Object InterfaceAlias, IPv4Address, IPv4DefaultGateway, DNSServer"; Type="PowerShell"; Category="Network-Topology"}
)

# Variable fuer die Verbindungsaudit-Ergebnisse
$global:connectionAuditResults = @{}

# Funktion zum Ausfuehren von PowerShell-Befehlen
function Invoke-PSCommand {
    param(
        [string]$Command
    )
    try {
        Write-DebugLog "Ausfuehren von PowerShell-Befehl: $Command" "CommandExec"
        
        # Spezielle Behandlung fuer bestimmte Befehle
        if ($Command -like "*Get-ComputerInfo*") {
            $result = Get-ComputerInfo | Format-List | Out-String
        } else {
            $result = Invoke-Expression -Command $Command | Format-Table -AutoSize | Out-String
        }
        
        Write-DebugLog "PowerShell-Befehl erfolgreich ausgefuehrt. Ergebnis-Laenge: $($result.Length)" "CommandExec"
        return $result
    }
    catch {
        $errorMsg = "Fehler bei der Ausfuehrung des Befehls: $Command`r`n$($_.Exception.Message)"
        Write-DebugLog "FEHLER: $errorMsg" "CommandExec"
        return $errorMsg
    }
}

# Funktion zum Ausfuehren von CMD-Befehlen
function Invoke-CMDCommand {
    param(
        [string]$Command
    )
    try {
        Write-DebugLog "Ausfuehren von CMD-Befehl: $Command" "CommandExec"
        $result = cmd /c $Command 2>&1 | Out-String
        Write-DebugLog "CMD-Befehl erfolgreich ausgefuehrt. Ergebnis-Laenge: $($result.Length)" "CommandExec"
        return $result
    }
    catch {
        $errorMsg = "Fehler bei der Ausfuehrung des Befehls: $Command`r`n$($_.Exception.Message)"
        Write-DebugLog "FEHLER: $errorMsg" "CommandExec"
        return $errorMsg
    }
}

# Funktion zum Pruefen, ob eine bestimmte Serverrolle installiert ist
function Test-ServerRole {
    param(
        [string]$FeatureName
    )
    
    try {
        Write-DebugLog "Pruefe Serverrolle: $FeatureName" "RoleCheck"
        $feature = Get-WindowsFeature -Name $FeatureName -ErrorAction SilentlyContinue
        if ($feature -and $feature.Installed) {
            Write-DebugLog "Serverrolle $FeatureName ist installiert" "RoleCheck"
            return $true
        }
        Write-DebugLog "Serverrolle $FeatureName ist NICHT installiert" "RoleCheck"
        return $false
    }
    catch {
        Write-DebugLog "FEHLER beim Pruefen der Serverrolle $FeatureName - $($_.Exception.Message)" "RoleCheck"
        return $false
    }
}
# VERBINDUNGSAUDIT FUNKTIONEN

# Hauptfunktion fuer die Verbindungsaudit-Durchfuehrung
function Start-ConnectionAuditProcess {
    # UI vorbereiten
    $btnRunConnectionAudit.IsEnabled = $false
    $btnExportConnectionHTML.IsEnabled = $false
    $btnCopyConnectionToClipboard.IsEnabled = $false
    $cmbConnectionCategories.IsEnabled = $false
    
    $rtbConnectionResults.Document = New-Object System.Windows.Documents.FlowDocument
    $progressBarConnection.Value = 0
    $txtStatus.Text = "Status: Verbindungsaudit laeuft..."
    
    # UI initial aktualisieren
    $window.Dispatcher.Invoke([Action]{
        $txtProgressConnection.Text = "Initialisiere Verbindungsaudit..."
        $progressBarConnection.Value = 0
    }, "Normal")
    
    # UI refresh erzwingen
    $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
    Start-Sleep -Milliseconds 300
    
    # Sammle ausgewaehlte Befehle
    $selectedCommands = @()
    foreach ($cmd in $connectionAuditCommands) {
        if ($connectionCheckboxes[$cmd.Name].IsChecked) {
            $selectedCommands += $cmd
        }
    }
    
    Write-DebugLog "Starte Verbindungsaudit mit $($selectedCommands.Count) ausgewaehlten Befehlen" "ConnectionAudit"
    
    $global:connectionAuditResults = @{}
    $progressStep = 100.0 / $selectedCommands.Count
    $currentProgress = 0
    
    # UI Update mit Anzahl der Befehle
    $window.Dispatcher.Invoke([Action]{
        $txtProgressConnection.Text = "Bereite $($selectedCommands.Count) Verbindungsaudit-Befehle vor..."
    }, "Normal")
    
    # UI refresh erzwingen
    $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
    Start-Sleep -Milliseconds 500
    
    for ($i = 0; $i -lt $selectedCommands.Count; $i++) {
        $cmd = $selectedCommands[$i]
        
        # UI aktualisieren - BEGINN des Befehls
        $window.Dispatcher.Invoke([Action]{
            $txtProgressConnection.Text = "Verarbeite: $($cmd.Name) ($($i+1)/$($selectedCommands.Count))"
            $progressBarConnection.Value = $currentProgress
        }, "Normal")
        
        # UI refresh erzwingen
        $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 200
        
        Write-DebugLog "Fuehre Verbindungsaudit aus ($($i+1)/$($selectedCommands.Count)): $($cmd.Name)" "ConnectionAudit"
        
        try {
            if ($cmd.Type -eq "PowerShell") {
                $result = Invoke-PSCommand -Command $cmd.Command
            } else {
                $result = Invoke-CMDCommand -Command $cmd.Command
            }
            
            $global:connectionAuditResults[$cmd.Name] = $result
            
            # Erfolg und Fortschrittsbalken aktualisieren
            $currentProgress += $progressStep
            $window.Dispatcher.Invoke([Action]{
                $progressBarConnection.Value = $currentProgress
                $txtProgressConnection.Text = "Abgeschlossen: $($cmd.Name) ($($i+1)/$($selectedCommands.Count))"
            }, "Normal")
            
        } catch {
            $errorMsg = "Fehler: $($_.Exception.Message)"
            $global:connectionAuditResults[$cmd.Name] = $errorMsg
            
            # Fehler und Fortschrittsbalken trotzdem aktualisieren
            $currentProgress += $progressStep
            $window.Dispatcher.Invoke([Action]{
                $progressBarConnection.Value = $currentProgress
                $txtProgressConnection.Text = "Fehler bei: $($cmd.Name) ($($i+1)/$($selectedCommands.Count))"
            }, "Normal")
            
            Write-DebugLog "FEHLER bei Verbindungsaudit $($cmd.Name): $($_.Exception.Message)" "ConnectionAudit"
        }
        
        # UI refresh nach jedem Befehl erzwingen
        $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 300
        
        # Zwischenstand der Ergebnisse aktualisieren
        if (($i + 1) % 3 -eq 0 -or $i -eq ($selectedCommands.Count - 1)) {
            $window.Dispatcher.Invoke([Action]{
                try {
                    Update-ConnectionResultsCategories
                    if ($cmbConnectionCategories.SelectedItem) {
                        $selectedCategory = $cmbConnectionCategories.SelectedItem.Tag
                        Show-ConnectionCategoryResults -Category $selectedCategory
                    } else {
                        Show-ConnectionCategoryResults -Category "Alle"
                    }
                }
                catch {
                    Write-DebugLog "FEHLER beim Zwischenupdate der Verbindungsaudit-Ergebnisanzeige: $($_.Exception.Message)" "ConnectionAudit"
                }
            }, "Normal")
            $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    
    # Verbindungsaudit abgeschlossen - Finale Updates
    $window.Dispatcher.Invoke([Action]{
        $progressBarConnection.Value = 100
        $txtProgressConnection.Text = "Verbindungsaudit vollstaendig abgeschlossen! $($selectedCommands.Count) Befehle ausgefuehrt."
        
        try {
            # Aktualisiere die Kategorien-Anzeige
            Update-ConnectionResultsCategories
            Show-ConnectionCategoryResults -Category "Alle"
        }
        catch {
            Write-DebugLog "FEHLER beim finalen Update der Verbindungsaudit-Ergebnisanzeige: $($_.Exception.Message)" "ConnectionAudit"
            try {
                Show-SimpleConnectionResults -Category "Alle"
            }
            catch {
                Write-DebugLog "FEHLER auch bei einfacher Verbindungsaudit-Anzeige: $($_.Exception.Message)" "ConnectionAudit"
            }
        }
        
        $txtStatus.Text = "Status: Verbindungsaudit abgeschlossen - $($global:connectionAuditResults.Count) Ergebnisse"
        
        # Buttons wieder aktivieren
        $btnRunConnectionAudit.IsEnabled = $true
        $btnExportConnectionHTML.IsEnabled = $true
        $btnCopyConnectionToClipboard.IsEnabled = $true
        $cmbConnectionCategories.IsEnabled = $true
    }, "Normal")
    
    # Finaler UI refresh
    $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
    [System.Windows.Forms.Application]::DoEvents()
    
    Write-DebugLog "Verbindungsaudit abgeschlossen mit $($global:connectionAuditResults.Count) Ergebnissen" "ConnectionAudit"
}

# Funktion zum Aktualisieren der Verbindungsaudit-Kategorien-ComboBox
function Update-ConnectionResultsCategories {
    Write-DebugLog "Aktualisiere Verbindungsaudit-Kategorien-ComboBox" "UI"
    
    $cmbConnectionCategories.Items.Clear()
    
    # "Alle" Option hinzufügen
    $allItem = New-Object System.Windows.Controls.ComboBoxItem
    $allItem.Content = "Alle Kategorien"
    $allItem.Tag = "Alle"
    $cmbConnectionCategories.Items.Add($allItem)
    
    # Einzelne Kategorien hinzufügen
    $categories = @{}
    if ($null -ne $connectionAuditCommands) {
        foreach ($cmd in $connectionAuditCommands) {
            $category = if ($cmd.Category) { $cmd.Category } else { "Allgemein" }
            if (-not $categories.ContainsKey($category)) {
                $categories[$category] = 0
            }
            if ($global:connectionAuditResults.ContainsKey($cmd.Name)) {
                $categories[$category]++
            }
        }
    }
    
    foreach ($category in $categories.Keys | Sort-Object) {
        if ($categories[$category] -gt 0) {
            $categoryItem = New-Object System.Windows.Controls.ComboBoxItem
            $categoryItem.Content = "$category ($($categories[$category]))"
            $categoryItem.Tag = $category
            $cmbConnectionCategories.Items.Add($categoryItem)
        }
    }
    
    # Ersten Eintrag auswählen
    if ($cmbConnectionCategories.Items.Count -gt 0) {
        $cmbConnectionCategories.SelectedIndex = 0
    }
}

# Funktion zum Anzeigen der Verbindungsaudit-Ergebnisse
function Show-ConnectionCategoryResults {
    param([string]$Category = "Alle")
    
    Write-DebugLog "Zeige Verbindungsaudit-Ergebnisse fuer Kategorie: $Category" "UI"
    
    if ($global:connectionAuditResults.Count -eq 0) {
        $rtbConnectionResults.Document = New-Object System.Windows.Documents.FlowDocument
        $emptyParagraph = New-Object System.Windows.Documents.Paragraph
        $emptyRun = New-Object System.Windows.Documents.Run("Keine Verbindungsaudit-Ergebnisse verfügbar. Führen Sie zuerst ein Verbindungsaudit durch.")
        $emptyRun.FontStyle = "Italic"
        $emptyRun.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(108, 117, 125))
        $emptyParagraph.Inlines.Add($emptyRun)
        $rtbConnectionResults.Document.Blocks.Add($emptyParagraph)
        return
    }
    
    try {
        # Versuche die formatierte Anzeige
        $document = Format-ConnectionRichTextResults -Results $global:connectionAuditResults -CategoryFilter $Category
        $rtbConnectionResults.Document = $document
        
        # Erzwinge Layout-Update
        $rtbConnectionResults.UpdateLayout()
        
        Write-DebugLog "Verbindungsaudit-Ergebnisse erfolgreich formatiert und angezeigt" "UI"
    }
    catch {
        Write-DebugLog "FEHLER bei der formatierten Verbindungsaudit-Anzeige: $($_.Exception.Message) - Verwende Fallback" "UI"
        
        # Fallback: Verwende einfache Textanzeige
        Show-SimpleConnectionResults -Category $Category
    }
}

# Funktion zum Formatieren der Verbindungsaudit-RichTextBox
function Format-ConnectionRichTextResults {
    param(
        [hashtable]$Results,
        [string]$CategoryFilter = "Alle"
    )
    
    Write-DebugLog "Formatiere Verbindungsaudit-Ergebnisse fuer Kategorie: $CategoryFilter" "UI"
    
    # Erstelle ein neues FlowDocument
    $document = New-Object System.Windows.Documents.FlowDocument
    $document.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI")
    $document.FontSize = 12
    $document.LineHeight = 18
    
    # Optimierte Layout-Einstellungen
    $document.PageWidth = [Double]::NaN
    $document.PageHeight = [Double]::NaN
    $document.ColumnWidth = [Double]::PositiveInfinity
    $document.TextAlignment = "Left"
    $document.PagePadding = New-Object System.Windows.Thickness(0)
    $document.IsOptimalParagraphEnabled = $true
    $document.IsHyphenationEnabled = $false
    
    # Gruppiere Ergebnisse nach Kategorien
    $categorizedResults = @{}
    foreach ($cmd in $connectionAuditCommands) {
        $category = if ($cmd.Category) { $cmd.Category } else { "Allgemein" }
        if (-not $categorizedResults.ContainsKey($category)) {
            $categorizedResults[$category] = @()
        }
        if ($Results.ContainsKey($cmd.Name)) {
            $categorizedResults[$category] += @{
                Name = $cmd.Name
                Result = $Results[$cmd.Name]
                Command = $cmd
            }
        }
    }
    
    # Bestimme welche Kategorien angezeigt werden sollen
    $categoriesToShow = if ($CategoryFilter -eq "Alle") { 
        $categorizedResults.Keys | Sort-Object 
    } else { 
        @($CategoryFilter) 
    }
    
    $totalItems = 0
    foreach ($category in $categoriesToShow) {
        if ($categorizedResults.ContainsKey($category)) {
            $categoryData = $categorizedResults[$category]
            $totalItems += $categoryData.Count
            
            # Kategorie-Header
            $categoryParagraph = New-Object System.Windows.Documents.Paragraph
            $categoryRun = New-Object System.Windows.Documents.Run("Kategorie: $category")
            $categoryRun.FontWeight = "Bold"
            $categoryRun.FontSize = 16
            $categoryRun.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(253, 126, 20)) # Orange für Verbindungsaudit
            $categoryParagraph.Inlines.Add($categoryRun)
            $categoryParagraph.Margin = New-Object System.Windows.Thickness(0, 15, 0, 10)
            $categoryParagraph.TextAlignment = "Left"
            $categoryParagraph.KeepTogether = $true
            $document.Blocks.Add($categoryParagraph)
            
            # Items in dieser Kategorie
            foreach ($item in $categoryData) {
                # Item-Header
                $itemParagraph = New-Object System.Windows.Documents.Paragraph
                $itemRun = New-Object System.Windows.Documents.Run("Eintrag: $($item.Name)")
                $itemRun.FontWeight = "SemiBold"
                $itemRun.FontSize = 13
                $itemRun.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(44, 62, 80))
                $itemParagraph.Inlines.Add($itemRun)
                $itemParagraph.Margin = New-Object System.Windows.Thickness(0, 10, 0, 5)
                $itemParagraph.TextAlignment = "Left"
                $itemParagraph.KeepTogether = $true
                $document.Blocks.Add($itemParagraph)
                
                # Kommando-Info (optional)
                if ($item.Command.Command) {
                    $cmdParagraph = New-Object System.Windows.Documents.Paragraph
                    $cmdRun = New-Object System.Windows.Documents.Run("Befehl: $($item.Command.Command)")
                    $cmdRun.FontStyle = "Italic"
                    $cmdRun.FontSize = 10
                    $cmdRun.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(108, 117, 125))
                    $cmdParagraph.Inlines.Add($cmdRun)
                    $cmdParagraph.Margin = New-Object System.Windows.Thickness(20, 0, 0, 5)
                    $cmdParagraph.TextAlignment = "Left"
                    $document.Blocks.Add($cmdParagraph)
                }
                
                # Ergebnis in einem optimierten Paragraph
                $resultParagraph = New-Object System.Windows.Documents.Paragraph
                $resultParagraph.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
                $resultParagraph.Padding = New-Object System.Windows.Thickness(15)
                $resultParagraph.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 248, 220)) # Heller Orange-Hintergrund
                $resultParagraph.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(253, 126, 20))
                $resultParagraph.BorderThickness = New-Object System.Windows.Thickness(1)
                $resultParagraph.TextAlignment = "Left"
                
                # Textbehandlung
                $resultText = $item.Result
                if ([string]::IsNullOrWhiteSpace($resultText)) {
                    $resultText = "Keine Daten verfügbar"
                }
                
                $resultRun = New-Object System.Windows.Documents.Run($resultText)
                $resultRun.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
                $resultRun.FontSize = 11
                $resultParagraph.Inlines.Add($resultRun)
                
                $document.Blocks.Add($resultParagraph)
            }
        }
    }
    
    return $document
}

# Fallback-Funktion für einfache Verbindungsaudit-Textanzeige
function Show-SimpleConnectionResults {
    param([string]$Category = "Alle")
    
    Write-DebugLog "Verwende einfache Textanzeige für Verbindungsaudit-Kategorie: $Category" "UI"
    
    # Erstelle einfaches FlowDocument
    $document = New-Object System.Windows.Documents.FlowDocument
    $document.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
    $document.FontSize = 11
    $document.PageWidth = [Double]::NaN
    $document.PageHeight = [Double]::NaN
    $document.ColumnWidth = [Double]::PositiveInfinity
    
    # Sammle alle relevanten Ergebnisse als einfachen Text
    $resultText = ""
    
    # Gruppiere nach Kategorien
    $categorizedResults = @{}
    foreach ($cmd in $connectionAuditCommands) {
        $cmdCategory = if ($cmd.Category) { $cmd.Category } else { "Allgemein" }
        if (-not $categorizedResults.ContainsKey($cmdCategory)) {
            $categorizedResults[$cmdCategory] = @()
        }
        if ($global:connectionAuditResults.ContainsKey($cmd.Name)) {
            $categorizedResults[$cmdCategory] += @{
                Name = $cmd.Name
                Result = $global:connectionAuditResults[$cmd.Name]
            }
        }
    }
    
    # Bestimme anzuzeigende Kategorien
    $categoriesToShow = if ($Category -eq "Alle") { 
        $categorizedResults.Keys | Sort-Object 
    } else { 
        @($Category) 
    }
    
    $totalItems = 0
    foreach ($cat in $categoriesToShow) {
        if ($categorizedResults.ContainsKey($cat)) {
            $categoryData = $categorizedResults[$cat]
            $totalItems += $categoryData.Count
            
            $resultText += "`n" + "="*60 + "`n"
            $resultText += "VERBINDUNGSAUDIT KATEGORIE: $cat`n"
            $resultText += "="*60 + "`n`n"
            
            foreach ($item in $categoryData) {
                $resultText += "-"*40 + "`n"
                $resultText += "EINTRAG: $($item.Name)`n"
                $resultText += "-"*40 + "`n"
                $resultText += "$($item.Result)`n`n"
            }
        }
    }
    
    # Erstelle einfachen Paragraph mit dem gesamten Text
    $paragraph = New-Object System.Windows.Documents.Paragraph
    $run = New-Object System.Windows.Documents.Run($resultText)
    $paragraph.Inlines.Add($run)
    $document.Blocks.Add($paragraph)
    
    $rtbConnectionResults.Document = $document
}
# Funktion zum HTML-Export der Verbindungsaudit-Ergebnisse
function Export-ConnectionAuditToHTML {
    param(
        [hashtable]$Results,
        [string]$FilePath
    )
    
    Write-DebugLog "Starte Verbindungsaudit HTML-Export nach: $FilePath" "Export"

    # Helper to replace Umlaute and escape HTML
    function Convert-ToDisplayString {
        param([string]$Text)
        if ([string]::IsNullOrEmpty($Text)) { return "" }
        $processedText = $Text -replace 'ä', 'ae' -replace 'ö', 'oe' -replace 'ü', 'ue' -replace 'Ä', 'Ae' -replace 'Ö', 'Oe' -replace 'Ü', 'Ue' -replace 'ß', 'ss'
        return [System.Security.SecurityElement]::Escape($processedText)
    }
    
    # Erweiterte Serverinformationen sammeln
    $osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $cpuInfoObj = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
    $totalRamBytes = (Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue | Measure-Object -Property Capacity -Sum).Sum
    
    $serverInfo = @{
        ServerName = $env:COMPUTERNAME
        ReportDate = Get-Date -Format "dd.MM.yyyy | HH:mm:ss"
        Domain = $env:USERDOMAIN
        User = $env:USERNAME
        OS = if ($osInfo) { "$($osInfo.Caption) $($osInfo.OSArchitecture)" } else { "N/A" }
        CPU = if ($cpuInfoObj) { $cpuInfoObj.Name } else { "N/A" }
        RAM = if ($totalRamBytes) { "{0:N2} GB" -f ($totalRamBytes / 1GB) } else { "N/A" }
    }

    # Gruppiere Ergebnisse nach Kategorien
    $groupedResults = @{}
    
    if ($null -ne $connectionAuditCommands) {
        foreach ($cmdDef in $connectionAuditCommands) {
            $categoryName = if ($cmdDef.Category) { $cmdDef.Category } else { "Allgemein" }
            
            if ($Results.ContainsKey($cmdDef.Name)) {
                if (-not $groupedResults.ContainsKey($categoryName)) {
                    $groupedResults[$categoryName] = @()
                }
                
                $groupedResults[$categoryName] += @{
                    Name = $cmdDef.Name
                    Result = $Results[$cmdDef.Name]
                    Command = $cmdDef
                }
            }
        }
    }

    # Navigationselemente und Tab-Inhalte generieren
    $sidebarNavLinks = ""
    $mainContentTabs = ""
    $firstTabId = $null
    
    # Erstelle eine sortierte Liste der Kategorien
    $sortedCategories = $groupedResults.Keys | Sort-Object
    
    foreach ($categoryKey in $sortedCategories) {
        $items = $groupedResults[$categoryKey]
        $displayCategory = Convert-ToDisplayString $categoryKey
        
        # Spezielle Anzeigenamen für Verbindungsaudit-Kategorien
        $displayCategoryName = switch ($categoryKey) {
            "TCP-Connections" { "TCP-Verbindungen" }
            "UDP-Connections" { "UDP-Verbindungen" }
            "External-Connections" { "Externe Verbindungen" }
            "Process-Network" { "Prozess-Netzwerk-Zuordnung" }
            "Local-Devices" { "Lokale Geraete" }
            "DNS-Info" { "DNS-Informationen" }
            "Geo-IP" { "Geo-IP Informationen" }
            "Firewall-Logs" { "Firewall-Protokolle" }
            "Network-Events" { "Netzwerk-Ereignisse" }
            "Domain-Users" { "Domaenen-Benutzer" }
            "Remote-Sessions" { "Remote-Sitzungen" }
            "Network-Topology" { "Netzwerk-Topologie" }
            default { $displayCategory }
        }
        
        $categoryIdPart = $categoryKey -replace '[^a-zA-Z0-9_]', ''
        if ($categoryIdPart.Length -eq 0) { 
            $categoryIdPart = "conncat" + ($categoryKey.GetHashCode() | ForEach-Object ToString X) 
        }
        $tabId = "tab_$categoryIdPart"

        if ($null -eq $firstTabId) { $firstTabId = $tabId }

        $sidebarNavLinks += @"
<li class="nav-item category-nav">
    <a href="#" class="nav-link" onclick="showTab('$tabId', this)">
        $displayCategoryName ($($items.Count))
    </a>
</li>
"@
        
        $tabContent = "<div id='$tabId' class='tab-content'>"
        $tabContent += "<h2 class='content-category-title'>$displayCategoryName</h2>"

        foreach ($item in $items) {
            $displayItemName = Convert-ToDisplayString $item.Name
            $displayItemResult = Convert-ToDisplayString $item.Result
            
            # Füge Kommando-Information hinzu
            $commandInfo = ""
            if ($item.Command -and $item.Command.Command) {
                $commandInfo = "<p class='command-info'><strong>Befehl:</strong> <code>$(Convert-ToDisplayString $item.Command.Command)</code></p>"
            }
            
            $tabContent += @"
<div class="section">
    <div class="section-header">
        <h3 class="section-title">$displayItemName</h3>
    </div>
    <div class="section-content">
        $commandInfo
        <pre>$displayItemResult</pre>
    </div>
</div>
"@
        }
        $tabContent += "</div>"
        $mainContentTabs += $tabContent
    }

    # HTML-Gesamtstruktur (angepasst für Verbindungsaudit mit orangem Design)
    $htmlOutput = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <title>$(Convert-ToDisplayString "Verbindungsaudit Bericht - $($serverInfo.ServerName)")</title>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            margin: 0; 
            background-color: #fff8e7;
            color: #333;
            line-height: 1.6;
        }
        .page-container {
            display: flex;
            flex-direction: column;
            min-height: 100vh;
        }
        .header {
            background: linear-gradient(135deg, #FD7E14, #E85D04);
            color: white;
            padding: 20px 40px;
            display: flex;
            flex-direction: column;
            align-items: center;
        }
        .header-top-row {
            display: flex;
            align-items: center;
            width: 100%;
            margin-bottom: 15px;
        }
        .header-logo {
            width: 125px;
            height: 75px;
            background-color: #e0e0e0;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: bold;
            color: #333;
            margin-right: 20px;
            flex-shrink: 0;
            border-radius: 4px;
        }
        .header-title { 
            margin: 0; 
            font-size: 2em; 
            font-weight: 500;
        }
        
        .header-info-cards-container {
            display: flex;
            flex-wrap: wrap;
            justify-content: center;
            gap: 15px;
            padding: 10px 0;
            width: 100%;
            max-width: 1200px;
        }
        .info-card {
            background-color: rgba(255, 255, 255, 0.1);
            border: 1px solid rgba(255, 255, 255, 0.2);
            border-radius: 6px;
            padding: 10px 15px;
            font-size: 0.85em;
            color: white;
            min-width: 150px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .info-card strong {
            display: block;
            font-size: 0.9em;
            color: #ffe0b3;
            margin-bottom: 3px;
        }

        .main-content-wrapper {
            display: flex;
            flex: 1;
            background-color: #fff8e7;
            margin: 0;
        }

        .sidebar {
            width: 280px;
            background-color: #ffffff; 
            padding: 20px;
            border-right: 1px solid #FD7E14;
            overflow-y: auto; 
            box-shadow: 2px 0 5px rgba(253,126,20,0.1);
            flex-shrink: 0;
        }
        .sidebar .nav-list {
            list-style: none;
            padding: 0;
            margin: 0;
        }
        .sidebar .nav-item {
            margin-bottom: 6px;
        }
        .sidebar .category-nav .nav-link {
            display: block;
            padding: 12px 15px;
            text-decoration: none;
            color: #33475b; 
            border-radius: 6px;
            transition: background-color 0.2s ease, color 0.2s ease;
            font-size: 0.95em;
            font-weight: 500;
            word-break: break-word;
            border-left: 3px solid transparent;
        }
        .sidebar .category-nav .nav-link:hover {
            background-color: #ffe0b3;
            color: #FD7E14;
            border-left-color: #FD7E14;
        }
        .sidebar .category-nav .nav-link.active {
            background-color: #FD7E14;
            color: white;
            font-weight: 600;
            border-left-color: #ffffff;
        }

        .content-area {
            flex: 1; 
            padding: 25px 35px; 
            overflow-y: auto;
            background-color: #ffffff; 
        }
        .content-category-title {
            font-size: 1.6em;
            color: #FD7E14;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #ffe0b3;
        }
        .tab-content {
            display: none;
        }
        .tab-content.active {
            display: block;
            animation: fadeIn 0.4s ease-in-out;
        }
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(8px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .section {
            margin-bottom: 30px;
            background: #ffffff; 
            border-radius: 5px;
            border: 1px solid #FFC107; 
            box-shadow: 0 1px 5px rgba(253,126,20,0.1); 
            overflow: hidden;
        }
        .section-header {
            background: #fff3cd; 
            padding: 12px 18px;
            border-bottom: 1px solid #FFC107;
        }
        .section-title {
            font-size: 1.15em; 
            font-weight: 600;
            color: #856404; 
            margin: 0;
        }
        .section-content {
            padding: 18px;
        }
        .command-info {
            background-color: #fff3cd;
            border-left: 4px solid #FFC107;
            padding: 10px 15px;
            margin-bottom: 15px;
            font-size: 0.9em;
        }
        .command-info code {
            background-color: #e9ecef;
            padding: 2px 6px;
            border-radius: 3px;
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            font-size: 0.85em;
            color: #495057;
        }
        pre { 
            background-color: #fffdf5; 
            padding: 12px; 
            border: 1px solid #FFC107; 
            border-radius: 4px;
            white-space: pre-wrap; 
            word-wrap: break-word;
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            font-size: 0.85em; 
            line-height: 1.5;
            overflow-x: auto;
            color: #333;
        }

        .footer-timestamp { 
            color: #505050;
            font-size: 0.8em;
            text-align: center;
            padding: 15px 40px;
            background-color: #ffe0b3;
            border-top: 1px solid #FD7E14;
        }
        .footer-timestamp a {
            color: #FD7E14;
            text-decoration: none;
        }
        .footer-timestamp a:hover {
            text-decoration: underline;
        }
    </style>
    <script>
        function showTab(tabId, clickedElement) {
            var i, contents, navLinks;
            contents = document.querySelectorAll('.tab-content');
            for (i = 0; i < contents.length; i++) {
                contents[i].classList.remove('active');
            }
            
            navLinks = document.querySelectorAll('.sidebar .category-nav .nav-link');
            for (i = 0; i < navLinks.length; i++) {
                navLinks[i].classList.remove('active');
            }
            
            var selectedTabContent = document.getElementById(tabId);
            if (selectedTabContent) {
                selectedTabContent.classList.add('active');
            }
            
            if (clickedElement) {
                clickedElement.classList.add('active');
            }
        }
        
        window.onload = function() {
            var firstNavLink = document.querySelector('.sidebar .nav-list .category-nav .nav-link');
            if (firstNavLink) {
                firstNavLink.click(); 
            } else {
                var firstContent = document.querySelector('.tab-content');
                if (firstContent) {
                    firstContent.classList.add('active');
                }
            }
        }
    </script>
</head>
<body>
    <div class="page-container">
        <header class="header">
            <div class="header-top-row">
                <div class="header-logo">LOGO</div>
                <h1 class="header-title">$(Convert-ToDisplayString "Verbindungsaudit Bericht")</h1>
            </div>
            <div class="header-info-cards-container">
                <div class="info-card"><strong>Hostname:</strong> $(Convert-ToDisplayString $serverInfo.ServerName)</div>
                <div class="info-card"><strong>$(Convert-ToDisplayString "Domäne"):</strong> $(Convert-ToDisplayString $serverInfo.Domain)</div>
                <div class="info-card"><strong>$(Convert-ToDisplayString "Betriebssystem"):</strong> $(Convert-ToDisplayString $serverInfo.OS)</div>
                <div class="info-card"><strong>CPU:</strong> $(Convert-ToDisplayString $serverInfo.CPU)</div>
                <div class="info-card"><strong>RAM:</strong> $(Convert-ToDisplayString $serverInfo.RAM)</div>
                <div class="info-card"><strong>$(Convert-ToDisplayString "Berichtsdatum"):</strong> $($serverInfo.ReportDate)</div>
                <div class="info-card"><strong>$(Convert-ToDisplayString "Benutzer"):</strong> $(Convert-ToDisplayString $serverInfo.User)</div>
            </div>
        </header>
        
        <div class="main-content-wrapper">
            <nav class="sidebar">
                <ul class="nav-list">
                    $sidebarNavLinks
                </ul>
            </nav>
            <main class="content-area">
                $mainContentTabs
            </main>
        </div>
        
        <footer class="footer-timestamp">
            $(Convert-ToDisplayString "Verbindungsaudit Bericht erstellt von easyWSAudit am $($serverInfo.ReportDate)") | <a href="https://psscripts.de" target="_blank">PSscripts.de</a> | Andreas Hepp
        </footer>
    </div>
</body>
</html>
"@

    $htmlOutput | Out-File -FilePath $FilePath -Encoding utf8
    Write-DebugLog "Verbindungsaudit HTML-Export abgeschlossen" "Export"
}

# Funktion zum Generieren des HTML-Exports
function Export-AuditToHTML {
    param(
        [hashtable]$Results,
        [string]$FilePath
    )
    
    Write-DebugLog "Starte HTML-Export nach: $FilePath" "Export"

    # Helper to replace Umlaute and escape HTML
    function Convert-ToDisplayString {
        param([string]$Text)
        if ([string]::IsNullOrEmpty($Text)) { return "" }
        $processedText = $Text -replace 'ä', 'ae' -replace 'ö', 'oe' -replace 'ü', 'ue' -replace 'Ä', 'Ae' -replace 'Ö', 'Oe' -replace 'Ü', 'Ue' -replace 'ß', 'ss'
        # Weitere Sonderzeichen koennten hier bei Bedarf behandelt werden
        return [System.Security.SecurityElement]::Escape($processedText)
    }
    
    # Erweiterte Serverinformationen sammeln
    $osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $cpuInfoObj = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
    $totalRamBytes = (Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue | Measure-Object -Property Capacity -Sum).Sum
    $driveC = Get-Volume -DriveLetter 'C' -ErrorAction SilentlyContinue
    
    $serverInfo = @{
        ServerName = $env:COMPUTERNAME
        ReportDate = Get-Date -Format "dd.MM.yyyy | HH:mm:ss" # Format beibehalten, da es ein Datum ist
        Domain = $env:USERDOMAIN
        User = $env:USERNAME
        OS = if ($osInfo) { "$($osInfo.Caption) $($osInfo.OSArchitecture)" } else { "N/A" }
        CPU = if ($cpuInfoObj) { $cpuInfoObj.Name } else { "N/A" }
        RAM = if ($totalRamBytes) { "{0:N2} GB" -f ($totalRamBytes / 1GB) } else { "N/A" }
        DiskCTotal = if ($driveC) { "{0:N2} GB" -f ($driveC.Size / 1GB) } else { "N/A" }
        DiskCFree = if ($driveC) { "{0:N2} GB" -f ($driveC.SizeRemaining / 1GB) } else { "N/A" }
    }

    # Gruppiere Ergebnisse nach Kategorien (gleiche Logik wie in der GUI)
    $groupedResults = @{}
    
    # Definiere die gewünschte Reihenfolge der Kategorien (System an erster Stelle)
    $categoryOrder = @(
        "System",
        "Hardware", 
        "Storage",
        "Network",
        "Security",
        "Services",
        "Tasks",
        "Events",
        "Features",
        "Software",
        "Updates",
        "Active-Directory",
        "DNS",
        "DHCP",
        "IIS",
        "WDS",
        "Hyper-V",
        "Cluster",
        "WSUS",
        "FileServices",
        "PrintServices",
        "RDS",
        "PKI",
        "ADFS",
        "ADLDS",
        "ADRMS",
        "DeviceAttestation",
        "VolumeActivation",
        "Backup",
        "NPAS",
        "HGS",
        "RemoteAccess",
        "InternalDB",
        "WindowsDefender",
        "WAS",
        "SearchService",
        "ServerEssentials",
        "Migration",
        "Identity",
        "FileSharing",
        "UserProfiles",
        "Firewall",
        "PowerManagement",
        "CredentialManager",
        "AuditPolicy",
        "GroupPolicy",
        "InstalledSoftware",
        "Environment"
    )
    
    if ($null -ne $commands) {
        foreach ($cmdDef in $commands) {
            # Verwende die Category direkt aus dem Befehl
            $categoryName = if ($cmdDef.Category) { $cmdDef.Category } else { "Allgemein" }
            
            if ($Results.ContainsKey($cmdDef.Name)) {
                if (-not $groupedResults.ContainsKey($categoryName)) {
                    $groupedResults[$categoryName] = @()
                }
                
                $groupedResults[$categoryName] += @{
                    Name = $cmdDef.Name
                    Result = $Results[$cmdDef.Name]
                    Command = $cmdDef
                }
            }
        }
    }

    # Navigationselemente und Tab-Inhalte generieren mit gewünschter Reihenfolge
    $sidebarNavLinks = ""
    $mainContentTabs = ""
    $firstTabId = $null
    
    # Erstelle eine sortierte Liste der Kategorien basierend auf der gewünschten Reihenfolge
    $sortedCategories = @()
    
    # Füge Kategorien in der gewünschten Reihenfolge hinzu
    foreach ($orderCat in $categoryOrder) {
        if ($groupedResults.ContainsKey($orderCat)) {
            $sortedCategories += $orderCat
        }
    }
    
    # Füge alle anderen Kategorien alphabetisch sortiert hinzu
    foreach ($categoryKey in ($groupedResults.Keys | Sort-Object)) {
        if ($categoryKey -notin $sortedCategories) {
            $sortedCategories += $categoryKey
        }
    }
    
    foreach ($categoryKey in $sortedCategories) {
        $items = $groupedResults[$categoryKey]
        $displayCategory = Convert-ToDisplayString $categoryKey
        
        # Spezielle Anzeigenamen für bessere Lesbarkeit
        $displayCategoryName = switch ($categoryKey) {
            "System" { "System-Informationen" }
            "Hardware" { "Hardware-Informationen" }
            "Storage" { "Speicher & Festplatten" }
            "Network" { "Netzwerk-Konfiguration" }
            "Security" { "Sicherheits-Einstellungen" }
            "Services" { "Dienste & Services" }
            "Tasks" { "Geplante Aufgaben" }
            "Events" { "Ereignisprotokoll" }
            "Features" { "Installierte Features" }
            "Software" { "Software & Programme" }
            "Updates" { "Windows Updates" }
            "Active-Directory" { "Active Directory" }
            "DNS" { "DNS-Server" }
            "DHCP" { "DHCP-Server" }
            "IIS" { "Internet Information Services (IIS)" }
            "WDS" { "Windows Deployment Services" }
            "Hyper-V" { "Hyper-V Virtualisierung" }
            "Cluster" { "Failover Clustering" }
            "WSUS" { "Windows Server Update Services" }
            "FileServices" { "Datei-Services" }
            "PrintServices" { "Druck-Services" }
            "RDS" { "Remote Desktop Services" }
            "PKI" { "Zertifikat-Services (PKI)" }
            "ADFS" { "Active Directory Federation Services" }
            "ADLDS" { "AD Lightweight Directory Services" }
            "ADRMS" { "AD Rights Management Services" }
            "DeviceAttestation" { "Device Health Attestation" }
            "VolumeActivation" { "Volume Activation Services" }
            "Backup" { "Windows Server Backup" }
            "NPAS" { "Network Policy and Access Services" }
            "HGS" { "Host Guardian Service" }
            "RemoteAccess" { "Remote Access Services" }
            "InternalDB" { "Windows Internal Database" }
            "WindowsDefender" { "Windows Defender" }
            "WAS" { "Windows Process Activation Service" }
            "SearchService" { "Windows Search Service" }
            "ServerEssentials" { "Windows Server Essentials" }
            "Migration" { "Migration Services" }
            "Identity" { "Windows Identity Foundation" }
            "FileSharing" { "Dateifreigaben" }
            "UserProfiles" { "Benutzer-Profile" }
            "Firewall" { "Windows Firewall" }
            "PowerManagement" { "Energieverwaltung" }
            "CredentialManager" { "Anmeldeinformationsverwaltung" }
            "AuditPolicy" { "Audit-Richtlinien" }
            "GroupPolicy" { "Gruppenrichtlinien" }
            "InstalledSoftware" { "Installierte Software" }
            "Environment" { "Umgebungsvariablen" }
            default { $displayCategory }
        }
        
        $categoryIdPart = $categoryKey -replace '[^a-zA-Z0-9_]', ''
        if ($categoryIdPart.Length -eq 0) { 
            $categoryIdPart = "cat" + ($categoryKey.GetHashCode() | ForEach-Object ToString X) 
        }
        $tabId = "tab_$categoryIdPart"

        if ($null -eq $firstTabId) { $firstTabId = $tabId }

        $sidebarNavLinks += @"
<li class="nav-item category-nav">
    <a href="#" class="nav-link" onclick="showTab('$tabId', this)">
        $displayCategoryName ($($items.Count))
    </a>
</li>
"@
        
        $tabContent = "<div id='$tabId' class='tab-content'>"
        $tabContent += "<h2 class='content-category-title'>$displayCategoryName</h2>"

        foreach ($item in $items) {
            $displayItemName = Convert-ToDisplayString $item.Name
            $displayItemResult = Convert-ToDisplayString $item.Result
            
            # Füge Kommando-Information hinzu, falls verfügbar
            $commandInfo = ""
            if ($item.Command -and $item.Command.Command) {
                $commandInfo = "<p class='command-info'><strong>Befehl:</strong> <code>$(Convert-ToDisplayString $item.Command.Command)</code></p>"
            }
            
            $tabContent += @"
<div class="section">
    <div class="section-header">
        <h3 class="section-title">$displayItemName</h3>
    </div>
    <div class="section-content">
        $commandInfo
        <pre>$displayItemResult</pre>
    </div>
</div>
"@
        }
        $tabContent += "</div>"
        $mainContentTabs += $tabContent
    }

    # HTML-Gesamtstruktur
    $htmlOutput = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <title>$(Convert-ToDisplayString "Windows Server Audit Bericht - $($serverInfo.ServerName)")</title>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            margin: 0; 
            background-color: #f0f2f5; /* Etwas hellerer Hintergrund */
            color: #333;
            line-height: 1.6;
        }
        .page-container {
            display: flex;
            flex-direction: column;
            min-height: 100vh;
        }
        .header {
            background: linear-gradient(135deg, #005a9e, #003966); /* Dunkleres Blau */
            color: white;
            padding: 20px 40px;
            display: flex;
            flex-direction: column; /* Ermoeglicht Top-Row und Info-Cards untereinander */
            align-items: center; /* Zentriert Info-Cards, falls sie schmaler sind */
        }
        .header-top-row {
            display: flex;
            align-items: center;
            width: 100%;
            margin-bottom: 15px;
        }
        .header-logo {
            width: 125px;
            height: 75px;
            background-color: #e0e0e0;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: bold;
            color: #333;
            margin-right: 20px;
            flex-shrink: 0;
            border-radius: 4px;
        }
        .header-title { 
            margin: 0; 
            font-size: 2em; 
            font-weight: 500; /* Etwas staerker */
        }
        
        .header-info-cards-container {
            display: flex;
            flex-wrap: wrap;
            justify-content: center; /* Zentriert die Karten */
            gap: 15px; /* Abstand zwischen den Karten */
            padding: 10px 0;
            width: 100%;
            max-width: 1200px; /* Begrenzt die Breite der Kartenreihe */
        }
        .info-card {
            background-color: rgba(255, 255, 255, 0.1); /* Leicht transparente Karten */
            border: 1px solid rgba(255, 255, 255, 0.2);
            border-radius: 6px;
            padding: 10px 15px;
            font-size: 0.85em;
            color: white; /* Textfarbe auf den Karten */
            min-width: 150px; /* Mindestbreite fuer Karten */
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .info-card strong {
            display: block;
            font-size: 0.9em;
            color: #b0dfff; /* Hellere Farbe fuer Label */
            margin-bottom: 3px;
        }

        .main-content-wrapper {
            display: flex;
            flex: 1;
            background-color: #f0f2f5; 
            margin: 0;
        }

        .sidebar {
            width: 280px; /* Etwas breiter fuer tiefere Navigation */
            background-color: #ffffff; 
            padding: 20px;
            border-right: 1px solid #d8dde3;
            overflow-y: auto; 
            box-shadow: 2px 0 5px rgba(0,0,0,0.05);
            flex-shrink: 0;
        }
        .sidebar .nav-list {
            list-style: none;
            padding: 0;
            margin: 0;
        }
        .sidebar .nav-item {
            margin-bottom: 6px;
        }
        .sidebar .category-nav .nav-link {
            display: block;
            padding: 12px 15px;
            text-decoration: none;
            color: #33475b; 
            border-radius: 6px;
            transition: background-color 0.2s ease, color 0.2s ease;
            font-size: 0.95em;
            font-weight: 500;
            word-break: break-word;
            border-left: 3px solid transparent;
        }
        .sidebar .category-nav .nav-link:hover {
            background-color: #e9ecef;
            color: #005a9e;
            border-left-color: #005a9e;
        }
        .sidebar .category-nav .nav-link.active {
            background-color: #0078d4;
            color: white;
            font-weight: 600;
            border-left-color: #ffffff;
        }

        .content-area {
            flex: 1; 
            padding: 25px 35px; 
            overflow-y: auto;
            background-color: #ffffff; 
        }
        .content-category-title { /* Stil fuer den Titel im Inhaltsbereich */
            font-size: 1.6em;
            color: #005a9e;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #eef1f5;
        }
        .tab-content {
            display: none;
        }
        .tab-content.active {
            display: block;
            animation: fadeIn 0.4s ease-in-out;
        }
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(8px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .section {
            margin-bottom: 30px;
            background: #ffffff; 
            border-radius: 5px;
            border: 1px solid #e7eaf0; 
            box-shadow: 0 1px 5px rgba(0,0,0,0.05); 
            overflow: hidden;
        }
        .section-header {
            background: #f7f9fc; 
            padding: 12px 18px; /* Etwas kompakter */
            border-bottom: 1px solid #e7eaf0;
        }
        .section-title {
            font-size: 1.15em; 
            font-weight: 600;
            color: #2c3e50; 
            margin: 0;
        }
        .section-content {
            padding: 18px;
        }
        .command-info {
            background-color: #f8f9fa;
            border-left: 4px solid #28a745;
            padding: 10px 15px;
            margin-bottom: 15px;
            font-size: 0.9em;
        }
        .command-info code {
            background-color: #e9ecef;
            padding: 2px 6px;
            border-radius: 3px;
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            font-size: 0.85em;
            color: #495057;
        }
        pre { 
            background-color: #fdfdff; 
            padding: 12px; 
            border: 1px solid #e0e4e9; 
            border-radius: 4px;
            white-space: pre-wrap; 
            word-wrap: break-word;
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            font-size: 0.85em; 
            line-height: 1.5;
            overflow-x: auto;
            color: #333;
        }

        .footer-timestamp { 
            color: #505050; /* Dunklerer Text */
            font-size: 0.8em;
            text-align: center;
            padding: 15px 40px;
            background-color: #e9ecef; /* Passend zum Rest */
            border-top: 1px solid #d8dde3;
        }
        .footer-timestamp a {
            color: #005a9e;
            text-decoration: none;
        }
        .footer-timestamp a:hover {
            text-decoration: underline;
        }
    </style>
    <script>
        function showTab(tabId, clickedElement) {
            var i, contents, navLinks;
            contents = document.querySelectorAll('.tab-content');
            for (i = 0; i < contents.length; i++) {
                contents[i].classList.remove('active');
            }
            
            navLinks = document.querySelectorAll('.sidebar .category-nav .nav-link');
            for (i = 0; i < navLinks.length; i++) {
                navLinks[i].classList.remove('active');
            }
            
            var selectedTabContent = document.getElementById(tabId);
            if (selectedTabContent) {
                selectedTabContent.classList.add('active');
            }
            
            if (clickedElement) {
                clickedElement.classList.add('active');
            }
        }
        
        window.onload = function() {
            // Ersten Kategorie-Link automatisch aktivieren
            var firstNavLink = document.querySelector('.sidebar .nav-list .category-nav .nav-link');
            if (firstNavLink) {
                firstNavLink.click(); 
            } else {
                // Fallback, falls keine Links vorhanden sind
                var firstContent = document.querySelector('.tab-content');
                if (firstContent) {
                    firstContent.classList.add('active');
                }
            }
        }
    </script>
</head>
<body>
    <div class="page-container">
        <header class="header">
            <div class="header-top-row">
                <div class="header-logo">LOGO</div>
                <h1 class="header-title">$(Convert-ToDisplayString "Windows Server Audit Bericht")</h1>
            </div>
            <div class="header-info-cards-container">
                <div class="info-card"><strong>Hostname:</strong> $(Convert-ToDisplayString $serverInfo.ServerName)</div>
                <div class="info-card"><strong>$(Convert-ToDisplayString "Domäne"):</strong> $(Convert-ToDisplayString $serverInfo.Domain)</div>
                <div class="info-card"><strong>$(Convert-ToDisplayString "Betriebssystem"):</strong> $(Convert-ToDisplayString $serverInfo.OS)</div>
                <div class="info-card"><strong>CPU:</strong> $(Convert-ToDisplayString $serverInfo.CPU)</div>
                <div class="info-card"><strong>RAM:</strong> $(Convert-ToDisplayString $serverInfo.RAM)</div>
                <div class="info-card"><strong>$(Convert-ToDisplayString "Festplatte C: Gesamt"):</strong> $(Convert-ToDisplayString $serverInfo.DiskCTotal)</div>
                <div class="info-card"><strong>$(Convert-ToDisplayString "Festplatte C: Frei"):</strong> $(Convert-ToDisplayString $serverInfo.DiskCFree)</div>
                <div class="info-card"><strong>$(Convert-ToDisplayString "Berichtsdatum"):</strong> $($serverInfo.ReportDate)</div>
                <div class="info-card"><strong>$(Convert-ToDisplayString "Benutzer"):</strong> $(Convert-ToDisplayString $serverInfo.User)</div>
            </div>
        </header>
        
        <div class="main-content-wrapper">
            <nav class="sidebar">
                <ul class="nav-list">
                    $sidebarNavLinks
                </ul>
            </nav>
            <main class="content-area">
                $mainContentTabs
            </main>
        </div>
        
        <footer class="footer-timestamp">
            $(Convert-ToDisplayString "Audit Bericht erstellt von easyWSAudit am $($serverInfo.ReportDate)") | <a href="https://psscripts.de" target="_blank">PSscripts.de</a> | Andreas Hepp
        </footer>
    </div>
</body>
</html>
"@

    $htmlOutput | Out-File -FilePath $FilePath -Encoding utf8
    Write-DebugLog "HTML-Export abgeschlossen" "Export"
}

# Variable fuer die Audit-Ergebnisse
$global:auditResults = @{}

# XAML UI Definition - Vereinfachte Version
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="easyWSAudit - Windows Server Audit Tool"
    Height="900" Width="1400" WindowStartupLocation="CenterScreen"
    Background="#F5F5F5">
    <Window.Resources>
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Padding" Value="15,10"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                CornerRadius="4" 
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#106EBE"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#CCCCCC"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style x:Key="CategoryHeader" TargetType="TextBlock">
            <Setter Property="FontSize" Value="16"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Foreground" Value="#0078D4"/>
            <Setter Property="Margin" Value="0,15,0,5"/>
        </Style>
        
        <Style x:Key="CheckboxStyle" TargetType="CheckBox">
            <Setter Property="Margin" Value="20,3,5,3"/>
            <Setter Property="Padding" Value="5,0,0,0"/>
        </Style>
    </Window.Resources>
    
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#0078D4" Padding="20">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                    <TextBlock Text="easyWSAudit" FontSize="28" Foreground="White" FontWeight="Light"/>
                    <TextBlock Text="Windows Server Audit Tool" FontSize="14" Foreground="#CCE7FF" Margin="0,5,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <TextBlock x:Name="txtServerName" Text="" FontSize="14" Foreground="White" VerticalAlignment="Center" Margin="0,0,20,0"/>
                    <Button Content="Vollstaendiges Audit" x:Name="btnFullAudit" Style="{StaticResource ModernButton}" Background="#28A745"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- Hauptinhalt -->
        <TabControl Grid.Row="1" Margin="0" Background="Transparent" BorderThickness="0">
            <TabItem Header="Audit Konfiguration" FontSize="14">
                <Grid Margin="20">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="400"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    
                    <!-- Linke Seite - Optionen -->
                    <Border Grid.Column="0" Background="White" CornerRadius="8" Padding="20" Margin="0,0,10,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <TextBlock Grid.Row="0" Text="Audit-Kategorien auswaehlen" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/>
                            
                            <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="0,0,0,15">
                                <StackPanel x:Name="spOptions"/>
                            </ScrollViewer>
                            
                            <StackPanel Grid.Row="2">
                                <Button Content="Alle auswaehlen" x:Name="btnSelectAll" Style="{StaticResource ModernButton}" Background="#28A745" Margin="0,0,0,5"/>
                                <Button Content="Alle abwaehlen" x:Name="btnSelectNone" Style="{StaticResource ModernButton}" Background="#DC3545" Margin="0,0,0,25"/>
                                <Button Content="Audit starten" x:Name="btnRunAudit" Style="{StaticResource ModernButton}" Background="#FFC107" Foreground="Black" FontWeight="Bold"/>
                            </StackPanel>
                        </Grid>
                    </Border>
                    
                    <!-- Rechte Seite - Fortschritt -->
                    <Border Grid.Column="1" Background="White" CornerRadius="8" Padding="20" Margin="10,0,0,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            
                            <TextBlock Grid.Row="0" Text="Audit-Fortschritt" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/>
                            
                            <StackPanel Grid.Row="1" Margin="0,0,0,15">
                                <ProgressBar x:Name="progressBar" Height="20" Margin="0,0,0,10"/>
                                <TextBlock x:Name="txtProgress" Text="Bereit fuer Audit" HorizontalAlignment="Center" FontSize="12" Foreground="#666"/>
                            </StackPanel>
                            
                            <Border Grid.Row="2" Background="#F8F9FA" CornerRadius="4" Padding="15">
                                <ScrollViewer VerticalScrollBarVisibility="Auto">
                                    <TextBlock x:Name="txtStatusLog" Text="Bereit..." FontFamily="Consolas" FontSize="11" Foreground="#495057"/>
                                </ScrollViewer>
                            </Border>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>
            
            <TabItem Header="Audit Ergebnisse" FontSize="14">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Toolbar mit Buttons -->
                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,15">
                        <Button Content="Als HTML exportieren" x:Name="btnExportHTML" Style="{StaticResource ModernButton}" Background="#17A2B8" IsEnabled="False"/>
                        <Button Content="In Zwischenablage" x:Name="btnCopyToClipboard" Style="{StaticResource ModernButton}" Background="#6C757D" IsEnabled="False" Margin="10,0,0,0"/>
                        <Button Content="Ergebnisse aktualisieren" x:Name="btnRefreshResults" Style="{StaticResource ModernButton}" Background="#28A745" IsEnabled="False" Margin="10,0,0,0"/>
                    </StackPanel>
                    
                    <!-- Kategorien-Auswahl -->
                    <Grid Grid.Row="1" Margin="0,0,0,15">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="250"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        
                        <TextBlock Grid.Column="0" Text="Kategorie:" VerticalAlignment="Center" Margin="0,0,10,0" FontWeight="SemiBold"/>
                        <ComboBox x:Name="cmbResultCategories" Grid.Column="1" Height="30" VerticalAlignment="Center"/>
                        <TextBlock x:Name="txtResultsSummary" Grid.Column="3" Text="" VerticalAlignment="Center" FontSize="12" Foreground="#6C757D"/>
                    </Grid>
                    
                    <!-- Ergebnisse-Anzeige -->
                    <Border Grid.Row="2" Background="White" CornerRadius="8" BorderThickness="1" BorderBrush="#DEE2E6">
                        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" 
                                     Padding="0" Margin="0">
                            <RichTextBox x:Name="rtbResults" Background="Transparent" BorderThickness="0" 
                                        IsReadOnly="True" Padding="20" FontFamily="Segoe UI" FontSize="12"
                                        HorizontalAlignment="Stretch" VerticalAlignment="Stretch"/>
                        </ScrollViewer>
                    </Border>
                </Grid>
            </TabItem>
            
            <TabItem Header="Verbindungsaudit" FontSize="14">
                <Grid Margin="20">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="400"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    
                    <!-- Linke Seite - Verbindungsaudit Optionen -->
                    <Border Grid.Column="0" Background="White" CornerRadius="8" Padding="20" Margin="0,0,10,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <StackPanel Grid.Row="0">
                                <TextBlock Text="Verbindungsaudit-Kategorien" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                <TextBlock Text="Analyse aktiver Netzwerkverbindungen, Geraete und Benutzer" FontSize="12" Foreground="#6C757D" TextWrapping="Wrap" Margin="0,0,0,15"/>
                            </StackPanel>
                            
                            <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="0,0,0,15">
                                <StackPanel x:Name="spConnectionOptions"/>
                            </ScrollViewer>
                            
                            <StackPanel Grid.Row="2">
                                <Button Content="Alle auswaehlen" x:Name="btnSelectAllConnection" Style="{StaticResource ModernButton}" Background="#28A745" Margin="0,0,0,5"/>
                                <Button Content="Alle abwaehlen" x:Name="btnSelectNoneConnection" Style="{StaticResource ModernButton}" Background="#DC3545" Margin="0,0,0,25"/>
                                <Button Content="Verbindungsaudit starten" x:Name="btnRunConnectionAudit" Style="{StaticResource ModernButton}" Background="#FD7E14" Foreground="White" FontWeight="Bold"/>
                            </StackPanel>
                        </Grid>
                    </Border>
                    
                    <!-- Rechte Seite - Verbindungsaudit Ergebnisse -->
                    <Border Grid.Column="1" Background="White" CornerRadius="8" Padding="20" Margin="10,0,0,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            
                            <TextBlock Grid.Row="0" Text="Verbindungsaudit-Ergebnisse" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/>
                            
                            <!-- Fortschritt fuer Verbindungsaudit -->
                            <StackPanel Grid.Row="1" Margin="0,0,0,15">
                                <ProgressBar x:Name="progressBarConnection" Height="20" Margin="0,0,0,10"/>
                                <TextBlock x:Name="txtProgressConnection" Text="Bereit fuer Verbindungsaudit" HorizontalAlignment="Center" FontSize="12" Foreground="#666"/>
                            </StackPanel>
                            
                            <!-- Toolbar fuer Verbindungsaudit -->
                            <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,15">
                                <Button Content="Export HTML" x:Name="btnExportConnectionHTML" Style="{StaticResource ModernButton}" Background="#17A2B8" IsEnabled="False"/>
                                <Button Content="In Zwischenablage" x:Name="btnCopyConnectionToClipboard" Style="{StaticResource ModernButton}" Background="#6C757D" IsEnabled="False" Margin="10,0,0,0"/>
                                <ComboBox x:Name="cmbConnectionCategories" Width="200" Height="30" VerticalAlignment="Center" Margin="20,0,0,0" IsEnabled="False"/>
                            </StackPanel>
                            
                            <!-- Ergebnisse-Anzeige fuer Verbindungsaudit -->
                            <Border Grid.Row="3" Background="#F8F9FA" CornerRadius="4" BorderThickness="1" BorderBrush="#DEE2E6">
                                <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" 
                                             Padding="0" Margin="0">
                                    <RichTextBox x:Name="rtbConnectionResults" Background="Transparent" BorderThickness="0" 
                                                IsReadOnly="True" Padding="20" FontFamily="Segoe UI" FontSize="12"
                                                HorizontalAlignment="Stretch" VerticalAlignment="Stretch"/>
                                </ScrollViewer>
                            </Border>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>
            
            <TabItem Header="Debug" FontSize="14">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Grid.Row="0" Text="Debug-Informationen" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/>
                    
                    <Border Grid.Row="1" Background="#1E1E1E" CornerRadius="8" BorderThickness="1" BorderBrush="#333333">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <TextBox x:Name="txtDebugOutput" Background="Transparent" BorderThickness="0" 
                                     FontFamily="Consolas" IsReadOnly="True" TextWrapping="Wrap"
                                     Foreground="#00FF00" Padding="15" FontSize="11"/>
                        </ScrollViewer>
                    </Border>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,15,0,0" HorizontalAlignment="Left">
                        <Button Content="Log-Datei oeffnen" x:Name="btnOpenLog" Style="{StaticResource ModernButton}"/>
                        <Button Content="Log leeren" x:Name="btnClearLog" Style="{StaticResource ModernButton}" Background="#DC3545" Margin="10,0,0,0"/>
                    </StackPanel>
                </Grid>
            </TabItem>
        </TabControl>
        
        <!-- Footer -->
        <Border Grid.Row="2" Background="#F8F9FA" BorderThickness="0,1,0,0" BorderBrush="#DEE2E6" Padding="20,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="txtStatus" Text="Status: Bereit" VerticalAlignment="Center" FontSize="12" Foreground="#6C757D"/>
                <TextBlock Grid.Column="1" Text="PSscripts.de | easyWSAudit v0.0.1 - 23.05.2025  |  Andreas Hepp" VerticalAlignment="Center" FontSize="12" Foreground="#6C757D"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Lade das XAML
Write-DebugLog "Lade XAML fuer UI..." "UI"
try {
    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    
    if ($null -eq $window) {
        throw "Window konnte nicht erstellt werden"
    }
    
    Write-DebugLog "XAML erfolgreich geladen, Fenster erstellt" "UI"
} catch {
    Write-DebugLog "FEHLER beim Laden des XAML: $($_.Exception.Message)" "UI"
    Write-Host "KRITISCHER FEHLER beim Laden der UI: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Details: $($_.Exception.ToString())" -ForegroundColor Red
    
    # Zeige eine Fehlermeldung und beende das Script
    if ($_.Exception.InnerException) {
        Write-Host "Innere Exception: $($_.Exception.InnerException.Message)" -ForegroundColor Red
    }
    
    Read-Host "Drücken Sie eine Taste zum Beenden"
    return
}

# Hole die UI-Elemente
Write-DebugLog "Suche UI-Elemente..." "UI"
$txtServerName = $window.FindName("txtServerName")
$btnFullAudit = $window.FindName("btnFullAudit")
$spOptions = $window.FindName("spOptions")
$btnSelectAll = $window.FindName("btnSelectAll")
$btnSelectNone = $window.FindName("btnSelectNone")
$btnRunAudit = $window.FindName("btnRunAudit")
$progressBar = $window.FindName("progressBar")
$txtProgress = $window.FindName("txtProgress")
$txtStatusLog = $window.FindName("txtStatusLog")
$cmbResultCategories = $window.FindName("cmbResultCategories")
$rtbResults = $window.FindName("rtbResults")
$txtResultsSummary = $window.FindName("txtResultsSummary")
$btnExportHTML = $window.FindName("btnExportHTML")
$btnCopyToClipboard = $window.FindName("btnCopyToClipboard")
$btnRefreshResults = $window.FindName("btnRefreshResults")
$txtStatus = $window.FindName("txtStatus")

# Verbindungsaudit-Elemente
$spConnectionOptions = $window.FindName("spConnectionOptions")
$btnSelectAllConnection = $window.FindName("btnSelectAllConnection")
$btnSelectNoneConnection = $window.FindName("btnSelectNoneConnection")
$btnRunConnectionAudit = $window.FindName("btnRunConnectionAudit")
$progressBarConnection = $window.FindName("progressBarConnection")
$txtProgressConnection = $window.FindName("txtProgressConnection")
$btnExportConnectionHTML = $window.FindName("btnExportConnectionHTML")
$btnCopyConnectionToClipboard = $window.FindName("btnCopyConnectionToClipboard")
$cmbConnectionCategories = $window.FindName("cmbConnectionCategories")
$rtbConnectionResults = $window.FindName("rtbConnectionResults")

# Debug-Elemente
$script:txtDebugOutput = $window.FindName("txtDebugOutput")
$btnOpenLog = $window.FindName("btnOpenLog")
$btnClearLog = $window.FindName("btnClearLog")

# Überprüfe kritische UI-Elemente
Write-DebugLog "Überprüfe UI-Elemente..." "UI"
if ($null -eq $window) { Write-DebugLog "FEHLER: window ist NULL!" "UI"; return }
if ($null -eq $spOptions) { Write-DebugLog "FEHLER: spOptions ist NULL!" "UI"; return }
if ($null -eq $txtServerName) { Write-DebugLog "FEHLER: txtServerName ist NULL!" "UI"; return }
if ($null -eq $spConnectionOptions) { Write-DebugLog "FEHLER: spConnectionOptions ist NULL!" "UI"; return }

Write-DebugLog "UI-Elemente erfolgreich initialisiert" "UI"

# Servername anzeigen
$txtServerName.Text = "Server: $env:COMPUTERNAME"

# Dictionary fuer Checkboxen (beide Audits)
$checkboxes = @{}
$connectionCheckboxes = @{}

# Erstelle die Checkboxen fuer die Verbindungsaudit-Optionen
Write-DebugLog "Erstelle Checkboxen fuer Verbindungsaudit-Optionen..." "UI"

$connectionCategories = @{}
foreach ($cmd in $connectionAuditCommands) {
    $categoryName = if ($cmd.Category) { $cmd.Category } else { "Allgemein" }
    if (-not $connectionCategories.ContainsKey($categoryName)) {
        $connectionCategories[$categoryName] = @()
    }
    $connectionCategories[$categoryName] += $cmd
}

# Iteriere über die Verbindungsaudit-Kategorien in alphabetischer Reihenfolge
foreach ($categoryKey in ($connectionCategories.Keys | Sort-Object)) {
    # Kategorie-Header
    $categoryHeader = New-Object System.Windows.Controls.TextBlock
    $categoryHeader.Text = "$categoryKey"
    
    # Setze Style direkt
    $categoryHeader.FontSize = 16
    $categoryHeader.FontWeight = "Bold"
    $categoryHeader.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(253, 126, 20)) # Orange für Verbindungsaudit
    $categoryHeader.Margin = New-Object System.Windows.Thickness(0, 15, 0, 5)
    
    try {
        $spConnectionOptions.Children.Add($categoryHeader)
        Write-DebugLog "Verbindungsaudit Kategorie-Header '$categoryKey' hinzugefügt" "UI"
    } catch {
        Write-DebugLog "FEHLER beim Hinzufügen des Verbindungsaudit Kategorie-Headers '$categoryKey': $($_.Exception.Message)" "UI"
        continue
    }
    
    # Checkboxen fuer diese Kategorie
    foreach ($cmd in $connectionCategories[$categoryKey]) {
        $checkbox = New-Object System.Windows.Controls.CheckBox
        $checkbox.Content = $cmd.Name
        $checkbox.IsChecked = $true # Standardmäßig aktiviert
        
        # Setze Style direkt
        $checkbox.Margin = New-Object System.Windows.Thickness(20, 3, 5, 3)
        $checkbox.Padding = New-Object System.Windows.Thickness(5, 0, 0, 0)
        
        # Ueberpruefe, ob diese Option mit einer Serverrolle verbunden ist
        if ($cmd.ContainsKey("FeatureName")) {
            $isRoleInstalled = Test-ServerRole -FeatureName $cmd.FeatureName
            if (-not $isRoleInstalled) {
                $checkbox.IsEnabled = $false
                $checkbox.Content = "$($cmd.Name) (Nicht installiert)"
                $checkbox.IsChecked = $false
            }
        }
        
        try {
            $spConnectionOptions.Children.Add($checkbox)
            $connectionCheckboxes[$cmd.Name] = $checkbox
            Write-DebugLog "Verbindungsaudit Checkbox '$($cmd.Name)' hinzugefügt (Kategorie: $categoryKey)" "UI"
        } catch {
            Write-DebugLog "FEHLER beim Hinzufügen der Verbindungsaudit Checkbox '$($cmd.Name)' (Kategorie: $categoryKey): $($_.Exception.Message)" "UI"
        }
    }
}
Write-DebugLog "Verbindungsaudit Checkboxen erstellt für $($connectionCheckboxes.Count) Optionen" "UI"

# Dictionary fuer Checkboxen
$checkboxes = @{}

# Erstelle die Checkboxen fuer die Audit-Optionen gruppiert nach Kategorien
Write-DebugLog "Erstelle Checkboxen fuer Audit-Optionen..." "UI"

# Überprüfe, ob spOptions verfügbar ist
if ($null -eq $spOptions) {
    Write-DebugLog "FEHLER: spOptions nicht gefunden - kann keine Checkboxen erstellen!" "UI"
    [System.Windows.MessageBox]::Show("UI-Initialisierungsfehler: spOptions nicht gefunden.", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    return
}

$categories = @{} # Verwende ein normales Hashtable für bessere Kompatibilität
foreach ($cmd in $commands) {
    $categoryName = if ($cmd.Category) { $cmd.Category } else { "Allgemein" }
    if (-not $categories.ContainsKey($categoryName)) {
        # Initialisiere den Eintrag für eine neue Kategorie mit einer Liste.
        $categories[$categoryName] = @()
    }
    # Füge den Befehl zur Liste der entsprechenden Kategorie hinzu
    $categories[$categoryName] += $cmd
}

# Iteriere über die Kategorien in alphabetischer Reihenfolge
foreach ($categoryKey in ($categories.Keys | Sort-Object)) {
    # Kategorie-Header
    $categoryHeader = New-Object System.Windows.Controls.TextBlock
    $categoryHeader.Text = "$categoryKey"
    
    # Setze Style direkt statt über FindResource
    $categoryHeader.FontSize = 16
    $categoryHeader.FontWeight = "Bold"
    $categoryHeader.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 120, 212))
    $categoryHeader.Margin = New-Object System.Windows.Thickness(0, 15, 0, 5)
    
    try {
        $spOptions.Children.Add($categoryHeader)
        Write-DebugLog "Kategorie-Header '$categoryKey' hinzugefügt" "UI"
    } catch {
        Write-DebugLog "FEHLER beim Hinzufügen des Kategorie-Headers '$categoryKey': $($_.Exception.Message)" "UI"
        continue # Springe zur nächsten Kategorie, falls der Header nicht hinzugefügt werden kann
    }
    
    # Checkboxen fuer diese Kategorie
    # $categories[$categoryKey] enthält eine Liste von Befehls-Hashtables für die aktuelle Kategorie
    foreach ($cmd in $categories[$categoryKey]) {
        $checkbox = New-Object System.Windows.Controls.CheckBox
        $checkbox.Content = $cmd.Name
        $checkbox.IsChecked = $true # Standardmäßig aktiviert
        
        # Setze Style direkt statt über FindResource
        $checkbox.Margin = New-Object System.Windows.Thickness(20, 3, 5, 3)
        $checkbox.Padding = New-Object System.Windows.Thickness(5, 0, 0, 0)
        
        # Ueberpruefe, ob diese Option mit einer Serverrolle verbunden ist
        if ($cmd.ContainsKey("FeatureName")) {
            # Test-ServerRole ist eine Funktion, die prüft, ob eine Windows-Funktion/Rolle installiert ist.
            # Diese Funktion ist außerhalb dieses Codeblocks definiert.
            $isRoleInstalled = Test-ServerRole -FeatureName $cmd.FeatureName
            if (-not $isRoleInstalled) {
                $checkbox.IsEnabled = $false
                $checkbox.Content = "$($cmd.Name) (Nicht installiert)"
                $checkbox.IsChecked = $false # Deaktiviere Checkbox, wenn zugehörige Rolle nicht installiert ist
            }
        }
        
        try {
            $spOptions.Children.Add($checkbox)
            $checkboxes[$cmd.Name] = $checkbox # Speichere eine Referenz zur Checkbox im globalen Hashtable
            Write-DebugLog "Checkbox '$($cmd.Name)' hinzugefügt (Kategorie: $categoryKey)" "UI"
        } catch {
            Write-DebugLog "FEHLER beim Hinzufügen der Checkbox '$($cmd.Name)' (Kategorie: $categoryKey): $($_.Exception.Message)" "UI"
            # Fahre mit der nächsten Checkbox fort, auch wenn eine fehlschlägt
        }
    }
}
Write-DebugLog "Checkboxen erstellt fuer $($checkboxes.Count) Optionen" "UI"

# Button-Event-Handler

# "Alle auswaehlen" Button
$btnSelectAll.Add_Click({
    Write-DebugLog "Alle Optionen auswaehlen" "UI"
    foreach ($key in $checkboxes.Keys) {
        if ($checkboxes[$key].IsEnabled) {
            $checkboxes[$key].IsChecked = $true
        }
    }
})

# "Alle abwaehlen" Button
$btnSelectNone.Add_Click({
    Write-DebugLog "Alle Optionen abwaehlen" "UI"
    foreach ($key in $checkboxes.Keys) {
        $checkboxes[$key].IsChecked = $false
    }
})

# Verbindungsaudit Button-Event-Handler

# "Alle auswaehlen" Button (Verbindungsaudit)
$btnSelectAllConnection.Add_Click({
    Write-DebugLog "Alle Verbindungsaudit-Optionen auswaehlen" "UI"
    foreach ($key in $connectionCheckboxes.Keys) {
        if ($connectionCheckboxes[$key].IsEnabled) {
            $connectionCheckboxes[$key].IsChecked = $true
        }
    }
})

# "Alle abwaehlen" Button (Verbindungsaudit)
$btnSelectNoneConnection.Add_Click({
    Write-DebugLog "Alle Verbindungsaudit-Optionen abwaehlen" "UI"
    foreach ($key in $connectionCheckboxes.Keys) {
        $connectionCheckboxes[$key].IsChecked = $false
    }
})

# "Verbindungsaudit starten" Button
$btnRunConnectionAudit.Add_Click({
    Write-DebugLog "Verbindungsaudit gestartet" "ConnectionAudit"
    Start-ConnectionAuditProcess
})

# ComboBox Selection Changed Event
$cmbResultCategories.Add_SelectionChanged({
    if ($cmbResultCategories.SelectedItem) {
        $selectedCategory = $cmbResultCategories.SelectedItem.Tag
        Show-CategoryResults -Category $selectedCategory
    }
})

# "Ergebnisse aktualisieren" Button
$btnRefreshResults.Add_Click({
    Write-DebugLog "Aktualisiere Ergebnisse-Anzeige" "UI"
    
    try {
        # Versuche die RichTextBox zurückzusetzen
        Reset-ResultsDisplay
        
        # Aktualisiere Kategorien
        Update-ResultsCategories
        
        # Zeige die ausgewählte Kategorie erneut an
        if ($cmbResultCategories.SelectedItem) {
            $selectedCategory = $cmbResultCategories.SelectedItem.Tag
            Show-CategoryResults -Category $selectedCategory
        } else {
            Show-CategoryResults -Category "Alle"
        }
        
        $txtStatus.Text = "Status: Ergebnisse erfolgreich aktualisiert"
    }
    catch {
        Write-DebugLog "FEHLER beim Aktualisieren der Ergebnisse: $($_.Exception.Message)" "UI"
        $txtStatus.Text = "Status: Fehler beim Aktualisieren der Ergebnisse"
        [System.Windows.MessageBox]::Show("Fehler beim Aktualisieren der Ergebnisse:`r`n$($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        
        # Fallback: Zeige Ergebnisse in einfacher Form
        Show-SimpleResults -Category "Alle"
    }
})

# Funktion zum Zurücksetzen der Ergebnisanzeige
function Reset-ResultsDisplay {
    Write-DebugLog "Setze Ergebnisanzeige zurück" "UI"
    
    try {
        # Erstelle ein komplett neues, leeres FlowDocument
        $newDocument = New-Object System.Windows.Documents.FlowDocument
        $newDocument.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI")
        $newDocument.FontSize = 12
        $newDocument.PageWidth = [Double]::NaN
        $newDocument.PageHeight = [Double]::NaN
        $newDocument.ColumnWidth = [Double]::PositiveInfinity
        
        # Setze das neue Dokument
        $rtbResults.Document = $newDocument
        
        # Erzwinge Update
        $rtbResults.UpdateLayout()
        
        Write-DebugLog "Ergebnisanzeige erfolgreich zurückgesetzt" "UI"
    }
    catch {
        Write-DebugLog "FEHLER beim Zurücksetzen der Ergebnisanzeige: $($_.Exception.Message)" "UI"
        throw $_
    }
}

# Hilfsfunktion zum Optimieren der RichTextBox für besseres Text-Wrapping
function Optimize-RichTextBoxLayout {
    try {
        Write-DebugLog "Optimiere RichTextBox-Layout" "UI"
        
        # Setze die RichTextBox-Eigenschaften für optimales Wrapping
        $rtbResults.HorizontalScrollBarVisibility = "Disabled"
        $rtbResults.VerticalScrollBarVisibility = "Disabled" # ScrollViewer übernimmt das Scrolling
        
        # Stelle sicher, dass das Document korrekt konfiguriert ist
        if ($rtbResults.Document) {
            $rtbResults.Document.PageWidth = [Double]::NaN
            $rtbResults.Document.PageHeight = [Double]::NaN
            $rtbResults.Document.ColumnWidth = [Double]::PositiveInfinity
            $rtbResults.Document.TextAlignment = "Left"
        }
        
        Write-DebugLog "RichTextBox-Layout optimiert" "UI"
    }
    catch {
        Write-DebugLog "WARNUNG: Konnte RichTextBox-Layout nicht optimieren: $($_.Exception.Message)" "UI"
        # Fehler ignorieren, da dies nur eine Optimierung ist
    }
}

# Vollstaendiges Audit Button
$btnFullAudit.Add_Click({
    Write-DebugLog "Vollstaendiges Audit gestartet" "Audit"
    # Alle verfuegbaren Optionen auswaehlen
    foreach ($key in $checkboxes.Keys) {
        if ($checkboxes[$key].IsEnabled) {
            $checkboxes[$key].IsChecked = $true
        }
    }
    # Audit starten
    Start-AuditProcess
})

# "Audit starten" Button
$btnRunAudit.Add_Click({
    Write-DebugLog "Benutzerdefiniertes Audit gestartet" "Audit"
    Start-AuditProcess
})

# Hauptfunktion fuer die Audit-Durchfuehrung (Synchron)
function Start-AuditProcess {
    # UI vorbereiten
    $btnRunAudit.IsEnabled = $false
    $btnFullAudit.IsEnabled = $false
    $btnExportHTML.IsEnabled = $false
    $btnCopyToClipboard.IsEnabled = $false
    $btnRefreshResults.IsEnabled = $false
    
    $rtbResults.Document = New-Object System.Windows.Documents.FlowDocument
    $txtStatusLog.Text = ""
    $progressBar.Value = 0
    $txtStatus.Text = "Status: Audit laeuft..."
    
    # UI initial aktualisieren
    $window.Dispatcher.Invoke([Action]{
        $txtProgress.Text = "Initialisiere Audit..."
        $txtStatusLog.Text = "=== Audit gestartet ===`r`n"
        $progressBar.Value = 0
    }, "Normal")
    
    # UI refresh erzwingen
    $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
    Start-Sleep -Milliseconds 300
    
    # Sammle ausgewaehlte Befehle
    $selectedCommands = @()
    foreach ($cmd in $commands) {
        if ($checkboxes[$cmd.Name].IsChecked) {
            $selectedCommands += $cmd
        }
    }
    
    Write-DebugLog "Starte Audit mit $($selectedCommands.Count) ausgewaehlten Befehlen" "Audit"
    
    $global:auditResults = @{}
    $allResults = ""
    $progressStep = 100.0 / $selectedCommands.Count
    $currentProgress = 0
    
    # UI Update mit Anzahl der Befehle
    $window.Dispatcher.Invoke([Action]{
        $txtProgress.Text = "Bereite $($selectedCommands.Count) Audit-Befehle vor..."
        $txtStatusLog.Text += "Anzahl ausgewaehlter Befehle: $($selectedCommands.Count)`r`n`r`n"
    }, "Normal")
    
    # UI refresh erzwingen
    $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
    Start-Sleep -Milliseconds 500
    
    for ($i = 0; $i -lt $selectedCommands.Count; $i++) {
        $cmd = $selectedCommands[$i]
        
        # UI aktualisieren - BEGINN des Befehls
        $window.Dispatcher.Invoke([Action]{
            $txtProgress.Text = "Verarbeite: $($cmd.Name) ($($i+1)/$($selectedCommands.Count))"
            $txtStatusLog.Text += "[$($i+1)/$($selectedCommands.Count)] $($cmd.Name)...`r`n"
            # Fortschritt am Anfang des Befehls anzeigen
            $progressBar.Value = $currentProgress
        }, "Normal")
        
        # UI refresh erzwingen - das ist der Schluessel!
        $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 200
        
        Write-DebugLog "Fuehre aus ($($i+1)/$($selectedCommands.Count)): $($cmd.Name)" "Audit"
        
        try {
            if ($cmd.Type -eq "PowerShell") {
                $result = Invoke-PSCommand -Command $cmd.Command
            } else {
                $result = Invoke-CMDCommand -Command $cmd.Command
            }
            
            $global:auditResults[$cmd.Name] = $result
            $allResults += "`r`n=== $($cmd.Name) ===`r`n$result`r`n"
            
            # Erfolg in Status-Log UND Fortschrittsbalken aktualisieren
            $currentProgress += $progressStep
            $window.Dispatcher.Invoke([Action]{
                $txtStatusLog.Text += "  [OK] Erfolgreich abgeschlossen`r`n"
                # Fortschrittsbalken NACH erfolgreichem Befehl aktualisieren
                $progressBar.Value = $currentProgress
                $txtProgress.Text = "Abgeschlossen: $($cmd.Name) ($($i+1)/$($selectedCommands.Count))"
            }, "Normal")
            
        } catch {
            $errorMsg = "Fehler: $($_.Exception.Message)"
            $global:auditResults[$cmd.Name] = $errorMsg
            $allResults += "`r`n=== $($cmd.Name) ===`r`n$errorMsg`r`n"
            
            # Fehler in Status-Log UND Fortschrittsbalken trotzdem aktualisieren
            $currentProgress += $progressStep
            $window.Dispatcher.Invoke([Action]{
                $txtStatusLog.Text += "  [FEHLER] $($_.Exception.Message)`r`n"
                # Fortschrittsbalken auch bei Fehler aktualisieren
                $progressBar.Value = $currentProgress
                $txtProgress.Text = "Fehler bei: $($cmd.Name) ($($i+1)/$($selectedCommands.Count))"
            }, "Normal")
            
            Write-DebugLog "FEHLER bei $($cmd.Name): $($_.Exception.Message)" "Audit"
        }
        
        # UI refresh nach jedem Befehl erzwingen - SEHR WICHTIG!
        $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 300
        
        # Zwischenstand der Ergebnisse aktualisieren (optional)
        if (($i + 1) % 3 -eq 0 -or $i -eq ($selectedCommands.Count - 1)) {
            $window.Dispatcher.Invoke([Action]{
                try {
                    # Aktualisiere die schöne Anzeige statt der alten Textbox
                    Update-ResultsCategories
                    if ($cmbResultCategories.SelectedItem) {
                        $selectedCategory = $cmbResultCategories.SelectedItem.Tag
                        Show-CategoryResults -Category $selectedCategory
                    } else {
                        Show-CategoryResults -Category "Alle"
                    }
                }
                catch {
                    Write-DebugLog "FEHLER beim Zwischenupdate der Ergebnisanzeige: $($_.Exception.Message)" "Audit"
                    # Bei Fehlern trotzdem fortfahren
                }
            }, "Normal")
            # UI refresh auch hier
            $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    
    # Audit abgeschlossen - Finale Updates
    $window.Dispatcher.Invoke([Action]{
        $progressBar.Value = 100
        $txtProgress.Text = "Audit vollstaendig abgeschlossen! $($selectedCommands.Count) Befehle ausgefuehrt."
        $txtStatusLog.Text += "`r`n" + "="*50 + "`r`n"
        $txtStatusLog.Text += "[FERTIG] Audit erfolgreich abgeschlossen!`r`n"
        $txtStatusLog.Text += "Ergebnisse: $($global:auditResults.Count) Eintraege`r`n"
        $txtStatusLog.Text += "="*50 + "`r`n"
        
        try {
            # Aktualisiere die schöne Kategorien-Anzeige
            Update-ResultsCategories
            Show-CategoryResults -Category "Alle"
        }
        catch {
            Write-DebugLog "FEHLER beim finalen Update der Ergebnisanzeige: $($_.Exception.Message)" "Audit"
            # Fallback auf einfache Anzeige
            try {
                Show-SimpleResults -Category "Alle"
                $txtStatusLog.Text += "[WARNUNG] Verwendet einfache Ergebnisanzeige aufgrund von Formatierungsproblemen`r`n"
            }
            catch {
                Write-DebugLog "FEHLER auch bei einfacher Anzeige: $($_.Exception.Message)" "Audit"
                $txtStatusLog.Text += "[FEHLER] Konnte Ergebnisse nicht anzeigen - siehe Debug-Log`r`n"
            }
        }
        
        $txtStatus.Text = "Status: Audit abgeschlossen - $($global:auditResults.Count) Ergebnisse"
        
        # Buttons wieder aktivieren
        $btnRunAudit.IsEnabled = $true
        $btnFullAudit.IsEnabled = $true
        $btnExportHTML.IsEnabled = $true
        $btnCopyToClipboard.IsEnabled = $true
        $btnRefreshResults.IsEnabled = $true
    }, "Normal")
    
    # Finaler UI refresh
    $window.Dispatcher.Invoke([Action]{}, "ApplicationIdle")
    [System.Windows.Forms.Application]::DoEvents()
    
    Write-DebugLog "Audit abgeschlossen mit $($global:auditResults.Count) Ergebnissen" "Audit"
}

# Export-Button-Funktionalitaet
$btnExportHTML.Add_Click({
    Write-DebugLog "HTML-Export gestartet" "Export"
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "HTML Files (*.html)|*.html"
    $saveFileDialog.Title = "Speichern Sie den Audit-Bericht"
    $saveFileDialog.FileName = "ServerAudit_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtStatus.Text = "Status: Exportiere HTML..."
        
        try {
            Export-AuditToHTML -Results $global:auditResults -FilePath $saveFileDialog.FileName
            $txtStatus.Text = "Status: Export erfolgreich abgeschlossen"
            [System.Windows.MessageBox]::Show("Bericht wurde erfolgreich exportiert:`r`n$($saveFileDialog.FileName)", "Export erfolgreich", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        } catch {
            $txtStatus.Text = "Status: Fehler beim Export"
            [System.Windows.MessageBox]::Show("Fehler beim Export:`r`n$($_.Exception.Message)", "Export Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
})

# Zwischenablage-Button
$btnCopyToClipboard.Add_Click({
    Write-DebugLog "Kopiere Ergebnisse in Zwischenablage" "UI"
    try {
        # Extrahiere Text aus der RichTextBox
        $textRange = New-Object System.Windows.Documents.TextRange($rtbResults.Document.ContentStart, $rtbResults.Document.ContentEnd)
        $plainText = $textRange.Text
        
        if ([string]::IsNullOrWhiteSpace($plainText)) {
            # Fallback: Erstelle Text aus den Rohdaten
            $allResults = ""
            foreach ($key in $global:auditResults.Keys | Sort-Object) {
                $allResults += "`r`n=== $key ===`r`n$($global:auditResults[$key])`r`n"
            }
            $plainText = $allResults
        }
        
        $plainText | Set-Clipboard
        $txtStatus.Text = "Status: Ergebnisse in Zwischenablage kopiert"
        [System.Windows.MessageBox]::Show("Audit-Ergebnisse wurden in die Zwischenablage kopiert.", "Kopiert", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    } catch {
        [System.Windows.MessageBox]::Show("Fehler beim Kopieren in die Zwischenablage: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

# Debug-Funktionen
$btnOpenLog.Add_Click({
    Write-DebugLog "Oeffne Log-Datei" "Debug"
    if (Test-Path $DebugLogPath) {
        Start-Process notepad.exe -ArgumentList $DebugLogPath
    } else {
        [System.Windows.MessageBox]::Show("Log-Datei nicht gefunden.", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

$btnClearLog.Add_Click({
    Write-DebugLog "Debug-Log wird geleert" "Debug"
    $script:txtDebugOutput.Text = ""
    if ($DEBUG) {
        $clearMessage = "=== Debug-Log geloescht: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") ==="
        Set-Content -Path $DebugLogPath -Value $clearMessage -Force
        Write-Host $clearMessage -ForegroundColor Cyan
    }
})

Write-DebugLog "UI-Initialisierung abgeschlossen" "Init"

# Funktion zum Formatieren der RichTextBox
function Format-RichTextResults {
    param(
        [hashtable]$Results,
        [string]$CategoryFilter = "Alle"
    )
    
    Write-DebugLog "Formatiere Ergebnisse fuer Kategorie: $CategoryFilter" "UI"
    
    # Erstelle ein neues FlowDocument
    $document = New-Object System.Windows.Documents.FlowDocument
    $document.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI")
    $document.FontSize = 12
    $document.LineHeight = 18
    
    # Optimierte Layout-Einstellungen für bessere Textanzeige
    $document.PageWidth = [Double]::NaN  # Unbeschränkte Breite
    $document.PageHeight = [Double]::NaN # Unbeschränkte Höhe
    $document.ColumnWidth = [Double]::PositiveInfinity # Keine Spalten-Begrenzung
    $document.TextAlignment = "Left"
    $document.PagePadding = New-Object System.Windows.Thickness(0)
    $document.IsOptimalParagraphEnabled = $true
    $document.IsHyphenationEnabled = $false
    
    # Gruppiere Ergebnisse nach Kategorien
    $categorizedResults = @{}
    foreach ($cmd in $commands) {
        $category = if ($cmd.Category) { $cmd.Category } else { "Allgemein" }
        if (-not $categorizedResults.ContainsKey($category)) {
            $categorizedResults[$category] = @()
        }
        if ($Results.ContainsKey($cmd.Name)) {
            $categorizedResults[$category] += @{
                Name = $cmd.Name
                Result = $Results[$cmd.Name]
                Command = $cmd
            }
        }
    }
    
    # Bestimme welche Kategorien angezeigt werden sollen
    $categoriesToShow = if ($CategoryFilter -eq "Alle") { 
        $categorizedResults.Keys | Sort-Object 
    } else { 
        @($CategoryFilter) 
    }
    
    $totalItems = 0
    foreach ($category in $categoriesToShow) {
        if ($categorizedResults.ContainsKey($category)) {
            $categoryData = $categorizedResults[$category]
            $totalItems += $categoryData.Count
            
            # Kategorie-Header
            $categoryParagraph = New-Object System.Windows.Documents.Paragraph
            $categoryRun = New-Object System.Windows.Documents.Run("Kategorie: $category")
            $categoryRun.FontWeight = "Bold"
            $categoryRun.FontSize = 16
            $categoryRun.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 120, 212))
            $categoryParagraph.Inlines.Add($categoryRun)
            $categoryParagraph.Margin = New-Object System.Windows.Thickness(0, 15, 0, 10)
            $categoryParagraph.TextAlignment = "Left"
            $categoryParagraph.KeepTogether = $true
            $document.Blocks.Add($categoryParagraph)
            
            # Items in dieser Kategorie
            foreach ($item in $categoryData) {
                # Item-Header
                $itemParagraph = New-Object System.Windows.Documents.Paragraph
                $itemRun = New-Object System.Windows.Documents.Run("Eintrag: $($item.Name)")
                $itemRun.FontWeight = "SemiBold"
                $itemRun.FontSize = 13
                $itemRun.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(44, 62, 80))
                $itemParagraph.Inlines.Add($itemRun)
                $itemParagraph.Margin = New-Object System.Windows.Thickness(0, 10, 0, 5)
                $itemParagraph.TextAlignment = "Left"
                $itemParagraph.KeepTogether = $true
                $document.Blocks.Add($itemParagraph)
                
                # Kommando-Info (optional)
                if ($item.Command.Command) {
                    $cmdParagraph = New-Object System.Windows.Documents.Paragraph
                    $cmdRun = New-Object System.Windows.Documents.Run("Befehl: $($item.Command.Command)")
                    $cmdRun.FontStyle = "Italic"
                    $cmdRun.FontSize = 10
                    $cmdRun.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(108, 117, 125))
                    $cmdParagraph.Inlines.Add($cmdRun)
                    $cmdParagraph.Margin = New-Object System.Windows.Thickness(20, 0, 0, 5)
                    $cmdParagraph.TextAlignment = "Left"
                    $document.Blocks.Add($cmdParagraph)
                }
                
                # Ergebnis in einem optimierten Paragraph für bessere Kontrolle
                $resultParagraph = New-Object System.Windows.Documents.Paragraph
                $resultParagraph.Margin = New-Object System.Windows.Thickness(0, 0, 0, 15)
                $resultParagraph.Padding = New-Object System.Windows.Thickness(15)
                $resultParagraph.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(248, 249, 250))
                $resultParagraph.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(222, 226, 230))
                $resultParagraph.BorderThickness = New-Object System.Windows.Thickness(1)
                $resultParagraph.TextAlignment = "Left"
                
                # Optimierte Textbehandlung - einfacher und robuster
                $resultText = $item.Result
                if ([string]::IsNullOrWhiteSpace($resultText)) {
                    $resultText = "Keine Daten verfügbar"
                }
                
                # Verwende einen einzigen Run für den gesamten Text
                $resultRun = New-Object System.Windows.Documents.Run($resultText)
                $resultRun.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
                $resultRun.FontSize = 11
                $resultParagraph.Inlines.Add($resultRun)
                
                $document.Blocks.Add($resultParagraph)
            }
        }
    }
    
    # Update Summary
    $window.Dispatcher.Invoke([Action]{
        $txtResultsSummary.Text = "Zeige $totalItems Einträge"
    }, "Normal")
    
    return $document
}

# Funktion zum Aktualisieren der Kategorien-ComboBox
function Update-ResultsCategories {
    Write-DebugLog "Aktualisiere Kategorien-ComboBox" "UI"
    
    $cmbResultCategories.Items.Clear()
    
    # "Alle" Option hinzufügen
    $allItem = New-Object System.Windows.Controls.ComboBoxItem
    $allItem.Content = "Alle Kategorien" # Emoji entfernt
    $allItem.Tag = "Alle"
    $cmbResultCategories.Items.Add($allItem)
    
    # Einzelne Kategorien hinzufügen
    # $commands wird als global/script-scoped Variable angenommen
    $categories = @{}
    if ($null -ne $commands) {
        foreach ($cmd in $commands) {
            $category = if ($cmd.Category) { $cmd.Category } else { "Allgemein" }
            if (-not $categories.ContainsKey($category)) {
                $categories[$category] = 0
            }
            if ($global:auditResults.ContainsKey($cmd.Name)) {
                $categories[$category]++
            }
        }
    }
    
    foreach ($category in $categories.Keys | Sort-Object) {
        if ($categories[$category] -gt 0) {
            $categoryItem = New-Object System.Windows.Controls.ComboBoxItem
            $categoryItem.Content = "$category ($($categories[$category]))" # Emoji entfernt
            $categoryItem.Tag = $category
            $cmbResultCategories.Items.Add($categoryItem)
        }
    }
    
    # Ersten Eintrag auswählen
    if ($cmbResultCategories.Items.Count -gt 0) {
        $cmbResultCategories.SelectedIndex = 0
    }
}

# Funktion zum Anzeigen der Ergebnisse
function Show-CategoryResults {
    param([string]$Category = "Alle")
    
    Write-DebugLog "Zeige Ergebnisse fuer Kategorie: $Category" "UI"
    
    if ($global:auditResults.Count -eq 0) {
        $rtbResults.Document = New-Object System.Windows.Documents.FlowDocument
        $emptyParagraph = New-Object System.Windows.Documents.Paragraph
        $emptyRun = New-Object System.Windows.Documents.Run("Keine Audit-Ergebnisse verfügbar. Führen Sie zuerst ein Audit durch.")
        $emptyRun.FontStyle = "Italic"
        $emptyRun.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(108, 117, 125))
        $emptyParagraph.Inlines.Add($emptyRun)
        $rtbResults.Document.Blocks.Add($emptyParagraph)
        return
    }
    
    try {
        # Versuche die formatierte Anzeige
        $document = Format-RichTextResults -Results $global:auditResults -CategoryFilter $Category
        $rtbResults.Document = $document
        
        # Erzwinge Layout-Update
        $rtbResults.UpdateLayout()
        
        Write-DebugLog "Ergebnisse erfolgreich formatiert und angezeigt" "UI"
    }
    catch {
        Write-DebugLog "FEHLER bei der formatierten Anzeige: $($_.Exception.Message) - Verwende Fallback" "UI"
        
        # Fallback: Verwende einfache Textanzeige
        Show-SimpleResults -Category $Category
    }
}

# Fallback-Funktion für einfache Textanzeige
function Show-SimpleResults {
    param([string]$Category = "Alle")
    
    Write-DebugLog "Verwende einfache Textanzeige für Kategorie: $Category" "UI"
    
    # Erstelle einfaches FlowDocument
    $document = New-Object System.Windows.Documents.FlowDocument
    $document.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
    $document.FontSize = 11
    $document.PageWidth = [Double]::NaN
    $document.PageHeight = [Double]::NaN
    $document.ColumnWidth = [Double]::PositiveInfinity
    
    # Sammle alle relevanten Ergebnisse als einfachen Text
    $resultText = ""
    
    # Gruppiere nach Kategorien
    $categorizedResults = @{}
    foreach ($cmd in $commands) {
        $cmdCategory = if ($cmd.Category) { $cmd.Category } else { "Allgemein" }
        if (-not $categorizedResults.ContainsKey($cmdCategory)) {
            $categorizedResults[$cmdCategory] = @()
        }
        if ($global:auditResults.ContainsKey($cmd.Name)) {
            $categorizedResults[$cmdCategory] += @{
                Name = $cmd.Name
                Result = $global:auditResults[$cmd.Name]
            }
        }
    }
    
    # Bestimme anzuzeigende Kategorien
    $categoriesToShow = if ($Category -eq "Alle") { 
        $categorizedResults.Keys | Sort-Object 
    } else { 
        @($Category) 
    }
    
    $totalItems = 0
    foreach ($cat in $categoriesToShow) {
        if ($categorizedResults.ContainsKey($cat)) {
            $categoryData = $categorizedResults[$cat]
            $totalItems += $categoryData.Count
            
            $resultText += "`n" + "="*60 + "`n"
            $resultText += "KATEGORIE: $cat`n"
            $resultText += "="*60 + "`n`n"
            
            foreach ($item in $categoryData) {
                $resultText += "-"*40 + "`n"
                $resultText += "EINTRAG: $($item.Name)`n"
                $resultText += "-"*40 + "`n"
                $resultText += "$($item.Result)`n`n"
            }
        }
    }
    
    # Erstelle einfachen Paragraph mit dem gesamten Text
    $paragraph = New-Object System.Windows.Documents.Paragraph
    $run = New-Object System.Windows.Documents.Run($resultText)
    $paragraph.Inlines.Add($run)
    $document.Blocks.Add($paragraph)
    
    $rtbResults.Document = $document
    
    # Update Summary
    $window.Dispatcher.Invoke([Action]{
        $txtResultsSummary.Text = "Zeige $totalItems Einträge (einfache Ansicht)"
    }, "Normal")
}

# Initialisiere die Ergebnisse-Anzeige
Show-CategoryResults -Category "Alle"

# Optimiere die RichTextBox für besseres Text-Wrapping
Optimize-RichTextBoxLayout

# Zeige das Fenster an
$txtStatus.Text = "Status: Bereit fuer Audit"
Write-DebugLog "Zeige Hauptfenster" "UI"
$null = $window.ShowDialog()
Write-DebugLog "Anwendung geschlossen" "Shutdown"

# Initialisiere die Ergebnisse-Anzeige
Show-CategoryResults -Category "Alle"

# Verbindungsaudit Export-Button-Funktionalitaet
$btnExportConnectionHTML.Add_Click({
    Write-DebugLog "Verbindungsaudit HTML-Export gestartet" "Export"
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "HTML Files (*.html)|*.html"
    $saveFileDialog.Title = "Speichern Sie den Verbindungsaudit-Bericht"
    $saveFileDialog.FileName = "ConnectionAudit_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtStatus.Text = "Status: Exportiere Verbindungsaudit-HTML..."
        
        try {
            Export-ConnectionAuditToHTML -Results $global:connectionAuditResults -FilePath $saveFileDialog.FileName
            $txtStatus.Text = "Status: Verbindungsaudit-Export erfolgreich abgeschlossen"
            [System.Windows.MessageBox]::Show("Verbindungsaudit-Bericht wurde erfolgreich exportiert:`r`n$($saveFileDialog.FileName)", "Export erfolgreich", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        } catch {
            $txtStatus.Text = "Status: Fehler beim Verbindungsaudit-Export"
            [System.Windows.MessageBox]::Show("Fehler beim Export:`r`n$($_.Exception.Message)", "Export Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
})

# Verbindungsaudit Zwischenablage-Button
$btnCopyConnectionToClipboard.Add_Click({
    Write-DebugLog "Kopiere Verbindungsaudit-Ergebnisse in Zwischenablage" "UI"
    try {
        # Extrahiere Text aus der RichTextBox
        $textRange = New-Object System.Windows.Documents.TextRange($rtbConnectionResults.Document.ContentStart, $rtbConnectionResults.Document.ContentEnd)
        $plainText = $textRange.Text
        
        if ([string]::IsNullOrWhiteSpace($plainText)) {
            # Fallback: Erstelle Text aus den Rohdaten
            $allResults = ""
            foreach ($key in $global:connectionAuditResults.Keys | Sort-Object) {
                $allResults += "`r`n=== $key ===`r`n$($global:connectionAuditResults[$key])`r`n"
            }
            $plainText = $allResults
        }
        
        $plainText | Set-Clipboard
        $txtStatus.Text = "Status: Verbindungsaudit-Ergebnisse in Zwischenablage kopiert"
        [System.Windows.MessageBox]::Show("Verbindungsaudit-Ergebnisse wurden in die Zwischenablage kopiert.", "Kopiert", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    } catch {
        [System.Windows.MessageBox]::Show("Fehler beim Kopieren in die Zwischenablage: $($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

# Verbindungsaudit ComboBox Selection Changed Event
$cmbConnectionCategories.Add_SelectionChanged({
    if ($cmbConnectionCategories.SelectedItem) {
        $selectedCategory = $cmbConnectionCategories.SelectedItem.Tag
        Show-ConnectionCategoryResults -Category $selectedCategory
    }
})

# Initialisiere die Verbindungsaudit-Ergebnisse-Anzeige
Show-ConnectionCategoryResults -Category "Alle"

# Initialisiere die Ergebnisse-Anzeige
Show-CategoryResults -Category "Alle"
