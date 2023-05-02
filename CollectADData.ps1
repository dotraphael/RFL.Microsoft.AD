<#
    .SYSNOPSIS
        Collect Data from a Active Directory environment to generate a Discovery Report using the ExportADData.ps1 PowerShell

    .DESCRIPTION
        Collect Data from a Active Directory environment to generate a Discovery Report using the ExportADData.ps1 PowerShell

    .PARAMETER SaveTo
        Where the .json files will be saved

    .NOTES
        Name: CollectADData.ps1
        Author: Raphael Perez
        DateCreated: 02 May 2020 (v0.1)
        Website: http://www.endpointmanagers.com
        WebSite: https://github.com/dotraphael/RFL.Microsoft.AD
        Twitter: @dotraphael
        Twitter: @dotraphael

    .LINK
        http://www.endpointmanagers.com
        http://www.rflsystems.co.uk
        https://github.com/dotraphael/HealthCheckToolkit_Community

    .EXAMPLE
        .\CollectADData.ps1 -SaveTo 'c:\temp\CollectADData'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = 'Please provide the path to where the json files will be saved to')]
    [ValidateNotNullOrEmpty()]
	[ValidateScript({ if (Test-Path $_ -PathType 'Container') { $true } else { throw "$_ is not a valid folder path" }  })]
    [String] $SaveTo
)

$Error.Clear()
#region Functions
#region Test-RFLAdministrator
Function Test-RFLAdministrator {
<#
    .SYSNOPSIS
        Check if the current user is member of the Local Administrators Group

    .DESCRIPTION
        Check if the current user is member of the Local Administrators Group

    .NOTES
        Name: Test-RFLAdministrator
        Author: Raphael Perez
        DateCreated: 28 November 2019 (v0.1)

    .EXAMPLE
        Test-RFLAdministrator
#>
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    (New-Object Security.Principal.WindowsPrincipal $currentUser).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}
#endregion

#region Set-RFLLogPath
Function Set-RFLLogPath {
<#
    .SYSNOPSIS
        Configures the full path to the log file depending on whether or not the CCM folder exists.

    .DESCRIPTION
        Configures the full path to the log file depending on whether or not the CCM folder exists.

    .NOTES
        Name: Set-RFLLogPath
        Author: Raphael Perez
        DateCreated: 28 November 2019 (v0.1)

    .EXAMPLE
        Set-RFLLogPath
#>
    if ([string]::IsNullOrEmpty($script:LogFilePath)) {
        $script:LogFilePath = $env:Temp
    }

    $script:ScriptLogFilePath = "$($script:LogFilePath)\$($Script:LogFileFileName)"
}
#endregion

#region Write-RFLLog
Function Write-RFLLog {
<#
    .SYSNOPSIS
        Write the log file if the global variable is set

    .DESCRIPTION
        Write the log file if the global variable is set

    .PARAMETER Message
        Message to write to the log

    .PARAMETER LogLevel
        Log Level 1=Information, 2=Warning, 3=Error. Default = 1

    .NOTES
        Name: Write-RFLLog
        Author: Raphael Perez
        DateCreated: 28 November 2019 (v0.1)

    .EXAMPLE
        Write-RFLLog -Message 'This is an information message'

    .EXAMPLE
        Write-RFLLog -Message 'This is a warning message' -LogLevel 2

    .EXAMPLE
        Write-RFLLog -Message 'This is an error message' -LogLevel 3
#>
param (
    [Parameter(Mandatory = $true)]
    [string]$Message,

    [Parameter()]
    [ValidateSet(1, 2, 3)]
    [string]$LogLevel=1
)
    $TimeNow = Get-Date   
    $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
    $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'
    if ([string]::IsNullOrEmpty($MyInvocation.ScriptName)) {
        $ScriptName = ''
    } else {
        $ScriptName = $MyInvocation.ScriptName | Split-Path -Leaf
    }

    $LineFormat = $Message, $TimeGenerated, $TimeNow.ToString('MM-dd-yyyy'), "$($ScriptName):$($MyInvocation.ScriptLineNumber)", $LogLevel
    $Line = $Line -f $LineFormat

    $Line | Out-File -FilePath $script:ScriptLogFilePath -Append -NoClobber -Encoding default
    $HostMessage = '{0} {1}' -f $TimeNow.ToString('dd-MM-yyyy HH:mm'), $Message
    switch ($LogLevel) {
        2 { Write-Host $HostMessage -ForegroundColor Yellow }
        3 { Write-Host $HostMessage -ForegroundColor Red }
        default { Write-Host $HostMessage }
    }
}
#endregion

#region Clear-RFLLog
Function Clear-RFLLog {
<#
    .SYSNOPSIS
        Delete the log file if bigger than maximum size

    .DESCRIPTION
        Delete the log file if bigger than maximum size

    .NOTES
        Name: Clear-RFLLog
        Author: Raphael Perez
        DateCreated: 28 November 2019 (v0.1)

    .EXAMPLE
        Clear-RFLLog -maxSize 2mb
#>
param (
    [Parameter(Mandatory = $true)][string]$maxSize
)
    try  {
        if(Test-Path -Path $script:ScriptLogFilePath) {
            if ((Get-Item $script:ScriptLogFilePath).length -gt $maxSize) {
                Remove-Item -Path $script:ScriptLogFilePath
                Start-Sleep -Seconds 1
            }
        }
    }
    catch {
        Write-RFLLog -Message "Unable to delete log file." -LogLevel 3
    }    
}
#endregion

#region Get-ScriptDirectory
function Get-ScriptDirectory {
<#
    .SYSNOPSIS
        Get the directory of the script

    .DESCRIPTION
        Get the directory of the script

    .NOTES
        Name: ClearGet-ScriptDirectory
        Author: Raphael Perez
        DateCreated: 28 November 2019 (v0.1)

    .EXAMPLE
        Get-ScriptDirectory
#>
    Split-Path -Parent $PSCommandPath
}
#endregion

#region fncObjMerge
Function fncObjMerge {
<#
    .SYSNOPSIS
        Merge two objects and add a "SourceDomain" property with the value from the DomainName parameter

    .DESCRIPTION
        Merge two objects and add a "SourceDomain" property with the value from the DomainName parameter

    .NOTES
        Name: fncObjMerge
        Author: Raphael Perez
        DateCreated: 02 May 2023 (v0.1)

    .EXAMPLE
        fncObjMerge
#>
	[CmdletBinding()]
	param(
        [object]$inputobject,

        [ValidateNotNullOrEmpty()]
        [string]$DomainName
	)
    if ($null -ne $inputobject) {
        $fields = ($inputobject | Get-Member | Where-Object {$_.MemberType -in ('Property','NoteProperty')} | select Name)
        $objReturn = @()

        foreach($item in $inputobject) {
            $objMerge1 = New-Object -TypeName PSObject
            $objMerge1 | Add-Member -NotePropertyName SourceDomain -NotePropertyValue $DomainName

            foreach($keyitem in $fields.Name) { 
                $objMerge1 | Add-Member -NotePropertyName $keyitem -NotePropertyValue  $item.$keyitem
            }
            $objReturn += $objMerge1
        }
        $objReturn
    }
}
#endregion

#region Get-RFLWindowsFeature
Function Get-RFLWindowsFeature {
<#
    .SYSNOPSIS
        Get the Windows Feature

    .DESCRIPTION
        Get the Windows Feature

    .NOTES
        Name: Get-RFLFeature
        Author: Raphael Perez
        DateCreated: 02 May 2023 (v0.1)

    .EXAMPLE
        Get-RFLFeature -Name
#>
	[CmdletBinding()]
	param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$Name,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Workstation', 'Server')]
        [String]$OSType
	)
    if ($OSType -eq 'Workstation') {
        Get-WindowsCapability -online -Name $Name
    } else {
        Get-WindowsFeature -Name $Name
    }
}
#endregion

#endregion

#region Variables
$script:ScriptVersion = '0.1'
$script:LogFilePath = $env:Temp
$Script:LogFileFileName = 'CollectADData.log'
$script:ScriptLogFilePath = "$($script:LogFilePath)\$($Script:LogFileFileName)"
$script:WorkstationFeatures = @('Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0', 'Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0', 'Rsat.Dns.Tools~~~~0.0.1.0', 'Rsat.DHCP.Tools~~~~0.0.1.0')
$script:ServerFeatures = @('RSAT-AD-PowerShell', 'RSAT-DNS-Server', 'RSAT-DHCP', 'GPMC')
$Script:Modules = @()
$Script:CurrentFolder = (Get-Location).Path
#endregion

#region Main
try {
    Set-RFLLogPath
    Clear-RFLLog 25mb

    Write-RFLLog -Message "*** Starting ***"
    Write-RFLLog -Message "Script version $($script:ScriptVersion)"
    Write-RFLLog -Message "Running as $($env:username) $(if(Test-RFLAdministrator) {"[Administrator]"} Else {"[Not Administrator]"}) on $($env:computername)"

	Write-RFLLog -Message "Please refer to the RFL.Microsoft.AD github website for more detailed information about this project." -LogLevel 2
	Write-RFLLog -Message "Do not forget to update your report configuration file after each new release." -LogLevel 2
	Write-RFLLog -Message "Documentation: https://github.com/dotraphael/RFL.Microsoft.AD" -LogLevel 2
	Write-RFLLog -Message "Issues or bug reporting: https://github.com/dotraphael/RFL.Microsoft.AD/issues" -LogLevel 2

    $PSCmdlet.MyInvocation.BoundParameters.Keys | ForEach-Object { 
        Write-RFLLog -Message "Parameter '$($_)' is '$($PSCmdlet.MyInvocation.BoundParameters.Item($_))'"
    }

    $PSVersionTable.Keys | ForEach-Object { 
        Write-RFLLog -Message "PSVersionTable '$($_)' is '$($PSVersionTable.Item($_) -join (', '))'"
    }    

    if ($PSVersionTable.item('PSVersion').Tostring() -lt '4.0') {
        throw "The requested operation requires PowerShell 4.0 or newer"
    }

    if (-not (Test-RFLAdministrator)) {
        throw "The requested operation requires elevation: Run PowerShell console as administrator"
    }

    Write-RFLLog -Message "Validating required Windows Features"
    $OSInfo = Get-WmiObject win32_OperatingSystem
    Write-RFLLog -Message "    OS Information: $($OSInfo.Caption)"    
    if ($OSInfo.Caption -like '*Server*') {
        $OSType = 'Server'
    } else {
        $OSType = 'Workstation'
    }
    
    $Continue = $true
	if ($OSType -eq 'WorkStation') {
        $script:WorkstationFeatures | ForEach-Object {
            $Feature = Get-RFLWindowsFeature -Name $_ -OSType $OSType
            if ($Feature.State -ne 'Installed') {
                Write-RFLLog -Message "    Feature $($_) not installed. Use Add-WindowsCapability -online -Name '$($_)' to install the required windows feature" -LogLevel 3
                $Continue = $false
            } else {
                Write-RFLLog -Message "    Feature $($_) installed"
            }
        }
    } elseif ($OSType -eq 'Server') {
        $script:ServerFeatures | ForEach-Object {
            $Feature = Get-RFLWindowsFeature -Name $_ -OSType $OSType
            if ($Feature.InstallState -ne 'Installed') {
                Write-RFLLog -Message "    Feature $($_) not installed. Use Install-WindowsFeature -Name '$($_)' to install the required windows feature" -LogLevel 3
                $Continue = $false
            } else {
                Write-RFLLog -Message "    Feature $($_) installed"
            }
        }
    }

    if (-not $Continue) {
        throw "The requested operation requires missing Windows Feature. Install the missing Features and try again"
    }

    Write-RFLLog -Message "Getting list of installed modules"
    $InstalledModules = Get-Module -ListAvailable -ErrorAction SilentlyContinue
    $InstalledModules | ForEach-Object { 
        Write-RFLLog -Message "    Module: '$($_.Name)', Type: '$($_.ModuleTYpe)', Verison: '$($_.Version)', Path: '$($_.ModuleBase)'"
    }

    Write-RFLLog -Message "Validating required PowerShell Modules"
    $Continue = $true
    $Script:Modules | ForEach-Object {
        $Module = $InstalledModules | Where-Object {$_.Name -eq $_}
        if ($null -eq $Module) {
            Write-RFLLog -Message "    Module $($_) not installed. Use Install-Module $($_) -force to install the required powershell modules" -LogLevel 3
            $Continue = $false
        } else {
            Write-RFLLog -Message "    Module $($_) installed. Type: '$($_.ModuleTYpe)', Verison: '$($_.Version)', Path: '$($_.ModuleBase)'"
        } 
    }
    if (-not $Continue) {
        throw "The requested operation requires missing PowerShell Modules. Install the missing PowerShell modules and try again"
    }

    Write-RFLLog -Message "Current Folder '$($Script:CurrentFolder)'"
    Set-Location $SaveTo

    Write-RFLLog -Message "All checks completed successful. Starting collecting data for report"

    #region forest
    Write-RFLLog -Message "Collecting Forest Root Information"
    $Global:RootDSE = Get-ADRootDSE
    $Global:RootDSE | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath .\RootDSE.json -Force
    $Global:ForestName = $RootDSE.ldapServiceName.Split(':')[0]

    #Write-RFLLog -Message "Collecting Forest Information"
    $data = Get-ADForest
    $Global:RootDomain = $Data.RootDomain.toUpper()
    $Global:Domains = @()
    $Global:Domains += $Global:RootDomain
    $Global:Domains += $Data.Domains | Where-Object {$_ -ne $Global:RootDomain}

    #Write-RFLLog -Message "Collecting Domain Root Information"
    $rootDomainInfo = Get-ADDomain -Identity $data.RootDomain
    $DomainDN = $rootDomainInfo.DistinguishedName
    $ForestNetbiosName = $rootDomainInfo.NetBIOSName

    #Write-RFLLog -Message "Collecting Tombstone Life Time information"
    $TombstoneLifetime = Get-ADObject "CN=Directory Service,CN=Windows NT,CN=Services,CN=Configuration,$($DomainDN)" -Properties tombstoneLifetime | Select-Object -ExpandProperty tombstoneLifetime

    #Write-RFLLog -Message "Collecting Schema Version Information"
    $ADVersion = Get-ADObject $Global:RootDSE.schemaNamingContext -property objectVersion | Select-Object -ExpandProperty objectVersion
    $server = switch ($ADVersion) {
        '88' { 'Windows Server 2019 or 2022' }
        '87' { 'Windows Server 2016' }
        '69' { 'Windows Server 2012 R2' }
        '56' { 'Windows Server 2012' }
        '47' { 'Windows Server 2008 R2' }
        '44' { 'Windows Server 2008' }
        '31' { 'Windows Server 2003 R2' }
        '30' { 'Windows Server 2003' }
        '13' { 'Windows Server 2000' }
        default { $ADVersion }
    }

    Write-RFLLog -Message "Export Collected Information"
    $OutObj = $data | select @{Name="ForestNetbiosName";Expression = {$ForestNetbiosName}}, @{Name="Forest Name";Expression = {$_.RootDomain}}, @{Name="Forest Functional Level";Expression = {$_.ForestMode.ToString()}}, @{Name="Schema Version";Expression = {"$($server) - (ObjectVersion = $($ADVersion))"}}, @{Name="Schema Master";Expression = {$_.SchemaMaster}}, @{Name="Tombstone Lifetime (days)";Expression = {$TombstoneLifetime}}, @{Name="Domain Naming Master";Expression = {$_.DomainNamingMaster}}, @{Name="Domains";Expression = {$_.Domains -join '; '}}, @{Name="Global Catalogs";Expression = {$_.GlobalCatalogs -join '; '}}, @{Name="Domains Count";Expression = {$_.Domains.Count}}, @{Name="Global Catalogs Count";Expression = {$_.GlobalCatalogs.Count}}, @{Name="Sites Count";Expression = {$_.Sites.Count}}, @{Name="Application Partitions";Expression = {$_.ApplicationPartitions -join '; '}}, @{Name="Partitions Container";Expression = {[string]$_.PartitionsContainer}}, @{Name="SPN Suffixes";Expression = {$_.SPNSuffixes -join '; '}}, @{Name="UPN Suffixes";Expression = {$_.UPNSuffixes -join '; '}}
    $OutObj | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath .\ADForest.json -Force
    #endregion

    #region DCs
    Write-RFLLog -Message "Collecting DCs Information"
    $DCs = @()
    foreach($DomainName in $Global:Domains) {
        $data = Get-ADDomainController -Server $DomainName -Filter * | select ComputerObjectDN, DefaultPartition, Domain, Enabled, Forest, HostName, InvocationId, IPv4Address, IPv6Address, IsGlobalCatalog, IsReadOnly, LdapPort, Name, NTDSSettingsObjectDN, OperatingSystem, OperatingSystemHotfix, OperatingSystemServicePack, OperatingSystemVersion, @{Name="OperationMasterRoles";Expression = {$_.OperationMasterRoles -join ', '}},     @{Name="Partitions";Expression = {$_.Partitions -join ', '}}, ServerObjectDN, ServerObjectGuid, Site, SslPort
        $DCs += fncObjMerge -DomainName $DomainName -inputobject $data
    }
    Write-RFLLog -Message "Exporting Collected Information"
    $DCs | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\ADDomainController.json' -Force
    #endregion

    #region Optional Features
    Write-RFLLog -Message "Collecting Optional Features Information"
    $data = Get-ADOptionalFeature -Filter * | select Name, @{Name="RequiredForestMode";Expression = {$_.RequiredForestMode.ToString()}}, EnabledScopes, @{Name="Enabled";Expression = {Switch (($_.EnabledScopes).count) {0 {'No'};default {'Yes'};} }}

    Write-RFLLog -Message "Export Collected Information"
    $data | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath .\ForestOptionalFeatures.json -Force
    #endregion

    #region replication sites
    Write-RFLLog -Message "Collecting Replicaiton Sites Information"
    $OutObj = @()
    $data = Get-ADReplicationSite -Filter * -Properties Name, Description, Subnets, createTimeStamp, TopologyCleanupEnabled, TopologyDetectStaleEnabled, RedundantServerTopologyEnabled
    foreach($item in $Data) {
	    $OutObj += $item | select Name, DistinguishedName, Description, @{N="DomainControllers";E={($DCs | where-object {$_.Site -eq $item.Name}).Count}}, TopologyCleanupEnabled, TopologyDetectStaleEnabled, RedundantServerTopologyEnabled, @{Name="Subnets";Expression = {(($_.Subnets | Get-ADReplicationSubnet) | select Name).Name}}, @{Name="SubnetCount";Expression = {$_.Subnets.Count}}, InterSiteTopologyGenerator, ManagedBy, ObjectClass, ObjectGUID, ReplicationSchedule, UniversalGroupCachingRefreshSite, @{Name="Creation Date";Expression = {$_.createTimeStamp.ToShortDateString()}}
    }

    Write-RFLLog -Message "Export Collected Information"
    $OutObj | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath .\ADSites.json -Force
    #endregion

    #region Replication Subnet
    Write-RFLLog -Message "Collecting Replicaiton Subnet Information"
    $data = Get-ADReplicationSubnet -Filter * -Properties Name, Description, Site, createTimeStamp, DistinguishedName
    $OutObj = $data | select Name, DistinguishedName, Description, Site, @{Name="SiteName";Expression = { if ($null -eq $_.Site) { '' } else { Get-ADObject $_.Site | Select-Object -ExpandProperty Name }}}, @{Name="Creation Date";Expression = {$_.createTimeStamp.ToShortDateString()}}

    Write-RFLLog -Message "Export Collected Information"
    $OutObj | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath .\ADSubnets.json -Force
    #endregion

    #region Replication SiteLink
    Write-RFLLog -Message "Collecting Replication SiteLInk Information"
    $data = Get-ADReplicationSiteLink -Filter * -Properties Name, Cost, ReplicationFrequencyInMinutes, InterSiteTransportProtocol, siteList
    $OutObj = $data | select @{Name="Site Link Name";Expression = {$_.Name}}, Cost, @{Name="Replication Frequency";Expression = { "$($_.ReplicationFrequencyInMinutes) min" }}, @{Name="Transport Protocol";Expression = { $_.InterSiteTransportProtocol }}, @{Name="Sites";Expression = { ($_.siteList | Get-ADReplicationSite | select Name).Name }}

    Write-RFLLog -Message "Export Collected Information"
    $OutObj | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath .\ADSiteLinks.json -Force
    #endregion

    #region Replication Partner
    Write-RFLLog -Message "Collecting Replication Partner Information"
    $data = Get-ADReplicationPartnerMetadata -Target * -Scope Server -Partition * -PartnerType Both -ErrorAction Ignore

    Write-RFLLog -Message "Export Collected Information"
    $data | select Server, Partition, Partner, @{Name="PartnerType";Expression = {$_.PartnerType.ToString()}}, @{Name="IntersiteTransportType";Expression = {$_.IntersiteTransportType.ToString()}}, LastReplicationSuccess | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath .\ADReplicationPartnerMetadata.json -Force
    #endregion

    #region AD Domain
    Write-RFLLog -Message "Collecting AD Domain Information"
    $OutObj = @()
    foreach($DomainName in $Global:Domains) {
        $data = Get-ADDomain -Identity $DomainName
        foreach ($Item in $Data) {
            #https://techcommunity.microsoft.com/t5/ask-the-directory-services-team/managing-rid-pool-depletion/ba-p/399736
            $RIDPool = Get-ADObject -server $item.PDCEmulator -Identity "CN=RID Manager$,CN=System,$($item.DistinguishedName)" -Properties rIDAvailablePool -ErrorAction SilentlyContinue
            if ($RIDPool) {
                [int32]$CompleteSIDS = $RIDPool.rIDAvailablePool / ([math]::Pow(2,32))
                [int64]$temp64val = $CompleteSIDS * ([math]::Pow(2,32))
                $RIDsIssued = ($RIDPool.rIDAvailablePool) - $temp64val #not using [int32] as it could return error - InvalidArgument: Cannot convert value "4611686014132446208" to type "System.Int32". Error: "Value was either too large or too small for an Int32."
                $RIDsRemaining = $CompleteSIDS - $RIDsIssued
            }

            $OutObj += [pscustomobject]([ordered] @{
                'SourceDomain' = $DomainName
                'Domain Name' = $Item.Name
                'NetBIOS Name' = $Item.NetBIOSName
                'Distinguished Name' = $Item.DistinguishedName
                'Domain SID' = $Item.DomainSID.Value
                'Domain Functional Level' = $Item.DomainMode.ToString()
                'Domains' = $Item.Domains
                'Forest' = $Item.Forest
                'Parent Domain' = $Item.ParentDomain
                'Replica Directory Servers' = $Item.ReplicaDirectoryServers
                'Child Domains' = $Item.ChildDomains
                'Computers Container' = $Item.ComputersContainer
                'Deleted Objects Container' = $Item.ComputersContainer
                'Domain Controllers Container' = $Item.DomainControllersContainer
                'Systems Container' = $Item.SystemsContainer
                'Users Container' = $Item.UsersContainer
                'ReadOnly Replica Directory Servers' = $Item.ReadOnlyReplicaDirectoryServers
                'ms-DS-MachineAccountQuota' = (Get-ADObject -server $item.PDCEmulator -Identity $item.DistinguishedName -Properties ms-DS-MachineAccountQuota -ErrorAction SilentlyContinue).'ms-DS-MachineAccountQuota'
                'RID Issued' = $RIDsIssued
                'RID Available' = $RIDsRemaining
            })
        }
    }

    Write-RFLLog -Message "Export Collected Information"
    $OutObj | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\ADDomain.json' -Force
    #endregion

    #region Domain Trust
    Write-RFLLog -Message "Collecting Domain Trust Information"
    $OutObj = @()
    foreach($DomainName in $Global:Domains) {
	    $data = Get-ADTrust -Filter * -Server $DomainName | select Name, DistinguishedName, Source, Target, @{Name="Direction";Expression = {$_.Direction.ToString()}}, IntraForest, SelectiveAuthentication, SIDFilteringForestAware, SIDFilteringQuarantined, @{Name="TrustType";Expression = {$_.TrustType.ToString()}}, UplevelOnly
        $OutObj += fncObjMerge -DomainName $DomainName -inputobject $data
    }

    Write-RFLLog -Message "Export Collected Information"
    $OutObj | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\ADDomainTrust.json' -Force
    #endregion

    #region FSMO
    Write-RFLLog -Message "Collecting FSMO Information"
    $OutObj = @()
    foreach($DomainName in $Global:Domains) {
        $data = Get-ADDomain -Identity $DomainName | Select-Object Name,InfrastructureMaster, RIDMaster, PDCEmulator
        $OutObj += fncObjMerge -DomainName $DomainName -inputobject $data
    }

    Write-RFLLog -Message "Export Collected Information"
    $OutObj | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\ADDomainFSMO.json' -Force
    #endregion

    #region Replication Failure
    Write-RFLLog -Message "Collecting Replication Failure Information"
    $OutObj = @()
    foreach($DomainName in $Global:Domains) {
        $data = Get-ADReplicationFailure -Target $DomainName -Scope Forest -ErrorAction SilentlyContinue
        $OutObj += fncObjMerge -DomainName $DomainName -inputobject $data
    }

    Write-RFLLog -Message "Export Collected Information"
    $OutObj | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\ADReplicationFailure.json' -Force
    #endregion 

    #region Global Catalog
    Write-RFLLog -Message "Collecting Global Catalog Information"
    $OutObj = @()
    foreach($DomainName in $Global:Domains) {
        $data = Get-ADDomainController -Server $DomainName -Filter {IsGlobalCatalog -eq "True"}
        $OutObj += fncObjMerge -DomainName $DomainName -inputobject $data
    }

    Write-RFLLog -Message "Export Collected Information"
    $OutObj | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\GlobalCatalog.json' -Force
    #endregion

    #region SPN
    Write-RFLLog -Message "Collecting SPN Information"
    $OutObj = @()
    foreach($DomainName in $Global:Domains) {
        $data = Get-ADObject -LDAPFilter "ServicePrincipalName=*" -Properties ServicePrincipalName -Server $DomainName
        $OutObj += fncObjMerge -DomainName $DomainName -inputobject $data
    }

    Write-RFLLog -Message "Export Collected Information"
    $OutObj | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\SPN.json' -Force
    #endregion

    #region Default Password Policy
    Write-RFLLog -Message "Collecting Default Password Policy Information"
    $OutObj = @()
    foreach($DomainName in $Global:Domains) {
        $Data = Get-ADDefaultDomainPasswordPolicy -Identity $DomainName
        $OutObj += fncObjMerge -DomainName $DomainName -inputobject $data
    }

    Write-RFLLog -Message "Export Collected Information"
    $outobj | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\DefaultPasswordPolicy.json' -Force
    #endregion

    #region Fined Grained Password Policies
    Write-RFLLog -Message "Collecting Fine Grained Password Policies Information"
    $OutObj = @()
    foreach($DomainName in $Global:Domains) {
        $Data = Get-ADFineGrainedPasswordPolicy -Server $DomainName -Filter {Name -like "*"} -Properties AppliesTo, CanonicalName, CN, ComplexityEnabled, Created, Description, Deleted, isDeleted, DistinguishedName, LockoutDuration, LockoutObservationWindow, LockoutThreshold, Name, MaxPasswordAge, MinPasswordAge, MinPasswordLength, objectClass, objectGuid, PasswordHistoryCount, ReversibleEncryptionEnabled -Searchbase (Get-ADDomain -Identity $DomainName).distinguishedName
        $OutObj += fncObjMerge -DomainName $DomainName -inputobject $data
    }

    Write-RFLLog -Message "Export Collected Information"
    $outobj | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\FinedGrainedPasswordPolicies.json' -Force
    #endregion

    #region AD Service Accounts -- todo: correct properties only and not *
    Write-RFLLog -Message "Collecting AD Service Accounts Information"
    $OutObj = @()
    foreach($DomainName in $Global:Domains) {
        $Data = Get-ADServiceAccount -Server $DomainName -Filter * -Properties *
        $OutObj += fncObjMerge -DomainName $DomainName -inputobject $data
    }

    Write-RFLLog -Message "Export Collected Information"
    $outobj | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\GroupManagedServiceAccounts.json' -Force
    #endregion

    #region AD Objects
    Write-RFLLog -Message "Collecting AD Objects Information"
    $OutObj = @()
    foreach($DomainName in $Global:Domains) {
        $Data = Get-ADObject -Server $DomainName -Filter *
        $OutObj += fncObjMerge -DomainName $DomainName -inputobject $data
    }

    Write-RFLLog -Message "Export Collected Information"
    $outobj | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\ADObject.json' -Force
    #endregion

    #region KMS Servers
    Write-RFLLog -Message "Collecting KMS Servers Information"
    $KMSServers = @()
    foreach($DomainName in $Global:Domains) {
        $data = Resolve-DnsName -name "_vlmcs._tcp.$($DomainName)" -Type SRV -ErrorAction SilentlyContinue | Where-Object {$_.Type -eq 'SRV'}
        $KMSServers += fncObjMerge -DomainName $DomainName -inputobject $data
    }

    Write-RFLLog -Message "Export Collected Information"
    $KMSServers | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\KMSServers.json' -Force
    #endregion

    #region KMS Keys
    Write-RFLLog -Message "Collecting KMS Keys Information"
    $OutObj = @()
    foreach ($Item in $KMSServers) {
        $OutObj += Get-WmiObject -ComputerName "$($item.NameTarget)" -Class SoftwareLicensingProduct | select @{Name="SourceDomain";Expression = {$item.SourceDomain}}, @{Name="ServerName";Expression = {$item.NameTarget}}, ApplicationID, Name, Description, ProductKeyChannel, PartialProductKey, RequiredClientCount, GenuineStatus, KeyManagementServiceCurrentCount, KeyManagementServiceFailedRequests, KeyManagementServiceLicensedRequests, KeyManagementServiceLookupDomain, KeyManagementServiceMachine, KeyManagementServiceNonGenuineGraceRequests, KeyManagementServiceNotificationRequests, KeyManagementServiceOOBGraceRequests, KeyManagementServiceOOTGraceRequests, KeyManagementServicePort, KeyManagementServiceProductKeyID, KeyManagementServiceTotalRequests, KeyManagementServiceUnlicensedRequests, LicenseDependsOn, LicenseFamily, LicenseIsAddon, LicenseStatus, LicenseStatusReason
    }

    Write-RFLLog -Message "Export Collected Information"
    $OutObj | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\KMSKeys.json' -Force
    #endregion

    #region DNS Servers
    Write-RFLLog -Message "Collecting DNS Servers Information"
    $DNSServers = @()
    foreach($DomainName in $Global:Domains) {
        $DNSServers += Resolve-DnsName -name "$($DomainName)" -Type NS -ErrorAction SilentlyContinue | Where-Object {$_.Type -eq 'NS'} | select @{Name="DomainName";Expression = { $_.Name }}, @{Name="Type";Expression = { $_.Type.ToString() }}, TTL, @{Name="Section";Expression = { $_.Section.ToString() }}, NameHost, @{Name="ConnectionSuccessful";Expression = { try { $outnull = Get-DNSServer -ComputerName $_.NameHost; $true } catch { $false } }}
    }

    Write-RFLLog -Message "Export Collected Information"
    $DNSServers | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\DNSServers.json' -Force
    #endregion

    #region DNS Fowarder
    Write-RFLLog -Message "Collecting DNS Fowarder Information"
    $OutObj = @()
    foreach ($Item in ($DNSServers | Where-Object {$_.ConnectionSuccessful -eq $true})) {
        $OutObj += Get-DnsServerForwarder -ComputerName $Item.NameHost | select @{Name="Server";Expression = {$item.NameHost}}, @{Name="Domain";Expression = {$item.DomainName}}, UseRootHint, Timeout, EnableReordering, @{Name="IPAddress";Expression = {$_.IPAddress -join ', '}}, @{Name="ReorderedIPAddress";Expression = {$_.ReorderedIPAddress -join ', '}}
    }

    Write-RFLLog -Message "Export Collected Information"
    $OutObj | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\DnsServerForwarder.json' -Force
    #endregion

    #region DNS ServerZone
    Write-RFLLog -Message "Collecting DNS ServerZone Information"
    $OutObj = @()
    foreach ($Item in ($DNSServers | Where-Object {$_.ConnectionSuccessful -eq $true})) {
        $OutObj += Get-DnsServerZone -ComputerName $Item.NameHost | select @{Name="SourceDomain";Expression = {$item.DomainName}}, @{Name="Server";Expression = {$item.NameHost}}, ZoneName, ZoneType, IsAutoCreated, IsDsIntegrated, IsReverseLookupZone, IsSigned
    }

    Write-RFLLog -Message "Export Collected Information"
    $OutObj | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\DnsServerZone.json' -Force
    #endregion

<#
    #region DNS ServerZone
    Write-RFLLog -Message "Collecting DNSServerZone Information"
    'DNS ServerZone'
    $OutObj = @()
    foreach ($Item in ($DNSServers | Where-Object {$_.ConnectionSuccessful -eq $true})) {
        $OutObj += Get-DnsServerZone -ComputerName $Item.NameHost | select @{Name="SourceDomain";Expression = {$item.SourceDomain}}, @{Name="Server";Expression = {$item.NameHost}}, ZoneName, ZoneType, IsAutoCreated, IsDsIntegrated, IsReverseLookupZone, IsSigned
    }

    Write-RFLLog -Message "Export Collected Information"
    $OutObj | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\DnsServerZone.json' -Force
    #endregion
#>
    #region DHCP Servers
    Write-RFLLog -Message "Collecting DHCP Servers Information"
    #$DomainInfo = Get-ADDomain -Identity $Global:RootDomain
    #$DHCPServers = Get-ADObject -Server $Global:RootDomain -SearchBase "cn=configuration,$($DomainInfo.DistinguishedName)" -Filter "objectclass -eq 'dhcpclass' -AND Name -ne 'dhcproot'" | select @{Name="DnsName";Expression = {$_.Name}}, @{Name="IPAddress";Expression = {(Resolve-DnsName $_.Name).IPAddress}}
    $DHCPServers = Get-DhcpServerInDC | select IPAddress, DnsName, @{Name="ConnectionSuccessful";Expression = { try { $outnull = Get-DhcpServerSetting -ComputerName $_.DnsName; $true } catch { $false } }}

    Write-RFLLog -Message "Export Collected Information"
    $DHCPServers | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\DHCPServers.json' -Force
    #endregion

    #region DHCP v4 Scopes
    Write-RFLLog -Message "Collecting DHCP v4 Scopes Information"
    $OutObj = @()
    foreach ($Item in ($DHCPServers | where-object {$_.ConnectionSuccessful -eq $true})) {
        $scopeList = Get-DhcpServerv4Scope –ComputerName $Item.DnsName | Select @{Name="ServerName";Expression = {$Item.DnsName}}, ScopeId, Name, SubnetMask, StartRange, EndRange, LeaseDuration, State
        Foreach ($dhcpScope in $scopeList) {
            $OutObj += Get-DhcpServerv4ScopeStatistics -ComputerName $Item.DnsName -ScopeId $dhcpScope.ScopeId | Select @{Name="ServerName";Expression = {$Item.DnsName}}, @{Name="ScopeId";Expression = {$dhcpScope.ScopeId}}, @{Name="Name";Expression = {$dhcpScope.Name}}, @{Name="SubnetMask";Expression = {$dhcpScope.SubnetMask}}, @{Name="StartRange";Expression = {$dhcpScope.StartRange}}, @{Name="EndRange";Expression = {$dhcpScope.EndRange}}, @{Name="LeaseDuration";Expression = {$dhcpScope.LeaseDuration}}, @{Name="State";Expression = {$dhcpScope.State}}, Free, InUse, Reserved, PercentageInUse
        }
    }

    Write-RFLLog -Message "Export Collected Information"
    $OutObj | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\DhcpServerv4Scopes.json' -Force
    #endregion

    #todo: get scope v6

    #region OUList & Containers
    Write-RFLLog -Message "Collecting OU List and Containers Information"
    $OUList = @()
    foreach($DomainName in $Global:Domains) {
        $data = Get-ADOrganizationalUnit -Server $DomainName -Properties * -Filter * | select ObjectGUID, DistinguishedName, CanonicalName, @{Name="SourceDomain";Expression = {$DomainName}}, @{Name="ObjectType";Expression = {"OrganizationalUnit"}}
        $OUList += $data | foreach-object { New-Object PSObject -Property @{
		    'ObjectGUID' = $_.ObjectGUID
		    'DistinguishedName' = $_.DistinguishedName
		    'CanonicalName' = $_.CanonicalName
		    'SourceDomain' = $_.SourceDomain
		    'ObjectType' = $_.ObjectType
	    }}
    }

    foreach($DomainName in $Global:Domains) {
        $data = Get-ADObject -Server $DomainName -SearchScope OneLevel -Filter 'objectClass -eq "container"' -Properties CanonicalName| select ObjectGUID, DistinguishedName, CanonicalName, @{Name="SourceDomain";Expression = {$DomainName}}, @{Name="ObjectType";Expression = {"Container"}}
        $OUList += $data | foreach-object { New-Object PSObject -Property @{
		    'ObjectGUID' = $_.ObjectGUID
		    'DistinguishedName' = $_.DistinguishedName
		    'CanonicalName' = $_.CanonicalName
		    'SourceDomain' = $_.SourceDomain
		    'ObjectType' = $_.ObjectType
	    }}
    }

    Write-RFLLog -Message "Export Collected Information"
    $OUList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\ADOUList.json' -Force
    #endregion

    #region GPOLinks
    Write-RFLLog -Message "Collecting GPO Links Information"
    $GPOLinkList = @()
    foreach($DomainName in $Global:Domains) {
	    $DomainInfo = Get-ADDomain -Identity $DomainName

	    #Get Links based on AD
	    $GPOLinkList += Get-GPInheritance -Target $DomainInfo.DistinguishedName | select-object -ExpandProperty GpoLinks | select GpoId, DisplayName, Enabled, Enforced, Target, Order
	
	    #Get Links based on OU
	    $GPOLinkList += (Get-ADOrganizationalUnit -Server $DomainName -filter * | Get-GPInheritance).GpoLinks | select GpoId, DisplayName, Enabled, Enforced, Target, Order
    }

    #get  GPOlinks based on sites
    $getADO = @{
	    LDAPFilter = "(Objectclass=site)"
	    properties = "Name"
	    SearchBase = $RootDSE.ConfigurationNamingContext
    }
    $sites = Get-ADObject @getADO #using get-adobject because Get-ADReplicationSite does not have a GetGPOLinks method

    $gpm = New-Object -ComObject "GPMGMT.GPM"
    $gpmConstants = $gpm.GetConstants()

    $gpmdomain = $gpm.GetDomain($Global:ForestName, "", $gpmConstants.UseAnyDC)
    $SiteContainer = $gpm.GetSitesContainer($Global:ForestName, $Global:ForestName, $null, $gpmConstants.UseAnyDC)

    foreach($SiteContainerItem in $SiteContainer) {
	    foreach($SiteItem in $Sites) {
		    $Site = $SiteContainerItem.GetSite($SiteItem.Name)
		    $GPOLinkList += $site.GetGPOLinks() | Select @{Name="GpoId";Expression = {$_.GpoId}},@{Name="DisplayName";Expression = {$gpmdomain.GetGPO($_.gpoid).Displayname}},Enabled,Enforced,@{Name="Target";Expression = {$Site.Path}},@{Name="Order";Expression = {$_.SOMLinkOrder}}
	    }
    }

    Write-RFLLog -Message "Export Collected Information"
    $GPOLinkList | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\ADGPOLink.json' -Force
    #endregion

    #region GPOList
    Write-RFLLog -Message "Collecting GPOList Information"
    $GPOList = @()
    foreach($DomainName in $Global:Domains) {
        $GPOList += Get-GPO -Domain $DomainName -All | select @{Name="SourceDomain";Expression = {$DomainName}}, Id,DisplayName,Path,Owner,DomainName,@{N="User";E={$_.User.Enabled.ToString()}},@{N="Computer";E={$_.Computer.Enabled.ToString()}},@{N="GpoStatus";E={$_.GpoStatus.ToString()}},@{N="WmiFilter";E={$_.WmiFilter.Name}},Description, CreationTime, ModificationTime
    }

    $OutObj = @()
    $arrGPOLinkList = $GPOLinkList | group-object GpoId
    foreach($GPOItem in $GPOList) {
	    $GPOInfo = Get-GPOReport -ReportType xml -Domain $GPOItem.DomainName -Guid $GPOItem.Id

	    $arrGPOLinkItem = $arrGPOLinkList | where-object {$_.Name -eq $GPOItem.Id}
	    $linkCount = $arrGPOLinkItem.Count
	    $LinksDisabledCount = ($arrGPOLinkItem.Group | Where-Object {$_.Enabled -eq $false} | Measure-Object).Count
	    $LinksEnabledCount = ($arrGPOLinkItem.Group | Where-Object {$_.Enabled -eq $true} | Measure-Object).Count

	    $validXML = $false
	    try {
		    $xml = [xml]$GPOInfo
		    $validXML = $true
	    } catch {
            Write-RFLLog -Message ("Invalid GPO XML. {0} - {1}. Error {2}" -f $DomainItem, $GPOItem.DisplayName. $_) -LogLevel 3
		    $ValidXML = $false
	    }
	    $OutObj += New-Object -TypeName PSObject -Property @{
		    'SourceDomain' = $GPOItem.SourceDomain
		    'Computer' = $GPOItem.Computer
		    'CreationTime' = $GPOItem.CreationTime
		    'Description' = $GPOItem.Description
		    'DisplayName' = $GPOItem.DisplayName
		    'DomainName' = $GPOItem.DomainName
		    'GpoStatus' = $GPOItem.GpoStatus
		    'Id' = $GPOItem.Id
		    'ModificationTime' = $GPOItem.ModificationTime
		    'Owner' = $GPOItem.Owner
		    'Path' = $GPOItem.Path
		    'User' = $GPOItem.User
		    'WmiFilter' = $GPOItem.WmiFilter
		    'ValidXML' = $validXML
		    'LinksCount' = $linkCount
		    'LinksDisabledCount' = $LinksDisabledCount
		    'LinksEnabledCount' = $LinksEnabledCount
	    }
    }

    Write-RFLLog -Message "Export Collected Information"
    $OutObj | select SourceDomain,Id,DisplayName,Path,Owner,DomainName,@{N="User";E={$_.User.ToString()}},@{N="Computer";E={$_.Computer.ToString()}},@{N="GpoStatus";E={$_.GpoStatus.ToString()}},@{N="WmiFilter";E={$_.WmiFilter.Name}},Description,@{N="CreationTime";E={[datetime]::parseexact(($_.CreationTime).ToString('dd/MM/yyyy HH:mm:ss'), 'dd/MM/yyyy HH:mm:ss', $null)}},@{N="ModificationTime";E={[datetime]::parseexact(($_.ModificationTime).ToString('dd/MM/yyyy HH:mm:ss'), 'dd/MM/yyyy HH:mm:ss', $null)}},ValidXML,LinksCount,LinksDisabledCount,LinksEnabledCount | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\ADGPO.json' -Force
    #endregion

    #region AD User Priviliged
    Write-RFLLog -Message "Collecting AD User Priviliged Information"
    $OutObj = @()
    foreach($DomainName in $Global:Domains) {
        $Data = Get-ADUser -Server $DomainName -Filter {AdminCount -eq "1"} -Properties CannotChangePassword,PasswordNeverExpires,PasswordNotRequired,pwdLastSet,SmartcardLogonRequired,SIDHistory,lastlogontimestamp,LastLogonDate,AccountExpirationDate,LockedOut
        $OutObj += fncObjMerge -DomainName $DomainName -inputobject $data
    }

    Write-RFLLog -Message "Export Collected Information"
    $outobj | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\ADUserPriviliged.json' -Force
    #endregion

    #region AD User
    Write-RFLLog -Message "Collecting AD User Information"
    $userList = @()
    foreach($DomainName in $Global:Domains) {
        $Data = Get-AdUser -Server $DomainName -Filter * -Properties CannotChangePassword, PasswordNeverExpires, SmartcardLogonRequired, SIDHistory, LastLogonDate, PasswordNotRequired, LockedOut, pwdLastSet, ObjectGUID, GivenName, Sn, SamAccountName, DisplayName, LastLogonTimestamp, PasswordLastSet, Description, Company, Title, Enabled, EmailAddress, Department, info, Office, physicalDeliveryOfficeName, CanonicalName, DistinguishedName, AccountExpirationDate
        $userList += fncObjMerge -DomainName $DomainName -inputobject $data
    }

    Write-RFLLog -Message "Export Collected Information"
    $userList | select CannotChangePassword, PasswordNeverExpires, SmartcardLogonRequired, SIDHistory, LastLogonDate, PasswordNotRequired, LockedOut, pwdLastSet, ObjectGUID, GivenName, Sn, SamAccountName, DisplayName, @{N="LastLogonTimestamp";E={[datetime]::FromFileTime($_.LastLogonTimestamp).ToString('dd/MM/yyyy HH:mm:ss')}}, @{N="PasswordLastSet";E={$_.PasswordLastSet.ToString('dd/MM/yyyy HH:mm:ss')}}, Description, Company, Title, Enabled, EmailAddress, Department, info, Office, physicalDeliveryOfficeName, CanonicalName, DistinguishedName, @{N="AccountExpirationDate";E={[datetime]::parseexact(($_.AccountExpirationDate).ToString('dd/MM/yyyy HH:mm:ss'), 'dd/MM/yyyy HH:mm:ss', $null)}}, @{N="Domain";E={$_.CanonicalName.Substring(0,$_.CanonicalName.IndexOf('/'))}} | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath '.\ADUsers.json' -Force
    #endregion

    #region AD Computer
    Write-RFLLog -Message "Collecting AD Computer Information"
    $ComputerList = @()
    foreach($DomainName in $Global:Domains) {
        $Data = Get-ADComputer -Server $DomainName -Properties ObjectGUID, SamAccountName, LastLogonTimestamp, PasswordLastSet, createTimeStamp, OperatingSystem, OperatingSystemVersion, Description, CanonicalName, DistinguishedName, Enabled -Filter *
        $ComputerList += fncObjMerge -DomainName $DomainName -inputobject $data
    }

    Write-RFLLog -Message "Export Collected Information"
    $ComputerList | select ObjectGUID, SamAccountName, @{N="LastLogonTimestamp";E={[datetime]::FromFileTime($_.LastLogonTimestamp.ToString('dd/MM/yyyy HH:mm:ss'))}}, @{N="PasswordLastSet";E={$_.PasswordLastSet.ToString('dd/MM/yyyy HH:mm:ss')}}, @{N="createTimeStamp";E={$_.createTimeStamp.ToString('dd/MM/yyyy HH:mm:ss')}}, OperatingSystem, OperatingSystemVersion, Description, CanonicalName, DistinguishedName, Enabled, @{N="Domain";E={$_.CanonicalName.Substring(0,$_.CanonicalName.IndexOf('/'))}} | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath .\ADComputers.json -Force
    #endregion

    #region AD Contacts
    Write-RFLLog -Message "Collecting AD Contacts Information"
    $OutObj = @()
    foreach($DomainName in $Global:Domains) {
        $Data = Get-ADObject -Filter 'objectClass -eq "contact"' -Server $DomainName -Properties ObjectGUID, DistinguishedName, Name, CN, Description, CanonicalName, Name, SamAccountName
        $OutObj += fncObjMerge -DomainName $DomainName -inputobject $data
    }

    Write-RFLLog -Message "Export Collected Information"
    $outobj | select ObjectGUID, DistinguishedName, Name, CN, Description, CanonicalName, SamAccountName, @{N="Domain";E={$_.CanonicalName.Substring(0,$_.CanonicalName.IndexOf('/'))}} | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath .\ADContacts.json -Force
    #endregion

    #region AD Printers
    Write-RFLLog -Message "Collecting AD Printers Information"
    $OutObj = @()
    foreach($DomainName in $Global:Domains) {
        $Data = Get-ADObject -Filter 'objectClass -eq "printQueue"' -Server $DomainName -Properties *
        $OutObj += fncObjMerge -DomainName $DomainName -inputobject $data
    }

    Write-RFLLog -Message "Export Collected Information"
    $outobj | Select-Object -Unique | select CanonicalName,CN,@{N="Created";E={[datetime]::parseexact(($_.Created).ToString('dd/MM/yyyy HH:mm:ss'), 'dd/MM/yyyy HH:mm:ss', $null)}},Deleted,Description,DisplayName,DistinguishedName,driverName,location,@{N="Modified";E={[datetime]::parseexact(($_.Modified).ToString('dd/MM/yyyy HH:mm:ss'), 'dd/MM/yyyy HH:mm:ss', $null)}},Name,ObjectGUID,serverName,shortServerName,uNCName,@{N="url";E={$_.url -join ','}} | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath .\ADPrinters.json -Force
    #endregion

    #region AD Groups
    Write-RFLLog -Message "Collecting AD Groups Information"
    $GroupList = @()
    foreach($DomainName in $Global:Domains) {
        $Data = Get-ADGroup -Server $DomainName -Properties ObjectGUID, GroupScope, CanonicalName, CN, Description, DisplayName, DistinguishedName, Name, SamAccountName, Members -Filter *
        $GroupList += fncObjMerge -DomainName $DomainName -inputobject $data
    }

    Write-RFLLog -Message "Export Collected Information"
    $GroupList = $GroupList | select ObjectGUID, CanonicalName, CN, Description, DisplayName, DistinguishedName, Name, SamAccountName, @{N="Domain";E={$_.CanonicalName.Substring(0,$_.CanonicalName.IndexOf('/'))}}, @{N="MemberCount";E={($_.members | Measure-Object).Count}},Members,@{N="MembersEnabled";E={[int]0}}, @{N="MembersDisabled";E={[int]0}}, @{N="MembersUnknown";E={[int]0}}
    #endregion

    #region AD Group Membership
    Write-RFLLog -Message "Collecting AD Group Membership Information"
    $arrGroupList = $GroupList | where-object {$_.MemberCount -gt 0}
    $Total = ($arrGroupList | measure-object).Count

    $GroupMembership = @()
    $i = 1
    foreach($GroupItem in $arrGroupList) {
        Write-RFLLog -Message ('GroupMembership: {0} - {1} of {2}' -f $GroupItem.CanonicalName, $i, $Total)
	    $j = 1
	    $Totalj = $GroupItem.Member.Count	
	    $users = $userList | Where-Object {$_.DistinguishedName -in $GroupItem.Members}
	    $groups = $GroupList | Where-Object {$_.DistinguishedName -in $GroupItem.Members}
	    $computers = $ComputerList | Where-Object {$_.DistinguishedName -in $GroupItem.Members}
		
	    foreach($GroupMembershipItem in $GroupItem.Members) {
		    $user = $users | Where-Object {$_.DistinguishedName -eq $GroupMembershipItem} | Select-Object -First 1
		    $group = $groups | Where-Object {$_.DistinguishedName -eq $GroupMembershipItem} | Select-Object -First 1
		    $computer = $computers | Where-Object {$_.DistinguishedName -eq $GroupMembershipItem} | Select-Object -First 1
		
		    if ($user) {
			    if ($user.CanonicalName) {
				    $MemberDomain = $user.CanonicalName.Substring(0,$user.CanonicalName.IndexOf('/'));
			    } else {
				    $MemberDomain = ''
			    }
			
			    $GroupMembership += New-Object -TypeName PSObject -Property @{
				    'MembershipTotal' = $Totalj
				    'MembershipCount' = $j
				    'GroupCanonicalName' = $GroupItem.CanonicalName;
				    'GroupDomain' = $GroupItem.CanonicalName.Substring(0,$GroupItem.CanonicalName.IndexOf('/'));
				    'MemberCanonicalName' = $user.CanonicalName;
				    'MemberObjectGUID' = $user.ObjectGUID;
				    'MemberSamAccountName' = $user.SamAccountName;
				    'MemberDistinguishedName' = $user.DistinguishedName;
				    'MemberDomain' = $MemberDomain
				    'MemberFound' = $true;
				    'MemberEnabled' = $user.Enabled
				    'MemberType' = 'User'
			    }
		    } elseif ($group) {
			    $GroupMembership += New-Object -TypeName PSObject -Property @{
				    'MembershipTotal' = $Totalj
				    'MembershipCount' = $j
				    'GroupCanonicalName' = $GroupItem.CanonicalName;
				    'GroupDomain' = $GroupItem.CanonicalName.Substring(0,$GroupItem.CanonicalName.IndexOf('/'));
				    'MemberCanonicalName' = $group.CanonicalName;
				    'MemberObjectGUID' = $group.ObjectGUID;
				    'MemberSamAccountName' = $group.SamAccountName;
				    'MemberDistinguishedName' = $group.DistinguishedName;
				    'MemberDomain' = $group.Domain
				    'MemberFound' = $true
				    'MemberEnabled' = $null
				    'MemberType' = 'Group'
			    }
		    } elseif ($computer) {
			    if ($computer.CanonicalName) {
				    $MemberDomain = $computer.CanonicalName.Substring(0,$computer.CanonicalName.IndexOf('/'));
			    } else {
				    $MemberDomain = ''
			    }
		
			    $GroupMembership += New-Object -TypeName PSObject -Property @{
				    'MembershipTotal' = $Totalj
				    'MembershipCount' = $j
				    'GroupCanonicalName' = $GroupItem.CanonicalName;
				    'GroupDomain' = $GroupItem.CanonicalName.Substring(0,$GroupItem.CanonicalName.IndexOf('/'));
				    'MemberCanonicalName' = $computer.CanonicalName;
				    'MemberObjectGUID' = $computer.ObjectGUID;
				    'MemberSamAccountName' = $computer.SamAccountName;
				    'MemberDistinguishedName' = $computer.DistinguishedName;
				    'MemberDomain' = $MemberDomain
				    'MemberFound' = $true
				    'MemberEnabled' = $computer.Enabled
				    'MemberType' = 'Computer'
			    }
		    } else {
			    $GroupMembership += New-Object -TypeName PSObject -Property @{
				    'MembershipTotal' = $Totalj
				    'MembershipCount' = $j
				    'GroupCanonicalName' = $GroupItem.CanonicalName;
				    'GroupDomain' = $GroupItem.CanonicalName.Substring(0,$GroupItem.CanonicalName.IndexOf('/'));
				    'MemberCanonicalName' = '';
				    'MemberObjectGUID' = '';
				    'MemberSamAccountName' = '';
				    'MemberDistinguishedName' = $GroupMembershipItem
				    'MemberDomain' = '';
				    'MemberFound' = $false;
				    'MemberType' = 'Unknown'
			    }
		    }
		    $j++
	    }
	    $i++
    }

    $arrGroupMembership = $GroupMembership | group-object GroupCanonicalName

    $i = 1
    Write-RFLLog -Message ('Checing number of members on each group')
    $GroupList | where-object {$_.MemberCount -gt 0} | foreach-object {
	    $GroupItem = $_
        Write-RFLLog -Message ('Group: {0} - {1} of {2}' -f $GroupItem.CanonicalName, $i, $Total)	    
	    $groups = $arrGroupMembership | where-object {$_.Name -eq $GroupItem.CanonicalName}
	
	    $_.MembersUnknown = ($groups.Group | Where-Object {$_.MemberFound -eq $false} | Measure-Object).Count
	    $_.MembersEnabled = ($groups.Group | Where-Object {($_.MemberFound -eq $true) -and ($_.MemberEnabled -eq $true)} | Measure-Object).Count
	    $_.MembersDisabled = ($groups.Group | Where-Object {($_.MemberFound -eq $true) -and ($_.MemberEnabled -eq $false)} | Measure-Object).Count
	    $i++
    }

    Write-RFLLog -Message "Export Exporting Groups (with members) Information"
    $GroupList | select ObjectGUID, CanonicalName, CN, Description, DisplayName, DistinguishedName, Name, SamAccountName, @{N="Domain";E={$_.CanonicalName.Substring(0,$_.CanonicalName.IndexOf('/'))}}, @{N="MemberCount";E={($_.members | Measure-Object).Count}},MembersUnknown,MembersEnabled,MembersDisabled,Members | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath .\GroupList.json -Force

    Write-RFLLog -Message "Export Exporting Groups (no members) Information"
    $GroupList | select ObjectGUID, @{N="GroupScope";E={$_.GroupScope.ToString()}}, CanonicalName, CN, Description, DisplayName, DistinguishedName, Name, SamAccountName, @{N="Domain";E={$_.CanonicalName.Substring(0,$_.CanonicalName.IndexOf('/'))}}, @{N="MemberCount";E={($_.members | Measure-Object).Count}},MembersUnknown,MembersEnabled,MembersDisabled | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath .\GroupListNoMembers.json -Force

    Write-RFLLog -Message "Export Exporting Group Membership Information"
    $GroupMembership | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath .\GroupMembership.json -Force
    #endregion

    #region Report
    Write-RFLLog -Message "Collecting AD Usage Report Information"
    $i = 1
    $Total = $OUList.Count
    $Report = @()
    foreach($OUItem in $OUList) {
        Write-RFLLog -Message ('OU: {0} - {1} of {2}' -f $OUItem.CanonicalName, $i, $Total)	    
	    $DomainName = $OUItem.SourceDomain 
	    $users = Get-AdUser -Server $DomainName -Properties ObjectGUID, GivenName, Sn, SamAccountName, DisplayName, LastLogonTimestamp, PasswordLastSet, Description, Company, Title, Enabled, EmailAddress, Department, info, Office, physicalDeliveryOfficeName, CanonicalName, DistinguishedName, AccountExpirationDate -Filter * -Searchbase $OUItem.DistinguishedName -SearchScope OneLevel
	    $ChildOUs = (Get-ADOrganizationalUnit -Server $DomainName -SearchBase $OUItem.DistinguishedName -Filter 'ObjectClass -eq "organizationalUnit"' -SearchScope OneLevel | Measure-Object).Count
	    $Computers = Get-ADComputer -Server $DomainName -Properties ObjectGUID, SamAccountName, LastLogonTimestamp, PasswordLastSet, createTimeStamp, OperatingSystem, OperatingSystemVersion, Description, CanonicalName, DistinguishedName, Enabled -Filter * -Searchbase $OUItem.DistinguishedName -SearchScope OneLevel
	    $Contacts = Get-ADObject -Filter 'objectClass -eq "contact"' -Server $DomainName -Properties ObjectGUID, DistinguishedName, Name, CN, Description, CanonicalName, Name, SamAccountName -Searchbase $OUItem.DistinguishedName -SearchScope OneLevel
	    $Groups = Get-ADGroup -Server $DomainName -Properties ObjectGUID, GroupScope, CanonicalName, CN, Description, DisplayName, DistinguishedName, Name, SamAccountName, Member -Filter * -Searchbase $OUItem.DistinguishedName -SearchScope OneLevel
	    $LinkedGPOs = $GPOLinkList | where-object {$_.Target -eq $OUItem.DistinguishedName} | select GpoId, DisplayName, Enabled, Enforced, Target, GPODomainName
	    $Printers = Get-ADObject -Filter 'objectClass -eq "printQueue"' -Server $DomainName -Properties * -Searchbase $OUItem.DistinguishedName
	
	    #Generate report
	    $Report += New-Object -TypeName PSObject -Property @{
		    'OUName' = $OUItem.CanonicalName
		    'ChildOUs' = [int]$ChildOUs
		    'Accounts' = ($users | Measure-Object).Count
		    'Accounts Active' = ($users | where-object {$_.Enabled -eq $true} | Measure-Object).Count
		    'Accounts Inactive' = ($users | where-object {$_.Enabled -eq $false} | Measure-Object).Count
		    'Computers' = ($Computers | Measure-Object).Count;
		    'Computers Active' = ($Computers | where-object {$_.Enabled -eq $true} | Measure-Object).Count
		    'Computers Inactive' = ($Computers | where-object {$_.Enabled -eq $false} | Measure-Object).Count
		    'Contacts' = ($Contacts | Measure-Object).Count
		    'Groups' = ($Groups | Measure-Object).Count
		    'GPOs' = ($LinkedGPOs | Measure-Object).Count
		    'GPOs Active' = ($LinkedGPOs | where-object {$_.Enabled -eq $true} | Measure-Object).Count
		    'GPOs Inactive' = ($LinkedGPOs | where-object {$_.Enabled -eq $false} | Measure-Object).Count
		    'GPOs Enforced' = ($LinkedGPOs | where-object {$_.Enforced -eq $true} | Measure-Object).Count
		    'Printers' = ($Printers | Measure-Object).Count;
		    'Printers Deleted' = ($Printers | where-object {$_.Deleted -eq $true} | Measure-Object).Count
	    }
	    $i++
    }

    Write-RFLLog -Message "Export Collected Information"
    $Report | select 'OUName', 'ChildOUs', 'Contacts', 'Accounts', 'Accounts Active', 'Accounts Inactive', 'Computers', 'Computers Active', 'Computers Inactive', 'Groups', 'GPOs', 'GPOs Active', 'GPOs Inactive', 'GPOs Enforced', Printers, 'Printers Deleted', @{N="Domain";E={$_.OUName.Substring(0,$_.OUName.IndexOf('/'))}} | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath .\ADReport.json -Force
    #endregion

    Write-RFLLog -Message "All Data have been collected and saved to '$($SaveTo)'" -LogLevel 2
	Write-RFLLog -Message "You can now use the 'ExportADData.ps1' PowerShell script to create the report file" -LogLevel 2
} catch {
    Write-RFLLog -Message "An error occurred $($_)" -LogLevel 3
    Exit 3000
} finally {
    Set-Location $Script:CurrentFolder
    Write-RFLLog -Message "*** Ending ***"
}
#endregion