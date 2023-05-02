<#
    .SYSNOPSIS
        Export Active Directory data collected using the ColectADData.ps1 to a word or html format

    .DESCRIPTION
        Export Active Directory data collected using the ColectADData.ps1 to a word or html format

    .PARAMETER SourceData
        Where the .json files are located

    .PARAMETER OutputFormat
        Format to export. Possible options are Word and HTML

    .PARAMETER DormantDays
        number of days to classify an object as dormant

    .PARAMETER PasswordDays
        how often an computer object must change its password

    .PARAMETER CompanyName
        Company Name to be added onto the report's header

    .PARAMETER CompanyWeb
        Company URL to be added onto the report's header

    .PARAMETER CompanyEmail
        Company E-mail to be added onto the report's header

    .NOTES
        Name: ExportADData.ps1
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
        .\ExportADData.ps1 -SourceData 'c:\temp\CollectADData' -OutputFormat @('Word')
        .\ExportADData.ps1 -SourceData 'c:\temp\CollectADData' -OutputFormat @('Word') -DormantDays 90 -PasswordDays 30 
        .\ExportADData.ps1 -SourceData 'c:\temp\CollectADData' -OutputFormat @('Word') -DormantDays 90 -PasswordDays 30 -CompanyName 'RFL Systems' -CompanyWeb 'www.rflsystems.co.uk' -CompanyEmail 'team@rflsystems.co.uk'

#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = 'Please provide the path to where the json files are located')]
    [ValidateNotNullOrEmpty()]
	[ValidateScript({ if (Test-Path $_ -PathType 'Container') { $true } else { throw "$_ is not a valid folder path" }  })]
    [String] $SourceData,

    [Parameter(Mandatory = $true, HelpMessage = 'Please provide the format you wish to export the report to')]
    [ValidateNotNullOrEmpty()]
    [ValidateSet('Word', 'HTML')]
    [string[]] $OutputFormat,

    [Parameter(Mandatory = $false, HelpMessage = 'Please provide the number of days to classify an object as dormant')]
    [int] $DormantDays = 90,

    [Parameter(Mandatory = $false, HelpMessage = 'Please provide how often an computer object must change its password')]
    [int] $PasswordDays = 30,

    [Parameter(Mandatory = $false, HelpMessage = 'Please provide the company name')]
    [string] $CompanyName = '',
    [Parameter(Mandatory = $false, HelpMessage = 'Please provide the company web')]
    [string] $CompanyWeb = '',
    [Parameter(Mandatory = $false, HelpMessage = 'Please provide the company email')]
    [string] $CompanyEmail = ''
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

#endregion

#region Variables
$script:ScriptVersion = '0.1'
$script:LogFilePath = $env:Temp
$Script:LogFileFileName = 'ExportADData.log'
$script:ScriptLogFilePath = "$($script:LogFilePath)\$($Script:LogFileFileName)"
$Script:Modules = @('PScribo')
$Script:CurrentFolder = (Get-Location).Path
$Script:dormanttime = ((Get-Date).AddDays($DormantDays*-1)).Date
$Script:passwordtime = ((Get-Date).AddDays($PasswordDays*-1)).Date
$Script:OutputFolderPath = $SourceData
$Global:ExecutionTime = Get-Date
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

    #if (-not (Test-RFLAdministrator)) {
    #    throw "The requested operation requires elevation: Run PowerShell console as administrator"
    #}

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
            Write-RFLLog -Message "    Module $($_) installed. Type: '$($Module.ModuleTYpe)', Verison: '$($Module.Version)', Path: '$($Module.ModuleBase)'"
        } 
    }
    if (-not $Continue) {
        throw "The requested operation requires missing PowerShell Modules. Install the missing PowerShell modules and try again"
    }

    Write-RFLLog -Message "Current Folder '$($Script:CurrentFolder)'"
    Set-Location $SourceData

    Write-RFLLog -Message "All checks completed successful. Starting collecting data for report"

    #region main script
$class = @"
    public class Num2Word
    {
        public static string NumberToText( int n)
          {
           if ( n < 0 )
              return "Minus " + NumberToText(-n);
           else if ( n == 0 )
              return "";
           else if ( n <= 19 )
              return new string[] {"One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", 
                 "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", 
                 "Seventeen", "Eighteen", "Nineteen"}[n-1] + " ";
           else if ( n <= 99 )
              return new string[] {"Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", 
                 "Eighty", "Ninety"}[n / 10 - 2] + " " + NumberToText(n % 10);
           else if ( n <= 199 )
              return "One Hundred " + NumberToText(n % 100);
           else if ( n <= 999 )
              return NumberToText(n / 100) + "Hundreds " + NumberToText(n % 100);
           else if ( n <= 1999 )
              return "One Thousand " + NumberToText(n % 1000);
           else if ( n <= 999999 )
              return NumberToText(n / 1000) + "Thousands " + NumberToText(n % 1000);
           else if ( n <= 1999999 )
              return "One Million " + NumberToText(n % 1000000);
           else if ( n <= 999999999)
              return NumberToText(n / 1000000) + "Millions " + NumberToText(n % 1000000);
           else if ( n <= 1999999999 )
              return "One Billion " + NumberToText(n % 1000000000);
           else 
              return NumberToText(n / 1000000000) + "Billions " + NumberToText(n % 1000000000);
        }
    }
"@
    Add-Type -TypeDefinition $class

    $Global:RootDSE = get-content .\RootDSE.json | convertfrom-json -AsHashtable
    $Global:ForestName = $RootDSE.ldapServiceName.Split(':')[0]
    $Global:ReportFile = '{0}' -f $Global:ForestName
	Write-RFLLog -Message "Report file name: $($Global:ReportFile)"

    $data = get-content .\ADForest.json | convertfrom-json
    $Global:RootDomain = $data.'Forest Name'.toUpper()
    $Global:Domains = @()
    $Global:Domains += $Global:RootDomain.tolower()
    $Global:Domains += $data.Domains.Split(';').Trim() | Where-Object {$_ -ne $Global:RootDomain}

    #region Report
    $Global:WordReport = Document $Global:ReportFile {
        #region style
        DocumentOption -EnableSectionNumbering -PageSize A4 -DefaultFont 'Arial' -MarginLeftAndRight 71 -MarginTopAndBottom 71 -Orientation Portrait
        Style -Name 'Title' -Size 24 -Color '0076CE' -Align Center
        Style -Name 'Title 2' -Size 18 -Color '00447C' -Align Center
        Style -Name 'Title 3' -Size 12 -Color '00447C' -Align Left
        Style -Name 'Heading 1' -Size 16 -Color '00447C'
        Style -Name 'Heading 2' -Size 14 -Color '00447C'
        Style -Name 'Heading 3' -Size 12 -Color '00447C'
        Style -Name 'Heading 4' -Size 11 -Color '00447C'
        Style -Name 'Heading 5' -Size 11 -Color '00447C'
        Style -Name 'Heading 6' -Size 11 -Color '00447C'
        Style -Name 'Normal' -Size 10 -Color '565656' -Default
        Style -Name 'Caption' -Size 10 -Color '565656' -Italic -Align Left
        Style -Name 'Header' -Size 10 -Color '565656' -Align Center
        Style -Name 'Footer' -Size 10 -Color '565656' -Align Center
        Style -Name 'TOC' -Size 16 -Color '00447C'
        Style -Name 'TableDefaultHeading' -Size 10 -Color 'FAFAFA' -BackgroundColor '0076CE'
        Style -Name 'TableDefaultRow' -Size 10 -Color '565656'
        Style -Name 'Critical' -Size 10 -BackgroundColor 'F25022'
        Style -Name 'Warning' -Size 10 -BackgroundColor 'FFB900'
        Style -Name 'Info' -Size 10 -BackgroundColor '00447C'
        Style -Name 'OK' -Size 10 -BackgroundColor '7FBA00'

        Style -Name 'HeaderLeft' -Size 10 -Color '565656' -Align Left -BackgroundColor BDD6EE
        Style -Name 'HeaderRight' -Size 10 -Color '565656' -Align Right -BackgroundColor E7E6E6
        Style -Name 'FooterRight' -Size 10 -Color '565656' -Align Right -BackgroundColor BDD6EE
        Style -Name 'FooterLeft' -Size 10 -Color '565656' -Align Left -BackgroundColor E7E6E6
        Style -Name 'TitleLine01' -Size 18 -Color '565656' -Align Left -BackgroundColor BDD6EE
        Style -Name 'TitleLine02' -Size 10 -Color '565656' -Align Left -BackgroundColor BDD6EE
        Style -Name '1stPageRowStyle' -Size 10 -Color '565656' -Align Left -BackgroundColor E7E6E6

        # Configure Table Styles
        $TableDefaultProperties = @{
            Id = 'TableDefault'
            HeaderStyle = 'TableDefaultHeading'
            RowStyle = 'TableDefaultRow'
            BorderColor = '0076CE'
            Align = 'Left'
            CaptionStyle = 'Caption'
            CaptionLocation = 'Below'
            BorderWidth = 0.25
            PaddingTop = 1
            PaddingBottom = 1.5
            PaddingLeft = 2
            PaddingRight = 2
        }

        TableStyle @TableDefaultProperties -Default
        TableStyle -Name Borderless -HeaderStyle Normal -RowStyle Normal -BorderWidth 0
        TableStyle -Name 1stPageTitle -HeaderStyle Normal -RowStyle 1stPageRowStyle -BorderWidth 0

        # Microsoft AD Cover Page Layout
        # Header & Footer
        Header -FirstPage {
            $Obj = [ordered] @{
                "CompanyName" = $CompanyName
                "CompanyWeb" = $CompanyWeb
                "CompanyEmail" = $CompanyEmail
            }
            [pscustomobject]$Obj | Table -Style Borderless -list -ColumnWidths 50, 50 
        }

        Header -Default {
            $hashtableArray = @(
                [Ordered] @{ "Private and Confidential" = "Active Directory"; '__Style' = 'HeaderLeft'; "Private and Confidential__Style" = 'HeaderRight';}
            )
            Table -Hashtable $hashtableArray -Style Borderless -ColumnWidths 30, 70 -list
        }

        Footer -Default {
            $hashtableArray = @(
                [Ordered] @{ " " = 'Page <!# PageNumber #!> of <!# TotalPages #!>'; '__Style' = 'FooterLeft'; " __Style" = 'FooterRight';}
            )
            Table -Hashtable $hashtableArray -Style Borderless -ColumnWidths 30, 70 -list
        }

        BlankLine -Count 11
        $LineCount = 32 + $LineCount

        # Microsoft Logo Image
        Try {
            Image -Text 'Microsoft Logo' -Align 'Center' -Percent 20 -Base64 "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"
            BlankLine -Count 2
        } Catch {
            Write-RFLLog -Message ".NET Core is required for cover page image support. Please install .NET Core from https://dotnet.microsoft.com/en-us/download" -LogLevel 3
        }

        # Add Report Name
        $Obj = [ordered] @{
            " " = ""
            " __Style" = "TitleLine02"
            "  " = "AD Report"
            "  __Style" = "TitleLine01"
            "   " = "Report Generated on $($Global:ExecutionTime.ToString("dd/MM/yyyy")) and $($Global:ExecutionTime.ToString("HH:mm:ss"))"
            "   __Style" = "TitleLine02"
            "    " = ""
            "    __Style" = "TitleLine02"
        }
        [pscustomobject]$Obj | Table -Style 1stPageTitle -list -ColumnWidths 10, 90 
        PageBreak

        # Add Table of Contents
        TOC -Name 'Table of Contents'
        PageBreak

        #region Executive Summary
        $sectionName = 'Introduction'
        Write-RFLLog -Message "Export $($SectionName)"
        Section -Style Heading2 $sectionName {
	        try {
		        Paragraph "This document describes the overall configuration of the $($Global:ForestName) forest."
		        BlankLine
                PageBreak
	        }
	        catch {
                Write-RFLLog -Message "An error occurred $($_)" -LogLevel 3
	        }
        }
        #endregion
        #endregion

        #region AD Forest
        $SectionName = "Active Directory Forest"
        Write-RFLLog -Message "Export $($SectionName)"
        Section -Style Heading1 $SectionName {
            $OutObj = get-content .\ADForest.json | convertfrom-json
            $childCount = ([int]$OutObj.'Domains Count')-1
            $childDomains = ($OutObj.Domains.Split(';').trim() | Where-Object {$_ -ne $outobj.'Forest Name'}) -join ', '
            if ($childCount -eq 0) {
                $textParagraph = "The $($outobj.'ForestNetbiosName') forest comprises one stand alone domain with no child domain. The forest root domain is $($OutObj.'Forest Name')."
            } else {
                $textParagraph = "The $($outobj.'ForestNetbiosName') forest comprises one forest root domain and $([Num2Word]::NumberToText($childCount)) child domain. The forest root domain is $($OutObj.'Forest Name') and the child "
                if ($childCount -eq 1) {
                    $textParagraph += "domain is $($childDomains)."
                } else {
                    $textParagraph += "domains are $($childDomains)."
                }
            }

            Paragraph $textParagraph
		    BlankLine

            #region ADForest
            Write-RFLLog -Message "Export ADForest"
            $TableParams = @{
                Name = "Forest Summary - $($OutObj.'Forest Name')"
                List = $true
                ColumnWidths = 40, 60
            }
            $TableParams['Caption'] = "- $($TableParams.Name)"
            $OutObj | Table @TableParams
            #endregion

            #region Optional Features
            $sectionName = 'Optional Features'
            Write-RFLLog -Message "Export $($SectionName)"
	        Section -Style Heading2 $sectionName { 
                $OutObj = get-content .\ForestOptionalFeatures.json | convertfrom-json 

                $TableParams = @{
                    Name = "Optional Features - $($Global:RootDomain)"
                    List = $false
                    ColumnWidths = 40, 30, 30
                }
                $TableParams['Caption'] = "- $($TableParams.Name)"
                $OutObj | select Name, RequiredForestMode, Enabled | Table @TableParams
            }
            #endregion 

            #todo: pictures of each domain with the domain controllers
        }
        #endregion

        #region Domain Controllers
        $SectionName = "Domain Controllers"
        Write-RFLLog -Message "Export $($SectionName)"
        Section -Style Heading1 $SectionName {
            Paragraph "The following table lists the domain controllers in the forest:"
            $OutObj = get-content '.\ADDomainController.json' | convertfrom-json | select Domain, Name, IPv4Address, IsGlobalCatalog, IsReadOnly, OperatingSystem, Site

            $TableParams = @{
                Name = "$sectionName"
                List = $false
            }
            $TableParams['Caption'] = "- $($TableParams.Name)"
            $OutObj | Table @TableParams

            #region Domain Controller Object Count
            $sectionName = 'Domain Controller Object Count'
            Write-RFLLog -Message "Export $($SectionName)"
	        Section -Style Heading2 $sectionName { 
                $OutObj = get-content '.\ADDomainController.json' | convertfrom-json
                $TableOut = @()
                foreach($item in ($OutObj | Group-Object SourceDomain)) {
                    $TableOut += $item | select Name, Count, @{Name="Global Catalogs";Expression = { ($item.Group | Where-Object {$_.IsGlobalCatalog -eq $true}).Count }}, @{Name="Read Only";Expression = { ($item.Group | Where-Object {$_.IsReadOnly -eq $true}).Count }}
                }
                $TableParams = @{
                    Name = $sectionName
                    List = $false
                }
                $TableParams['Caption'] = "- $($TableParams.Name)"
                $TableOut | Table @TableParams
            }
            #endregion 

            #region FSMO
            $sectionName = 'FSMO (Flexible Single Master Operations)'
            Write-RFLLog -Message "Export $($SectionName)"
	        Section -Style Heading2 $sectionName { 
                $OutObj = @()
                $data = get-content .\ADForest.json | convertfrom-json
                $OutObj += New-Object PSObject -Property @{
                    Domain = $data.'Forest Name'.tolower()
                    Role = 'Schema Master'
                    Hoster = $data.'Schema Master'
                }

                $OutObj += New-Object PSObject -Property @{
                    Domain = $data.'Forest Name'.tolower()
                    Role = 'Domain Naming Master'
                    Hoster = $data.'Domain Naming Master'
                }

                get-content '.\ADDomainFSMO.json' | convertfrom-json | ForEach-Object {
                    $OutObj += New-Object PSObject -Property @{
                        Domain = $_.SourceDomain.tolower()
                        Role = 'PDC Emulator'
                        Hoster = $_.PDCEmulator
                    }
                    $OutObj += New-Object PSObject -Property @{
                        Domain = $_.SourceDomain.tolower()
                        Role = 'Infrastructure Master'
                        Hoster = $_.InfrastructureMaster
                    }
                    $OutObj += New-Object PSObject -Property @{
                        Domain = $_.SourceDomain.tolower()
                        Role = 'RID Master'
                        Hoster = $_.RIDMaster
                    }
                }

                $TableParams = @{
                    Name = "FSMO Server"
                    List = $false
                }
                $TableParams['Caption'] = "- $($TableParams.Name)"
                $OutObj | select Domain, Role, Hoster | Sort-Object -Property 'Domain', Role | Table @TableParams
            }
            #endregion
        }
        #endregion

        #region Domains and Trusts
        $SectionName = "Domain and Trusts"
        Section -Style Heading1 $SectionName {
            $OutObj = get-content '.\ADDomainTrust.json' | convertfrom-json | select SourceDomain, Name, Direction

            if ($null -ne $OutObj) {
                foreach($item in ($OutObj | Group-Object SourceDomain)) {
                    $sectionName = $item.Name
                    Write-RFLLog -Message "Export $($SectionName)"
	                Section -Style Heading2 $sectionName { 
                        Paragraph "The following table illustrates the Active Directory trust relationships in $($sectionName)"

                        $TableParams = @{
                            Name = "Domain and Trusts - $($item)"
                            List = $false
                        }
                        $TableParams['Caption'] = "- $($TableParams.Name)"
                        $item.Group | select Name, Direction | Table @TableParams
                    }
                }
                ##to-do picture of the AD Trusts
            } else {
                Paragraph "No $($SectionName) found"
            }
        }
        #endregion 

        #region Active Directory Site Topology
        $SectionName = "Active Directory Site Topology"
        Section -Style Heading1 $SectionName {

            #region Sites
            $sectionName = 'Active Directory Sites'
            Write-RFLLog -Message "Export $($SectionName)"
	        Section -Style Heading2 $sectionName { 
                Paragraph "The following table illustrates the Active Directory sites in $($Global:ForestName)"
                $OutObj = get-content '.\ADSites.json' | convertfrom-json | select Name, @{N="Total DCs";E={$_.DomainControllers}}, @{N="Cleanup";E={$_.TopologyCleanupEnabled}}, @{N="Stale Link Detection";E={$_.TopologyDetectStaleEnabled}}, @{N="Redundant Server Topology";E={$_.RedundantServerTopologyEnabled}}, @{N="Subnets";E={$_.Subnets.Value}}
                $TableParams = @{
                    Name = $sectionName
                    List = $false
                }
                $TableParams['Caption'] = "- $($TableParams.Name)"
                $OutObj | Table @TableParams
            }
            #endregion

            #region SiteLinks
            $OutObj = get-content '.\ADSiteLinks.json' | convertfrom-json | select 'Site Link Name', Cost, 'Replication Frequency', 'Transport Protocol', @{N="Sites";E={$_.Sites.Value -join ', '}}
            $sectionName = 'Active Directory Site Links'
            Write-RFLLog -Message "Export $($SectionName)"
            if ($null -ne $OutObj) {
	            Section -Style Heading2 $sectionName { 
                    Paragraph "The following table illustrates the Active Directory sites in $($Global:ForestName)"

                    $TableParams = @{
                        Name = $sectionName
                        List = $false
                    }
                    $TableParams['Caption'] = "- $($TableParams.Name)"
                    $OutObj | Table @TableParams
                }
            } else {
                Paragraph "No $($SectionName) found"
            }
            #endregion

            ##to-do picture of the AD Site Topology
        }
        #endregion 

        #region AD Replication
        $SectionName = "Active Directory Replication"
        Section -Style Heading1 $SectionName {
            Paragraph "The following table illustrates the Active Directory Replication in $($Global:ForestName)"

            $replicationPartner = get-content '.\ADReplicationPartnerMetadata.json' | convertfrom-json
            $dcs = get-content '.\ADDomainController.json' | convertfrom-json | select Domain, Name, IPv4Address, IsGlobalCatalog, IsReadOnly, OperatingSystem, Site, NTDSSettingsObjectDN
            $replicationFailures += get-content '.\ADReplicationFailure.json' | convertfrom-json

            $outobj = @()
            foreach($dataItem in $replicationPartner) {
                $outobj += New-Object -TypeName PSObject -Property @{
		            'Server' = $dataItem.Server
		            'Site' = ($dcs | Where-Object {$dataItem.Server -eq "$($_.Name).$($_.Domain)"}).Site.Trim()
		            'Partition' = $dataItem.Partition
		            'Partner' = $dataItem.Partner.Split(',')[1].Replace('CN=','')
		            'PartnerSite' = ($dcs | Where-Object {$_.NTDSSettingsObjectDN -eq $dataItem.Partner}).Site
		            'PartnerType' = $dataItem.PartnerType
		            'IntersiteTransportType' = $dataItem.IntersiteTransportType
		            'LastReplicationSuccess' = $dataItem.LastReplicationSuccess
		            'FailureCount' = ($replicationFailures | Where-Object {($_.Server -eq $dataItem.Server) -and ($_.Partner -eq $dataItem.Partner)}).FailureCount
	            }
            }

            if ($null -ne $OutObj) {
                $TableParams = @{
                    Name = $SectionName
                    List = $false
                }
                $TableParams['Caption'] = "- $($TableParams.Name)"
                $OutObj | Sort-Object -Property 'Server', Partition, Partner | Table @TableParams
            } else {
                Paragraph "No $($SectionName) found"
            }
        }
        #endregion

        #region DNS Information
        $SectionName = "DNS Information"
        Section -Style Heading1 $SectionName {
            Paragraph "The following table lists the domain DNS Servers in the forest:"

            #region DNS Servers
            $OutObj = get-content '.\DNSServers.json' | convertfrom-json | select @{Name="Domain";Expression = { $_.DomainName }}, @{Name="Host";Expression = { $_.NameHost }}, @{Name="Connection Successful";Expression = { $_.ConnectionSuccessful }}, Type, TTL, Section
            if ($null -ne $OutObj) {
                $TableParams = @{
                    Name = "$sectionName"
                    List = $false
                }
                $TableParams['Caption'] = "- $($TableParams.Name)"
                $OutObj | select Domain, Host, 'Connection Successful' | Table @TableParams
            } else {
                Paragraph "No $($SectionName) found"
            }
            #endregion

            #region DNS Conditional Forwarders
            $sectionName = 'DNS Conditional Forwarders'
            Write-RFLLog -Message "Export $($SectionName)"
	        Section -Style Heading2 $sectionName {
                Paragraph "The following tables contains the Active Directory Servers DNS conditional forwarders configured"

                $OutObj = get-content '.\DnsServerForwarder.json' | convertfrom-json | select @{Name="Domain";Expression = { $_.domain.tolower() }}, Server, UseRootHint, Timeout, IPAddress 
                if ($null -ne $OutObj) {
                    $TableParams = @{
                        Name = $sectionName
                        List = $false
                    }
                    $TableParams['Caption'] = "- $($TableParams.Name)"
                    $OutObj | Sort-Object -Property 'Domain', Server | Table @TableParams
                } else {
                    Paragraph "No $($SectionName) found"
                }
            }
            #endregion

            #region DNS Zones
            $sectionName = 'DNS Zones'
            Write-RFLLog -Message "Export $($SectionName)"
	        Section -Style Heading2 $sectionName {
                Paragraph "The following tables contains the Active Directory Zones configured"
                $OutObj = get-content '.\DnsServerZone.json' | convertfrom-json | select @{Name="Domain";Expression = { $_.SourceDomain }}, Server, ZoneName, ZoneType, IsAutoCreated, IsDsIntegrated, IsReverseLookupZone, IsSigned
                if ($null -ne $OutObj) {
                    $TableParams = @{
                        Name = $sectionName
                        List = $false
                    }
                    $TableParams['Caption'] = "- $($TableParams.Name)"
                    $OutObj | Sort-Object -Property Domain, Server, ZoneType | Table @TableParams
                } else {
                    Paragraph "No $($SectionName) found"
                }
            }
            #endregion
        }
        #endregion

        #region Group Policy Objects
        $SectionName = "Group Policy Objects"
        Section -Style Heading1 $SectionName {
            #Paragraph ""
            $OutObj = get-content '.\ADGPO.json' | convertfrom-json
            foreach($item in ($OutObj | Group-Object SourceDomain)) {
                #region All GPOS
                $sectionName = "$($item.Name) - All GPOs"
                Write-RFLLog -Message "Export $($SectionName)"
	            Section -Style Heading2 $sectionName { 
                    #Paragraph ""
                    $TableParams = @{
                        Name = $sectionName
                        List = $false
                    }
                    $TableParams['Caption'] = "- $($TableParams.Name)"
                    $item.Group | select DomainName, DisplayName, GpoStatus, LinksCount, LinksDisabledCount, LinksEnabledCount | Table @TableParams
                }
                #endregion

                #region Unlinked GPOs
                $sectionName = "$($item.Name) - Unlinked GPOs"
                Write-RFLLog -Message "Export $($SectionName)"
	            Section -Style Heading2 $sectionName { 
                    $arrGPOList = $item.Group | where-object {$_.LinksCount -eq 0}
                    if ($null -ne $arrGPOList) {
                        #Paragraph ""
                        $TableParams = @{
                            Name = $sectionName
                            List = $false
                        }
                        $TableParams['Caption'] = "- $($TableParams.Name)"
                        $arrGPOList | select DomainName, DisplayName, GpoStatus | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                }
                #endregion

                #region GPO Link Information
                $sectionName = "$($item.Name) - GPO Link Information"
                $gpoLink = get-content '.\ADGPOLink.json' | convertfrom-json | Where-Object {$_.GPOID -in $item.Group.ID}

                Write-RFLLog -Message "Export $($SectionName)"
	            Section -Style Heading2 $sectionName { 
                    if ($null -ne $gpoLink) {
                        #Paragraph ""
                        $TableParams = @{
                            Name = $sectionName
                            List = $false
                        }
                        $TableParams['Caption'] = "- $($TableParams.Name)"
                        $gpoLink | select DisplayName, Enabled, Enforced, Target, Order | Table @TableParams
                    } else {
                        Paragraph "No $($sectionName) found"
                    }
                }
                #endregion
            }
        }
        #endregion

        #region DHCP Information
        $SectionName = "DHCP Information"
        Section -Style Heading1 $SectionName {
            #region DHCP Servers
            $OutObj = get-content '.\DHCPServers.json' | convertfrom-json | select @{Name="IPAddress";Expression = { $_.IpAddress.IPAddressToString}}, DnsName, ConnectionSuccessful
            if ($null -ne $OutObj) {
                Paragraph "The following table lists the domain DHCP Servers in the forest:"
                $TableParams = @{
                    Name = $sectionName
                    List = $false
                }
                $TableParams['Caption'] = "- $($TableParams.Name)"
                $OutObj | select DnsName, IPAddress, ConnectionSuccessful | Table @TableParams
            } else {
                Paragraph "No $($sectionName) found"
            }
            #endregion

            #region DHCP v4 Scope
            $sectionName = 'DHCP v4 Scopes'
            Write-RFLLog -Message "Export $($SectionName)"
	        Section -Style Heading2 $sectionName {
                #Paragraph ""

                $OutObj = get-content '.\DhcpServerv4Scopes.json' | convertfrom-json | select ServerName, @{Name="ScopeId";Expression = { $_.ScopeId.IPAddressToString}}, Name, @{Name="SubnetMask";Expression = { $_.SubnetMask.IPAddressToString}}, @{Name="StartRange";Expression = { $_.StartRange.IPAddressToString}}, @{Name="EndRange";Expression = { $_.EndRange.IPAddressToString}}, @{Name="LeaseDuration";Expression = { ("{0} days, {1} hours, {2} minutes, {3} seconds" -f $_.LeaseDuration.Days, $_.LeaseDuration.Hours, $_.LeaseDuration.Minutes, $_.LeaseDuration.Seconds ) }}, State, Free, InUse, Reserved, PercentageInUse
                if ($null -ne $OutObj) {
                    $TableParams = @{
                        Name = $sectionName
                        List = $false
                    }
                    $TableParams['Caption'] = "- $($TableParams.Name)"
                    $OutObj | Sort-Object -Property 'ServerName', ScopeId | Table @TableParams
                } else {
                    Paragraph "No $($sectionName) found"
                }
            }
            #endregion

            #todo: get scope v6

        }
        #endregion

        #region Password Policies
        $SectionName = "Password Policies"
        Section -Style Heading1 $SectionName {
            Paragraph "The following table lists the password policies in the forest:"

            #region Default Password Policy
            $sectionName = 'Default Password Policy'
            Write-RFLLog -Message "Export $($SectionName)"
	        Section -Style Heading2 $sectionName {
                #Paragraph ""
                $OutObj = get-content '.\DefaultPasswordPolicy.json' | convertfrom-json | select @{Name="Domain";Expression = { $_.SourceDomain}}, ComplexityEnabled, @{Name="LockoutDuration";Expression = { ("{0} days, {1} hours, {2} minutes, {3} seconds" -f $_.LockoutDuration.Days, $_.LockoutDuration.Hours, $_.LockoutDuration.Minutes, $_.LockoutDuration.Seconds ) }}, @{Name="LockoutObservationWindow";Expression = { ("{0} days, {1} hours, {2} minutes, {3} seconds" -f $_.LockoutObservationWindow.Days, $_.LockoutObservationWindow.Hours, $_.LockoutObservationWindow.Minutes, $_.LockoutObservationWindow.Seconds ) }}, LockoutThreshold, @{Name="MaxPasswordAge";Expression = { ("{0} days, {1} hours, {2} minutes, {3} seconds" -f $_.MaxPasswordAge.Days, $_.MaxPasswordAge.Hours, $_.MaxPasswordAge.Minutes, $_.MaxPasswordAge.Seconds ) }}, @{Name="MinPasswordAge";Expression = { ("{0} days, {1} hours, {2} minutes, {3} seconds" -f $_.MinPasswordAge.Days, $_.MinPasswordAge.Hours, $_.MinPasswordAge.Minutes, $_.MinPasswordAge.Seconds ) }}, PasswordHistoryCount, ReversibleEncryptionEnabled
                if ($null -ne $OutObj) {
                    $TableParams = @{
                        Name = "$sectionName"
                        List = $false
                    }
                    $TableParams['Caption'] = "- $($TableParams.Name)"
                    $OutObj | Table @TableParams
                } else {
                    Paragraph "No $($SectionName) found"
                }
            }
            #endregion

            #region Fined Grained Password Policies
            $sectionName = 'Fined Grained Password Policy'
            Write-RFLLog -Message "Export $($SectionName)"
	        Section -Style Heading2 $sectionName {
                #Paragraph ""
                $OutObj = get-content '.\FinedGrainedPasswordPolicies.json' | convertfrom-json | select @{Name="Domain";Expression = { $_.SourceDomain}}, Created, Name, Description, @{Name="AppliesTo";Expression = { $_.AppliesTo -join ', '}},  ComplexityEnabled, @{Name="LockoutDuration";Expression = { ("{0} days, {1} hours, {2} minutes, {3} seconds" -f $_.LockoutDuration.Days, $_.LockoutDuration.Hours, $_.LockoutDuration.Minutes, $_.LockoutDuration.Seconds ) }}, @{Name="LockoutObservationWindow";Expression = { ("{0} days, {1} hours, {2} minutes, {3} seconds" -f $_.LockoutObservationWindow.Days, $_.LockoutObservationWindow.Hours, $_.LockoutObservationWindow.Minutes, $_.LockoutObservationWindow.Seconds ) }}, LockoutThreshold, @{Name="MaxPasswordAge";Expression = { ("{0} days, {1} hours, {2} minutes, {3} seconds" -f $_.MaxPasswordAge.Days, $_.MaxPasswordAge.Hours, $_.MaxPasswordAge.Minutes, $_.MaxPasswordAge.Seconds ) }}, @{Name="MinPasswordAge";Expression = { ("{0} days, {1} hours, {2} minutes, {3} seconds" -f $_.MinPasswordAge.Days, $_.MinPasswordAge.Hours, $_.MinPasswordAge.Minutes, $_.MinPasswordAge.Seconds ) }}, PasswordHistoryCount, Precedence, ReversibleEncryptionEnabled
                if ($null -ne $OutObj) {
                    $TableParams = @{
                        Name = "$sectionName"
                        List = $false
                    }
                    $TableParams['Caption'] = "- $($TableParams.Name)"
                    $OutObj | Table @TableParams
                } else {
                    Paragraph "No $($SectionName) found"
                }
            }
            #endregion

        }
        #endregion

        #region Computer Accounts
        $sectionName = 'Computer Accounts'
        Write-RFLLog -Message "Export $($SectionName)"
	    Section -Style Heading1 $sectionName { 
            $OutObj = get-content '.\ADComputers.json' | convertfrom-json

            #region Computer Count
            $sectionName = 'Computer Count'
            Write-RFLLog -Message "Export $($SectionName)"
	        Section -Style Heading2 $sectionName { 
                $TableOut = @()
                foreach($item in ($OutObj | Group-Object Domain)) {
                    $TableOut += $item | select Name, Count, @{Name="Servers";Expression = { ($item.Group | Where-Object {$_.OperatingSystem -like '*server*'}).Count }}, @{Name="Workstations";Expression = { ($item.Group | Where-Object {$_.OperatingSystem -notlike '*server*'}).Count }}
                }
                $TableParams = @{
                    Name = $sectionName
                    List = $false
                }
                $TableParams['Caption'] = "- $($TableParams.Name)"
                $TableOut | Table @TableParams
            }
            #endregion 

            #region Computer Objects
            $sectionName = 'Computer Status'
            Write-RFLLog -Message "Export $($SectionName)"
	        Section -Style Heading2 $sectionName { 
                #$OutObj = get-content '.\ADComputers.json' | convertfrom-json
                $TableOut = @()
                foreach($item in ($OutObj | Group-Object Domain)) {
                    $TableOut += $item | select Name, Count, @{Name="Enabled";Expression = { ($item.Group | Where-Object {$_.Enabled -eq $true}).Count }}, @{Name="Disabled";Expression = { ($item.Group | Where-Object {$_.Enabled -eq $false}).Count }}
                }
                $TableParams = @{
                    Name = "$sectionName"
                    List = $false
                }
                $TableParams['Caption'] = "- $($TableParams.Name)"
                $TableOut | Table @TableParams
            }
            #endregion 

            #region Status of Computer Objects
            $sectionName = 'Status of Computer Objects'
            Write-RFLLog -Message "Export $($SectionName)"
	        Section -Style Heading2 $sectionName { 
                #$OutObj = get-content '.\ADComputers.json' | convertfrom-json
                $TableOut = @()
                foreach($item in ($OutObj | Group-Object Domain)) {
                    $TableOut += $item | select Name, Count, @{Name="'Password Age (> $($PasswordDays) days)'";Expression = { ($item.Group | Where-Object {$_.PasswordLastSet -le $passwordtime}).Count }}, @{Name="Dormant (> $($DormantDays) days)";Expression = { ($item.Group | Where-Object {[datetime]::FromFileTime($_.LastLogonTimestamp) -lt $dormanttime}).Count }}, @{Name="SidHistory";Expression = { ($item.Group | Where-Object {$_.SIDHistory.Count -ne 0}).Count }}
                }
                $TableParams = @{
                    Name = $sectionName
                    List = $false
                }
                $TableParams['Caption'] = "- $($TableParams.Name)"
                $TableOut | Table @TableParams
            }
            #endregion

            #region Operating System
            $sectionName = 'Operating System'
            Write-RFLLog -Message "Export $($SectionName)"
	        Section -Style Heading2 $sectionName { 
                #$OutObj = get-content '.\ADComputers.json' | convertfrom-json
                $TableOut = @()
                foreach($item in ($OutObj | Group-Object Domain, OperatingSystem)) {
                    $TableOut += $item | select @{Name="Domain";Expression = { $item.Name.Split(',')[0] }}, @{Name="Operating System";Expression = { $item.Name.Split(',')[1] }}, Count
                }
                $TableParams = @{
                    Name = $sectionName
                    List = $false
                }
                $TableParams['Caption'] = "- $($TableParams.Name)"
                $TableOut | Table @TableParams
            }
            #endregion
        }
        #endregion

        #region User Accounts
        $sectionName = 'User Accounts'
        Write-RFLLog -Message "Export $($SectionName)"
	    Section -Style Heading1 $sectionName { 
            $Users = get-content '.\ADUsers.json' | convertfrom-json

            #region User Objects
            $sectionName = 'User Objects'
            Write-RFLLog -Message "Export $($SectionName)"
	        Section -Style Heading2 $sectionName { 
                $GroupObj = get-content '.\GroupList.json' | convertfrom-json
                $PrivUser = get-content '.\ADUserPriviliged.json' | convertfrom-json

                $TableOut = @()
                foreach($item in $Global:Domains) {
                    $TableOut += $item | select @{Name="Domain";Expression = { $item }}, @{Name="Users";Expression = { ($Users | Group-Object Domain | Where-Object {$_.Name -eq $item}).Count }}, @{Name="Privileged Users";Expression = { ($PrivUser | Group-Object SourceDomain | Where-Object {$_.Name -eq $item}).Count }}, @{Name="Groups";Expression = { ($GroupObj | Group-Object Domain | Where-Object {$_.Name -eq $item}).Count }}
                }

                $TableParams = @{
                    Name = "$sectionName"
                    List = $false
                }
                $TableParams['Caption'] = "- $($TableParams.Name)"
                $TableOut | Table @TableParams
            }
            #endregion

            #region User Account Status
            $sectionName = 'User Account Status'
            Write-RFLLog -Message "Export $($SectionName)"
	        Section -Style Heading2 $sectionName { 
                #$Users = get-content '.\ADUsers.json' | convertfrom-json
                $TableOut = @()
                foreach($item in ($Users | Group-Object Domain)) {
                    $TableOut += $item | select Name, Count, @{Name="Enabled";Expression = { ($item.Group | Where-Object {$_.Enabled -eq $true}).Count }}, @{Name="Disabled";Expression = { ($item.Group | Where-Object {$_.Enabled -eq $false}).Count }}
                }
                $TableParams = @{
                    Name = "$sectionName"
                    List = $false
                }
                $TableParams['Caption'] = "- $($TableParams.Name)"
                $TableOut | Table @TableParams
            }
            #endregion 

            #region Status of User Account
            $sectionName = 'Status of User Account'
            Write-RFLLog -Message "Export $($SectionName)"
	        Section -Style Heading2 $sectionName { 
                #$Users = get-content '.\ADUsers.json' | convertfrom-json
                $TableOut = @()
                foreach($item in ($Users | Group-Object Domain)) {
                    $TableOut += $item | select Name, Count, @{Name="Cannot Change Password";Expression = { ($item.Group | Where-Object {$_.CannotChangePassword -eq $true}).Count }}, @{Name="Password Never Expires";Expression = { ($item.Group | Where-Object {$_.PasswordNeverExpires -eq $true}).Count }}, @{Name="Must Change Password at Logon";Expression = { ($item.Group | Where-Object {$_.pwdLastSet -eq 0}).Count }}, @{Name="Smartcard Logon Required";Expression = { ($item.Group | Where-Object {$_.SmartcardLogonRequired -eq $true}).Count }}, @{Name="SidHistory";Expression = { ($item.Group | Where-Object {$_.SIDHistory.Count -ne 0}).Count }}, @{Name="Never Logged in";Expression = { ($item.Group | Where-Object {($_.lastlogontimestamp -eq '01/01/1601 00:00:00')}).Count }}, @{Name="Dormant (> 90 days)";Expression = { ($item.Group | Where-Object {$_.LastLogonTimestamp -lt $dormanttime}).Count }}, @{Name="Password Not Required";Expression = { ($item.Group | Where-Object {$_.PasswordNotRequired -eq $true}).Count }}, @{Name="Account Lockout";Expression = { ($item.Group | Where-Object {$_.LockedOut -eq $true}).Count }}
                }
                $TableParams = @{
                    Name = "$sectionName"
                    List = $false
                }
                $TableParams['Caption'] = "- $($TableParams.Name)"
                $TableOut | Table @TableParams
            }
            #endregion 
        }
        #endregion 

        #region Domain Count Objects
        $sectionName = 'Domain Count Objects'
        Write-RFLLog -Message "Export $($SectionName)"
	    Section -Style Heading1 $sectionName { 
            #region AD Object Type Count
            $sectionName = 'AD Object Type Count'
            Write-RFLLog -Message "Export $($SectionName)"
	        Section -Style Heading2 $sectionName { 
                $outObj = get-content '.\ADObject.json' | convertfrom-json
                $TableOut = @()
                foreach($item in ($outObj | Group-Object SourceDomain,ObjectClass)) {
                    $TableOut += $item | select @{Name="Domain";Expression = { $item.Name.Split(',')[0] }}, @{Name="Object Type";Expression = { $item.Name.Split(',')[1] }}, Count
                }
                $TableParams = @{
                    Name = $sectionName
                    List = $false
                }
                $TableParams['Caption'] = "- $($TableParams.Name)"
                $TableOut | sort-object Domain, 'Object Type' | Table @TableParams
            }
            #endregion 

            #region SPN Count
            $sectionName = 'SPN Count'
            Write-RFLLog -Message "Export $($SectionName)"
	        Section -Style Heading2 $sectionName { 
                $Excluded = @('kadmin/changepw')
                $SPNCache = [ordered] @{}
                foreach ($Object in (get-content '.\SPN.json' | convertfrom-json -AsHashtable)) {
                    foreach ($SPN in $Object.ServicePrincipalName) {
                        if (-not $SPNCache[$SPN]) {
                            $SPNCache[$SPN] = [PSCustomObject] @{
                                Name = $SPN
                                SourceDomain = $object.SourceDomain
                                Duplicate = $false
                                Count = 0
                                Excluded = $false
                                List = [System.Collections.Generic.List[Object]]::new()
                            }
                        }
                        if ($SPN -in $Excluded) {
                            $SPNCache[$SPN].Excluded = $true
                        }
                        $SPNCache[$SPN].List.Add($Object)
                        $SPNCache[$SPN].Count++
                    }
                }

                foreach ($SPN in $SPNCache.Values) {
                    if ($SPN.Count -gt 1 -and $SPN.Excluded -ne $true) {
                        $SPN.Duplicate = $true
                    }
                }

                $outObj = $SPNCache.Values | select SourceDomain,Name,Duplicate,Count,Excluded,List

                $TableOut = @()
                foreach($item in ($outObj | Group-Object SourceDomain)) {
                    $TableOut += $item | select @{Name="Domain";Expression = { $item.Name }}, @{Name="Duplicated";Expression = { ($item.Group | Where-Object {$_.Duplicate -eq $true}).Count }}, @{Name="Not Duplicated";Expression = {  ($item.Group | Where-Object {$_.Duplicate -eq $false}).Count }}
                }
                $TableParams = @{
                    Name = $sectionName
                    List = $false
                }
                $TableParams['Caption'] = "- $($TableParams.Name)"
                $TableOut | sort-object Domain, 'Object Type' | Table @TableParams
            }
            #endregion 

            #region GMSA - Group Managed Service Accounts
            $sectionName = 'GMSA - Group Managed Service Accounts'
            Write-RFLLog -Message "Export $($SectionName)"
	        Section -Style Heading2 $sectionName { 
                if ($null -ne $OutObj) {
                    $outObj = get-content '.\GroupManagedServiceAccounts.json' | convertfrom-json
                    $TableParams = @{
                        Name = $sectionName
                        List = $false
                    }
                    $TableParams['Caption'] = "- $($TableParams.Name)"
                    $outObj | select @{Name="Domain";Expression = { $item.Name.Split(',')[0] }}, Name, DistinguishedName, Enabled, HostComputers, UserPrincipalName | sort-object Domain | Table @TableParams
                } else {
                    Paragraph "No $($sectionName) found"
                }
            }
            #endregion 
        }
        #endregion 
    }
    #endregion

    #region Export File
    foreach($OutPutFormatItem in $OutputFormat) {
        Write-RFLLog -Message "Exporting report format $($OutPutFormatItem) to $($OutputFolderPath)"
	    $Document = $Global:WordReport | Export-Document -Path $OutputFolderPath -Format:$OutPutFormatItem -Options @{ TextWidth = 240 } -PassThru
    }
    #endregion

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