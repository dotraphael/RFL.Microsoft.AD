# Active Directory Documentation
Automatic create an Active Directory Documentation to simplify the life of admins and consultants.

# Usage
CollectADData.ps1 - On a elevated PowerShell session, it collects AD Data and saves to json files. It can be run from any Windows Device (Workstation, Server or Domain Controller). It only needs to have the Windows Features: PowerShell module for Active Directory, DNS Server Management, DHCP Server Management and GPMC.

ExportADData.ps1 - Exports the Active Directory data collected using the ColectADData.ps1 to a word or html format. It uses the PScribo PowerShell module (https://github.com/iainbrighton/PScribo). It can be run from any Windows Device (Workstation, Server or Domain Controller). As it uses external PowerShell module, it is recommended not to run from a Domain Controller.

# Examples
.\ExportADData.ps1 -SaveTo 'c:\temp\CollectADData'
**Collects the Information and save the json files to 'c:\temp\CollectADData'

.\ExportADData.ps1 -SourceData 'c:\temp\CollectADData' -OutputFormat @('Word')
**Exports the data collected by CollectADData.ps1 script onto 'c:\temp\CollectADData' and export to word format

.\ExportADData.ps1 -SourceData 'c:\temp\CollectADData' -OutputFormat @('Word') -DormantDays 90 -PasswordDays 30 
**Exports the data collected by CollectADData.ps1 script onto 'c:\temp\CollectADData' and export to word format. It classify an object as dormant if the password was not changed for over 90 days and a computer object that did not change its password for 30 days. 

.\ExportADData.ps1 -SourceData 'c:\temp\CollectADData' -OutputFormat @('Word', 'HTML') -DormantDays 90 -PasswordDays 30 -CompanyName 'RFL Systems' -CompanyWeb 'www.rflsystems.co.uk' -CompanyEmail 'team@rflsystems.co.uk'
**Exports the data collected by CollectADData.ps1 script onto 'c:\temp\CollectADData' and export to word format. It classify an object as dormant if the password was not changed for over 90 days and a computer object that did not change its password for 30 days. Add the company details to the header

# Documentation
Access our Wiki at https://github.com/dotraphael/RFL.Microsoft.AD/wiki

# Issues and Support
Access our Issues at https://github.com/dotraphael/RFL.Microsoft.AD/issues
