###########################################################
# AUTHOR  : Victor Ashiedu
# WEBSITE : iTechguides.com
# BLOG    : iTechguides.com/blog-2/
# CREATED : 08-08-2014 
# UPDATED : 05-12-2017 
# COMMENT : This script exports Active Directory users
#           to a a csv file. 
###########################################################


#Force Starting of Powershell script as Administrator 
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))

{   
$arguments = "& '" + $myinvocation.mycommand.definition + "'"
Start-Process powershell -Verb runAs -ArgumentList $arguments
Break
}

#Import AD modules
If ((Get-Module | Where-Object {$_.Name -Match "ActiveDirectory"}).Count -eq 0 ) {
	Write-Host ("Loading Active Directory Plugins") -foregroundcolor "Green"
	Import-Module "ActiveDirectory"  -ErrorAction SilentlyContinue
} Else {
	Write-Host ("Active Directory Plug-ins Already Loaded") -foregroundcolor "Green"
}



#Define CSV and log file location variables
#they have to be on the same location as the script

$csvfile = ((Split-Path -Parent -Path $MyInvocation.MyCommand.Definition) + "\" + $MyInvocation.MyCommand.Name + "_" + (Get-Date -format yyyyMMdd-hhmm) + ".csv")


#Sets the OU to do the base search for all user accounts, change as required.
$SearchBase = (Get-ADDomain).DistinguishedName

#Get Admin accountb credential

#$GetAdminact = Get-Credential

#Define variable for a server with AD web services installed
$ADServer = (Get-ADDomain).PDCEmulator


# Where-Object {$_.info -NE 'Migrated'} #ensures that updated users are never exported.
# Where-Object {$_.Enabled -eq 'TRUE'} #Only get enabled users


Get-ADUser -server $ADServer -searchbase $SearchBase -Filter * -Properties GivenName,Surname,sAMAccountName,DisplayName,StreetAddress,City,st,PostalCode,Country,Title,Company,Description,Department,physicalDeliveryOfficeName,telephoneNumber,pager,Mail,homeMDB,Manager,homeDirectory,Enabled,lastLogon,whencreated,msTSExpireDate,pwdLastSet  | 
Select-Object @{Label = "First Name";Expression = {$_.GivenName}},
@{Label = "Last Name";Expression = {$_.Surname}},
@{Label = "Logon Name";Expression = {$_.sAMAccountName}},
@{Label = "Display Name";Expression = {$_.DisplayName}},
@{Label = "Full address";Expression = {$_.StreetAddress}},
@{Label = "City";Expression = {$_.City}},
@{Label = "State";Expression = {$_.st}},
@{Label = "Post Code";Expression = {$_.PostalCode}},
@{Label = "Country/Region";Expression = {if (($_.Country -eq 'GB')  ) {'United Kingdom'} Else {''}}},
@{Label = "Job Title";Expression = {$_.Title}},
@{Label = "Company";Expression = {$_.Company}},
@{Label = "Directorate";Expression = {$_.Description}},
@{Label = "Department";Expression = {$_.Department}},
@{Label = "Office";Expression = {$_.physicalDeliveryOfficeName}},
@{Label = "Phone";Expression = {$_.telephoneNumber}},
@{Label = "Jack";Expression = {$_.pager}},
@{Label = "Email";Expression = {$_.Mail}},
@{Label = "Mail Store";Expression = {($_.homeMDB).SubString(3,($_.homeMDB.Indexof(",")-3))}},
@{Label = "Manager";Expression = {%{(Get-AdUser $_.Manager -server $ADServer -Properties DisplayName).DisplayName}}},
@{Label = "Home Directory";Expression = {$_.homeDirectory}},
@{Label = "Account Status";Expression = {if (($_.Enabled -eq 'TRUE')  ) {'Enabled'} Else {'Disabled'}}}, # the 'if statement# replaces $_.Enabled
@{Label = "Last LogOn Date";Expression = {[DateTime]::FromFileTime($_.lastLogon)}},
@{Label = "Days Since Last LogOn";Expression = {$(([DateTime]::FromFileTime($_.lastLogon)) - (Get-Date)).Days}},
@{Label = "Creation Date";Expression = {$_.whencreated}}, 
@{Label = "Days Since Creation";Expression = {$(([DateTime]($_.whencreated)) - (Get-Date)).Days}}, 
@{Label = "RDS CAL Expiration Date";Expression = {$_.msTSExpireDate}}, 
@{Label = "Days to RDS CAL Expiration";Expression = {$(([DateTime]($_.msTSExpireDate)) - (Get-Date)).Days}}, 
@{Label = "Last Password Change";Expression = {[DateTime]::FromFileTime($_.pwdLastSet)}}, 
@{Label = "Days from last password change";Expression = {$(([DateTime]::FromFileTime($_.pwdLastSet)) - (Get-Date)).Days}} | 
Export-Csv -Path $csvfile -NoTypeInformation
