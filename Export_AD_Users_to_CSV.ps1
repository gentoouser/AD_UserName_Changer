###########################################################
# AUTHOR  : Victor Ashiedu
# WEBSITE : iTechguides.com
# BLOG    : iTechguides.com/blog-2/
# CREATED : 08-08-2014 
# UPDATED : 27-11-2017 
# COMMENT : This script exports Active Directory users
#           to a a csv file. v2.1 adds the condition to 
#           ignore all users with the info (Notes) field
#           found on the Telephones tab containing the 
#           word 'Migrated'. 
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


#Define location of my script variable
#the -parent switch returns one directory lower from directory defined. 
#below will return up to ImportADUsers folder 
#and since my files are located here it will find it.
#It failes withpout appending "*.*" at the end

$path = Split-Path -parent "\\wwt-fp03.wwtps.com\teams\it\Active Directory\*.*"

#Create a variable for the date stamp in the log file

$LogDate = get-date -f yyyyMMddhhmm

#Define CSV and log file location variables
#they have to be on the same location as the script

$csvfile = $path + "\ALLADUsers_$logDate.csv"

#import the ActiveDirectory Module

#Import-Module ActiveDirectory


#Sets the OU to do the base search for all user accounts, change as required.
#Simon discovered that some users were missing
#I decided to run the report from the root of the domain

$SearchBase = "DC=WWTPS,DC=com"

#Get Admin accountb credential

#$GetAdminact = Get-Credential

#Define variable for a server with AD web services installed

$ADServer = 'wwt-ad01.wwtps.com'

#Find users that are not disabled
#To test, I moved the following users to the OU=ADMigration:
#Philip Steventon (kingston.gov.uk/RBK Users/ICT Staff/Philip Steventon) - Disabled account
#Joseph Martins (kingston.gov.uk/RBK Users/ICT Staff/Joseph Martins) - Disabled account
#may have to get accountb status with another AD object

#Define "Account Status" 
#Added the Where-Object clause on 23/07/2014
#Requested by the project team. This 'flag field' needs
#updated in the import script when users fields are updated
#The word 'Migrated' is added in the Notes field, on the Telephone tab.
#The LDAB object name for Notes is 'info'. 

#$AllADUsers = Get-ADUser -server $ADServer `
#-Credential $GetAdminact -searchbase $SearchBase `
#-Filter * -Properties * | Where-Object {$_.info -NE 'Migrated'} #ensures that updated users are never exported.

$AllADUsers = Get-ADUser -server $ADServer `
-searchbase $SearchBase `
-Filter * -Properties * | Where-Object {$_.info -NE 'Migrated'} #ensures that updated users are never exported.

$AllADUsers |
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
@{Label = "Office";Expression = {$_.OfficeName}},
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
#Export CSV report

Export-Csv -Path $csvfile -NoTypeInformation
