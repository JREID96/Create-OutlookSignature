<# Outlook 2016 Signature Generator
This program will generate signatures for Outlook automatically
Jarred Reid - 2019

Changelog:

v0.1 - Alpha release

#>


# Determine script location for PowerShell
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
cd $ScriptDir

$replySigTxt = '.\Reply Signature.txt'
$regSigTxt = '.\Regular Signature.txt'

$regSigHtm = '.\Regular Signature.htm'
$replySigHtm = '.\Reply Signature.htm'

#$replySigRtf = '.\Reply Signature.rtf'
#$regSigRtf = '.\Regular Signature.rtf'

$alias = Read-Host -prompt "Set the alias"

$name = Get-ADUser $alias -Properties name | select -expand name
$title = Get-ADUser $alias -Properties title | select -expand title
$email = Get-ADUser $alias -Properties UserPrincipalName | select -expand UserPrincipalName
$department = Get-ADUser $alias -Properties Department | select -expand Department
$office = Get-ADUser $alias -Properties Office | select -expand Office

$phone = Get-ADUser $alias -Properties telephoneNumber | select -expand telephoneNumber
$ext = $phone.split(" ")
$ext = $ext[1]

### Write-Host "$name | $title | $email | $department | $office | $ext"

### Append data to TXT files
(Get-Content $regsigtxt) | foreach-Object {$_ -replace "name",$name -replace "title",$title -replace "department",$department -replace "ext",$ext -replace "email",$email} | Set-Content $regsigtxt

(Get-Content $replysigtxt) | foreach-Object {$_ -replace "name",$name -replace "title",$title -replace "department",$department -replace "ext",$ext -replace "email",$email} | Set-Content $replysigtxt

### Append data to HTM files
(Get-Content $regsightm) | foreach-Object {$_ -replace "name",$name -replace "title",$title -replace "department",$department -replace "ext",$ext -replace "email",$email} | Set-Content $regsightm

(Get-Content $replysightm) | foreach-Object {$_ -replace "name",$name -replace "title",$title -replace "department",$department -replace "ext",$ext -replace "email",$email} | Set-Content $replysightm

### Append data to RTF files
#(Get-Content $replySigRtf) | foreach-Object {$_ -replace "Repname1",$name -replace "Reptitle1",$title -replace "Repdept1",$department -replace "Repext1",$ext -replace "Repemail1",$email} | Set-Content $replySigRtf

#(Get-Content $regSigRtf) | foreach-Object {$_ -replace "Repname1",$name -replace "Reptitle1",$title -replace "Repdept1",$department -replace "Repext1",$ext -replace "Repemail1",$email} | Set-Content $regSigRtf

### Signature injection

$pc = Read-Host -prompt "Enter the PC name"

$path = "\\$pc\C$\users\$alias\AppData\Roaming\Microsoft\Signatures"

Copy-Item $ScriptDir -Destination $path -Recurse

#>
