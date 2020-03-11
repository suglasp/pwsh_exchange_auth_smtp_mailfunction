
#
# Pieter De Ridder
# Create a Powershell credential file used for Authenticated SMTP access
#
# Notes on the output file (.clixml):
# This file will be unique per host, because it uses for hashing the Windows Data Protection API.
# On Linux and Mac OS the password will not be hashed.
# For each machine that needs to use the Send-SmtpExchange function, you will want to run this script.
# Ref : https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/export-clixml?view=powershell-7
#


# global vars
$global:Workfolder = $PSScriptRoot
$global:Hostname = ($env:COMPUTERNAME).ToLowerInvariant()

# generate a clixml hashed credentials file
$credfile = "$($global:Workfolder)\$($global:Hostname)_smtp_cred.clixml"
$CredUserName = "smtp_account@$($env:USERDNSDOMAIN)".ToLowerInvariant()
$credForExport = Get-Credential -Credential $CredUserName
$credForExport | Export-CliXml $credfile
$credForExport = $null


# example code to load the clixml credential into memory
#If (Test-Path($CredFile)) {
#    $CredFromFile = Import-CliXml $CredFile
#}
#$CredFromFile = $null

