
#
# Pieter De Ridder
# Exchange On-Premise Mailing function
# Created : 22/10/2019
# Updated : 11/03/2020
#
# Function that can do authenticated mailing to a on-premise Exchange in the AD/LDAP Domain
#
# Tested with Exchange 2016 in a single forest/single domain testing environment
#


#region global vars
[string]$global:Workfolder = "$($PSScriptRoot)"                                                  # Current work folder

# AD domain vars
[string]$global:DOMAINROOT       = $env:USERDNSDOMAIN.ToLowerInvariant()                         # FQDN

# smtp server vars (mailing)
$global:DefaultSmtpServer        = "smtp.$($env:USERDNSDOMAIN)".ToLowerInvariant()               # smtp.domain.com
$global:DefaultSmtpServerSSLPort = 587                                                           # smtp through authenticated SSL TCP port 587
$global:DefaultUserName          = "smtp_account@$($env:USERDNSDOMAIN)".ToLowerInvariant()       # smtp_account@domain.com. Important NOT to use domain\username for Exchange (in MS language, use UPN (User Principale Name) and NOT Down-Level Logon name)!
$global:SmtpCredFile             = "$($global:Workfolder)\$($env:COMPUTERNAME)_smtp_cred.clixml" # PWSH clixml file holding the SMTP credentials
$global:ToAddressesFile          = "$($global:Workfolder)\smtp_addresses.txt"                    # list with To mail adresses
$global:ToAddresses              = @()                                                           # private list to hold the To Mail adresses
$global:SmtpSubjectKeynote       = "[TEST MAIL] some demo mail"                                  # Mail subject

# other vars relevant to this script
[string]$global:SomeFile         = "$($global:Workfolder)\test.txt"                              # a test file to include a mail attachment
#endregion



#
# Function : Send-SmtpExchange
# method for sending authenticated SMTP mail 
# 
Function Send-SmtpExchange {  
    Param(
        [string]$Subject,
        [string]$Body,
        [string[]]$Attachments
    )

    # create a var to hold credential
    $privateSMTPCred = $null

    Write-Host "Sending SMTP mail..."

    # read credentials file
    If (Test-Path $global:SmtpCredFile) {
        $privateSMTPCred = Import-CliXml $global:SmtpCredFile
    } else {
        Write-Warning "[!] Failed to load SMTP credentials file!"
    }

    # read mail addresses To send to file
    If (Test-Path $global:ToAddressesFile) {
        $global:ToAddresses = Get-Content -Path $global:ToAddressesFile
    } else {
        Write-Warning "[!] Failed to read SMTP Address file!"
    }

    # send mail through Exchange SSL
    If (($global:ToAddresses) -And ($privateSMTPCred)) {
        $props = @{
           To=$global:ToAddresses
           Subject=$Subject
           Body=$Body
           From=$($global:DefaultUserName)
           SmtpServer=$($global:DefaultSmtpServer)
           Port=$($global:DefaultSmtpServerSSLPort)
        }

        # add is not empty
        If ($Attachments) { $props.Attachments=$Attachments }

	# try sending the mail
        try {
            Send-MailMessage @props -UseSsl -Credential $privateSMTPCred
			Write-Host "Mail send."
        } catch {
            $_
        }
    } else {
        If (-not ($global:ToAddresses)) {
            Write-Warning "SMTP Address list empty!"
        }

        If (-not ($privateSMTPCred)) {
            Write-Warning "SMTP credential problem!"
        }

        Write-Host "[!] SMTP Mailing failed!"
    }
}




# --- TEST THE FUNCTION Send-SmtpExchange ---
If (Test-Path $global:SomeFile) {
    $mailBody = "SOME TEST MAIL FROM HOST $($env:COMPUTERNAME)"

    Send-SmtpExchange -Subject "$($global:SmtpSubjectKeynote)" -Body $mailBody -Attachments @($global:SomeFile)
} Else {
    Write-Warning "No file $($global:SomeFile)!"
}
