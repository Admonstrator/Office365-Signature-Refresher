<# 
.SYNOPSIS Set a Exchange Online signature for org-wide users.
.DESCRIPTION
.NOTES Use at your own risk!
.COMPONENT Requires Module ActiveDirectory, ExchangeOnlineManagement
.Parameter LoginCredentialFile Path to Clixml-credential file
.Parameter PlainSignatureFile Path to plain signature file
.Parameter HTMLSignatureFile Path to HTML signature file
#>

# Declare your variables here
# Yeah, I know. That is not how it should be - maybe a later version will be suitable for cli parameters instead

# Companyname - must be the same like in company attribute in your local AD:
$CompanyName = "Example Corp"
# This file must be created by Export-Clixml - it contains the credentials for Exchange Online:
$FileLoginCredential = "$PSScriptRoot\login.cred"
# This file contains the plain text signature:
$FilePlainSignature = "$PSScriptRoot\signature.txt"
# This file contains the HTML text signature:
$FileHTMLSignature = "$PSScriptRoot\signature.html"
# This file will contain the log output:
$FileLog = "$PSScriptRoot\script.log"

$ErrorActionPreference = "Stop"

Try {
    Start-Transcript -Path $FileLog -IncludeInvocationHeader
}
Catch {
    Start-Transcript -Path $FileLog -IncludeInvocationHeader
}

if (Test-Path -Path $FileLoginCredential) {
    $UserCredential = Import-Clixml -path $FileLoginCredential
}
else {
    Write-Error "Credential file is missing!"
    Exit
}

Write-Host Starting signature refresh cycle for 365 OWA ...
Write-Host Terminate old sessions ...
Get-PSSession | where computername -like "outlook.office365.com" | Remove-PSSession
Write-Host Import ActiveDirectory module ...
import-module activedirectory
Write-Host Import Exchange module ...
import-module ExchangeOnlineManagement
Write-Host Import credentials ...
Write-Host Connecting to Exchange Online Management Services ...
Connect-ExchangeOnline -Credential $UserCredential -ShowProgress $false

# Get all active users from AD where Company is set and mail attribute exists:
Write-Host Refreshing signatures ...
$allusers = Get-ADUser -filter 'mail -like "*" -and Company -eq $CompanyName -and enabled -eq "true" -and ObjectClass -eq "User"' -Properties displayname, samaccountname, office, officephone, fax, mail, city, streetaddress, postalcode, title, mobilephone
foreach ($user in $allusers) {
    if (Test-Path -Path $FileHTMLSignature) {
        # Replacing variables in HTML signature
        $SignatureHTML = Get-Content($FileHTMLSignature) -encoding utf8
        $SignatureHTML = $SignatureHTML.Replace("[Display Name]", $user.displayname)
        $SignatureHTML = $SignatureHTML.Replace("[Job Title]", $user.title)
        $SignatureHTML = $SignatureHTML.Replace("[Street]", $user.streetaddress)
        $SignatureHTML = $SignatureHTML.Replace("[Zip]", $user.postalcode)
        $SignatureHTML = $SignatureHTML.Replace("[City]", $user.city)
        if ($user.officephone -like "") {
            $SignatureHTML = $SignatureHTML.Replace("Phone: [Telephone No]<br>", "")
        }
        else {
            $SignatureHTML = $SignatureHTML.Replace("[Telephone No]", $user.officephone)
        }
        if ($user.mobilephone -like "") {
            $SignatureHTML = $SignatureHTML.Replace("Mobile: [Mobile No]<br>", "")
        }
        else {
            $SignatureHTML = $SignatureHTML.Replace("[Mobile No]", $user.mobilephone)
        }
        if ($user.Fax -like "") {
            $SignatureHTML = $SignatureHTML.Replace("Fax: [Fax No]<br>", "")
        }
        else {
            $SignatureHTML = $SignatureHTML.Replace("[Fax No]", $user.fax)
        }
        $SignatureHTML = $SignatureHTML.Replace("[EmailAddress]", $user.mail)
    }

    if (Test-Path -Path $FilePlainSignature) {
        # Replacing variables in plain signature
        $SignaturePlain = Get-Content($FilePlainSignature) -encoding utf8
        $SignaturePlain = $SignaturePlain.Replace("[Display Name]", $user.displayname)
        $SignaturePlain = $SignaturePlain.Replace("[Job Title]", $user.title)
        $SignaturePlain = $SignaturePlain.Replace("[Street]", $user.streetaddress)
        $SignaturePlain = $SignaturePlain.Replace("[Zip]", $user.postalcode)
        $SignaturePlain = $SignaturePlain.Replace("[City]", $user.city)
        if ($user.officephone -like "") {
            $SignaturePlain = $SignaturePlain.Replace("Phone: [Telephone No]", "")
        }
        else {
            $SignaturePlain = $SignaturePlain.Replace("[Telephone No]", $user.officephone)
        }
        if ($user.mobilephone -like "") {
            $SignaturePlain = $SignaturePlain.Replace("Mobile: [Mobile No]", "")
        }
        else {
            $SignaturePlain = $SignaturePlain.Replace("[Mobile No]", $user.mobilephone)
        }
        if ($user.Fax -like "") {
            $SignaturePlain = $SignaturePlain.Replace("Fax: [Fax No]", "")
        }
        else {
            $SignaturePlain = $SignaturePlain.Replace("[Fax No]", $user.fax)
        }
        $SignaturePlain = $SignaturePlain.Replace("[EmailAddress]", $user.mail)
    }
    # Push signatures
    # You can push other settings as well, see https://docs.microsoft.com/de-de/powershell/module/exchange/set-mailboxmessageconfiguration?view=exchange-ps
    # For pushing a default font add something like -DefaultFontName:Tahoma -DefaultFontSize:11
    Set-MailboxMessageConfiguration -AlwaysShowBcc:$true -AutoAddSignatureOnMobile:$true -IsFavoritesFolderTreeCollapsed:$true -IsMailRootFolderTreeCollapsed:$true -LinkPreviewEnabled:$false -MailFolderPaneExpanded:$true -UseDefaultSignatureOnMobile:$true -SignatureHtml:$SignatureHTML -DefaultFormat:html -AutoAddSignatureOnReply:$true -AutoAddSignature:$true -SignatureText:$SignaturePlain -SignatureTextOnMobile:$SignaturePlain -Identity $user.mail
    Write-Host Signature set for $user.samaccountname -ForegroundColor Green
}


Write-Host Terminate active session ...
Disconnect-ExchangeOnline -Confirm:$false
Write-Host Refresh done!
Stop-Transcript