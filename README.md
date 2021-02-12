# Office365-Signature-Refresher
## Synopsis
Setting up org-wide signatures for O365 easily by using Powershell.

## Motivation
Small or medium-sized businesses often struggle with setting up org-wide signatures in Office Outlook for Web. They are many paid softwares available but often there is not enough budget for such an easy task. To support sysadmins I release my script for free.

## Information
**Please be aware that this script is provided "as is" and therefor there might be problems, issues and more problems. Make sure to understand the script and test before rollout!**

**Please make sure to fully understand how this script works before running it against your local AD. Mistakes should not kill anything - but better safe than sorry!**

## Requirements

* Office365 account with admin permissions
* Powershell modules: `ActiveDirectory`, `ExchangeOnlineManagement`
* Users in AD with correct attribute data in the following fields: 
  * company
  * displayname
  * title
  * streetaddress
  * postalcode
  * city
  * officephone
  * mobilephone
  * fax
  * mail

## How it works
The script will read all AD users with the following criterias:
* Company is set to ``$Company`` (see inline)
* User account is active
* Object type is 'User'

Afterwards it will replace all placeholders in the template files by attributes of the user and push it to the Exchange Online Service.

## Preparation
Install the EXO V2 module to connect to Exchange Online Powershell by using MFA and modern authentification. See [Connect to Exchange Online PowerShell
](https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps) for more information.

If you can successfully connect to Exchange Online you are ready to go!

## Templates
This script ships with two template files. `signature.html` contains the HTML version and `signature.txt` the plain version of your signature. The HTML version is used while creating pure HTML emails, the plain one by creating plan text emails. Please make sure that you always edit both files!

All variables are enclosed in square brackets. A part of the signature contains static information (f.e. name of the company) and is marked by hashtags. Please edit the file accordingly.

## Usage
* Clone this project.

* Create inside the project folder an CliXML export containing the credentials of **your O365 admin account**. See [Export-Clixml](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/export-clixml?view=powershell-7.1) for more information.

  ```powershell
  Credentials = Get-Credentials
  $Credentials | Export-Clixml ./credentials.xml
  ```

* Edit the template files `signature.html` and `signature.txt` to fit your needs.
* Read the script file and edit the variables to fit your needs. (Sorry, no parameters right now, only old-fashioned variables inline)
* Execute the script manually or by using the task planner on a regular base.