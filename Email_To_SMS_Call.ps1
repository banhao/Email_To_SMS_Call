 
<#PSScriptInfo

.VERSION 1.0.0

.GUID 134de175-8fd8-4938-9812-053ba39eed83

.AUTHOR banhao@gmail.com

.COMPANYNAME

.COPYRIGHT

.TAGS

.LICENSEURI

.PROJECTURI

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#> 

<# 

.DESCRIPTION 
"Email_To_SMS_Call.ps1" is used to monitor a specific mailbox that some important emails you don't want to miss and check later. When the condition matched, the script will send a SMS message to you or make a phone call to you.

It can be used to monitor IT system warning emails or security alert emails to different team by different time range. 

This script use the Twilio API to send the SMS and make phone call. So please register a Twilio phone number first.
https://www.twilio.com/

Before you run the script Install the Exchange Web Services Managed API 2.2. 
https://www.microsoft.com/en-us/download/details.aspx?id=42951

Please check the License before you download this script, if you don't agree with the License please don't download and use this script. https://github.com/banhao/Email_To_SMS_Call/blob/master/LICENSE

#> 


<#
.SYNOPSIS
  <>

.DESCRIPTION
  <>

.PARAMETER <Parameter_Name>
  <>

.INPUTS
  <>

.OUTPUTS
  <>

.NOTES
 
  Version:        1.0.0
  Author:         <HAO BAN/banhao@gmail.com>
  Creation Date:  <01/24/2020>
  Purpose/Change: Initial version
  
.EXAMPLE
How to config the init.xml
<EMAILFROM> is used to define which sender you want to monitor. It should be unique in the XML file.
	<EMAILSUBJECT> is used to define which email subject you want to monitor. You can define different <EMAILSUBJECT> under the <EMAILFROM>. If you want to match any Subject from one Email Address, the use the "*"
			<DaysOfWeek> is from Monday to Sunday, If you want to match all days then use "*"
				<HOUR> is hour range. 24 hours format 00-23. The script will get the email received hour and compare with this element.
					<TWILIOPHONETO> is the phone number you want to call.

The following example will monitor the sender's email address is from "example@example.com". When the Subject is match "TEST 1" and the Day is Saturday from 09 to 20 it will call "SaturdaySupportTeam1PhoneNumber", from 00-08 it will call "SaturdaySupportTeam2PhoneNumber", from 21-23 it will call "SaturdaySupportTeam3PhoneNumber". If it is Sunday, at anytime it will call "SundaySupportTeamPhoneNumber"

		<EMAILFROM>example@example.com
			<EMAILSUBJECT>TEST 1
				<DaysOfWeek>Saturday
					<HOUR>09-20
						<TWILIOPHONETO>SaturdaySupportTeam1PhoneNumber</TWILIOPHONETO>
					</HOUR>
					<HOUR>00-08
						<TWILIOPHONETO>SaturdaySupportTeam2PhoneNumber</TWILIOPHONETO>
					</HOUR>
					<HOUR>21-23
						<TWILIOPHONETO>SaturdaySupportTeam3PhoneNumber</TWILIOPHONETO>
					</HOUR>
				</DaysOfWeek>
				<DaysOfWeek>Sunday
					<HOUR>*
						<TWILIOPHONETO>SundaySupportTeamPhoneNumber</TWILIOPHONETO>
					</HOUR>
				</DaysOfWeek>
			</EMAILSUBJECT>
		</EMAILFROM>



#>
#-------------------------------------------------------------------------------------------------------------------------------------------------------
#variables
cls

$ADUSERNAME = Select-Xml -Path .\init.xml -XPath "//SETTINGS/ADUSERNAME" | foreach {$_.node.InnerXML}
$DOMAIN = Select-Xml -Path .\init.xml -XPath "//SETTINGS/DOMAIN" | foreach {$_.node.InnerXML}
$EMAILADDRESS = Select-Xml -Path .\init.xml -XPath "//SETTINGS/EMAILADDRESS" | foreach {$_.node.InnerXML}
#$EXCHANGESRV = Select-Xml -Path .\init.xml -XPath "//SETTINGS/EXCHANGESRV" | foreach {$_.node.InnerXML}
$EWSDLLPATH = Select-Xml -Path .\init.xml -XPath "//SETTINGS/EWSDLLPATH" | foreach {$_.node.InnerXML}
$INTERVAL = [int]$(Select-Xml -Path .\init.xml -XPath "//SETTINGS/INTERVAL" | foreach {$_.node.InnerXML})
$TWILIOSMSAPIURL = Select-Xml -Path .\init.xml -XPath "//SETTINGS/TWILIOSMSAPIURL" | foreach {$_.node.InnerXML}
$TWILIOVOICEAPIURL = Select-Xml -Path .\init.xml -XPath "//SETTINGS/TWILIOVOICEAPIURL" | foreach {$_.node.InnerXML}
$VOICEURL = Select-Xml -Path .\init.xml -XPath "//SETTINGS/VOICEURL" | foreach {$_.node.InnerXML}
$TWILIOSID = Select-Xml -Path .\init.xml -XPath "//SETTINGS/TWILIOSID" | foreach {$_.node.InnerXML}
$TWILIOTOKEN = Select-Xml -Path .\init.xml -XPath "//SETTINGS/TWILIOTOKEN" | foreach {$_.node.InnerXML}
$TWILIOPHONEFROM = Select-Xml -Path .\init.xml -XPath "//SETTINGS/TWILIOPHONEFROM" | foreach {$_.node.InnerXML}
$xmlEMAILFROM = $(Select-Xml -Path .\init.xml -XPath "//SETTINGS/POLICY/EMAILFROM" | foreach {$_.node}).'#text'
$xmlSETTINGS = [xml](Get-Content -Path .\init.xml)
$ADCREDENTIAL = Get-Credential -credential $ADUSERNAME

function xmlEMAIL-SUBJECT-Node {
	if ( [string]$($xmlEMAILSUBJECT.'#text').Split('',[System.StringSplitOptions]::RemoveEmptyEntries) -eq "*" ){
		$xmlDaysOfWeek = $($xmlEMAILSUBJECT | where-object{$_.parentnode.'#text' -match $SENDERADDRESS}).DaysOfWeek
		xmlEMAIL-DaysOfWeek-Node
	}else{
		if ( $([string]$xmlEMAILSUBJECT.'#text').contains($EMAIL.Subject) ){
			$xmlDaysOfWeek = $($xmlEMAILSUBJECT | where-object{$_.'#text' -match $EMAIL.Subject}).DaysOfWeek
			xmlEMAIL-DaysOfWeek-Node
		}
	}
}

function xmlEMAIL-DaysOfWeek-Node {
	if ( [string]$($xmlDaysOfWeek.'#text').Split('',[System.StringSplitOptions]::RemoveEmptyEntries) -eq "*" ){
		$xmlHOUR = $($xmlDaysOfWeek | where-object{$_.parentnode.'#text' -match $EMAIL.Subject}).HOUR
		xmlEMAIL-HOUR-Node
	}else{
		if ( $([string]$xmlDaysOfWeek.'#text').contains($EMAILDAYOFWEEK) ){
			$xmlHOUR = $($xmlDaysOfWeek | where-object{$_.'#text' -match $EMAILDAYOFWEEK}).HOUR
			xmlEMAIL-HOUR-Node
		}
	}
}

function xmlEMAIL-HOUR-Node {
	if ( [string]$($xmlHOUR.'#text').Split('',[System.StringSplitOptions]::RemoveEmptyEntries) -eq "*" ){
		$global:TWILIOSMSTO = $xmlHOUR.TWILIOPHONETO
	}else{
		for ( $i = 0 ; $i -lt $($xmlHOUR.length); $i++ ){	##Get phone number by HOUR##
			if ( $EMAILHOUR -ge $xmlHOUR[$i].'#text'.Split("{-}")[0] -and $EMAILHOUR -le $xmlHOUR[$i].'#text'.Split("{-}")[1] ){
				$global:TWILIOSMSTO = $xmlHOUR[$i].TWILIOPHONETO
			}
		}
	}
}


function MAIN {
	date
	Import-Module $EWSDLLPATH
	$SERVICE = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2)
	$SERVICE.Credentials = New-Object Net.NetworkCredential($ADUSERNAME, $ADCREDENTIAL.PASSWORD, $DOMAIN)
	$SERVICE.AutodiscoverUrl($EMAILADDRESS)
	$INBOX = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($SERVICE,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
	$PROPERTYSET = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
	$PROPERTYSET.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
	if ( $INBOX.TotalCount -ne 0 ){
		$ITEMS = $INBOX.FindItems($INBOX.TotalCount)
		foreach ( $EMAIL in $ITEMS.Items ){
			# only get unread emails
			if( $EMAIL.isread -eq $false ){
				# load the property set to get to the body
				$EMAIL.load($PROPERTYSET)
				if ( $([string]$xmlEMAILFROM).contains($EMAIL.From.Address) ){	##Email address is in the XML file##
					$EMAILDAYOFWEEK = $(Get-Date -Date "$($EMAIL.DateTimeReceived)").DayOfWeek	##Get the email received day of week##
					$EMAILHOUR = Get-Date -Date "$($EMAIL.DateTimeReceived)" -Format HH	##Get the email received hour##
					$SENDERADDRESS = [string]$($EMAIL.From.Address)
					$xmlEMAILSUBJECT = $xmlSETTINGS.SelectNodes("//SETTINGS/POLICY/EMAILFROM[contains(.,'$SENDERADDRESS')]/EMAILSUBJECT")
					xmlEMAIL-SUBJECT-Node
					$SMSBODY = $($EMAIL.Subject) + "," + $($EMAIL.DateTimeReceived)
					$TWILIOSMSBODY = @{ To = $TWILIOSMSTO; From = $TWILIOPHONEFROM; Body = $SMSBODY }
					$TWILIOVOICEBODY = @{ To = $TWILIOSMSTO; From = $TWILIOPHONEFROM; Url = $VOICEURL }
					$TWILIOCREDENTIAL = New-Object System.Management.Automation.PSCredential($TWILIOSID, $($TWILIOTOKEN | ConvertTo-SecureString -asPlainText -Force))
					$SMSRESULTS = Invoke-WebRequest -Uri $TWILIOSMSAPIURL -Method 'Post' -Credential $TWILIOCREDENTIAL -Body $TWILIOSMSBODY
					$CALLRESULTS = Invoke-WebRequest -Uri $TWILIOVOICEAPIURL -Method 'Post' -Credential $TWILIOCREDENTIAL -Body $TWILIOVOICEBODY
					$EMAIL.isRead = $true
					$EMAIL.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve)
				}
			}
		}
	}else{ Write-OutPut "==============There is no email in the inbox==================" }
}

# Main Procedure
if ( $INTERVAL -eq 0 ){
	MAIN
}else{
	while($true){
		MAIN
		Write-Host -NoNewline "==============After"$INTERVAL" seconds will check again=============="
		""
		Start-Sleep -s $INTERVAL
	}
}
