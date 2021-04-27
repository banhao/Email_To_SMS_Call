# Email To SMS Call
[![Minimum Supported PowerShell Version](https://img.shields.io/badge/PowerShell-5.1+-purple.svg)](https://github.com/PowerShell/PowerShell) ![Cross Platform](https://img.shields.io/badge/platform-windows-lightgrey)
[![PowerShell Gallery Version](https://img.shields.io/powershellgallery/v/Email_To_SMS_Call)](https://www.powershellgallery.com/packages/Email_To_SMS_Call) [![PowerShell Gallery](https://img.shields.io/powershellgallery/dt/Email_To_SMS_Call)](https://www.powershellgallery.com/packages/Email_To_SMS_Call)

"Email_To_SMS_Call.ps1" is used to monitor a specific mailbox that some important emails you don't want to miss and check later. When the condition matched, the script will send a SMS message to you or make a phone call to you.

It can be used to monitor IT system warning emails or security alert emails to different team by different time range. 

This script use the Twilio API to send the SMS and make phone call. So please register a Twilio phone number first.
https://www.twilio.com/

Before you run the script Install the Exchange Web Services Managed API 2.2. 
https://www.microsoft.com/en-us/download/details.aspx?id=42951
