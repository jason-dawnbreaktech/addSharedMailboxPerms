This script is relatively easy to use. Just make sure you're on a PC where you can install cmdlets and change the excecution policy. 

Set-ExecutionPolicy remotesigned -force
Install-Module -Name ExchangeOnlineManagement
Import-Module -Name ExchangeOnlineManagement
Update-Module -Name ExchangeOnlineManagement 

Now run the script with powershell. It will prompt you for the calendar, user, permissions, confirm your entry, then display the calendar permissions after the change.
It will also prompt you to see if you'd like to make more changes :) 
