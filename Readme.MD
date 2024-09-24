# First install the Exchange Module

```
Set-ExecutionPolicy remotesigned -force
Install-Module -Name ExchangeOnlineManagement
Import-Module -Name ExchangeOnlineManagement
Update-Module -Name ExchangeOnlineManagement 
```
Now run the script with powershell. It will prompt you for the calendar, user, permissions, confirm your entry, then display the calendar permissions after the change.
It will also prompt you to see if you'd like to make more changes :) 
