# quick-add-outlook.ps1 [link to script](https://github.com/tvs-dk/quick-add-outlook/blob/main/calendar.ps1)
## A Powershell script to quickly add a full-day event (private or not private) to your Outlook calendar by either: 
 * writing how many days from now to add the event by writing 1d, 2d etc. 
 * writing the date in the format DD/MM/YYYY



### Further features:
* if the event falls on a weekend, the user is asked if the event should be moved to the following Monday

* If the PC is protected from running Powershell script, you can deactivate that by running the command
```
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser
```
### Or write the following command. Followed by "Bypass"
```
Set-ExecutionPolicy -Scope CurrentUser
```

### or run only this trusted script by the command:
```
powershell.exe -noprofile -executionpolicy bypass -file .\calendar.ps1
```

### Protip:
Put the script into a function. The name of the function in this example is "Get-Cal"

```
function Get-Cal {
#put script here script
}
```

and save this script in your user folder as a .ps1
```
C:\Users\your-user-name>
```
and load the script as a module in your Powershell profile. See how to do that here: [Setting-up-the-Powershell-profile](https://github.com/tvs-dk/Notes-for-Powershell/wiki/Setting-up-the-Powershell-profile)
```
Import-Module .\calendar.ps1
```
Then you can always load the script in a new Powershell window with the command:
```
Get-Cal
```
