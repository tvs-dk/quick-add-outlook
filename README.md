# quick-add-outlook.ps1
## A Powershell script to quickly add a full-day event (private or not private) to your Outlook calendar by either: 
 * writing how many days from now to add the event by writing 1d, 2d etc. 
 * writing the date in the format DD/MM/YYYY

[link to script](https://github.com/tvs-dk/quick-add-outlook/blob/main/calendar.ps1)

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
