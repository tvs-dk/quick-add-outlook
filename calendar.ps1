# Load the Outlook COM object
$outlook = New-Object -ComObject Outlook.Application

# Prompt for date or number of days
$dateString = Read-Host "Enter date for event (DD/MM/YYYY) or number of days from now (e.g. 1d, 2d, 30d):"
$title = Read-Host "Enter event title"

# Prompt to ask if the event is private
$private = Read-Host "Is the event private? (y/n)"

# Check if the input is a date or number of days
if ($dateString -match "^\d+d$") {
    # Input is number of days
    $days = [int] ($dateString.Substring(0, $dateString.Length - 1))
    $date = (Get-Date).AddDays($days)
}
else {
    # Input is a date
    $date = [datetime]::ParseExact($dateString, "dd/MM/yyyy", $null)
}

# Check if the date falls on a weekend
if ($date.DayOfWeek -eq [System.DayOfWeek]::Saturday) {
    $moveToMonday = Read-Host "Date falls on a Saturday. Move event to following Monday? (y/n)"
    if ($moveToMonday -eq "y") {
        $date = $date.AddDays(2)
    }
}
elseif ($date.DayOfWeek -eq [System.DayOfWeek]::Sunday) {
    $moveToMonday = Read-Host "Date falls on a Sunday. Move event to following Monday? (y/n)"
    if ($moveToMonday -eq "y") {
        $date = $date.AddDays(1)
    }
}

# Create a new appointment item
$appointment = $outlook.CreateItem([Microsoft.Office.Interop.Outlook.OlItemType]::olAppointmentItem)

# Set the appointment properties
$appointment.Start = $date
$appointment.End = $date
$appointment.AllDayEvent = $true
$appointment.Subject = $title

# Set the event as private if necessary
if ($private -eq "y") {
    $appointment.Sensitivity = [Microsoft.Office.Interop.Outlook.OlSensitivity]::olPrivate
}

# Save the appointment to the calendar
$appointment.Save()

# Display a confirmation message
Write-Host "Appointment added to calendar: $title on $date as a whole day event"
Write-Host "Event privacy: $(if ($private -eq 'y') { 'Private' } else { 'Not private' })"
