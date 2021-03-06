#Date and time stuff
$FSTimeFormat = "yyyy\-MM\-dd\THH\-mm\-ss"
$ReadableTimeFormat = "dd\/MM\/yyyy\ HH\:mm\:ss"
$StartDate = Date
$StartTime = Get-Date -Date $StartDate -f $FSTimeFormat
$ThisMonth = Get-Date -f MMM

#Set file and folder locations
$ReportsFolder = "${PSScriptRoot}\Reports"
If (Test-Path -Path "${ENV:ProgramFiles(x86)}") {
	$ProgramFiles = "${ENV:ProgramFiles(x86)}"
} else {
	$ProgramFiles = "${ENV:ProgramFiles}"
}
$SIMSDotNetFolder = "${ProgramFiles}\SIMS\SIMS .net"

#Network
$DBServer = "<SERVERNAME>\<INSTANCE>"
$DBName = "<DATABASE>"

#SIMS Reports
$SIMSStaffUsersReport = "SIMS2Oliver Staff"
$SIMSStudentUsersReport = "SIMS2Oliver Students"

#Report output files
$SIMSStaffList = "${ReportsFolder}\SIMSStaffList.csv"
$SIMSStudentList = "${ReportsFolder}\SIMSStudentList.csv"
$OutputFile = "${ReportsFolder}\SIMS2OliverImport.csv"
$StaffPhotosFile = "${ReportsFolder}\SIMS2OliverStaffPhotos.csv"

#Commands
$CommandReporter = "${SIMSDotNetFolder}\CommandReporter.exe"

#Settings
#$FirstRecord = 221
#$LastRecord = 230
