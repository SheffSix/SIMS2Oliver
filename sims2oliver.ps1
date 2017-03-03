<#

.SYNOPSIS
Creates a CSV file to synchronise Oliver borrowers for students and staff using reports generated from SIMS.

#>
. "${PSScriptRoot}\coreFunctions.ps1"
. "${PSScriptRoot}\Config.ps1"
Function RunSIMSReport($ReportName, $OutputFile) {
	If (Test-Path -Path ${OutputFile}){
		Remove-Item -Path ${OutputFile} -Force -ErrorAction SilentlyContinue
	}
	speak "Running SIMS report ${Reportname}..." -Time
	&  $CommandReporter /TRUSTED /SERVERNAME:"$DBServer" /DATABASENAME:"$DBName" /REPORT:"$ReportName" /OUTPUT:"$OutputFile"
	if ($? -and (Test-Path -Path ${OutputFile})) {
		speak "${OutputFile} created successfully." -Time
		speak ""
	} else {
		speak "Report failed" -Time
		speak ""
	}
}

#Run Reports
RunSIMSReport $SIMSStaffUsersReport $SIMSStaffList
RunSIMSReport $SIMSStudentUsersReport $SIMSStudentList

$WriteStream = [System.IO.StreamWriter] "${OutputFile}"
$PhotoStaff = [System.IO.StreamWriter] "${StaffPhotosFile}"
#$WriteStream.WriteLine("studentCode,address,personalName,familyName,year,group,birthday,gender,type,mailTitle,email,phone,Alias")

If (Test-Path -Path $SIMSStaffList) {
	$ReadStream = [System.IO.StreamReader] "${SIMSStaffList}"
	$User = $ReadStream.ReadLine()
} else {
	speak "ERROR: ${SIMSStaffList} does not exist." -Time
	bork 1
}
$Record = 1
While (!($ReadStream.EndOfStream)) {
	$User = $ReadStream.Readline()
	If ((!($FirstRecord) -or $Record -ge $FirstRecord) -and (!($LastRecord) -or $Record -le $LastRecord)) {
		$User = $User.Split(",").Replace("""","").Trim()
		$Fields = @($User).Count
		#If ($Fields -eq 16) {
			$PhotoStaff.WriteLine("$($User[0]),ID$($User[0])")
			$User[0] = "ID$($User[0])"
			$User[3] = Get-Date -Date $($User[3]) -f "dd/MM/yyy"
			$WriteStream.Write("""$($User[0])"",") # studentCode
			$WriteStream.Write("""")													# address
			If ($User[9] -ne "") { $WriteStream.Write("Apartment $($User[9])`n") }		# apartment
			If ($User[10] -ne "") { $WriteStream.Write("$($User[10])`n") }					# house name
			If ($User[11] -ne "") { $WriteStream.Write("$($User[11]) ") }					# house nubmer
			If ($User[12] -ne "") { $WriteStream.Write("$($User[12])") }					# street name
			For ($i=13; $i -lt $Fields; $i++) {
				If ($User[$i] -ne "") { $WriteStream.Write("`n$($User[$i])") }
			}
			$WriteStream.Write(""",")													# end address
			$WriteStream.Write("""$($User[1])"",") # personalName
			$WriteStream.Write("""$($User[2])"",") # familyName
			$WriteStream.Write(""""",") # year
			$WriteStream.Write(""""",") # group
			$WriteStream.Write("""$($User[3])"",") # birthday
			$WriteStream.Write("""$($User[4])"",") # gender
			$WriteStream.Write("""Staff"",") # type
			$WriteStream.Write("""$($User[5])"",") # mailTitle
			$WriteStream.Write("""$($User[6])"",") # email
			$WriteStream.Write("""$($User[7])"",") # phone
			$WriteStream.Write("""$($User[0])"",") # alias
			$WriteStream.WriteLine("""Staff""") # borrower loan category
		#} Else {
		# If ($Fields -gt 16) {
			# echo "Incorrect number of fields, $Fields, for $($User[0])"
			# ForEach ($Line in $User) {
				# write-host "	$Line"
			# }
		# }
	}
	$Record ++
	$User = $null
}
$ReadStream.Close()

###
# $writestream.close()
# exit
###

If (Test-Path -Path $SIMSStudentList) {
	$ReadStream = [System.IO.StreamReader] "${SIMSStudentList}"
	$User = $ReadStream.ReadLine()
} else {
	speak "ERROR: ${SIMSStudentList} does not exist." -Time
	bork 1
}
$Record = 1
While (!($ReadStream.EndOfStream)) {
	$User = $ReadStream.Readline()
	If ((!($FirstRecord) -or $Record -ge $FirstRecord) -and (!($LastRecord) -or $Record -le $LastRecord)) {
		$User = $User.Split(",").Replace("""","").Trim()
		$Fields = @($User).Count
		#If ($Fields -eq 18) {
			$User[5] = Get-Date -Date $($User[5]) -f "dd/MM/yyy"
			# Switch ($User[3]) {
				# 9 { $BorrowerLoanCategory = "Years 9 and 10" }
				# 10 { $BorrowerLoanCategory = "Years 9 and 10" }
				# 11 { $BorrowerLoanCategory = "Year 11" }
				# 12 { $BorrowerLoanCategory = "Years 12 and 13" }
				# 13 { $BorrowerLoanCategory = "Years 12 and 13" }
				# default	{ $BorrowerLoanCategory = "Years 7 and 8" }
			# }
			$BorrowerLoanCategory = ""
			$WriteStream.Write("""$($User[0])"",") # studentCode
			$WriteStream.Write("""")													# address
			If ($User[10] -ne "") { $WriteStream.Write("Apartment $($User[10])`n") }		# apartment
			If ($User[11] -ne "") { $WriteStream.Write("$($User[11])`n") }				# house name
			If ($User[12] -ne "") { $WriteStream.Write("$($User[12]) ") }					# house nubmer
			If ($User[13] -ne "") { $WriteStream.Write("$($User[13])") }					# street name
			For ($i=14; $i -lt $Fields; $i++) {
				If ($User[$i] -ne "") { $WriteStream.Write("`n$($User[$i])") }
			}			
			$WriteStream.Write(""",")													# end address
			$WriteStream.Write("""$($User[1])"",") # personalName
			$WriteStream.Write("""$($User[2])"",") # familyName
			$WriteStream.Write("""$($User[3])"",") # year
			$WriteStream.Write("""$($User[4])"",") # group
			$WriteStream.Write("""$($User[5])"",") # birthday
			$WriteStream.Write("""$($User[6])"",") # gender
			$WriteStream.Write("""Student"",") # type
			$WriteStream.Write("""$($User[7])"",") # mailTitle
			$WriteStream.Write("""$($User[8])"",") # email
			$WriteStream.Write("""$($User[9])"",") # phone
			$WriteStream.Write("""$($User[0])"",") # alias
			$WriteStream.WriteLine("""$BorrowerLoanCategory""") # borrower loan category
		# } Else 
		if ($Fields -gt 18) {
			echo "Incorrect number of fields, $Fields, for $($User[0])"
			ForEach ($Line in $User) {
				write-host "	$Line"
			}
		}
	}
	$Record ++
	$User = $null
}
$ReadStream.Close()
$PhotoStaff.Close()
$WriteStream.Close()

#Finish
bork