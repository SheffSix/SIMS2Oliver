Function Speak() {
	Param (
		[Parameter(Position=1)][string]$Message,
		[switch]$Time,
		[switch]$NoNewLine
	)
	$TheTime = Now
	If ($TIME) {
		$Message = "${TheTime} ${Message}"
	}
	If ($NoNewLine) {
		write-host -NoNewLine $Message
	} Else {
		write-host $Message
	}
}

Function Bork ([INT]$error = 0) {
	speak "Finished." -Time
	# Close opened text streams
	If ($ReadStream) {speak "Closing ${ReadStream}"; $ReadStream.Close(); $ReadStream = $NULL}
	If ($WriteStream) {speak "Closing ${WriteStream}"; $WriteStream.Close(); $WriteStream = $NULL}
	exit $error
}

Function Now {
	$Out = Get-Date -Date $(Date) -f $ReadableTimeFormat
	Return $Out
}
