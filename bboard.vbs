' Baseboard data getter
' Outputs JSON
' V 1.1

' CONSTANTS
Const ForReading   = 1
Const ForWriting   = 2
Const ForAppending = 8

Const LogMaxSize   = 16777216 ' bytes

Const LogPath      = "C:\Program Files\Zabbix Agent\\Scripts\ScriptData\Logs\bboard.log"
Const LogPrevPath  = "C:\Program Files\Zabbix Agent\Scripts\ScriptData\Logs\bboard_prev.log"

Const OutPath      = "C:\Program Files\Zabbix Agent\Scripts\ScriptData\bboard_out.txt"

' VARIABLES 
Set objShell       = WScript.CreateObject("WScript.Shell")
Set objFSO         = CreateObject("Scripting.FileSystemObject")

' FUNCTIONS
Function FormatNow
	dnow = Now()
	logday = Day(dnow)
	If logday < 10 Then logday = "0" & logday
	logmonth = Month(dnow)
	If logmonth < 10 Then logmonth = "0" & logmonth
	loghour = Hour(dnow)
	If loghour < 10 Then loghour = "0" & loghour
	logminute = Minute(dnow)
	If logminute < 10 Then logminute = "0" & logminute
	logsec = Second(dnow)
	If logsec < 10 Then logsec = "0" & logsec
	FormatNow = logday & "/" & logmonth & "/" & Year(dnow) & " " & _
				loghour & ":" &logminute & ":" & logsec
End Function

Sub LogAddLine(line)
	If objFSO.FileExists(LogPath) Then
		Set objFile = objFSO.GetFile(LogPath)
		If ObjFile.Size < LogMaxSize Then
			Set objFile = Nothing
			Set outputFile = objFSO.OpenTextFile(LogPath, ForAppending, True, -1)
			outputFile.WriteLine(FormatNow & " - " & line)
			outputFile.Close
			Set outputFile = Nothing
		Else
			Set objFile = Nothing
			objFSO.CopyFile LogPath, LogPrevPath, True
			Set outputFile = objFSO.CreateTextFile(LogPath, ForWriting, True)
			outputFile.WriteLine(FormatNow & " - " & line)
			outputFile.Close
			Set outputFile = Nothing
		End If
	Else
		Set outputFile = objFSO.CreateTextFile(LogPath, True, -1)
		outputFile.WriteLine(FormatNow & " - " & line)
		outputFile.Close
		Set outputFile = Nothing
	End If
End Sub

Function GetWMICData(section, jsonsection)
	Set objExecObject = objShell.Exec("wmic baseboard get " & section)
	strOutput = objExecObject.StdOut.ReadAll
	outspl = Split(strOutput, Chr(10))
	outspl(1) = Replace(outspl(1), Chr(13), "")
	Set objExecObject = Nothing
	GetWMICData = jsonsection & """" & Trim(outspl(1)) & """"
	Set outspl = Nothing
End Function

' SCRIPT
LogAddLine "Script started"
jsonout = GetWMICData("manufacturer", "{""vendor"":")
LogAddLine "Vendor obtained"
jsonout = jsonout & GetWMICData("product", ",""model"":")
LogAddLine "Model obtained"
jsonout = jsonout & GetWMICData("version", ",""rev"":")
LogAddLine "Revision obtained"
jsonout = jsonout & GetWMICData("serialnumber", ",""serial"":")
LogAddLine "SN obtained"
jsonout = jsonout & "}"
Set outFile = objFSO.CreateTextFile(OutPath, True, False)
outFile.Write jsonout
outFile.Close
LogAddLine "Script finished"
Set objShell = Nothing
Set objFSO = Nothing