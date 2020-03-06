Set objShell = CreateObject("Wscript.Shell")
strComputer = "."
strID = ucase(InputBox("Select your system drive (Ex. C for C:\)", "Drive Information", "C"))
If len(strID) <> 1 Then
	MsgBox "Please enter only one character", vbCritical, "Drive"
Elseif Asc(strID) < 67 Or Asc(strID) > 90 Then
    	MsgBox "Unable to validate! Wrong drive letter??", vbCritical, "Something went wrong!?"
Else	
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery _
    	("Select * From Win32_LogicalDisk Where DeviceID = '" & strID & ":'")

	For Each objItem in colItems
    		MsgBox objItem.VolumeSerialNumber, vbInformation, "Your DriveID is:"
Next
    intMessage = Msgbox ("Thanks for using the DriveID Viewer!" & vbcrlf & "Would you like to visit DigitalBrekke?" & vbcrlf & "" & vbcrlf & "We have lots of other tips and tricks", _
    vbYesNo, "Thank You!")
    If intMessage = vbYes then
    objShell.Run("https://www.digitalbrekke.com")
    Else
    Wscript.Quit
    End If
End If