Function GetSerialNumber()
Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
Set colItems = objWMIService.ExecQuery("Select * From Win32_DiskDrive",,48)
For Each objItem in colItems 
	If objItem.MediaType = "Removable Media" Then
		s = s & "SerialNumber: " & objItem.SerialNumber & vbcrlf 
		s = s & "DeviceID: " & objItem.DeviceID & vbcrlf 
		s = s & "Model: " & objItem.Model & vbcrlf
	End If 
Next

MsgBox s

End Function
GetSerialNumber()