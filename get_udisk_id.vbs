Dim flag
'Set fso = CreateObject("Scripting.FileSystemObject")
'Set colDrivers = fso.Drives
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * From Win32_USBHub")

For each objItem in colItems
	uDeviceID = objItem.DeviceID
	'Wscript.Echo "UDisk Serial Number:" & uDeviceID
	flag = InStr(1, uDeviceID, "USB\VID", vbTextCompare)
	'Wscript.Echo flag
	If flag > 0 Then
		flag = InStrRev(uDeviceID, "\")
		'Wscript.Echo flag
		uDeviceID = Mid(uDeviceID, flag + 1)
		If InStr(1, uDeviceID, "&", vbTextCompare) = 0 Then
			Wscript.Echo "UDisk Serial Number:" & uDeviceID
		End If
		' uFolder = uDeviceID
		' If Not fso.FolderExists(Target_Folder & uFolder) Then 
		' 	fso.CreateFolder(Target_Folder & uFolder)
		' End If
	End If
Next