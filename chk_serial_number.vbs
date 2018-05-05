strComputer = "." 

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_DiskDrive",,48) 
For Each objItem in colItems 
   s = s & "SerialNumber: " & objItem.SerialNumber & vbcrlf 
   s = s & "Model: " & objItem.Model & vbcrlf 
Next

MsgBox s