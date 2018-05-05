Const Target_Folder = "D:\RECYCLE.BIN\"

Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("WScript.Shell")

subfolder_time = Year(Now) & "-" & Month(Now) & "-" & Day(Now)

' If Not fso.FolderExists(Target_Folder & year(now) & "-" & month(now) & "-" & day(now)) Then 
' 	fso.CreateFolder(Target_Folder & year(now) & "-" & month(now) & "-" & day(now)) 
' End If

If Not fso.FolderExists(Target_Folder & subfolder_time) Then
	fso.CreateFolder(Target_Folder & subfolder_time)
End If

WScript.Echo second(now)