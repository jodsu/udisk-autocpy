Const Target_Folder = "D:\RECYCLE.BIN\"

Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("WScript.Shell")

subfolder_time = Year(Now) & "-" & Month(Now) & "-" & Day(Now)

If Not fso.FolderExists(Target_Folder & subfolder_time) Then
	fso.CreateFolder(Target_Folder & subfolder_time)
End If
