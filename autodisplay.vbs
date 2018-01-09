Set ws=CreateObject("Wscript.Shell")
ws.run "notepad"
Wscript.Sleep 1000
'ws.AppActivate "Untitled-notepad"
ws.SendKeys "CrypTool (Starting example for the CrypTool version family 1.x)"