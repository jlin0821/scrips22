
vbs app
         set WshShell = WScript.CreateObject("WScript.Shell")
         WshShell.AppActivate "lss"
         WScript.Sleep 50
         WshShell.SendKeys "+^{TAB}"
         WScript.Sleep 50
         WshShell.AppActivate "Documents"

