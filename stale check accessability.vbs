         set WshShell = WScript.CreateObject("WScript.Shell")
         WshShell.AppActivate "access"
         WScript.Sleep 50
         WshShell.SendKeys "{HOME}"
         WScript.Sleep 150
         WshShell.SendKeys "^c"
         WshShell.AppActivate "lss"
         WScript.Sleep 350
         WshShell.SendKeys("{HOME}^f")
         WScript.Sleep 150
         WshShell.SendKeys("^v")
         WScript.Sleep 150
         WshShell.SendKeys("~")

