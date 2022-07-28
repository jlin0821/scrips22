         set WshShell = WScript.CreateObject("WScript.Shell")
         WshShell.AppActivate "access"
         WScript.Sleep 50
         WshShell.SendKeys "^c"
         WshShell.AppActivate "lss"
         WshShell.SendKeys("+{F10}")
         WScript.Sleep 50
         WshShell.SendKeys "p"