         set WshShell = WScript.CreateObject("WScript.Shell")
         WshShell.AppActivate "access"
         WScript.Sleep 150
         WshShell.SendKeys "{DOWN}{RIGHT}^c"
         WshShell.AppActivate "lss"
         WScript.Sleep 150
         WshShell.SendKeys "^{TAB}"
         WScript.Sleep 150
         WshShell.SendKeys("+{F10}")
         WScript.Sleep 150
         WshShell.SendKeys "p"
         WScript.Sleep 150
         WshShell.SendKeys "~"