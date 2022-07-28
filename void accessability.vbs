         set WshShell = WScript.CreateObject("WScript.Shell")
         WshShell.AppActivate "lss"
         WScript.Sleep 50
         WshShell.SendKeys "v{LEFT}{UP}{RIGHT}"