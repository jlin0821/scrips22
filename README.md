
vbs app
         set WshShell = WScript.CreateObject("WScript.Shell")
         WshShell.AppActivate "lss"
         WScript.Sleep 50
         WshShell.SendKeys "+^{TAB}"
         WScript.Sleep 50
         WshShell.AppActivate "Documents"

def string_check(string_length, my_string):
   result_string = []
   words = my_string.split(" ")
   for x in words:
      if len(x) > string_length:
         result_string.append(x)
   return result_string
string_length = 5
my_string =""

result_string = string_check(string_length, my_string)
print(' '.join(result_string))
