set WshShell = WScript.CreateObject("WScript.Shell")
x=inputbox("Choose your key:")
b=inputbox("How many times do you want the key to be pressed?")
a=msgbox("The key you chose is " & x & ". Do you wish to continue?",4,"Confirmation")
If a = 7 Then

Else
 For eh = 1 To b
   WScript.Sleep 1
   WshShell.SendKeys x
 Next
End If