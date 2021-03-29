Option Explicit
Dim objShell, Racey, intCount
Set objShell = CreateObject("WScript.Shell")
objShell.Run "notepad" 
Wscript.Sleep 1500
Racey = 1000
intCount=0

Do While intCount < 7 
objShell.SendKeys "Hello" 
objShell.SendKeys "{ENTER}"
Wscript.Sleep 800
objShell.SendKeys "Insert text to type" 
objShell.SendKeys "{ENTER}" 
Wscript.Sleep 800
objShell.SendKeys "Insert text to type" 
objShell.SendKeys "{ENTER}"
Wscript.Sleep 800
objShell.SendKeys "Insert text to type" 
objShell.SendKeys "{ENTER}"
WScript.Sleep Racey
intCount = intCount + 1
Racey = Racey - 100
Loop
Wscript.Sleep 1500
objShell.SendKeys "%F"
WScript.Sleep 1500
objShell.SendKeys "x"
WScript.Sleep 1500
objShell.SendKeys "{TAB}"
WScript.Sleep 500
objShell.SendKeys "{ENTER}"

WScript.Quit 
