dim strPath

Set objShell = CreateObject("Wscript.Shell")  

strPath = Wscript.Arguments(0)  
strPath = "explorer.exe /e," & strPath  
objShell.Run strPath 

msgbox strPath