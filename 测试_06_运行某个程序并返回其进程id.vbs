Set objShell = CreateObject("WScript.Shell")

' Run the program asynchronously
Set objProcess = objShell.Exec("notepad")

' Print the process ID
WScript.Echo objProcess.ProcessID
