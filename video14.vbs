Option Explicit
dim fso


set fso=CreateObject("Scripting.FileSystemObject")

WScript.Echo(fso.FileExists("C:\Users\fabiana.kraft\OneDrive - Vela Industries Group\Desktop\LinksTreinamentoCCLAS.txt"))

WScript.Echo(fso.FileExists("C:\Users\fabiana.kraft\OneDrive - Vela Industries Group\Desktop\LinksTreinamentoCCLASs.txt"))

WScript.Echo(fso.folderExists("C:\Users\fabiana.kraft\OneDrive - Vela Industries Group\Desktop"))