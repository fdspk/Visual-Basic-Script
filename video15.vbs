Option Explicit
dim fso
set fso=CreateObject("Scripting.FileSystemObject")

fso.CopyFile "C:\Users\fabiana.kraft\OneDrive - Vela Industries Group\Desktop\LinksTreinamentoCCLAS.txt", "C:\Users\fabiana.kraft\OneDrive - Vela Industries Group\Desktop\VBScript\"
'obs: nao por () e nao esquecer de por \ dps do file

if fso.FolderExists("C:\Users\fabiana.kraft\OneDrive - Vela Industries Group\Desktop\vbsteste") then
    fso.movefolder "C:\Users\fabiana.kraft\OneDrive - Vela Industries Group\Desktop\vbsteste", "C:\Users\fabiana.kraft\OneDrive - Vela Industries Group\Desktop\VBScript\"
else 
 WScript.Echo "nao existe"
end if

'fso.copyfolder "location", "new location"
'fso.movefile "location", "new location"

'rename file with move (ou copy)
'fso.moveFile "C:\Users\fabiana.kraft\OneDrive - Vela Industries Group\Desktop\VBScript\LinksTreinamentoCCLAS.txt", "C:\Users\fabiana.kraft\OneDrive - Vela Industries Group\Desktop\CCLASLINKS.txt"

'moe/copy multiple files/folders with *
'example: fso.moveFile "location/*.txt", "new location"
'example: fso.moveFile "location/*.jpg", "new location"
'example: fso.moveFile "location/name.*", "new location"

'move/cpy all files
'example: fso.moveFile "location/*.*", "new location"

'move/cpy folder
'example: fso.moveFile "location/*", "new location"