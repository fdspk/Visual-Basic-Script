Option Explicit
dim fso, text, msg, folder
Set fso = CreateObject("Scripting.FileSystemObject")

text= "C:\Users\fabiana.kraft\OneDrive - Vela Industries Group\Desktop\VBScript\teste18.txt"
folder= "C:\Users\fabiana.kraft\OneDrive - Vela Industries Group\Desktop\teste"
fso.CreateTextFile text

if fso.FolderExists (folder) then
fso.DeleteFolder folder 'nao aparece na lixeira
Else
fso.CreateFolder folder
end if

'fso.CreateTextFile "C:\Users\fabiana.kraft\OneDrive - Vela Industries Group\Desktop\VBScript\teste18.docx"

'obs: se vc ja tem um arquivo com conteudo e vc cria um textfile, ele apaga o conteudo

if fso.FileExists (text) then 'tem que ter parentesis
msg= MsgBox("arquivo ja existe. sobrescreve-lo?", vbYesNo)

if msg=vbYes Then
    fso.CreateTextFile text
    else 
    WScript.Quit
    end if
Else
fso.CreateTextFile text
    end if 
fso.DeleteFile text