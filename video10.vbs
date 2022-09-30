a= WScript.ScriptName ' da o nome do arquivo
b=WScript.ScriptFullName 'da o path
c= len(b)-len(a) ' len da path sem o arquivo
d=Left(b,c)'path sem o nome do arquivo
MsgBox a
MsgBox b
MsgBox(d)
CreateObject("wscript.shell").run (d & "video1.vbs")

'outro jeito

Msgbox(CreateObject("wscript.shell").currentdirectory & "\") ' da o path sem o nome do arquivo

'se vc nao sabe o username, use %username% no lugar
