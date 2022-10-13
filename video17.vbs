Option Explicit
dim obj, path
set obj= CreateObject("wscript.shell")

'msgbox obj.SpecialFolders("Desktop") & "\"
 
path=obj.SpecialFolders("Desktop")

obj.run path 
'nao funciona :/ sera que e por causa do onedrive? os outros atalhos alem de desktop nem rodam...