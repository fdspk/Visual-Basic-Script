Option Explicit
dim obj, nlink
set obj=CreateObject("wscript.shell")

set nlink=obj.CreateShortcut("C:\Users\fabiana.kraft\OneDrive - Vela Industries Group\Desktop\novoatalho.lnk")

nlink.targetpath="C:\Users\fabiana.kraft\Downloads"
'nlink.Description="teste"
'nLink.IconLocation = "C:\windows\system32\SHELL32.dll,137"
nlink.save