Option Explicit
dim obj

set obj= CreateObject("wscript.shell")

dim a,b,c

a=MsgBox("Open a program?", vbYesNo+vbQuestion+vbSystemModal)

if a=vbYes Then
obj.run "mspaint.exe"
end If

b=MsgBox("Open a folder?", vbYesNo+vbQuestion+vbSystemModal)

if b=vbYes Then
obj.run "C:\Users\FabianaDatamine\Desktop\VBScript"
end If

c=MsgBox("Open a file?", vbYesNo+vbQuestion+vbSystemModal)

if c=vbYes Then
obj.run """C:\Users\FabianaDatamine\Desktop\VBScript\nome com espaco.txt"""
end If
