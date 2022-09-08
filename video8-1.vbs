Option Explicit
dim obj, browse
set obj=CreateObject("wscript.shell")
MsgBox"parte 2"
browse=InputBox("O que voce quer pesquisar?")
obj.run"msedge.exe"
WScript.Sleep 1000
obj.SendKeys browse
obj.SendKeys "{enter}"
