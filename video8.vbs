'sendkeys (manda teclas do teclado uma a uma)
'createobject("wscript.shell").sendkeys ""
Option Explicit
dim obj, browse
set obj=CreateObject("wscript.shell")
obj.run """C:\Users\FabianaDatamine\Desktop\VBScript\nome com espaco.txt"""
WScript.Sleep 1000 'espera 1 segundo senao n da tempo pro txt abrir e escreve no vscode
obj.SendKeys"isto nao estava aqui antes"
obj.SendKeys"{enter}"
WScript.Sleep 1000
obj.SendKeys"olha a magica {enter}"
WScript.Sleep 1000
obj.SendKeys"^s" 'salva ctrl é ^

