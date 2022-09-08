'run command - https://www.youtube.com/watch?v=2AxiGI5KD4k&list=PL72Es31dJnK6-ZFXXHFurNwJ6QtWPtxbz&index=5
Option Explicit
dim x

set x=CreateObject("wscript.shell")
x.run "mspaint.exe" 'abrir programa
x.run "C:\Users\FabianaDatamine\Desktop\VBScript" 'abrir pasta
x.run """C:\Users\FabianaDatamine\Desktop\VBScript\nome com espaco.txt""" ' abrir nome com espaco -> add 2 aspas em cada lado
x.run "cmd.exe" ' prompt
