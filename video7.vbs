'do-while do-until
'exit do é o break
'wscript.quit é o exit

Option Explicit

dim a, password
a=0
Do while a<3
MsgBox a
a=a+1
Loop

MsgBox "parte 2"
a=0
Do Until a=3
MsgBox a
a=a+1
Loop

MsgBox"parte3"
do until password="senha" ' obs: coloquei or password=vbcancel e nao da, pq vbcancel retorna "" cadeia vazia e se por "" ele nem entra no loop pq ja e valor default :(
password=InputBox("Enter Password")
Loop
