'concatenacao
fName= "Fabi"
lName="Kraft"
num1= "12"
num2=13

WScript.Echo fName +" "+ LName
WScript.Echo fName & " " &LName

'diferenca operadores qdo tem num
WScript.Echo num1 + num2
WScript.Echo num1 & num2

'vbtab e vblf (ou chr(9) e chr(10))

WScript.Echo "da tab com vbtab" & vbTab & "e pula linha com vblf" & vbLf &"nova linha"

'operadores obs: n funciona com WScript.echo, so com msgbox

a=2
b=1
c=2

msgbox a=b
MsgBox a=c
MsgBox a<b
MsgBox a>b
MsgBox a>=b
MsgBox a<=b
MsgBox a<>b '!=
