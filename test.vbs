'wscript.echo "test"
'MsgBox("test")
'Variaveis - letras no inicio e meio e _ e numeros no meio, nao pode espaço, nem .
'tamanho 255 char max
'name = "Fabi" ' n precisa definir se e string, int e tals
'age = 28
'digit = 34
'age = digit 'copia o valor de digit e poe em age
'name = 21 'pode mudar tipo da variavel dps tbm

'declarar variaveis
Option Explicit 'no inicio
dim ano, mes, dia
ano=2002
WScript.Echo ano
'dim hora = 3 ERRADO
dim hora : hora= 3 'CERTO
WScript.Echo hora

'operadores aritméticos

ano= 2002-2
mes=7+1
dia=2*4
hora=7/2
dim minuto : minuto=7\2 'divisao inteira
dim seg : seg = 7 mod 2 'resto (%)
dim pow : pow = 2^5 'potencia (raiz quadrada é elevado a 1/2)
dim raiz : raiz = 2^(1/2)
WScript.Echo ano, mes, dia, hora, minuto, seg, pow, raiz

