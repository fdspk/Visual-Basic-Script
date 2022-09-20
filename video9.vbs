'strings: len, left(corta pela esq) and right (corta pela dir)
MsgBox(len("Hello world")) 'e basicamente o strlen
s="hello my friend"
MsgBox(left(s,5)) 'hello
MsgBox(s)'hello my friend
MsgBox(right(s,6)) 'friend
MsgBox(right(s,9)) 'my friend
s="hello world"
MsgBox(left(s,5)) 'hello
MsgBox(right(s,5)) 'world
a="I love dogs"
b="cats are stupid"
MsgBox(left(a,7) & left(b,4) & right(a,5) & right(b, 11))