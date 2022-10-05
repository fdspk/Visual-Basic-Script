'subrotina

call subrotina("welcome", "hello")
call func2

sub subrotina(x, y)
msgbox y,20,x
end sub

sub func2
msgbox "Bye"
end sub

call subrotina("ola",InputBox("insira seu nome"))
