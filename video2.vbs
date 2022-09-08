'continuação video1
'_ move para proxima secao
a=MsgBox("oi",vbAbortRetryIgnore +_
 vbExclamation+vbDefaultButton2+vbSystemModal,_
 "teste") 'em vez de por o num, poe o comando  obs: no defalutbutton o mouse nao vai pro botao
'o vbsystemmodal faz com que a janela fique por cima independente do q vc tem aberto

if a=vbAbort then MsgBox"quit", vbCritical
