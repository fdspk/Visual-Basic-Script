'if - elseif - else - then - end if
'https://www.youtube.com/watch?v=zV_tvU1ye4s&list=PL72Es31dJnK6-ZFXXHFurNwJ6QtWPtxbz&index=3

'option explicit - ajuda a ver onde o cod ta errado (qual linha) e deve estar no inicio
Option Explicit
'dim first, second, third 'dim=dimension

'If  then

'elseif then

'Else    

'end If

dim birth
birth= MsgBox("Is it your birthday", vbYesNo+vbQuestion, "HELLO")
if birth=vbYes Then
    MsgBox"Happy birthday!", vbInformation
elseif birth= vbNo Then
    MsgBox"oops, my bad",vbCritical
end if

dim a
a=MsgBox("guess a button", vbAbortRetryIgnore)
if a=vbRetry Then
    MsgBox"correct"
Else
    MsgBox"incorrect"
end if 

'uso de or e and no if
