'video 6 - input box
'https://www.youtube.com/watch?v=JcPXsSy-5x8&list=PL72Es31dJnK6-ZFXXHFurNwJ6QtWPtxbz&index=6
'inputbox "message","title","input field", x position, y position
Option Explicit
dim name, age
name= InputBox("what is your name?", "Information", "Write your name here")

if name="" then
MsgBox"please type something in"
else
MsgBox "hello " & name, vbOKOnly, "this is your name"
end if

'elseif not 

age= InputBox("what is your age?", "Information", "Write your age here")

if IsNumeric(age) then
MsgBox age, vbOKOnly, "this is your age"
else
MsgBox "invalid"
end if

if age>=18 Then
MsgBox "you are an adult"
end if
