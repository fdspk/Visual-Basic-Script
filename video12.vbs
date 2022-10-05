Option Explicit
Dim ie,x
set x= CreateObject("wscript.shell")
set ie= CreateObject("InternetExplorer.Application")
ie.Navigate"https://www.bing.com/" 
ie.visible=1
ie.Toolbar=0
ie.statusbar=0
ie.Height=720
ie.Width=1000
ie.top=0
ie.left=0
ie.resizable=0

sub waitForLoad
do while ie.busy
wscript.sleep200
loop
end sub
'WScript.Sleep 5000  'ou
'call waitForLoad
'x.sendkeys"dog"
'x.sendkeys"{enter}"
