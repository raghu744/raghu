set x=createobject("wscript.shell")
x.sendkeys "^"+"{ESC}"
wscript.sleep 1000
x.sendkeys "cmd"
wscript.sleep 1000
x.sendkeys "{ENTER}"
