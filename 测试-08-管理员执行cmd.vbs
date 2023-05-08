dim wsh

set wsh=createobject("wscript.shell")
wsh.run("cmd /c runas /user:administrator cmd")
wscript.sleep 500

wsh.sendkeys("qwe123{enter}")		rem 最好使用 微软输入法
wscript.sleep 500

