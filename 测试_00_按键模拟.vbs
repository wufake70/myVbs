Rem 消息发送
Dim wsh,msg,y,t

set wsh=createobject("wscript.shell")

wsh.popup "准备好要发送消息的窗口","0","提示窗口","1"

y 	= InputBox("发送次数:")
msg = InputBox("发送的内容:")
t 	= InputBox("时间间隔(ms):")

Rem if y = "1" then  MsgBox "999" 		1与"1" 	相等


wsh.popup "两秒后运行....","2","提示窗口"

wscript.Sleep 2000
for i = 1 to y 
wscript.sleep t
wsh.sendKeys msg
wsh.sendKeys "{enter}"				Rem %s alt+s 类似于回车发送消息 只能发送英文字符
Next
wscript.quit
