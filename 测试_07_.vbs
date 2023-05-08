Set Wrap = CreateObject("DynamicWrapper")
Wrap.Register "user32.dll","FindWindow","i=ss","f=s", "r=l"
WindowHandle = Wrap.FindWindow("", "无标题 - 记事本")
if WindowHandle=0 then 
	MsgBox "发现窗口" 
else 
	MsgBox "没发现"
end if