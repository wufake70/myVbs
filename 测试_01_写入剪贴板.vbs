Set Ws=CreateObject("Wscript.Shell")


rem Clipboard="MsHta vbscript:ClipBoardData.setData(""Text"","""&"你好世界"&""")(Window.Close)"
rem Ws.Run(Clipboard)


rem 使用dos命令 粘贴到剪贴板
rem 该方法的弊端是 cmd窗体会 闪现并闪退
rem 一定要加上 "cmd /c",并且注意空格

rem clip = "cmd /c echo " + "你敢 嘛？？" & "| clip"

rem 使用数组拼接
dim li
li = Array("cmd /c echo ","你敢  嘛？？","| clip")
clip = join(li,"")


rem 这里exec 和run 都可以运行 dos命令
rem Ws.exec(clip)
ws.run clip,0		rem 隐藏dos 窗口

rem run的第二个参数解析
rem 0 隐藏一个窗口并激活另一个窗口。
rem 1 激活并显示窗口。如果窗口处于最小化或最大化状态，则系统将其还原到原始大小和位置。第一次显示该窗口时，应用程序应指定此标志。
rem 2 激活窗口并将其显示为最小化窗口。
rem 3 激活窗口并将其显示为最大化窗口。
rem 4 按最近的窗口大小和位置显示窗口。活动窗口保持活动状态。
rem 5 激活窗口并按当前的大小和位置显示它。
rem 6 最小化指定的窗口，并按照 Z 顺序激活下一个顶部窗口。
rem 7 将窗口显示为最小化窗口。活动窗口保持活动状态。
rem 8 将窗口显示为当前状态。活动窗口保持活动状态。
rem 9 激活并显示窗口。如果窗口处于最小化或最大化状态，则系统将其还原到原始大小和位置。还原最小化窗口时，应用程序应指定此标志。
rem 10 根据启动应用程序的程序状态来设置显示状态。