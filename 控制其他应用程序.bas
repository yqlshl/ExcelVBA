Attribute VB_Name = "控制其他应用程序"

Sub 操作其他应用程序() '必须在excelVBA中运行，wordVBA中没有sendkeys方法

Dim labelshop
Rem vbNormalFocus表示正常显示，vbMaximizedFocus最大化模式，vbMinimizedFocus最小化
labelshop = Shell("Explorer.exe 程序的地址", vbNormalFocus)

Rem excelVBA中用wait方法，wordVBA只能用sleep api
Application.wait (Now + timevale("00:00:05"))


Application.SendKeys "按键名" 或者 变量


Application.SendKeys "我是一串字符,或者一篇文章.如果字符串过长,可以作用变量"

Application.SendKeys "^P"                '调出打印窗口
Application.SendKeys "5"                  '设置打印的份数
Application.SendKeys "~"                '回车开始打印

Application.SendKeys 参数附表:
键名
        代码

Backspace
{BACKSPACE} 或 {BS}

Break
{BREAK}

Caps Lock
{CAPSLOCK}

清除
{CLEAR}

Delete 或 Del
{DELETE} 或 {DEL}

向下键
{DOWN}

结束
{END}

Enter (数字小键盘)
{ENTER}

Enter
~（波形符）

Esc
{ESCAPE} 或 {ESC}

帮助
{HELP}

主页
{HOME}

Ins
{INSERT}

向左键
{LEFT}

Num Lock
{NUMLOCK}

PageDown
{PGDN}

PageUp
{PGUP}

Return
{RETURN}

向右键
{RIGHT}

Scroll Lock
{SCROLLLOCK}

Tab
{TAB}

向上键
{UP}

F1 到 F15
{F1} 到 {F15}
</table>
还可指定与 Shift 和/或 Ctrl 和/或 Alt 组合使用的键。若要指定与其他键组合使用的键，可使用下表。
[Table]
要组合的键
在键代码之前添加

Shift
+（加号）

Ctrl
^（插入符号）

Alt
%（百分号）



End Sub

Sub 关闭()
  Set msoft = GetObject("winmgmts:").execquery("select * from win32_process where name like '%.exe'")
    For Each S In msoft       '用s.name可以列出所有运行的exe文件
      If S.Name = "LabelShop.exe" Then
         S.Terminate
      End If
    Next
End Sub






