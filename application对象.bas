Attribute VB_Name = "模块8"
Sub application方法()
'wait方法——注意时间格式
application.Wait "18:23:00" '在此时间再运行
newHour = Hour(Now())
newMinute = Minute(Now())
newSecond = Second(Now()) + 10
waitTime = TimeSerial(newHour, newMinute, newSecond)
application.Wait waitTime
If application.Wait(Now + TimeValue("0:00:10")) Then
    MsgBox "Time expired"
End If
'volatile易失性函数，随时发生变化，如now、today，rand，
'Union合并单元格
'Undo撤消最后一次用户界面操作,本示例必须放在宏的第一行
'SetDefaultChart在16版中没有此方法。
'Save 后面无参数

End Sub
Function myFUN(cell As Range)
application.Volatile '如果没有这句话，如果b5为cell那么b6变化就不会影响结果，因为b6不是myFUN的参数。
myFUN = cell.Value + cell.Offset(1, 0).Value

End Function
Sub ahl()
application.setdefalutchart FormatName:="Monthly Sales"
End Sub
