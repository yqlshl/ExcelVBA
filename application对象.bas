Attribute VB_Name = "ģ��8"
Sub application����()
'wait��������ע��ʱ���ʽ
application.Wait "18:23:00" '�ڴ�ʱ��������
newHour = Hour(Now())
newMinute = Minute(Now())
newSecond = Second(Now()) + 10
waitTime = TimeSerial(newHour, newMinute, newSecond)
application.Wait waitTime
If application.Wait(Now + TimeValue("0:00:10")) Then
    MsgBox "Time expired"
End If
'volatile��ʧ�Ժ�������ʱ�����仯����now��today��rand��
'Union�ϲ���Ԫ��
'Undo�������һ���û��������,��ʾ��������ں�ĵ�һ��
'SetDefaultChart��16����û�д˷�����
'Save �����޲���

End Sub
Function myFUN(cell As Range)
application.Volatile '���û����仰�����b5Ϊcell��ôb6�仯�Ͳ���Ӱ��������Ϊb6����myFUN�Ĳ�����
myFUN = cell.Value + cell.Offset(1, 0).Value

End Function
Sub ahl()
application.setdefalutchart FormatName:="Monthly Sales"
End Sub
