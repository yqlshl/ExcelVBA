Attribute VB_Name = "��������Ӧ�ó���"

Sub ��������Ӧ�ó���() '������excelVBA�����У�wordVBA��û��sendkeys����

Dim labelshop
Rem vbNormalFocus��ʾ������ʾ��vbMaximizedFocus���ģʽ��vbMinimizedFocus��С��
labelshop = Shell("Explorer.exe ����ĵ�ַ", vbNormalFocus)

Rem excelVBA����wait������wordVBAֻ����sleep api
Application.wait (Now + timevale("00:00:05"))


Application.SendKeys "������" ���� ����


Application.SendKeys "����һ���ַ�,����һƪ����.����ַ�������,�������ñ���"

Application.SendKeys "^P"                '������ӡ����
Application.SendKeys "5"                  '���ô�ӡ�ķ���
Application.SendKeys "~"                '�س���ʼ��ӡ

Application.SendKeys ��������:
����
        ����

Backspace
{BACKSPACE} �� {BS}

Break
{BREAK}

Caps Lock
{CAPSLOCK}

���
{CLEAR}

Delete �� Del
{DELETE} �� {DEL}

���¼�
{DOWN}

����
{END}

Enter (����С����)
{ENTER}

Enter
~�����η���

Esc
{ESCAPE} �� {ESC}

����
{HELP}

��ҳ
{HOME}

Ins
{INSERT}

�����
{LEFT}

Num Lock
{NUMLOCK}

PageDown
{PGDN}

PageUp
{PGUP}

Return
{RETURN}

���Ҽ�
{RIGHT}

Scroll Lock
{SCROLLLOCK}

Tab
{TAB}

���ϼ�
{UP}

F1 �� F15
{F1} �� {F15}
</table>
����ָ���� Shift ��/�� Ctrl ��/�� Alt ���ʹ�õļ�����Ҫָ�������������ʹ�õļ�����ʹ���±�
[Table]
Ҫ��ϵļ�
�ڼ�����֮ǰ���

Shift
+���Ӻţ�

Ctrl
^��������ţ�

Alt
%���ٷֺţ�



End Sub

Sub �ر�()
  Set msoft = GetObject("winmgmts:").execquery("select * from win32_process where name like '%.exe'")
    For Each S In msoft       '��s.name�����г��������е�exe�ļ�
      If S.Name = "LabelShop.exe" Then
         S.Terminate
      End If
    Next
End Sub






