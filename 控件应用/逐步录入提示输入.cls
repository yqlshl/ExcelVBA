VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'=============列表框单击事件============
Private Sub ListBox1_Click()
'单击后，把值赋予活动单元格
ActiveCell.Value = ListBox1.Value
'之后隐藏列表框和文本框
ListBox1.Visible = False
TextBox1.Visible = False
End Sub
'==============文本框输入内容事件===============
Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Dim arr
ListBox1.Visible = True
'把sheet5中a列内容写入数组中
arr = Sheet5.Range("a1", Sheet5.Cells(Rows.Count, 1).End(xlUp))
'清理列表框中的内容，运行前必须清理固定模式
ListBox1.Clear
'进行循环，如果在文本框中输入内容，则列表框中出现相应的选项
For Each ar In arr
' instr函数第二参数是被查找的文本，第三参数为需要查找的文本
If InStr(1, ar, TextBox1.Value) Then
    ListBox1.AddItem ar
End If
Next ar
'设置文本框的值赋予活动单元格的值
ActiveCell.Value = TextBox1.Value
End Sub
'===================工作表中单元格选择事件===================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'设置文本框和单元格的位置
If Target.Column = 1 Then
    TextBox1.Visible = True
    ListBox1.Visible = False
    TextBox1.Height = ActiveCell.Height
    TextBox1.Width = ActiveCell.Width
    TextBox1.Top = ActiveCell.Top
    TextBox1.Left = ActiveCell.Left
    ListBox1.Top = ActiveCell.Top
    ListBox1.Left = ActiveCell(1, 2).Left
    '激活文本框，就可以点击的时候直接输入
    TextBox1.Activate
    '每次都清空文本框的内容
    TextBox1.Value = ""
    '每次都清空列表框内容
    ListBox1.Clear
Else
    '不是选择a列，则全部都是隐藏状态
    ListBox1.Visible = False
    TextBox1.Visible = False
End If
End Sub
