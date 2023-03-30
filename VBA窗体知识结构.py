
VBA窗体知识结构
	窗体属性
		外观：backcolor 背景色
			bordercolor 边框颜色
			borderstyle 边框类型
			caption 说明文本
			specialeffect 外观
		图片：picture 选择图片相当于filedilog
			picturealignment 图片对齐的位置
			picturesizemode 原有方式显示图片、在保持图片的原始比例的同时更改图片的大小，以及拉伸图片以填充空间
			picturetiling 平铺图片，ture就是把图片填充全部控件，图片重复填充空间
		位置：height 高度
			width 宽度
			left 左边距离
			top 顶端距离
			startupposition 窗体为与桌面，或office程序的那个位置
			specialeffect 图片为与窗体的那个位置有5个值
			whatsthisbutton
			whatsthishelp
			zoom

		其他：showmodel 打开窗体是否编辑工作表
			mouseicon 鼠标在窗体中的样式，需要mousepointer设置为99
			tag 窗体的附加信息，解释说明
			cycle 选0时，使用tab键，会遍历所有的该窗体范围内所有的控件；选1时，使用tab键，只会遍历某一窗体范围内的控件，tab不会跳出范围，到最后一个控件后会在回到该窗体范围的第一个控件
			drawbuffer 
			enabled
			helpcontextid
			keepscrollbarsvisible
			righttoleft
			scrollbars
			scrollheight
			scrollleft
			scrolltop
			scrollwidth
			Controls(控件名称).text=""

	窗体事件
		加载事件 initialize()
			msgbox "加载"
			me.caption=now
			me.tag=now
		激活事件 activate()
			msgbox "窗体激活了！"
			msgbox me.tag
		失去焦点事件 deactivate()
			msgbox "离开了！"
		关闭按钮事件 queryclose(cancel参数，closemode参数均为整数)
			msgbox closemode 是0
			if closemode=0 then cancel=true ‘禁用关闭
		离开后的事件 terminate()
			msgbox "欢迎下次光临！"
		单击事件 click()
			msgbox "你单击了窗体"
		双击事件 dblclick(cancel参数：控件返回值，值是0或-1)
			双击和单击事件不能同时存在
		Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
		AddControl(ByVal Control As MSForms.Control)
		BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal State As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
		BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
		KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
		KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
		KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
		Layout()
		MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
		MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
		MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
			鼠标移动到窗体事件
			
		RemoveControl(ByVal Control As MSForms.Control)
		Resize()
		Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
		Zoom(Percent As Integer)
		



