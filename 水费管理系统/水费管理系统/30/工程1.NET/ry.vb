Option Strict Off
Option Explicit On
Friend Class Form21
	Inherits System.Windows.Forms.Form
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim i As Object
		MsgBox("已入库！",  , "提示")
		Text5.Text = CStr(Val(Text5.Text) + Val(Text1(5).Text))
		'UPGRADE_ISSUE: Data 方法 Data1.UpdateRecord 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
		Data1.UpdateRecord()
		Text1(0).Focus()
		Text1(0).Text = ""
		For i = 1 To 9
			Label1(i).Visible = False
		Next i
		For i = 1 To 6
			Text1(i).Visible = False
		Next i
		Text8.Visible = False
		
	End Sub
	
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Dim i As Object
		Dim SQL As Object
		'UPGRADE_WARNING: 未能解析对象 SQL 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
		SQL = "select * from 用户管理 where 总户号='" & Trim(Text1(0).Text) & "'"
		'UPGRADE_ISSUE: Data 属性 Data3.RecordSource 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
		'UPGRADE_WARNING: 未能解析对象 SQL 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
		Data3.RecordSource = SQL
		Data3.Refresh()
		'UPGRADE_ISSUE: Data 属性 Data3.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
		If Data3.Recordset.EOF Then
			MsgBox("没有此总户号！请重新输入[总户号]！",  , "提示")
			Text1(0).Text = ""
			Text1(0).Focus()
		Else
			For i = 1 To 9
				Label1(i).Visible = True
			Next i
			For i = 1 To 6
				Text1(i).Visible = True
			Next i
			Text8.Visible = True
			'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			Data1.Recordset.AddNew()
			Text1(7).Text = Text6.Text
			Text1(1).Text = Text3.Text
			Text1(2).Text = Text4.Text
			Text1(4).Text = Text2.Text
			Text1(6).Text = DateString
			Text1(3).Focus()
			Text8.Text = Text7.Text
		End If
	End Sub
	
	
	
	Private Sub Form21_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Width = VB6.TwipsToPixelsX(7350)
		Me.Height = VB6.TwipsToPixelsY(7395)
		Me.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(MDIForm1.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2), VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(MDIForm1.Height) - VB6.PixelsToTwipsY(Me.Height)) / 4), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		Label5.Text = CStr(Today)
	End Sub
	
	Private Sub Picture1_Click()
		Dim Picture1 As Object
		'UPGRADE_WARNING: 未能解析对象 Picture1.Picture 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
		Picture1.Picture = My.Computer.Clipboard.GetImage()
	End Sub
	
	'UPGRADE_WARNING: 初始化窗体时可能激发事件 Text1.TextChanged。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"”
	Private Sub Text1_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text1.TextChanged
		Dim Index As Short = Text1.GetIndex(eventSender)
		If Index = 3 Or Index = 4 Then
			Text1(5).Text = CStr(Val(Text1(3).Text) * Val(Text1(4).Text))
		End If
	End Sub
End Class