Option Strict Off
Option Explicit On
Friend Class Form11
	Inherits System.Windows.Forms.Form
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim SQL As Object
		MDIForm1.DL.Text = "注销"
		'UPGRADE_WARNING: 未能解析对象 SQL 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
		SQL = "select * from user where user ='" & Trim(Text1.Text) & "'"
		'UPGRADE_ISSUE: Data 属性 Data1.RecordSource 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
		'UPGRADE_WARNING: 未能解析对象 SQL 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
		Data1.RecordSource = SQL
		Data1.Refresh()
		'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
		If Data1.Recordset.EOF Then
			MsgBox("没有此用户！",  , "提示")
			'UPGRADE_ISSUE: Data 属性 Data1.RecordSource 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			Data1.RecordSource = "user"
			Data1.Refresh()
			
		Else
			MDIForm1.StatusBar1.Panels(1).Text = "用户名：" & Trim(Text1.Text)
			'UPGRADE_WARNING: 未能解析对象 yfm 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
			yfm = Trim(Text1.Text)
			If Trim(Text2.Text) = Trim(Text4.Text) Then
				'UPGRADE_WARNING: 未能解析对象 qxqx 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
				qxqx = Text5.Text
				'UPGRADE_WARNING: 未能解析对象 qxqx 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
				If qxqx = 2 Then
					MDIForm1.sfgl.Enabled = True
					MDIForm1.yfgl.Enabled = True
					MDIForm1.DYFW.Enabled = True
					MDIForm1.hjsz.Enabled = True
				Else
					'UPGRADE_WARNING: 未能解析对象 qxqx 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
					If qxqx = 3 Then
						MDIForm1.sfgl.Enabled = True
						MDIForm1.DYFW.Enabled = True
					Else
						MDIForm1.mmxg.Enabled = True
						MDIForm1.sfgl.Enabled = True
						MDIForm1.yfgl.Enabled = True
						MDIForm1.DYFW.Enabled = True
						MDIForm1.hjsz.Enabled = True
					End If
				End If
				Me.Close()
			Else
				MsgBox("密码错误！",  , "提示")
			End If
		End If
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Me.Close()
	End Sub
	
	Private Sub Form11_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Width = VB6.TwipsToPixelsX(4950)
		Me.Height = VB6.TwipsToPixelsY(3810)
		Me.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(MDIForm1.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2), VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(MDIForm1.Height) - VB6.PixelsToTwipsY(Me.Height)) / 4), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
	End Sub
End Class