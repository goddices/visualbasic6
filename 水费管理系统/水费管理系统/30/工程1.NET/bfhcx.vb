Option Strict Off
Option Explicit On
Friend Class Form32
	Inherits System.Windows.Forms.Form
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim SQL As Object
		If Option1.Checked = True Then
			'UPGRADE_WARNING: 未能解析对象 SQL 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
			SQL = "select * from 用户管理 where 总户号='" & Trim(Text1.Text) & "'"
			'UPGRADE_ISSUE: Data 属性 Data1.RecordSource 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			'UPGRADE_WARNING: 未能解析对象 SQL 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
			Data1.RecordSource = SQL
			Data1.Refresh()
			'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			If Data1.Recordset.EOF Then
				MsgBox("没有此总户号！",  , "提示")
			End If
		End If
		
		If Option2.Checked = True Then
			'UPGRADE_WARNING: 未能解析对象 SQL 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
			SQL = "select * from 用户管理 where 户名='" & Trim(Text1.Text) & "'"
			'UPGRADE_ISSUE: Data 属性 Data1.RecordSource 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			'UPGRADE_WARNING: 未能解析对象 SQL 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
			Data1.RecordSource = SQL
			Data1.Refresh()
			'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			If Data1.Recordset.EOF Then
				MsgBox("没有此户名！",  , "提示")
			End If
		End If
		
		If Option1.Checked = False And Option2.Checked = False Then
			MsgBox("请选择查询的项目后再进行查询！",  , "提示")
		End If
	End Sub
	
	Private Sub Form32_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Width = VB6.TwipsToPixelsX(8085)
		Me.Height = VB6.TwipsToPixelsY(6300)
		Me.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(MDIForm1.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2), VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(MDIForm1.Height) - VB6.PixelsToTwipsY(Me.Height)) / 4), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
	End Sub
End Class