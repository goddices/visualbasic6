Option Strict Off
Option Explicit On
Friend Class Form12
	Inherits System.Windows.Forms.Form
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		If Command1.Text = "新增" Then
			Command1.Text = "确定"
			'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			Data1.Recordset.AddNew()
			Text2.Focus()
			Command2.Enabled = False
			Command3.Enabled = False
		Else
			Command1.Text = "新增"
			'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			Data1.Recordset.Update()
			'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			Data1.Recordset.MoveLast()
			Command2.Enabled = True
			Command3.Enabled = True
		End If
		'Download by http://down.liehuo.net
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
		Data1.Recordset.Delete()
		'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
		Data1.Recordset.MovePrevious()
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
		Data1.Recordset.Edit()
		'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
		Data1.Recordset.Update()
	End Sub
	
	Private Sub Form12_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Width = VB6.TwipsToPixelsX(4965)
		Me.Height = VB6.TwipsToPixelsY(4335)
		Me.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(MDIForm1.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2), VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(MDIForm1.Height) - VB6.PixelsToTwipsY(Me.Height)) / 4), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		
	End Sub
End Class