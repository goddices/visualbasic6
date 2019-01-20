Option Strict Off
Option Explicit On
Friend Class Form31
	Inherits System.Windows.Forms.Form
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim Index As Short = Command1.GetIndex(eventSender)
		If Index = 0 Then
			'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			Data1.Recordset.MoveFirst()
			Command1(2).Enabled = False
			Command1(1).Enabled = True
		End If
		If Index = 1 Then
			Command1(2).Enabled = True
			'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			Data1.Recordset.MoveNext()
			'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			If Data1.Recordset.EOF Then
				'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
				Data1.Recordset.MoveLast()
				Command1(1).Enabled = False
			End If
		End If
		If Index = 2 Then
			Command1(1).Enabled = True
			'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			Data1.Recordset.MovePrevious()
			'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			If Data1.Recordset.BOF Then
				'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
				Data1.Recordset.MoveFirst()
				Command1(2).Enabled = False
			End If
		End If
		If Index = 3 Then
			'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			Data1.Recordset.MoveLast()
			Command1(1).Enabled = False
			Command1(2).Enabled = True
		End If
		If Index = 4 Then
			'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			Data1.Recordset.AddNew()
			'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			Text1(0).Text = CStr(Data1.Recordset.RecordCount + 1)
			Text1(1).Focus()
			Command1(4).Enabled = False
			Command1(5).Enabled = False
			Command1(6).Enabled = False
		End If
		If Index = 5 Then
			'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			Data1.Recordset.Delete()
			'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			Data1.Recordset.MoveNext()
			'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			If Data1.Recordset.EOF Then
				'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
				Data1.Recordset.MoveLast()
				Command1(1).Enabled = False
			End If
		End If
		If Index = 6 Then
			'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			Data1.Recordset.Edit()
			Command1(4).Enabled = False
			Command1(5).Enabled = False
			Command1(6).Enabled = False
		End If
		If Index = 7 Then
			'UPGRADE_ISSUE: Data 方法 Data1.UpdateRecord 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			Data1.UpdateRecord()
			'UPGRADE_ISSUE: Data 属性 Data1.Recordset 未升级。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"”
			Data1.Recordset.MoveLast()
			Command1(1).Enabled = False
			Command1(2).Enabled = True
			Command1(4).Enabled = True
			Command1(5).Enabled = True
			Command1(6).Enabled = True
		End If
	End Sub
	'Download by http://down.liehuo.net
	Private Sub Form31_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Width = VB6.TwipsToPixelsX(7095)
		Me.Height = VB6.TwipsToPixelsY(5340)
		Me.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(MDIForm1.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2), VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(MDIForm1.Height) - VB6.PixelsToTwipsY(Me.Height)) / 4), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
	End Sub
End Class