Option Strict Off
Option Explicit On
Friend Class MDIForm1
	Inherits System.Windows.Forms.Form
	Public Sub bjyf_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles bjyf.Click
		Form31.Show()
	End Sub
	'Download by http://down.liehuo.net
	Public Sub cxsf_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cxsf.Click
		Form22.Show()
	End Sub
	
	Public Sub DL_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles DL.Click
		Me.mmxg.Enabled = False
		Me.sfgl.Enabled = False
		Me.yfgl.Enabled = False
		Me.DYFW.Enabled = False
		Me.hjsz.Enabled = False
		Form11.Show()
	End Sub
	
	Public Sub dqjg_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dqjg.Click
		Form51.Show()
	End Sub
	
	Public Sub DRjfqh_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles DRjfqh.Click
		Dim DataReport1 As Object
		'UPGRADE_WARNING: 未能解析对象 DataReport1.Show 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
		DataReport1.Show()
	End Sub
	
	Private Sub HELP_Click()
		Dim Form13 As Object
		'UPGRADE_WARNING: 未能解析对象 Form13.Show 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
		Form13.Show()
	End Sub
	
	
	
	Public Sub jnsf_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles jnsf.Click
		Form21.Show()
	End Sub
	
	Public Sub llyf_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles llyf.Click
		Form32.Show()
	End Sub
	
	Public Sub mmxg_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mmxg.Click
		Form12.Show()
	End Sub
	
	Public Sub TC_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TC.Click
		Me.Close()
	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		StatusBar1.Panels(2).Text = "当前时间：" & TimeOfDay
	End Sub
End Class