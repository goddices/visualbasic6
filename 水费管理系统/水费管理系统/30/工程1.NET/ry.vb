Option Strict Off
Option Explicit On
Friend Class Form21
	Inherits System.Windows.Forms.Form
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim i As Object
		MsgBox("����⣡",  , "��ʾ")
		Text5.Text = CStr(Val(Text5.Text) + Val(Text1(5).Text))
		'UPGRADE_ISSUE: Data ���� Data1.UpdateRecord δ������ �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"��
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
		'UPGRADE_WARNING: δ�ܽ������� SQL ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
		SQL = "select * from �û����� where �ܻ���='" & Trim(Text1(0).Text) & "'"
		'UPGRADE_ISSUE: Data ���� Data3.RecordSource δ������ �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"��
		'UPGRADE_WARNING: δ�ܽ������� SQL ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
		Data3.RecordSource = SQL
		Data3.Refresh()
		'UPGRADE_ISSUE: Data ���� Data3.Recordset δ������ �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"��
		If Data3.Recordset.EOF Then
			MsgBox("û�д��ܻ��ţ�����������[�ܻ���]��",  , "��ʾ")
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
			'UPGRADE_ISSUE: Data ���� Data1.Recordset δ������ �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"��
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
		'UPGRADE_WARNING: δ�ܽ������� Picture1.Picture ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
		Picture1.Picture = My.Computer.Clipboard.GetImage()
	End Sub
	
	'UPGRADE_WARNING: ��ʼ������ʱ���ܼ����¼� Text1.TextChanged�� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"��
	Private Sub Text1_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text1.TextChanged
		Dim Index As Short = Text1.GetIndex(eventSender)
		If Index = 3 Or Index = 4 Then
			Text1(5).Text = CStr(Val(Text1(3).Text) * Val(Text1(4).Text))
		End If
	End Sub
End Class