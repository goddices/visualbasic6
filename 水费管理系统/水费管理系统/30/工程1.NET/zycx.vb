Option Strict Off
Option Explicit On
Friend Class Form22
	Inherits System.Windows.Forms.Form
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim SQL As Object
		If Option1(0).Checked = True Then
			'UPGRADE_WARNING: δ�ܽ������� SQL ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			SQL = "select * from ˮ�ѹ��� where �ܻ���='" & Trim(Text1.Text) & "'"
		Else
			If Option1(1).Checked = True Then
				'UPGRADE_WARNING: δ�ܽ������� SQL ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
				SQL = "select * from ˮ�ѹ��� where ����='" & Trim(Text1.Text) & "'"
			Else
				If Option1(2).Checked = True Then
					'UPGRADE_WARNING: δ�ܽ������� SQL ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
					SQL = "select * from ˮ�ѹ��� where �ɷ�����='" & VB6.Format(DTPicker1.Value, "yyyy-mm-dd") & "'"
				End If
			End If
		End If
		'UPGRADE_ISSUE: Data ���� Data1.RecordSource δ������ �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"��
		'UPGRADE_WARNING: δ�ܽ������� SQL ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
		Data1.RecordSource = SQL
		Data1.Refresh()
		'UPGRADE_ISSUE: Data ���� Data1.Recordset δ������ �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"��
		If Data1.Recordset.EOF Then
			MsgBox("û����Ҫ��ѯ�Ľ���ˮ�������",  , "��ʾ")
		End If
	End Sub
	'Download by http://down.liehuo.net
	Private Sub Form22_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Width = VB6.TwipsToPixelsX(10320)
		Me.Height = VB6.TwipsToPixelsY(6525)
		Me.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(MDIForm1.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2), VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(MDIForm1.Height) - VB6.PixelsToTwipsY(Me.Height)) / 4), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		DTPicker1.Value = Today
	End Sub
	
	'UPGRADE_WARNING: ��ʼ������ʱ���ܼ����¼� Option1.CheckedChanged�� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"��
	Private Sub Option1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option1.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = Option1.GetIndex(eventSender)
			Dim i As Object
			For i = 0 To 2
				If Option1(0).Checked = True Or Option1(1).Checked = True Then
					Text1.Visible = True
					DTPicker1.Visible = False
				Else
					If Option1(2).Checked = True Then
						Text1.Visible = False
						DTPicker1.Visible = True
					Else
						MsgBox("��ѡ���ѯ���",  , "��ʾ")
					End If
				End If
			Next i
			Command1.Visible = True
		End If
	End Sub
End Class