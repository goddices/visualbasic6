Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.Compatibility

Module DataEnvironment_DataEnvironment1_Module
	Friend DataEnvironment1 As DataEnvironment_DataEnvironment1 = New DataEnvironment_DataEnvironment1()
End Module

Friend Class DataEnvironment_DataEnvironment1
	Inherits VB6.BaseDataEnvironment
	Public WithEvents Connection1 As ADODB.Connection
	Public WithEvents rs���սɷ���� As ADODB.Recordset
	Private m_���սɷ���� As ADODB.Command
	Public WithEvents rs�ɷ���ϸ As ADODB.Recordset
	Private m_�ɷ���ϸ As ADODB.Command
	Public Sub New()
		MyBase.New()
		Dim par As ADODB.Parameter
		
		
		Connection1 = New ADODB.Connection()
		Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\ˮ�ѹ���ϵͳ\water.mdb;Persist Security Info=False;"
		m_Connections.Add(Connection1, "Connection1")
		m_���սɷ���� = New ADODB.Command()
		rs���սɷ���� = New ADODB.Recordset()
		m_���սɷ����.Name = "���սɷ����"
		m_���սɷ����.CommandText = "SELECT * FROM ˮ�ѹ��� WHERE �ɷ����� = format(DATE (), ""yyyy-mm-dd"")"
		m_���սɷ����.CommandType = ADODB.CommandTypeEnum.adCmdText
		rs���սɷ����.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rs���սɷ����.CursorType = ADODB.CursorTypeEnum.adOpenStatic
		rs���սɷ����.LockType = ADODB.LockTypeEnum.adLockReadOnly
		rs���սɷ����.Source = m_���սɷ����
		m_Commands.Add(m_���սɷ����, "���սɷ����")
		m_Recordsets.Add(rs���սɷ����, "���սɷ����")
		m_�ɷ���ϸ = New ADODB.Command()
		rs�ɷ���ϸ = New ADODB.Recordset()
		m_�ɷ���ϸ.Name = "�ɷ���ϸ"
		m_�ɷ���ϸ.CommandText = "SELECT * FROM ˮ�ѹ��� "
		m_�ɷ���ϸ.CommandType = ADODB.CommandTypeEnum.adCmdText
		rs�ɷ���ϸ.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rs�ɷ���ϸ.CursorType = ADODB.CursorTypeEnum.adOpenStatic
		rs�ɷ���ϸ.LockType = ADODB.LockTypeEnum.adLockReadOnly
		rs�ɷ���ϸ.Source = m_�ɷ���ϸ
		m_Commands.Add(m_�ɷ���ϸ, "�ɷ���ϸ")
		m_Recordsets.Add(rs�ɷ���ϸ, "�ɷ���ϸ")
	End Sub
	Public Sub ���սɷ����()
		If Connection1.State = ADODB.ObjectStateEnum.adStateClosed Then
			Connection1.Open()
		End If
		If rs���սɷ����.State = ADODB.ObjectStateEnum.adStateOpen Then
			rs���սɷ����.Close()
		End If
		m_���սɷ����.ActiveConnection = Connection1
		rs���սɷ����.Open()
	End Sub
	Public Sub �ɷ���ϸ()
		If Connection1.State = ADODB.ObjectStateEnum.adStateClosed Then
			Connection1.Open()
		End If
		If rs�ɷ���ϸ.State = ADODB.ObjectStateEnum.adStateOpen Then
			rs�ɷ���ϸ.Close()
		End If
		m_�ɷ���ϸ.ActiveConnection = Connection1
		rs�ɷ���ϸ.Open()
	End Sub
	'Download by http://www.codefans.net
End Class