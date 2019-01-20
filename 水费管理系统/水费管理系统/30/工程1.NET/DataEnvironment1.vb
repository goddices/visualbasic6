Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.Compatibility

Module DataEnvironment_DataEnvironment1_Module
	Friend DataEnvironment1 As DataEnvironment_DataEnvironment1 = New DataEnvironment_DataEnvironment1()
End Module

Friend Class DataEnvironment_DataEnvironment1
	Inherits VB6.BaseDataEnvironment
	Public WithEvents Connection1 As ADODB.Connection
	Public WithEvents rs当日缴费情况 As ADODB.Recordset
	Private m_当日缴费情况 As ADODB.Command
	Public WithEvents rs缴费明细 As ADODB.Recordset
	Private m_缴费明细 As ADODB.Command
	Public Sub New()
		MyBase.New()
		Dim par As ADODB.Parameter
		
		
		Connection1 = New ADODB.Connection()
		Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\水费管理系统\water.mdb;Persist Security Info=False;"
		m_Connections.Add(Connection1, "Connection1")
		m_当日缴费情况 = New ADODB.Command()
		rs当日缴费情况 = New ADODB.Recordset()
		m_当日缴费情况.Name = "当日缴费情况"
		m_当日缴费情况.CommandText = "SELECT * FROM 水费管理 WHERE 缴费日期 = format(DATE (), ""yyyy-mm-dd"")"
		m_当日缴费情况.CommandType = ADODB.CommandTypeEnum.adCmdText
		rs当日缴费情况.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rs当日缴费情况.CursorType = ADODB.CursorTypeEnum.adOpenStatic
		rs当日缴费情况.LockType = ADODB.LockTypeEnum.adLockReadOnly
		rs当日缴费情况.Source = m_当日缴费情况
		m_Commands.Add(m_当日缴费情况, "当日缴费情况")
		m_Recordsets.Add(rs当日缴费情况, "当日缴费情况")
		m_缴费明细 = New ADODB.Command()
		rs缴费明细 = New ADODB.Recordset()
		m_缴费明细.Name = "缴费明细"
		m_缴费明细.CommandText = "SELECT * FROM 水费管理 "
		m_缴费明细.CommandType = ADODB.CommandTypeEnum.adCmdText
		rs缴费明细.CursorLocation = ADODB.CursorLocationEnum.adUseClient
		rs缴费明细.CursorType = ADODB.CursorTypeEnum.adOpenStatic
		rs缴费明细.LockType = ADODB.LockTypeEnum.adLockReadOnly
		rs缴费明细.Source = m_缴费明细
		m_Commands.Add(m_缴费明细, "缴费明细")
		m_Recordsets.Add(rs缴费明细, "缴费明细")
	End Sub
	Public Sub 当日缴费情况()
		If Connection1.State = ADODB.ObjectStateEnum.adStateClosed Then
			Connection1.Open()
		End If
		If rs当日缴费情况.State = ADODB.ObjectStateEnum.adStateOpen Then
			rs当日缴费情况.Close()
		End If
		m_当日缴费情况.ActiveConnection = Connection1
		rs当日缴费情况.Open()
	End Sub
	Public Sub 缴费明细()
		If Connection1.State = ADODB.ObjectStateEnum.adStateClosed Then
			Connection1.Open()
		End If
		If rs缴费明细.State = ADODB.ObjectStateEnum.adStateOpen Then
			rs缴费明细.Close()
		End If
		m_缴费明细.ActiveConnection = Connection1
		rs缴费明细.Open()
	End Sub
	'Download by http://www.codefans.net
End Class