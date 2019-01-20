VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   5475
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim conn As New ADODB.Connection
 Dim rs As New ADODB.Recordset

Dim cc As New Class1

Private Sub Command1_Click()
 MsgBox rs("选项1").Value
End Sub

Private Sub Form_Load()
 
 cc.SetConn conn
 
 
 cc.SetRecordSet rs, conn
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
conn.Close
Set conn = Nothing
Set cc = Nothing
End Sub
