VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form EditBook 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�༭�޸�ͼ��"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "EditBook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditBook.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditBook.frx":08A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditBook.frx":0D02
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditBook.frx":1162
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditBook.frx":15C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditBook.frx":1A22
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EditBook.frx":1E82
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "�ر�(&C)"
      Height          =   495
      Left            =   2760
      TabIndex        =   17
      Top             =   4200
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   240
      ScaleHeight     =   1575
      ScaleWidth      =   6375
      TabIndex        =   13
      Top             =   3360
      Width           =   6375
      Begin VB.CommandButton cmdOkCancel 
         BackColor       =   &H0000C000&
         Caption         =   "ȡ��"
         Height          =   495
         Index           =   1
         Left            =   3840
         Picture         =   "EditBook.frx":22E2
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdOkCancel 
         BackColor       =   &H0000C000&
         Caption         =   "ȷ��"
         Height          =   495
         Index           =   0
         Left            =   5040
         Picture         =   "EditBook.frx":25EC
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label labFlag 
         AutoSize        =   -1  'True
         Caption         =   "ȷʵҪ�޸ĵ�ǰ��¼��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   1920
         TabIndex        =   16
         Top             =   240
         Width           =   2640
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   360
      ScaleHeight     =   1335
      ScaleWidth      =   6375
      TabIndex        =   11
      Top             =   3360
      Width           =   6375
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   2565
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   4524
         ButtonWidth     =   1455
         ButtonHeight    =   1455
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ǰ"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ǰһ��"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��һ��"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               ImageIndex      =   7
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ͼ�������Ϣ"
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6375
      Begin VB.TextBox txtBookNum 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   2760
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtBookName 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1290
         Width           =   2535
      End
      Begin VB.TextBox txtBookChu 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtCost 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4800
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "EditBook.frx":2A2E
         Left            =   4800
         List            =   "EditBook.frx":2A30
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ͼ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   0
         Left            =   1560
         TabIndex        =   10
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��  ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��  ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   2
         Left            =   3960
         TabIndex        =   8
         Top             =   2280
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��  ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   3
         Left            =   3960
         TabIndex        =   7
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   4
         Left            =   360
         TabIndex        =   6
         Top             =   2280
         Width           =   780
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   600
         Picture         =   "EditBook.frx":2A32
         Top             =   480
         Width           =   480
      End
   End
End
Attribute VB_Name = "EditBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rst As Recordset
Dim Rec As Integer
Dim StrFlag As String
Dim NumFlag As Boolean
Dim db1 As Database
Dim rst1 As Recordset
Dim i As Integer
Private Sub cmdOkCancel_Click(Index As Integer)
Select Case Index
    Case 0
        If StrFlag = "�޸�" Then
            rst.Edit
            WriteIn
            rst.Update
            Disp
            Picture2.Visible = False
            Picture1.Visible = True
            SetTxt (False)
        ElseIf StrFlag = "ɾ��" Then
            rst.Delete
            rst.MovePrevious
            If rst.BOF Then rst.MoveNext
            Disp
            Picture2.Visible = False
            Picture1.Visible = True
        End If
    Case 1
        Disp
        Picture2.Visible = False
        Picture1.Visible = True
        SetTxt (False)
End Select
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set db = Workspaces(0).OpenDatabase(App.Path & "\DataBase\Data.mdb", False)
Set rst = db.OpenRecordset("Book", dbOpenTable)
rst.Index = "ͼ����"

Set db1 = Workspaces(0).OpenDatabase(App.Path & "\DataBase\Data.mdb", False)
Set rst1 = db1.OpenRecordset("Type", dbOpenTable)

Rec = rst.RecordCount
If Rec = 0 Then
    Toolbar1.Enabled = False
    SetTxt (False)
End If
SetTxt (False)
rst.MoveFirst
Disp
TypeAdd
Picture1.Visible = True
Picture2.Visible = False
NumFlag = False
End Sub
Private Sub Disp()
txtBookNum = rst.Fields("ͼ����") & vbNullString
txtBookName = rst.Fields("����") & vbNullString
txtCost = rst.Fields("�۸�") & Empty
txtBookChu = rst.Fields("������") & vbNullString
Combo1.Text = rst.Fields("���") & vbNullString
End Sub
Private Sub Kong()
txtBookNum = ""
txtBookName = ""
txtBookChu = ""
Combo1.Text = ""
End Sub
Private Sub SetTxt(Bool As Boolean)
txtBookNum.Enabled = Bool
txtCost.Enabled = Bool
txtBookName.Enabled = Bool
txtBookChu.Enabled = Bool
Combo1.Enabled = Bool
End Sub

Private Sub Form_Unload(Cancel As Integer)
rst.Close
rst1.Close
db1.Close
db.Close
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        rst.MoveFirst
        Disp
    Case 2
        rst.MovePrevious
        If rst.BOF Then
            rst.MoveNext
            Exit Sub
        End If
        Disp
    Case 3
        rst.MoveNext
        If rst.EOF Then
            rst.MovePrevious
            Exit Sub
        End If
        Disp
    Case 4
        rst.MoveLast
        Disp
    Case 10
        StrFlag = "�޸�"
        SetTxt (True)
        labFlag.Caption = "��ȷʵҪ�޸ĵ�ǰ��¼��"
        Picture1.Visible = False
        Picture2.Visible = True
    Case 11
        StrFlag = "ɾ��"
        labFlag.Caption = "��ȷʵҪɾ����ǰ��¼��"
        Picture1.Visible = False
        Picture2.Visible = True
    Case 12
        SearchNum.Show (1)
        If SearchFlag = True Then
            rst.Seek "=", BookBianHao
            If rst.NoMatch Then
                MsgBox "û�д�ͼ���ţ�", 0 + 48, "����ʧ��"
                Exit Sub
            End If
            Disp
            SearchFlag = False
        End If
End Select
End Sub
Private Sub WriteIn()
rst.Fields("ͼ����") = txtBookNum
rst.Fields("����") = txtBookName
rst.Fields("�۸�") = Val(txtCost)
rst.Fields("������") = txtBookChu
rst.Fields("���") = Combo1.Text
End Sub
Private Sub TypeAdd()
rst1.MoveLast
rst1.MoveFirst
For i = 1 To rst1.RecordCount
    Combo1.AddItem rst1.Fields("���")
    rst1.MoveNext
    If rst1.EOF Then Exit Sub
Next
End Sub
