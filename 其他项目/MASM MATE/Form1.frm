VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "masm mate  6.11"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   9105
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame flnk 
      Caption         =   "link 参数"
      Height          =   735
      Left            =   2160
      TabIndex        =   7
      Top             =   1200
      Width           =   6735
      Begin VB.TextBox Txt_lib 
         Height          =   375
         Left            =   4680
         TabIndex        =   11
         Text            =   "kernel32.lib"
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Opt_console 
         Caption         =   "console"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Opt_win 
         Caption         =   "windows"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "-entry:start"
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame fml 
      Caption         =   "ml 参数"
      Height          =   735
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   6735
      Begin VB.TextBox Txt_asm 
         Height          =   375
         Left            =   4680
         TabIndex        =   8
         Text            =   "asms\3-1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton asmfile 
         Caption         =   "..."
         Height          =   375
         Left            =   5880
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chk_nologo 
         Caption         =   "-nologo"
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chk_coff 
         Caption         =   "-coff"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chk_c 
         Caption         =   "-c"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdlnk 
      Caption         =   "执行 LINK 命令"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdml 
      Caption         =   "执行 ML 命令"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mlStr As String
Dim lnkStr As String

Private Sub Command1_Click()
    Shell "ml.exe -c -coff g\2-1.asm", vbNormalFocus
End Sub

Private Sub Command2_Click()
    Shell "link.exe -subsystem:console -entry:start -out:2-1.exe 2-1.obj kernel32.lib"
End Sub

Private Sub cmdml_Click()
    
End Sub

Private Sub Form_Load()
    mlStr = "ml.exe"
    lnkStr = "link.exe"
End Sub

Private Sub generate_ml_parameter()
    If chk_c.Value = 1 Then mlStr = mlStr & " -c"
    If chk_coff.Value = 1 Then mlStr = mlStr & " -coff"
    If Txt_asm = "" Then
        MsgBox "请选择源文件"
    Else
        mlStr = mlStr & " " & Txt_asm.Text
    End If
End Sub
