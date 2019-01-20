VERSION 5.00
Begin VB.Form FormDEBUG 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   6345
   ClientLeft      =   15435
   ClientTop       =   4605
   ClientWidth     =   3075
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   3975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2040
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      ItemData        =   "FormDEBUG.frx":0000
      Left            =   240
      List            =   "FormDEBUG.frx":0013
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "FormDEBUG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form1.Hide
Me.Hide
Set Form1 = Nothing
End
End Sub

Private Sub List1_Click()
    Dim S As Long
    

    
    S = List1.ListIndex
    
    Select Case S
      Case 0
        Call Scene1 '玩具室
      Case 1
        Call Scene2 '机构
      Case 2
        Call Scene3 '机器人
       Case 3
        Call Scene4 '机器人2
        Case 4
        Call Scene5 '永动机
      Case Else '
        Call Scene1
    End Select
    
    
End Sub
