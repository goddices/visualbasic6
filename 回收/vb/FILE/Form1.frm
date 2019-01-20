VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   7725
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "添加txt1.txt到12.jpg"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   3255
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3个文件"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Long
Dim ar() As Byte

Private Sub Command1_Click()

Dim data As Byte

Open App.Path & "\12.jpg" For Binary As #1
Get #1, LOF(1), data
Print "12.jpg"; data
Close #1


 
Open App.Path & "\3.jpg" For Binary As #1
Get #1, LOF(1), data
Print "3.jpg"; data
Close #1

 
Open App.Path & "\txt1.txt" For Binary As #1
Get #1, LOF(1), data
Print "txt1.txt"; data
Close #1

Print vbNewLine
End Sub

Private Sub Command2_Click()

 
Dim str As String

Dim r As Variant

Open App.Path & "\12.jpg" For Binary As #1
Open App.Path & "\txt1.txt" For Binary As #2
ReDim ar(LOF(2)) As Byte
 
 
 
Put #1, LOF(1) + 1, ";"
 
For i = 1 To LOF(2)
    Get #2, , ar(i)
    Put #1, LOF(1) + 1, ar(i)
Next

Close #1, #2

 'Text1.Text = Text1.Text & str
End Sub
 
 

Private Sub Command3_Click()
Dim str As String
Dim ar As Byte
Open App.Path & "\12.jpg" For Binary As #1
For i = LOF(1) To LOF(1) - 1000
    Get #1, i, ar
    str = str & ar
    'If ar(i) = ";" Then Exit For
    
Next
Text1.Text = str
Close #1
End Sub

Private Sub Form_Load()
i = -1
End Sub
