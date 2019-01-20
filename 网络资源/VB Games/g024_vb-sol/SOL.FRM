VERSION 5.00
Begin VB.Form Sol 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Solitaire"
   ClientHeight    =   5835
   ClientLeft      =   900
   ClientTop       =   2565
   ClientWidth     =   8865
   ForeColor       =   &H80000008&
   Icon            =   "SOL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   389
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   591
   Begin VB.PictureBox picInfoLine 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   -15
      Picture         =   "SOL.frx":030A
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   591
      TabIndex        =   14
      Top             =   5580
      Width           =   8895
   End
   Begin VB.PictureBox Hold 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   6
      Left            =   7545
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   13
      Top             =   1605
      Width           =   1065
   End
   Begin VB.PictureBox Hold 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   5
      Left            =   6315
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   12
      Top             =   1605
      Width           =   1065
   End
   Begin VB.PictureBox Hold 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   4
      Left            =   5085
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   11
      Top             =   1605
      Width           =   1065
   End
   Begin VB.PictureBox Hold 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   3
      Left            =   3855
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   10
      Top             =   1605
      Width           =   1065
   End
   Begin VB.PictureBox Hold 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   2
      Left            =   2625
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   9
      Top             =   1605
      Width           =   1065
   End
   Begin VB.PictureBox Hold 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   1
      Left            =   1395
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   8
      Top             =   1605
      Width           =   1065
   End
   Begin VB.PictureBox Hold 
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   0
      Left            =   165
      Picture         =   "SOL.frx":08AC
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   7
      Top             =   1605
      Width           =   1065
   End
   Begin VB.PictureBox Source 
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   2
      Left            =   225
      Picture         =   "SOL.frx":16AE
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   6
      Top             =   105
      Width           =   1065
   End
   Begin VB.PictureBox Source 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Index           =   1
      Left            =   195
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   5
      Top             =   90
      Width           =   1065
   End
   Begin VB.PictureBox Source 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Index           =   0
      Left            =   165
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   4
      Top             =   75
      Width           =   1065
   End
   Begin VB.PictureBox Home 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   3
      Left            =   7545
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   3
      Top             =   75
      Width           =   1065
   End
   Begin VB.PictureBox Home 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   2
      Left            =   6315
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   2
      Top             =   75
      Width           =   1065
   End
   Begin VB.PictureBox Home 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   1
      Left            =   5085
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   1
      Top             =   75
      Width           =   1065
   End
   Begin VB.PictureBox Home 
      BorderStyle     =   0  'None
      Height          =   1440
      Index           =   0
      Left            =   3855
      Picture         =   "SOL.frx":24B0
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   0
      Top             =   75
      Width           =   1065
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About VB Solitaire..."
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Sol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CalcPtOnLine(ByVal x0%, ByVal y0%, ByVal x1%, ByVal y1%, ByVal t!, X%, Y%)

  'USE:  Calc point on line determined by (x0,y0) and (x1,y1)
  'IN:   (x0,y0) = first endpoint of line
  '      (x1,y1) = second endpoint of line
  '      t = parameter between 0.0 and 1.0: t = 0.0 is (x0,y0),
  '      t = 1.0 is (x1,y1), t = 0.5 is midpoint, and so on
  'OUT:  (x,y) = calculated point on line for parameter t
  'NOTE: Add 0.5 for rounding to nearest pixel.

  X = Int(x0 + t * (x1 - x0) + 0.5)
  Y = Int(y0 + t * (y1 - y0) + 0.5)

End Sub

Private Sub Form_Load()

  Dim i%         'to traverse picture arrays

  'Initialize State.
  State = WAITING

  'Must use Twips to set initial form size, even though
  'DrawMode will be PIXELS for rest of program.
  Sol.Width = XWIN_SIZE * Screen.TwipsPerPixelX
  Sol.Height = YWIN_SIZE * Screen.TwipsPerPixelY
  Sol.DrawMode = PIXELS
  
  'Settings for source rectangles at top left.
  For i = 0 To 2
    Source(i).ScaleMode = PIXELS
    Source(i).BorderStyle = NONE
    Source(i).Left = SOURCE_LEFT + 2 * i
    Source(i).Top = SOURCE_TOP + i
    Source(i).Width = CARD_SIZEX
    Source(i).Height = CARD_SIZEY
  Next i
  
  'Settings for home rectangles across top right.
  For i = 0 To 3
    Home(i).ScaleMode = PIXELS
    Home(i).BorderStyle = NONE
    Home(i).Left = HOME_LEFT + i * HOME_OFFSET
    Home(i).Top = HOME_TOP
    Home(i).Width = CARD_SIZEX
    Home(i).Height = CARD_SIZEY
  Next i

  'Settings for hold rectangles in second row.
  For i = 0 To 6
    Hold(i).ScaleMode = PIXELS
    Hold(i).BorderStyle = NONE
    Hold(i).Left = HOLD_LEFT + i * HOLD_OFFSET
    Hold(i).Top = HOLD_TOP
    Hold(i).Width = CARD_SIZEX
    Hold(i).Height = CARD_SIZEY
  Next i

  'Set properties for info picture box at bottom of form.
  picInfoLine.BorderStyle = FIXED_SINGLE
  picInfoLine.ScaleMode = PIXELS
  picInfoLine.BackColor = RGB(255, 255, 255)
  picInfoLine.Left = -1
  picInfoLine.Top = YWIN_SIZE - INFO_LINE_OFFSET
  picInfoLine.Height = INFO_LINE_HEIGHT
  picInfoLine.Width = XWIN_SIZE

  'The three card bitmaps (Dots, Palms, AceHearts) are loaded
  'once each into a picture control at design time, then copied
  'here (at run time) into other controls where they appear.
  'This way each is stored only once in the FRX and EXE files.

  'Copy Dots card from 0th HomeRect to others.
  For i = 1 To 3
    Home(i).Picture = Home(0).Picture
  Next i

  'Copy Palms card from 2nd SourceRect to others.
  For i = 0 To 1
    Source(i).Picture = Source(2).Picture
  Next i

  'Copy Palms card from 2nd SourceRect to HoldRects 1 thru 6.
  For i = 1 To 6
    Hold(i).Picture = Source(2).Picture
  Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub Hold_DblClick(Index As Integer)

  'If double-click HoldRect, move to 0th Home position.
  If Index = 0 Then
    Hold(0).Left = Home(0).Left
    Hold(0).Top = Home(0).Top
  End If

End Sub

Private Sub Hold_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Index = 0 And State = WAITING Then
    State = DRAGGING
    'Set down point so can calculate deltas during drag.
    x0 = X: y0 = Y
    'Set dragged card's ZOrder so it's on top.
    Hold(0).ZOrder 0
  End If

End Sub

Private Sub Hold_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim dx, dy     'deltas from last position

  If Index = 0 And State = DRAGGING Then
    'Calculate deltas from last position.
    dx = X - x0: dy = Y - y0
    'Reposition card according to deltas.
    Hold(0).Left = Hold(0).Left + dx
    Hold(0).Top = Hold(0).Top + dy
  End If

End Sub

Private Sub Hold_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  Dim Overlap%      'holds number of overlapped card or -1 if none
  Dim xc%, yc%      'card location at button up
  Dim dx%, dy%      'deltas as positive values
  Dim np%           'number of pixels in line back to hold position
  Dim i%            'to calculate points along line
  Dim mx%, my%      'coords of points along line

  If Index = 0 And State = DRAGGING Then
    State = WAITING

    'Grab current card location.
    xc = Hold(0).Left:  yc = Hold(0).Top

    'Exit if final position same as original position (at HoldRect) to prevent
    '0-division later (no need to walk back since at original position).
    If (xc = HOLD_LEFT) And (yc = HOLD_TOP) Then
      Exit Sub
    End If

    'If HoldRect's current position overlaps with one of the HomeRects,
    'put it on the HomeRect and exit sub.
    Overlap = HomeHoldOverlap()
    If Overlap >= 0 Then
      Hold(0).Left = Home(Overlap).Left
      Hold(0).Top = Home(Overlap).Top
      Exit Sub
    End If

    'If get to here, no overlap: walk HoldRect back to original position.

    'Calculate deltas as positive values.
    dx = Abs(xc - HOLD_LEFT): dy = Abs(yc - HOLD_TOP)
    'Calculate number of pixels in line (largest delta).
    If (dx < dy) Then
      np = dy
    Else
      np = dx
    End If
    'Walk card back along the line to the original location.
    For i = 0 To np Step SKIP_FACTOR
      CalcPtOnLine xc, yc, HOLD_LEFT, HOLD_TOP, i / np, mx, my
      Hold(0).Left = mx:  Hold(0).Top = my
    Next i
    'One last time to make sure card ends up on HoldRect.
    Hold(0).Left = HOLD_LEFT:  Hold(0).Top = HOLD_TOP
  End If

End Sub

Private Function HomeHoldOverlap%()

  'USE:  Determine overlap of dragging HoldRect and HomeRects
  'RET:  Return number of overlapped HomeRect (0-3) or -1 if no overlap.

  Dim i%        'to traverse HomeRects
  Dim ToLeft%   'boolean saying if HoldRect to left of HomeRect
  Dim ToRight%  'boolean saying if HoldRect to right of HomeRect
  Dim Above%    'boolean saying if HoldRect above HomeRect
  Dim Below%    'boolean saying if HoldRect below HomeRect
  Dim ret%      'return value (0 to 3 or -1)

  For i = 0 To 3
    'Calculate booleans for relative position.
    ToLeft = (Hold(0).Left + CARD_SIZEX) < Home(i).Left
    ToRight = Hold(0).Left > (Home(i).Left + CARD_SIZEX)
    Above = (Hold(0).Top + CARD_SIZEY) < Home(i).Top
    Below = Hold(0).Top > (Home(i).Top + CARD_SIZEY)
    'Condition true iff overlap on i'th HomeRect
    If Not (ToLeft Or ToRight Or Above Or Below) Then
      'Set ret to i or i+1 depending on greatest overlap.
      If Hold(0).Left < (Home(i).Left + BOUNDARY) Then
        ret = i
      Else
        ret = i + 1
      End If
      'But if above calculated 4, set back to 3.
      If ret = 4 Then
        ret = 3
      End If
      HomeHoldOverlap = ret
      Exit Function
    End If
  Next i

  'If got to here, no overlap.
  HomeHoldOverlap = -1

End Function

Private Sub mnuAbout_Click()

  'Show About Box.
  frmAbout.Show MODAL

End Sub

Private Sub mnuExit_Click()
  End
End Sub

