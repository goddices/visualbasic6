VERSION 5.00
Begin VB.Form frmBlt 
   AutoRedraw      =   -1  'True
   Caption         =   "Menace"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   2055
   ScaleWidth      =   3030
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1275
      Top             =   1425
   End
End
Attribute VB_Name = "frmBlt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Transparent Blit
Option Compare Text
Option Explicit

Dim u As Long

Dim blnend As Boolean

' Win32
Const IMAGE_BITMAP = 0
Const LR_LOADFROMFILE = &H10
Const LR_CREATEDIBSECTION = &H2000
Const SRCCOPY = &HCC0020
Private Type BITMAP
        bmType          As Long
        bmWidth         As Long
        bmHeight        As Long
        bmWidthBytes    As Long
        bmPlanes        As Integer
        bmBitsPixel     As Integer
        bmBits          As Long
End Type

' GDI32
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
' USER32
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Const ResolutionX = 640     ' Width for the display mode
Const ResolutionY = 480     ' Height for the display mode

Dim dd As DirectDraw2               ' DirectDraw object
Dim ddsdFront As DDSURFACEDESC      ' Front surface description
Dim ddsFront As DirectDrawSurface2  ' Front buffer
Dim ddsBack As DirectDrawSurface2   ' Back buffer

Dim aDDS As DirectDrawSurface2      ' Images to blit
Dim tDDS As DirectDrawSurface2      ' tiles to blit

Dim ddCaps As DDSCAPS               ' Capabilities for search

Dim fx As DDBLTFX

'hold the sprites
Dim spnx%(40), spny%(40), spnw%(40), spnh%(40), spnox%(40), spnoy%(40)

Dim mode%      'mode% is the current behaviour
Dim anim%      'amount through the given behaviour animation
Dim animshift% 'flag to indicate if blocks should be pushed during anim

'dim sprite behaviour guff
Dim bname$(30)  'name of behaviour (arbitrary 30 behaviour limit)
Dim bcells%(30) 'number of cells in the behaviour
Dim bchar%(30, 30) '30 behaviours, with max 30 cells in the anim
Dim bxo%(30, 30)   'x offset
Dim byo%(30, 30)   'y offset

'now the block array x and y in pixels
Dim blockx%(30)    'up to 30 blocks on a map
Dim blocky%(30)
Dim blockcell%(30)
Dim blockmode%(30) '0=none, 1=left, 2=right, 3=fall
Dim blockcount%    'number of blocks on this level
            
Dim level%

'hold the map
Dim map(40, 6) As Integer '40 wide, 6 high
Dim mapl%, mapv% 'left margin



' Loads a bitmap in a DirectDraw surface
Private Function CreateDDSFromBitmap(dd As DirectDraw2, ByVal strFile As String) As DirectDrawSurface2
    Dim hbm As Long                 ' Handle on bitmap
    Dim bm As BITMAP                ' Bitmap header
    Dim ddsd As DDSURFACEDESC       ' Surface description
    Dim dds As DirectDrawSurface2   ' Created surface
    Dim hdcImage As Long            ' Handle on image
    Dim mhdc As Long                ' Handle on surface context
    Dim clr As Long                 'hold the colour top left to be made transparent
    ' Load bitmap
    hbm = LoadImage(ByVal 0&, strFile, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
    ' Get bitmap info
    GetObject hbm, Len(bm), bm
    ' Fill surface description
    With ddsd
        .dwSize = Len(ddsd)
        .dwFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .DDSCAPS.dwCaps = DDSCAPS_OFFSCREENPLAIN
        .dwWidth = bm.bmWidth
        .dwHeight = bm.bmHeight
    End With
    ' Create surface
    dd.CreateSurface ddsd, dds, Nothing
    ' Create memory device
    hdcImage = CreateCompatibleDC(ByVal 0&)
    ' Select the bitmap in this memory device
    SelectObject hdcImage, hbm
    ' Restore the surface
    dds.Restore
    ' Get the surface's DC
    dds.GetDC mhdc
    ' Copy from the memory device to the DirectDrawSurface
    StretchBlt mhdc, 0, 0, ddsd.dwWidth, ddsd.dwHeight, hdcImage, 0, 0, bm.bmWidth, bm.bmHeight, SRCCOPY
    'get the top left colour
    clr = GetPixel(mhdc, 0, 0)
    ' Release the surface's DC
    dds.ReleaseDC mhdc
    ' Release the memory device and the bitmap
    DeleteDC hdcImage
    DeleteObject hbm
    
    'make surface transparent
    Dim mhddck As DDCOLORKEY
    mhddck.dwColorSpaceLowValue = clr 'really works only for 24 bit colour
    mhddck.dwColorSpaceHighValue = clr 'but as sprites have black is all 0 at any rate
    dds.SetColorKey DDCKEY_SRCBLT, mhddck
    
    ' Returns the new surface
    Set CreateDDSFromBitmap = dds
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If mode% = 1 Or mode% = 2 Then 'walking left or right
    'shift = 1 then push, = 2 then jump
    Select Case KeyCode
        Case vbKeyR
            'restart  the level
            loadlevel level%
        Case vbKeyControl
            Select Case Shift
            Case 2
                If mode% = 1 Then mode% = 5 Else mode% = 7 'jump
                anim% = 0 'amount through the jump
                animshift% = 0 'don't shift blocks
            Case 3 'jump & shift
                If mode% = 1 Then mode% = 5 Else mode% = 7 'jump
                anim% = 0 'amount through the jump
                animshift% = 1 'shift blocks
            End Select
        Case vbKeyEscape
            blnend = True
        Case vbKeyLeft
            Select Case Shift
            Case 0 'walk normally
                mode% = 1 'walkleft
                anim% = anim% + 1 'walk
                If anim% > 7 Then anim% = 1
                If map((mapl% + 294) \ 60, mapv% \ 60) = 0 Then
                    mapl% = mapl% - 6
                End If
                If mapl% < -300 Then mapl% = -300
                If mapl% Mod 60 = 0 Then
                    If map((mapl% + 300) \ 60, (mapv% + 10) \ 60) = 0 Then
                        mode% = 3 'fallleft
                    End If
                End If
            Case 1 'shift box
                mode% = 9 'shift left
                anim% = 0 'amount through the jump
                animshift% = 1 'shift blocks
            End Select
        Case vbKeyRight
            Select Case Shift
            Case 0 'walk normally
                mode% = 2 'walkright
                anim% = anim% + 1 'walk
                If anim% > 7 Then anim% = 1
                If map((mapl% + 365) \ 60, mapv% \ 60) = 0 Then
                    mapl% = mapl% + 6
                End If
                If mapl% > 2400 Then mapl% = 2400
                If mapl% Mod 60 = 0 Then
                    If map((mapl% + 300) \ 60, (mapv% + 10) \ 60) = 0 Then
                        mode% = 4 'fall
                    End If
                End If
            Case 1 'shift box
                mode% = 10 'shift right
                anim% = 0 'amount through the jump
                animshift% = 1 'shift blocks
            End Select
    End Select
End If
End Sub

Private Sub Form_Load()
    Dim a%, g$, bi%
    mapv% = 0
    mode% = 3 'fallleft
    anim% = 1 'start of animation
    
    ShowCursor 0
    ' Create the DirectDraw object
    DirectDrawCreate ByVal 0&, dd, Nothing
    ' This app is full screen and will change the display mode
    dd.SetCooperativeLevel Me.hwnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN
    ' Set the display mode
    dd.SetDisplayMode ResolutionX, ResolutionY, 8, 0, 0
    ' Load images (in a real app don't load the surrounding empty space !)
    Set aDDS = CreateDDSFromBitmap(dd, App.Path & "\menace.BMP")
    Set tDDS = CreateDDSFromBitmap(dd, App.Path & "\tiles.BMP")
    
    ' Fill front buffer description structure...
    With ddsdFront
        ' Structure size
        .dwSize = Len(ddsdFront)
        ' Use DDSD_CAPS and BackBufferCount
        .dwFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
        ' Primary, flipable surface
        .DDSCAPS.dwCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX Or DDSCAPS_SYSTEMMEMORY
        ' One back buffer (you can try 2)
        .dwBackBufferCount = 1
    End With
    ' Create front buffer
    dd.CreateSurface ddsdFront, ddsFront, Nothing
    
    ' Retrieve the back buffer object
    ddCaps.dwCaps = DDSCAPS_BACKBUFFER
    ddsFront.GetAttachedSurface ddCaps, ddsBack
    
  
    'load up the sprites
    Open App.Path & "\menace.spr" For Random As #1 Len = 2
    For a% = 0 To 39
        Get #1, a% * 6 + 1, spnox%(a% + 1)
        Get #1, a% * 6 + 2, spnoy%(a% + 1)
        Get #1, a% * 6 + 3, spnx%(a% + 1)
        Get #1, a% * 6 + 4, spny%(a% + 1)
        Get #1, a% * 6 + 5, spnw%(a% + 1)
        Get #1, a% * 6 + 6, spnh%(a% + 1)
    Next
    Close #1
    
    
    'load up sprite behaviours
    Open App.Path & "\spritebe.txt" For Input As #1
    Do
        Line Input #1, g$
        If left$(g$, 1) = "*" Then
            'record follows
            Input #1, bi%
            Input #1, bname$(bi%), bcells%(bi%)
            For a% = 1 To bcells%(bi%)
                Input #1, bchar%(bi%, a%), bxo%(bi%, a%), byo%(bi%, a%)
            Next
        End If
    Loop Until EOF(1)
    Close #1
    
    'load the map level
    level% = 1
    loadlevel level%
   
    
    Timer1.Enabled = -1
End Sub

' Draw next frame
Private Sub DrawNextFrame()
    Dim a%, b%, xx%, ofx%, sp%, cbx%, cby%
    Dim t As RECT
    'On Error Resume Next
    ' Clear the back buffer
    With fx
        .dwSize = Len(fx)
        .dwFillColor = RGB(0, 0, 0)
    End With
    t.top = 0
    t.left = 0
    t.bottom = ResolutionY
    t.Right = ResolutionX
    ddsBack.Blt t, Nothing, t, DDBLT_COLORFILL, fx
      
    'draw the map
    ofx% = mapl% \ 60
    For a% = 0 To 11
        If a% + ofx% >= 0 And a% + ofx% <= 39 Then
            For b% = 6 To 1 Step -1
                drawblock map(a% + ofx%, b% - 1), a% * 60 - (mapl% Mod 60), b% * 60
            Next
        End If
    Next
    
    'move moveable blocks
    For a% = 1 To 30
        If blockmode%(a%) > 0 Then 'something to do
            'yeah I know, cheap and nasty,
            'moving blocks go on top of the map... oh well
            drawblock blockcell%(a%), (blockx%(a%) - mapl%) - 300, blocky%(a%)
            Select Case blockmode%(a%)
                Case 1, 2 'left,right
                    If blockmode%(a%) = 1 Then blockx%(a%) = blockx%(a%) - 6 Else blockx%(a%) = blockx%(a%) + 6
                    If blockx%(a%) Mod 60 = 0 Then
                        'now check above old spot!
                        If blockmode%(a%) = 1 Then cbx% = (blockx%(a%) - 235) \ 60 Else cbx% = (blockx%(a%) - 355) \ 60
                        cby% = (blocky%(a%) - 10) \ 60
                        blockmode%(a%) = 3 'fall
                        checkabove cbx%, cby%
                    End If
                Case 3 'fall till hit something
                    If map((blockx%(a%) - 300) \ 60, (blocky%(a%) + 10) \ 60) = 0 Then
                        'keep falling
                        blocky%(a%) = blocky%(a%) + 10
                    Else
                        'hit ground so add to map
                        map((blockx%(a%) - 300) \ 60, (blocky%(a%) - 5) \ 60) = blockcell%(a%)
                        blockmode%(a%) = 0 'stop falling
                        checkforgroup
                    End If
            End Select
        End If
    Next
    
    
    'place the character
    Select Case mode%
        Case 1, 2
        Case 3, 4 'fall
            anim% = anim% + 1
            If anim% > bcells%(mode%) Then anim% = 1
            If map((mapl% + 300) \ 60, (mapv% + 10) \ 60) = 0 And map((mapl% + 355) \ 60, (mapv% + 10) \ 60) = 0 Then
                mapv% = mapv% + 10
            Else
                If mode% = 3 Then mode% = 6 Else mode% = 8 'land
            End If
            If mapv% > 350 Then
                mapv% = 350
                If mode% = 3 Then mode% = 6 Else mode% = 8 'land
            End If
        Case 5, 7 'jump left or right
            anim% = anim% + 1
            If anim% > bcells%(mode%) Then
                If mode% = 5 Then mode% = 3 Else mode% = 4
                anim% = 1
            Else
                If anim% <= 15 Then 'on the way up
                    If mode% = 5 Then 'left
                        If map((mapl% + 300) \ 60, (mapv% - 60) \ 60) = 0 And map((mapl% + 355) \ 60, (mapv% - 60) \ 60) = 0 Then
                            mapv% = mapv% + byo%(mode%, anim%)
                            If mapv% < 1 Then mapv% = 1
                        Else
                            anim% = 15
                        End If
                        If map((mapl% + 300 + bxo%(mode%, anim%)) \ 60, mapv% \ 60) = 0 And map((mapl% + 300 + bxo%(mode%, anim%)) \ 60, (mapv% - 50) \ 60) = 0 Then
                            mapl% = mapl% + bxo%(mode%, anim%)
                        Else
                            mode% = 3: anim% = 1
                            If animshift% Then startmove 1
                        End If
                    Else 'right
                        If map((mapl% + 355) \ 60, (mapv% - 60) \ 60) = 0 And map((mapl% + 300) \ 60, (mapv% - 60) \ 60) = 0 Then
                            mapv% = mapv% + byo%(mode%, anim%)
                            If mapv% < 1 Then mapv% = 1
                        Else
                            anim% = 15
                        End If
                        If map((mapl% + 355 + bxo%(mode%, anim%)) \ 60, mapv% \ 60) = 0 And map((mapl% + 355 + bxo%(mode%, anim%)) \ 60, (mapv% - 50) \ 60) = 0 Then
                            mapl% = mapl% + bxo%(mode%, anim%)
                            If mapv% < 1 Then mapv% = 1
                        Else
                            mode% = 4: anim% = 1
                            If animshift% Then startmove 2
                        End If
                    End If
                    
                Else 'on the way down
                    If mode% = 5 Then 'left
                        If map((mapl% + 300) \ 60, (mapv% + 15) \ 60) = 0 And map((mapl% + 355) \ 60, (mapv% + 15) \ 60) = 0 Then
                            mapv% = mapv% + byo%(mode%, anim%)
                            If mapv% < 1 Then mapv% = 1
                        Else
                            mode% = 6: anim% = 1 'land
                        End If
                        If map((mapl% + 300 + bxo%(mode%, anim%)) \ 60, mapv% \ 60) = 0 And map((mapl% + 300 + bxo%(mode%, anim%)) \ 60, (mapv% - 50) \ 60) = 0 Then
                            mapl% = mapl% + bxo%(mode%, anim%)
                        Else
                            mode% = 3: anim% = 1 'fall
                            If animshift% Then startmove 1
                        End If
                    Else 'right
                        If map((mapl% + 355) \ 60, (mapv% + 15) \ 60) = 0 And map((mapl% + 300) \ 60, (mapv% + 15) \ 60) = 0 Then
                            mapv% = mapv% + byo%(mode%, anim%)
                            If mapv% < 1 Then mapv% = 1
                        Else
                            mode% = 4: anim% = 1 'fall
                        End If
                        If map((mapl% + 355 + bxo%(mode%, anim%)) \ 60, mapv% \ 60) = 0 And map((mapl% + 355 + bxo%(mode%, anim%)) \ 60, (mapv% - 50) \ 60) = 0 Then
                            mapl% = mapl% + bxo%(mode%, anim%)
                        Else
                            mode% = 4: anim% = 1 'fall
                            If animshift% Then startmove 2
                        End If
                    End If
                End If
            End If
        Case 9, 10 'shift blocks
            If anim% = 1 Then 'start pushing
                'make blocks moveable
                If mode% = 9 Then startmove 1 Else startmove 2
            End If
            anim% = anim% + 1
            If anim% > bcells%(mode%) Then
                anim% = 1
                If mode% = 9 Then mode% = 1 Else mode% = 2
            End If
        Case Else
            anim% = anim% + 1
            If anim% > bcells%(mode%) Then
                Select Case mode%
                    Case 6: mode% = 1 'land now walk left
                    Case 8: mode% = 2 'land now walk right
                    Case Else: mode% = 1
                End Select
                anim% = 1
            End If
    End Select
    
        
    sp% = bchar%(mode%, anim%)
    If sp% = 0 Then sp% = 1
    
    t.top = spny%(sp%)
    t.left = spnx%(sp%)
    t.bottom = spny%(sp%) + spnh%(sp%)
    t.Right = spnx%(sp%) + spnw%(sp%)
       
   ddsBack.BltFast 295 + spnox%(sp%), mapv% + spnoy%(sp%) + 10, aDDS, t, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    
    'If Err.Number <> 0 Then
        ' Just in case
    '    Stop
    'End If
    ' Flip the buffers
    Do
        ddsFront.Flip Nothing, 0
        If Err.Number = DDERR_SURFACELOST Then ddsFront.Restore
    Loop Until Err.Number = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)

End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = 0

    'Render loop
    While Not blnend
        DrawNextFrame
        u = DoEvents
    Wend
    
    'clean up
    Set aDDS = Nothing
    Set tDDS = Nothing
    dd.FlipToGDISurface
    dd.RestoreDisplayMode
    dd.SetCooperativeLevel 0, DDSCL_NORMAL
    Set ddsBack = Nothing
    Set ddsFront = Nothing
    Set dd = Nothing
    ShowCursor 1
    
    Unload frmBlt

End Sub
Sub loadlevel(mapnum%)
Dim a%, b%, nfile%
mapv% = 0
mapl% = 0
mode% = 3
Open App.Path & "\map.dat" For Random As #1 Len = 2
nfile% = LOF(1) / 2
blockcount% = 0
For a% = 0 To 39
    For b% = 0 To 5
        Get #1, 1 + (a% + (b% * 40)) + (mapnum% - 1) * 2500, map(a%, b%)
        If map(a%, b%) > 3 Then blockcount% = blockcount% + 1
    Next
Next

Close #1
For a% = 1 To 30
    blockmode%(a%) = 0
Next
End Sub

Sub drawblock(cl%, ByVal xx%, ByVal yy%)
Dim a%, b%
Dim t As RECT
    
Select Case cl%
    Case 0
        t.top = -99
    Case 1:
        t.top = 0
        t.left = 0
        t.bottom = 75
        t.Right = 80
    Case 2:
        t.top = 0
        t.left = 80
        t.bottom = 75
        t.Right = 160
    Case 3:
        t.top = 0
        t.left = 160
        t.bottom = 75
        t.Right = 240
    Case 4:
        t.top = 75
        t.left = 0
        t.bottom = 150
        t.Right = 80
    Case 5:
        t.top = 75
        t.left = 80
        t.bottom = 150
        t.Right = 160
    Case 6:
        t.top = 75
        t.left = 160
        t.bottom = 150
        t.Right = 240
End Select
If t.top >= 0 Then
    If xx% < 0 Then
        t.left = t.left + Abs(xx%)
        xx% = 0
    End If
    If xx% > 560 Then
        t.Right = t.Right - (xx% - 560)
    End If
    If t.Right > t.left Then ddsBack.BltFast xx%, yy%, tDDS, t, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If
End Sub
Sub startmove(initdir%)
Dim cbx%, cby%, a%
If initdir% = 1 Then cbx% = (mapl% + 290) \ 60 Else cbx% = (mapl% + 369) \ 60
cby% = (mapv% - 30) \ 60
If map(cbx%, cby%) > 3 Then 'moveable block
    If (initdir% = 1 And map%(cbx% - 1, cby%) = 0) Or (initdir% = 2 And map%(cbx% + 1, cby%) = 0) Then
        For a% = 1 To 30
            If blockmode%(a%) = 0 Then 'spare block
                blockx%(a%) = cbx% * 60 + 300
                blocky%(a%) = cby% * 60 + 60
                blockcell%(a%) = map(cbx%, cby%)
                If initdir% = 1 Then blockmode%(a%) = 1 Else blockmode%(a%) = 2
                Exit For
            End If
        Next
        map(cbx%, cby%) = 0 'remove block of map, as now is in motion
        'when the block stops, it will be readded to the map
    End If
End If

End Sub

Sub checkforgroup()
Dim a%, b%, cc%, ct%, flag%
For a% = 0 To 39
    ct% = 0: cc% = 0
    For b% = 0 To 5
        If map(a%, b%) > 3 Then
            If map(a%, b%) <> cc% Then ct% = 0
            cc% = map(a%, b%)
            ct% = ct% + 1
            If ct% = 3 Then
                map(a%, b%) = 0
                map(a%, b% - 1) = 0
                map(a%, b% - 2) = 0
                
                checkabove a%, b%
                checkabove a%, b% - 1
                checkabove a%, b% - 2
                
                flag% = True
                blockcount% = blockcount% - 3
                If blockcount% < 3 Then
                    level% = level% + 1
                    loadlevel level%
                End If
                Exit For
            End If
        Else
            ct% = 0
        End If
    Next
Next
If Not flag% Then
    For b% = 0 To 5
        ct% = 0: cc% = 0
        For a% = 0 To 39
            If map(a%, b%) > 3 Then
                If map(a%, b%) <> cc% Then ct% = 0
                cc% = map(a%, b%)
                ct% = ct% + 1
                If ct% = 3 Then
                    map(a%, b%) = 0
                    map(a% - 1, b%) = 0
                    map(a% - 2, b%) = 0
                
                    checkabove a%, b%
                    checkabove a% - 1, b%
                    checkabove a% - 2, b%
                    
                    flag% = True
                    blockcount% = blockcount% - 3
                    If blockcount% < 3 Then
                        level% = level% + 1
                        loadlevel level%
                    End If
                    Exit For
                End If
            Else
                ct% = 0
            End If
        Next
    Next
End If
End Sub
Sub checkabove(cbx%, ByVal cby%)
Dim b%
Do While cby% > 0
    cby% = cby% - 1
    If map%(cbx%, cby%) > 3 Then
        For b% = 1 To 30
            If blockmode%(b%) = 0 Then 'spare block
                blockx%(b%) = cbx% * 60 + 300
                blocky%(b%) = cby% * 60 + 60
                blockcell%(b%) = map(cbx%, cby%)
                map(cbx%, cby%) = 0
                blockmode%(b%) = 3 'fall
                Exit For
            End If
        Next
    Else
        Exit Do
    End If
Loop
End Sub
