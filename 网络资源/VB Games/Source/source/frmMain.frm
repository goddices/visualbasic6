VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Particles"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'//WARNING: D3DInit and D3DUtil have been altered
'//they are no equal to microsofts version. This one runs at fixed resolution

Private Sub Form_KeyPress(KeyAscii As Integer)
    Form_Unload 0
End Sub

Private Sub Form_Load()
    Dim m_bInit As Boolean
    
    Randomize
    
    Me.Show
    DoEvents
    
    m_bInit = D3DUtil_Init(Me.hwnd, False, 0, 0, D3DDEVTYPE_HAL, Me)
    If Not (m_bInit) Then End

    '//Reset states
    ResetStates
    
    Dim myNewFire As colFireFX
    Set myNewFire = New colFireFX
    myNewFire.ParticleCounts = 150
    myNewFire.ReLocate 150, 200
    myNewFire.Begin

    Dim myBlueThrust As colBlueThrust
    Set myBlueThrust = New colBlueThrust
    myBlueThrust.ParticleCounts = 100
    myBlueThrust.ReLocate 350, 200
    myBlueThrust.Begin
    
    Dim myFlower As colRedFlower
    Set myFlower = New colRedFlower
    myFlower.ParticleCounts = 150
    myFlower.ReLocate 550, 200
    myFlower.Begin
    
    
    Dim myAtomic As colAtomic
    Set myAtomic = New colAtomic
    myAtomic.ParticleCounts = 300
    myAtomic.ReLocate 750, 200
    myAtomic.Begin
    

    Dim mySmoke As colSmokeFX
    Set mySmoke = New colSmokeFX
    mySmoke.ParticleCounts = 20
    mySmoke.ReLocate 150, 400
    mySmoke.Begin
    

    Dim myRedTwirl As colRedTwirl
    Set myRedTwirl = New colRedTwirl
    myRedTwirl.ParticleCounts = 20
    myRedTwirl.ReLocate 350, 400
    myRedTwirl.Begin
    
    Dim myBlueTwirl As colBlueTwirl
    Set myBlueTwirl = New colBlueTwirl
    myBlueTwirl.ParticleCounts = 20
    myBlueTwirl.ReLocate 550, 400
    myBlueTwirl.Begin
    
    
    Dim myGalaxy As colGalaxy
    Set myGalaxy = New colGalaxy
    myGalaxy.ParticleCounts = 300
    myGalaxy.ReLocate 750, 400
    myGalaxy.Begin
    
    Dim myHeart As colFireyHeart
    Set myHeart = New colFireyHeart
    myHeart.ParticleCounts = 314
    myHeart.ReLocate 150, 600
    myHeart.Begin
    
    Dim myExplosion As colBlueExplosion
    Set myExplosion = New colBlueExplosion
    myExplosion.ParticleCounts = 200
    myExplosion.ReLocate 350, 600
    myExplosion.Begin
    
    Dim myGreenPlasma As colGreenPlasma
    Set myGreenPlasma = New colGreenPlasma
    myGreenPlasma.ParticleCounts = 300
    myGreenPlasma.ReLocate 550, 600
    myGreenPlasma.Begin
    
    Dim myWormhole As colWormhole
    Set myWormhole = New colWormhole
    myWormhole.ParticleCounts = 400
    myWormhole.ReLocate 750, 600
    myWormhole.Begin
    
    Dim mySign As colAnimatedSign
    Set mySign = New colAnimatedSign
    mySign.ParticleCounts = 1000
    mySign.ReLocate 512, 384
    mySign.Begin
    
    Set myTexture = g_d3dx.CreateTextureFromFileEx(g_dev, App.Path & "\particle.bmp", 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, D3DColorARGB(255, 0, 0, 0), ByVal 0, ByVal 0)
    
    Do

        
        g_dev.Clear 0, ByVal 0&, D3DCLEAR_TARGET, 0, 1, 0
        g_dev.BeginScene
        
            myNewFire.Update
            myNewFire.Render

            myBlueThrust.Update
            myBlueThrust.Render

            myFlower.Update
            myFlower.Render

            myAtomic.Update
            myAtomic.Render
            
            mySmoke.Update
            mySmoke.Render

            myRedTwirl.Update
            myRedTwirl.Render

            myBlueTwirl.Update
            myBlueTwirl.Render

            myGalaxy.Update
            myGalaxy.Render

            myHeart.Update
            myHeart.Render

            myExplosion.Update
            myExplosion.Render

            myGreenPlasma.Update
            myGreenPlasma.Render

            myWormhole.Update
            myWormhole.Render
            
            mySign.Update
            mySign.Render

        
        g_dev.EndScene
        g_dev.Present ByVal 0&, ByVal 0&, 0, ByVal 0&
               
        DoEvents
    Loop
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    End
End Sub
