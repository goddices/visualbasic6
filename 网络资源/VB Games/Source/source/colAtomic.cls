VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colAtomic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Particles() As Particle
Private vertsPoints() As typeTRANSLITVERTEX

Private mCol As Collection

Private sngX As Single, sngY As Single
Public ParticleCounts As Long

Private lngFloat0 As Long
Private lngFloat1 As Long
Private lngFloatSize As Long

Private sngProgression As Single
Private intDirection As Integer

Private lngPreviousFrame As Long    '//For timing

Public Sub Begin()
    '//We initialize our stuff here
    Dim I As Long
    
    lngFloat0 = FtoDW(0)
    lngFloat1 = FtoDW(1)
    lngFloatSize = FtoDW(15) '//Size of our flame particles..
    
    '//Redim our particles to the particlecount
    ReDim Particles(0 To ParticleCounts)
    
    '//Redim vertices to the particle count
    '//Point sprites, so 1 per particle
    ReDim vertsPoints(0 To ParticleCounts)
        
    '//Now generate all particles
    For I = 0 To ParticleCounts
        Set Particles(I) = New Particle
        vertsPoints(I).rhw = 1
        Call Reset(I)

    Next I

    intDirection = 1

    '//Set initial time
    lngPreviousFrame = timeGetTime

End Sub

Public Sub Reset(I As Long)
    Dim X As Single, Y As Single
    Dim r As Single
    
    r = 10 + Sin(2 * (I / 10)) * 50
    X = sngX + r * Cos(I / 30)
    Y = sngY + r * Sin(I / 30)
    
    '//This is were we will reset individual particles.
    With Particles(I)
        Call .ResetIt(X, Y, 1, 1, 0, 0, 2)
        Call .ResetColor(0.4, 0.25, 1, 1, 0.3 + (0.2 * Rnd))  '+ Rnd * 0.3)
    End With
End Sub

Public Sub Update()
    Dim I As Long
    Dim sngElapsedTime As Single
    
    '//We calculate the time difference here
    sngElapsedTime = (timeGetTime - lngPreviousFrame) / 100
    lngPreviousFrame = timeGetTime
    
    For I = 0 To ParticleCounts
        With Particles(I)
            Call .UpdateParticle(sngElapsedTime)
            
            '//If the particle is invisible, reset it again.
            If .sngA <= 0 Then
                Call Reset(I)
            End If
        
            vertsPoints(I).color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
            vertsPoints(I).X = .sngX
            vertsPoints(I).Y = .sngY
        
        End With
    Next I
End Sub

Public Sub Render()
    With g_dev
        '//Set the render states for using point sprites
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1 'True
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0 'True
        .SetRenderState D3DRS_POINTSIZE, lngFloatSize
        .SetRenderState D3DRS_POINTSIZE_MIN, lngFloat0
        .SetRenderState D3DRS_POINTSCALE_A, lngFloat0
        .SetRenderState D3DRS_POINTSCALE_B, lngFloat0
        .SetRenderState D3DRS_POINTSCALE_C, lngFloat1
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        .SetRenderState D3DRS_ALPHABLENDENABLE, 1
        
        '//Set up the vertex shader
        .SetVertexShader D3DFVF_TLVERTEX
        
        '//Set our texture
        .SetTexture 0, myTexture
        
        '//And draw all our particles :D
        .DrawPrimitiveUP D3DPT_POINTLIST, ParticleCounts, vertsPoints(0), Len(vertsPoints(0))
        
        '//Reset states back for normal rendering
        .SetRenderState D3DRS_ALPHABLENDENABLE, 0
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 0 'False
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0 'False
    End With
End Sub

Public Sub ReLocate(sngNewX As Single, sngNewY As Single)
    sngX = sngNewX
    sngY = sngNewY
End Sub



