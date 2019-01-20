Attribute VB_Name = "GameEngineData"
Option Explicit

Public Enum DefMsgType                         '消息类型
    MSG_Keydown
    MSG_KeyUp
    MSG_KeyPress
    MSG_MouseClick
    MSG_MouseDblClick
    MSG_MouseUp
    MSG_MouseDown
    MSG_MouseMove
    MSG_MouseScroll
End Enum

Public Enum EngineSetupConst        '
    DX_SOUND_USE = 1
    DX_INPUT_USE = 2
End Enum

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type InputMessage
    MsgStyle As DefMsgType
    x As Integer
    y As Integer
    Value As Integer
End Type

Public Type LARGE_INTEGER
    low As Long
    hight As Long
End Type

Public Enum STD_BLT_MODE
    STD_BLT = 0
    FAST_BLT = 1
    FAST_BLT_EX = 2
    ALPHA_BLT = 3
    ADDTIVE_BLT = 4
    SUB_BLT = 5
    MASK_BLT = 6
    EDGELINE_BLT = 7
    ROTE_BLT = 8
    TEXT_BLT = 9
    ZBUFFER_BLT = 10
End Enum

Public Type TStdBlt
    lpSurface As Long                      '请必须保持lpSurface的首位置
    ID As Long                             '我们将使用该ID标识物体，通常用于屏幕识别
    bltmode As Long
    x As Long
    y As Long
    z As Long
    effect As Long
    frame As Long
    k As Single
    Reserve As Long
End Type

Public Const BLT_SPEED_MODE = 0             '传统的 BLT模式
Public Const ALHPA_SPEED_MODE = 1           '加速效果模式,建议在窗体下运行
Public Const SYS_KEYCOLOR = 63519           'RGBtoDDColor(255,0,255)&HF81F

Public g_DD As DirectDraw7
Public g_DX7 As DirectX7
Public g_Hwnd As Long
Public g_ViewRect As RECT
Public g_Clipper As RECT
Public g_MainSurface As CSurface
Public g_Screen As CScreen
Public g_Inputs As CInput
Public g_Sounds As CSoundWav
Public g_Err_Description As String

Public g_RBitsMask As Long
Public g_GBitsMask As Long
Public g_BBitsMask As Long
Public g_XPiexlsPerWord As Long
Public g_YPiexlsPerWord As Long

Public g_Mode As Integer
Public g_Windowed As Boolean

Public DEBUG_STICK As LARGE_INTEGER

'======================Alpha DLL  声明========================================
'注意：对于指针变量可以声明为 ByRef data As Any
'..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll

#If RealseVersion = 0 Then
    Public Declare Function bitsmovl Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
        (ByVal data As Long, ByVal Bits As Byte) As Long
    Public Declare Function bitsmovr Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
        (ByVal data As Long, ByVal Bits As Byte) As Long
    Public Declare Function getbit Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
        (ByVal data As Long, ByVal Bits As Byte) As Long
    Public Declare Sub setpixel Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (lpsrc As Long, ByVal x As Long, ByVal y As Long, ByVal iDstPitch As Long, ByVal color As Long)
    Public Declare Function GetPixel Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       Alias "getpixel" (lpsrc As Long, ByVal x As Long, ByVal y As Long, ByVal iDstPitch As Long) As Long
    Public Declare Function blendcolor Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (ByVal color0 As Long, ByVal color1 As Long, ByVal alph As Byte) As Long
    Public Declare Sub colorblend_565 Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (ByVal alpha As Byte, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, ByVal iSrcPitch As Long, _
       lpDst As Long, ByVal dx As Long, ByVal dy As Long, ByVal iDstPitch As Long, _
       ByVal W As Long, ByVal H As Long, ByVal keycolor As Long, ByVal blendcolor As Long)
    Public Declare Sub mask_565_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
        (lpsrc As Long, ByVal iSrcX As Long, ByVal iSrcY As Long, ByVal iSrcPitch As Long, _
        lpDst As Long, ByVal iDstX As Long, ByVal iDstY As Long, ByVal iDstPitch As Long, _
        ByVal iDstW As Long, ByVal iDstH As Long, ByVal mask As Long, ByVal keycolor As Long)

    Public Declare Sub bltfast Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (lpDst As Long, ByVal dx As Long, ByVal dy As Long, ByVal iDstPitch As Long, _
       ByVal W As Long, ByVal H As Long, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
       ByVal iSrcPitch As Long)
       
    Public Declare Sub Qmemcpy Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (lpDst As Long, lpsrc As Long, ByVal nQWORDs As Long)
       
    Public Declare Sub Qmemset Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (lpDst As Long, ByVal c As Long, ByVal nQWORDs As Long)
       
    Public Declare Sub memcopy_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (lpDst As Long, lpsrc As Long, ByVal nbytes As Long)
       
    Public Declare Sub memrecopy_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (lpDst As Long, lpsrc As Long, ByVal nbytes As Long)

    Public Declare Sub fast_additive_565_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (lpsrc As Long, ByVal sx As Long, ByVal sy As Long, ByVal iSrcPitch As Long, _
       lpDst As Long, ByVal dx As Long, ByVal dy As Long, ByVal iDstPitch As Long, _
       ByVal W As Long, ByVal H As Long)

    Public Declare Sub blt_to_lighttable_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (lplight As Long, ByVal dx As Long, ByVal dy As Long, ByVal W As Long, ByVal H As Long, ByVal iDstPitch As Long, _
       lpsrc As Long, ByVal sx As Long, ByVal sy As Long, ByVal iSrcPitch As Long)

    Public Declare Sub bltfast_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (lpDst As Long, ByVal dx As Long, ByVal dy As Long, ByVal iDstPitch As Long, _
       ByVal W As Long, ByVal H As Long, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
       ByVal iSrcPitch As Long, ByVal keycolor As Long)
    Public Declare Sub bltzoom_565_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (lpDst As Long, ByVal dx As Long, ByVal dy As Long, ByVal iDstPitch As Long, _
       ByVal dw As Long, ByVal dh As Long, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
       ByVal iSrcPitch As Long, ByVal sw As Long, ByVal sh As Long, ByVal keycolor As Long)
    'New
    Public Declare Sub bltzoom_additive_565_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (ByVal alpha As Byte, lpDst As Long, ByVal dx As Long, ByVal dy As Long, ByVal iDstPitch As Long, _
       ByVal dw As Long, ByVal dh As Long, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
       ByVal iSrcPitch As Long, ByVal sw As Long, ByVal sh As Long, ByVal keycolor As Long)
    Public Declare Sub bltzoom_ablend_565_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (ByVal alpha As Byte, lpDst As Long, ByVal dx As Long, ByVal dy As Long, ByVal iDstPitch As Long, _
       ByVal dw As Long, ByVal dh As Long, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
       ByVal iSrcPitch As Long, ByVal sw As Long, ByVal sh As Long, ByVal keycolor As Long)
       
    Public Declare Sub DrawAlpha Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
        (ByVal x As Long, ByVal y As Long, ByVal sx As Long, ByVal sy As Long, _
        ByVal W As Long, ByVal H As Long, ByVal spith As Long, ByVal dpith As Long, _
        ByVal alpha As Byte, ByVal keycolor As Long, dst As Long, src As Long)
    Public Declare Sub AddActive Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
        (ByVal x As Long, ByVal y As Long, ByVal sx As Long, ByVal sy As Long, _
        ByVal W As Long, ByVal H As Long, ByVal spith As Long, ByVal dpith As Long, _
        ByVal alpha As Byte, ByVal keycolor As Long, dst As Long, src As Long)
    Public Declare Sub DrawRect Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
        (ByVal x As Long, ByVal y As Long, ByVal W As Long, ByVal H As Long, _
        ByVal spith As Long, ByVal alpha As Byte, src As Long) ' Byte)
    Public Declare Sub ablend_565_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (ByVal alpha As Byte, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
       ByVal iSrcPitch As Long, lpDst As Long, ByVal dx As Long, ByVal dy As Long, _
        ByVal iDstPitch As Long, ByVal W As Long, ByVal H As Long, ByVal keycolor As Long)
    Public Declare Sub scanx_565 Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (lpsrc As Long, ByVal sx As Long, ByVal sy As Long, ByVal iSrcPitch As Long, _
        lpDst As Long, ByVal dx As Long, ByVal dy As Long, ByVal iDstPitch As Long, _
        ByVal W As Long, ByVal H As Long, ByVal color As Long, ByVal keycolor As Long)
    Public Declare Sub addlightex_565_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (lpsrc As Long, ByVal sx As Long, ByVal sy As Long, ByVal iSrcPitch As Long, _
        lpDst As Long, ByVal dx As Long, ByVal dy As Long, ByVal iDstPitch As Long, _
        ByVal W As Long, ByVal H As Long, ByVal keycolor As Long, ByVal color As Long)
    Public Declare Sub additive_565_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (ByVal alpha As Byte, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
        ByVal iSrcPitch As Long, lpDst As Long, ByVal dx As Long, ByVal dy As Long, _
        ByVal iDstPitch As Long, ByVal W As Long, ByVal H As Long, ByVal keycolor As Long)
        
    Public Declare Sub addcolor_565_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (lpsrc As Long, ByVal sx As Long, ByVal sy As Long, ByVal iSrcPitch As Long, _
        ByVal W As Long, ByVal H As Long, ByVal color As Long)
   
    Public Declare Sub light_565_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
       ByVal iSrcPitch As Long, lpDst As Long, ByVal dx As Long, ByVal dy As Long, _
        ByVal iDstPitch As Long, ByVal W As Long, ByVal H As Long, ByVal keycolor As Long)
    Public Declare Sub fastlight_565_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
        (lpDst As Long, lpsrc As Long, lpTable As Byte, ByVal x As Long, ByVal y As Long, _
        ByVal iDstPitch As Long, ByVal iSrcPitch As Long, ByVal iTablePitch As Long, ByVal W As Long, ByVal H As Long)
    Public Declare Sub memset_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
        (lpDst As Any, ByVal pitch As Long, ByVal W As Long, ByVal H As Long, ByVal data As Byte)
    Public Declare Sub fastmemset Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
        (lpDst As Any, ByVal bytesize As Long, ByVal data As Byte)
 
    Public Declare Sub subitive_565_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (ByVal alpha As Byte, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
       ByVal iSrcPitch As Long, lpDst As Long, ByVal dx As Long, ByVal dy As Long, _
        ByVal iDstPitch As Long, ByVal W As Long, ByVal H As Long, ByVal keycolor As Long)
    Public Declare Sub alpharect_565_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (ByVal alpha As Byte, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
       ByVal iSrcPitch As Long, ByVal W As Long, ByVal H As Long)
    Public Declare Sub halfablend_565_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
       (lpsrc As Long, ByVal sx As Long, ByVal sy As Long, ByVal iSrcPitch As Long, _
       lpDst As Long, ByVal dx As Long, ByVal dy As Long, ByVal iDstPitch As Long, _
       ByVal W As Long, ByVal H As Long)
    Public Declare Sub rotate_tran Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
        (lpscreen As Any, ByVal screenpitch As Long, lpbmp As Any, ByVal bmppitch As Long, ByVal x As Long, ByVal y As Long, ByVal dw As Long, ByVal dh As Long, ByVal sx As Long, ByVal sy As Long, ByVal sw As Long, ByVal sh As Long, ByVal angle As Single, ByVal keycolor As Long)
    Public Declare Sub gray_565_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
        (lpscreen As Any, ByVal screenpitch As Long, lpbmp As Any, ByVal bmppitch As Long, ByVal x As Long, ByVal y As Long, _
        ByVal sx As Long, ByVal sy As Long, ByVal W As Long, ByVal H As Long, ByVal keycolor As Long)
    
    Public Declare Sub renderipple Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
            (lpscreen As Long, ByVal screenpitch As Long, lpbmp As Long, ByVal bmppitch As Long, lpbuf As Any, ByVal W As Long, ByVal H As Long)
    Public Declare Sub ripplespread Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
            (lpbuf As Any, lpoldbuf As Any, ByVal W As Long, ByVal H As Long)
    Public Declare Sub blur_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
            (lpscreen As Any, ByVal screenpitch As Long, lpbmp As Any, ByVal bmppitch As Long, ByVal W As Long, ByVal H As Long)
    Public Declare Sub blur_c Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
            (lpscreen As Any, ByVal screenpitch As Long, ByVal x As Long, ByVal y As Long, ByVal W As Long, ByVal H As Long)

    Public Declare Sub zbuffer_blt_mmx Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
        (lpscreen As Long, lpzbuffer As Long, ByVal x As Long, ByVal y As Long, ByVal z As Long, _
        ByVal scrw As Long, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
        ByVal spitch As Long, ByVal sw As Long, ByVal sh As Long, ByVal keycolor As Long)
        
    Public Declare Sub memsetw Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
        (lpDst As Any, ByVal wordsize As Long, ByVal data As Integer)

    Public Declare Sub scan_linexy Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
    (lpDst As Any, ByVal x As Long, ByVal y As Long, ByVal dpitch As Long, _
    lpsrc As Any, ByVal sx As Long, ByVal sy As Long, ByVal spitch As Long, _
    ByVal W As Long, ByVal H As Long, ByVal color As Long, ByVal keycolor As Long)
    'extern _stdcall rle_blt(char *lpdst,long dpitch,long x,long y,char *lpsrc,long pointernum)
    Public Declare Sub rle_blt Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
        (lpDst As Any, ByVal dpitch As Long, ByVal H As Long, ByVal x As Long, ByVal y As Long, _
        lpsrc As Any, ByVal pointernum As Long)
        
    Public Declare Function RGB565 Lib "..\..\EngineSourceCodes\DD_Alpha\Release\alpha.dll" _
        (RGB555 As Integer) As Integer
    
#Else
    Public Declare Function bitsmovl Lib "alpha.dll" _
        (ByVal data As Long, ByVal Bits As Byte) As Long
    Public Declare Function bitsmovr Lib "alpha.dll" _
        (ByVal data As Long, ByVal Bits As Byte) As Long
    Public Declare Function getbit Lib "alpha.dll" _
        (ByVal data As Long, ByVal Bits As Byte) As Long
    Public Declare Sub setpixel Lib "alpha.dll" _
       (lpsrc As Long, ByVal x As Long, ByVal y As Long, ByVal iDstPitch As Long, ByVal color As Long)
    Public Declare Function GetPixel Lib "alpha.dll" _
       Alias "getpixel" (lpsrc As Long, ByVal x As Long, ByVal y As Long, ByVal iDstPitch As Long) As Long
    Public Declare Function blendcolor Lib "alpha.dll" _
       (ByVal color0 As Long, ByVal color1 As Long, ByVal alph As Byte) As Long
    Public Declare Sub colorblend_565 Lib "alpha.dll" _
       (ByVal alpha As Byte, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, ByVal iSrcPitch As Long, _
       lpDst As Long, ByVal dx As Long, ByVal dy As Long, ByVal iDstPitch As Long, _
       ByVal W As Long, ByVal H As Long, ByVal keycolor As Long, ByVal blendcolor As Long)
    Public Declare Sub mask_565_mmx Lib "alpha.dll" _
        (lpsrc As Long, ByVal iSrcX As Long, ByVal iSrcY As Long, ByVal iSrcPitch As Long, _
        lpDst As Long, ByVal iDstX As Long, ByVal iDstY As Long, ByVal iDstPitch As Long, _
        ByVal iDstW As Long, ByVal iDstH As Long, ByVal mask As Long, ByVal keycolor As Long)

    Public Declare Sub bltfast Lib "alpha.dll" _
       (lpDst As Long, ByVal dx As Long, ByVal dy As Long, ByVal iDstPitch As Long, _
       ByVal W As Long, ByVal H As Long, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
       ByVal iSrcPitch As Long)
       
    Public Declare Sub Qmemcpy Lib "alpha.dll" _
       (lpDst As Long, lpsrc As Long, ByVal nQWORDs As Long)
       
    Public Declare Sub Qmemset Lib "alpha.dll" _
       (lpDst As Long, ByVal c As Long, ByVal nQWORDs As Long)
       
    Public Declare Sub memcopy_mmx Lib "alpha.dll" _
       (lpDst As Long, lpsrc As Long, ByVal nbytes As Long)
       
    Public Declare Sub memrecopy_mmx Lib "alpha.dll" _
       (lpDst As Long, lpsrc As Long, ByVal nbytes As Long)

    Public Declare Sub fast_additive_565_mmx Lib "alpha.dll" _
       (lpsrc As Long, ByVal sx As Long, ByVal sy As Long, ByVal iSrcPitch As Long, _
       lpDst As Long, ByVal dx As Long, ByVal dy As Long, ByVal iDstPitch As Long, _
       ByVal W As Long, ByVal H As Long)

    Public Declare Sub blt_to_lighttable_mmx Lib "alpha.dll" _
       (lplight As Long, ByVal dx As Long, ByVal dy As Long, ByVal W As Long, ByVal H As Long, ByVal iDstPitch As Long, _
       lpsrc As Long, ByVal sx As Long, ByVal sy As Long, ByVal iSrcPitch As Long)

    Public Declare Sub bltfast_mmx Lib "alpha.dll" _
       (lpDst As Long, ByVal dx As Long, ByVal dy As Long, ByVal iDstPitch As Long, _
       ByVal W As Long, ByVal H As Long, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
       ByVal iSrcPitch As Long, ByVal keycolor As Long)
    Public Declare Sub bltzoom_565_mmx Lib "alpha.dll" _
       (lpDst As Long, ByVal dx As Long, ByVal dy As Long, ByVal iDstPitch As Long, _
       ByVal dw As Long, ByVal dh As Long, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
       ByVal iSrcPitch As Long, ByVal sw As Long, ByVal sh As Long, ByVal keycolor As Long)
    'New
    Public Declare Sub bltzoom_additive_565_mmx Lib "alpha.dll" _
       (ByVal alpha As Byte, lpDst As Long, ByVal dx As Long, ByVal dy As Long, ByVal iDstPitch As Long, _
       ByVal dw As Long, ByVal dh As Long, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
       ByVal iSrcPitch As Long, ByVal sw As Long, ByVal sh As Long, ByVal keycolor As Long)
    Public Declare Sub bltzoom_ablend_565_mmx Lib "alpha.dll" _
       (ByVal alpha As Byte, lpDst As Long, ByVal dx As Long, ByVal dy As Long, ByVal iDstPitch As Long, _
       ByVal dw As Long, ByVal dh As Long, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
       ByVal iSrcPitch As Long, ByVal sw As Long, ByVal sh As Long, ByVal keycolor As Long)
       
    Public Declare Sub DrawAlpha Lib "alpha.dll" _
        (ByVal x As Long, ByVal y As Long, ByVal sx As Long, ByVal sy As Long, _
        ByVal W As Long, ByVal H As Long, ByVal spith As Long, ByVal dpith As Long, _
        ByVal alpha As Byte, ByVal keycolor As Long, dst As Long, src As Long)
    Public Declare Sub AddActive Lib "alpha.dll" _
        (ByVal x As Long, ByVal y As Long, ByVal sx As Long, ByVal sy As Long, _
        ByVal W As Long, ByVal H As Long, ByVal spith As Long, ByVal dpith As Long, _
        ByVal alpha As Byte, ByVal keycolor As Long, dst As Long, src As Long)
    Public Declare Sub DrawRect Lib "alpha.dll" _
        (ByVal x As Long, ByVal y As Long, ByVal W As Long, ByVal H As Long, _
        ByVal spith As Long, ByVal alpha As Byte, src As Long) ' Byte)
    Public Declare Sub ablend_565_mmx Lib "alpha.dll" _
       (ByVal alpha As Byte, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
       ByVal iSrcPitch As Long, lpDst As Long, ByVal dx As Long, ByVal dy As Long, _
        ByVal iDstPitch As Long, ByVal W As Long, ByVal H As Long, ByVal keycolor As Long)
    Public Declare Sub scanx_565 Lib "alpha.dll" _
       (lpsrc As Long, ByVal sx As Long, ByVal sy As Long, ByVal iSrcPitch As Long, _
        lpDst As Long, ByVal dx As Long, ByVal dy As Long, ByVal iDstPitch As Long, _
        ByVal W As Long, ByVal H As Long, ByVal color As Long, ByVal keycolor As Long)
    Public Declare Sub addlightex_565_mmx Lib "alpha.dll" _
       (lpsrc As Long, ByVal sx As Long, ByVal sy As Long, ByVal iSrcPitch As Long, _
        lpDst As Long, ByVal dx As Long, ByVal dy As Long, ByVal iDstPitch As Long, _
        ByVal W As Long, ByVal H As Long, ByVal keycolor As Long, ByVal color As Long)
    Public Declare Sub additive_565_mmx Lib "alpha.dll" _
       (ByVal alpha As Byte, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
        ByVal iSrcPitch As Long, lpDst As Long, ByVal dx As Long, ByVal dy As Long, _
        ByVal iDstPitch As Long, ByVal W As Long, ByVal H As Long, ByVal keycolor As Long)
        
    Public Declare Sub addcolor_565_mmx Lib "alpha.dll" _
       (lpsrc As Long, ByVal sx As Long, ByVal sy As Long, ByVal iSrcPitch As Long, _
        ByVal W As Long, ByVal H As Long, ByVal color As Long)
   
    Public Declare Sub light_565_mmx Lib "alpha.dll" _
       (lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
       ByVal iSrcPitch As Long, lpDst As Long, ByVal dx As Long, ByVal dy As Long, _
        ByVal iDstPitch As Long, ByVal W As Long, ByVal H As Long, ByVal keycolor As Long)
    Public Declare Sub fastlight_565_mmx Lib "alpha.dll" _
        (lpDst As Long, lpsrc As Long, lpTable As Byte, ByVal x As Long, ByVal y As Long, _
        ByVal iDstPitch As Long, ByVal iSrcPitch As Long, ByVal iTablePitch As Long, ByVal W As Long, ByVal H As Long)
    Public Declare Sub memset_mmx Lib "alpha.dll" _
        (lpDst As Any, ByVal pitch As Long, ByVal W As Long, ByVal H As Long, ByVal data As Byte)
    Public Declare Sub fastmemset Lib "alpha.dll" _
        (lpDst As Any, ByVal bytesize As Long, ByVal data As Byte)
 
    Public Declare Sub subitive_565_mmx Lib "alpha.dll" _
       (ByVal alpha As Byte, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
       ByVal iSrcPitch As Long, lpDst As Long, ByVal dx As Long, ByVal dy As Long, _
        ByVal iDstPitch As Long, ByVal W As Long, ByVal H As Long, ByVal keycolor As Long)
    Public Declare Sub alpharect_565_mmx Lib "alpha.dll" _
       (ByVal alpha As Byte, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
       ByVal iSrcPitch As Long, ByVal W As Long, ByVal H As Long)
    Public Declare Sub halfablend_565_mmx Lib "alpha.dll" _
       (lpsrc As Long, ByVal sx As Long, ByVal sy As Long, ByVal iSrcPitch As Long, _
       lpDst As Long, ByVal dx As Long, ByVal dy As Long, ByVal iDstPitch As Long, _
       ByVal W As Long, ByVal H As Long)
    Public Declare Sub rotate_tran Lib "alpha.dll" _
        (lpscreen As Any, ByVal screenpitch As Long, lpbmp As Any, ByVal bmppitch As Long, ByVal x As Long, ByVal y As Long, ByVal dw As Long, ByVal dh As Long, ByVal sx As Long, ByVal sy As Long, ByVal sw As Long, ByVal sh As Long, ByVal angle As Single, ByVal keycolor As Long)
    Public Declare Sub gray_565_mmx Lib "alpha.dll" _
        (lpscreen As Any, ByVal screenpitch As Long, lpbmp As Any, ByVal bmppitch As Long, ByVal x As Long, ByVal y As Long, _
        ByVal sx As Long, ByVal sy As Long, ByVal W As Long, ByVal H As Long, ByVal keycolor As Long)
    
    Public Declare Sub renderipple Lib "alpha.dll" _
            (lpscreen As Long, ByVal screenpitch As Long, lpbmp As Long, ByVal bmppitch As Long, lpbuf As Any, ByVal W As Long, ByVal H As Long)
    Public Declare Sub ripplespread Lib "alpha.dll" _
            (lpbuf As Any, lpoldbuf As Any, ByVal W As Long, ByVal H As Long)
    Public Declare Sub blur_mmx Lib "alpha.dll" _
            (lpscreen As Any, ByVal screenpitch As Long, lpbmp As Any, ByVal bmppitch As Long, ByVal W As Long, ByVal H As Long)
    Public Declare Sub blur_c Lib "alpha.dll" _
            (lpscreen As Any, ByVal screenpitch As Long, ByVal x As Long, ByVal y As Long, ByVal W As Long, ByVal H As Long)

    Public Declare Sub zbuffer_blt_mmx Lib "alpha.dll" _
        (lpscreen As Long, lpzbuffer As Long, ByVal x As Long, ByVal y As Long, ByVal z As Long, _
        ByVal scrw As Long, lpsrc As Long, ByVal sx As Long, ByVal sy As Long, _
        ByVal spitch As Long, ByVal sw As Long, ByVal sh As Long, ByVal keycolor As Long)
        
    Public Declare Sub memsetw Lib "alpha.dll" _
        (lpDst As Any, ByVal wordsize As Long, ByVal data As Integer)

    Public Declare Sub scan_linexy Lib "alpha.dll" _
    (lpDst As Any, ByVal x As Long, ByVal y As Long, ByVal dpitch As Long, _
    lpsrc As Any, ByVal sx As Long, ByVal sy As Long, ByVal spitch As Long, _
    ByVal W As Long, ByVal H As Long, ByVal color As Long, ByVal keycolor As Long)
    'extern _stdcall rle_blt(char *lpdst,long dpitch,long x,long y,char *lpsrc,long pointernum)
    Public Declare Sub rle_blt Lib "alpha.dll" _
        (lpDst As Any, ByVal dpitch As Long, ByVal H As Long, ByVal x As Long, ByVal y As Long, _
        lpsrc As Any, ByVal pointernum As Long)
        
    Public Declare Function RGB565 Lib "alpha.dll" _
        (RGB555 As Integer) As Integer
#End If

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'GDI——API 声明
Public Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal HDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal HDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Public Declare Function GetDCOrgEx Lib "gdi32" (ByVal HDC As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
'可以使用下面这个函数检测CPU是否支持MMX,SSE指令:
'int CheckMMX( )
'{
'    int isMMX=0;
'    __asm
'    {
'        mov eax,1;
'        cpuid;
'        test edx,00800000h;
'        jz  NotSupport;
'        mov  isMMX,1;
'NotSupport:
'    }
'    return isMMX;
'}
'int CheckSSE()
'{
'    int isSSE = 0;
'    _asm
'{
'        mov eax, 1
'        cpuid
'        shr edx,0x1A
'        jnc NotSupport
'        mov isSSE, 1
'NotSupport:
'}
'    return isSSE;
'}

Public Sub DDColortoRGB(ByVal DDcolor As Long, R As Byte, G As Byte, B As Byte)
    'rgb 5-6-5
    R = (DDcolor And g_RBitsMask) / 256
    G = (DDcolor And g_GBitsMask) / 8
    B = (DDcolor And g_BBitsMask) * 8
End Sub

Public Function RGBtoDDColor(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte) As Long
    RGBtoDDColor = ((R / 255) * g_RBitsMask And g_RBitsMask) + ((G / 255) * g_GBitsMask And g_GBitsMask) + ((B / 255) * g_BBitsMask And g_BBitsMask)
End Function

Public Function CheckRect(srcRect As RECT, ByVal x As Long, ByVal y As Long) As Boolean
    CheckRect = x > srcRect.Left And x < srcRect.Right And y > srcRect.Top And y < srcRect.Bottom
End Function

Public Function MoveRect(srcRect As RECT, ByVal x As Long, ByVal y As Long) As RECT
    MoveRect.Left = x
    MoveRect.Top = y
    MoveRect.Right = x + srcRect.Right - srcRect.Left
    MoveRect.Bottom = y + srcRect.Bottom - srcRect.Top
End Function
Public Sub setbit0(data As Long, ByVal Bits As Byte)
    data = data And (Not bitsmovl(1, Bits - 1))
End Sub

Public Sub setbit1(data As Long, ByVal Bits As Byte)
    data = data Or bitsmovl(1, Bits - 1)
End Sub

Public Sub SetScreenFont(font As StdFont)
    g_MainSurface.DD_Surface.SetFont font
End Sub

Public Sub StickStart()
    '开始性能计时器
    QueryPerformanceCounter DEBUG_STICK
End Sub

Public Function StickEnd() As Long
    '计时结束
    Dim tmpStick As LARGE_INTEGER
    QueryPerformanceCounter tmpStick
    StickEnd = tmpStick.low - DEBUG_STICK.low
End Function


'*************************************************************************
'**函 数 名：GetResPath
'**输    入：无
'**输    出：(String) -
'**功能描述：得到ＳＤＫ的演示资源目录，便于ＳＤＫ的演示使用，用户可以删除
'**作    者：王慧平
'**日    期：2005-07-05 19:11:55
'**修 改 人：
'**日    期：
'**版    本：V1.0.0
'*************************************************************************
Public Function GetResPath() As String
    Dim srcPath As String
    srcPath = App.Path
    GetResPath = Left(srcPath, InStrRev(srcPath, "\")) & "\res\"
End Function
