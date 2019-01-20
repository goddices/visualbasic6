VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'impactX Game Engine v1.0.0
'本类模块用于多媒体文件的回放和处理
'使用本类模块必须遵守:
'你可以免费使用本引擎及代码
'使用本引擎后的责任由使用者承担
'你可以任意拷贝本引擎代码，但必须保证其完整性
'希望我能得到你使用本引擎制作出的程序
'使用DirectShow，必须在工程->引用菜单中添加ActiveMovie control type library
'Davy.xu sunicdavy@sina.com qq:20998333
Option Explicit
Private m_objBasicAudio  As IBasicAudio         'Basic Audio Object
Private m_objBasicVideo As IBasicVideo          'Basic Video Object
Private m_objMediaEvent As IMediaEvent        'MediaEvent Object
Private m_objVideoWindow As IVideoWindow   'VideoWindow Object
Private m_objMediaControl As IMediaControl    'MediaControl Object
Private m_objMediaPosition As IMediaPosition 'MediaPosition Object
Private m_dblStartPosition As Double
Private m_dblRunLength As Double
Private m_boolVideoRunning As Boolean
Private m_Vol As Integer
Private m_Bal As Integer
Private m_hWnd As Long
Private m_Width As Integer
Private m_Height As Integer
Private m_Top As Integer
Private m_Left As Integer
'初始化设定DShow的对象参数
Public Sub InitDXShow(hWnd As Long, Width As Integer, Height As Integer, Optional Left As Integer = 0, Optional Top As Integer = 0)
    m_hWnd = hWnd
    m_Width = Width
    m_Height = Height
    m_Top = Top
    m_Left = Left
End Sub
'载入媒体,支持媒体类型为mpg,avi,wav,mov,mp3
Public Sub LoadMedia(Pathname As String)
On Local Error GoTo ErrLine
    If Mid(Pathname, 2, 1) <> ":" Then Pathname = App.Path & "\" & Pathname
    If Len(Dir(Pathname)) = 0 Then
        Debug.Print "[PlayMeida]Err:文件不存在!"
        Debug.Print Pathname
        End
    End If
    
    Set m_objMediaControl = New FilgraphManager
    Call m_objMediaControl.RenderFile(Pathname)
    Set m_objBasicAudio = m_objMediaControl
    m_objBasicAudio.Volume = (m_Vol - 100) * 40
    m_objBasicAudio.Balance = m_Bal * 50
    
    Set m_objVideoWindow = m_objMediaControl
    m_objVideoWindow.WindowStyle = CLng(&H6000000)
    m_objVideoWindow.Top = m_Top
    m_objVideoWindow.Left = m_Left
    m_objVideoWindow.Width = m_Width
    m_objVideoWindow.Height = m_Height
    m_objVideoWindow.Owner = m_hWnd
    
    Set m_objMediaEvent = m_objMediaControl '播放,停止,暂停的控制对象
    
    Set m_objMediaPosition = m_objMediaControl '媒体位置控制对象
    m_dblStartPosition = 0
    m_objMediaPosition.Rate = 1
    m_dblRunLength = Round(m_objMediaPosition.Duration, 2)
            
    DoEvents
    Exit Sub
ErrLine:
    Err.Clear
    Resume Next
End Sub
'音量的获取和设定
Public Property Get Volume() As Integer
    Volume = m_Vol
End Property
Public Property Let Volume(ByVal Vol As Integer)
    m_Vol = Vol
    m_objBasicAudio.Volume = (Vol - 100) * 40
End Property
'播放进度的获取和设置
Public Property Get MediaPosition() As Double
    MediaPosition = m_objMediaPosition.CurrentPosition
End Property
Public Property Let MediaPosition(ByVal Position As Double)
    m_objMediaPosition.CurrentPosition = Position
End Property
'声道平衡的获取和设置
Public Property Get Balance() As Integer
    Balance = m_Bal
End Property
Public Property Let Balance(ByVal bal As Integer)
    m_Bal = bal
    m_objBasicAudio.Balance = bal * 50
End Property
'获取媒体播放长度
Public Property Get Duration() As Double
    Duration = m_dblRunLength
End Property
'检测媒体是否在播放
Public Property Get isPlaying() As Boolean
    isPlaying = IIf(m_objMediaPosition.CurrentPosition < m_dblRunLength, True, False)
End Property
'播放媒体
Public Sub PlayMedia()
    If CLng(m_objMediaPosition.CurrentPosition) < CLng(m_dblStartPosition) Then
        m_objMediaPosition.CurrentPosition = m_dblStartPosition
    ElseIf CLng(m_objMediaPosition.CurrentPosition) = CLng(m_dblRunLength) Then
        m_objMediaPosition.CurrentPosition = m_dblStartPosition
    End If
    Call m_objMediaControl.Run
    m_boolVideoRunning = True
    DoEvents
    DoEvents
End Sub
'暂停播放
Public Sub PauseMedia()
    Call m_objMediaControl.Pause
    m_boolVideoRunning = False
End Sub
'停止播放
Public Sub StopMedia()
    Call m_objMediaControl.Stop
    m_boolVideoRunning = False
    m_objMediaPosition.CurrentPosition = 0
End Sub
'卸载DShow
Public Sub UnloadDXShow()
            m_boolVideoRunning = False
            DoEvents
            
            If Not m_objMediaControl Is Nothing Then
               m_objMediaControl.Stop
            End If
            
'            If Not m_objVideoWindow Is Nothing Then
'               m_objVideoWindow.Left = Screen.Width * 8
'               m_objVideoWindow.Height = Screen.Height * 8
'               m_objVideoWindow.Owner = 0
'            End If
            
            If Not m_objBasicAudio Is Nothing Then Set m_objBasicAudio = Nothing
            If Not m_objBasicVideo Is Nothing Then Set m_objBasicVideo = Nothing
            If Not m_objMediaControl Is Nothing Then Set m_objMediaControl = Nothing
            If Not m_objVideoWindow Is Nothing Then Set m_objVideoWindow = Nothing
            If Not m_objMediaPosition Is Nothing Then Set m_objMediaPosition = Nothing
End Sub
Private Sub Class_Initialize()
    m_Vol = 100
End Sub
