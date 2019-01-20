Attribute VB_Name = "ModstdPicture"
Option Explicit

Private Declare Function CreateStreamOnHGlobal Lib "ole32.dll" ( _
   ByVal hGlobal As Long, _
   ByVal fDeleteOnRelease As Long, _
   lpIStream As IUnknown) As Long

Private Declare Function OleLoadPicture Lib "oleaut32.dll" ( _
   ByVal lpStream As IUnknown, _
   ByVal lSize As Long, _
   ByVal fRunmode As Long, _
   riid As Any, _
   lpIPicture As IPicture) As Long
'Download by http://www.codefans.net
Public Function BytesToPicture(PictureData() As Byte) As StdPicture

     Dim IID_IPicture(3) As Long
     Dim oPicture As IPicture
     Dim nResult As Long
     Dim oStream As IUnknown
     Dim hGlobal As Long

     ' Array f¨¹llen um den KlassenID (CLSID) IID_IPICTURE
     IID_IPicture(0) = &H7BF80980
     IID_IPicture(1) = &H101ABF32
     IID_IPicture(2) = &HAA00BB8B
     IID_IPicture(3) = &HAB0C3000

   ' Stream erstellen
     Call CreateStreamOnHGlobal(VarPtr(PictureData(LBound(PictureData))), 0, oStream)

   ' OLE IPicture-Objekt erstellen
     nResult = OleLoadPicture(oStream, 0, 0, IID_IPicture(0), oPicture)
     If nResult = 0 Then
         Set BytesToPicture = oPicture
     End If

End Function




