Attribute VB_Name = "Modfile"
Option Explicit
'Download by http://www.codefans.net
Public Function ExtractFilePath(Value As String) As String '��ȡ�ļ����ڵ�·��

On Error Resume Next

 Dim tmpCount As Integer
 Dim MainCount As Integer
 
 tmpCount = Len(Value)
 
 For MainCount = 0 To Len(Value)
 
  If Mid$(Value, tmpCount, 1) <> "\" Then
   
   tmpCount = tmpCount - 1
  
  Else
  
   ExtractFilePath = Left$(Value, tmpCount)
   
   Exit Function
   
  End If
  
 Next MainCount
  
End Function

Public Function ExtractFileName(TarStrings As String) As String    '��һ�������ļ�����·������ȡ�ļ���

 On Error Resume Next
 
 Dim Tmp As String
 Dim MainCount As Integer
 Dim tmpCount As Integer
 
 
  Tmp = ""
  tmpCount = Len(TarStrings)
  
  For MainCount = 0 To Len(TarStrings)
  
   
   If Mid$(TarStrings, tmpCount, 1) <> "\" Then
   
    Tmp = Mid$(TarStrings, tmpCount, 1) + Tmp
    tmpCount = tmpCount - 1
   
   Else
    
    ExtractFileName = Tmp
    Exit Function
    
   End If

  Next MainCount

End Function

Public Function ExtractFileExt(Value As String) As String  '��ȡ�ļ��ĺ����

 On Error Resume Next
 
  Dim Tmp As String
  Dim tmpCount As Integer
  Dim MainCount As Integer
  
  tmpCount = Len(Value)
  
  For MainCount = 0 To Len(Value)
    
   If Mid$(Value, tmpCount, 1) <> "." Then
   
    Tmp = Mid$(Value, tmpCount, 1) + Tmp
    tmpCount = tmpCount - 1
  
   Else
  
    If Tmp <> "" Then ExtractFileExt = Tmp
          
    Exit Function
   
   End If
  
 Next MainCount

End Function

Public Function ExtractMainFileName(Value As String) As String
'���ļ����л�ȡ���ļ���
    Dim i As Integer
    Dim intCount As Integer
    Dim Tmp As String
    
    intCount = Len(Value)
    
    For i = 0 To Len(Value)
        If Mid$(Value, intCount, 1) <> "." Then
            'Tmp = Mid$(Value, intCount, 1) & Tmp
            intCount = intCount - 1
            
        Else
            ExtractMainFileName = Left$(Value, intCount - 1)
            Exit Function
        
        End If
        
    Next
    
End Function



Public Function FileList(ByVal strPath As String, Optional ByVal FileExt As String = "*") As String()
'����:  strPath - �б��ļ���Ŀ¼
'       FileExt - �ļ���չ��,֧��*����������չ��,��Ŀ¼�µ�ȫ���ļ�
'����ֵ���ļ����б��ַ�����
    Dim strFileList() As String         'һ��Ŀ¼�µ��ļ����б�����

    Dim fso As New FileSystemObject
    Dim Folder1 As Folder
    
    Dim F As Files
    Dim F1 As File
    
    Dim intFileCount As Integer
    Dim i As Integer
    
    i = 0
    
    FileExt = LCase$(FileExt)   'ת��Сд
    
    
    Set Folder1 = fso.GetFolder(strPath)
    Set F = Folder1.Files
    
    intFileCount = F.Count
    If intFileCount > 0 Then ReDim strFileList(intFileCount - 1) Else ReDim strFileList(0)
    
    For Each F1 In F
'        If FileExt = "*" Then
'            Debug.Print F1.Name
'            strFileList(i) = F1.Name
'            i = i + 1
'        Else
'            If LCase$(ExtractFileExt(F1.Name)) = FileExt Then
'                Debug.Print F1.Name
'                strFileList(i) = F1.Name
'                i = i + 1
'            End If
'
'        End If
        
        If FileExt = "*" Or (FileExt <> "*" And LCase$(ExtractFileExt(F1.Name)) = FileExt) Then
'            Debug.Print F1.Name
            strFileList(i) = F1.Name
            i = i + 1
        
        End If
            
    Next
    
    If i > 0 Then ReDim Preserve strFileList(i - 1) Else ReDim strFileList(0)
    
    FileList = strFileList
    
    
End Function


