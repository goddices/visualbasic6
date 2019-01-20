<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>苏武牧羊张江点</title>
<%

dim r
dim rtxt2
dim a 
dim text
Dim fso, f1, ts, s
Const ForReading = 1

r=request.QueryString("rtxt1")
Set fso = CreateObject("Scripting.FileSystemObject")


if r<>"" then 
	rtxt2=rtxt2 & r 
	a=server.mappath("../")
	a= a & "\wwwroot\"
	Set f1 = fso.OpenTextFile( a & "testfile.txt", 1,True)	  
	s = f1.ReadLine
	Set f1 = fso.OpenTextFile( a & "testfile.txt", 2,True)	
	 f1.WriteLine s & rtxt2
	 f1.WriteBlankLines(1)
	 f1.Close
end if
%>
</head>

<body>
 
</body>
</html>
