<%
'//�ۼ��� : gigatera(gigatera@gigatera.co.kr)
'//�ۼ��� : 2010-05-31
'//��  �� : ����Ÿ���̽� ���� �Լ� ����

Public Function GetDbConn() 
'MS-SQL ��� ������ ���� �Լ�
'oConn�̶� ������ ��� ����ȴ�
	ConnStr = "provider=sqloledb;server=1.209.150.20;uid=hoseo;pwd=Tjdnfghtj@$366o;database=hoseo"
	Err.Clear 
	On Error Resume Next
		Set oConn = Server.CreateObject("Adodb.Connection")
		oConn.CursorLocation = 3 'AdUseClient
		oConn.Open(ConnStr)
	If err.number <> 0 Then
		GetDbConn = False
	Else
		GetDbConn = True
	End If
End Function

Public Function GetRs()
'���ڵ���� ���´�
	Err.Clear
	On Error Resume Next
		Set oRs = Server.CreateObject("Adodb.RecordSet")
	If err.number  <> 0 Then
		GetRs = False
	Else
		GetRs = True
	End If
End Function

Public Function GetRs2()
'���ڵ���� ���´�
	Err.Clear
	On Error Resume Next
		Set oRs2 = Server.CreateObject("Adodb.RecordSet")
	If err.number  <> 0 Then
		GetRs2 = False
	Else
		GetRs2 = True
	End If
End Function

Public Function GetRs3()
'���ڵ���� ���´�
	Err.Clear
	On Error Resume Next
		Set oRs3 = Server.CreateObject("Adodb.RecordSet")
	If err.number  <> 0 Then
		GetRs3 = False
	Else
		GetRs3 = True
	End If
End Function

Public Sub SetFreeObj(ByRef obj)  
'��ü ��Ʈ����
'��ü�� �޸𸮿��� �����ִ� �Լ�
	If Not obj Is Nothing Then
		Set obj = Nothing
	End If
End Sub





'���� ���� ���
Public Function GetDbConn2() 
	ConnStr = "provider=sqloledb;server=119.206.205.35;uid=Syma;pwd=kyunkk1100;database=Yewonmusic"
	Err.Clear 
	On Error Resume Next
		Set oConn2 = Server.CreateObject("Adodb.Connection")
		oConn2.CursorLocation = 3 'AdUseClient
		oConn2.Open(ConnStr)
	If err.number <> 0 Then
		GetDbConn2 = False
	Else
		GetDbConn2 = True
	End If
End Function

'MMS ���
Public Function GetDbConnM() 
	ConnStr = "provider=sqloledb;server=1.209.150.5;uid=mms;pwd=mmsbkbosory6;database=mms"
	Err.Clear 
	On Error Resume Next
		Set oConnM = Server.CreateObject("Adodb.Connection")
		oConnM.CursorLocation = 3 'AdUseClient
		oConnM.Open(ConnStr)
	If err.number <> 0 Then
		GetDbConnM = False
	Else
		GetDbConnM = True
	End If
End Function


'HOSEO ���
Public Function GetDbConnH() 
	ConnStr = "provider=sqloledb;server=1.209.150.5;uid=hoseo;pwd=hoseobkbosory6;database=hoseo"
	Err.Clear 
	On Error Resume Next
		Set oConnH = Server.CreateObject("Adodb.Connection")
		oConnH.CursorLocation = 3 'AdUseClient
		oConnH.Open(ConnStr)
	If err.number <> 0 Then
		GetDbConnH = False
	Else
		GetDbConnH = True
	End If
End Function
%>