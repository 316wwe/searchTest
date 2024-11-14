<!--METADATA TYPE= "typelib"  NAME= "ADODB Type Library"  FILE="C:\Program Files\Common Files\SYSTEM\ADO\msado15.dll"  -->
<OBJECT RUNAT=server PROGID=ADODB.Connection id=Conn></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Command id=Cmd></OBJECT>
<%

Dim DBSTR_Execute, DBSTR_DataShape, DB_SERVER, DB_Name, DB_User, DB_PWD
Dim SMS_DBSTR_Execute, SMS_DBSTR_DataShape, SMS_DB_SERVER, SMS_DB_Name, SMS_DB_User, SMS_DB_PWD
Dim HS_DBSTR_Execute, HS_DBSTR_DataShape, HS_DB_SERVER, HS_DB_Name, HS_DB_User, HS_DB_PWD

DB_SERVER = "1.209.150.20"
DB_Name = "hoseo"
DB_User = "hoseo"
DB_PWD = "Tjdnfghtj@$366o"


DBSTR_Execute = "Driver={SQL Server};Server=" & DB_SERVER & ";Database=" & DB_Name & ";UID=" & DB_User & ";pwd=" & DB_PWD & ";"

'//기본 테이터 연결용
Class Database
	Private strConn
	
	Private Sub Class_Initialize
		strConn = DBSTR_Execute
	End Sub
	
	Public Sub ConnOpen()
		Conn.Open strConn
		Cmd.ActiveConnection = Conn
	END Sub

	Private Sub Class_Terminate
		IF Conn.State Then Conn.Close
	END Sub
End Class

SMS_DB_SERVER = "222.108.174.182"
SMS_DB_Name = "sms"
SMS_DB_User = "sms"
SMS_DB_PWD =  "art3200!%!%"

SMS_DBSTR_Execute = "Driver={SQL Server};Server=" & SMS_DB_SERVER & ";Database=" & SMS_DB_Name & ";UID=" & SMS_DB_User & ";pwd=" & SMS_DB_PWD & ";"

'//기본 테이터 연결용
Class SMS_Database
	Private strConn
	
	Private Sub Class_Initialize
		strConn = SMS_DBSTR_Execute
	End Sub
	
	Public Sub ConnOpen()
		Conn.Open strConn
		Cmd.ActiveConnection = Conn
	END Sub

	Private Sub Class_Terminate
		IF Conn.State Then Conn.Close
	END Sub
End Class


HS_DB_SERVER = "1.209.150.5"
HS_DB_Name = "hoseo"
HS_DB_User = "hoseo"
HS_DB_PWD = "hoseobkbosory6"

HS_DBSTR_Execute = "Driver={SQL Server};Server=" & HS_DB_SERVER & ";Database=" & HS_DB_Name & ";UID=" & HS_DB_User & ";pwd=" & HS_DB_PWD & ";"

'//기본 테이터 연결용
Class HS_Database
	Private strConn
	
	Private Sub Class_Initialize
		strConn = HS_DBSTR_Execute
	End Sub
	
	Public Sub ConnOpen()
		Conn.Open strConn
		Cmd.ActiveConnection = Conn
	END Sub

	Private Sub Class_Terminate
		IF Conn.State Then Conn.Close
	END Sub
End Class
%>