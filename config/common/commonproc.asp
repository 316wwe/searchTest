<%
'//�ۼ��� : gigatera
'//�ۼ��� : 2010-05-31
'//��   �� : ���� ���ν���/�Լ� ����


'-------------------------------------------------------------------------
'This is Procedure Section
'START
'-------------------------------------------------------------------------

Sub HistoryBack(ByVal Msg)
	'�ڹٽ�ũ��Ʈ�� Msg�� alert�� history.back�ϴ� ���ν���
	With Response
		.Write "<script language=""javascript"">" & chr(13)
		.Write "	alert("""&Msg&""");" & chr(13)
		.Write " history.back();" & chr(13)
		.Write "</script>" & chr(13)
	End With
End Sub


Sub initArrary(ByRef arr,  ByVal initVal)
	'�迭 �ʱ�ȭ ���ν���
	For i=0 To Ubound(arr) Step 1
		arr(i) = Trim(initVal)
	Next
End Sub

'-------------------------------------------------------------------------
'END
'This is Procedure Section
'-------------------------------------------------------------------------



'-------------------------------------------------------------------------
'This is Function Section
'START
'-------------------------------------------------------------------------

Public Function GetDate(ByVal dateOPT)
	
	'��¥�� ��ü/��/��/���� ������ �������� ���Ϲ��� �� �ִ� �Լ�
	
	Res = ""

	Select Case Trim(dateOPT)
		Case 0 '��¥�� 2003-01-27���� �������� ����
			Res = Date()
		Case 1 '2003������ �⸸ ����
			Res = Split(CStr(Date()),"-")(CInt(dateOPT)-1)
		Case 2 '01������ ���� ����
			Res = Split(CStr(Date()),"-")(CInt(dateOPT)-1)
		Case 3 '27������ �ϸ� ����
			Res = Split(CStr(Date()),"-")(CInt(dateOPT)-1)
		Case Else
			Res = ""
	End Select
	
	GetDate = Res

End Function



Public Function GetWeekDay(myDate) '���� ���ϱ�
	
	'������ �����ִ� �Լ�
	Res = ""
	Select Case WeekDay(myDate)
		Case 1
			Res = "��"
		Case 2
			Res = "��"
		Case 3
			Res = "ȭ"
		Case 4
			Res = "��"
		Case 5
			Res = "��"
		Case 6
			Res = "��"
		Case 7
			Res = "��"
		Case Else
			Res = "��"
	End Select

	GetWeekDay = Res

End Function




function getLastDayInMonth(i_year, i_month)

	'�Ѵ��� �ѳ�¥ ����Լ�
	Dim now_first_date            : now_first_date = i_year &"-"& RIGHT("0"& i_month,2) &"-01"
	Dim next_first_date            : next_first_date = DateAdd("m",1,now_first_date)
	Dim now_last_date            : now_last_date = DateAdd("d",-1,next_first_date)
	Dim now_month_days        : now_month_days = Day(now_last_date)

	getLastDayInMonth = now_month_days

end function



'====================================

'Description : ���ǽ����� �Ͽ����̶� �����Ѵ�

'====================================

 '-- �ش糯¥�� �ش��ϴ� ��¥�� �Ͽ��ϰ� ����ϳ�¥ ���ؿ�( ���� chkdate : 2009-01-01)
 FUNCTION week_day(chkdate)
  SELECT CASE weekday(CDate(chkdate))
   CASE 1 : temp1 = CDate(chkdate) - 0 : temp2 = CDate(chkdate) + 6 '��
   CASE 2 : temp1 = CDate(chkdate) - 1 : temp2 = CDate(chkdate) + 5 '��
   CASE 3 : temp1 = CDate(chkdate) - 2 : temp2 = CDate(chkdate) + 4 'ȭ
   CASE 4 : temp1 = CDate(chkdate) - 3 : temp2 = CDate(chkdate) + 3 '��
   CASE 5 : temp1 = CDate(chkdate) - 4 : temp2 = CDate(chkdate) + 2 '��
   CASE 6 : temp1 = CDate(chkdate) - 5 : temp2 = CDate(chkdate) + 1 '��

   CASE 7 : temp1 = CDate(chkdate) - 6 : temp2 = CDate(chkdate) + 0 '��
  END SELECT
  week_day = temp1 &"|"& temp2
 END FUNCTION

 

 '-- �ش���� ��������
 FUNCTION last_day(year,month)
 Dim temp
  temp = CDate(LEFT(dateadd("m",1,year &"-"& month &"-01"),7) &"-01")-1
  last_day = Split(temp,"-")(2)
 END FUNCTION
 
 '-- ������� ����( ���� getdate : 2009-01-01)
 FUNCTION now_week(getdate)
  now_week = int((Day(getdate) - weekday(getdate)+13)/7)
 END FUNCTION
 
 '-- �ش���� ���ޱ����� ������
 FUNCTION month_week(year, month)
  week = 0
  FOR i = 1 TO month - 1
   week = week + now_week(year &"-"& i &"-"& last_day(year,i))
  NEXT
  month_week = week
 END FUNCTION
 
 '-- �ش���� �������� �Ͽ���(�ȵ�;;)
 FUNCTION week_sunday(year,week)
  week_sunday = CDate(year &"-01-01")+(week-1)*7+1-WEEKDAY(CDate(year& "-01-01")+(week-1)*7)
 END FUNCTION

 

 '-- �ش���� �������� �Ͽ���(���� now_month_week("2009","02","2") : 2009��2��2����)

FUNCTION now_month_week(year,month,week_cnt)
  a = CDate(year &"-01-01") '-- ���۱�����
  b = Datepart("ww",year &"-"& month &"-01",1,3) '-- �ش���� �������� ����
  c = 7 * b '--7�� ������ ���ϸ� ���� �� ���� ����
  d = DateAdd("d",c,a) '--���۱����Ͽ��� ��¥�� c ��ŭ �����ش�
  '-- �ش���� �������� �Ͽ����̸� ù���̹Ƿ� ���ϴ��������� 7�� ���ش�.
  IF weekday(year &"-"& month &"-01") = 1 THEN
   d = CDate(d) + (7 * week_cnt) - 7
  ELSE
   d = CDate(d) + (7 * week_cnt)
  END IF
  now_month_week = d
 END FUNCTION






Function IsSet(ByVal Val)

	'������ ���� �����ϴ����� ���θ� �Ǵ����ִ� �Լ�

	Res = True
	If ( Trim(CStr(Val))="" Or IsNull(Trim(CStr(Val))) Or IsEmpty(Trim(CStr(Val))) ) Then
		Res = False
	End If

	IsSet = Res

End Function 



Function GetStrReplace(ByVal strVal,ByVal strLength)
'strVal ��Ʈ�� ������ strLength��ŭ �߶� �������ִ� �Լ�
'strLength �� 0�̸� strVal �ڸ��� �ʰ� ����
	Dim strRet

	If (CInt(strLength>0)) Then '0�̸� ��Ʈ�� �ڸ��� ����
		If (Len(strVal)>CInt(strLength)) Then
			strRet = Mid(strVal,1,strLength) & "..."
		Else
			strRet = strVal
		End If
	Else 
		strRet = strVal
	End If

	strRet = Replace(strRet,"'","''")

	GetStrReplace = strRet

End Function 



Function MailSender(Sender,Reciever,Cust_Name,Title,Body,Attach) '���� ����..
	'CDO2000��ü�� �̿��� ���� ������ �Լ�
	Dim iMsg, iConf, Flds
	Set iMsg  = CreateObject("CDO.Message")
	Set iConf = CreateObject("CDO.Configuration")
	Set Flds  = iConf.Fields 

	With Flds
		.Item(cdoSendUsing)				= 25
		.Item(cdoSMTPServer)			= "119.70.15310" '���� �ش� smtp�����Ƿ� ����
		.Item(cdoSMTPConnectionTimeout)	= 10
		.Item(cdoSMTPAuthenticate)		= cdoBasic
		.Item(cdoSendUserName)			= "gigatera"
		.Item(cdoSendPassword)			= "12341" 
		.Item(cdoURLGetLatestVersion)   = True
		.Update
	End With

	Set iMsg.Configuration = iConf   

	With iMsg        
		'.From			= Cust_Name & "<" & Sender & ">"
		'.To				= "��������" & "<" & Reciever & ">" 
		.From			= Sender
		.To				= Reciever
		.Subject		= Title    
		.HTMLBody       = Body   
		If (Attach<>"") Then
			.AddAttachment  Attach
		End If
		.Send     
	End With

End Function



Function ChContent(CheckValue, tag)
'�Խ��ǿ��� �۾��� �Ҷ� html���� ���ο� ���� ��ȯ���� �������ִ� �Լ�
	Dim Content
	If Cint(tag)=0 Then
		Content = Server.HTMLEncode(CheckValue)
		'Content = Replace(CheckValue,chr(13),"<br>")
	Else
		Content = Replace(CheckValue,chr(13),"<br>")
	End If
	
	ChContent = Content
End Function



Function GetFileSystemObject()
 '���Ͻý��� ������Ʈ�� ���´�
 '��ȯ�� SetFreeObj(fso)�� �Ѵ�
	Err.Clear 
	On Error Resume Next
		Set fso = CreateObject("Scripting.FileSystemObject")	
	If err.number <> 0 Then
		GetFileSystemObject = False
	Else
		GetFileSystemObject = True
	End If
End Function



Function GetTextFile(ByVal FilePath,ByVal IOMode,ByVal CreateMode)
'FilePath�� �־��� �ؽ�Ʈ ������ �о�´�
'IOMode >> ForReading = 1, ForWriting = 2, ForAppending = 8 
'CreateMode >> True, False
	Dim Path
	
	Err.Clear 
	On Error Resume Next
		Path = Server.MapPath(FilePath)
		Set fp = fso.OpenTextFile(Path,IOMode,CreateMode) 
		lpstr = fp.ReadAll()
	If err.number <> 0 Then
		GetTextFile = False
	Else
		GetTextFile = lpstr
	End If
End Function



Function decreaseTitle(title,size)
	'�Խ��� ������ �ʹ� �涧 ���������� + "..." ���� ��ü�ϴ� �Լ�
	Dim tmpStr : tmpStr = ""
	if (len(title)>size) then
		tmpStr = Left(title,size) & "..."
		'tmpStr = Left(title,size)
	else
		tmpStr = title
	end if
	decreaseTitle = tmpStr
End Function


Function CheckWord(cw)
	' ���� ��ȯ
	cw = replace(cw,"&","&amp;")
	cw = replace(cw,"<","&lt;")
	cw = replace(cw,">","&gt;")
	cw = replace(cw,chr(34),"&quot;")
	cw = replace(cw,"'","''")	
	cw = trim(cw)
	CheckWord = cw
End Function

Function GetFileExt(filename)
	'���� Ȯ���� ���ϱ�
	Dim Res : Res = ""

	if (filename<>"") then
		Res = Split(filename,".")(1)
	end if

	GetFileExt = Res
End Function

Function IsImage(ext) 
	'�̹�������
	Dim Res  : Res = false
	Dim ImgExts : ImgExts = Array("jpe","jpeg","jpg","gif","bmp","png")
	
	for i=0 to Ubound(ImgExts) step 1
		if trim(LCase(ext))=trim(ImgExts(i)) then
			Res = True
		end if
	Next
	IsImage = Res
End Function


function setThumbnail(w,h,path,filename)
	
	if GetFileExt(filename)<>"" then

		if IsImage(GetFileExt(filename)) then

			Set Image = Server.CreateObject("Nanumi.ImagePlus")
			'Response.Write "Here : " & Server.MapPath(path)
			Image.OpenImageFile Server.MapPath(path) & "\" & filename
			Image.ChangeSize w, h
			Image.SaveFile Server.MapPath(path) & "\thumb\" & filename
			Image.Dispose
			Set Image = Nothing

		end if

	end if

end function


function setThumbnail2(w,h,path,filename)
	
	if GetFileExt(filename)<>"" then

		if IsImage(GetFileExt(filename)) then

			Set Image = Server.CreateObject("Nanumi.ImagePlus")
			Image.OpenImageFile Server.MapPath(path) & "\" & filename
			Image.KeepAspect = False
			Image.ChangeSize w, h
			Image.SaveFile Server.MapPath(path) & "\thumb\" & filename
			Image.Dispose
			Set Image = Nothing

		end if

	end if

end Function



Function IsMov(ext) 
	'�̹�������
	Dim Res  : Res = false
	Dim MovExts : MovExts = Array("mov","wmv","asf","mpg","mpe","mpeg","avi")
	
	for i=0 to Ubound(MovExts) step 1
		if trim(LCase(ext))=trim(MovExts(i)) then
			Res = True
		end if
	Next
	IsMov = Res
End Function

Function IsSnd(ext) 
	'�̹�������
	Dim Res  : Res = false
	Dim SndExts : SndExts = Array("wma","mp3","wav")
	
	for i=0 to Ubound(SndExts) step 1
		if trim(LCase(ext))=trim(SndExts(i)) then
			Res = True
		end if
	Next
	IsSnd = Res
End Function


Function convertDT(dates) 
	'��¥ ���̱׷��̼ǽ� �Է°����� ��¥ ������ ��ȯ���ִ� �Լ�	
	'yyyy-mm-dd hh:mm:ss �������� ��ȯ�� ����
	Dim reg_date
	reg_date = trim(dates)
	reg_date = Replace(reg_date,"���� 12","00")
	reg_date = Replace(reg_date,"���� 1","1")
	reg_date = Replace(reg_date,"���� 2","2")
	reg_date = Replace(reg_date,"���� 3","3")
	reg_date = Replace(reg_date,"���� 4","4")
	reg_date = Replace(reg_date,"���� 5","5")
	reg_date = Replace(reg_date,"���� 6","6")
	reg_date = Replace(reg_date,"���� 7","7")
	reg_date = Replace(reg_date,"���� 8","8")
	reg_date = Replace(reg_date,"���� 9","9")
	reg_date = Replace(reg_date,"���� 10","10")
	reg_date = Replace(reg_date,"���� 11","11")
	
	reg_date = Replace(reg_date,"���� 1","13")
	reg_date = Replace(reg_date,"���� 2","14")
	reg_date = Replace(reg_date,"���� 3","15")
	reg_date = Replace(reg_date,"���� 4","16")
	reg_date = Replace(reg_date,"���� 5","17")
	reg_date = Replace(reg_date,"���� 6","18")
	reg_date = Replace(reg_date,"���� 7","19")
	reg_date = Replace(reg_date,"���� 8","20")
	reg_date = Replace(reg_date,"���� 9","21")
	reg_date = Replace(reg_date,"���� 10","22") : reg_date = Replace(reg_date,"130","22")
	reg_date = Replace(reg_date,"���� 11","23") : reg_date = Replace(reg_date,"131","23")
	reg_date = Replace(reg_date,"���� 12","12") : reg_date = Replace(reg_date,"132","12")
	
	convertDT = reg_date
End Function

Function getGUID() '������ guid �� ���
  Dim tmpTemp
  tmpTemp = Right(String(4,48) & Year(Now()),4)
  tmpTemp = tmpTemp & Right(String(4,48) & Month(Now()),2)
  tmpTemp = tmpTemp & Right(String(4,48) & Day(Now()),2)
  tmpTemp = tmpTemp & Right(String(4,48) & Hour(Now()),2)
  tmpTemp = tmpTemp & Right(String(4,48) & Minute(Now()),2)
  tmpTemp = tmpTemp & Right(String(4,48) & Second(Now()),2)
  getGUID = tmpTemp
End Function

'-------------------------------------------------------------------------
'END
'This is Function Section
'-------------------------------------------------------------------------


Function IsFlash(ext) 
	'�̹�������
	Dim Res  : Res = false
	Dim FlashExts : FlashExts = Array("swf")
	
	for i=0 to Ubound(FlashExts) step 1
		if trim(LCase(ext))=trim(FlashExts(i)) then
			Res = True
		end if
	Next
	IsFlash = Res
End Function

Function GetAspUploadObject(mus)
 '���Ͻý��� ������Ʈ�� ���´�
 '��ȯ�� SetFreeObj(fso)�� �Ѵ�
	Err.Clear 
	On Error Resume Next
		Set theForm = Server.CreateObject("ABCUpload4.XForm")

		theForm.MaxUploadSize = mus * 1024 * 1024
		theForm.AbsolutePath = true     ' ���ε�� ������ ���� ��θ� ����Ѵ�.
		theForm.CodePage = 949			' ���ε�� �ѱ��� �����Ѵ�.
		theForm.Overwrite = false

	If err.number <> 0 Then
		GetAspUploadObject = False
	Else
		GetAspUploadObject = True
	End If
End Function


Function SetAspUploadPath(path)
	
	uploadPath = path
	'Response.Write "path : " & path & "<br>"
	
	if not fso.folderexists(server.mappath(Path)) then 	
		fso.createfolder(server.mappath(Path)) '�ش� �Խ��� ���� ���ε�� ���丮 ����
	end if
End Function


function SetAspUploadOk(ByRef theField, ByVal upPath)
	
	FileName = ""
	if Len(theField.SafeFileName) > 0 then
		bExist = true
		countFileName = 0
		While bExist
			FileName = getGUID & "_" & countFileName & "." & theField.FileType
			saveFileName = Server.MapPath(upPath) & "\" & getGUID & "_" & countFileName & "." & theField.FileType
			If (not Fso.FileExists(saveFileName)) Then
				bExist = False
			else
				countFileName = countFileName  + 1
			End If
		Wend

		theField.Save saveFileName
	End if
	
	SetAspUploadOk = FileName

End function


Public Function IsValue(ByVal Val)
'������ ���� �����ϴ����� ���θ� �Ǵ����ִ� �Լ�
	Res = True
	If ( Trim(CStr(Val))="" Or IsNull(Trim(CStr(Val))) Or IsEmpty(Trim(CStr(Val))) ) Then
		Res = False
	End If
	IsValue = Res
End Function 


Function saveQry(ByRef qry)
	' ���� ���� �α� / update, delete, insert �� �����Ѵ�...
	Dim logRes : logRes = false

	if trim(Request.Cookies("koreaart_uid"))<>"" and trim(Request.ServerVariables("REMOTE_HOST"))<>"" and trim(qry)<>"" then
		Dim logConn,logQry

		Set logConn = Server.CreateObject("Adodb.Connection")
		logConn.CursorLocation = 3
		logConn.Open("provider=sqloledb;server=sql8ssd-011.localnet.kr;uid=korea3200_hsart;pwd=koreabkbosory6;database=korea3200_hsart")
		
		logQry = ""
		logQry = "insert into tbl_log_intranet_qry(uid,ip,qry) values ('"& trim(Request.Cookies("koreaart_uid")) &"','"& trim(Request.ServerVariables("REMOTE_HOST")) &"','"& Replace(trim(qry),"'","''") &"');"
		'Response.Write logQry
		
		err.Clear
		logConn.BeginTrans
		On error resume next
			logConn.Execute(logQry)
		if err.number <> 0 then
			logConn.RollBackTrans
			logRes = false
		else
			logConn.CommitTrans
			logRes = true
		end if

		if not logConn is nothing then set logConn = nothing
	end if

	saveQry = logRes

End Function


Sub die(byval msg)
	Response.Write "<table border='0' width='100%' height='1000'><tr><td align='center' valign='middle' class='text_msg'>"
	Response.Write msg
	'Response.Write "<br>2���� �ڷ� �̵��մϴ�"
	Response.Write "</td></tr></table>"
	Response.Write "<script language='javascript'> "
	Response.Write "	function goBack() { "
	Response.Write "		history.back(); "
	Response.Write "	} "
	Response.Write "	setTimeout('goBack()',2000); "
	Response.Write "</script> "
	Response.End
end sub

Sub msg(byval txts)
	Response.Write "<table border='0' width='100%' height='1000'><tr><td align='center' valign='middle' class='text_msg'>"
	Response.Write txts
	Response.Write "</td></tr></table>"
end sub


function Req(ByVal QStr)

	Dim ret : ret = trim(QStr)

	ret = Replace(ret,"'","''")
	ret = Replace(ret,"create","�ϨިѨͨ��")
	ret = Replace(ret,"insert","�ըڨߨѨި�")
	ret = Replace(ret,"drop","�Шިۨ�")
	ret = Replace(ret,"pangolin","�ܨͨڨӨۨبը�")
	ret = Replace(ret,"unicode","��ڨըϨۨШ�")
	ret = Replace(ret,"substr","�ߨ�Ψߨ��")
	ret = Replace(ret,"char","�ϨԨͨ�")	
	ret = Replace(ret,"xp_","���_")

	Req = ret 

end function


 '------------------------HtmlTagRemover -- HTML �ױ� ���� �Լ� -------by Andy---------
 ' �Ķ���� ���� : (ó���ҹ��ڿ�, �ڸ�����)
 ' cutlen = 0 �ϰ�� ��ü ���ڿ�
 '---------------------------------------------------------------------------------------
 function HtmlTagRemover(content, cutlen)
	Dim tmpb,length,htmlRemovedContent
  j=1
  tmpb=2
  length = len(content)
  htmlRemovedContent = content

  Do while length > 0
   k = mid(htmlRemovedContent,j,1)

   if k="<" then
    tmpb = 0
   elseif k = ">" then
    tmpb = 1
   end if

   if tmpb = 0 then
    htmlRemovedContent = left(htmlRemovedContent,j-1) & mid(htmlRemovedContent,j+1)
   elseif tmpb = 1 then
    htmlRemovedContent = left(htmlRemovedContent,j-1) & mid(htmlRemovedContent,j+1)
    tmpb = 2
   else
    j=j+1
   end if
 
   length = length -1
  loop
   
  if cutlen <> 0 then
   htmlRemovedContent = left(htmlRemovedContent, cutlen)
  end if

  HtmlTagRemover = htmlRemovedContent

 end function
 '------------------------HtmlTagRemover -- HTML �ױ� ���� �Լ� -------by Andy----------

' response.write HtmlTagRemover("ABCDEF<img src='/ZYXWVUTSRQPO/'>GHIJKL",20)




function toHanDate(ByVal engdate)

	Dim ret : ret = ""
	Dim tmpDate

	if trim(engdate)="" then
		ret = ""
	else
		tmpDate = split(engdate,"-")
		ret = trim(tmpDate(0)) & "�� " & trim(tmpDate(1)) & "�� " & trim(tmpDate(2)) & "�� " 
	end if

	toHanDate = ret

end function



function setAccountPrint(acc)

	Dim prt_acc : prt_acc = ""
	
	if ( trim(acc)<>"" ) then

		if ( len(trim(acc))=14 ) then

			prt_acc = Mid(acc,1,3)
			prt_acc = prt_acc & "-"
			prt_acc = prt_acc & Mid(acc,4,6)
			prt_acc = prt_acc & "-"
			prt_acc = prt_acc & Mid(acc,10,2)
			prt_acc = prt_acc & "-"
			prt_acc = prt_acc & Mid(acc,12,3)

		end if

	end if


	setAccountPrint = prt_acc

end function


Function htmlToEncode(str)
  str = Replace(Trim(str),"'","&acute;")
  str = Replace(str,"""","&quot;")
  str = Replace(str,"<","&lt;")
  str = Replace(str,">","&gt;")  
  htmlToEncode = str
End Function

Function htmlToDecode(str)  
  str = Replace(Trim(str),"&acute;","'")  
  str = Replace(Trim(str),"&quot;","""")
  str = Replace(str,"&lt;","<")
  str = Replace(str,"&gt;",">")
  htmlToDecode = str
End Function


'===============================================
'���ڸ޽��� ȣ�� �⺻ �Լ�(ȣ������ ������ �߼�)
'===============================================
Function SendMsg(sender, receiver, content)
	Dim data
	DIM URL
	DIM xmlhttp
	Dim result

	URL = "http://www.koreaart.ac.kr/intranet/fresh/sms/hp_ok.asp"
	data = "ispop=0&sender=" & sender & "&receiver=" & receiver & "&contents=" & Server.URLEncode(content)

	SET xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")

	xmlhttp.open "POST", URL, False
	xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xmlhttp.send data

	result = xmlhttp.responsetext

	SET xmlhttp = nothing

	SendMsg = result
End Function



%>