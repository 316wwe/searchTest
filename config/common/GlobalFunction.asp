<% 
'======================  �α��� ���� �Լ� ���� ====================================
'//�α��� äũ �Լ� �� : Call LoginCheck(���ư� ������, �����ڵ�)
 Sub LoginCheck(GoURL)	
 
	Dim cryptDll

	IF ckUserID = "" THEN
		GoURL = replace(GoURL,"&","$")
		response.redirect "/login/loginForm.asp?GoURL=" & GoURL
		response.end()
	END IF
	
End Sub


'//�α��� äũ �Լ� �� : Call LoginCheck(���ư� ������)
 Sub AdminLoginCheck(GoURL)
	Dim cryptDll
 	IF ckAdminID = "" THEN
		GoURL = replace(GoURL,"&","$")
		ScriptFunction("top.location.href = '/shopManage/login/loginForm.asp?GoURL=" & GoURL & "'")
	End IF
End Sub


'//�˾����������� �α��� äũ �Լ� �� : Call PopLoginCheck(javascirpt)
Sub PopLoginCheck(Message)									
	if ckUserid = "" then
		Response.Write("<html>")
		Response.Write("<head>")
		Response.Write("<meta http-equiv=""Content-Type"" content=""text/html; charset=euc-kr"">")
		Response.Write("</head>")
		Response.Write("<body bgcolor=#FFFFFF>")
		Response.Write("<SCRIPT LANGUAGE=""JavaScript"">")
		Response.Write(Message)
		Response.Write("</SCRIPT>")
		Response.Write("</body>")	
		Response.Write("</html>")
		response.end
	end if
End Sub

'======================  �α��� ���� �Լ� �� ====================================

'��ũ��Ʈ �޼��� ó�� �Լ�
'��뿹 : Call ScriptFunction("alert('â�� �ݰڽ��ϴ�.'); this.close();")

Sub ScriptFunction(Message)
	Response.Write("<html>")
	Response.Write("<head>")
	Response.Write("<meta http-equiv=""Content-Type"" content=""text/html; charset=euc-kr"">")
	Response.Write("</head>")
	Response.Write("<body bgcolor=#FFFFFF>")
	Response.Write("<SCRIPT LANGUAGE=""JavaScript"">")
	Response.Write(Message)
	Response.Write("</SCRIPT>")
	Response.Write("</body>")	
	Response.Write("</html>")
	response.end
End Sub

'���� ��ȯ
Public Function  dbTOweb(CheckValue)
	CheckValue = replace(CheckValue, "&" , "&amp;")
	CheckValue = replace(CheckValue, "<", "&lt;")
	CheckValue = replace(CheckValue, ">", "&gt;")
	CheckValue = replace(CheckValue, "'", "&quot;")
	CheckValue = replace(CheckValue, chr(13)&chr(10), "<br>")
	dbTOweb = CheckValue
End Function


'======================  ��Ű ���� �Լ� ���� ====================================

'//*********************************************
'// sName : ��Ű �̸� 
'// sVal  : ��Ű��
'// dExp  : ���ѳ�¥.
'//*********************************************


'// ���� ��Ű ����
Sub SetCookies(sName, sVal, dExp)
	Response.Cookies(sName)  = sVal

	'// ��Ű ���� ������ ����
	Response.Cookies(sName).domain  = SITEDOMAIN
	Response.Cookies(sName).path = "/"
	If  Trim( dExp ) <> "" then
		Response.Cookies(sName).expires = dExp
	End if	
End Sub 

'// ��Ű ���ֱ�
Sub DeleteCookies(sName)
	Response.Cookies(sName)  = ""
	Response.Cookies(sName).domain  = SITEDOMAIN
	Response.Cookies(sName).path = "/"
	Response.Cookies(sName).expires = DateAdd("d", -1, Now())
End sub

''// ���ƿ� �� ����   
'Sub SetCookUrl(sUrl)
'	DeleteCookies "JsGoURL"
'	SetCookies "GoURL", sUrl, ""
'End Sub


'// ��Ű�� ��������
Function GetCookies(sName)
	GetCookies = Request.Cookies(sName)
End Function


'======================  ��Ű ���� �Լ� �� ====================================



'======================  ����¡ ���� �Լ� ���� ====================================
'//���� ����¡
'function GoPage(SearchType, SearchKeyword, Sort_Type, Page)
'{
'	location.href = "/Mynetwork/MyPage_ProPoseList.asp?Page=" + Page + "&SearchType=" + SearchType + "&SearchKeyword=" + SearchKeyword + "&Sort_Type=" + Sort_Type;
'}
'Call Paging("GoPage('" & SearchType & "','" & SearchKeyword & "','" & Sort_Type & "','")
'//

Sub Paging(Query, Page, PageCount)
	Dim startpage, currentpage
	if (Page mod 10) = 0 then
		startpage = Page - 9
	else
		startpage = fix(Page/10) * 10 + 1
	end if

	if startpage > 10 then
	    	response.write "<a href=""javascript:" & Query  & startpage - 1 & "');""><img src=""/images/common/icon_list_prev.gif"" border=""0""></a>&nbsp;&nbsp;"
	end if
	'response.write "["
	for currentpage = startpage to startpage + 9
		if currentpage > PageCount then
			exit for
		elseif currentpage <> startpage then
			response.write "&nbsp;"
		end if
		if Cint(currentpage) = Cint(Page) then
			response.write " <span style=font-family:Gulim,verdana,arial,helvetica; font-size:9pt; color:rgb(118,118,118); line-height:150%; text-decoration:none;><strong><a href=""javascript:" & Query & currentpage & "');"">" & currentpage & "</a></strong></font>&nbsp;"
		else
			response.write " <a href=""javascript:" & Query & currentpage & "');"" class=""main"">" & currentpage & "</a>&nbsp;"
		end if

	next
	'response.write "]"
	if (startpage + 9) < PageCount then
	      	response.write " /<b> total [<a href=""javascript:" & Query & PageCount & "');"">"&PageCount&"</a>]</b>&nbsp;&nbsp;<a href=""javascript:" & Query & startpage + 10 & "');""><img src=""/images/common/icon_list_next.gif"" border=""0""></a>"
	end if
End Sub


'���� ����¡2 (2�� �����󿡼� �����)
Sub Paging2(Query, Page, PageCount)
	Response.write "<center>"
	Dim startpage, currentpage
	if (Page mod 10) = 0 then
		startpage = Page - 9
	else
		startpage = fix(Page/10) * 10 + 1
	end if


	if startpage > 10 then
	    	response.write "<a href=""javascript:" & Query  & startpage - 1 & "');""><img src=""/images/common/prev.gif"" border=""0""></a>&nbsp;&nbsp;"
	end if

	for currentpage = startpage to startpage + 9

		if Cint(currentpage) > Cint(PageCount) then
			exit for
		elseif Cint(currentpage) <> Cint(startpage) then
			response.write "&nbsp;"
		end if
		if Cint(currentpage) = Cint(Page) then
			response.write " <a href=""javascript:" & Query & currentpage & "');""><b><font color=""#FF3300"">" & currentpage & "</b></font></a>&nbsp;"
		else
			response.write " <a href=""javascript:" & Query & currentpage & "');"" class=""main"">" & currentpage & "</a>&nbsp;"
		end if

	next
	'response.write "]"
	if (startpage + 9) < Cint(PageCount) then
	      	response.write " /<b> total [<a href=""javascript:" & Query & PageCount & "');"">"&PageCount&"</a>]</b>&nbsp;&nbsp;<a href=""javascript:" & Query & startpage + 10 & "');""><img src=""/images/common/next.gif"" border=""0""></a>"
	end if
	Response.write "</center>"
End Sub

'======================  ����¡ ���� �Լ� �� ====================================

'//=======================	�Է»����� DB�� �ֱ� ���� Ư�������� ��ȯ 
Function ChkValToDB( m_replVal )
'	m_replVal = Replace(m_replVal, "&nbsp;", chr(32))
'	m_replVal = Replace(m_replVal, "<br>", chr(13))
	m_replVal = Replace(m_replVal, "&" , "&amp;")
	m_replVal = Replace(m_replVal, "<", "&lt;")
	m_replVal = Replace(m_replVal, ">", "&gt;")
'	m_replVal = Replace(m_replVal, "'", "''")
	ChkValToDB = m_replVal
End Function

'//=======================	DB���� �Է»����� �ѷ��ٶ�....��ȯ
Function ChkDBToVal( m_replDB )
	m_replDB = Replace(m_replDB, "&" , "&amp;")
	m_replDB = Replace(m_replDB, "&lt;", "<")
	m_replDB = Replace(m_replDB, "&gt;", ">")	
	m_replDB = Replace(m_replDB, "''", "'")
	m_replDB = Replace(m_replDB, """", "'")
	m_replDB = Replace(m_replDB,"&amp;", "&" )
	m_replDB = Replace(m_replDB, chr(32), "&nbsp;" )
	m_replDB = Replace(m_replDB, chr(13), "<br>")
	ChkDBToVal = m_replDB
End Function

'//======================= ����Ʈ �ѷ��ٶ� ���ڿ� ���� üũ 
Function strLeft(strVal, strCnt)
	dim strNewTitle, intRTitle, intSeq, intTitleCnt, intK
	strNewTitle = strVal
	intRTitle = 0
	intSeq = 0
	intTitleCnt = strCnt'18 - (2*arr_List(8,loop_int))
																		
	if len(strNewTitle) > intTitleCnt then 
	    for intK = 1 to len(strNewTitle)
	        if (intRTitle > (intTitleCnt * 2) or intSeq > (intTitleCnt * 2)) then exit for   

	        if (asc(mid(strNewTitle,intK,1)) >= 0 and asc(mid(strNewTitle,intK,1)) <=255) then 
	            intRTitle = intRTitle + 1 
	            intSeq = intSeq + 1
	        else 
				intRTitle = intRTitle + 2  
	        end if
	    next
	    if (intRTitle > (intTitleCnt * 2)) then 
	        if (intSeq mod 2 = 0) then  
				strLeft = Leftb(strNewTitle,(intTitleCnt * 2) + intSeq) & ".."
	        else  
	            strLeft = Leftb(strNewTitle,(intTitleCnt * 2) + intSeq - 1) & ".."
	        end if
	    else 
			strLeft = strNewTitle
	    end if
	else
		strLeft = strNewTitle
	end if
End Function

'=======================�ѱ�/���� ���� üũ �Լ�	============================
Public Function HLen(str)
	Dim i , chlen 
	
	For i = 1 To Len(str)
		If Asc(Mid(str, i, 1)) < 0 Then
			chlen = chlen + 2
		Else
			chlen = chlen + 1
		End If
	Next 
	
	HLen = chlen
End Function



'======================= Purpose : ��¥���˺�ȯ =======================
Function MakeDate8(strDate)
	Dim temp_strDate
	temp_strDate	=	Left(strDate,4) & "-" & Mid(strDate,5,2) & "-" & Mid(strDate,7,2)
	MakeDate8		=	temp_strDate
End Function


'============ Purpose : ���糯¥�� yyyyMMddhhmmss �������� ��ȯ =====
Function MakeDate14()	
	Dim curYear
	Dim curMonth
	Dim curDate
	Dim curHour
	Dim curMin
	Dim curSec
	
	curYear = Year(now)
	curMonth = Month(now)
	If Len(curMonth)<2 Then
		curMonth = "0" & curMonth
	End If
	curDate = Day(now)
	If Len(curDate)<2 Then
		curDate = "0" & curDate
	End If
	curHour = Hour(now)
	If Len(curHour)<2 Then
		curHour = "0" & curHour
	End If
	curMin = Minute(now)
	If Len(curMin)<2 Then
		curMin = "0" & curMin
	End If
	curSec = Second(now)
	If Len(curSec)<2 Then
		curSec = "0" & curSec
	End If
	
	MakeDate14 = curYear & curMonth & curDate & curHour & curMin & curSec
End Function



'======================= �� �������� �Ѿ�� ���� byte�� ���� äũ =============
Function TextLenCheck(TXT, TLen)
	if HLen(TXT) > TLen OR Len(Trim(TXT)) = 0 then
		TextLenCheck = False 
	else
		TextLenCheck = True
	end if
End Function

'======================= �޴��� ��ȣ ������ �´��� Ȯ�� ===================
Function IsPhoneNumber(phoneNum)	
	dim strDst	

	strDst = Left(phoneNum,3)

	if strDst = "010" or strDst = "011" or strDst = "016" or strDst = "018"  or strDst = "017" or strDst = "019" then
       IsPhoneNumber = True
	else
	   IsPhoneNumber = False
	end if

	if  Len(phoneNum) <> 11 then
	   IsPhoneNumber = False
	   
	end if	
End Function


'======================= �ֹι�ȣ 13�ڸ��� ����� ===================
Function MakeJumin13(jumin)
	jumin = Replace(jumin, " ", "")	'������ ������ ���ְ�
	jumin = Replace(jumin, "-", "")	'�뽬(-) �� ������ ���ְ�
	
	MakeJumin13 = jumin
End Function

'======================= �ֹ� ��ȣ - ���̱� (�Է����� : 7312151512456 => ��ȯ: 731215-1312456) ===================
Function MakeJumin14(jumin)
	MakeJumin14 = Left(jumin,6) & "-" & Right(jumin,7)
End Function


'======================= �ֹ� ��ȣ - ���̰� Remarking (�Է����� : 7312151512456 => ��ȯ: 731215-*******) =========
Function MakeJuminRemarking14(jumin)
	MakeJuminRemarking14 = Left(jumin,6) & "-" & "*******"
End Function




Public Function showListControl(Byval pURL, Byval lngRecordCount, Byval iPageSize, Byval iPageNo)
                                
	Dim iPrevPageNo, iNextPageNo, iLastPageNo, iPrtLastPgNo, iPrtFirstPgNo, i, sHTML
	Dim iBlockPage, iTemp, iLoop

	'��ϻ�����
	iBlockPage = 10

	' ������ ������ ���
	If lngRecordCount > 1 Then
		iLastPageNo = (lngRecordCount / iPageSize) 
		iLastPageNo = Round(iLastPageNo + 0.41)
	Else
		iLastPageNo = 1
	End If
	
		
	sHTML = ""

	sHTML = sHTML & "<table width='336' height='11' align='center' border='0' cellpadding='0' cellspacing='0'>"
	sHTML = sHTML & "	<tr>"
	sHTML = sHTML & "		<td align='center'>"

	If iPageNo>1 Then
		sHTML = sHTML & "			<a href='" & pURL & "'><img src='/images/life/starmovie_listbtn_first.gif' border='0' align='absmiddle'></a><img src='/images/life/null.gif' width='3' height='1'>"
	Else
		sHTML = sHTML & "			<img src='/images/life/starmovie_listbtn_first.gif' border='0' align='absmiddle'><img src='/images/life/null.gif' width='3' height='1'>"
	End If
	
	iTemp = Int((iPageNo - 1) / iBlockPage) * iBlockPage + 1


	If iTemp = 1 Then
		sHTML = sHTML & "<img src='/images/life/null.gif' width='3' height='1'><img src='/images/life/starmovie_listbtn_prev.gif' align='absmiddle'>"
	Else 
		sHTML = sHTML & "<a href='" & pURL & "&PageNo=" & iTemp - iBlockPage & "'><img src='/images/life/null.gif' width='3' height='1'><img src='/images/life/starmovie_listbtn_prev.gif' align='absmiddle'></a>"
	End If


	sHTML = sHTML & "			<img src='/images/life/null.gif' width='13' height='1'>"
	
    iLoop = 1
	
	Do Until iLoop > iBlockPage Or iTemp > iLastPageNo
        If iTemp = CInt(iPageNo) Then
            sHTML = sHTML & "<span><font color=#ff0000>" & iTemp & "</font></span>"
        Else
            sHTML = sHTML & " <a href='" & pURL & "&PageNo=" & iTemp & "'>" & iTemp & "</a> "
        End If

		If iTemp<>iLastPageNo then	
			sHTML = sHTML & " | " 
		End If
		
        iTemp = iTemp + 1
        iLoop = iLoop + 1
    Loop

	sHTML = sHTML & "			<img src='/images/life/null.gif' width='13' height='1'>"


	If iTemp > iLastPageNo Then
		sHTML = sHTML & "			<img src='/images/life/starmovie_listbtn_next.gif' align='absmiddle'><img src='/images/life/null.gif' width='3' height='1'>"
	Else
		sHTML = sHTML & "			<a href='" & pURL & "&PageNo=" & iTemp & "'><img src='/images/life/starmovie_listbtn_next.gif' align='absmiddle'><img src='/images/life/null.gif' width='3' height='1'></a>"
	End If


	If iPageNo < iLastPageNo Then
		sHTML = sHTML & "			<img src='/images/life/starmovie_listbtn_last.gif' align='absmiddle'>"
	Else
		sHTML = sHTML & "			<a href='" & pURL & "&PageNo=" & iTemp & "'><img src='/images/life/starmovie_listbtn_last.gif' align='absmiddle'></a>"
	End if


	sHTML = sHTML & "		</td>"
	sHTML = sHTML & "	<tr>"
	sHTML = sHTML & "<table>"

	
	showListControl = sHTML

End Function 







Public Function showListControl2(Byval pURL, Byval lngRecordCount, Byval iPageSize, Byval iPageNo)
                                
	Dim iPrevPageNo, iNextPageNo, iLastPageNo, iPrtLastPgNo, iPrtFirstPgNo, i, sHTML
	Dim iBlockPage, iTemp, iLoop
	
	iBlockPage = 10
	
	' ������ ������ ���
	If lngRecordCount > 1 Then
		iLastPageNo = (lngRecordCount / iPageSize) 
		iLastPageNo = Round(iLastPageNo + 0.5)
	Else
		iLastPageNo = 1
	End If
	
	If iPageNo < 0 Then
		iPageNo = 1
	End If
	
	iTemp = Int((iPageNo - 1) / iBlockPage) * iBlockPage + 1

	sHTML = "<br/><table border='0' cellpadding='0' cellspacing='0'><tr>"
	
    If iPageNo > 1 Then
        sHTML = sHTML & "<td style='padding-right:10px;'><a href='" & pURL & "&PageNo=1'><img src='/intranet/new2013/images/prev.gif' border='0'></a></td><td class='txt_page'>"
    Else
		sHTML = sHTML & "<td style='padding-right:10px;'><img src='/intranet/new2013/images/prev.gif' border='0'></td><td class='txt_page'>"
    End If
    
    iLoop = 1
	
	Do Until iLoop > iBlockPage Or iTemp > iLastPageNo
        If iTemp = CInt(iPageNo) Then
            sHTML = sHTML & "<span><font color=#ff0000>" & iTemp & "</font></span>"
        Else
            sHTML = sHTML & " <a href='" & pURL & "&PageNo=" & iTemp & "'>" & iTemp & "</a> "
        End If

		If iTemp<>iLastPageNo then	
			sHTML = sHTML & " | " 
		End If
		
        iTemp = iTemp + 1
        iLoop = iLoop + 1
    Loop
        
 '   sHTML = sHTML & "<td width='104'>"
    If iPageNo < iLastPageNo Then
        sHTML = sHTML & "</td><td style='padding-left:10px;'><a href='" & pURL & "&PageNo=" & iPageNo+1 & "'><img src='/intranet/new2013/images/next.gif' border='0'></a></td>"
    Else
		sHTML = sHTML & "</td><td style='padding-left:10px;'><img src='/intranet/new2013/images/next.gif' border='0'></td>"
    End If
	
	sHTML =sHTML & "</tr></table>"
	
	showListControl2 = sHTML
End Function 



'������ ����¡
Public Function showListControl3(Byval lngRecordCount, Byval iPageSize, Byval iPageNo)
                                
	Dim iPrevPageNo, iNextPageNo, iLastPageNo, iPrtLastPgNo, iPrtFirstPgNo, i, sHTML
	Dim iBlockPage, iTemp, iLoop
	
	iBlockPage = 10
	
	' ������ ������ ���
	If lngRecordCount > 1 Then
		iLastPageNo = CInt((lngRecordCount-1) / iPageSize)+1
	Else
		iLastPageNo = 1
	End If
	
	If iPageNo < 0 Then
		iPageNo = 1
	End If
	
	iTemp = Int((iPageNo - 1) / iBlockPage) * iBlockPage + 1

	sHTML = "<br/><table border='0' cellpadding='0' cellspacing='0'><tr>"
	
    If iPageNo > 1 Then
        sHTML = sHTML & "<td style='padding-right:10px;'><a href=javascript:goPaging('"&iPageNo-1&"')><img src='/intranet/new2013/images/prev.gif' border='0'></a></td><td class='txt_page'>"
    Else
		sHTML = sHTML & "<td style='padding-right:10px;'><img src='/intranet/new2013/images/prev.gif' border='0'></td><td class='txt_page'>"
    End If
    
    iLoop = 1
	
	Do Until iLoop > iBlockPage Or iTemp > iLastPageNo
        If iTemp = CInt(iPageNo) Then
            sHTML = sHTML & "<span><font color=#ff0000>" & iTemp & "</font></span>"
        Else
            sHTML = sHTML & " <a href=javascript:goPaging('"&iTemp&"')>" & iTemp & "</a> "
        End If

		If iTemp<>iLastPageNo then	
			sHTML = sHTML & " | " 
		End If
		
        iTemp = iTemp + 1
        iLoop = iLoop + 1
    Loop

       
 '   sHTML = sHTML & "<td width='104'>"
    If CLng(iPageNo) < CLng(iLastPageNo) Then
        sHTML = sHTML & "</td><td style='padding-left:10px;'><a href=javascript:goPaging('"&iPageNo+1&"')><img src='/intranet/new2013/images/next.gif' border='0'></a></td>"
    Else
		sHTML = sHTML & "</td><td style='padding-left:10px;'><img src='/intranet/new2013/images/next.gif' border='0'></td>"
    End If
	
	sHTML =sHTML & "</tr></table>"
	
	showListControl3 = sHTML
End Function 

%>
