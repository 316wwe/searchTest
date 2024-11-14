<!--#include virtual="/lib/encoding.asp"-->
<!--#include virtual="/lib/dbClass.asp"-->
<!--#include virtual="/lib/globalFunction.asp"-->
<!--#include virtual="/lib/commonvar.asp"-->



<%
	Dim hs : hs = Replace(trim(Request.Form("hs")),"'","''")
%>
<!DOCTYPE html>
<html lang="ko">
<head>
<title>고등학교 검색</title>
<link rel="stylesheet" type="text/css" href="../../config/css/style.css">
<script language="javascript" src="../../config/js/commonproc.js"></script>
<meta name="viewport" content="width=device-width,initial-scale=1.0,minimum-scale=1.0,maximum-scale=1.0,user-scalable=no" />
<script language="javascript">
window.onload = function()
{
	var form = document.forms[0];
	form.hs.focus();
}
function goSearch(form)
{
	form.action = "et_hs2.asp";
	form.submit();
}
function goSelect(code, hs, charger)
{
	opener.setHs(code, hs, charger);
	self.close();
}

function goNonameSet()
{
	opener.document.getElementById("hs").value = document.getElementById("hs").value;
	self.close();
}
</script>
</head>

<body leftmargin="0" topmargin="0">
<form method="post">

	<table width="466" height="363"  border="0" cellpadding="0" cellspacing="0">
	<tr>
	<td align="center" bgcolor="#f8dfe0">
	
		<table width="456" height="353" border="0" cellpadding="0" cellspacing="0">
		<tr>
		<td width="30" valign="top" bgcolor="#ffffff"></td>
		<td width="396" valign="top" bgcolor="#ffffff">
		
			<table width="396" height="353"  border="0" cellpadding="0" cellspacing="0">
			<tr>
			<td height="12"></td>
			</tr>
			
			<tr>
			<td><img src="images/pop_txt01.gif" width="101" height="19">
			
			
			&nbsp;<span style="color:green;"> (문의 02-2063-0700)</span>
			</td>
			</tr>
			
			<tr>
			<td height="10"></td>
			</tr>
			
			<tr>
			<td><img src="images/pop_line.gif" width="392" height="1"></td>
			</tr>
			
			<tr>
			<td height="15"></td>
			</tr>
			
			<tr>
			<td height="112" valign="top">
			
				<table width="396" height="267" border="0" cellpadding="0" cellspacing="0">
				<tr>
				<td width="4" rowspan="13"></td>
				<td height="1" bgcolor="#E0E0DF"></td>
				<td width="4" rowspan="13"></td>
				</tr>
				
				<tr>
				<td width="388" height="21" bgcolor="#FBFBFB"></td>
				</tr>
				
				<tr>
				<td height="22" bgcolor="#FBFBFB" style="padding-left:30;"><span style="font-size:12;">고등학교명 </span>&nbsp;&nbsp;<input name="hs" type="text" style='border:1 solid #DDDBDE; color=#676767 font-size:9pt; background-color:#ffffff' value="<%=hs%>" size="24" style="ime-mode:active;font-size:12;">
				&nbsp;<a href="javascript:goSearch(document.forms[0]);" onFocus="this.blur();"><img src="images/btn_search02.gif" width="37" height="20" border="0" align="absmiddle"></a>
		
				
				</td>
				</tr>
				
				<tr>
				<td height="21" bgcolor="#FBFBFB" align="center"></td>
				</tr>
				
				<tr>
				<td height="1" bgcolor="#FBFBFB">
				  
					<table width="100%"  border="0" cellspacing="0" cellpadding="0">
					<tr>
					<td width="25"></td>
					<td width="338"></td>
					<td width="25"></td>
					</tr>
					</table>
				  
				</td>
				</tr>
				
				<tr>
				<td height="23" bgcolor="#FBFBFB"></td>
				</tr>
				
				<tr>
				<td height="11" bgcolor="#FBFBFB" style="padding-left:30;"><img src="images/pop_txt03.gif" width="47" height="13">
				<span style="color:green;"> (결과는 최신학교명으로 출력됩니다)</span>
				</td>
				</tr>
				
				<tr>
				<td height="9" bgcolor="#FBFBFB"></td>
				</tr>
				
				<tr>
				<td height="21" bgcolor="#FBFBFB" style="padding-left:30;" class="TA">
				

				<span style=" border:1 dashed solid #D3D3D3; WIDTH: 330px;HEIGHT:100px; font-size:12px;overflow-y:scroll;">
<%
				if (hs<>"") Then

					SET ClsDB = New DataBase
					ClsDB.ConnOpen()

					With Cmd
						.CommandType = adCmdStoredProc
						.CommandText = "dbo.USP_HIGHSCHOOL_LIST"
						.parameters.Append .CreateParameter("@name",		adVarchar,	adParamInput,	100,	hs)

						SET Rs = .Execute
						if not Rs.EOF then
							ResultList = RS.GetRows
						end If
						
						Rs.close
						SET Rs = nothing
			
						.parameters.Delete("@name")
					End With

					SET ClsDB = Nothing
					
					If isArray(ResultList) Then 
						For loop_int=0 To Ubound(ResultList,2)
							hsCode	= ResultList(0, loop_int)
							hs		= ResultList(1, loop_int)
							oldHs	= ResultList(2, loop_int)
							charger = ResultList(3, loop_int)
							sido	= ResultList(4, loop_int)
							gugun	= ResultList(5, loop_int)
							isMou	= ResultList(6, loop_int)

%>
							<a href="javascript:goSelect('<%=hsCode%>','<%=hs%>', '<%=charger%>');" title="선택하세요!">(<%=sido%>)&nbsp;<%=hs%></a><br>
<%
						Next

					Else
						Response.Write "검색결과가 없습니다. <br><br><a href=javascript:goNonameSet() style='color:red;font-size:16;'><b>검색한 고등학교명을 사용하려면 여기를 클릭하세요</b></a>"
					End if
				Else
					Response.Write "검색결과가 없습니다."
				End If
%>
				</span>


				</td>
				</tr>
				
				<tr>
				<td height="17" bgcolor="#FBFBFB"></td>
				</tr>
				
				<tr>
				<td height="18" align="center" bgcolor="#FBFBFB"><a href="javascript:self.close();" onFocus="this.blur();"><img src="images/btn_can01.gif" width="70" height="23" border="0"></a></td>
				</tr>
				
				<tr>
				<td height="22" bgcolor="#FBFBFB"></td>
				</tr>
				
				<tr>
				<td height="1" bgcolor="#E0E0DF"></td>
				</tr>
				</table>
			
			</td>
			</tr>
			
			<tr>
			<td height="14"></td>
			</tr>
			</table>
			
		</td>
		<td width="30" valign="top" bgcolor="#ffffff"></td>
		</tr>
		</table>
	
	</td>
	</tr>
  </table>

</form>
</body>
</html>
