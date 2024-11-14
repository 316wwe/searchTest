<%@Language="VBScript"%>
<!--#include virtual="/config/common/commonvar.asp"-->
<!--#include virtual="/config/common/const.asp"-->
<!--#include virtual="/config/common/commonproc.asp"-->
<!--#include virtual="/config/common/dbconf.asp"-->

<%
Dim address : address = trim(Request.Form("address"))

GetDbConn
GetRs
%>
<html>
<head>
<title>▒ 우편번호 검색 ▒</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" type="text/css" href="/config/css/style.css">
<script language="javascript" src="/config/js/commonproc.js"></script>
<script language="javascript">
window.onload = function()
{
	var form = document.forms[0];
	form.address.focus();
}
function goSearch(form)
{
	form.action = "addr.asp";
	form.submit();
}
function goSelect(zipcode,address)
{
	var zips = String(zipcode).split('-');
	
	var pform = opener.document.forms[0];
	pform.c_zip1.value = zips[0];
	pform.c_zip2.value = zips[1];
	pform.c_addr1.value = address;
	pform.c_addr2.value = "";
	pform.c_addr2.focus();
	self.close();
}
</script>
</head>

<body leftmargin="0" topmargin="0">
<form method="post">

	<table width="466" height="363"  border="0" cellpadding="0" cellspacing="0">
	<tr>
	<td align="center" bgcolor="#EADEEA">
	
		<table width="456" height="353" border="0" cellpadding="0" cellspacing="0">
		<tr>
		<td width="30" valign="top" bgcolor="#ffffff"></td>
		<td width="396" valign="top" bgcolor="#ffffff">
		
			<table width="396" height="353"  border="0" cellpadding="0" cellspacing="0">
			<tr>
			<td height="12"></td>
			</tr>
			
			<tr>
			<td><img src="/intranet/images/center/text_add_search.gif" width="103" height="17"></td>
			</tr>
			
			<tr>
			<td height="10"></td>
			</tr>
			
			<tr>
			<td><img src="/intranet/images/center/text_pw_input01.gif" width="294" height="18"></td>
			</tr>
			
			<tr>
			<td height="15"></td>
			</tr>
			
			<tr>
			<td height="112" valign="top">
			
				<table width="396" height="267" border="0" cellpadding="0" cellspacing="0">
				<tr>
				<td width="4" rowspan="13" background="../../images/center/add_line_l.gif"></td>
				<td height="1" bgcolor="#E0E0DF"></td>
				<td width="4" rowspan="13" background="../../images/center/add_line_r.gif"></td>
				</tr>
				
				<tr>
				<td width="388" height="21" bgcolor="#FBFBFB"></td>
				</tr>
				
				<tr>
				<td height="22" bgcolor="#FBFBFB" style="padding-left:30;"><span style="font-size:12;">동이름 입력</span>&nbsp;&nbsp;<input name="address" type="text" style='border:1 solid #DDDBDE; color=#676767 font-size:9pt; background-color:#ffffff' value="<%=address%>" size="24" style="ime-mode:active;font-size:12;">
				&nbsp;<a href="javascript:goSearch(document.forms[0]);" onFocus="this.blur();"><img src="/intranet/images/button/bt_add_search.gif" width="60" height="20" border="0" align="absmiddle"></a> </td>
				</tr>
				
				<tr>
				<td height="21" bgcolor="#FBFBFB"></td>
				</tr>
				
				<tr>
				<td height="1" bgcolor="#FBFBFB">
				  
					<table width="100%"  border="0" cellspacing="0" cellpadding="0">
					<tr>
					<td width="25"></td>
					<td width="338"  background="/intranet/images/center/bg_line.gif"></td>
					<td width="25"></td>
					</tr>
					</table>
				  
				</td>
				</tr>
				
				<tr>
				<td height="23" bgcolor="#FBFBFB"></td>
				</tr>
				
				<tr>
				<td height="11" bgcolor="#FBFBFB" style="padding-left:30;"><img src="/intranet/images/center/text_04.gif" width="62" height="11"></td>
				</tr>
				
				<tr>
				<td height="9" bgcolor="#FBFBFB"></td>
				</tr>
				
				<tr>
				<td height="21" bgcolor="#FBFBFB" style="padding-left:30;" class="TA">
				

				<span style=" border:1 dashed solid #D3D3D3; WIDTH: 330px;HEIGHT:100px; font-size:12px;overflow-y:scroll;">
				<%
				if (address<>"") then
					oQry = ""
					oQry = "select zipcode,(isnull(sido,'')+' '+isnull(gugun,'')+' '+isnull(dong,'')+' '+isnull(ri,'')+' '+isnull(bldg,'')) as address,(isnull(sido,'')+' '+isnull(gugun,'')+' '+isnull(dong,'')+' '+isnull(ri,'')+' '+isnull(bldg,'')+' '+isnull(st_bunji,'')+ (case when isnull(ed_bunji,'')<>'' then '~'+isnull(ed_bunji,'') else '' end) )   as address from tbl_zipcode where dong like '%"&address&"%';"
					'Response.Write oQry
					oRs.Open oQry,oConn,3,1
					if not ors.eof and not ors.bof then
						while not ors.eof
							%>
							<a href="javascript:goSelect('<%=oRs(0)%>','<%=oRs(1)%>');" title="선택하세요!">(<%=oRs(0)%>)&nbsp;<%=oRs(2)%></a><br>
							<%
							oRs.Movenext
						wend
					else
						Response.Write "검색결과가 없습니다."
					end if
				else
					Response.Write "검색결과가 없습니다."
				end if
				%>
				</span>


				</td>
				</tr>
				
				<tr>
				<td height="17" bgcolor="#FBFBFB"></td>
				</tr>
				
				<tr>
				<td height="18" align="center" bgcolor="#FBFBFB"><a href="javascript:self.close();" onFocus="this.blur();"><img src="/intranet/images/button/bt_cancle_center.gif" width="61" height="18" border="0"></a></td>
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
<%
SetFreeObj(oRs)
SetFreeObj(oConn)
%>