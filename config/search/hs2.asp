<%@Language="VBScript"%>
<!--#include virtual="/config/common/commonvar.asp"-->
<!--#include virtual="/config/common/const.asp"-->
<!--#include virtual="/config/common/commonproc.asp"-->
<!--#include virtual="/config/common/dbconf.asp"-->

<%
Dim hs : hs = Replace(trim(Request.Form("hs")),"'","''")

GetDbConn
GetRs
%>
<html>
<head>
<title>�� �����б��� �˻� ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">

<meta name="viewport" content="width=device-width,initial-scale=1.0,minimum-scale=1.0,maximum-scale=1.0,user-scalable=no" />
<meta name="viewport" content="height=device-height,width=device-width" />
<meta http-equiv="Cache-Control" content="no-cache" />
<meta http-equiv="Pragma" content="no-cache" />

<link rel="stylesheet" type="text/css" href="../../config/css/style.css">
<script language="javascript" src="../../config/js/commonproc.js"></script>
<script language="javascript">
window.onload = function()
{
	var form = document.forms[0];
	form.hs.focus();
}
function goSearch(form)
{
	form.action = "hs2.asp";
	form.submit();
}
function goSelect(code, hs)
{
	opener.document.getElementById("hs_code").value = code;
	opener.document.getElementById("hs").value = hs;
	self.close();
}

function goNonameSet()
{
	opener.document.getElementById("hs").value = document.getElementById("hs").value;
	self.close();
}
</script>
<style type="text/css">
.box{width:40%;height:20px;}
</style>
</head>

<body leftmargin="0" topmargin="0">
<form method="post">
<div style="margin:auto">
<div style="width:98%;border:4px solid #f8dfe0;">
	
		<table width="98%" height="353" border="0" cellpadding="0" cellspacing="0">
		<tr>
		<td width="30" valign="top" bgcolor="#ffffff"></td>
		<td valign="top" bgcolor="#ffffff">
		
			<table width="100%" height="353"  border="0" cellpadding="0" cellspacing="0">
			<tr>
			<td height="12"></td>
			</tr>
			
			<tr>
			<td><img src="images/pop_txt01.gif" width="101" height="19">&nbsp;<span style="color:green;"> (���� 02-2063-0700)</span></td>
			</tr>
			
			<tr>
			<td height="10"></td>
			</tr>			
			<tr>
			<td height="112" valign="top">
			
				<table width="100%" height="267" border="0" cellpadding="0" cellspacing="0">
				<tr>
				<td width="4" rowspan="13" background="/intranet/images/center/add_line_l.gif"></td>
				<td height="1" bgcolor="#E0E0DF"></td>
				<td width="4" rowspan="13" background="/intranet/images/center/add_line_r.gif"></td>
				</tr>
				
				<tr>
				<td width="100%" height="21" bgcolor="#FBFBFB"></td>
				</tr>
				
				<tr>
				<td height="22" bgcolor="#FBFBFB" style="padding-left:10;"><span style="font-size:12;">�����б���</span>&nbsp;&nbsp;<input name="hs" type="text" style='border:1 solid #DDDBDE; color=#676767 font-size:9pt; background-color:#ffffff' value="<%=hs%>" size="24" style="ime-mode:active;font-size:12;"  class="box">
				&nbsp;<a href="javascript:goSearch(document.forms[0]);" onFocus="this.blur();"><img src="/images/common/btn_search02.gif" width="37" height="20" border="0" align="absmiddle"></a> </td>
				</tr>
				
				<tr>
				<td height="21" bgcolor="#FBFBFB"></td>
				</tr>
				
				<tr>
				<td height="1" bgcolor="#FBFBFB">
				  
					<table width="100%"  border="0" cellspacing="0" cellpadding="0">
					<tr>
					<td width="25"></td>
					<td   background="/intranet/images/center/bg_line.gif">&nbsp;</td>
					<td width="25"></td>
					</tr>
					</table>
				  
				</td>
				</tr>
				
				<tr>
				<td height="23" bgcolor="#FBFBFB"></td>
				</tr>
				
				<tr>
				<td height="11" bgcolor="#FBFBFB" style="padding-left:30;"><img src="/images/common/pop_txt03.gif" width="47" height="13"></td>
				</tr>
				
				<tr>
				<td height="9" bgcolor="#FBFBFB"></td>
				</tr>
				
				<tr>
				<td height="21" bgcolor="#FBFBFB" style="padding-left:30;" class="TA">
				

				<span style=" border:1 dashed solid #D3D3D3; WIDTH: 100%;HEIGHT:100px; font-size:12px;overflow-y:scroll;">
				<%
				if (hs<>"") then
					oQry = ""
					oQry = "select HsCode, Name, Sido from dbo.TB_HighSchool where (Name like '%"&hs&"%') or (OldName1 like '%"&hs&"%') or (OldName2 like '%"&hs&"%')  or (OldName3 like '%"&hs&"%');"
					'Response.Write oQry
					oRs.Open oQry,oConn,3,1
					if not ors.eof and not ors.bof then
						while not ors.eof
							%>
							<a href="javascript:goSelect('<%=oRs("HsCode")%>','<%=oRs("Name")%>');" title="�����ϼ���!">(<%=oRs("Sido")%>)&nbsp;<%=oRs("Name")%></a><br>
							<%
							oRs.Movenext
						wend
					else
						Response.Write "�˻������ �����ϴ�. <br><br><a href=javascript:goNonameSet() style='color:red;font-size:16;'><b>�˻��� �����б����� ����Ϸ��� ���⸦ Ŭ���ϼ���</b></a>"
					end if
				else
					Response.Write "�˻������ �����ϴ�."
				end if
				%>
				</span>
				</td>
				</tr>
				
				<tr>
				<td height="17" bgcolor="#FBFBFB"></td>
				</tr>
				
				<tr>
				<td height="18" align="center" bgcolor="#FBFBFB"><a href="javascript:self.close();" onFocus="this.blur();"><img src="/images/common/btn_can01.gif" width="70" height="23" border="0"></a></td>
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
	
  </div>
</div>
</form>
</body>
</html>
<%
SetFreeObj(oRs)
SetFreeObj(oConn)
%>