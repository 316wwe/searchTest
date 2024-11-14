<%
Dim koreaart_uid, koreaart_tm, koreaart_ip, koreaart_macaddress, koreaart_macenable
Dim koreaart_pfs, koreaart_webalba, koreaart_staff, koreaart_cr


koreaart_uid = trim(Request.Cookies("koreaart_uid"))
koreaart_tm = trim(Request.Cookies("koreaart_tm"))
koreaart_ip = trim(Request.Cookies("koreaart_ip"))
koreaart_macaddress = trim(Request.Cookies("koreaart_macaddress"))
koreaart_macenable = trim(Request.Cookies("koreaart_macenable"))


koreaart_pfs = Request.Cookies("koreaart_pfs")
koreaart_webalba = Request.Cookies("koreaart_webalba")
koreaart_staff = Request.Cookies("koreaart_staff")
koreaart_cr = Request.Cookies("koreaart_cr")

'If Request.Cookies("koreaart_pfs")<>"68F1C7AE-E2F0-45A3-BF70-3CAB9234C579" And Request.Cookies("koreaart_webalba") <> "A0458800-C2F3-4275-820A-5172AC48BC3B" And Request.Cookies("koreaart_staff") <> "A255EF63-D740-40B0-B76C-1CB1E0ED9DDC" And Request.Cookies("koreaart_cr") <> "FRGGH82H-K9R5-LKG9-JEYR-12POFUGKEHE5" And Request.Cookies("koreaart_tm") <> "4BE08EC5-5487-45DA-A98B-2DD41A7423F5" And Request.Cookies("koreaart_staff") <> "GFIDS5G3-G384-GGLE-PR3R-1CB1E0ED9DDC" And Request.Cookies("koreaart_staff") <> "B2FEG4GE-A380-5DFD-4GSG-2CD8GDA78ADC" And Request.Cookies("koreaart_staff") <> "RTDGDEAG-DD45-45GT-TTEE-34RTDFFHFHHH" Then

If koreaart_pfs<>"34FEKUER-OIEP-Q8GB-9EPE-3CAB9234CLEG" And koreaart_webalba <> "A0458800-C2F3-4275-820A-5172AC48BC3B" And koreaart_staff <> "F8EIE9GQ-D840-40B0-KKWR-1CB1E0ED9DDC" And koreaart_cr <> "FRGGH82H-K9R5-LKG9-JEYR-12POFUGKEHE5" And koreaart_tm <> "4BE08EC5-5487-45DA-A98B-2DD41A7423F5" And koreaart_staff <> "GFIDS5G3-G384-GGLE-PR3R-1CB1E0ED9DDC" And koreaart_staff <> "B2FEG4GE-A380-5DFD-4GSG-2CD8GDA78ADC" And koreaart_staff <> "RTDGDEAG-DD45-45GT-TTEE-34RTDFFHFHHH" Then
%>
	<script language="javascript">
		alert("정상적인 접근이 아닙니다.\n\n자동로그아웃처리 됩니다");
		location.replace("/intranet/logout_ok.asp");
	</script>
<%
End If


'쿠키를 유출의심이 들면 login_ok2.asp에서 쿠키값을 바꾸고 여기도 바꾼다.
If koreaart_staff="F8EIE9GQ-D840-40B0-KKWR-1CB1E0ED9DDC" Then '입학관리부
	koreaart_staff = "입학관리부"
End if

If koreaart_pfs="34FEKUER-OIEP-Q8GB-9EPE-3CAB9234CLEG" Then
	koreaart_pfs = "교수"
End if


'if trim(Request.ServerVariables("REMOTE_ADDR"))="119.70.156.40" or trim(Request.ServerVariables("REMOTE_ADDR"))="119.70.156.41" then
	'Response.Write "koreaart_uid : " &  koreaart_uid & "<br>"
	'Response.Write "koreaart_tm : " &  koreaart_tm & "<br>"
	'Response.Write "koreaart_ip : " &  koreaart_ip & "<br>"
	'Response.Write "koreaart_macaddress : " &  koreaart_macaddress & "<br>"
	'Response.Write "koreaart_macenable : " &  koreaart_macenable & "<br>"
'end if

if (koreaart_uid="" or koreaart_ip="") then
	%>
	<script language="javascript">
		alert("로그인정보가 올바르지 않습니다\n\n자동로그아웃처리 됩니다");
		location.replace("/intranet/logout_ok.asp");
	</script>
	<%
	Response.End
end if
%>

<%if trim(koreaart_macenable)="1" then%>
<OBJECT id="iSysInfo" classid="clsid:8DAA3668-D06F-48BC-9DC2-3626B5B57DEF" codebase="/config/iSysInfo.CAB#version=1,0,0,4">
	<param name="copyright" value="http://isulnara.com">
</OBJECT>
	<script language="javascript">
	function Installed()
	{
		try
		{
			return (new ActiveXObject('iSysInfo.iSysInfoX'));
		}
		catch (e)
		{
			return false;
		}
	}

	if ( !Installed() )
	{
		goWhenNoModule();
	}
	else
	{
		if ("<%=Left(koreaart_macaddress,17)%>"!=String(iSysInfo.MacAddress).substring(0,17)) 
		{
			//로그인할때 맥어드레스와 현재 맥어드레스가 다르면 로그아웃 처리한다. 여기 아래를 오픈해야 합니다
			<%if date>Cdate("2011-02-10") then%>
			alert("::: 경고 :::\n\n보안주소가 다릅니다\n\nPC가 바뀌었거나 인증을 하지 않았거나, 내부 하드웨어가 변경되었습니다\n\n\n\n자동로그아웃됩니다");
			location.replace("/intranet/logout_ok.asp");
			<%end if%>
		}
	}

	function goWhenNoModule()
	{
		alert("보안모듈이 설치되지 않았습니다\n\n상단의 노란색 설치바를 클릭후 추가기능설치를 클릭하세요\n\n상단 노란색바가 보이지 않을시는 잠시 기다리면 됩니다\n\n먼저 보안모듈을 설치해야 로그인 및 사용할수 있습니다");

		//var index = String(location.href).split('/')[String(location.href).split('/').length-1];	
		location.replace("/intranet/logout_ok.asp");
	}
	</script>
<%end if%>