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
		alert("�������� ������ �ƴմϴ�.\n\n�ڵ��α׾ƿ�ó�� �˴ϴ�");
		location.replace("/intranet/logout_ok.asp");
	</script>
<%
End If


'��Ű�� �����ǽ��� ��� login_ok2.asp���� ��Ű���� �ٲٰ� ���⵵ �ٲ۴�.
If koreaart_staff="F8EIE9GQ-D840-40B0-KKWR-1CB1E0ED9DDC" Then '���а�����
	koreaart_staff = "���а�����"
End if

If koreaart_pfs="34FEKUER-OIEP-Q8GB-9EPE-3CAB9234CLEG" Then
	koreaart_pfs = "����"
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
		alert("�α��������� �ùٸ��� �ʽ��ϴ�\n\n�ڵ��α׾ƿ�ó�� �˴ϴ�");
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
			//�α����Ҷ� �ƾ�巹���� ���� �ƾ�巹���� �ٸ��� �α׾ƿ� ó���Ѵ�. ���� �Ʒ��� �����ؾ� �մϴ�
			<%if date>Cdate("2011-02-10") then%>
			alert("::: ��� :::\n\n�����ּҰ� �ٸ��ϴ�\n\nPC�� �ٲ���ų� ������ ���� �ʾҰų�, ���� �ϵ��� ����Ǿ����ϴ�\n\n\n\n�ڵ��α׾ƿ��˴ϴ�");
			location.replace("/intranet/logout_ok.asp");
			<%end if%>
		}
	}

	function goWhenNoModule()
	{
		alert("���ȸ���� ��ġ���� �ʾҽ��ϴ�\n\n����� ����� ��ġ�ٸ� Ŭ���� �߰���ɼ�ġ�� Ŭ���ϼ���\n\n��� ������ٰ� ������ �����ô� ��� ��ٸ��� �˴ϴ�\n\n���� ���ȸ���� ��ġ�ؾ� �α��� �� ����Ҽ� �ֽ��ϴ�");

		//var index = String(location.href).split('/')[String(location.href).split('/').length-1];	
		location.replace("/intranet/logout_ok.asp");
	}
	</script>
<%end if%>