<%
Dim THIS_URL, THIS_SERVER_IP, LOGIN_RETURN_URL
Dim ckUserID, ckUserName, ckUserNickName, ckAdminID, ckAuthStr, ckAdminName
Dim MenuNum

Const SITEDOMAIN						= ".matcha.co.kr" 
Const SITEFULLDOMAIN				=	"www.matcha.co.kr" 

THIS_URL = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
THIS_SERVER_IP = Request.ServerVariables("LOCAL_ADDR")

	ckUserID						= Request.Cookies("userID")						'// ȸ�����̵�
	ckUserName					= Request.Cookies("name")							'// ȸ����
	ckUserNickName				= Request.Cookies("nickName")					'// �г���
	ckAdminID						= Request.Cookies("adminID")						'// ������ ���̵�
	ckAdminName					= Request.Cookies("adminName")					'// �����ڸ�
	ckAuthStr						= Request.Cookies("authority")						'// ������ ����
%>