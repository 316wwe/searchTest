<%
Dim THIS_URL, THIS_SERVER_IP, LOGIN_RETURN_URL
Dim ckUserID, ckUserName, ckUserNickName, ckAdminID, ckAuthStr, ckAdminName
Dim MenuNum

Const SITEDOMAIN						= ".matcha.co.kr" 
Const SITEFULLDOMAIN				=	"www.matcha.co.kr" 

THIS_URL = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
THIS_SERVER_IP = Request.ServerVariables("LOCAL_ADDR")

	ckUserID						= Request.Cookies("userID")						'// 회원아이디
	ckUserName					= Request.Cookies("name")							'// 회원명
	ckUserNickName				= Request.Cookies("nickName")					'// 닉네임
	ckAdminID						= Request.Cookies("adminID")						'// 관리자 아이디
	ckAdminName					= Request.Cookies("adminName")					'// 관리자명
	ckAuthStr						= Request.Cookies("authority")						'// 관리자 권한
%>