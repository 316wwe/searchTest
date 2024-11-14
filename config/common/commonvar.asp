<%
'//작성자 : gigatera
'//작성일 : 2010-05-31
'//설   명 : 공통변수 모음
'option Explicit

Response.AddHeader "P3P","CP='NOI DSP NID TAIo PSAa OUR IND UNI OTC TST'"

Dim ConnStr '디비연결 문자열을 저장할 문자열 변수		
Dim oMail '메일링을 할때 smtp객체를 얻어오는 객체 변수
Dim oConn '디비연결값을 리턴받는 디비연결 객체 변수
Dim oConn2 '디비연결값을 리턴받는 디비연결 객체 변수
Dim oConnM '디비연결값을 리턴받는 디비연결 객체 변수
Dim oConnH '디비연결값을 리턴받는 디비연결 객체 변수
Dim oRs '레코드셋을 얻어오는 레코드셋 객체 변수
Dim oRs2 '레코드셋을 얻어오는 레코드셋 객체 변수
Dim oRs3 '레코드셋을 얻어오는 레코드셋 객체 변수
Dim oCmd '케멘드 객체를 얻어오는 커맨드 객체 변수
Dim oQry '쿼리문을 저장하는 쿼리 스트링 변수
Dim exQry '커맨드 객체를 사용하지 않고, stored procedure를 사용할 때 쓰는 쿼리 스트링 변수
Dim Cnt '카운트 정수형 변수

Dim Res 'on error resume 문 등에서 사용하는 에러 체크 boolean 변수
Dim Chk '값이 존재하는지의 여부를 따질 때 사용하는 정수형 변수
Dim i 'for 문에서 사용하는 정수형 변수
Dim j 'for 문에서 사용하는 정수형 변수
Dim k 'for 문에서 사용하는 정수형 변수
Dim l'for 문에서 사용하는 정수형 변수
Dim m'for 문에서 사용하는 정수형 변수
Dim n'for 문에서 사용하는 정수형 변수
Dim z'for 문에서 사용하는 정수형 변수
Dim view '보이기/숨기기 같은 곳에서 사용하는 boolean 변수

Dim fso   '파일 시스템 객체(file system object)
Dim fp     '파일 포인터 객체(file pointer)
Dim lpstr '텍스트 파일을 읽어드릴 스트링 변수(long pointer string)

' asp 업로드 컴포넌트 
Dim Image
Dim theForm, theField, bExist , countFileName, saveFileName, FileName
Dim uploadPath
Dim GetPreUrl

Dim board_titles : board_titles = Array("입학QnA","공지사항","학교소식","언론보도","학사자료실")

Dim fcColors : fcColors = Array("F6BD0F","8BBA00","FF8E46","008E8E","D64646","8E468E","588526","B3AA00","008ED6","9D080D","A186BE","F6BD0F","8BBA00","FF8E46","008E8E","D64646","8E468E","588526","B3AA00","008ED6","9D080D","A186BE","F6BD0F","8BBA00","FF8E46","008E8E","D64646","8E468E","588526","B3AA00","008ED6","9D080D","A186BE","F6BD0F","8BBA00","FF8E46","008E8E","D64646","8E468E","588526","B3AA00","008ED6","9D080D","A186BE","F6BD0F","8BBA00","FF8E46","008E8E","D64646","8E468E","588526","B3AA00","008ED6","9D080D","A186BE","F6BD0F","8BBA00","FF8E46","008E8E","D64646","8E468E","588526","B3AA00","008ED6","9D080D","A186BE","F6BD0F","8BBA00","FF8E46","008E8E","D64646","8E468E","588526","B3AA00","008ED6","9D080D","A186BE","F6BD0F","8BBA00","FF8E46","008E8E","D64646","8E468E","588526","B3AA00","008ED6","9D080D","A186BE","F6BD0F","8BBA00","FF8E46","008E8E")


Dim entrance
Dim koreaart_files : koreaart_files="/intranet/files/"
%>


<%
Dim myAddr 
myAddr = trim(Request.ServerVariables("SERVER_NAME")) & trim(Request.ServerVariables("SCRIPT_NAME"))
if trim(Request.ServerVariables("QUERY_STRING"))<>"" then
	myAddr = myAddr & "?" & trim(Request.ServerVariables("QUERY_STRING"))
end if


Dim OffCharger(8) '오프라인 직원 리스트	
OffCharger(0) = "강철민"
OffCharger(1) = "강학성"
OffCharger(2) = "고석준"
OffCharger(3) = "류혜정"
OffCharger(4) = "서윤주"
OffCharger(5) = "주세용"
OffCharger(6) = "최항진"
OffCharger(7) = "홍경식"
                                                                                          
Dim offChargerCnt
offChargerCnt = 8

Dim OffTeam(8) '오프팀
OffTeam(0) = "1팀"
OffTeam(1) = "2팀"
OffTeam(2) = "2팀"
OffTeam(3) = "1팀"
OffTeam(4) = "2팀"
OffTeam(5) = "1팀"
OffTeam(6) = "1팀"
OffTeam(7) = "2팀"


Dim OnCharger(2) '온라인 직원 리스트	
OnCharger(0) = "김정분"
OnCharger(1) = "김도훈"

Dim onChargerCnt
onChargerCnt = 2

Dim CounselCharger(4) '상담직 직원 리스트	
CounselCharger(0) = "이정아"
CounselCharger(1) = "손승민"
CounselCharger(2) = "권정민"
CounselCharger(3) = "정명숙"

Dim counselChargerCnt
counselChargerCnt = 4

'알림톡 관련
Dim Kakao_Nation_Code, Callback_No, Kakao_Sender_Key
Kakao_Nation_Code = "82"
Callback_No = "02-2063-0700"
Kakao_Sender_Key = "8376e81990574a32b9f3ea3c8ff027a3ea4501dd"
%>