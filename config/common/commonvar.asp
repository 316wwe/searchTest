<%
'//�ۼ��� : gigatera
'//�ۼ��� : 2010-05-31
'//��   �� : ���뺯�� ����
'option Explicit

Response.AddHeader "P3P","CP='NOI DSP NID TAIo PSAa OUR IND UNI OTC TST'"

Dim ConnStr '��񿬰� ���ڿ��� ������ ���ڿ� ����		
Dim oMail '���ϸ��� �Ҷ� smtp��ü�� ������ ��ü ����
Dim oConn '��񿬰ᰪ�� ���Ϲ޴� ��񿬰� ��ü ����
Dim oConn2 '��񿬰ᰪ�� ���Ϲ޴� ��񿬰� ��ü ����
Dim oConnM '��񿬰ᰪ�� ���Ϲ޴� ��񿬰� ��ü ����
Dim oConnH '��񿬰ᰪ�� ���Ϲ޴� ��񿬰� ��ü ����
Dim oRs '���ڵ���� ������ ���ڵ�� ��ü ����
Dim oRs2 '���ڵ���� ������ ���ڵ�� ��ü ����
Dim oRs3 '���ڵ���� ������ ���ڵ�� ��ü ����
Dim oCmd '�ɸ�� ��ü�� ������ Ŀ�ǵ� ��ü ����
Dim oQry '�������� �����ϴ� ���� ��Ʈ�� ����
Dim exQry 'Ŀ�ǵ� ��ü�� ������� �ʰ�, stored procedure�� ����� �� ���� ���� ��Ʈ�� ����
Dim Cnt 'ī��Ʈ ������ ����

Dim Res 'on error resume �� ��� ����ϴ� ���� üũ boolean ����
Dim Chk '���� �����ϴ����� ���θ� ���� �� ����ϴ� ������ ����
Dim i 'for ������ ����ϴ� ������ ����
Dim j 'for ������ ����ϴ� ������ ����
Dim k 'for ������ ����ϴ� ������ ����
Dim l'for ������ ����ϴ� ������ ����
Dim m'for ������ ����ϴ� ������ ����
Dim n'for ������ ����ϴ� ������ ����
Dim z'for ������ ����ϴ� ������ ����
Dim view '���̱�/����� ���� ������ ����ϴ� boolean ����

Dim fso   '���� �ý��� ��ü(file system object)
Dim fp     '���� ������ ��ü(file pointer)
Dim lpstr '�ؽ�Ʈ ������ �о�帱 ��Ʈ�� ����(long pointer string)

' asp ���ε� ������Ʈ 
Dim Image
Dim theForm, theField, bExist , countFileName, saveFileName, FileName
Dim uploadPath
Dim GetPreUrl

Dim board_titles : board_titles = Array("����QnA","��������","�б��ҽ�","��к���","�л��ڷ��")

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


Dim OffCharger(8) '�������� ���� ����Ʈ	
OffCharger(0) = "��ö��"
OffCharger(1) = "���м�"
OffCharger(2) = "����"
OffCharger(3) = "������"
OffCharger(4) = "������"
OffCharger(5) = "�ּ���"
OffCharger(6) = "������"
OffCharger(7) = "ȫ���"
                                                                                          
Dim offChargerCnt
offChargerCnt = 8

Dim OffTeam(8) '������
OffTeam(0) = "1��"
OffTeam(1) = "2��"
OffTeam(2) = "2��"
OffTeam(3) = "1��"
OffTeam(4) = "2��"
OffTeam(5) = "1��"
OffTeam(6) = "1��"
OffTeam(7) = "2��"


Dim OnCharger(2) '�¶��� ���� ����Ʈ	
OnCharger(0) = "������"
OnCharger(1) = "�赵��"

Dim onChargerCnt
onChargerCnt = 2

Dim CounselCharger(4) '����� ���� ����Ʈ	
CounselCharger(0) = "������"
CounselCharger(1) = "�ս¹�"
CounselCharger(2) = "������"
CounselCharger(3) = "�����"

Dim counselChargerCnt
counselChargerCnt = 4

'�˸��� ����
Dim Kakao_Nation_Code, Callback_No, Kakao_Sender_Key
Kakao_Nation_Code = "82"
Callback_No = "02-2063-0700"
Kakao_Sender_Key = "8376e81990574a32b9f3ea3c8ff027a3ea4501dd"
%>