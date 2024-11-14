<!--#include virtual="/config/inc/intranet_non_layout.asp"-->

<script>
	$(function() {
		$("#btnSubmit").on("click", function() {
			var institute_name = $("#institute_name").val();
			if (institute_name == "") {
				alert("학원명을 입력해 주세요\n예)프리");
			}
			$("form").submit();
		});
	});

	function goSelect(code, name) {
		opener.setInstitute(code, name);
		self.close();
	}
</script>

<%
	Dim institute_name
	institute_name = Replace(trim(Request.Form("institute_name")),"'","''")
%>


<form method="post" method="post" action="et_ins.asp">
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
								<td><img src="images/pop_institute_title.gif">
								&nbsp;<span style="color:green;"></span>
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
											<td height="22" bgcolor="#FBFBFB" style="padding-left:30;">
												<span style="font-size:12;">학원명 </span>&nbsp;&nbsp;<input name="institute_name" id="institute_name" type="text" style='border:1 solid #DDDBDE; color=#676767 font-size:9pt; background-color:#ffffff' value="<%=hs%>" size="24" style="ime-mode:active;font-size:12;">&nbsp;<img src="images/btn_search02.gif" width="37" height="20" border="0" align="absmiddle" id="btnSubmit" style="cursor:pointer;">
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
											<td height="11" bgcolor="#FBFBFB" style="padding-left:30;"><img src="images/pop_txt03.gif" width="47" height="13"></td>
										</tr>
										<tr>
											<td height="9" bgcolor="#FBFBFB"></td>
										</tr>
										<tr>
											<td height="21" bgcolor="#FBFBFB" style="padding-left:30;" class="TA">
												<span style=" border:1 dashed solid #D3D3D3; WIDTH: 330px;HEIGHT:100px; font-size:12px;overflow-y:scroll;">
<%
												If institute_name<>"" Then
													
													Set ClsDB = New DataBase
													ClsDB.ConnOpen()

													sql = "SELECT idx, name, sido, gugun FROM tb_institute WHERE name LIKE '%" & institute_name & "%' AND hoseo_gubun='hac' ORDER BY name ASC"

													With Cmd
														.CommandType = adCmdText
														.CommandText = sql
														
														SET Rs = .Execute
														isInstituteList = 0
														if not Rs.EOF Then
															isInstituteList = 1
															ResultList = RS.GetRows
														end If

														Rs.close
														SET Rs = nothing
													End With

													Set ClsDB = Nothing

													If isInstituteList = "1" Then
%>
														<table cellspacing="1" cellpadding="5" width="350" style="border-top:1px solid #000000; background-color:#B3B3B3;">
															<tr class="tcolorgray txtblack tacenter">
																<td>코드</td>
																<td>학원명</td>
																<td>시도</td>
																<td>구군</td>
															</tr>
<%
															For loop_int=0 To Ubound(ResultList,2)
																idx		= ResultList(0, loop_int)
																name	= ResultList(1, loop_int)
																sido	= ResultList(2, loop_int)
																gugun	= ResultList(3, loop_int)
%>
																<tr>
																	<td class="tcolorwhite txtblack tacenter"><%=idx%></td>
																	<td class="tcolorwhite txtblack taleft"><a href="javascript:goSelect('<%=idx%>','<%=name%>')"><%=name%></a></td>
																	<td class="tcolorwhite txtblack tacenter"><%=sido%></td>
																	<td class="tcolorwhite txtblack tacenter"><%=gugun%></td>
																</tr>
<%

															Next
%>
														</table>
<%
													Else
														Response.Write "검색 결과가 없습니다"
													End if


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