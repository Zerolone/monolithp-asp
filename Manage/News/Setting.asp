<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim ThisPage, Code, TheString, TheReplaceString
	Dim News_Pic_Folder, News_Pic_Files, News_Pic_Folder2, News_Pic_FileType, News_Pic_AutoSave, News_Template, News_Folder, News_Folder2
	ThisPage				= Request.ServerVariables("PATH_INFO")
	Code						= Request("Code")
	if Code="Edit" then
		TheString					= "{����}"

		News_Pic_Folder		= Request("News_Pic_Folder")
		News_Pic_Files		= Request("News_Pic_Files")
		News_Pic_Folder2	= Request("News_Pic_Folder2")
		News_Pic_FileType	= Request("News_Pic_FileType")
		News_Pic_AutoSave	= Request("News_Pic_AutoSave")
		News_Template			= Request("News_Template")
		News_Folder				= Request("News_Folder")
		News_Folder2			= Request("News_Folder2")

		TheReplaceString	=	News_Pic_Folder & "|" & News_Pic_Files & "|" & News_Pic_Folder2 & "|" & News_Pic_FileType & "|" & News_Pic_AutoSave & "|" & News_Template & "|" & News_Folder & "|" & News_Folder2 & "|"
		
		Call	FSOFileReplaceString("NewsSettingTemplate.asp", "NewsSetting.asp", TheString , TheReplaceString)
		'Response.write "<font color=red>�޸ĳɹ���</font>"
		Response.write "<script>alert('�޸ĳɹ���');window.open('" & ThisPage & "', '_self')</script>"
	end if
%>
<!-- #include file ="NewsSetting.asp" -->
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
</head>
<p>&nbsp;��</p>
<form method="POST" action="<%=ThisPage%>?Code=Edit" name="NewsSettingFrm">
	<table border="0" width="100%" id="table1" cellspacing="0" cellpadding="0">
		<tr>
			<td>&nbsp;&nbsp;ͼƬ�ϴ�����λ�ã�<input type="text" name="News_Pic_Folder" size="30" class="InputBox" style="width=450" value="<%=NewsSettingArr(0)%>"> ������Ŀ¼:<%=server.mappath("/")%>
			</td>
		</tr>
		<tr>
			<td>&nbsp;&nbsp;ͼƬ�ϴ��������ƣ�<input type="text" name="News_Pic_Files" size="30" class="InputBox" style="width=450" value="<%=NewsSettingArr(1)%>"> 0��Ϊ������</td>
		</tr>
		<tr>
			<td>&nbsp;&nbsp;ͼƬ��Ӧ��վλ�ã�<input type="text" name="News_Pic_Folder2" size="30" class="InputBox" style="width=450" value="<%=NewsSettingArr(2)%>"></td>
		</tr>
		<tr>
			<td>&nbsp;&nbsp;�ļ����������ͣ�<input type="text" name="News_Pic_FileType" size="30" class="InputBox" style="width=450" value="<%=NewsSettingArr(3)%>"></td>
		</tr>
		<tr>
			<td>&nbsp;&nbsp;�ļ����������ͣ�<input type="text" name="News_Pic_AutoSave" size="30" class="InputBox" style="width=450" value="<%=NewsSettingArr(4)%>"> 0Ϊ�Զ��������� 1ΪȡԴ�ļ���</td>
		</tr>
		<tr>
			<td>&nbsp;&nbsp;����ϵͳĬ��ģ�棺<select size="1" name="News_Template" class="InputBox"  style="width=450">
			<%
			'---------------------0-----------------1------------------2
			Sql	= "Select [Template_ID], [Template_Level], [Template_Title] From [Monolith_Template] Order By [Template_Level]"
			Response.write Sql
			Rs.Open Sql, Conn, 1, 1
			Do While Not (Rs.Bof Or Rs.Eof)
			%>
			<option value="<%=Rs(0)%>"<% if Cint(NewsSettingArr(5)) = Rs(0) Then Response.write " selected"%>><%LoopNBSP(Len(Rs(1))-2)%><%=Rs(1)%>-<%=Rs(2)%></option>
			<%
				Rs.MoveNext
			Loop
			Call CloseRs
			Call CloseConn
			%>					</select> ����ϵͳĬ�ϵ�ģ��</td>
		</tr>
		<tr>
			<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;���ű���Ŀ¼��<input type="text" name="News_Folder" size="30" class="InputBox" style="width=450" value="<%=NewsSettingArr(6)%>"> ������Ŀ¼:<%=server.mappath("/")%></td>
		</tr>
		<tr>
			<td>&nbsp;&nbsp;���Ŷ�Ӧ��վĿ¼��<input type="text" name="News_Folder2" size="30" class="InputBox" style="width=450" value="<%=NewsSettingArr(7)%>"> </td>
		</tr>
	</table>
	<p>��</p>
	<p><input type="submit" value=" �� �� �� �� " name="B1" class="InputBox"> <input type="reset" value=" ȡ �� �� �� " name="B2" class="InputBox"></p>
</form>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->