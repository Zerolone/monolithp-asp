<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"--><%
	Dim ThisPage, Code, TheString, TheReplaceString
	Dim Info_Pic_Folder, Info_Pic_Files, Info_Pic_Folder2, Info_Pic_FileType, Info_Pic_AutoSave, Info_Template, Info_Folder, Info_Folder2
	ThisPage				= Request.ServerVariables("PATH_INFO")
	Code						= Request("Code")
	if Code="Edit" then
		TheString					= "{����}"

		Info_Pic_Folder		= Request("Info_Pic_Folder")
		Info_Pic_Files		= Request("Info_Pic_Files")
		Info_Pic_Folder2	= Request("Info_Pic_Folder2")
		Info_Pic_FileType	= Request("Info_Pic_FileType")
		Info_Pic_AutoSave	= Request("Info_Pic_AutoSave")
		Info_Template			= Request("Info_Template")
		Info_Folder				= Request("Info_Folder")
		Info_Folder2			= Request("Info_Folder2")

		TheReplaceString	=	Info_Pic_Folder & "|" & Info_Pic_Files & "|" & Info_Pic_Folder2 & "|" & Info_Pic_FileType & "|" & Info_Pic_AutoSave & "|" & Info_Template & "|" & Info_Folder & "|" & Info_Folder2 & "|"
		
		Call	FSOFileReplaceString("InfoSettingTemplate.asp", "InfoSetting.asp", TheString , TheReplaceString)

		Response.write "<script>alert('�޸ĳɹ���');window.open('" & ThisPage & "', '_self')</script>"
	end if
%>
<!-- #include file ="InfoSetting.asp" -->

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
</head>

<form method="POST" action="<%=ThisPage%>?Code=Edit" name="InfoSettingFrm">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;��Ϣ����ϵͳ���� 
			&gt;&gt; ��Ϣ����ϵͳ��������</font></b></td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" align="right">
			<font color="#FFFFFF">ͼƬ����λ�ã�</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Info_Pic_Folder" size="30" class="InputBox" style="width=450" value="<%=InfoSettingArr(0)%>"> 
			������Ŀ¼:<%=server.mappath("/")%> </td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" align="right">
			<font color="#FFFFFF">ͼƬ�������ƣ�</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Info_Pic_Files" size="30" class="InputBox" style="width=450" value="<%=InfoSettingArr(1)%>"> 
			0��Ϊ������</td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" align="right">
			<font color="#FFFFFF">ͼƬ��վλ�ã�</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Info_Pic_Folder2" size="30" class="InputBox" style="width=450" value="<%=InfoSettingArr(2)%>"></td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" align="right">
			<font color="#FFFFFF">�ļ��������ͣ�</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Info_Pic_FileType" size="30" class="InputBox" style="width=450" value="<%=InfoSettingArr(3)%>"></td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" align="right">
			<font color="#FFFFFF">�ļ��������ƣ�</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Info_Pic_AutoSave" size="30" class="InputBox" style="width=450" value="<%=InfoSettingArr(4)%>"> 
			0Ϊ�Զ��������� 1ΪȡԴ�ļ���</td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" align="right">
			<font color="#FFFFFF">��Ϣϵͳģ�棺</font></td>
			<td bgcolor="#FFFFFF">
			<select size="1" name="Info_Template" class="InputBox" style="width=450">
			<%
			'---------------------0-----------------1------------------2
			Sql	= "Select [Template_ID], [Template_Level], [Template_Title] From [Monolith_Template] Order By [Template_Level]"
			Response.write Sql
			Rs.Open Sql, Conn, 1, 1
			Do While Not (Rs.Bof Or Rs.Eof)
			%>
			<option value="<%=Rs(0)%>" <% if cint(infosettingarr(5)) = rs(0) then response.write " selected"%>>
			<%LoopNBSP(Len(Rs(1))-2)%><%=Rs(1)%>-<%=Rs(2)%></option>
			<%
				Rs.MoveNext
			Loop
			Call CloseRs
			Call CloseConn
			%></select> ��ϢϵͳĬ�ϵ�ģ��</td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" align="right">
			<font color="#FFFFFF">��Ϣ����Ŀ¼�� 
			</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Info_Folder" size="30" class="InputBox" style="width=450" value="<%=InfoSettingArr(6)%>"> 
			������Ŀ¼:<%=server.mappath("/")%></td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" align="right">
			<font color="#FFFFFF">��Ϣ��վĿ¼�� 
			</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Info_Folder2" size="30" class="InputBox" style="width=450" value="<%=InfoSettingArr(7)%>"></td>
		</tr>
		<tr>
			<td bgcolor="#999999" width="100" height="20" align="right">��</td>
			<td bgcolor="#FFFFFF">
			<input type="submit" name="Submit0" value=" �� �� (Alt + Y) " class="InputBox" accesskey="Y">
			<input type="button" value=" ȡ �� (Alt + N) " name="B2" onclick="window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);" class="InputBox">
			</td>
		</tr>
	</table>
</form>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->