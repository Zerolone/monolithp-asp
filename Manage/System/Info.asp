<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<body>

<%
	Dim CounterIp
	CounterIp = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
	If CounterIp = "" Then CounterIp = Request.ServerVariables("REMOTE_ADDR")
%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="4"><b><font color="#FFFFFF">&nbsp;��������ͻ�����Ϣ</font></b></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="20">
		<td width="12%" align="right">&nbsp;�� �� ��&nbsp; �� �ͣ� </td>
		<td width="38%">&nbsp;<%=Request.ServerVariables("OS")%></td>
		<td width="12%" align="right">&nbsp;�� �� ��&nbsp; I&nbsp; P�� </td>
		<td width="36%">&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
	</tr>
	<tr height="20">
		<td width="12%" align="right">&nbsp;�� �� ��&nbsp; �� �ڣ� </td>
		<td width="38%">&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
		<td width="12%" align="right">&nbsp;�� �� ��&nbsp; ʱ�䣺 </td>
		<td width="36%">&nbsp;<%=now%></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="20">
		<td width="12%" align="right">&nbsp;��&nbsp; ��&nbsp; ��&nbsp;&nbsp; ���� </td>
		<td width="38%">&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
		<td width="12%" align="right">&nbsp;������ CPU������ </td>
		<td width="36%">&nbsp;<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%></td>
	</tr>
	<tr height="20">
		<td width="12%" align="right">&nbsp;������ ����ϵͳ�� </td>
		<td width="38%">&nbsp;<%=Request.ServerVariables("OS")%></td>
		<td width="12%" align="right">��</td>
		<td width="36%">��</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="20">
		<td width="12%" align="right">&nbsp;IIS&nbsp;&nbsp;&nbsp;&nbsp; ��&nbsp;&nbsp; ���� </td>
		<td width="38%">&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
		<td width="12%" align="right">&nbsp;վ �� ����·���� </td>
		<td width="36%">&nbsp;<%=Request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
	</tr>
	<tr height="20">
		<td width="12%" align="right">&nbsp;��&nbsp; �� ��ʱʱ�䣺</td>
		<td width="38%">&nbsp;<%=Server.ScriptTimeout%> ��</td>
		<td width="12%" align="right">&nbsp;�������������棺 </td>
		<td width="36%">&nbsp;<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %>��</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="20">
		<td width="12%" align="right">Session��ʱʱ�䣺</td>
		<td width="38%">&nbsp;<%=session.Timeout%> ����</td>
		<td width="12%" align="right">ϵ ͳ ռ�ÿռ䣺</td>
		<td width="36%">&nbsp;<%
	Dim Fs, f
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set f = fs.GetFolder(server.mappath("/"))
	response.write Formatnumber(f.size/1048576,2,-1)&"&nbsp;MB"
	set f=nothing
%></td>
	</tr>
	<tr height="20">
		<td width="12%" align="right">&nbsp;���ݿ�(ADO)ʹ�ã� </td>
		<td width="38%"><b><font color="#CC0000" class="red">&nbsp;<%
      	if ObjInstall("adodb.connection")=false then 
      	  Response.write "��"
      	else
      	  Response.write "��"
      	  Response.write " �汾��" & Conn.veRsion
      	end if
      %></font></b> ��</td>
		<td width="12%" align="right">&nbsp;FSO�� �� �� д�� </td>
		<td width="36%"><b><font color="#CC0000" class="red">&nbsp;<%
      	if ObjInstall("scripting.filesystemobject")=false then 
      	  Response.write "��"
      	else
      	  Response.write "��"
      	end if
      %></font></b>��</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="20">
		<td width="12%" align="right">&nbsp;Jmail&nbsp; ���֧�֣� </td>
		<td width="38%"><b><font color="#CC0000" class="red">&nbsp;<%
      	if ObjInstall("JMail.SMTPMail")=false then 
      	  Response.write "��"
      	else
      	  Response.write "��"
      	end if
      %></font></b> ��</td>
		<td width="12%" align="right">&nbsp;CDONTS���֧�֣� </td>
		<td width="36%"><b><font color="#CC0000" class="red">&nbsp;<%
      	if ObjInstall("CDONTS.NewMail")=false then 
      	  Response.write "��"
      	else
      	  Response.write "��"
      	end if
      %></font></b> ��</td>
	</tr>
	<tr height="20">
		<td width="12%" align="right">�� ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; IP��</td>
		<td width="38%">&nbsp;<%=CounterIp%></td>
		<td width="12%" align="right">��</td>
		<td width="36%">��</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="20">
		<td width="12%" align="right">�� ����������ͣ�</td>
		<td width="86%" colspan="3">&nbsp;<%=Request.ServerVariables("HTTP_USER_AGENT")%></td>
	</tr>
	<tr height="20">
		<td width="98%" colspan="4">��</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="20">
		<td width="98%" align="right" colspan="4">��</td>
	</tr>
	<tr height="20">
		<td width="98%" colspan="4">��</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="20">
		<td width="98%" align="right" colspan="4">��</td>
	</tr>
	<tr height="20">
		<td width="98%" colspan="4">��</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="20">
		<td width="98%" align="right" colspan="4">��</td>
	</tr>
	<tr height="20">
		<td width="98%" colspan="4">��</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="20">
		<td width="98%" align="right" colspan="4">��</td>
	</tr>
	<tr height="20">
		<td width="98%" colspan="4">��</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="20">
		<td width="98%" align="right" colspan="4">��</td>
	</tr>
	<tr height="20">
		<td width="98%" colspan="4">��</td>
	</tr>
	<tr bgcolor="#FFFFFF" height="20">
		<td width="98%" align="right" colspan="4">��</td>
	</tr>
</table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->

</body>