<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp"--><!-- #include virtual ="/Manage/Check.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">

<%
	Dim Remark_ID
	
	Remark_ID = Request("Remark_ID")
	CheckNum Remark_ID
	
	Dim Page
	Page			= Request("Page")
	
	'--------------------0------------------1---------------2
	Sql = "Select [Remark_UserName], [Remark_Content], [News_ID] From [Monolith_News_Remark] Where [Remark_ID]=" & Remark_ID
	Rs.Open Sql, Conn, 1, 1
	If Rs.Bof or Rs.Eof then
	  Response.write "������۲����ڣ�"
		Response.end
	end if	
%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<form method="POST" action="Del.asp" name="FrmDel">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;�������۹��� &gt;&gt; �����޸�</font></b></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�� �� ID��</font></td>
			<td bgcolor="#FFFFFF">ϵͳ�Զ�����</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�� �� �ˣ�</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Remark_UserName" size="40" style="width=250px;" value="<%=Rs(0)%>" class="InputBox" disabled>
			<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�������ݣ�</font></td>
			<td bgcolor="#FFFFFF">
			<font color="#FF0000">
			<textarea rows="2" name="Remark_Content" cols="20" style="width=500;height:100" class="InputBox" disabled><%=Monolith_Decode(Rs(1))%></textarea></font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF">��������ʾ���ݣ�</td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF">
			<input type="hidden" value="<%=Remark_ID%>" name="Remark_ID">
			<input type="hidden" value="<%=Rs(2)%>" name="News_ID">
			<input type="hidden" value="<%=Page%>" name="Page">
			<input type="submit" value=" ȷ �� ɾ ��  (Alt + Y) " name="B1" class="InputBox" accesskey="Y">
			<input type="submit" value=" �� �� (Alt + N)" name="B2" class="InputBox" onclick="window.open('Default.asp','_self')" accesskey="N"></td>
		</tr>
	</form>
</table>
<%
	Call CloseRs
	Call CloseConn
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->