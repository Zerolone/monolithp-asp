<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"--><%
	Dim ThisPage, GuestBook_ID

  ThisPage					= Request.ServerVariables("PATH_INFO")
	GuestBook_ID			= Request("GuestBook_ID")	
	CheckNum(GuestBook_ID)

	if Request("Action") = "Reply" then
		'-----------------------0---------------------1-------------------2-----------------------3
		Sql = "Select [GuestBook_Replayer], [GuestBook_Replay], [GuestBook_ReplayDate], [GuestBook_Active] From [Monolith_GuestBookII] Where [GuestBook_ID] = " & GuestBook_ID
		Rs.Open Sql, Conn, 1, 3
		if not (rs.eof and rs.bof) then
			Rs(0)		= Session("Admin_Name")
			Rs(1)		= Request("GuestBook_Replay")
			Rs(2) 	= Now()
			Rs(3)		= Request("GuestBook_Active")
			Rs.update
		end if
		rs.close

		Response.write("<script language='javascript'>alert('�޸ĳɹ���')</script>")
	end if

	'-----------------------0---------------------1---------------------2---------------------3-------------------4-----------------------5
	Sql = "Select [GuestBook_UserName], [GuestBook_UserMail], [GuestBook_Content], [GuestBook_Postdate], [GuestBook_Replay], [GuestBook_ReplayDate], [GuestBook_Active] From [Monolith_GuestBookII] Where [GuestBook_ID] = " & [GuestBook_ID]
	Rs.Open Sql, Conn, 1, 1
	if not (Rs.eof and Rs.Bof) then 
	%>
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">

<form name="repl" method="post" action="?Action=Reply&GuestBook_id=<%=GuestBook_id%>">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;���԰���� &gt;&gt; ��˻ظ�����</font></b></td>
		</tr>
	<tr>
		<td align="right" width="100" height="20" bgcolor="#999999">
		<font color="#FFFFFF">�������ڣ�</font></td>
		<td bgcolor="#FFFFFF"> &nbsp;<%=Rs(3)%></td>
	</tr>
	<tr>
		<td align="right" width="100" height="20" bgcolor="#999999">
		<font color="#FFFFFF">�� �� �ˣ�</font></td>
		<td bgcolor="#FFFFFF"> &nbsp;<%=Rs(0)%></td>
	</tr>
	<tr>
		<td align="right" width="100" height="20" bgcolor="#999999">
		<font color="#FFFFFF">��&nbsp; &nbsp; ����</font></td>
		<td bgcolor="#FFFFFF"> &nbsp;<%=Rs(1)%></td>
	</tr>
	<tr>
		<td align="right" width="100" height="20" bgcolor="#999999">
		<font color="#FFFFFF">�������ݣ�</font></td>
		<td bgcolor="#FFFFFF"> &nbsp;<%=Rs(2)%></td>
	</tr>
	<tr>
		<td align="right" width="100" valign="top" height="20" bgcolor="#999999"><font color="#FFFFFF">��&nbsp;&nbsp;&nbsp; ����</font></td>
		<td valign="top" bgcolor="#FFFFFF"> &nbsp;<%if isnull(Rs(4)) then response.write "<b>�����ԣ�δ�ظ�</b>" else response.write Rs(4)%></td>
	</tr>
	<tr>
		<td align="right" width="100" valign="top" height="20" bgcolor="#999999">
		<font color="#FFFFFF"><br>
		���»ظ���<font></td>
		<td bgcolor="#FFFFFF"><textarea name="GuestBook_Replay" cols="60" rows="5"><%=Rs(4)%></textarea>
		</td>
	</tr>
	<tr>
		<td align="right" width="100" height="20" bgcolor="#999999">
		<font color="#FFFFFF">�Ƿ�ͨ����</font></td>
		<td bgcolor="#FFFFFF"> &nbsp;<input type="radio" name="GuestBook_Active" id="GuestBook_Active_1" value="true" <% if Rs(6) then Response.write "checked" %>><label for="GuestBook_Active_1">ͨ�� </label>
		<input type="radio" name="GuestBook_Active" value="false" id="GuestBook_Active_0" <% if Not Rs(6) then Response.write "checked" %> ><label for="GuestBook_Active_0">��ͨ�� </label></td>
	</tr>
	<tr>
		<td align="right" width="100" height="20" bgcolor="#999999">��</td>
		<td bgcolor="#FFFFFF">
			<input type="submit" name="Submit0" value=" �� �� (Alt + Y) " class="InputBox" accesskey="Y">
			<input type="button" value=" ȡ �� (Alt + N) " name="B2" onclick="window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);" class="InputBox">&nbsp; </td>
	</tr>
</table>
	<form>
<%
	end if	
	Call CloseRs
	Call CloseConn
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->