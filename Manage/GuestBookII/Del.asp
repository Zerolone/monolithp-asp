<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"--><%
	Dim GuestBook_ID

	GuestBook_ID			= Request("GuestBook_ID")	
	CheckNum(GuestBook_ID)

	Dim Page
	Page	= Request("Page")



	'-----------------------0---------------------1---------------------2---------------------3-------------------4-----------------------5
	Sql = "Select [GuestBook_UserName], [GuestBook_UserMail], [GuestBook_Content], [GuestBook_Postdate], [GuestBook_Replay], [GuestBook_ReplayDate], [GuestBook_Active] From [Monolith_GuestBookII] Where [GuestBook_ID] = " & [GuestBook_ID]
	Rs.Open Sql, Conn, 1, 1
	if not (Rs.eof and Rs.Bof) then 
	%>
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<form name="repl" method="post" action="Del_Submit.asp?Action=Reply&GuestBook_id=<%=GuestBook_id%>">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;留言板管理 &gt;&gt; 留言删除</font></b></td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">
			<font color="#FFFFFF">留言日期：</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(3)%></td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">
			<font color="#FFFFFF">留 言 人：</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(0)%></td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">
			<font color="#FFFFFF">邮&nbsp;&nbsp;&nbsp; 件：</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(1)%></td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">
			<font color="#FFFFFF">留言内容：</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(2)%></td>
		</tr>
		<tr>
			<td align="right" width="100" valign="top" height="20" bgcolor="#999999">
			<font color="#FFFFFF">回&nbsp;&nbsp;&nbsp; 复：</font></td>
			<td valign="top" bgcolor="#FFFFFF">&nbsp;<%if isnull(Rs(4)) then response.write "<b>新留言，未回复</b>" else response.write Rs(4)%></td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">
			<font color="#FFFFFF">是否通过：</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(6)%></td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">　</td>
			<td bgcolor="#FFFFFF">
			<input type="hidden" name="Page" value="<%=Page%>">
			<input type="submit" name="Submit0" value=" 删 除 (Alt + Y) " class="InputBox" accesskey="Y">
			<input type="reset" value=" 取 消 (Alt + N) " name="B2" class="InputBox" accesskey="N" onclick="window.open('Default.asp?Page=<%=Page%>','_self');">&nbsp;
			</td>
		</tr>
	</table>
</form>
<form>
</form>
<%
	end if	
Call CloseRs
Call CloseConn
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->