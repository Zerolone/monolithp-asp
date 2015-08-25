<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp"--><!-- #include virtual ="/Manage/Check.asp"--><%
	Dim ThisPage, Page
	Dim Bulletin_ID
	Page							= Request("Page")
  ThisPage					= Request.ServerVariables("PATH_INFO")
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<%
	Dim TrBgColor
	
	Bulletin_ID = Request("Bulletin_ID")
	CheckNum Bulletin_ID
	
	Sql = "Select [Bulletin_Title], [Bulletin_Title2], [Bulletin_Content], [Bulletin_Content2] From [Monolith_Bulletin] Where [Bulletin_ID]=" & Bulletin_ID
	Rs.Open Sql, Conn, 1, 1
	If Rs.Bof or Rs.Eof then
	  Response.write "这个广告不存在！"
		Response.end
	end if	
%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<form method="POST" action="Del_Submit.asp"  name="FrmDel">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;公告管理 &gt;&gt; 公告删除</font></b></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">公 告&nbsp;&nbsp;&nbsp;&nbsp; ID：</font></td>
			<td bgcolor="#FFFFFF">系统自动生成</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">公 告&nbsp; 标 题：</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Bulletin_Title" size="40" style="width=250px;" value="<%=Rs(0)%>" class="InputBox" disabled>
			<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">公 告 副标题：</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Bulletin_Title2" size="40" style="width=250px;" value="<%=Rs(1)%>" class="InputBox" disabled></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">公告略缩内容：</font></td>
			<td bgcolor="#FFFFFF"><font color="#FF0000">
			<textarea rows="2" name="Bulletin_Content2" cols="20" style="width=500;height:100" class="InputBox" disabled><%=Rs(3)%></textarea></font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF">（一般为调用内容）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">公 告&nbsp; 内 容：</font></td>
			<td bgcolor="#FFFFFF"><font color="#FF0000">
			<textarea rows="2" name="Bulletin_Content" cols="20" style="width=500;height:100" class="InputBox" disabled><%=Rs(2)%></textarea> 
			*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF">（公告显示内容）</td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF"><input type="hidden" value="Edit" name="Code">
			<input type="hidden" value="<%=Bulletin_ID%>" name="Bulletin_ID">
			<input type="hidden" value="<%=Page%>" name="Page">
			<input type="submit" value=" 删 除 确 认 (Alt + Y) " name="B1" class="InputBox" accesskey="Y">
			<input type="reset" value=" 返 回 (Alt + N) " name="B2" class="InputBox" accesskey="N" onclick="window.open('Default.asp?Page=<%=Page%>','_self');"></td>
		</tr>
	</form>
</table>
<table><tr><td height="144"></td></tr></table>
<%
	Call CloseRs
	Call CloseConn
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->