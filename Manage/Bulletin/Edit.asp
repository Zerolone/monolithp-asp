<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp"--><!-- #include virtual ="/Manage/Check.asp"--><%
	Dim ThisPage, Page
	Dim Bulletin_ID
	
	Page							= Request("Page")

  ThisPage					= Request.ServerVariables("PATH_INFO")
	if Request("Code") = "Edit" then
		Bulletin_ID = Request("Bulletin_ID")
		CheckNum Bulletin_ID
		
		Dim Bulletin_Title, Bulletin_Title2, Bulletin_Content, Bulletin_Content2

		Bulletin_Title			= Request("Bulletin_Title")
		Bulletin_Title2			= Request("Bulletin_Title2")
		Bulletin_Content		= Request("Bulletin_Content")
		Bulletin_Content2		= Request("Bulletin_Content2")

		Bulletin_Title			= MonolithEncode(Bulletin_Title)
		Bulletin_Title2			= MonolithEncode(Bulletin_Title2)
		Bulletin_Content		= MonolithEncode(Bulletin_Content)
		Bulletin_Content2		= MonolithEncode(Bulletin_Content2)
		
		'---------------------0------------------1-----------------2--------------------3
		Sql = "Select [Bulletin_Title], [Bulletin_Title2], [Bulletin_Content], [Bulletin_Content2] From [Monolith_Bulletin] Where [Bulletin_ID]=" & Bulletin_ID
		Rs.Open Sql, Conn, 1, 3
		Rs(0)								= Bulletin_Title
		Rs(1)								= Bulletin_Title2
		Rs(2)								= Bulletin_Content
		Rs(3)								= Bulletin_Content2
		Rs.Update
		Rs.Close
		
		Response.Redirect("Default.asp?Page=" & Page)

	end if
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<script language="javascript">
function check()
{
  var Check_Bulletin_Title	= document.FrmAddMenu.Bulletin_Title;
  if (Check_Bulletin_Title.value == "")
  {
 	alert("请输入公告标题！");
	Check_Bulletin_Title.focus();
	return false;
  }
  
  var Check_Bulletin_Content	= document.FrmAddMenu.Bulletin_Content;
  if (Check_Bulletin_Content.value == "")
  {
 	alert("请输入公告内容！");
	Check_Bulletin_Content.focus();
	return false;
  }
}
</script>
<%
	Dim TrBgColor
	
	Bulletin_ID = Request("Bulletin_ID")
	CheckNum Bulletin_ID
	
	'----------------------0-----------------1------------------2-------------------3
	Sql = "Select [Bulletin_Title], [Bulletin_Title2], [Bulletin_Content], [Bulletin_Content2] From [Monolith_Bulletin] Where [Bulletin_ID]=" & Bulletin_ID
	Rs.Open Sql, Conn, 1, 1
	If Rs.Bof or Rs.Eof then
	  Response.write "这个广告不存在！"
		Response.end
	end if	
%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<form method="POST" action="Edit.asp" onsubmit="return check()" name="FrmAddMenu">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;公告管理 &gt;&gt; 公告修改</font></b></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">公 告&nbsp;&nbsp;&nbsp;&nbsp; ID：</font></td>
			<td bgcolor="#FFFFFF"><%=Bulletin_ID%>（系统自动生成）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">公 告&nbsp; 标 题：</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Bulletin_Title" size="40" style="width=250px;" value="<%=Rs(0)%>" class="InputBox">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">公 告 副标题：</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Bulletin_Title2" size="40" style="width=250px;" value="<%=Rs(1)%>" class="InputBox"></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">公告略缩内容：</font></td>
			<td bgcolor="#FFFFFF"><font color="#FF0000">
			<textarea rows="2" name="Bulletin_Content2" cols="20" style="width=500;height:100" class="InputBox"><%=Rs(3)%></textarea></font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF">（一般为调用内容）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">公 告&nbsp; 内 容：</font></td>
			<td bgcolor="#FFFFFF"><font color="#FF0000">
			<textarea rows="2" name="Bulletin_Content" cols="20" style="width=500;height:100" class="InputBox"><%=Rs(2)%></textarea> 
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
			<input type="submit" value=" 提 交 (Alt + S) " name="B1" class="InputBox" accesskey="A">
			<input type="reset" value=" 返 回 (Alt + N) " name="B2" class="InputBox" accesskey="N" onclick="window.open('Default.asp?Page=<%=Page%>','_self');"></td>
		</tr>
	</form>
</table>
<table><tr><td height="144"></td></tr></table>
<%
	Call CloseRs
	Call CloseConn
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->