<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp"-->
<!-- #include virtual ="/Inc/MD5.asp"-->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim ThisPage
	Dim From_ID
	From_ID = Request("From_ID")
	CheckNum From_ID	
		
	Dim Page

	Page	= Request("Page")

  ThisPage					= Request.ServerVariables("PATH_INFO")
	if Request("Code") = "Edit" then
		Dim From_Name, From_Order
		
		From_Name		= Request("From_Name")
		From_Order	= Request("From_Order")

		'--------------------0--------------1
		Sql = "Select [From_Name], [From_Order] From [Monolith_News_From] Where [From_ID]=" & From_ID
		Rs.Open Sql, Conn, 1, 3
		Rs(0)		= From_Name
		Rs(1)		= From_Order
		Rs.Update
		Rs.Close

		Response.write("<script language='javascript'>alert('修改成功！')</script>")

	end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
</head>

<body>

<%
	Dim i
		
	'--------------------0--------------1
	Sql = "Select [From_Name], [From_Order] From [Monolith_News_From] Where [From_ID]=" & From_ID
	Rs.Open Sql, Conn, 1, 1
	If Rs.Bof or Rs.Eof then
	  Response.write "这个作者不存在！"
		Response.end
	end if
%>
<form name="form1" method="post" action="<%=ThisPage%>?Code=Edit">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;新闻系统设置 &gt;&gt; 新闻</font></b><font color="#FFFFFF"><b>来源管理 
		&gt;&gt; 新闻来源修改</b></font></td>
		</tr>
		<tr>
			<td width="100" height="20" bgcolor="#999999" align="right">
			<font color="#FFFFFF">新闻作者：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="From_Name" type="text" size="40" class="InputBox" value="<%=Rs(0)%>" style="width=250"></td>
		</tr>
		<tr>
			<td width="100" height="20" bgcolor="#999999" align="right">
			<font color="#FFFFFF">顺&nbsp;&nbsp;&nbsp; 序：</font></td>
			<td bgcolor="#FFFFFF">
			<select name="From_Order" size="1" class="InputBox" style="width=250">
			<%For i=1 to 100%>
			<option value="<%=i%>" <% if Rs(1) = i then Response.Write "selected" %>><%=i%></option>
			<%Next%>
			</select> </td>
		</tr>
		<tr>
			<td bgcolor="#999999" width="100" height="20" align="right">　</td>
			<td bgcolor="#FFFFFF">
			<input type="submit" name="Submit0" value=" 修 改 (Alt + Y) " class="InputBox" accesskey="Y">
			<input type="button" value=" 取 消 (Alt + N) " name="B2" onclick="window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);" class="InputBox">
			<input type="hidden" name="From_ID" value="<%=From_ID%>">
			<input type="hidden" name="Page" value="<%=Page%>"></td>
		</tr>
	</table>
</form>
<%
		Call CloseRs
		Call CloseConn
%><table><tr><td height="330"></td></tr></table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->