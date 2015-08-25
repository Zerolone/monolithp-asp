<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp"-->
<!-- #include virtual ="/Inc/MD5.asp"-->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim ThisPage
	Dim Users_ID
	Users_ID = Request("Users_ID")
	CheckNum Users_ID	
		
	Dim Page

	Page	= Request("Page")

  ThisPage					= Request.ServerVariables("PATH_INFO")
	if Request("Code") = "Edit" then
		Dim Users_Password, Users_Level_Group, Users_Level, Users_Group, Users_Active
		
		Users_Password		= Request("Users_Password")
		if Users_Password	<> "" then Users_Password	= MD5(Users_Password, 32)
		Users_Level_Group	= Request("Users_Level_Group")
		Users_Level_Group	= Split(Users_Level_Group, "|")
		Users_Level				=	Users_Level_Group(0)
		Users_Group				=	Users_Level_Group(1)
		Users_Active			= Request("Users_Active")

		'-------------------0-----------------1--------------2--------------3
		Sql = "Select [Users_Password], [Users_Level], [Users_Group], [Users_Active] From [Monolith_Users] Where [Users_ID]=" & Users_ID
		Rs.Open Sql, Conn, 1, 3
		if Users_Password	<> "" then Rs(0)		= Users_Password
		Rs(1)		= Users_Level
		Rs(2)		= Users_Group
		Rs(3)		= Users_Active
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
	Dim TrBgColor
	
	Users_ID = Request("Users_ID")
	CheckNum Users_ID

	'-------------------0-------------1-----------------2--------------3--------------4------------------5
	Sql = "Select [Users_Name], [Users_Password], [Users_Level], [Users_Group], [Users_LastLogin], [Users_Active] From [Monolith_Users] Where [Users_ID]=" & Users_ID
	Rs.Open Sql, Conn, 1, 1
	If Rs.Bof or Rs.Eof then
	  Response.write "这个会员不存在！"
		Response.end
	end if	

%>
<form name="form1" method="post" action="<%=ThisPage%>?Code=Edit">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;会员管理 &gt;&gt; 修改会员</font></b></td>
		</tr>
		<tr>
			<td width="100" height="20" bgcolor="#999999" align="right">
			<font color="#FFFFFF">会员名称：</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(0)%>
			</td>
		</tr>
		<tr>
			<td width="100" height="20" bgcolor="#999999" align="right">
			<font color="#FFFFFF">会员密码：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Users_Password" type="text" size="40" class="InputBox">
			不修改请留空</td>
		</tr>
		<tr>
			<td bgcolor="#999999" width="100" height="20" align="right">
			<font color="#FFFFFF">会员等级：</font></td>
			<td bgcolor="#FFFFFF">
			<select size="1" name="Users_Level_Group" class="InputBox" style="width=250">
        <%
			Dim Rs1
			Set Rs1		= Server.CreateObject("ADODB.Recordset")
			'-------------------0-----------1
			Sql = "Select [Level_ID], [Level_Name] From [Monolith_Users_Level] Order By [Level_Order] Asc"
			Rs1.Open Sql, Conn ,1, 1
			Do While Not (Rs1.Bof Or Rs1.Eof)
		%>
        <option value="<%=Rs1(0)%>|<%=Rs1(1)%>" <% if Rs1(0) = Rs(2) then Response.write "Selected"  %>><%=Rs1(1)%></option>
        <%
				Rs1.MoveNext
			Loop
			Rs1.Close
			Set Rs1=Nothing
		%>
					</select></td>
		</tr>
		<tr>
			<td height="20" bgcolor="#999999" width="100" align="right">
			<font color="#FFFFFF">最后登陆：</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(4)%>
			</td>
		</tr>
		<tr>
			<td bgcolor="#999999" width="100" height="20" align="right">
			<font color="#FFFFFF">是否激活：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Users_Active" type="radio" value="0" id="Users_Active_0" <% if not rs(5) then response.write "checked" %>><label for="Users_Active_0">禁用</label>
			<input name="Users_Active" type="radio" value="1" id="Users_Active_1" <% if rs(5) then response.write "checked" %>><label for="Users_Active_1">激活</label>
			</td>
		</tr>
		<tr>
			<td bgcolor="#999999" width="100" height="20" align="right">　</td>
			<td bgcolor="#FFFFFF">
			<input type="submit" name="Submit0" value=" 修 改 (Alt + Y) " class="InputBox" accesskey="Y">
			<input type="button" value=" 取 消 (Alt + N) " name="B2" onclick="window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);" class="InputBox">
			<input type="hidden" name="Users_ID" value="<%=Users_ID%>">
			<input type="hidden" name="Page" value="<%=Page%>"></td>
		</tr>
	</table>
</form>
<%
		Call CloseRs
		Call CloseConn
%><table><tr><td height="330"></td></tr></table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->