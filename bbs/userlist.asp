<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!-- #include file="connection.asp" -->
<!-- #include file="std.asp" -->
<%
	OnlySysOp
	Alive "用户管理"
%>
<html >
<head>
<meta http-equiv="pragma" content="no-cache" >
<link rel="stylesheet" type="text/css" href="main.css">
</head>
<body class=clblue>
<%		dim rsUserInfo,strSQL,iUser,i,byOrder,pgNo
		set rsUserInfo=server.CreateObject("ADODB.RecordSet")

		pgNo=request("pgNo")
		if pgNo="" then pgNo=1
		byOrder=request("order")
		if byOrder="" then byOrder="userid"
		strSQL="select * from UserInfo order by "+byOrder

		rsUserInfo.CursorLocation = 3
		rsUserInfo.CacheSize = 50
		rsUserInfo.Open strSQL,myconn,1
		iUser=rsUserInfo.RecordCount
		rsUserInfo.MoveFirst
        rsUserInfo.PageSize = 50
        Dim TotalPages
        TotalPages = rsUserInfo.PageCount
		rsUserInfo.AbsolutePage = pgNo
		Dim Count
		dim dNo,uNo
		dNo=pgNo-1
		uNo=pgNo+1
		if dNo=0 then dNo=1
		if uNo>totalpages then uNo=totalpages
%>
<table  class=main cellspacing=0 width=100%>
<tr class=usertitle><td align=right>
	第<%=pgNo%>页，共<%=totalpages%>页&nbsp;&nbsp;&nbsp;&nbsp;
	<a HREF=userlist.asp?order=<%=byOrder%>&pgNo=<%=dNo%>>上一页</a>&nbsp;&nbsp;
	<a HREF=userlist.asp?order=<%=byOrder%>&pgNo=<%=UNo%>>下一页</a>&nbsp;&nbsp;
</td></tr>
</table>
<table  class=main cellspacing=2 align=center>
	<tr class=usertitle>
		<td width='4%' ><center><a class=user href=userlist.asp?order=id>ID</a></center></td>
		<td width='10%' ><center><a class=user href=userlist.asp?order=userid>用户ID</a></center></td>
		<td width='14%' ><center><a class=user href=userlist.asp?order=nickname>呢称</a></center></td>
		<td width='15%' ><center><a class=user href=userlist.asp?order=Email>Email</a></center></td>
		<td width='4%' ><center><a class=user href=userlist.asp?order=iLogin>上站</a></center></td>
		<td width='4%' ><center><a class=user href=userlist.asp?order=iBook>文章</a></center></td>
		<td width='4%' ><center><a class=user href=userlist.asp?order=iPerience>经验</a></center></td>
		<td width='10%' ><center><a class=user href=userlist.asp?order=lastlogin>最后上站时间</a></center></td>
	</tr>
<%	Do While Not rsUserInfo.EOF And Count < rsUserInfo.PageSize
		if count mod 2=1 then %>
			<tr class=userline1>
<%		else %>
			<tr class=userline2>
<%		end if  %>
		<td align=right>
		<%		if isSysop(session("userid")) then%><a class=userlist href=adduser.asp?method=delete&userid=<%=rsUserInfo("userid")%>><%end if%>
		   <%=rsUserInfo("id")%>
		</a></td>
		<td>
<%		if isSysop(session("userid")) then%>
		<a class=userlist href=adduser.asp?method=modify&userid=<%=rsUserInfo("userid")%>>
<%		end if%>
		<%=rsUserInfo("userid")%></a></td>
		<td><%=rsUserInfo("nickname")%></td>
		<td><%=rsUserInfo("Email")%></td>
		<td align=right><%=rsUserInfo("iLogin")%></td>
		<td align=right><%=rsUserInfo("ibook")%></td>
		<td align=right><%=rsUserInfo("iperience")%></td>
		<td align=center><%=formatdatetime(rsUserInfo("lastlogin"),2)%></td>
	</tr>
<%	rsUserinfo.movenext
	Count=Count+1
	Loop 
	rsUserInfo.close
	myconn.close
%>
</table>
</body>
</html>
<%
	set rsUserInfo=Nothing
	set myconn=Nothing
%>