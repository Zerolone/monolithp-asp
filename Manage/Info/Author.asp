<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>无标题文档</title>
<script language="javascript" src="/Js/All.js"></script>
</head>

<body>

<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="3"><b><font color="#FFFFFF">&nbsp;信息发布系统设置 &gt;&gt; 信息作者</font></b><font color="#FFFFFF"><b>管理 
		&gt;&gt; 信息作者察看</b></font></td>
	</tr>
	<tr height="20" bgcolor="#FFFFFF">
		<td width="71%" bgcolor="#999999"><font color="#FFFFFF">&nbsp;作者名称</font></td>
		<td width="9%" bgcolor="#999999"><font color="#FFFFFF">&nbsp;顺序</font></td>
		<td width="12%" bgcolor="#999999"><font color="#FFFFFF">&nbsp;操作</font></td>
	</tr>
	<%
		Dim ThisPage , Page, Pages, j, i
	  ThisPage							= Request.ServerVariables("PATH_INFO")
		Page									= clng(request("Page"))
		Dim TrBgColor
		'--------------------0------------1--------------2
		Sql = "Select [Author_ID], [Author_Name], [Author_Order] From [Monolith_Info_Author] Order By [Author_Order] Asc"
		Rs.Open Sql, Conn ,1, 1
		if Not (Rs.bof Or Rs.Eof) then
			Rs.PageSize=19
			if page=0 then page=1
			pages=rs.pagecount
			if page > pages then page=pages
			rs.AbsolutePage=page
			for j=1 to rs.PageSize 
				if j mod 2 = 0 then 
					TrBgColor	= ""
				else 
					TrBgColor	= "bgcolor=""#FFFFFF"""
				end if
		%>
	<tr <%=trbgcolor%>>
		<td height="22" width="71%">&nbsp;<%=Rs(1)%></a></td>
		<td height="22">&nbsp;<%=Rs(2)%></td>
		<td>｜<a href="Author_Edit.asp?Author_ID=<%=Rs(0)%>" target="tabWin" onclick="fnClick();">修改</a>｜<a href="Author_Del.asp?Author_ID=<%=Rs(0)%>&Page=<%=Page%>">删除</a>｜</td>
	</tr>
	<%
			Rs.MoveNext
    	if rs.eof then exit for
		next
		
		if j < Rs.PageSize then
		  for i=1 to Rs.PageSize-j
				if i mod 2 = 0 then 
					TrBgColor	= ""
				else 
					TrBgColor	= "bgcolor=""#FFFFFF"""
				end if
	%>
	<tr <%=trbgcolor%> height="22">
		<td width="71%">　</td>
		<td width="9%">　</td>
		<td width="12%">　</td>
	</tr>
	<%	
			next
		end if
		if TrBgColor="" then
			TrBgColor	= "bgcolor=""#FFFFFF"""
		else
			TrBgColor = ""
		end if
	%>
	<tr <%=trbgcolor%> height="30" valign="bottom">
		<form method="Post" action="<%=ThisPage%>" style="MARGIN-BOTTOM:0px">
			<td colspan="3" align="center">
			<%
			  if Page<2 then
			%>
				｜&nbsp;<strike>首页</strike>&nbsp;｜
				<strike>上一页</strike>&nbsp;｜
			<%else%>
				｜&nbsp;<a href="<%=ThisPage%>&page=1">首页</a>&nbsp;｜
				<a href="<%=ThisPage%>&page=<%=Page-1%>">上一页</a>&nbsp;｜
			<%
				end if
				if rs.pagecount-page<1 then
			%>
				<strike>下一页</strike>
				｜
				<strike>尾页</strike>
				｜
			<%else%>
				<a href="<%=ThisPage%>&page=<%=page+1%>">下一页</a>&nbsp;｜
				<a href="<%=ThisPage%>&page=<%=rs.pagecount%>">尾页</a>
				｜
			<%end if%>
			页次：<strong><font color=red><%=Page%></font>/<%=rs.pagecount%></strong>&nbsp;页&nbsp;｜
			共<b><font color="#FF0000"><%=rs.recordcount%></font></b>条记录&nbsp;｜
			<b><%=rs.pagesize%></b>条记录/页&nbsp;｜
			转到：<input type="text" name="page" size=4 maxlength=10 class="InputBox" value="<%=page%>">
			<input class="InputBox" type="submit"  value=" Goto (Alt + G) "  name="cndok" accesskey="G">
		<%
	end if
	Call CloseRs
	Call CloseConn
%> </td>
		</form>
	</tr>
</table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->