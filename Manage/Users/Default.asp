<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>无标题文档</title>
</head>
<script language="javascript" src="/Js/All.js"></script>
<body>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="5"><b><font color="#FFFFFF">&nbsp;会员管理 &gt;&gt; 会员察看</font></b></td>
	</tr>  
	<tr height="20">
    <td bgcolor="#999999"><font color="#FFFFFF">&nbsp;用户名</font></td>
    <td width="1" bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;用户权限</nobr></font></td>
    <td width="1" bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;最后登陆时间</nobr></font></td>
    <td width="1" bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;是否激活</nobr></font></td>
    <td width="50" bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;操作</nobr></font></td>
  </tr>
	<%
		Dim Page, Pages, j, ThisPage
	  ThisPage							= Request.ServerVariables("PATH_INFO")
		Page									= clng(request("Page"))
		Dim i, TrBgColor
		'-------------------0-----------1-------------2--------------3------------------4
		Sql = "Select [Users_ID], [Users_Name], [Users_Group], [Users_LastLogin], [Users_Active] From [Monolith_Users] Order By [Users_ID] Desc"
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
  <tr  <%=TrBgColor%> height="21">
    <td width="18%">&nbsp;<%=Rs(1)%></td>
    <td>&nbsp;<%=Rs(2)%></td>
    <td>&nbsp;<%=Rs(3)%></td>
    <td>&nbsp;<%=Rs(4)%></td>
    <td>｜<a href="Edit.asp?Users_ID=<%=Rs(0)%>&Page=<%=Page%>" target="tabWin" onclick="fnClick();">修改</a>｜<a href="Del.asp?Users_ID=<%=Rs(0)%>&Page=<%=Page%>">删除</a>｜</td>
  </tr>
	<%
    Rs.MoveNext
    if rs.eof then exit for
  next

		Dim ColorMod
		
		if TrBgColor="" then
			ColorMod=0
		else
			ColorMod=1
		end if	

		if j < Rs.PageSize then
		  for i=1 to Rs.PageSize-j
				if i mod 2 = ColorMod then 
					TrBgColor	= ""
				else 
					TrBgColor	= "bgcolor=""#FFFFFF"""
				end if
	%>
	<tr <%=trbgcolor%> height="22">
		<td width="18%" >　</td>
		<td >　</td>
		<td >　</td>
		<td >　</td>
		<td >　</td>
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
			<td colspan="5" align="center">
			<%
			  ThisPage = ThisPage & "?1=1"
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