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
		<td colspan="8"><b><font color="#FFFFFF">&nbsp;广告管理</font></b></td>
	</tr>  
	<tr height="20">
    <td bgcolor="#999999" width="18%"><font color="#FFFFFF">&nbsp;广告标题</font></td>
    <td width="3%" bgcolor="#999999"><font color="#FFFFFF">&nbsp;顺序</font></td>
    <td width="18%" bgcolor="#999999"><font color="#FFFFFF">&nbsp;所属</font></td>
    <td width="6%" bgcolor="#999999"><font color="#FFFFFF">&nbsp;点击数</font></td>
    <td width="9%" bgcolor="#999999"><font color="#FFFFFF">&nbsp;尺寸</font></td>
    <td width="25%" bgcolor="#999999"><font color="#FFFFFF">&nbsp;URL</font></td>
    <td width="5%" bgcolor="#999999"><font color="#FFFFFF">&nbsp;使用</font></td>
    <td width="12%" bgcolor="#999999"><font color="#FFFFFF">&nbsp;操作</font></td>
  </tr>
	<%
		Dim Page, Pages, j, ThisPage
	  ThisPage							= Request.ServerVariables("PATH_INFO")
		Page									= clng(request("Page"))
		Dim i, TrBgColor
		'---------------------0----------------------1---------------------2---------------------------3-----------------4------------------------5----------------------6--------------------7-----------------------8
		Sql = "SELECT [Advertisement_ID], [Advertisement_Title], [Advertisement_Order], [Advertisement_Class_ID], [Advertisement_Hits], [Advertisement_Width], [Advertisement_Height], [Advertisement_Url], [Advertisement_Active] FROM [Monolith_Advertisement] Order By [Advertisement_Class_ID] Asc, [Advertisement_Active] DESC, [Advertisement_Order] ASC "
		Rs.Open Sql, Conn ,1, 3
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
    <td>&nbsp;<%=Rs(5)&"*"&Rs(6)%></td>
    <td>&nbsp;<%=Rs(7)%></td>
    <td>&nbsp;<%=Rs(8)%></td>
    <td>｜<a href="View.asp?Advertisement_ID=<%=Rs(0)%>" target="tabWin" onclick="fnClick();">察看</a>｜<a href="Edit.asp?Advertisement_ID=<%=Rs(0)%>&Page=<%=Page%>">修改</a>｜<a href="Del.asp?Advertisement_ID=<%=Rs(0)%>&Page=<%=Page%>">删除</a>｜</td>
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
			<td colspan="8" align="center">
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