<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<script language="javascript" src="/Js/All.js"></script>
<script language="javascript">
<!--
function ConfirmDel()
{
   if(confirm("确定要删除此评论吗？"))
     return true;
   else
     return false;
}
-->
</script>
<body>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="4"><b><font color="#FFFFFF">&nbsp;新闻评论管理</font></b></td>
	</tr>
	  <tr height="20" bgcolor="#FFFFFF">
    <td width="65%"  bgcolor="#999999"><font color="#FFFFFF">&nbsp;所属新闻标题</font></td>
    <td width="25%"  bgcolor="#999999"><font color="#FFFFFF">&nbsp;评论人</font></td>
    <td width="9%"  bgcolor="#999999"><font color="#FFFFFF">&nbsp;操作</font></td>
  </tr>
	<%
		Dim ThisPage , Page, Pages, j
	  ThisPage							= Request.ServerVariables("PATH_INFO")
		Page									= Request("Page")
		Dim i, TrBgColor
		'-----------------------------------0-----------------------------------1
		Sql = "Select [Monolith_News_Remark].[Remark_ID], [Monolith_News_Remark].[Remark_UserName]," 
		'----------------------------2
		Sql = Sql & " [Monolith_News].[News_Title]"
		Sql = Sql & " From [Monolith_News], [Monolith_News_Remark]"
		Sql = Sql & " Where [Monolith_News_Remark].[News_ID]=[Monolith_News].[News_ID]"
		Sql = Sql & " Order By [Monolith_News].[News_ID] Desc, [Monolith_News_Remark].[Remark_ID] Desc"
		Rs.Open Sql, Conn, 1, 1
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
  <tr  <%=TrBgColor%> height="22">
    <td width="*">&nbsp;<%=Rs(2)%></td>
    <td>&nbsp;<%=Rs(1)%></td>
    <td>｜<a href="View.asp?Remark_ID=<%=Rs(0)%>&Page=<%=Page%>">察看</a>｜<a href="View.asp?Remark_ID=<%=Rs(0)%>&Page=<%=Page%>">删除</a>｜</td>
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
		<td width="65%">　</td>
		<td width="25%">　</td>
		<td width="9%">　</td>
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
			<td colspan="4" align="center">
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