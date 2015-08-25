<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual = "/Manage/News/NewsSetting.asp" --><!-- #include virtual ="/Manage/Check.asp"-->

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>无标题文档</title>
<script language="javascript" src="/Js/All.js"></script>
<script language="JavaScript">
function cform(){
 if(!confirm("是否要删除新闻！"))
 return false;
}
</script>
</head>

<body>

<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="4"><b><font color="#FFFFFF">&nbsp;新闻管理 &gt;&gt; 新闻察看</font></b></td>
	</tr>
	<tr height="20" bgcolor="#FFFFFF">
		<td width="71%" bgcolor="#999999"><font color="#FFFFFF">&nbsp;新闻标题</font></td>
		<td width="9%" bgcolor="#999999"><font color="#FFFFFF">&nbsp;栏目编号</font></td>
		<td width="6%" bgcolor="#999999"><font color="#FFFFFF">&nbsp;作者</font></td>
		<td width="12%" bgcolor="#999999"><font color="#FFFFFF">&nbsp;操作</font></td>
	</tr>
	<%
		Dim News_Class, News_Class_Len, News_Date_Arr, News_Url, News_Area, News_Area_Len, News_Area_Code, i
		News_Class						= Request("News_Class")
		News_Area							= Request("News_Area")
		News_Class_Len				= Len(News_Class)		
		
		if News_Area <> "" then
			News_Area_Len					= Len(News_Area)
			News_Area_Code					= 1
			For i=2 to News_Area_Len
				News_Area_Code	= 10 * News_Area_Code
			Next
		end if
		
		Dim ThisPage , Page, Pages, j
	  ThisPage							= Request.ServerVariables("PATH_INFO")
		Page									= clng(request("Page"))
		Dim TrBgColor
		'-----------------0------------1-------------2--------------3--------------4--------------5
		Sql = "Select [News_ID], [News_Title], [News_Class], [News_Author], [News_FileName], [News_Date] From [Monolith_News] Where [News_Apply]=1 "
		if News_Class	<> "" then Sql = Sql + " And Left([News_Class]," & News_Class_Len & ") ='" & News_Class & "' "
		if News_Area	<> "" then Sql = Sql + " And Right([News_Area], " & News_Area_Len & ") >= " & News_Area_Code

		if News_Area	<> "" then
			Sql = Sql + " Order By [News_Order] Asc"
		else
			Sql	= Sql + " Order By [News_ID] Desc"
		end if
		
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

				News_Date_Arr = Split(Rs(5), "-")
				News_Url = NewsSettingArr(7) & "/" & News_Date_Arr(0) & "/" & News_Date_Arr(1) & News_Date_Arr(2) & "/" & Rs(4)
				
		%>
	<tr <%=trbgcolor%>>
		<td height="22" width="71%">&nbsp;<a href="<%=News_Url%>" title="<%=Rs(1)%>" target="_blank"><%=CutLongStrDot(Rs(1), 60, 1)%></a></td>
		<td height="22">&nbsp;<%=Rs(2)%></td>
		<td height="22">&nbsp;<%=Rs(3)%></td>
		<td>｜<a href="#" onclick="ffClick('News/Add.asp?News_ID=<%=Rs(0)%>', '<%=Rs(0)%>');">修改</a>｜<a onclick="return cform();" href="News_Del.asp?News_ID=<%=Rs(0)%>&Page=<%=Page%>&News_Class=<%=News_Class%>">删除</a>｜<a href="#" onclick="ffClick('News/News_ToHtml.asp?News_ID=<%=Rs(0)%>&amp;Page=<%=Page%>', '<%="生成"&Rs(0)%>');">HTML</a>｜</td>
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
		<td width="6%">　</td>
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
			<input type="hidden" name="News_Class" value="<%=News_Class%>">
			<input type="hidden" name="News_Area" value="<%=News_Area%>">
			<td colspan="4" align="center">
			<%
			  ThisPage = ThisPage & "?1=1"
			  if News_Class	<> "" then ThisPage = ThisPage + "&News_Class=" + News_Class
			  if News_Area	<> "" then ThisPage = ThisPage + "&News_Area="  + News_Area
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