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
		<td colspan="5"><b><font color="#FFFFFF">&nbsp;信息管理 &gt;&gt; 信息察看</font></b></td>
	</tr>
	<tr height="20" bgcolor="#FFFFFF">
		<td width="*" bgcolor="#999999"><font color="#FFFFFF">&nbsp;信息标题</font></td>
		<td width="1" bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;栏目编号</nobr></font></td>
		<td width="1" bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;栏目名称</nobr></font></td>
		<td width="1" bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;作者</nobr></font></td>
		<td width="1" bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;操作</nobr></font></td>
	</tr>
	<%
		Dim Info_Class, Info_Class_Len, Info_Date_Arr, Info_Url, Info_Area, Info_Area_Len, Info_Area_Code, i
		Info_Class						= Request("Info_Class")
		Info_Area							= Request("Info_Area")
		Info_Class_Len				= Len(Info_Class)		
		
		if Info_Area <> "" then
			Info_Area_Len					= Len(Info_Area)
			Info_Area_Code					= 1
			For i=2 to Info_Area_Len
				Info_Area_Code	= 10 * Info_Area_Code
			Next
		end if
		
		Dim ThisPage , Page, Pages, j
	  ThisPage							= Request.ServerVariables("PATH_INFO")
		Page									= clng(request("Page"))
		Dim TrBgColor
		'-----------------0------------1-------------2------------3------------------4--------------5----------------6
		Sql = "Select [Info_ID], [Info_Title], [Info_Class], [Info_Class_Name], [Info_Author], [Info_FileName], [Info_Date] From [Monolith_Info] Where [Info_Apply]=1 "
		if Info_Class	<> "" then Sql = Sql + " And Left([Info_Class]," & Info_Class_Len & ") ='" & Info_Class & "' "
		if Info_Area	<> "" then Sql = Sql + " And Right([Info_Area], " & Info_Area_Len & ") >= " & Info_Area_Code

		if Info_Area	<> "" then
			Sql = Sql + " Order By [Info_Order] Asc"
		else
			Sql	= Sql + " Order By [Info_ID] Desc"
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

				Info_Date_Arr = Split(Rs(6), "-")
				Info_Url = NewsSettingArr(7) & "/" & Info_Date_Arr(0) & "/" & Info_Date_Arr(1) & Info_Date_Arr(2) & "/" & Rs(5)
				
		%>
	<tr <%=trbgcolor%>>
		<td height="22">&nbsp;<a href="<%=Info_Url%>" title="<%=Rs(1)%>" target="_blank"><%=CutLongStrDot(Rs(1), 60, 1)%></a></td>
		<td height="22">&nbsp;<%=Rs(2)%></td>
		<td height="22">&nbsp;<%=Rs(3)%></td>
		<td height="22">&nbsp;<%=Rs(4)%></td>
		<td>｜<a href="Add.asp?Info_ID=<%=Rs(0)%>" target="tabWin" onclick="fnClick();">修改</a>｜<a onclick="return cform();" href="Info_Del.asp?Info_ID=<%=Rs(0)%>&Page=<%=Page%>&Info_Class=<%=Info_Class%>">删除</a>｜<a href="Info_ToHtml.asp?Info_ID=<%=Rs(0)%>&Page=<%=Page%>" target="tabWin" onclick="fnClick();">HTML</a>｜</td>
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
		<td>　</td>
		<td>　</td>
		<td>　</td>
		<td>　</td>
		<td>　</td>
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
			<input type="hidden" name="Info_Class" value="<%=Info_Class%>">
			<input type="hidden" name="Info_Area" value="<%=Info_Area%>">
			<td colspan="5" align="center">
			<%
			  ThisPage = ThisPage & "?1=1"
			  if Info_Class	<> "" then ThisPage = ThisPage + "&Info_Class=" + Info_Class
			  if Info_Area	<> "" then ThisPage = ThisPage + "&Info_Area="  + Info_Area
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