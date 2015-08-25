<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" --><!-- #include virtual ="/Manage/Check.asp"--><%
	Dim News_Class
	News_Class						= Request("News_Class")
	Dim ThisPage , Page, Pages, j
  ThisPage							= Request.ServerVariables("PATH_INFO")
	Page									= clng(request("Page"))
	Dim i, TrBgColor
	
	Dim News_ID
	News_ID	= Request("News_ID")
'	Response.write "<hr>" & News_ID & "<hr>"
	if News_ID <> "" and Request("submit") = " 通 过 审 核 (Alt + Y) " then
	  Sql="update [Monolith_News] set [News_Apply]=1 where [News_ID] in (" & News_ID & ")"
	  Conn.Execute(Sql)
'	  Response.write "<script language='javascript'>alert('文章通过审核！');location.href=""" & ThisPage & """;</script>"
	end if
%>
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

function CheckBoxSelect(CheckBoxName, Code)
{
	switch (Code){
	case 1:
	  for(i=0; i<= CheckBoxName.length-1;i++)
	  {
	  	CheckBoxName(i).checked=true;
	  }
	  break;

	case 2:
	  for(i=0; i<= CheckBoxName.length-1;i++)
	  {
	  	CheckBoxName(i).checked=CheckBoxName(i).checked==true?false:true;
	  }
	  break;

	case 3:
	  for(i=0; i<= CheckBoxName.length-1;i++)
	  {
	  	CheckBoxName(i).checked=false;
	  }
	  break;

	 }
}
</script>
</head>

<body>

<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="4"><b><font color="#FFFFFF">&nbsp;新闻审核管理</font></b></td>
	</tr>
	<tr height="20" bgcolor="#FFFFFF">
		<td width="82%" bgcolor="#999999"><font color="#FFFFFF">&nbsp;新闻标题</font></td>
		<td width="8%" bgcolor="#999999"><font color="#FFFFFF">&nbsp;新闻类别</font></td>
		<td width="9%" bgcolor="#999999"><font color="#FFFFFF">&nbsp;新闻作者</font></td>
	</tr>
	<%
		'-----------------0------------1-------------2--------------3
		Sql = "Select [News_ID], [News_Title], [News_Class], [News_Author] From [Monolith_News] Where [News_Apply]=0 "
		if News_Class	<> "" then Sql = Sql + " And [News_Class] Like '" & News_Class & "%' "
		Sql	= Sql + " Order By [News_ID] Desc"
		
		Rs.Open Sql, Conn ,1, 1
		if Not (Rs.bof Or Rs.Eof) then
			Rs.PageSize=17
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
	<form action="<%=ThisPage%>" method="POST" name="form2" style="MARGIN-BOTTOM:0px">
		<tr <%=trbgcolor%>>
			<td height="22" width="82%">
			<input type="checkbox" name="News_ID" value="<%=Rs(0)%>" id="News_ID_<%=Rs(0)%>"><label for="News_ID_<%=Rs(0)%>">
			<%=Rs(1)%></label></td>
			<td width="8%"><%=Rs(2)%></td>
			<td width="9%"><%=Rs(3)%></td>
		</tr>
		<%
			Rs.MoveNext
			i = i +1
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
			<td width="82%">　</td>
			<td width="8%">　</td>
			<td width="6%">　</td>
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
		<tr <%=trbgcolor%>>
			<td colspan="3" height="44">
			<p align="center">
			<input type="button" value=" 全 选 (Alt + A) " onclick="CheckBoxSelect(News_ID, 1)" class="InputBox" accesskey="A">
			<input type="button" value=" 反 选 (Alt + I) " onclick="CheckBoxSelect(News_ID, 2)" class="InputBox" accesskey="I">
			<input type="button" value=" 全 不 选 (Alt + N) " onclick="CheckBoxSelect(News_ID, 3)" class="InputBox" accesskey="N">
			<input name="Submit" type="submit" value=" 通 过 审 核 (Alt + Y) " class="InputBox" accesskey="Y"></p>
			</td>
		</tr>
	</form>
	<%
		if TrBgColor="" then
			TrBgColor	= "bgcolor=""#FFFFFF"""
		else
			TrBgColor = ""
		end if
	%>
	<tr <%=trbgcolor%> height="30" valign="bottom">
		<form method="Post" action="<%=ThisPage%>" style="MARGIN-BOTTOM:0px">
			<input type="hidden" name="News_Class" value="<%=News_Class%>">
			<td colspan="3" align="center"><%
			  ThisPage = ThisPage & "?1=1"
			  if News_Class	<> "" then ThisPage = ThisPage + "&News_Class=" + News_Class
			  if Page<2 then
			%> ｜ <strike>首页</strike> ｜ <strike>上一页</strike> ｜ <%else%> ｜ 
			<a href="<%=ThisPage%>&page=1">首页</a> ｜ 
			<a href="<%=ThisPage%>&page=<%=Page-1%>">上一页</a> ｜ <%
				end if
				if rs.pagecount-page<1 then
			%> <strike>下一页</strike> ｜ <strike>尾页</strike> ｜ <%else%>
			<a href="<%=ThisPage%>&page=<%=page+1%>">下一页</a> ｜ 
			<a href="<%=ThisPage%>&page=<%=rs.pagecount%>">尾页</a> ｜ <%end if%> 页次：<strong><font color="red"><%=Page%></font>/<%=rs.pagecount%></strong> 
			页 ｜ 共<b><font color="#FF0000"><%=rs.recordcount%></font></b>条记录 ｜ <b>
			<%=rs.pagesize%></b>条记录/页 ｜ 转到：<input type="text" name="page" size="4" maxlength="10" class="InputBox" value="<%=page%>">
			<input class="InputBox" type="submit" value=" Goto " name="cndok"> <%
	end if
	Call CloseRs
	Call CloseConn
%> </td>
		</form>
	</tr>
</table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->

</body>

</html>
