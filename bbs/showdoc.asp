<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!-- #include file="connection.asp" -->
<!-- #include file="string.asp" -->
<!-- #include file="std.asp" -->
<%
'接受参数：
'	request("id")	      文章的ID
'	request("bid")		  所在版面ID
'   
	OnlyLogin
	Alive "品味文章"
%>
<html >
<head>
<meta http-equiv="pragma" content="no-cache" >
<link rel="stylesheet" type="text/css" href="main.css">
<style type="text/css">
<!--
.yellow2 {  font-size: 10pt; color: #FFFF66}
-->
</style>
</head>
<body class=clblue background="images/black.gif">
<%	dim rsUserInfo,rsBoard,rsBook,strSQL,i,recNo,j,strSQL1,hitss,boardMaster,method,pgNo,bid
	dim strBoardName,strBoardCname,rsSortBook2,rsSortBook1,strPrePage,strPostPage,strDate
	dim strSQL2,sortID1,sortID2
	pgNo=request("pageNo")
	bid=cstr(request("bid"))
	method=request("method")
	if method="" then method="normalt"
	
	set rsBook=Server.CreateObject("ADODB.RecordSet")
	set rsSortBook1=Server.CreateObject("ADODB.RecordSet")
	set rsSortBook2=Server.CreateObject("ADODB.RecordSet")
	set rsBoard=Server.CreateObject("ADODB.RecordSet")
	set rsUserInfo=Server.CreateObject("ADODB.RecordSet")
	'rsBoard.Open "update board set iclick=iclick+1 where id="+bid,myconn,1
	rsBoard.Open "select * from Board where id="+bid ,myconn,1
	if not rsBoard.EOF and not rsBoard.BOF then
		boardMaster=LCASE(rsBoard("userid"))
		strBoardName=rsBoard("name")
		strBoardCname=rsBoard("cname")
	end if

	if method<>"new" then
		if session("ListMethod")=1 then
			strSQL="select * from book where fid="+request("id")+" order by type asc,wdate asc"
		else
			strSQL="select * from book where id="+request("id")
		end if
	else
		strSQL="select * from book where id="+request("id")
	end if 
	rsBook.Open strSQL,myconn,1
	recNo=rsBook.RecordCount
	if recNo>0 then
	hitss=rsBook("hits")

	if method<>"new" then
		if session("ListMethod")=1 then
			strSQL1="select max(id) as idd from book where type=0 and bid="+bid+" and id<"+cstr(rsBook("id"))
			strSQL2="select min(id) as idd from book where type=0 and bid="+bid+" and id>"+cstr(rsBook("id"))
		else
			strSQL1="select max(id) as idd from book where bid="+bid+" and id<"+cstr(rsBook("id"))
			strSQL2="select min(id) as idd from book where bid="+bid+" and id>"+cstr(rsBook("id"))
		end if
	else
			strSQL1="select max(id) as idd from book where id<"+cstr(rsBook("id"))
			strSQL2="select min(id) as idd from book where id>"+cstr(rsBook("id"))
	end if 
	rsSortBook1.Open strSQL1,myconn,1
	rsSortBook2.Open strSQL2,myconn,1
	if rsSortBook1.RecordCount>0 and not isnull(rsSortBook1("idd")) then
		sortID1=cstr(rsSortBook1("idd"))
		rsSortBook1.close
		rsSortBook1.Open "select bid,id from book where id="+sortID1,myconn,1
		if rssortbook1.recordcount>0 then strPostPage="method="+method+"&bid="+cstr(rsSortBook1("bid"))+"&pageNo="+pgNo+"&id="+ cstr(rsSortBook1("id"))
	end if
	if rsSortBook2.RecordCount>0 and not isnull(rsSortBook2("idd")) then
		sortID2=cstr(rsSortBook2("idd"))
		rsSortBook2.close
	    rsSortBook2.Open "select bid,id from book where id="+sortID2,myconn,1
		if rssortbook2.recordcount>0 then strPrePage="method="+method+"&bid="+cstr(rsSortBook2("bid"))+"&pageNo="+pgNo+"&id="+ cstr(rsSortBook2("id"))
	end if
	rsSortBook1.Close
	rsSortBook2.Close
	if session("ListMethod")=1 and method<>"new" then
		strPrePage=""
		strPostPage=""
	end if 
%> <br>
<span class="docMain"></span>
<table width="95%" border="0" align="center" class=docMain >
  <tr>
     
    <td  width="20%" align="lift" >点击次数：<span class="yellow2"><%=hitss+1%>次</span></td>
     <td  width="40%" align="center" >【讨论区: <%=strBoardCname%>】</td>
     <td  width="40%" align="right" >
<% if method="new" then 
		if strPrePage<>"" then %>
			<a class=showdoc href="showdoc.asp?<%=strPrePage%>">上篇</a>&nbsp;
<%		else %>
		    上篇&nbsp;
<% 		end if	
	else %>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%  end if %>

<% if method="new" then 
		if strPostPage<>"" then %>
			<a class=showdoc href="showdoc.asp?<%=strPostPage%>">下篇</a>&nbsp;
<%		else %>
		    下篇&nbsp;
<% 		end if	
	else %>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%  end if %>
<% if method="bb_hotbook" then%>
<a class=showdoc href="bb_hotbook.asp?hotmethod=<%=request("hotmethod")%>">返回</a><br></td>
<% else%>
	<a class=showdoc href="listTopic.asp?method=<%=method%>&pageNo=<%=pgNo%>&bid=<%=bid%>">返回</a><br></td>
<% end if%>
  </tr>
</table>
<%	if recNo>0 then
	for i=1 to recNo
		rsUserInfo.Open "select * from UserInfo where userid='"+rsBook("userid")+"'",myconn,1
%>
<hr width="95%"  >
<span class="docTopic"></span> <span class="docTopic"></span> <span class="docTopic"></span>
<table width="92%" border="0" align=center class=docTopic >
  <tr> 
    <td align="left" width=75% > 发信人： <% if rsUserInfo.RecordCount>0 then %> <a class=showdoc href="mailto:<%=rsUserInfo("email")%>"><%=rsBook("userid")%>(<%=rsUserInfo("nickname")%>)</a> 
      <% else%> <%=rsBook("userid")%> <% end if 
		rsUserInfo.close
	%> <br>
    </td>
	<td aligh width=25%></td>
  </tr>
  <tr>
    <td align="left" width="75%" >标&nbsp;&nbsp;题<span class="yellow2">： <%=rsBook("title")%></span><br>
    </td>
	<td width="25%"></td>
  </tr>
  <tr>
	<td align="left" width="75%" >发信站：<%=Application("Title")%> <font size="2">(发表于：<%=rsBook("wdate")%>)</font><br>
      	</td>
	<td align="right" width="25%" > <%if ((Ucase(session("userid"))=Ucase(rsBook("Userid"))) and rsBook("userid")<>"guest") or isBMaster(Boardmaster) or issysop(session("userid")) then %> 
      <a class=showdoc href="post.asp?method=<%=method%>&bid=<%=bid%>&fid=<%=request("id")%>&id=<%=rsBook("id")%>&pageNo=<%=pgNo%>&pmethod=DELETE" >删除</a>&nbsp;&nbsp; 
      <a class=showdoc href="post.asp?method=<%=method%>&bid=<%=bid%>&fid=<%=request("id")%>&id=<%=rsBook("id")%>&pageNo=<%=pgNo%>&pmethod=EDIT" >修改</a>&nbsp;&nbsp; 
      <%end if%> <a class=showdoc href="post.asp?method=<%=method%>&bid=<%=bid%>&fid=<%=request("id")%>&id=<%=rsBook("id")%>&pageNo=<%=pgNo%>&pmethod=REPLY" >回复</a> 
    </td>
  </tr>
</table>

<table width=92% border=0 class=docContent align=center  cellpadding="10" cellspacing="1" >
  <tr>
    <td class="yellow2"> <% dim strrContent
	   strrContent=showbody(rsBook("content"))
	   strrContent=Replace(strrContent,"<br>:","<br>"&"<font color=#408080>:")
	   strrContent=Replace(strrContent,"<br>","</font><br>")	
		%> <%=strrContent%> <span class="yellow2"></span> </td>
  </tr>
</table>

<span class="docContent"></span><%  rsBook.Movenext
	next 
	end if
	rsBook.Close
	strSQL="Update book set hits=hits+1 where id="+request("id")
	rsBook.Open strSQL,myconn
%> 
<hr width="95%"  >

<table width="95%" bgcolor="#7AB8FC" border="0" align="center" class=docMain >
  <tr>
     <td  width="20%" align="lift" >点击次数：<%=hitss+1%>次</td>
     
    <td  width="40%" align="center" >讨论区: <%=session("bname")%>[<%=strBoardCname%>]</td>
     <td  width="40%" align="right" >
<% if method="new" then 
		if strPrePage<>"" then %>
			<a class=showdoc href="showdoc.asp?<%=strPrePage%>">上篇</a>&nbsp;
<%		else %>
		    上篇&nbsp;
<% 		end if	
	else %>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%  end if %>
<% if method="new" then 
		if strPostPage<>"" then %>
			<a class=showdoc href="showdoc.asp?<%=strPostPage%>">下篇</a>&nbsp;
<%		else %>
		    下篇&nbsp;
<% 		end if	
	else %>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%  end if %>
<% if method="bb_hotbook" then%>
<a class=showdoc href="bb_hotbook.asp?hotmethod=<%=request("hotmethod")%>">返回</a>
<% else%>
	<a class=showdoc href="listTopic.asp?method=<%=method%>&pageNo=<%=pgNo%>&bid=<%=bid%>">返回</a>
<% end if%>
  <br></td></tr>
</table>
<% end if%>
</body>
</html>
<%
	myconn.Close
	set rsUserInfo=nothing
	set rsBoard=nothing
	set rsBook=nothing
	set myconn=nothing
	set rsSortBook2=nothing
	set rsSortBook1=nothing
%>