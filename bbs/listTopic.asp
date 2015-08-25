<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!-- #include file="connection.asp" -->
<!-- #include file="std.asp" -->
<%
'接受参数：
'	request("pageno")	  所要显示的页数
'	request("bid")		  版面ID
'   
	OnlyLogin
	Alive "品味文章"
%>
<%	dim rsBoard,rsBook,strSQL,i,recNo,j,rsPos,dispNo,bid,pgNO,method
	dim preNo,postNo,MaxPage
	dim strMaster
	bid= request("bid")
	pgNo=request("pageNo")
	method=request("method")
	if bid="" then bid="1"
	if pgNo="" then pgNo=1
	if method="" then method="normalt"
%>
<html>
<link rel="stylesheet" type="text/css" href="main.css">
<style type="text/css">
<!--
.black2 {  font-size: 10pt; color: #000000; text-decoration: none; font-family: "幼圆"}
.red2 {  font-family: "黑体"; font-size: 10pt; color: #FF0000; text-decoration: none}
.white2 {  font-family: "幼圆"; font-size: 10pt; color: #FFFFFF; text-decoration: none}
-->
</style>

<body class=clblue background="images/black.gif">
<%	
	set rsBook=Server.CreateObject("ADODB.RecordSet")
	set rsBoard=Server.CreateObject("ADODB.RecordSet")

	if method="new" then strSQL="select * from book order by wdate desc"
	if method="normalt" then strSQL="select * from book where bid="+bid+" and type=0 order by lastdate desc"
	if method="normala" then strSQL="select * from book where bid="+bid+" order by wdate desc"
	rsBook.CursorLocation = 3
    rsBook.CacheSize = session("OnePage")
	rsBook.Open strSQL,myconn,1
	if rsBook.RecordCount >0 then
	rsBook.MoveFirst
	rsBook.AbsolutePage = CInt(pgNo)
	end if
    rsBook.PageSize = session("OnePage")
    MaxPage = rsBook.PageCount
	

	preNo=pgNo-1
	postNo=pgNo+1
	if method<>"new" then
		strSQL="select * from Board where id="+bid
		rsBoard.Open strSQL,myconn
		session("bname")=rsBoard("name")
		session("cname")=rsBoard("cname")
		strMaster=trim(rsBoard("userid"))
		if isNull(strMaster) or len(strMaster)<=2 then
			strMaster="申请中"
		else
			response.write(strmaster)
			strMaster=mid(strMaster,2,len(strMaster)-2)
		end if
		rsBoard.close
	end if
%><br>
<center><div align=center>
<%  if request("method")="new" then %>
	<table width=95% border=0 class="docTopic" >
      <tr>
        <td align=center width =95% class=listtitle height="23"><font color="#FFFFFF"><font color="#000000">新文章列表</font></td>
      </tr></table>
    <%	end if %> 
    <table border=0  cellpadding="0" cellspacing="0" width=95% align=center class="docTopic">
      <tr class="docTopic"> 
        <td align=left class=listtopic width="52%">
          <font  class="white2"> 
          <%  if request("method")<>"new" then %> 
          &nbsp;【<%=session("cname")%>】&nbsp;&nbsp;版主:<%=strMaster%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="post.asp?pmethod=new&bid=<%=bid%>" class=white2 target=frmMain>发表文章</a>&nbsp;&nbsp; 
          <%	end if %>
         </font>
         
        </td>
        <td align=right  class=docTopic width="48%"><font color="#FFFFFF" class="black2"> 
          第<span class="docTopic"><%=pgNo%>页[<%=(pgNo-1)*session("OnePage")+1%>-<%=pgNo*session("OnePage")%>]&nbsp;共<%=Maxpage%>页<%=rsBook.RecordCount%>篇 
          <%if preNo>0 then %> <A href=ListTopic.asp?method=<%=method%>&pageNO=<%=preNo%>&bid=<%=bid%> class=TopicPage1 >[上一页]</a> 
          <%else   %> [上一页] <%end if %> <%if postNo<=MaxPage then %> <A href=ListTopic.asp?method=<%=method%>&pageNO=<%=postNo%>&bid=<%=bid%> class=TopicPage1 >[下一页]</a> 
          <%else   %> [下一页] <%end if %> </span></font></td>
      </tr>
    </table>
    <table width=95% cellspacing=2 align=center bgcolor="#000000" class="clblack">
      <tr> <%if session("ListMethod")=1 and method<>"new" then%> 
        <td width='6%' class="black2" > 
          <center>
            <font class="white2">回帖</font> 
          </center>
        </td>
        <%else%> 
        <td width='18%' class="black2" > 
          <center>
            <font class="white2">讨论区</font> 
          </center>
        </td>
        <%end if%> 
        <td width='58%' > 
          <center>
            <font class="white2">主&nbsp;&nbsp; 题</font> 
          </center>
        </td>
        <td width='14%' > 
          <center>
            <font class="white2">作&nbsp; 者</font> 
          </center>
        </td>
        <td width='5%' > 
          <center>
            <font class="white2">点击</font> 
          </center>
        </td>
        <%if method<>"new" then%> 
        <td width='15%' > 
          <center>
            <font class="white2">最后回应时间</font> 
          </center>
        </td>
        <%end if%> </tr>
      <%	dim amon,aday,ahour,amin,tztime,wdate,rsUserInfo,lastDay,bcname
	lastDay=" "

'	  set rsUserInfo=Server.CreateObject("ADODB.RecordSet")
'	  rsUserInfo.Open "Select * from UserInfo",myconn,1
	j=0
	Do While Not rsBook.EOF And j < rsBook.PageSize
		if j mod 2=1 then %> 
      <tr class=listline1> <%		else			  %> 
      <tr class=bline2> <%		end if			  	
		if method="normalt" then
			tztime=cdate(rsBook("lastDate")) 
		else
			tztime=cdate(rsBook("wdate"))
		end if 
		amon=right("0"+cstr(month(tztime)),2)
		aday=right("0"+cstr(day(tztime)),2)
		ahour=right("0"+cstr(hour(tztime)),2)
		amin=right("0"+cstr(minute(tztime)),2)
		if lastDay<>amon+"."+aday then
			wdate=amon+"."+aday+"-"+ahour+":"+amin
		else
			wdate="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+ahour+":"+amin
		end if
		lastday=amon+"."+aday
		if session("ListMethod")=1 and method<>"new" then
%> 
        <td align="center" > <font color="#FFFFFF"><span class="black2"><%if rsBook("reply")>0 and method<>"new" then%></span><span class="black2"> 
          +<%=right("0"&cstr(rsBook("reply")),2)%></span></font> <span class="black2"><font color="#FFFFFF"><%end if%> 
          </font></span></td>
        <%else 
			bcname=" "
		    rsBoard.Open "select * from board where id="+cstr(rsBook("bid")),myconn,1
			if rsBoard.RecordCount>0 then bcname=rsBoard("cname")%> 
        <td align=center class=listtopic1><font color="#FFFFFF" class="black2">【<%=bcname%>】</font></td>
        <%	rsBoard.Close
		end if %> 
        <td><font color="#FFFFFF">&nbsp;<a class=info href=showdoc.asp?method=<%=method%>&bid=<%=rsBook("bid")%>&pageNo=<%=request("pageNo")%>&id=<%=rsBook("id")%>&reply=<%=rsBook("reply")%>><span class="red2">●</span>&nbsp;<span class="listtopic1"><span class="listtopic1"><span class="listtopic1"><span class="listtopic1"><%=rsBook("Title")%></span></span></span></span></a></font></td>
        <td><span class="black2"><font color="#FFFFFF" class="black2"><%=rsBook("userid")%></font></span></td>
        <td align=right class="black2"><font color="#FFFFFF"><span class="black2"><span class="black2"><span class="black2"><span class="black2"></span></span></span><%=rsBook("hits")%></span></font></td>
        <%if  method<>"new" then%> 
        <td class="black2"><font color="#FFFFFF" class="black2"><span class="black2"><%=wdate%><span class="black2"></span></span></font></td>
        <%end if%> </tr>
      <%
		rsBook.MoveNext
		j=j+1
	  Loop 
	  rsBook.Close
	if j=0 then %> 
      <tr bgcolor='#eeeffe'> <%if session("ListMethod")=1 and method<>"new" then %> 
        <td bgcolor="#000000" class="black2">&nbsp;</td>
        <%end if%> 
        <td bgcolor="#000000" class="black2"> 
          <center>
            <font color="#FFFFFF" class="black2"> ------&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;该板没有任何文章&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;------ 
            </font> 
          </center>
        </td>
        <td bgcolor="#000000" class="black2">&nbsp;</td>
        <td bgcolor="#000000" class="black2">&nbsp;</td>
        <td bgcolor="#000000" class="black2">&nbsp;</td>
      </tr>
      <%	end if %> 
    </table>
</div></center>
<p>
</body>
</html>
<%
	myconn.Close
	set rsBoard=nothing
	set rsBook=nothing
	set myconn=nothing
%>
