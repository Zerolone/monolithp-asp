<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!-- #include file="connection.asp" -->
<html >
<body>
<%		dim rsBook,strSQL,iUser,i,byOrder,deleteid,misspassid
		dim rsBoard,rsBook1,strFtitle,strBname
		set rsBook=server.CreateObject("ADODB.RecordSet")
		set rsBook1=server.CreateObject("ADODB.RecordSet")
		set rsBoard=server.CreateObject("ADODB.RecordSet")
		strSQL="select * from Book order by wdate where show=0"
		rsBook.Open strSQL,myconn,1
		iUser=rsBook.RecordCount
		for i=1 to iUser
			rsBoard.Open "select * from Board where id="+cstr(rsBook("bid")),myconn,1
			if rsBoard.RecordCount>0 then 
				strBname=rsBoard("cname")
				rsBoard.Close
				strFtitle=""
				if rsBook("id")<>rsBook("fid") then
					rsBook1.Open "select * from Book where id="+cstr(rsBook("fid")),myconn,1
					if rsBook1.RecordCount>0 then
						strFtitle=rsBook1("title")
					end if
					rsBook1.Close
				end if
%>
<br>==========title:<%=rsBook("title")%>
<br>======content:<%=showbody(rsBook("content"))%>
<br>======bname:<%=showbody(strBname)%>
<br>======ftitle:<%=strFtitle%>
<br>======userid:<%=rsBook("userid")%>
<br>======
<%			end if	
			rsBook.MoveNext
		next %>
</body>
</html>
