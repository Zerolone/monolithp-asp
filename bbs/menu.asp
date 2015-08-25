<% option explicit%>
<!-- #include file="connection.asp" -->
<!-- #include file="std.asp" -->
<%	OnlyLogin  %>
<html>
<head>
<meta http-equiv="pragma" content="no-cache">
<link rel="stylesheet" type="text/css" href="main.css">
</head>
<body class=clblack>
<%	dim rsBoard,strSQL,i,LastType,recNo
	set rsBoard=Server.CreateObject("ADODB.Recordset")
	rsBoard.Open "select Board.id as id,BoardType.id as BtID,Board.name as name,BoardType.name as Btname,cname from Board,BoardType where Board.BoardType=BoardType.id order by BoardType.id,Board.name",myconn,1
	recNo=rsBoard.RecordCount
	LastType=0
%>
<br><br><br>
<p class="menu">
<%	for i=1 to recNo
			if rsBoard("Btid")<>LastType and Application("MenuMethod")=1 then   %>
<img src="image/line_sub.gif" border=0 align=absmiddle><img src="image/folder.gif" border=0 align=absmiddle ><%=rsBoard("Btname")%><br>			
<%			LastType=rsBoard("Btid")
			end if
			if Application("MenuMethod")=1 then
%>
			<img src="image/line_05.gif" border=0 align=absmiddle><img src="image/reddot.gif" border=0 align=absmiddle >
<%			else%>
			&nbsp;&nbsp;¡ñ
<%			end if%>
			<a class=menu href="listTopic.asp?pageno=1&bid=<%=rsBoard("id")%>" target= "frmMain"><%=rsBoard("cname")%></a><br>
<%			rsBoard.MoveNext
	next
	rsBoard.Close
%>
<br><br>
</p>
</body>
</html>