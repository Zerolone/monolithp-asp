<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim Bulletin_ID
	Bulletin_ID = Request("Bulletin_ID")
	CheckNum Bulletin_ID
	
	Dim Page
	Page				= Request("Page")
	
	Sql = "Delete From [Monolith_Bulletin] Where [Bulletin_ID]=" & Bulletin_ID
	Conn.Execute(Sql)
	Call CloseConn
	Response.Redirect "Default.asp?Page=" & Page
%>