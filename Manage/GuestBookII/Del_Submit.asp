<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim GuestBook_ID
	GuestBook_ID = Request("GuestBook_ID")
	CheckNum GuestBook_ID
	
	Dim Page
	Page	= Request("Page")	

	Sql = "Delete From [Monolith_GuestBookII] Where [GuestBook_ID]=" & GuestBook_ID
	Conn.Execute(Sql)
	Call CloseConn

	Response.Redirect "Default.asp?Page=" & Page

%>