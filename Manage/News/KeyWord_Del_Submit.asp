<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim KeyWord_ID
	KeyWord_ID = Request("KeyWord_ID")
	CheckNum KeyWord_ID
	
	Dim Page
	Page	= Request("Page")	

	Sql = "Delete From [Monolith_News_KeyWord] Where [KeyWord_ID]=" & KeyWord_ID
	Conn.Execute(Sql)
	Call CloseConn

	Response.Redirect "KeyWord.asp?Page=" & Page
%>