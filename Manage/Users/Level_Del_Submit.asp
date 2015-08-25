<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim Level_ID, Cmd
	Level_ID = Request("Level_ID")
	CheckNum Level_ID
	
	Dim Page
	Page	= Request("Page")
	
	Sql = "Delete From [Monolith_Users_Level] Where [Level_ID]=" & Level_ID
	Conn.Execute(Sql)
	Call CloseConn

	Response.Redirect "Level.asp?Page=" & Page
%>