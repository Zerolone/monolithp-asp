<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" --><!-- #include virtual ="/Manage/Check.asp"-->

<%
	Dim Vote_ID
	Vote_ID = Request("Vote_ID")
	CheckNum Vote_ID
	
	Dim Page
	Page		= Request("Page")

	Sql = "Delete From [Monolith_Vote] Where [Vote_ID]=" & Vote_ID
	Conn.Execute(Sql)
	Call CloseConn

	Response.Redirect "Default.asp?Page=" & Page
%>