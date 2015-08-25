<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<%
	dim News_ID, Page, Info_Class
	News_ID			= Request("News_ID")
	Page				= Request("Page")
	Info_Class	= Request("Info_Class")
	CheckNum News_ID
	Sql="DELETE FROM [Monolith_Info] where [News_ID]="& News_ID
	Conn.Execute (Sql)
	Call CloseConn
	Response.Redirect "News_List.asp?Page=" & Page & "&Info_Class=" + Info_Class
%>