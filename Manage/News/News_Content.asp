<!-- #include virtual ="/Inc/Monolith_Conn.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim News_ID
	News_ID	= Request("News_ID")
	if News_ID <> "" then
		Sql = "Select [News_Content] From [Monolith_News] Where [News_ID]=" & News_ID
		Rs.Open Sql, Conn, 1, 1
		if Not (Rs.Bof Or Rs.Eof) then
			Response.write Rs(0)
		end if
		Call CloseRs
		Call CloseConn
	end if
%>