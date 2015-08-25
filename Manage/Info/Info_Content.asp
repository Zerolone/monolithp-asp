<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim Info_ID
	Info_ID	= Request("Info_ID")
	if Info_ID <> "" then
		Sql = "Select [Info_Content] From [Monolith_Info] Where [Info_ID]=" & Info_ID
		Rs.Open Sql, Conn, 1, 1
		if Not (Rs.Bof Or Rs.Eof) then
			Response.write Rs(0)
		end if
		Call CloseRs
		Call CloseConn
	end if
%>