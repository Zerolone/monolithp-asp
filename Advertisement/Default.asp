<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<%
	Dim Advertisement_No
	Advertisement_No	= Request("Advertisement_No")
	CheckNum Advertisement_No

	'---------------------------------0--------------------1
	Sql = "Select Top 1 [Advertisement_Url], [Advertisement_Hits] From [Monolith_Advertisement] Where [Advertisement_No]=" & Advertisement_No
	Rs.Open Sql, Conn, 1, 3
	if Not (Rs.Bof Or Rs.Eof) then
		Rs(1) = Rs(1) + 1
		Rs.Update
		Response.Redirect(Trim(Rs(0)))
	else
	  Response.write "这个广告不存在！"
		Response.end
	end if

	Call CloseRs
	Call CloseConn	
%>