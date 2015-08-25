<!-- #include virtual="/Inc/Monolith_Conn.asp"-->
<!-- #include virtual="/Inc/Monolith_Function.asp"-->
<!-- #include virtual="/Inc/LoginUserCheck.asp"-->
<%
  Dim TheString

	if Users_ID = "" then 
		Response.write "ÇëÏÈµÇÂ½£¡"
		Response.end
	end if
	
	Dim Users_BYear, Users_BMonth, Users_BDay, Users_WebSite, Users_QQ, Users_Msn, Users_From, Users_Interest
	Users_BYear			= CheckStr(Request("Users_BYear"))
	Users_BMonth		= CheckStr(Request("Users_BMonth"))
	Users_BDay			= CheckStr(Request("Users_BDay"))
	Users_WebSite		= CheckStr(Request("Users_WebSite"))
	Users_QQ				= CheckStr(Request("Users_QQ"))
	Users_Msn				= CheckStr(Request("Users_Msn"))
	Users_From			= CheckStr(Request("Users_From"))
	Users_Interest	= CheckStr(Request("Users_Interest"))
	

	'------------------0--------------1---------------2-------------3----------------4-----------5------------6-------------7
  Sql= "Select [Users_BYear], [Users_BMonth], [Users_BDay], [Users_WebSite], [Users_QQ], [Users_Msn], [Users_From], [Users_Interest] From [Monolith_Users] where [Users_ID] = " & Users_ID
	Rs.Open Sql, Conn, 1, 3
	if Not (Rs.Bof Or Rs.Eof) then
		Rs(0)			= Users_BYear
		Rs(1)			= Users_BMonth
		Rs(2)			= Users_BDay
		Rs(3)			= Users_WebSite
		Rs(4)			= Users_QQ
		Rs(5)			= Users_Msn
		Rs(6)			= Users_From
		Rs(7)			= Users_Interest
		Rs.Update		
	end if
	
	Call CloseRs
	Call CloseConn
	
	Response.Redirect "Personal.asp"
%>