<!-- #include virtual="/Head.asp"-->
<!-- #include virtual="/Inc/Monolith_Function.asp"-->
<!-- #include virtual="/Inc/LoginUserCheck.asp"-->
<%
  Dim TheString

	if Users_ID = "" then 
		Response.write "ÇëÏÈµÇÂ½£¡"
		Response.end
	end if

 	TheString =	Monolith_Template_10
		'------------------0--------------1---------------2-------------3----------------4-----------5------------6-------------7
	  Sql= "Select [Users_BYear], [Users_BMonth], [Users_BDay], [Users_WebSite], [Users_QQ], [Users_Msn], [Users_From], [Users_Interest] From [Monolith_Users] where [Users_ID] = " & Users_ID
		Rs.Open Sql, Conn, 1, 1
		if Not (Rs.Bof Or Rs.Eof) then
			TheString = Replace(TheString , "{Users_BYear}", 		Rs(0)) 
			TheString = Replace(TheString , "{Users_BMonth}", 	Rs(1)) 
			TheString = Replace(TheString , "{Users_BDay}", 		Rs(2)) 
 			TheString = Replace(TheString , "{Users_WebSite}", 	Rs(3)) 
 			TheString = Replace(TheString , "{Users_QQ}", 			Rs(4)) 
 			TheString = Replace(TheString , "{Users_Msn}", 			Rs(5)) 
 			TheString = Replace(TheString , "{Users_From}", 		Rs(6)) 
 			TheString = Replace(TheString , "{Users_Interest}",	Rs(7)) 
		end if
	Response.write TheString

  TheString	= Monolith_Template_Foot 
  TheString	= Replace(TheString, "{Now}",	Now())
  TheString	= Replace(TheString, "{ExecuteTime}",	FormatNumber((Timer()-Startime)*1000,5) & "ºÁÃë¡£")
  TheString	= Replace(TheString, "{SqlQueryNum}",	SqlQueryNum)
  Response.write TheString
%>