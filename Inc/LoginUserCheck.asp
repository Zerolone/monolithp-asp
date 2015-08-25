<%
	Dim Users_Name, Users_Password, Users_Group, Users_Level, Users_RegTime, Users_Subject, Users_Sex, IsLogin, Users_ID
	
	IsLogin			= 0
	
	if Request.Cookies.count<>0 then
	  Users_Name			= request.Cookies(Monolith_CookiesName)("Users_Name")
	  Users_Password	= request.Cookies(Monolith_CookiesName)("Users_Password")
	
		Users_Name			= Monolith_Encode(Users_Name)
		Users_Password	= Monolith_Encode(Users_Password)
		
		'------------------0--------------1--------------2----------------3----------------4------------5
	  Sql= "Select [Users_Group], [Users_Level], [Users_RegTime], [Users_Subject], [Users_Sex], [Users_ID] From [Monolith_Users] where [Users_Name] = '" & Users_Name & "' and [Users_Password]='" & Users_Password & "'"
		Rs.Open Sql, Conn, 1, 1
		if Not (Rs.Bof Or Rs.Eof) then
			Users_Group			= Rs(0)
			Users_Level			= Rs(1)
			Users_RegTime		= Rs(2)
			Users_Subject		= Rs(3)
			Users_Sex				= Rs(4)
			Users_ID				= Rs(5)
			IsLogin					= 1
  	end if
  	Rs.Close
  	
 	  'SqlÖ´ÐÐ´ÎÊý+1
	  SqlQueryNum	= SqlQueryNum + 1
	end if
%>