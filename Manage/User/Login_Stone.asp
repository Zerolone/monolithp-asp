<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/MD5.asp"--><!-- #include virtual ="/Manage/Check.asp"-->
<%
  Dim User_Name, User_Password, Code, CanLogin

  Code=Request("Code")
	CanLogin=false
	
  Select case Code
    Case "Exit"
		  Session("User_Name")			= ""
		  Session("User_Password")	= ""
	  	Session("User_Login")		= 0
		  Session.Abandon
		Case "Login"
			User_Name			= Replace(Request("User_Name"),"'","''")
			User_Password	= Replace(Request("User_Password"),"'","''")
	  	User_Password	= MD5(User_Password,32)
	
		  Sql = "Select [User_Name] From [Monolith_User] where [User_Name] ='" & User_Name & "' And [User_Password] = '" & User_Password & "'"
			Rs.Open Sql, Conn, 1, 3
			If Rs.Bof Or Rs.Eof then
				Session("User_Name")			= ""
		    Session("User_Password")	= ""
				Session("User_Login")		= 0
	 	    Session.Abandon
		 else
			  Session("User_Name")			= User_Name
				Session("User_Password")	= User_Password
				Session("User_Login")			= 1
				CanLogin									= true
		  end if
	case else
	end Select

	if CanLogin then
		Response.write "µÇÂ½³É¹¦"
	else
		Response.write "µÇÂ½Ê§°Ü"
	end if
%>