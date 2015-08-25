<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<%
	Dim Info_Users_Level_Arr, i

	Dim TheString, CanView
	TheString		= ""
	CanView			= "false"
	
	Dim Users_Level
	Users_Level	= Request.Cookies(Monolith_CookiesName)("Users_Level")
	Users_Level	= Cint(Users_Level)
	
	Dim Info_ID
	Info_ID			= Request("Info_ID")
	CheckNum(Info_ID)
	
	Dim Info_Class
	Info_Class	= Request("Info_Class")

	if Users_Level = "" then
	  TheString = "只有会员才可以察看，<a href='/Login.asp' target='_blank' >点击这里登陆</a>"
	else
		if Info_Class <> "" then
			Sql	= "Select [Info_Users_Level] From [Monolith_Info_Class] Where [Info_Class_Level]='" & Info_Class & "'"
			Rs.Open Sql, Conn, 1, 1
			if Not(Rs.Bof Or Rs.Eof) then
				Info_Users_Level_Arr	= Split(Rs(0), ",")
				For i = LBound(Info_Users_Level_Arr) to UBound(Info_Users_Level_Arr)
					if Info_Users_Level_Arr(i) = 0 or Cint(Info_Users_Level_Arr(i)) = Users_Level then CanView = "True"
				Next
			end if
			Rs.Close
		
			if CanView then
				Sql	= "Select [Info_Content] From [Monolith_Info] Where [Info_ID]=" & Info_ID & " And [Info_Class]='" & Info_Class & "'"
				Rs.Open Sql, Conn , 1, 1
				if Not(Rs.Bof Or Rs.Eof) then TheString = MonolithEncodeII(Rs(0))
				Rs.Close
			else
			  TheString = "你的会员权限不能察看。"
			end if
		end if
	end if
%>
document.write("<%=TheString%>");