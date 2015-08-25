<!-- #include virtual ="/Head_NoHead.asp" -->
<!-- #include virtual ="/Inc/MD5.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<%
  Dim Users_Name, Users_Password, Users_Email, Users_ID, Users_Sex
  Dim ErrorStatus
  ErrorStatus			= 1

  Users_Name			= CheckStr(Request("Users_Name"))
  Users_Password	= CheckStr(Request("Users_Password"))
  Users_Email			= CheckStr(Request("Users_Email"))
  Users_Sex				= CheckStr(Request("Users_Sex"))
  
  if Users_Name = "" or Users_Password = "" or Users_Email = "" then
  	ErrorStatus = 0
  else
  	Sql="Select [Users_Name] From [Monolith_Users] Where [Users_Name] = '" & Users_Name & "'"
  	Rs.open Sql, Conn, 1, 1
  	if not (rs.eof and rs.bof) then ErrorStatus		= 0
  	rs.close
		if ErrorStatus = 1 then 
			'-----------------0-------------1-----------------2----------------3------------------4--------------5-------------6-------------7--------------8-----------9----------------10
    	Sql="Select [Users_Name], [Users_PassWord], [Users_Subject], [Users_LastLogin], [Users_Group], [Users_Face], [Users_From], [Users_Email], [Users_ID], [Users_RegTime], [Users_Sex], " 
			'-----------------11-------------12-------------13--------------14------------15---------------16----------17-----------18------------19
    	Sql= Sql & "[Users_Level], [Users_BYear], [Users_BMonth], [Users_BDay], [Users_WebSite], [Users_QQ], [Users_Msn], [Users_From], [Users_Interest] From [Monolith_Users]"
			rs.open sql, conn, 1, 3
			rs.addnew
			rs(0)		= Users_Name
			rs(1)		= md5(Users_Password,32)
			rs(2)		= 0
			rs(3)		= Now()
			rs(4)		= "初级会员"
			rs(5)		= 1
			rs(6)		= ""
			rs(7)		= Users_Email
			rs(9)		= Now()
			rs(10)	= Users_Sex
			rs(11)	= 1
			rs(12)	= 1990
			rs(13)	= 1
			rs(14)	= 1
			rs(15)	= "http://"
			rs(16)	= "10000"
			rs(17)	= "YourName@Msn.com"
			rs(18)	= "中国"
			rs(19)	= "无"
		
			Users_ID	= Rs(8)
			rs.update
			rs.close
  	end if
  end if
  
  if ErrorStatus = 1 then 
 		Sql = "update [Monolith_Stats] set [Stats_UserNumber] = [Stats_UserNumber] + 1"
		Sql = Sql & ", [Stats_NewUser] 			='" & Users_Name & "'"
		Sql = Sql & ", [Stats_NewUserID] 		=" & Users_ID
		Conn.execute(Sql)

		Dim TheString

		TheString	= Monolith_Template_06_Step3
		Response.write TheString
  else
  	Response.redirect "Reg2.asp"
  end if		
%>