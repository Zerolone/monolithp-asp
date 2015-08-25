<!-- #include file ="../Inc/Monolith_Conn.asp" -->
<!-- #include file ="../Inc/MD5.asp" -->
<!-- #include file ="../Inc/Monolith_Function.asp" -->
<%
  Dim Users_Name, Users_Password, Users_Email, Users_ID
  Dim ErrorStatus
  ErrorStatus			= 1

  Users_Name			= CheckStr(Request("Users_Name"))
  Users_Password	= CheckStr(Request("Users_Password"))
  Users_Email			= CheckStr(Request("Users_Email"))
  
  if Users_Name = "" or Users_Password = "" or Users_Email = "" then
  	ErrorStatus = 0
  else
  	Sql="Select [Users_Name] From [Monolith_Users] Where [Users_Name] = '" & Users_Name & "'"
  	Rs.open Sql, Conn, 1, 1
  	if not (rs.eof and rs.bof) then ErrorStatus		= 0
  	rs.close
		if ErrorStatus = 1 then 
			'-----------------0-------------1-----------------2----------------3------------------4--------------5-------------6-------------7--------------8
    	sql="Select [Users_Name], [Users_PassWord], [Users_Subject], [Users_LastLogin], [Users_Group], [Users_Face], [Users_From], [Users_Email], [Users_ID] From [Monolith_Users]"
			rs.open sql, conn, 1, 3
			rs.addnew
			rs(0)		= Users_Name
			rs(1)		= md5(Users_Password,32)
			rs(2)		= 0
			rs(3)		= now
			rs(4)		= "初级会员"
			rs(5)		= ""
			rs(6)		= ""
			rs(7)		= Users_Email
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
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="refresh" content="3; url=Default.asp" />
<title>MyBBS (Power By Monolith)</title>
<link href="../Css/BBS.css" rel="stylesheet" type="text/css">
</head>

<body>

<table border="0" width="100%" height="100%">
	<tr>
		<td align="center">　<table width="50%" cellspacing="0" cellpadding="0" style="border: 1px solid #000000">
			<tr>
				<td class="pformstrip"><b>&nbsp;感谢!</b></td>
			</tr>
			<tr>
				<td height="62" class="row4">
				<p>&nbsp;注册成功， 请登陆<br>
&nbsp;请稍侯,正在为你转换页面...<b> </b></p>
				</td>
			</tr>
			<tr>
				<td class="pformstrip" height="28" align="center">(<a href="Default.asp">如果你不想等待请按这里</a>) 
				　</td>
			</tr>
		</table>
		</td>
	</tr>
</table>

</body>

</html>
<%  else
  	Response.redirect "Reg2.asp"
  end if		
%>