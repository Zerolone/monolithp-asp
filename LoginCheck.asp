<!-- #include file ="Inc/Monolith_Conn.asp" -->
<!-- #include file ="Inc/Monolith_Function.asp" -->
<!-- #include file ="Inc/MD5.asp" --><%
  Dim Users_Name, Users_Password, CookiesTime
  Dim ErrorMessage, ErrorStatus

  ErrorStatus	= 0
  
  Users_Name				= CheckStr(Request("Users_Name"))
  Users_Password		= CheckStr(Request("Users_Password"))
  CookiesTime				= CheckStr(Request("CookiesTime"))
  
  '------------------0-----------1-------------2-----------------3------------------4
  Sql= "Select [Users_ID], [Users_Name], [Users_Password], [Users_LastLogin], [Users_Level] From [Monolith_Users] Where [Users_Name]='" & Users_Name & "'"
  Rs.Open Sql, Conn, 1, 3
  if err.number <> 0 then
    response.write "数据库操作失败，请稍候重试！"
		response.end
  else
    if Not(Rs.Eof Or Rs.Bof) then
      if Rs(2) = md5(Users_Password, 32) then
        Response.cookies(Monolith_CookiesName)("Users_ID")				= Rs(0)
        Response.cookies(Monolith_CookiesName)("Users_Name")			= Rs(1)
        Response.cookies(Monolith_CookiesName)("Users_Password")	= Rs(2)
        Response.cookies(Monolith_CookiesName)("Users_LastLogin")	= Rs(3)
        Response.cookies(Monolith_CookiesName)("Users_Level")			= Rs(4)

				if CookiesTime = "1" then
		  		CookiesTime	= dateadd("d",365,date())
		  		Response.cookies(Monolith_CookiesName).expires	= CookiesTime
				end if
				ErrorStatus = 1
				Rs(3)	= Now()
				Rs.Update
			end if
		end if
  end if
  Call CloseRs
  Call CloseConn
		
  if ErrorStatus = 1 then 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="refresh" content="3; url=Default.asp" />
<title>Monolith Portal Beta v1.0</title>
<link href="Css/Monolith.css" rel="stylesheet" type="text/css">
</head>

<body background="images/bg.GIF">

<table border="0" width="100%" height="100%">
	<tr>
		<td align="center">　<table width="50%" cellspacing="0" cellpadding="0" style="border: 1px solid #000000">
			<tr>
				<td class="Board_Foot" height="25" ><b>&nbsp;感谢!</b></td>
			</tr>
			<tr>
				<td height="57" class="row1">
				<p>&nbsp;你现在登陆为:<%=Request.Cookies(Monolith_CookiesName)("Users_Name")%><br>
&nbsp;请稍侯,正在为你转换页面...<b> </b></p>
				</td>
			</tr>
			<tr>
				<td class="Board_Foot" height="28" align="center">(<a href="Default.asp">如果你不想等待请按这里</a>) 
				　</td>
			</tr>
		</table>
		</td>
	</tr>
</table>
</body>
</html>
<%else%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="refresh" content="3; url=Login.asp" />
<title>Monolith Portal Beta v1.0</title>
<link href="Css/Monolith.css" rel="stylesheet" type="text/css">
</head>

<body>

<table border="0" width="100%" height="100%">
	<tr>
		<td align="center">　<table width="50%" cellspacing="0" cellpadding="0" style="border: 1px solid #000000">
			<tr>
				<td class="Board_Foot" height="25" ><b>&nbsp;登陆失败!</b></td>
			</tr>
			<tr>
				<td height="57" class="row1">
				<p>&nbsp;请正确输入你的登陆信息,<br>
&nbsp;请稍侯,正在为你转换页面...<b> </b></p>
				</td>
			</tr>
			<tr>
				<td class="Board_Foot" height="28" align="center">(<a href="Login.asp">如果你不想等待请按这里</a>) 
				　</td>
			</tr>
		</table>
		</td>
	</tr>
</table>
</body>
</html>
<%end if%>