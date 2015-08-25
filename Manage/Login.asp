<!--#include virtual ="/Inc/Monolith_Conn.asp" -->
<!--#include virtual="/Inc/MD5.asp"-->
<!--#include virtual="/inc/Monolith_Function.asp"-->
<!--#include virtual="/Manage/Check_WB.asp"-->
<%
  Dim Code, IP, Admin_Code
  '检验ip
  IP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
  if IP = ""  then IP=Request.ServerVariables("remote_addr")
  if instr(ip,"'")>0 then response.end '非法ip
  
  Code=Request.QueryString("Code")
  Select case Code
    Case "Exit"
	  Session("Admin_Name")		= ""
	  Session("Admin_Password")	= ""
	  Session("Admin_Login")		= 0
	  Session.Abandon
	  Response.redirect "Login.asp"
	
	case "Login"
	  Dim Admin_Name, Admin_Password
	  '防止远程提交
	  'if StrComp(Request.ServerVariables("http_referer"),"http://"&Request.ServerVariables("server_name")&"/Login.asp") <> 0 then response.end
	  Admin_Name		= Monolith_Encode(Request.Form("Admin_Name"))
	  Admin_Code		= Monolith_Encode(Request.Form("Admin_Code"))
	  Admin_Password= Monolith_Encode(Request.Form("Admin_Password"))
	  Admin_Password= MD5(Admin_Password,32)
	  
	  '检验
	  if Not Pass_Name(Admin_Name) Or Not CheckNumII(Admin_Code) then call ErrMsg("您输入的用户名不存在或者密码错误或者随机验证码错误！","Login.asp")
	  
	  Sql = "Select [Admin_Name] From [Monolith_Admin] where [Admin_Name] ='" & Admin_Name & "' And [Admin_Password] = '" & Admin_Password & "'"
	  Rs.Open Sql, Conn, 1, 3
	  If Rs.Bof Or Rs.Eof or Session("Admin_Code")<>Admin_Code or not CheckNumII(Session("Admin_Code")) then
	    Session("Admin_Name")			= ""
	    Session("Admin_Password")	= ""
			Session("Admin_Login")		= 0
 	    Session.Abandon
			Response.Redirect "Login.asp"
	  else
	    Session("Admin_Name")			= Admin_Name
			Session("Admin_Password")	= Admin_Password
			Session("Admin_Login")		= 1
	    Response.Redirect "Default.asp"
	  end if
 
	case""
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Monolith后台管理系统 &gt;&gt; 管理员登陆</title>
</head>

<body bgcolor="#2E5984">

<form method="POST" action="Login.asp?Code=Login">
	<div align="center">
		<table border="0" width="100%" id="table1" height="100%" cellspacing="0" cellpadding="0">
			<tr>
				<td>
				<div align="center">
					<table border="0" width="418" height="256" id="table2" background="Images/Admin_Login.gif" cellspacing="0" cellpadding="0">
						<tr>
							<td width="418" colspan="3" height="70">
							<p align="center">
							<font face="Impact" size="6" color="#FF6633">
							Monolith System Login</font></p></td>
						</tr>
						<tr>
							<td width="418" colspan="3">
							<p>&nbsp; 默认帐号：Zerolone</p>
							<p>&nbsp; 默认密码：Zerolone(<font color="#FF0000">注意大小写</font>)<p>　</td>
						</tr>
						<tr>
							<td height="75" width="125" rowspan="3">
							<p align="center">
							<font face="黑体" color="#FFFFFF" size="4">管理员登陆</font></p>
							</td>
							<td height="11" width="205">
							</td>
							<td height="75" width="88" rowspan="3">
							<p align="center">
							<input border="0" src="Images/Admin_Login_Button.gif" name="I1" width="54" height="40" type="image" class="SubmitButton"></p>
							</td>
						</tr>
						<tr>
							<td height="18" width="205">
							<font face="黑体" color="#FFFFFF" size="3">用户名：</font><input name="Admin_Name" type="text" style="border:1px solid #000000; width=80;height=16" value="Zerolone"> <img src=/Inc/code.asp></td>
							</tr>
						<tr>
							<td height="34" width="205">
							<font face="黑体" color="#FFFFFF"  size="3">密&nbsp; 码：</font><input name="Admin_Password" type="password" style="border:1px solid #000000; width=80;height=16" value="Zerolone"> 
							<input name="Admin_Code" type="text" style="border:1px solid #000000; width=40;;height=16"></td>
						</tr>
					</table>
				</div>
				</td>
			</tr>
		</table>
	</div>
</form>

</body>

</html>
<%end Select%>