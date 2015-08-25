<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/MD5.asp"-->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" --><!-- #include virtual ="/Manage/Check.asp"-->

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<%
  dim Code, CanLogin, Ip
  '检验ip
  IP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
  if IP = ""  then IP=Request.ServerVariables("remote_addr")
  if instr(ip,"'")>0 then response.end '非法ip
  
  Code=Request.QueryString("Code")
  Select case Code
    Case "Exit"
	  Session("User_Name")		= ""
	  Session("User_Password")	= ""
	  Session("User_Login")		= 0
	  Session.Abandon
	  Response.redirect "User_Login.asp"
	
	case "Login"
	  dim User_Name, User_Password, User_Code
	  '防止远程提交
	  'if StrComp(Request.ServerVariables("http_referer"),"http://"&Request.ServerVariables("server_name")&"/Login.asp") <> 0 then Response.End
	  User_Name	= Replace(Request.Form("User_Name"),"'","''")
	  User_Code	= Replace(Request.Form("User_Code"),"'","''")
	  User_Password= Replace(Request.Form("User_Password"),"'","''")
	  User_Password= MD5(User_Password,32)
	  
	  '检验
	  if Not CheckNumII(User_Code) Or Len(User_Code) <> 4 then call errmsg("您输入的随机验证码错误！",16,"Login.asp")
	  
	  Sql = "Select [User_Level] From [Monolith_User] where [User_Name] ='" & User_Name & "' And [User_Password] = '" & User_Password & "'"
	  Rs.Open Sql, Conn, 1, 3
	  If Rs.Bof Or Rs.Eof or Session("User_Code")<>User_Code or not CheckNumII(session("User_Code")) then
 	    CanLogin=false
	    Session("User_Name")			= ""
	    Session("User_Password")	= ""
			Session("User_Login")			= 0
			Session("User_Level")			= ""
 	    Session.Abandon
	  else
	    Session("User_Name")			= User_Name
			Session("User_Password")	= User_Password
			Session("User_Login")			= 1
			Session("User_Level")			= Rs(0)
			
			CanLogin=true

'---------------
		Dim User_Level_Arr, User_Level_Name
		User_Level_Arr	= Split(Session("User_Level"), "|")
		User_Level_Arr(0) = 1
		Select Case User_Level_Arr(0)
			Case "1"
				User_Level_Name = "系统管理员"
			Case "2"
				User_Level_Name = "管理员"
			Case "3"
				User_Level_Name = "新闻系统管理员"
			Case "4"
				User_Level_Name = "广告系统管理员"
			Case else
				User_Level_Name = "游客"
	  End Select
'---------------
			
		end if
  
	  if CanLogin then
	    Response.write "<br><ul>"
	    Response.write "<li>登陆成功！现在您可以进行操作了！</li>"
	    Response.write "<li></li>"
	    Response.write "<li>Session(&quot;User_Session&quot;)=&quot;"& Session("User_Level") &"&quot;</li>"
	    Response.write "<li>您的管理登陆等级是：<a href=Info.asp?User_Level=" & User_Level_Arr(0) & "><font color=red>" & User_Level_Name & "</font></a></li>"
	    Response.write "</ul>"
	  else
	    response.redirect"Login.asp"
	  end if
	case ""
%>
<title>论坛登陆</title>
<table border=1 cellpadding=0 cellspacing=0 width=98%>
  <tr>
    <td width='100%' valign=top>
	<form method=POST action="?Code=Login"><br>
    &nbsp;用户名： 
            <input name=User_Name type=text id="User_Name" value="Zerolone" size=25 class="InputBox" style="width=160">
            &nbsp;·请输入管理员的用户名。<br>
	&nbsp;密&nbsp;&nbsp;码： 
            <input name=User_Password type=password id="User_Password" value="Zerolone" size=25 class="InputBox"  style="width=160;height=18">
            &nbsp;·请输入管理员密码<br>
	&nbsp;验证码： 
            <input name=User_Code type=text id="User_Code" size=25 class="InputBox"  style="width=160">
            &nbsp;·<img src=../Inc/code.asp>&nbsp;<font color=red>·如无法显示或看不清，请刷新</font><br>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<input type=submit value=" 提  交  " name=B1 class="InputBox">
	</form>
	</td>
  </tr>
</table>
<%end Select%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->