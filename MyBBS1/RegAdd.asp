<!-- #include file="Inc/Conn.asp"-->
<!-- #include file="Inc/MD5.asp"-->
<!-- #include file="Inc/Function.asp"-->
<%
  dim UserName, Password
  dim ErrorMessage, ErrorStatus
  ErrorMessage	= ""
  ErrorStatus	= 0

  UserName		= CheckStr(Request("UserName"))
  PassWord		= CheckStr(Request("Password"))
  
  if UserName = "" then
    ErrorMessage 	= ErrorMessage & "<li>输入的帐号为空。"
    ErrorStatus		= 1
  end if

  if Password = "" then
    ErrorMessage 	= ErrorMessage & "<li>输入的密码为空。"
    ErrorStatus		= 1
  end if

  sql="Select * From [MyBBS_Users] Where [Username] = '" & UserName & "'"
  rs.open sql, conn, 1, 3
  if not (rs.eof and rs.bof) then
    ErrorMessage 	= ErrorMessage & "<li>输入的帐号已经被注册。"
    ErrorStatus		= 1
  end if
  rs.close

  if ErrorStatus = 0 then 
    sql="Select * From [MyBBS_Users]"
	rs.open sql, conn, 1, 3
	rs.addnew
	rs("UserName")			= UserName
	rs("PassWord")			= md5(Password,32)
	rs.update
	rs.close
  end if
  
  if ErrorStatus = 0 then 
    Response.write "注册成功，请<a href='index.asp'>返回登陆</a>。"  
  else
    Response.write "注册失败，错误原因如下：<p>"
    Response.write ErrorMessage
    Response.write "<p><a href='javascript:history.back();'>请返回</a>"
  end if
%>
<!-- #include file="BoardTime.asp"-->