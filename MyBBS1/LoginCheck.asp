<!-- #include file="Inc/Conn.asp"-->
<!-- #include file="Inc/MD5.asp"-->
<!-- #include file="Inc/Function.asp"-->
<!-- #include file="Inc/Config.asp"-->
<%
  dim UserName, Password, CookiesTime
  dim ErrorMessage, ErrorStatus

  ErrorMessage	= ""
  ErrorStatus	= 0
  
  UserName		= CheckStr(Request("UserName"))
  Password		= CheckStr(Request("Password"))
  CookiesTime	= CheckStr(Request("CookiesTime"))
  
  sql= "Select * From [MyBBS_Users] Where [UserName]='" & UserName & "'"
  rs.open sql, conn, 1, 3
  if err.number <> 0 then
    response.write "���ݿ����ʧ�ܣ����Ժ����ԣ�"
	response.end
  else
    if not rs.eof and not rs.bof then
      if rs("Password") = md5(Password, 32) then
        Response.cookies(MyBBSCookiesName)("UserName")	= Rs("UserName")
        Response.cookies(MyBBSCookiesName)("Password")	= Rs("Password")
		if not cstr(CookiesTime)="0" then
		  CookiesTime	= dateadd("d",CookiesTime,date())
		  Response.cookies(MyBBSCookiesName).expires	= CookiesTime
		else
		  CookiesTime	= date 
		end if
	  else
	    ErrorMessage	= "<li>�������"
	    ErrorStatus	= 1
	  end if
	else
	    ErrorMessage	= "<li>�û�������"
	    ErrorStatus	= 1
	end if
  end if
  rs.close
		
  if ErrorStatus = 0 then 
    Response.write request.cookies(MyBBSCookiesName)("UserName") & "��½�ɹ�����<a href='index.asp'>����</a>��������ϱ��浽" & CookiesTime  
  else
    Response.write "��½ʧ�ܣ�����ԭ�����£�<p>"
    Response.write ErrorMessage
    Response.write "<p><a href='javascript:history.back();'>�뷵��</a>"
  end if		
%>
<!-- #include file="BoardTime.asp"-->