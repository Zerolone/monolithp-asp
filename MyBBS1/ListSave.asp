<!-- #include file="Inc/Conn.asp"-->
<!-- #include file="Inc/Config.asp"-->
<!-- #include file="Inc/Function.asp"-->
<!-- #include file="TopMenu.asp"-->
<%
  if Username="" then
    Response.write "���ȵ�¼˵��ע�ᣡ"
	Response.end
  end if

  dim BBSTitle, BBSBody, BBSBoardID, BBSPostUser, BBSPostTime, BBSLastPostUser, BBSLastPostTime, BBSRoot

  dim ThisPage, FormPage
  
  FormPage	= Request("FormPage")
  ThisPage	= Request.ServerVariables("PATH_INFO")
  if Request("Submit") = " �� �� �� �� " then 
    
	BBSTitle			= Trim(Request("BBSTitle"))
	BBSBody				= Trim(Request("BBSBody"))
	BBSBody				= Replace(BBSBody, chr(10) & chr(13) , "<br>")
	BBSBody				= Replace(BBSBody, vbcrlf , "<br>")
    BBSBoardID			= Trim(Request("BBSBoardID"))
	BBSPostUser			= UserName
	BBSPostTime			= Now
	BBSLastPostUser		= UserName
	BBSLastPostTime		= Now
 
	Sql = "Select * From [MyBBS_BBS]"
	rs.open sql, conn, 1, 3
	Rs.AddNew
	Rs("BBSTitle")			= BBSTitle
	Rs("BBSBody")			= BBSBody
	Rs("BBSBoardID")		= BBSBoardID
	Rs("BBSPostUser")		= BBSPostUser
	Rs("BBSPostTime")		= BBSPostTime
	Rs("BBSLastPostTime")	= BBSLastPostTime
	Rs("BBSLastPostUser")	= BBSLastPostUser
	Rs.Update
	Rs.close
	
	Sql = "update [MyBBS_Board] set [BoardTopic] = [BoardTopic] + 1"
	Sql = Sql & ", [BoardLastPostTime] 		='" & BBSLastPostTime & "'"
	Sql = Sql & ", [BoardLastPostUser] 		='" & BBSLastPostUser & "'"
	Sql = Sql & ", [BoardLastPostTopic] 	='" & BBSTitle & "'"
	Sql = Sql & " where [ID]=" & BBSBoardID
	Conn.execute(Sql)
		
    Response.write "<script language='javascript'>alert('��ӳɹ���');location.href='" & FormPage & "';</script>"
  end if
  
  if Request("Submit") = " �� �� �� �� �� �� " then 
    
	BBSTitle			= Server.HTMLEncode(Trim(Request("BBSTitle")))
	BBSBody				= Trim(Request("BBSBody"))
	BBSBody				= Replace(BBSBody, chr(10) & chr(13) , "<br>")
	BBSBody				= Replace(BBSBody, vbcrlf , "<br>")
'	BBSBody				= Server.HTMLEncode(BBSBody)
    BBSBoardID			= Trim(Request("BBSBoardID"))
	BBSPostUser			= UserName
	BBSPostTime			= Now
	BBSLastPostUser		= UserName
	BBSLastPostTime		= Now
	BBSRoot				= Trim(Request("BBSRoot"))
 
	Sql = "Select * From [MyBBS_BBS]"
	rs.open sql, conn, 1, 3
	Rs.AddNew
	Rs("BBSTitle")			= BBSTitle
	Rs("BBSBody")			= BBSBody
	Rs("BBSBoardID")		= BBSBoardID
	Rs("BBSPostUser")		= BBSPostUser
	Rs("BBSPostTime")		= BBSPostTime
	Rs("BBSLastPostTime")	= BBSLastPostTime
	Rs("BBSLastPostUser")	= BBSLastPostUser
	Rs("BBSRoot")			= BBSRoot
	Rs.Update
	Rs.close
	
	Sql = "update [MyBBS_Board] set [BoardReply] = [BoardReply] + 1"
	Sql = Sql & ", [BoardLastPostTime] 		='" & BBSLastPostTime & "'"
	Sql = Sql & ", [BoardLastPostUser] 		='" & BBSLastPostUser & "'"
	Sql = Sql & ", [BoardLastPostTopic] 	='" & BBSTitle & "'"
	Sql = Sql & " where [ID]=" & BBSBoardID
	Conn.execute(Sql)

	Sql = "update [MyBBS_BBS] set [BBSReply] = [BBSReply] + 1"
	Sql = Sql & ", [BBSLastPostTime] 		='" & BBSLastPostTime & "'"
	Sql = Sql & ", [BBSLastPostUser] 		='" & BBSLastPostUser & "'"
	Sql = Sql & " where [ID]=" & BBSRoot
	Conn.execute(Sql)

    FormPage	= Request.ServerVariables("HTTP_REFERER")
	
    Response.write "<script language='javascript'>alert('�ظ��ɹ���');location.href='" & FormPage & "';</script>"
  end if   
%>