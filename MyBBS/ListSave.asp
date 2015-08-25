<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual = "/Inc/Monolith_Function.asp" -->
<!-- #include virtual = "/Inc/LoginUserCheck.asp" -->
<%
  if IsLogin = 0 then
    Response.Write "请登陆！"
		Response.end
  end if

  Dim BBS_ID, BBS_Subject, BBS_Content, BBS_Board_ID, BBS_PostUser, BBS_PostTime, BBS_LastPostUser, BBS_LastPostTime, BBS_Root, BBS_Icon, BBS_SubjectDesc

  Dim ThisPage, FormPage
  
  FormPage	= Request("FormPage")
  ThisPage	= Request.ServerVariables("PATH_INFO")
  if Request("Submit") = " 发 表 新 的 主 题 " then 
    
		BBS_Subject				= Trim(Request("BBS_Subject"))
		BBS_Content				= Trim(Request("BBS_Content"))
		BBS_Content				= Replace(BBS_Content, chr(10) & chr(13) , "<br>")
		BBS_Content				= Replace(BBS_Content, vbcrlf , "<br>")
	  BBS_Board_ID			= Trim(Request("BBS_Board_ID"))
		BBS_PostUser			= Users_Name
		BBS_PostTime			= Now
		BBS_LastPostUser	= Users_Name
		BBS_LastPostTime	= Now
		BBS_Icon					= Request("BBS_Icon")
		BBS_SubjectDesc		= Request("BBS_SubjectDesc")
 
	 	'-------------------0-------------1----------------2----------------3-----------4-------------5--------------6-------------7-----------8-------------9-------------------10--------------11--------------12----------------13------------------14---------------15
		Sql = "Select [BBS_Subject], [BBS_Content], [BBS_Board_ID], [BBS_PostUser], [BBS_IsTop], [BBS_PostTime], [BBS_Reply], [BBS_View], [BBS_Root], [BBS_LastPostUser], [BBS_LastPostTime], [BBS_Icon], [BBS_SubjectDesc], [BBS_PostUserID], [BBS_LastPostUserID], [BBS_ID] From [MyBBS_BBS]"
		Rs.Open Sql, Conn, 1, 3
		Rs.AddNew
		Rs(0)		= BBS_Subject
		Rs(1)		= BBS_Content
		Rs(2)		= BBS_Board_ID
		Rs(3)		= BBS_PostUser
		Rs(4)		= 0
		Rs(5)		= BBS_PostTime
		Rs(6)		= 0
		Rs(7)		= 0
		Rs(8)		= 0
		Rs(9)		= BBS_LastPostUser
		Rs(10)	= BBS_LastPostTime
		Rs(11)	= BBS_Icon
		Rs(12)	= BBS_SubjectDesc
		Rs(13)	= Users_ID
		Rs(14)	= Users_ID
		
		Rs.Update
		BBS_ID	= Rs(15)
	
		Sql = "update [MyBBS_Board] set [Board_Topic] = [Board_Topic] + 1"
		Sql = Sql & ", [Board_LastPostTime] 		='" & BBS_LastPostTime & "'"
		Sql = Sql & ", [Board_LastPostUser] 		='" & BBS_LastPostUser & "'"
		Sql = Sql & ", [Board_LastPostUserID]		=" 	& Users_ID
		Sql = Sql & ", [Board_LastPostTopic] 		='" & BBS_Subject & "'"
		Sql = Sql & ", [Board_LastPostTopicID]	=" 	& BBS_ID
		Sql = Sql & " Where [Board_ID]=" & BBS_Board_ID
		Conn.execute(Sql)

		Sql = "update [Monolith_Users] set [Users_Subject] = [Users_Subject] + 1"
		Sql = Sql & " Where [Users_ID]=" & Users_ID
		Conn.execute(Sql)
		
    Response.write "<script language='javascript'>alert('主题发表成功！');location.href='" & FormPage & "';</script>"
    
    Call CloseRs
    Call CloseConn
  end if
  
  if Request("Submit") = " 加 入 回 复 " then
		BBS_Content				= Trim(Request("BBS_Content"))
		BBS_Content				= Replace(BBS_Content, chr(10) & chr(13) , "<br>")
		BBS_Content				= Replace(BBS_Content, vbcrlf , "<br>")
	  BBS_Board_ID			= Trim(Request("Board_ID"))
		BBS_PostUser			= Users_Name
		BBS_PostTime			= Now
		BBS_LastPostUser	= Users_Name
		BBS_LastPostTime	= Now
		BBS_Root					= Trim(Request("BBS_Root"))
		BBS_Icon					= Request("BBS_Icon")
 
 		'-------------------0-------------1----------------2----------------3-----------4-------------5--------------6-------------7-----------8-------------9-------------------10--------------11--------------12----------------13------------------14
		Sql = "Select [BBS_Subject], [BBS_Content], [BBS_Board_ID], [BBS_PostUser], [BBS_IsTop], [BBS_PostTime], [BBS_Reply], [BBS_View], [BBS_Root], [BBS_LastPostUser], [BBS_LastPostTime], [BBS_Icon], [BBS_SubjectDesc], [BBS_PostUserID], [BBS_LastPostUserID] From [MyBBS_BBS]"
		Rs.open Sql, Conn, 1, 3
		Rs.AddNew
		Rs(0)		= BBS_Subject
		Rs(1)		= BBS_Content
		Rs(2)		= BBS_Board_ID
		Rs(3)		= BBS_PostUser
		Rs(4)		= 0
		Rs(5)		= BBS_PostTime
		Rs(6)		= 0
		Rs(7)		= 0
		Rs(8)		= BBS_Root
		Rs(9)		= BBS_LastPostUser
		Rs(10)	= BBS_LastPostTime
		Rs(11)	= BBS_Icon
		Rs(12)	= BBS_SubjectDesc
		Rs(13)	= Users_ID
		Rs(14)	= Users_ID
		Rs.Update
		Rs.close
	
		Sql = "update [MyBBS_Board] set [Board_Reply] = [Board_Reply] + 1"
		Sql = Sql & ", [Board_LastPostTime] 		='" & BBS_LastPostTime & "'"
		Sql = Sql & ", [Board_LastPostUser] 		='" & BBS_LastPostUser & "'"
'		Sql = Sql & ", [Board_LastPostTopic] 		='" & BBS_Subject & "'"
'		Sql = Sql & ", [Board_LastPostTopicID]	=" 	& BBS_Root
		Sql = Sql & " where [Board_ID]=" & BBS_Board_ID
		Conn.execute(Sql)

		Sql = "update [MyBBS_BBS] set [BBS_Reply] = [BBS_Reply] + 1"
		Sql = Sql & ", [BBS_LastPostTime] 		='" & BBS_LastPostTime & "'"
		Sql = Sql & ", [BBS_LastPostUser] 		='" & BBS_LastPostUser & "'"
		Sql = Sql & " where [BBS_ID]=" & BBS_Root
		Conn.execute(Sql)

		Sql = "update [Monolith_Users] set [Users_Subject] = [Users_Subject] + 1"
		Sql = Sql & " Where [Users_ID]=" & Users_ID
		Conn.execute(Sql)
		
	  Response.write "<script language='javascript'>alert('回复成功！');location.href='ListView.asp?BBS_ID=" & BBS_Root & "&Board_ID=" & BBS_Board_ID & "';</script>"
  end if   
%>