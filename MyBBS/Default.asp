<!-- #include virtual="/MyBBS/Head.asp"-->
<!-- #include file="LoginString.asp"-->
<!-- #include file="BoardString.asp"-->
<!-- #include file="WelcomeBack.asp"-->
<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线-->
<%
	Dim i, TheString
	i = 0

  '----------------0-------------1--------------2---------------3-------------4--------------5---------------6-----------------7--------------------8---------------------9---------------------10------------------------11
  Sql= "select [Board_ID], [Board_Name], [Board_Level], [Board_Intro], [Board_Master], [Board_Topic], [Board_Reply], [Board_LastPostTime], [Board_LastPostUser], [Board_LastPostUserID], [Board_LastPostTopic], [Board_LastPostTopicID] From [MyBBS_Board] Order By [Board_Level] Asc"
  Rs.open Sql, Conn, 1, 1
	Do While Not (Rs.Bof Or Rs.Eof)
	  if Len(Rs(2)) = 2 then
	  	'根板块
	  	TheString	= MyBBS_Template_01_Head
			TheString	= Replace(TheString, "{Board_ID}", 		Rs(0))
			TheString = Replace(TheString, "{Board_Name}", 	Rs(1))
			if i > 0 then Response.write MyBBS_Template_01_Foot & "<br>"
			Response.write TheString
		else
			'分模版
			TheString	= MyBBS_Template_01_Body
			TheString	= Replace(TheString, "{Board_ID}",							Rs(0))
			TheString	= Replace(TheString, "{Board_Name}",						Rs(1))
			TheString	= Replace(TheString, "{Board_Intro}",						Rs(3))
			TheString	= Replace(TheString, "{Board_Master}",					Rs(4))
			TheString	= Replace(TheString, "{Board_Topic}", 					Rs(5))
			TheString	= Replace(TheString, "{Board_Reply}", 					Rs(6))
			TheString	= Replace(TheString, "{Board_LastPostTime}", 		Rs(7))
			TheString	= Replace(TheString, "{Board_LastPostUser}", 		Rs(8))
			TheString	= Replace(TheString, "{Board_LastPostUserID}",	Rs(9))
			TheString	= Replace(TheString, "{Board_LastPostTopic}", 	Rs(10))
			TheString	= Replace(TheString, "{Board_LastPostTopicID}", Rs(11))
			Response.write TheString
		end if

		i = i + 1
		Rs.MoveNext
	Loop
	Response.write MyBBS_Template_01_Foot & "<br>"

	Rs.Close
  'Sql执行次数+1
  SqlQueryNum	= SqlQueryNum + 1


	Dim Users_List, Users_Number
	Users_Number = 0
	Sql = "Select [Users_ID], [Users_Name] From [Monolith_Users] Where [Users_LastLogin] < #" & DateAdd("n", -30, Now) & "#"
	Rs.Open Sql, Conn, 1, 1
	Do while Not(Rs.Bof or Rs.Eof)
		Users_List		= Users_List & "<a href=""ViewUser.asp?Users_ID=" & Rs(0) & """>" & Rs(1) & "</a>, "
		Users_Number	= Users_Number + 1
		Rs.MoveNext
	Loop
	Rs.Close
	SqlQueryNum	= SqlQueryNum + 1

	TheString	= MyBBS_Template_01_Stats
	TheString	= Replace(TheString, "{Users_List}",							Users_List)
	TheString	= Replace(TheString, "{Users_Number}",						Users_Number)
	
  '------------------------0----------------1-------------------2----------------3------------------4-----------------5
  Sql= "Select Top 1 [Stats_Subject], [Stats_UserNumber], [Stats_NewUser], [Stats_NewUserID], [Stats_MostTime], [Stats_MostUsers] From [Monolith_Stats]"
  Rs.open Sql, Conn, 1, 1
  if Not(Rs.Bof Or Rs.Eof) then	
		TheString	= Replace(TheString, "{Stats_Subject}",					Rs(0))
		TheString	= Replace(TheString, "{Stats_UserNumber}",			Rs(1))
		TheString	= Replace(TheString, "{Stats_NewUser}",					Rs(2))
		TheString	= Replace(TheString, "{Stats_NewUserID}",				Rs(3))
		TheString	= Replace(TheString, "{Stats_MostTime}",				Rs(4))
		TheString	= Replace(TheString, "{Stats_MostUsers}",				Rs(5))
	end if
	Response.write TheString

  SqlQueryNum	= SqlQueryNum + 1
  Call CloseRs
  Call CloseConn

  TheString	= MyBBS_Template_Foot 
  TheString	= Replace(TheString, "{Now}",	Now())
  TheString	= Replace(TheString, "{ExecuteTime}",	FormatNumber((Timer()-Startime)*1000,5) & "毫秒。")
  TheString	= Replace(TheString, "{SqlQueryNum}",	SqlQueryNum)
  Response.write TheString
%>