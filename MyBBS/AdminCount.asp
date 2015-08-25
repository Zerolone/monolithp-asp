<!-- #include file="Inc/Conn.asp"-->
<!-- #include file="Inc/Config.asp"-->
<!-- #include file="Inc/Function.asp"-->
<%
  dim FormPage, Code
  
  FormPage	= Request.ServerVariables("HTTP_REFERER")
  Code		= Request("Code")
  dim ID, BoardTopic, BoardReply
    
  ID = Request("ID")
  
  if Code = "Topic" then 
    Sql = "Select Count(*) as [BoardTopic] From [MyBBS_BBS] where [BBSBoardID]= " & ID
	Sql = Sql & " And [BBSRoot]=0"
    Rs.open sql, conn , 1, 3
    if not (rs.bof or rs.eof) then BoardTopic = Rs(0)
    Rs.Close
    Sql = "update [MyBBS_Board] set [BoardTopic] = " & BoardTopic
	Sql = Sql & " Where [ID]=" & ID
    conn.execute(sql)
  end if

  if Code = "Reply" then 
    Sql = "Select Count(*) as [BoardReply] From [MyBBS_BBS] where [BBSBoardID]= " & ID
	Sql = Sql & " And [BBSRoot]<>0"
    Rs.open sql, conn , 1, 3
    if not (rs.bof or rs.eof) then BoardReply = Rs(0)
    Rs.Close
    Sql = "update [MyBBS_Board] set [BoardReply] = " & BoardReply
	Sql = Sql & " Where [ID]=" & ID
    conn.execute(sql)
  end if

  Response.write "<script language='javascript'>alert('ÐÞ¸Ä³É¹¦£¡');location.href='" & FormPage & "';</script>"
' SqlServer Could Be Like Is	update [MyBBS_Board] set [BoardTopic] = (Select count(*) From [MyBBS_BBS] where [BBSBoardID]= 2) 
%>