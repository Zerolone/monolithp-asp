<!-- #include file="Inc/Conn.asp"-->
<%
  dim BoardName, ID, BoardOrder
  BoardName			= Request("BoardName")
  ID				= Request("ID")
  BoardOrder		= Request("BoardOrder")
  
  dim Arr_BoardName, Arr_ID, Arr_BoardOrder
  Arr_BoardName		= Split(BoardName, ",")
  Arr_ID			= Split(ID, ",")
  Arr_BoardOrder	= Split(BoardOrder, ",")
  dim i
  for i = LBound(Arr_ID) to UBound(Arr_ID)
	Sql = "update [MyBBS_Board] set [BoardName]='" & Trim(Arr_BoardName(i)) & "', [BoardOrder] =" & Trim(Arr_BoardOrder(i)) & " where [ID]=" & Trim(Arr_ID(i))
	'Response.write sql & "<hr>"	
	Conn.execute(Sql)
  next
  Response.write "<script language='javascript'>alert('ÐÞ¸Ä³É¹¦£¡');location.href='AdminBoard.asp';</script>"
%>