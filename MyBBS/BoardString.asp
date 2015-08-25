<%
	Dim Board_ID, Board_Name
  Board_ID 		= Request("Board_ID")
  if Board_ID <> "" then
  	Sql = "Select Top 1 [Board_Name] From [MyBBS_Board] Where [Board_ID]=" & Board_ID
  	Rs.Open Sql, Conn, 1, 1
  	if Rs.Bof Or Rs.Eof then
  		Response.write "板块错误， 请联系管理员！"
  		Response.end
  	else
  		Board_Name = Rs(0)
  	end if
  	Rs.Close
  	
 	  'Sql执行次数+1
	  SqlQueryNum	= SqlQueryNum + 1
  end if
%>
<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线--><table cellspacing="0" cellpadding="0" width="98%" align="center">
	<tr>
		<td><b><a href="/MyBBS/Default.asp"><img src="/images/nav.gif" border="0"> MyBBS Board</a> 
		<%if Board_ID <> "" then%>
		&gt;&nbsp;<a href="List.asp?Board_ID=<%=Board_ID%>"><%=Board_Name%></a>
		<%end if%>
		</b>
		</td>
	</tr>
</table>