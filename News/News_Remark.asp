<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<%
	Dim News_ID, News_Hits, News_Remark, TheString
	News_ID					= Request("News_ID")

	if CheckNumII(News_ID) then
		'新闻点击数
		'----------------0------------1
		Sql="Select [News_Hits], [News_Remark] From [Monolith_News] Where [News_ID]= " & News_ID
		Rs.Open Sql, Conn, 1, 3
		News_Hits		= Rs(0)
		News_Remark	= Rs(1)
		Rs(0) = Rs(0) + 1
		Rs.update
		Rs.Close

		Sql	= "Select  Top 5 [Remark_UserName], [Remark_Content], [News_ID] From [Monolith_News_Remark] Where [News_ID]=" & News_ID & " Order By [Remark_ID] Desc"
		Rs.Open Sql, Conn, 1, 1
		Do While Not(Rs.Bof Or Rs.Eof)
			TheString			= TheString & "评论人：" & Rs(0) & "<br>评论内容：<br>" & Rs(1) & "<hr>"
			Rs.MoveNext
		Loop

		Call CloseRs
		Call CloseConn

	end if
%>
document.write("<hr>");
document.write("点击次数：");
document.write("<%=News_Hits+1%>");
document.write("  回复次数：")
document.write("<%=News_Remark%>");
document.write("<hr>");
document.write("最近5条评论：<hr><%=TheString%>");