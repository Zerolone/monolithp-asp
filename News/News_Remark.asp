<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<%
	Dim News_ID, News_Hits, News_Remark, TheString
	News_ID					= Request("News_ID")

	if CheckNumII(News_ID) then
		'���ŵ����
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
			TheString			= TheString & "�����ˣ�" & Rs(0) & "<br>�������ݣ�<br>" & Rs(1) & "<hr>"
			Rs.MoveNext
		Loop

		Call CloseRs
		Call CloseConn

	end if
%>
document.write("<hr>");
document.write("���������");
document.write("<%=News_Hits+1%>");
document.write("  �ظ�������")
document.write("<%=News_Remark%>");
document.write("<hr>");
document.write("���5�����ۣ�<hr><%=TheString%>");