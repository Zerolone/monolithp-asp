<!-- #include virtual =	"/Inc/Monolith_Conn.asp" -->
<%
	Response.write 1912 \ 10
	Response.write "<hr>"

	Sql = "Select * From [Monolith_News] Where Left(News_Title,1) ='一'"
	Rs.Open Sql, Conn, 1, 1 
	Response.write Rs(0)
	Response.write Rs(1)
	Response.write Rs(2)
	Response.write Rs(3)
	Response.write Rs(4)
	
	Response.write "<hr>"

	Dim NewsArea, j ,i

	NewsArea = Request("NewsArea")
	
	Response.write NewsArea
	Response.write "<hr>"
	NewsArea = Split(NewsArea, ",")
	
	j =0
	
	For i=LBound(NewsArea) to UBound(NewsArea)
		j = j + NewsArea(i)
	Response.write "<hr>"
		Response.write NewsArea(i)
	Next
	
	Response.write "<hr>"
	Response.write j
	
%>

<form action="Test.asp" method="post">
			显示位置:<input type="checkbox" name="NewsArea" value="1" id="NewsArea_01"><label for="NewsArea_01">首页</label><input type="checkbox" name="NewsArea" value="10" id="NewsArea_02"><label for="NewsArea_02">图片新闻</label><input type="checkbox" name="NewsArea" value="100" id="NewsArea_03"><label for="NewsArea_03">栏目首页</label><input type="submit" name="aaa">
</form>


