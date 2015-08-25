<!-- #include virtual ="/Inc/Monolith_Conn.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<br><br>
<%
	Dim ThisPage
	ThisPage       = Request.ServerVariables("PATH_INFO")

	if Request("Code") = "Run" then
		Sql = Request("SqlStr")
		Conn.execute Sql
		if err.number<>0 then
			response.write "语句执行出错："& err.description & err.clear
		else
			response.write "语句执行成功！"
		end if
		Response.write "<hr>"
		Call CloseConn
	end if
%><form method="POST" action="<%=ThisPage%>?Code=Run">
	<p>&nbsp;&nbsp;Sql语句：<input type="text" name="SqlStr" size="90" class="InputBox" value="<%=Sql%>"> <input type="submit" value=" 提 交 " name="B1" class="InputBox"> <input type="reset" value=" 重 置 " name="B2" class="InputBox"></p>
</form>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->
<p>&nbsp;&nbsp;更新栏目用的SQL：Update [Monolith_News] Set [News_Class]='' Where [News_Class]=''</p>