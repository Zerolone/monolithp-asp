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
			response.write "���ִ�г���"& err.description & err.clear
		else
			response.write "���ִ�гɹ���"
		end if
		Response.write "<hr>"
		Call CloseConn
	end if
%><form method="POST" action="<%=ThisPage%>?Code=Run">
	<p>&nbsp;&nbsp;Sql��䣺<input type="text" name="SqlStr" size="90" class="InputBox" value="<%=Sql%>"> <input type="submit" value=" �� �� " name="B1" class="InputBox"> <input type="reset" value=" �� �� " name="B2" class="InputBox"></p>
</form>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->
<p>&nbsp;&nbsp;������Ŀ�õ�SQL��Update [Monolith_News] Set [News_Class]='' Where [News_Class]=''</p>