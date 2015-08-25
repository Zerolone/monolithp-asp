<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<%
	Dim News_ID, News_FileName, Remark_UserName, Remark_Content
	News_ID					= Request("News_ID")
	News_FileName		= Request("News_FileName")
	Remark_UserName	= Monolith_Encode(Request("Remark_UserName"))
	Remark_Content	= Monolith_Encode(Request("Remark_Content"))
	
	CheckNum(News_ID)
	
	Sql	= "Select [Remark_UserName], [Remark_Content], [News_ID] From [Monolith_News_Remark]"
	Rs.Open Sql, Conn, 1, 3
	Rs.AddNew
	Rs(0)	= Remark_UserName
	Rs(1)	= Remark_Content
	Rs(2)	= News_ID
	Rs.Update
	Rs.Close
	
	Sql	= "Update [Monolith_News] Set [News_Remark]= [News_Remark]+1 Where [News_ID]=" & News_ID
	Conn.Execute(Sql)
	
	Call CloseConn
	Response.Redirect News_FileName
%>