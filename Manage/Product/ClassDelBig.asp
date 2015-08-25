<!-- #include virtual ="/Inc/Monolith_Conn.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim ProductClass_BigName

  ProductClass_BigName=trim(Request("ProductClass_BigName"))
  if ProductClass_BigName <> "" then
	sql="delete from [Monolith_ProductClass]	where [ProductClass_BigName]='" & ProductClass_BigName & "'"
	conn.Execute sql
	sql="delete from [Monolith_Product]			where [Product_BigName]='" & ProductClass_BigName & "'"
	conn.Execute sql
  end if
	Response.Redirect "ClassManage.asp"
%>