<!-- #include virtual ="/Inc/Monolith_Conn.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim ProductClass_MidName, ProductClass_BigName


  ProductClass_MidName=trim(Request("ProductClass_MidName"))
  ProductClass_BigName=trim(Request("ProductClass_BigName"))

  if ProductClass_MidName <> "" then
	sql="delete from [Monolith_ProductClass]	where [ProductClass_MidName]='" & ProductClass_MidName & "' And [ProductClass_BigName]='" & ProductClass_BigName &  "'"
	conn.Execute sql
	sql="delete from [Monolith_Product]			where [Product_MidName]='" & ProductClass_MidName & "' And [Product_BigName]='" & ProductClass_BigName &  "'"
	conn.Execute sql
  end if

	Response.Redirect "ClassManage.asp"
%>