<!--#include file="../Inc/conn.asp"-->
<!--#include file="../inc/fun.asp"-->
<!--#include file="Admin_Check.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<%
  Dim FileOther_ID
  Product_ID	= trim(Request("Product_ID"))
  FileOther_ID	= trim(Request("FileOther_ID"))
  if Product_ID <> "" and FileOther_ID <> "" then
	  sql="delete from [7Gem_Product_FileOther] where [FileOther_ID] in (" & FileOther_ID & ")"
	  conn.Execute sql
  end if
  response.redirect "Admin_Product_Modify.asp?Product_ID=" & Product_ID
%>