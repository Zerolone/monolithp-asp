<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim Product_ID, action

  Product_ID= trim(Request("Product_ID"))
  action	= Request("Action")

  if Product_ID <> "" then
    if Action=" ɾ �� " then
		  sql="delete from [Monolith_Product]			where [Product_ID] in (" & Product_ID & ")"
			conn.Execute sql
	  end if
	
		if Action = " �� �� " then
		  Sql="update [Monolith_Product] set [Product_Online] = false  where [Product_ID] in (" & Product_ID & ")"
			conn.Execute sql
		end if

	  if Action = " �� �� " then
		  Sql="update [Monolith_Product] set [Product_Online] = true  where [Product_ID] in (" & Product_ID & ")"
			conn.Execute sql
    end if

		Response.write "<script language=""javascript"">"
		Response.write "  alert(""�����ɹ�����ˢ��ҳ�档"");"
'		Response.write "  window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);"
		Response.write "</script>"
	end if
%>