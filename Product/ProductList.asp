<table width="100%" border="0" cellpadding=1 cellspacing=1>
<%
'  Sql="Select Distinct [ProductClass_BigName],[ProductClass_BigOrder],[ProductClass_MidName],[ProductClass_MidOrder] From [7Gem_ProductClass] Order By [ProductClass_BigOrder],[ProductClass_MidOrder]"
  Sql="Select [ProductClass_BigName],[ProductClass_BigOrder],[ProductClass_MidName],[ProductClass_MidOrder] From [7Gem_ProductClass] Order By [ProductClass_BigOrder],[ProductClass_MidOrder]"
  Set Rs=Server.CreateObject("ADODB.RecordSet")
  Rs.Open Sql,conn,1,3
  Do While Not (Rs.Bof Or Rs.eof)	
	ProductClass_BigName	= rs(0)
	ProductClass_MidName	= rs(2)
	if (ProductClass_BigName<>ProductClass_BigNameTmp) then
%>
  <tr>
    <td>&nbsp;--&gt;&nbsp;<a href='ProductListView.asp?ProductClass_BigName=<%=ProductClass_BigName%>'><%=ProductClass_BigName%></a></td>
  </tr>
<%
	end if
%>
  <tr>
    <td>&nbsp;&nbsp;<a href='ProductListView.asp?ProductClass_MidName=<%=ProductClass_MidName%>'><font color=#0066ff><img src="/images/arrow_01.gif" width="3" height="5" border="0" align="absmiddle"><%=ProductClass_MidName%></font></a></td>
  </tr>
<%
   ProductClass_BigNameTmp = rs(0)
   rs.movenext
  Loop
  Rs.Close
%>
</table>