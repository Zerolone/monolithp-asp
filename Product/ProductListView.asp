<!--#Include File ="../Inc/Conn.asp"-->
<meta http-equiv="Content-Type" content="text/html; chaRset=gb2312">
<body bgcolor="#E7E7E5">
<%
  Page=Request("page")
  pagecount=page
  ProductClass_BigName	= Request("ProductClass_BigName")
  ProductClass_MidName	= Request("ProductClass_MidName")
  ThisPage				= Request.ServerVariables("PATH_INFO")

  Sql="Select * from [7Gem_Product] where [Product_Online]=true "
  if ProductClass_BigName<>"" then Sql = Sql + "and [Product_BigName]='"&ProductClass_BigName&"'"
  if ProductClass_MidName<>"" then Sql = Sql + "and [Product_MidName]='"&ProductClass_MidName&"'"
  Sql = Sql + "order by [Product_BigName],[Product_MidName],[Product_AddDate] Asc"
  Rs.open Sql , conn, 1, 3
  n=0
  if Rs.bof and Rs.eof then
    response.write "���ã�Ŀǰ�̳���ʱû�� "& ProductClass_BigName & ProductClass_MidName &" ����Ʒ"
  end if

  if ProductClass_BigName <> "" then response.write ProductClass_BigName & " >> "
  if ProductClass_MidName <> "" then response.write ProductClass_MidName 
  if ProductClass_BigName ="" and ProductClass_MidName = "" then response.write "���в�Ʒ�е�:<b>" &Keywords&"</b>"
  if not isempty(Request("page")) then   
	pagecount=cint(Request("page"))
  else   
    pagecount=1
  end if
  Rs.pagesize=8
  Rs.AbsolutePage=pagecount
  Do While Not (Rs.eof Or Rs.bof)
%> 
<table cellspacing=1 cellpadding=2 width="540" border=0 class="new" bgcolor="#F5F5F5">
  <tr bgcolor="#FFFFFF">
    <td rowspan="4" align="center" valign="middle" width="161"><a href="ProductView.asp?Product_ID=<%=Rs("Product_ID")%>"><img src="Pic/<%=Rs("Product_ImgPrev")%>" width='<%=Rs("Product_ImgPrevWidth")%>' height='<%=Rs("Product_ImgPrevheight")%>' border=0></a></td>
	<td width="239">���ƣ�<a href='ProductView.asp?Product_ID=<%=Rs("Product_ID")%>'><%=Rs("Product_Name")%></a>&nbsp;</td><td>�ͺţ�<%=Rs("Product_Model")%></td>
  </tr>
  <tr>
    <td class=bodycontent colspan="2" bgcolor="#FFFFFF">�۸�
	<%
	  if Rs("Product_Price")=0 then
		response.write "<b><font color=red>������...</font></b>"
	  else
		response.write "<b><font color=red>" & Rs("Product_Price") & " Ԫ</font></b>"
	  end if
	  %>
	</td>
  </tr>
  <tr><td height=5 colspan="2" bgcolor="#FFFFFF">��飺<%=Rs("Product_Content")%></td></tr>
  <tr align="right">
  <td height="21" colspan="2" bgcolor="#FFFFFF"><a href="#" onClick="window.open('../Trade/add.asp?Product_ID=<%=Rs("Product_ID")%>','add','scrollbars=yes,resizable=no,width=650,height=450')">���빺�ﳵ</a></td>
  </tr>
</table>
<%
    n=n+1
    Rs.movenext
    if n>=Rs.pagesize then exit do
  loop
%>
<table width=99% border="0" cellspacing="1" cellpadding="2" align="center" bgcolor=#cfcfcf>
<form action="<%=ThisPage%>" method="post" >
  <tr>
    <td height="22">
      <div align="center">ҳ��: <b><font color=red><%=pagecount%></font>/<%=Rs.pagecount%></b> 
      <%
	    ThisPage = ThisPage & "?ProductClass_BigName=" & ProductClass_BigName & "&ProductClass_MidName=" & ProductClass_MidName & "&page="
	    if pagecount=1 and Rs.pagecount<>pagecount and Rs.pagecount<>0 then Response.Write "<a href=" & ThisPage & cstr(pagecount+1) & ">��һҳ</a>"
        if Rs.pagecount>1 and Rs.pagecount=pagecount then Response.write "<a href=" & ThisPage & cstr(pagecount-1) & ">��һҳ</a>"
        if pagecount<>1 and Rs.pagecount<>pagecount then
		  Response.write "<a href=" & ThisPage & cstr(pagecount-1) & ">��һҳ</a>"
		  Response.Write "<a href=" & ThisPage & cstr(pagecount+1) & ">��һҳ</a>"
        end if
	  %>
      &nbsp; ת���� 
        <select name="page">
          <%for i=1 to Rs.pagecount%> 
            <option value="<%=i%>"><%=i%></option>
          <%next%> 
        </select>ҳ
		<input type="submit" name="go" value="Go">
		<input type="hidden" name="ProductClass_BigName" value="<%=ProductClass_BigName%>">
		<input type="hidden" name="ProductClass_MidName" value="<%=ProductClass_MidName%>">
	  </div>
    </td>
  </tr>
</form>
</table>
<%
Rs.close
%> 