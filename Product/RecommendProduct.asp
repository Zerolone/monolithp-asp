<%
  Sql = "Select Top 2 * From [7Gem_Product] where [Product_Online] = True And Product_Remark = 1 order by [Product_AddDate] desc"
  Rs.open Sql , conn, 1, 3
  n=0
  if Rs.bof and Rs.eof then
    response.write "您好！目前商城暂时没有 "& ProductClass_BigName & ProductClass_MidName &" 类商品"
  end if
  Do While Not (Rs.eof Or Rs.bof)
	Product_ImgPrevWidth			= Rs("Product_ImgPrevWidth")
	Product_ImgPrevHeight			= Rs("Product_ImgPrevHeight")
	if Product_ImgPrevWidth		> 103 then Product_ImgPrevWidth		= 103
	if Product_ImgPrevHeight	> 103 then Product_ImgPrevHeight	= 103
%> 
<table width="365" border=0 class="new" bgcolor="#F5F5F5">
  <tr>
    <td rowspan="4" align="center" valign="middle"><a target="_blank" href="Product/ProductView.asp?Product_ID=<%=Rs("Product_ID")%>"><img src="Product/Pic/<%=Rs("Product_ImgPrev")%>" width="<%=Product_ImgPrevWidth%>" height="<%=Product_ImgPrevheight%>" border=0></a></td>
		<td>名称：<a href='../ProductView.asp?Product_ID=<%=Rs("Product_ID")%>'><%=Rs("Product_Name")%></a></td>
  </tr>
  <tr><td>型号：<%=Rs("Product_Model")%></td></tr>
  <tr>
    <td>价格：
		<%
			if Rs("Product_Price")=0 then
				response.write "<b><font color=red>备货中...</font></b>"
			else
			  response.write "<b><font color=red>" & Rs("Product_Price") & " 元</font></b>"
			end if
		%>
		</td>
	</tr>
  <tr><td align="right"><a href="#" onClick="window.open('Trade/add.asp?Product_ID=<%=Rs("Product_ID")%>','add','scrollbars=yes,resizable=no,width=650,height=450')">加入购物车</a></td>
  </tr>
</table>
<%
    Rs.movenext
  loop
	Rs.close
%> 