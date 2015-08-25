<!-- #include file="../Inc/Top.asp" -->
<!--#include file="../BBS/sk_conn.asp"-->
<!--#include file="../inc/fun.asp"-->
<link rel="stylesheet" type="text/css" href="../BBS/inc/style<%=GBL_Board_BoardStyle%>.css">
<link href="../Css/7Gem.css" rel="stylesheet" type="text/css">

<style>
<!--
	.l {FONT-FAMILY: 宋体; FONT-SIZE: 12px;}
	table {WORD-BREAK: break-all;BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 1px; }TD {BORDER-RIGHT: 0px; BORDER-TOP: 0px;}
-->
</style>
<table width="980" height="842" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="9" height="373" valign="top" bgcolor="f3f6f7"><img src="../images/2_left_line.gif" width="9" height="698"></td>
    <td width="212" valign="top" bgcolor="f3f6f7"> <table width="100" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td><!-- #include file="../Inc/Login.asp" --></td>
        </tr>
      </table>
			<!--  #include file="../Users/Menu.asp" -->
      <table width="100" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../images/left_menuline.gif" width="212" height="18"></td>
        </tr>
      </table>
			<!-- #include file="../Employ/Menu.asp" -->
      <table width="100" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../images/left_menuline.gif" width="212" height="18"></td>
        </tr>
      </table>
			<!-- #include file="../Download/Menu.asp" -->
      <table width="100" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../images/left_menuline.gif" width="212" height="18"></td>
        </tr>
      </table>
			<!-- #include file="../Link/Menu.asp" -->
      </td>
    <td width="29" valign="top" background="../images/2_midd_bg0.gif"><img src="../images/2_midd_line.gif" width="29" height="698"></td>
    <td width="132" valign="top" background="../images/2_midd_bg.gif"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="132" height="698">
      <param name="movie" value="../images/page_5.swf">
      <param name="quality" value="high">
      <embed src="../images/page_5.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="132" height="698"></embed>
    </object></td>
    <td width="615" valign="top">
		<table width="100" border=0 cellspacing=0 cellpadding=0 align=center>
      <tr>
        <td>&nbsp;</td>
      </tr>
    </table>		
		<table width="95%" border=0 cellspacing=0 cellpadding=0 align=center>
      <tr>
        <td><%=Replace(GBL_TableHeadString,"770","100%")%></td>
      </tr>
    </table>
		<div align=center><center>
     <table border='0' cellpadding='0' cellspacing='1' class=TBone width=95%>
 			<%
				  ThisPage					= Request.ServerVariables("PATH_INFO")
					Page							= Cint(Request("Page"))
					Dim ProductCount
				  Set Rs	= Server.CreateObject("ADODB.Recordset")
				  'Sql = "Select Top 2 * From [7Gem_Product] where [Product_Online] = True And Product_Remark = 1 order by [Product_AddDate] desc"
				  Sql = "Select * From [7Gem_Product] where [Product_Online] = True order by [Product_AddDate] desc"
					Rs.Open Sql, Conn ,1, 3
					if not (rs.bof or rs.eof) then
						rs.PageSize=7
						if page=0 then page=1
						pages=rs.pagecount
						if page > pages then page=pages
						rs.AbsolutePage=page  
						for j=1 to rs.PageSize 
							Product_ImgPrevWidth			= Rs("Product_ImgPrevWidth")
							Product_ImgPrevHeight			= Rs("Product_ImgPrevHeight")
							if Product_ImgPrevWidth		> 150 then
								Product_ImgPrevWidth		= 150
								Product_ImgPrevHeight		= (150 * Rs("Product_ImgPrevHeight")) / Rs("Product_ImgPrevWidth")
							end if
							if Product_ImgPrevHeight	> 100 then
								Product_ImgPrevHeight	= 100
								Product_ImgPrevWidth		= (100 * Rs("Product_ImgPrevWidth")) / Rs("Product_ImgPrevHeight")
							end if
							
			%> 
        <tr class=TBBG9>
	         <td width="100%" bgcolor="#F5F5F5">
						<table width="509" height="100" border=0 cellpadding="0" cellspacing="0" bgcolor="#F5F5F5" class="new">
							<tr>
								<td width="12" rowspan="4" valign="middle" bgcolor="#FFFFFF">&nbsp;&nbsp;&nbsp;</td>
								<td width="162" rowspan="4" valign="middle" bgcolor="#FFFFFF"><a target="_blank" href="../Product/ProductView.asp?Product_ID=<%=Rs("Product_ID")%>"><img src="../Product/Pic/<%=Rs("Product_ImgPrev")%>" width="<%=Product_ImgPrevWidth%>" height="<%=Product_ImgPrevheight%>" border=0></a></td>
								<td width="117" height="22">名称：<a href='../Product/ProductView.asp?Product_ID=<%=Rs("Product_ID")%>'><%=Rs("Product_Name")%></a></td>
							  <td width="200">产地：<%=Rs("Product_Area")%></td>
							</tr>
							<tr><td height="22">型号：<%=Rs("Product_Model")%></td>
							  <td>库存：<%=Rs("Product_Quantity")%></td>
							</tr>
							<tr><td height="22">价格：
							  <%
								if Rs("Product_Price")=0 then
									response.write "<b><font color=red>备货中...</font></b>"
								else
									response.write "<b><font color=red>" & Rs("Product_Price") & " 元</font></b>"
								end if
							%>
								</td>
							  <td>（此价格包含邮费）</td>
							</tr>
							<tr><td height="16" colspan="2" align="right"><a href="#" onClick="window.open('../Trade/add.asp?Product_ID=<%=Rs("Product_ID")%>','add','scrollbars=yes,resizable=no,width=650,height=450')">加入购物车</a></td></tr>
						</table>
					</td>
	       </tr>
 					<%
      				rs.movenext
	  					if rs.eof then exit for
						next
					end if					
				%> 
      </table>
		</center></div>
      <table width=95% border=0 cellspacing=0 cellpadding=0 align=center>
        <tr>
          <td><%=Replace(GBL_TableBottomString,"770","100%")%></td>
        </tr>
      </table>
			
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
  <form method=Post action="<%=ThisPage%>">  
    <td height="30" align="center"> 
    <%
	  ThisPage = ThisPage & "?1=1"
	  if Page<2 then
	    response.write "首页 上一页&nbsp;"
	  else
	    response.write "<a href=" & ThisPage & "&page=1>首页</a>&nbsp;"
		response.write "<a href=" & ThisPage & "&page=" & Page-1 & ">上一页</a>&nbsp;"
	  end if
	  if rs.pagecount-page<1 then
	    response.write "下一页 尾页"
	  else
	    response.write "<a href=" & ThisPage & "&page=" & (page+1) & ">下一页</a> "
		response.write "<a href=" & ThisPage & "&page="&rs.pagecount&">尾页</a>"
	  end if
	  response.write "&nbsp;页次：<strong><font color=red>"&Page&"</font>/"&rs.pagecount&"</strong>页 "
	  response.write "&nbsp;共<b><font color='#FF0000'>"&rs.recordcount&"</font></b>条记录 <b>"&rs.pagesize&"</b>条记录/页"
	  response.write " 转到：<input type='text' name='page' size=4 maxlength=10 class=input value="&page&">"
	  response.write " <input class=input type='submit'  value=' Goto '  name='cndok'></span></p>"
	%>
    </td>
	</form>
  </tr>
</table>
			
			</td>
  </tr>
</table>
<!-- #include file="../Inc/Down.asp" --><% 
  rs.close
  set rs=nothing
%>