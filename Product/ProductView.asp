<!-- #include file="../Inc/Top.asp" -->
<!--#include file="../BBS/sk_conn.asp"-->
<!--#include file="../inc/fun.asp"-->
<!--#Include File ="../Inc/Config.asp"-->

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
<%

  ThisPage		= Request.ServerVariables("PATH_INFO")
  Product_ID	= Request("Product_ID")

  if Product_ID	= "" then Response.Redirect "../Product/Default.asp"
  Set Rs	= Server.CreateObject("ADODB.Recordset")

  Sql="Select * from [7Gem_Product] where [Product_ID]=" & Product_ID
  Rs.Open Sql, conn, 1, 3
%>
		<div align=center><center>
     <table border='0' cellpadding='0' cellspacing='1' class=TBone width=95%>
        <tr class=TBBG9>
	         <td height="31" colspan=2 >目前位置：<%=Rs("Product_BigName")%> >> <%=Rs("Product_MidName")%> </td>
	       </tr>
        <tr class=TBBG9>
         <td width="570" colspan=2 align="center" height="1" background="../images/bg_xg_01.gif"></td>
       </tr>
       <%
  Rs.Close

  Sql="Select * from [7Gem_Product] where [Product_Online]=true and [Product_ID]="& Product_ID 
  Rs.Open Sql,conn,1,3
  if Rs.bof and Rs.eof then
	Response.write "<Tr><Td width=100% rowspan='2' height=150 align=center>您好！目前商城的暂无此商品</Td></Tr>"
  else
    Do While Not Rs.eof
	  Product_Content 	= Rs("Product_Content")
	  Product_MemoSpec	= Rs("Product_MemoSpec")
		Product_Code			= Rs("Product_Code")

	  if Product_Content	<> "" then Product_Content	= Autolink(Product_Content)
	  if Product_MemoSpec	<> "" then Product_MemoSpec	= Autolink(Product_MemoSpec)
%>
        <tr class=TBBG9>
         <td width=<%=Rs("Product_ImgFullWidth")%> align=center valign=top><a href="Pic/<%=Rs("Product_ImgFull")%>" target="_blank"><img src="Pic/<%=Rs("Product_ImgPrev")%>" width='<%=Rs("Product_ImgPrevWidth")%>' height='<%=Rs("Product_ImgPrevheight")%>' border=0></a><br>
             <%
'			如果有其他图
'	  if Rs("Product_FileOther")="1" then
'	    Set RsFileOther=Server.CreateObject("ADODB.Recordset")
'	    Sql = "Select * from [7Gem_Product_FileOther] where [FileOther_Product_Code]='" & Product_Code &"'"
'	    RsFileOther.Open Sql , Conn, 1, 3
'	    if Not (RsFileOther.bof Or RsFileOther.eof) then
'	    RsFileOther.Close
'    end if
	%>
         </td>
         <td width=491 valign=top>
           <table  width="333" border=0 align="center" cellpadding=4 cellspacing=0 class=new>
             <tr>
               <td colspan=2 valign="bottom"><font color=#ff6633 size=4><b><%=Rs("Product_Name")%></b></font></td>
             </tr>
             <tr>
               <td  height=1 colspan=2 align="right" bgcolor=#cccccc></td>
             </tr>
             <tr>
               <td width=55 align="right">型号：</td>
               <td width=278><font color=#ff6633 size=4><b><%=Rs("Product_Model")%></b></font></td>
             </tr>
             <tr>
               <td  height=1 colspan=2 align="right" bgcolor=#cccccc></td>
             </tr>
             <tr>
               <td align="right" width="55">简介：</td>
               <td width="278"><%=Product_Content%></td>
             </tr>
             <tr>
               <td  height=1 colspan=2 align="right" bgcolor=#cccccc></td>
             </tr>
             <tr>
               <td width=55 align="right">价格：</td>
               <td width=278>
                 <%
		    if Rs("Product_Price")=0 then
			  response.write "<b>备货中...</b>"
			else
			  response.write "<b>" & Rs("Product_Price") & "元</b>"
			end if
		  %>
               </td>
             </tr>
             <tr>
               <td valign=top bgcolor=#cccccc colspan=2 height=1></td>
             </tr>
             <tr>
               <td width=55 align="right" bgcolor="#f7f7f7">购买：</td>
               <td width=278 bgcolor="#f7f7f7"><a href="#" onClick="window.open('../Trade/add.asp?Product_ID=<%=Rs("Product_ID")%>','add','scrollbars=yes,resizable=no,width=650,height=450')">加入购物车</a></td>
             </tr>
           </table>
         </td>
       </tr>
        <tr class=TBBG9>
         <td background="../images/tb_bg09.gif" colspan=2 height="23">&nbsp;&nbsp;<b>详细说明:</b></td>
       </tr>
        <tr class=TBBG9>
         <td colspan=2 style="PADDING-LEFT:1em;PADDING-right:1em"><%=Product_MemoSpec%></td>
       </tr>
       <%	'获取商品点击次数
	  n=Rs("Product_ClickTimes")
	  n=n+1
	  Rs.movenext
    Loop

    '增加商品查看次数
    conn.execute "UPDATE [7Gem_Product] SET [Product_ClickTimes] = " & n & " WHERE [Product_ID] =" & Product_ID
  end if
  Rs.Close
%>
  <td width="77"></td>
      <td width="77"></td>
  </tr>
      </table>
		 </center></div>
      <table width=95% border=0 cellspacing=0 cellpadding=0 align=center>
        <tr>
          <td><%=Replace(GBL_TableBottomString,"770","100%")%></td>
        </tr>
      </table>
			</td>
  </tr>
</table>
<!-- #include file="../Inc/Down.asp" --><% 
'  rs.close
'  set rs=nothing
%>