<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>无标题文档</title>
</head>

<script language="javascript">
<!--
function JM_wu(ob){
  ob.style.display="none";
}
function JM_you(ob){
  ob.style.display="";
}

function uppic(Product_ImgCode)
{
  Product_Code=document.Product_AddForm.Product_Code.value;
  window.open("Admin_Upload.asp?Product_Code=" + Product_Code +"&Product_ImgCode="+Product_ImgCode,"upload", "left=12, top=12, width=600, height=100, resizable=1,menubar=1,scrollbars=1")
}
-->
</script><%
  function MakedownName()
	dim fname
	fname = now()
	fname = replace(fname,"-","")
	fname = replace(fname," ","") 
	fname = replace(fname,":","")
	fname = replace(fname,"PM","")
	fname = replace(fname,"AM","")
	fname = replace(fname,"上午","")
	fname = replace(fname,"下午","")
	fname = int(fname) + int((10-1+1)*Rnd + 1)
	MakedownName=fname
  end function 


	Dim ThisPage

  ThisPage				= Request.ServerVariables("PATH_INFO")

  if Request("Add") ="OK" then
    sql="Select * from [Monolith_Product]"
	rs.open sql, conn , 1, 3
	rs.Addnew
	rs("Product_Name")			= Request.Form("Product_Name")
	rs("Product_Model")			= Request.Form("Product_Model")
	rs("Product_PriceOrigin")	= Request.Form("Product_PriceOrigin")
	rs("Product_Price")			= Request.Form("Product_Price")
	rs("Product_BigName")		= Request.Form("Product_BigName")		
	rs("Product_MidName")		= Request.Form("Product_MidName")
	rs("Product_Area")			= Request.Form("Product_Area")
	rs("Product_Content")		= Request.Form("Product_Content")
	rs("Product_Long")			= Request.Form("Product_Long")
	rs("Product_Height")		= Request.Form("Product_Height")
	rs("Product_Width")			= Request.Form("Product_Width")
	rs("Product_ImgPrev")		= Request.Form("Product_ImgPrev")
	rs("Product_ImgPrevWidth")	= Request.Form("Product_ImgPrevWidth")
	rs("Product_ImgPrevHeight")	= Request.Form("Product_ImgPrevHeight")
	rs("Product_Height")		= Request.Form("Product_Height")
	rs("Product_ImgFull")		= Request.Form("Product_ImgFull")
	rs("Product_ImgFullHeight")	= Request.Form("Product_ImgFullHeight")
	rs("Product_ImgFullWidth")	= Request.Form("Product_ImgFullWidth")
	rs("Product_FileOther")		= Request.Form("Product_FileOther")
	rs("Product_MemoSpec")		= Request.Form("Product_MemoSpec")
	rs("Product_Remark")		= Request.Form("Product_Remark")
	rs("Product_Quantity")		= Request.Form("Product_Quantity")
	rs("Product_Code")			= Request.Form("Product_Code")
	rs.update	
	set rs=nothing
	Response.write "<script>alert('添加成功！');window.document.location.href='Default.asp';</script>"
	Response.end
  end if
%> 
<div align="center">
	<table width="100%" border="0" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="2">
		<b><font color="#FFFFFF">&nbsp;产品管理 &gt;&gt; 产品添加</font></b></td>
	</tr>
	<form name="Product_AddForm" method="post" action="<%=ThisPage%>?Add=OK">
		<tr>
			<td align="right" bgcolor="#999999" width="100" height="20">
			<font color="#FFFFFF">所 属&nbsp; 大 类：</font></td>
			<td  bgcolor="#FFFFFF">
			<select name="Product_BigName" class="InputBox">
			<option value selected>---选择大类---</option>
			<%
			Dim RsMid, SqlMid
		  set RsMid=server.createobject("adodb.recordset") 
		  Sql	= "Select distinct [ProductClass_BigName],[ProductClass_BigOrder] from Monolith_ProductClass order by [ProductClass_BigOrder]" 
		  SqlMid= "Select distinct [ProductClass_BigName],[ProductClass_BigOrder],[ProductClass_MidName],[ProductClass_MidOrder] from Monolith_ProductClass order by [ProductClass_BigOrder],[ProductClass_MidOrder] " 
		  Rs.Open Sql,conn,1,3
		  RsMid.Open SqlMid,conn,1,3
		  Do While Not (Rs.eof Or Rs.Bof)
		    Response.write "<option value='"&Rs("ProductClass_BigName")&"'>"&Rs("ProductClass_BigName")&"</option>"
			Rs.movenext
		  loop
		  Rs.close
		  set Rs=nothing
		%>
			</select> </td>
		</tr>
		<tr>
			<td align="right" bgcolor="#999999" width="100" height="20">
			<font color="#FFFFFF">所 属&nbsp; 中 类：</font></td>
			<td  bgcolor="#FFFFFF">
			<select name="Product_MidName" class="InputBox">
			<option value selected>---选择中类---</option>
			<%
		Do While Not RsMid.eof	
		  Response.write "<option value='"&RsMid("ProductClass_MidName")&"'>"&RsMid("ProductClass_BigName")&"-"&RsMid("ProductClass_MidName")&"</option>"
		  RsMid.movenext
		loop
		RsMid.close
		set RsMid=nothing
	  %>
			</select> </td>
		</tr>
		<tr>
			<td align="right" bgcolor="#999999" width="100" height="20">
			<font color="#FFFFFF">产 品&nbsp; 名 称：</font></td>
			<td  bgcolor="#FFFFFF">
			<input name="Product_Name" type="text" value size="40" maxlength="40" class="InputBox"> <input name="Product_Code" type="hidden" value="<%=MakedownName%>" class="InputBox"></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#999999" width="100" height="20">
			<font color="#FFFFFF">产 品&nbsp; 型 号：</font></td>
			<td  bgcolor="#FFFFFF">
			<input name="Product_Model" type="text" value class="InputBox"></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#999999" width="100" height="20">
			<font color="#FFFFFF">原&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 价：</font></td>
			<td  bgcolor="#FFFFFF">
			<input name="Product_PriceOrigin" type="text" value size="6" maxlength="5" class="InputBox"> 元</td>
		</tr>
		<tr>
			<td align="right" bgcolor="#999999" width="100" height="20">
			<font color="#FFFFFF">最 新&nbsp; 售 价：</font></td>
			<td  bgcolor="#FFFFFF">
			<input name="Product_Price" type="text" value size="6" maxlength="5" class="InputBox"> 元</td>
		</tr>
		<tr>
			<td align="right" bgcolor="#999999" width="100" height="20">
			<font color="#FFFFFF">产&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 地：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Product_Area" type="text" value="深圳" size="12" maxlength="10" class="InputBox"></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#999999" width="100" height="20">
			<font color="#FFFFFF">产 品&nbsp; 简 介：</font></td>
			<td bgcolor="#FFFFFF">
			<textarea name="Product_Content" cols="50" rows="3" class="InputBox"></textarea></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#999999" width="100" height="20">
			<font color="#FFFFFF">外 观&nbsp; 尺 寸：</font></td>
			<td bgcolor="#FFFFFF">
&nbsp;长<input name="Product_Long" type="text" value="11" size="5" maxlength="5" class="InputBox">&nbsp; 宽<input name="Product_Width" type="text" value="11" size="5" maxlength="5" class="InputBox">&nbsp; 高<input name="Product_Height" type="text" value="11" size="5" maxlength="5" class="InputBox"></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#999999" width="100" height="20">
			<font color="#FFFFFF">略 图&nbsp; 地 址：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Product_ImgPrev" type="text" value="/prod/none.gif" size="50" maxlength="100" class="InputBox"> <input type="button" value="上传" onclick="javascript:uppic('Product_ImgPrev');" class="InputBox"></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#999999" width="100" height="20">
			<font color="#FFFFFF">略 图&nbsp; 尺 寸：</font></td>
			<td bgcolor="#FFFFFF">
&nbsp;宽<input name="Product_ImgPrevWidth" type="text" value="81" size="5" maxlength="5" class="InputBox">&nbsp; 高<input name="Product_ImgPrevHeight" type="text" value="76" size="5" maxlength="5" class="InputBox"></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#999999" width="100" height="20">
			<font color="#FFFFFF">全 图&nbsp; 地 址：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Product_ImgFull" type="text" value="/prod/none_d.gif" size="50" maxlength="100" class="InputBox"> <input type="button" value="上传" onclick="javascript:uppic('Product_ImgFull');" class="InputBox"></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#999999" width="100" height="20">
			<font color="#FFFFFF">全 图&nbsp; 尺 寸：</font></td>
			<td bgcolor="#FFFFFF">
&nbsp;宽<input name="Product_ImgFullWidth" type="text" value="220" size="5" maxlength="5" class="InputBox">&nbsp; 高<input name="Product_ImgFullHeight" type="text" value="172" size="5" maxlength="5" class="InputBox"></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#999999" width="100" height="20">
			<font color="#FFFFFF">是 否&nbsp; 推 荐：</font></td>
			<td bgcolor="#FFFFFF">
			<input type="radio" name="Product_ReMark" checked value="0">不推荐<input type="radio" name="Product_ReMark" value="1">推荐</td>
		</tr>
		<tr>
			<td align="right" bgcolor="#999999" width="100" height="20">
			<font color="#FFFFFF">库 存&nbsp; 数 量：</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Product_Quantity" value="10" class="InputBox"></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#999999" width="100" height="20">
			<font color="#FFFFFF">首页推荐产品图片：</font></td>
			<script language="JavaScript">
    function createForm(textN,number)
    {
	  data = "";
	  inter = "'";
	  if (number < 11 && number > -1)
	  {
	    for (i=1; i <= number; i++)
		{
		  if (i < 10) spaces="      ";
		    else spaces="    ";
			data =data + "&nbsp;&nbsp;<Input name="+textN+i+" type=text value='' size='30' maxlength='100'><Input TYPE='button' value='上传"+i+"' onclick=javascript:uppic('Hug"+i+"')><br>";
		}
		if (document.layers)
		{
		  document.layers.cust.document.write(data);
		  document.layers.cust.document.close();
		}
		else
		{
		  if (document.all) {cust.innerHTML = data;}
		}
	  }
	  else
	    {window.alert("请不要超过10张图片.");}
	}
  </script><td bgcolor="#FFFFFF">
			<input type="radio" name="Product_FileOther" checked value="0" onclick="JM_wu(s1)">无<br>
			<input type="radio" name="Product_FileOther" value="1" onclick="JM_you(s1)">有 <span id="s1" style="display:none"><input name="Hug" type="text" value="1" size="1" maxlength="1" class="InputBox">
			<input type="button" value="首页推荐产品图片" onclick="createForm('Hug',document.Product_AddForm.Hug.value);" class="InputBox"><br>
			<span id="cust" style="position:relative;"></span></span></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#999999" width="100" height="20">
			<font color="#FFFFFF">详 细&nbsp; 说 明：</font></td>
			<td bgcolor="#FFFFFF">
			<textarea name="Product_MemoSpec" cols="65" rows="8" class="InputBox"></textarea></td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" height="20">
			　</td>
			<td  bgcolor="#FFFFFF">
			<input type="submit" value=" 添 加 " class="InputBox"> <input type="button" name="Submit2" value=" 返 回 " onClick="window.open('Default.asp', '_self');" class="InputBox"></td>
		</tr>
	</form>
</table>
</div>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->

</html>