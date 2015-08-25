<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp"-->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim ThisPage
	Dim Advertisement_ID, Cmd

	Dim Page
	Page	= Request("Page")

  ThisPage					= Request.ServerVariables("PATH_INFO")
	if Request("Code") = "Edit" then
		Dim Advertisement_FileName, Advertisement_FullName, Advertisement_Url, Advertisement_Width, Advertisement_Height, Advertisement_Color, TheString
		Dim Advertisement_No
		Dim LayerID
		
		Advertisement_FileName	= Request("Advertisement_FileName")
'		Advertisement_FullName	= "Http://Www.DigiChina.cn/Advertisement/Src/" & Advertisement_FileName
		Advertisement_FullName	= "/Advertisement/Src/" & Advertisement_FileName
		Advertisement_Width			= Request("Advertisement_Width")
		Advertisement_Height		= Request("Advertisement_Height")
		Advertisement_Url				= "http://" & Replace(Request("Advertisement_Url"), "http://", "")
		Advertisement_Color			= Request("Advertisement_Color")
		Advertisement_No				= Request("Advertisement_No")
		if Advertisement_No = "" then Advertisement_No = MakeRndNum
		LayerID = year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
				
		Dim FileExt, Advertisement_Content
		Advertisement_Content = ""
		FileExt=lcase(right(Advertisement_FileName,4))
		
		If FileEXT = ".swf" then
			Advertisement_Content	= Advertisement_Content + "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0' width='" & Advertisement_Width & "' height='" & Advertisement_height & "'>"
			Advertisement_Content	= Advertisement_Content + "<param name='movie' value='" & Advertisement_FullName & "'>"
			Advertisement_Content	= Advertisement_Content + "<param name='quality' value='high'>"
			Advertisement_Content	= Advertisement_Content + "<PARAM NAME=wmode VALUE=transparent>"
			Advertisement_Content	= Advertisement_Content + "</object>"
		Else
			Advertisement_Content	= Advertisement_Content + "<img border='0' width='" & Advertisement_Width & "' height='" & Advertisement_Height & "' src='" & Advertisement_FullName & "'>"
		End If
		
		TheString = ""
		TheString = TheString + "<table width='" & Advertisement_Width &"' border='0' cellspacing='0' cellpadding='0'>"
		TheString = TheString + "<tr>"
		TheString = TheString + "<td height='" & Advertisement_Height & "' bgcolor='" & Advertisement_Color & "'>"
		TheString = TheString + "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
		TheString = TheString + "<tr>"
		TheString = TheString + "<td>"
		TheString = TheString + "<div id='LayerTop_" & LayerID & "' style='position:absolute; width:0px; height:0px; z-index:1'>"
		TheString = TheString + "<div id='Layer_"    & LayerID & "' style='position:absolute; left:0px; top:0px; width:" & Advertisement_Width & "px; height:" & Advertisement_Height & "px; z-index:2; filter:alpha(opacity=0)'>"
		TheString = TheString + "<table width='100%' height='100%' border='0' cellspacing='0' cellpadding='0'>"
		TheString = TheString + "<tr>"
		TheString = TheString + "<td style='cursor:hand' onClick=window.open('/Advertisement/Default.asp?Advertisement_No=" & Advertisement_No & "');>&nbsp;</td>"
		TheString = TheString + "</tr>"
		TheString = TheString + "</table>"
		TheString = TheString + "</div>"
		TheString = TheString + "</div>"
		TheString = TheString + Advertisement_Content
		TheString = TheString + "</td>"
		TheString = TheString + "</tr>"
		TheString = TheString + "</table>"
		TheString = TheString + "</td>"
		TheString = TheString + "</tr>"
		TheString = TheString + "</table>"
		
		Advertisement_ID = Request("Advertisement_ID")
		CheckNum Advertisement_ID
		'-----------------------0---------------------1--------------------2------------------------3---------------------4----------------------5--------------------------6------------------------7---------------------------8--------------------9--------------------10--------------------11------
		Sql = "Select [Advertisement_Title], [Advertisement_Hits], [Advertisement_Width], [Advertisement_Height], [Advertisement_Url], [Advertisement_FileName], [Advertisement_Class_ID], [Advertisement_Active], [Advertisement_Content], [Advertisement_Color], [Advertisement_No], [Advertisement_Order] From [Monolith_Advertisement] Where [Advertisement_ID]=" & Advertisement_ID
		Rs.Open Sql, Conn, 1, 3
		Rs(0)		= Request("Advertisement_Title")
		Rs(1)		= Request("Advertisement_Hits")
		Rs(2)		= Advertisement_Width
		Rs(3)		= Advertisement_Height
		Rs(4)		= Advertisement_Url
		Rs(5)		= Advertisement_FileName
		Rs(6)		= Request("Advertisement_Class_ID")
		Rs(7)		= Request("Advertisement_Active")
		Rs(8)		= TheString
		Rs(9)		= Advertisement_Color
		Rs(10)	= Advertisement_No
		Rs(11)	= Request("Advertisement_Order")
		Rs.Update
		Call CloseRs
		Call CloseConn

		Response.Redirect("Default.asp?Page=" & Page)
		Response.end
	end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
</head>

<body>
<%
	Dim TrBgColor
	
	Advertisement_ID = Request("Advertisement_ID")
	CheckNum Advertisement_ID
	'---------------------0----------------------1---------------------2---------------------------3-----------------4------------------------5----------------------6-------------------------7-----------------------8----------------------9-----------------------10-------------------11---------------------12
	Sql = "SELECT [Advertisement_ID], [Advertisement_Title], [Advertisement_Hits], [Advertisement_Width], [Advertisement_Height], [Advertisement_Class_ID], [Advertisement_Active], [Advertisement_FileName], [Advertisement_Url], [Advertisement_Content], [Advertisement_Color], [Advertisement_No], [Advertisement_Order] FROM [Monolith_Advertisement] Where [Advertisement_ID]=" & Advertisement_ID
	Rs.Open Sql, Conn, 1, 1
	If Rs.Bof or Rs.Eof then
	  Response.write "这个广告不存在！"
		Response.end
	end if	

%>
<form name="form1" method="post" action="Del_Submit.asp">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="3">
			<b><font color="#FFFFFF">&nbsp;广告管理 &gt;&gt; 删除广告</font></b></td>
		</tr>  <tr>
    <td width="100" height="20" bgcolor="#999999" align="right"><font color="#FFFFFF">标&nbsp;&nbsp;&nbsp; 题：</font></td>
    <td bgcolor="#FFFFFF">
      <input name="Advertisement_Title" type="text" id="Advertisement_Title" value="<%=Trim(Rs(1))%>" size="40" class="InputBox" disabled>
    </td>
    <td rowspan="10" bgcolor="#FFFFFF">　</td>
  </tr>
  <tr>
    <td bgcolor="#999999" width="100" height="20" align="right"><font color="#FFFFFF">点&nbsp;&nbsp;&nbsp; 击：</font></td>
    <td bgcolor="#FFFFFF">
      <input name="Advertisement_Hits" type="text" id="Advertisement_Hits" value="<%=Rs(2)%>" size="40" class="InputBox" disabled>
    </td>
    </tr>
  <tr>
    <td height="20" bgcolor="#999999" width="100" align="right"><font color="#FFFFFF">长&nbsp;&nbsp;&nbsp; 度：</font></td>
    <td bgcolor="#FFFFFF">
      <input name="Advertisement_Width" type="text" id="Advertisement_Width" value="<%=Rs(3)%>" size="40" class="InputBox" disabled>
    </td>
    </tr>
  <tr>
    <td bgcolor="#999999" width="100" height="20" align="right"><font color="#FFFFFF">宽&nbsp;&nbsp;&nbsp; 度：</font></td>
    <td bgcolor="#FFFFFF">
      <input name="Advertisement_Height" type="text" id="Advertisement_Height" value="<%=Rs(4)%>" size="40" class="InputBox" disabled>
    </td>
    </tr>
  <tr>
    <td bgcolor="#999999" width="100" height="20" align="right"><font color="#FFFFFF">所&nbsp;&nbsp;&nbsp; 属：</font></td>
    <td bgcolor="#FFFFFF">
      <select name="Advertisement_Class_ID" id="Advertisement_Class_ID" class="InputBox" disabled>
        <%
			Dim Rs1
			Set Rs1		= Server.CreateObject("ADODB.Recordset")
			Sql = "Select [Advertisement_Class_Name], [Advertisement_Class_ID] From [Monolith_Advertisement_Class] Order By [Advertisement_Class_Order]"
			Rs1.Open Sql, Conn ,1, 3
			Do While Not (Rs1.Bof Or Rs1.Eof)
		%>
        <option value="<%=Rs1(1)%>" <% if Trim(Rs1(1)) = Trim(Rs(5)) then Response.write "Selected"  %>><%=Trim(Rs1(0))%></option>
        <%
				Rs1.MoveNext
			Loop
			Rs1.Close
			Set Rs1=Nothing
		%>
      </select>
    </td>
    </tr>	
  <tr>
    <td bgcolor="#999999" width="100" height="20" align="right"><font color="#FFFFFF">使&nbsp;&nbsp;&nbsp; 用：</font></td>
    <td bgcolor="#FFFFFF">
      <input name="Advertisement_Active" type="radio" value="0" id="Advertisement_Active_0" <% if Rs(6)=0 then Response.write "checked" %>><label for="Advertisement_Active_0" disabled>不使用</label>
      <input type="radio" name="Advertisement_Active" value="1" id="Advertisement_Active_1" <% if Rs(6)=1 then Response.write "checked" %>><label for="Advertisement_Active_1" disabled>使用</label>
    </td>
    </tr>	
  <tr>
    <td bgcolor="#999999" width="100" height="20" align="right"><font color="#FFFFFF">文 件 名：</font></td>
    <td bgcolor="#FFFFFF">
      <input name="Advertisement_FileName" type="text" id="Advertisement_FileName" value="<%=Trim(Rs(7))%>" size="40" class="InputBox" disabled> </td>
    </tr>
  <tr>
    <td bgcolor="#999999" width="100" height="20" align="right"><font color="#FFFFFF">U&nbsp;&nbsp; r&nbsp; l：</font></td>
    <td bgcolor="#FFFFFF">
      <input name="Advertisement_Url" type="text" id="Advertisement_Url" value="<%=Trim(Rs(8))%>" size="40" class="InputBox" disabled>
    </td>
    </tr>
  <tr>
    <td bgcolor="#999999" width="100" height="20" align="right"><font color="#FFFFFF">背景颜色：</font></td>
    <td bgcolor="#FFFFFF">
      <input name="Advertisement_Color" type="text" id="Advertisement_Color" value="<%=Rs(10)%>" size="40" class="InputBox" disabled>&nbsp;
      </td>
    </tr>
  <tr>
    <td bgcolor="#999999" width="100" height="20" align="right"><font color="#FFFFFF">顺&nbsp;&nbsp;&nbsp; 序：</font></td>
    <td bgcolor="#FFFFFF"><input name="Advertisement_Order" type="text" id="Advertisement_Order" value="<%=Rs(12)%>" size="40" class="InputBox" disabled></td>
  </tr>
  <tr>
    <td bgcolor="#999999" width="100" height="20" align="right">　</td>
    <td colspan="2" bgcolor="#FFFFFF">
			<input type="submit" name="Submit0" value=" 删 除 广 告 (Alt + D) " class="InputBox" accesskey="D">
			<input type="reset" value=" 返 回 (Alt + N) " name="B2" class="InputBox" accesskey="N" onclick="window.open('Default.asp?Page=<%=Page%>','_self');">
			<input type="hidden" name="Advertisement_ID" value="<%=Advertisement_ID%>">
    <input type="hidden" name="Page" value="<%=Page%>">
  </tr>
</table>
</form>
<%
		Call CloseRs
		Call CloseConn
%><table><tr><td height="198"></td></tr></table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->