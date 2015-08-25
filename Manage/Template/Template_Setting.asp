<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<script language="javascript" src="/Js/All.js"></script>
</head>

<body onselectstart="return false;">

<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="4"><b><font color="#FFFFFF">&nbsp;信息发布系统 &gt;&gt; </font></b>
		<font color="#FFFFFF"><b>信息栏目管理</b></font></td>
	</tr>
	<tr>
		<td>
<%
	Dim Template_Level_Len, i, j
	Dim Tablecellspacing, Tablecellpadding
	Tablecellspacing	= 1
	Tablecellpadding	= 2
	Template_Level_Len	= 0
	
	'--------------------0------------------1------------------2-----------------3------------------4
	Sql	= "Select [Template_ID], [Template_Parent_ID], [Template_Level], [Template_Title], [Template_HasChild] From [Monolith_Template] Order By [Template_Level]"

	Rs.Open Sql, Conn ,1 ,1

	Do While Not (Rs.Bof Or Rs.Eof)
'A
		if Template_Level_Len = 0 then
			Response.write "<table border=""0"" width=""100%"" cellspacing=""" & Tablecellspacing & """ cellpadding=""" & Tablecellpadding & """>"	& Chr(13)
		end if

'B
		if Template_Level_Len = Len(Rs(2)) then
			Response.write "</td>"	& Chr(13)
			Response.write "</tr>"	& Chr(13)
		end if

'C
		if Template_Level_Len < Len(Rs(2)) And Template_Level_Len <> 0 then
			Response.write "<table bgcolor=""#FFFFFF""  border=""0"" width=""100%"" cellspacing=""" & Tablecellspacing & """ cellpadding=""" & Tablecellpadding & """ style=""display:none"" id=""Node_" & Rs(1) & """>"	& Chr(13)
		end if

'D
		if Template_Level_Len > Len(Rs(2)) then
			j = (Template_Level_Len-Len(Rs(2))) \ 2
		  for i = 1 to j
	  		Response.write "</td>"															& Chr(13)
	  		Response.write "</tr>"															& Chr(13)
		  	Response.write "</table>"														& Chr(13)
			next
	  		Response.write "</td>"															& Chr(13)
	  		Response.write "</tr>"															& Chr(13)
		end if

		Response.write "<tr>"																	& Chr(13)
		Response.write "<td>"																	& Chr(13)
		Response.write "<img src=""../images/blank.gif"" height=""1"" width=""" & (Len(Rs(2))-2) / 2 * 10 & """ >"	& Chr(13)

		if Rs(4) then
			Response.write "<span class=""Menu"" onmouseover=""HL_Menu(this,0)"" onmouseout=""HL_Menu(this,1)"" onclick=""ExtendNode('Node_" & Rs(0) & "');""><img  id=""Img_Node_" & Rs(0) & """ src=""../images/shrink.gif"">&nbsp;" & Rs(3) & "</span>"	& Chr(13)
		else
			Response.write "<span class=""Menu"" onmouseover=""HL_Menu(this,0)"" onmouseout=""HL_Menu(this,1)""><img src=""../images/leaf.gif"">&nbsp;" & Rs(3) & "</span>"														& Chr(13)
		end if

		Dim Template_Title_Len, Template_Title_Normal
		Template_Title_Normal	=	40
		Template_Title_Len		=	CountStr(Rs(3))

		if Template_Title_Len < Template_Title_Normal then
			LoopNBSP(Template_Title_Normal - Template_Title_Len)
		end if


'		if Not (Left(Rs(2),2) = "98" or Left(Rs(2),2) = "99") then
			Response.write "<a href=""Template_Add.asp?Template_Level_Code=Same_Level&Template_Parent_ID="	& Rs(1) & """ target=""tabWin"" onclick=""fnClick();"">添加同级(" & Rs(1) & ")</a>" & Chr(13)
			Response.write "<a href=""Template_Add.asp?Template_Level_Code=Lower_Level&Template_Parent_ID="	& Rs(0) & "&Template_Level=" & Rs(2) & """ target=""tabWin"" onclick=""fnClick();"">添加下级(" & Rs(1) & ")</a>" & Chr(13)
			Response.write "<a href=""Template_Edit.asp?Template_ID=" & Rs(0) & """ target=""tabWin"" onclick=""fnClick();"">修改(" & Rs(1) & ")</a>"			& Chr(13)
			if Not Rs(4) then
				Response.write "<a href=""Template_Del.asp?Template_ID=" & Rs(0) & """>删除(" & Rs(1) & ")</a>"			& Chr(13)
			end if
'		end if


		Template_Level_Len	=  Len(Rs(2))
		Rs.MoveNext
	Loop

	Response.write "</td>"																		& Chr(13)
	Response.write "</tr>"																		& Chr(13)
 	Response.write "</table>"																	& Chr(13)
	Response.write "</td>"																		& Chr(13)
	Response.write "</tr>"																		& Chr(13)
 	Response.write "</table>"																	& Chr(13)
%></td>
	</tr>
</table><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" --></body></html>