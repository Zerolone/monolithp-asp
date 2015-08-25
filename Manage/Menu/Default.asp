<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
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
		<td colspan="4"><b><font color="#FFFFFF">&nbsp;系统管理 &gt;&gt; 系统菜单管理</font></b></td>
	</tr>
	<tr>
		<td><%
	Dim Tree_Level_Len, i, j
	Dim Tablecellspacing, Tablecellpadding
	Tablecellspacing	= 1
	Tablecellpadding	= 2
	Tree_Level_Len	= 0
	
	'------------------0-------------1---------------2-----------3-------------4------------5---------------6----------
	Sql	= "Select [Tree_ID], [Tree_Parent_ID], [Tree_Level], [Tree_Title], [Tree_Url], [Tree_Target], [Tree_HasChild] From [Monolith_Tree] Order By [Tree_Level]"
	Rs.Open Sql, Conn ,1 ,1

	Do While Not (Rs.Bof Or Rs.Eof)
'A
		if Tree_Level_Len = 0 then
			Response.write "<table border=""0"" width=""100%"" cellspacing=""" & Tablecellspacing & """ cellpadding=""" & Tablecellpadding & """>"	& Chr(13)
		end if

'B
		if Tree_Level_Len = Len(Rs(2)) then
			Response.write "</td>"	& Chr(13)
			Response.write "</tr>"	& Chr(13)
		end if

'C
		if Tree_Level_Len < Len(Rs(2)) And Tree_Level_Len <> 0 then
			Response.write "<table  bgcolor=""#FFFFFF"" border=""0"" width=""100%"" cellspacing=""" & Tablecellspacing & """ cellpadding=""" & Tablecellpadding & """ style=""display:none"" id=""Node_" & Rs(1) & """>"	& Chr(13)
		end if

'D
		if Tree_Level_Len > Len(Rs(2)) then
			j = (Tree_Level_Len-Len(Rs(2))) \ 2
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
		if Rs(6) then 
			Response.write "<span class=""Menu"" onmouseover=""HL_Menu(this,0)"" onmouseout=""HL_Menu(this,1)"" onclick=""ExtendNode('Node_" & Rs(0) & "');""><img  id=""Img_Node_" & Rs(0) & """ src=""../images/shrink.gif"">&nbsp;" & Rs(3) & "</span>"	& Chr(13)
		else
			Response.write "<span class=""Menu"" onmouseover=""HL_Menu(this,0)"" onmouseout=""HL_Menu(this,1)""><img src=""../images/leaf.gif"">&nbsp;" & Rs(3) & "</span>"														& Chr(13)
		end if

		Dim Tree_Title_Len, Tree_Title_Normal
		Tree_Title_Normal	=	30
		Tree_Title_Len		=	CountStr(Rs(3))

		if Tree_Title_Len < Tree_Title_Normal then
			LoopNBSP(Tree_Title_Normal - Tree_Title_Len)
		end if


'		if Not (Left(Rs(2),2) = "98" or Left(Rs(2),2) = "99") then
			Response.write "<a href=""Add.asp?Tree_Level_Code=Same_Level&Tree_Parent_ID="	& Rs(1) & """				target=""tabWin"" onclick=""fnClick();"">添加同级(" & Rs(1) & ")</a>" 	& Chr(13)
			Response.write "<a href=""Add.asp?Tree_Level_Code=Lower_Level&Tree_Parent_ID=" & Rs(0) & "&Tree_Level=" & Rs(2) & """ 	target=""tabWin"" onclick=""fnClick();"">添加下级(" & Rs(1) & ")</a>" 	& Chr(13)
			Response.write "<a href=""Edit.asp?Tree_ID=" & Rs(0) & """								target=""tabWin"" onclick=""fnClick();"">修改(" & Rs(1) & ")</a>"	& Chr(13)
			if Not Rs(6) then		Response.write "<a href=""Del.asp?Tree_ID=" & Rs(0) & """						>删除(" & Rs(1) & ")</a>"				& Chr(13)
'		end if

		Tree_Level_Len	=  Len(Rs(2))
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
</table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->
