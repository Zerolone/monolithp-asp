<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<link href="../Css/Manage.Css" rel="stylesheet" type="text/css">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<script language="javascript" src="/Js/All.js"></script>
<script language="javascript" src="/Js/LeftBar.js"></script>
</head>
<body onresize="showOrHide()" onselectstart="return false;" scroll="no">
<table border="0" bgcolor="#0099FF" width="100%" height="21" cellspacing="0" cellpadding="0">
	<tr>
		<td width=* class="Menu"> </td>
		<td align="center"	title="同步左侧菜单"	width="65"	class="Menu" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)"  onclick="javascript:location.reload();"><font color="#FFFFFF"> 同 步 </font></td>
		<td align="right"		title="关闭左侧菜单" 	width="15"	class="Menu" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)"  onclick="HideLeftBar();"><img name="toolBar" src="images/Menu_Hide.gif"></td>
	</tr>
</table>
<%
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
			Response.write "<table border=""0"" width=""100%"" cellspacing=""" & Tablecellspacing & """ cellpadding=""" & Tablecellpadding & """ style=""display:none"" id=""Node_" & Rs(1) & """>"	& Chr(13)
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
		Response.write "<img src=""images/blank.gif"" height=""1"" width=""" & (Len(Rs(2))-2) / 2 * 10 & """ >"	& Chr(13)
		if Rs(6) then 
			Response.write "<span class=""Menu"" onmouseover=""HL_Menu(this,0)"" onmouseout=""HL_Menu(this,1)"" onclick=""ExtendNode('Node_" & Rs(0) & "');""><img  id=""Img_Node_" & Rs(0) & """ src=""images/shrink.gif"">&nbsp;" & Rs(3) & "</span>"	& Chr(13)
		else
			Response.write "<span class=""Menu"" onmouseover=""HL_Menu(this,0)"" onmouseout=""HL_Menu(this,1)"" onclick=""ffClick('" & Rs(4) & "','" & Rs(3) & "');""><img src=""images/leaf.gif"">&nbsp;" & Rs(3) & "</span>"				& Chr(13)
		end if

		Tree_Level_Len	= Len(Rs(2))
		Rs.MoveNext
	Loop

	Response.write "</td>"																		& Chr(13)
	Response.write "</tr>"																		& Chr(13)
 	Response.write "</table>"																	& Chr(13)
	Response.write "</td>"																		& Chr(13)
	Response.write "</tr>"																		& Chr(13)
 	Response.write "</table>"																	& Chr(13)
%>