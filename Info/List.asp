<!-- #include virtual =	"/Inc/Monolith_Conn.asp" -->
<!-- #include virtual = "/Inc/Monolith_Function.asp" -->
<!-- #include virtual = "/Inc/LoginUserCheck.asp" -->
<!-- #include virtual = "/Manage/Info/InfoSetting.asp" -->
<%
	'---��վ�ַ���----��ʱ�ַ���1---��ʱ�ַ���2---��ʱ�ַ���3---��ʱ�ַ���4
	Dim TheString,    TempString1,  TempString2,  TempString3,  TempString4

	TheString		= ""
	'����ҳ
	TheString		= TheString & LoadTemplate("02_03_List")

	'��ͷ
	TheString		= Replace(TheString, "{Monolith_Head}", LoadTemplate("02_02_Default_Head"))

	'Css
	TheString		= Replace(TheString, "{Monolith_Css}", 	LoadTemplate("01_01_Css"))

	'-----------------------------------------------------------------------------------------------------------------------
	'ģ�飺����
	TheString		= Replace(TheString, "{Monolith_Module_01_01}", 	LoadTemplate("03_01_Module_Nav"))
	if IsLogin = 1 then
	  TheString	= Replace(TheString, "{Display_Code_Login_Show}",		"Block")
	  TheString	= Replace(TheString, "{Display_Code_Login_NoShow}",	"none")
	else
	  TheString	= Replace(TheString, "{Display_Code_Login_Show}",		"none")
	  TheString	= Replace(TheString, "{Display_Code_Login_NoShow}",	"Block")
	end if
	'-----------------------------------------------------------------------------------------------------------------------

	'-----------------------------------------------------------------------------------------------------------------------
	'ģ�飺�û���Ϣ
	TempString1	= ""
	if Users_Name <> "" then
		TempString1	= LoadTemplate("03_03_Module_UserInfo")
		TempString1 = Replace(TempString1, "{Users_Name}", 					Users_Name)
		TempString1 = Replace(TempString1, "{Users_Group}", 				Users_Group)
		TempString1 = Replace(TempString1, "{Users_RegTime}",				Users_RegTime)
		TempString1 = Replace(TempString1, "{Users_Level}", 				Users_Level)
		TempString1 = Replace(TempString1, "{Users_Subject}", 			Users_Subject)
		TempString1 = Replace(TempString1, "{Users_Sex}", 					Users_Sex)
	end if
	TheString		= Replace(TheString, "{Monolith_Module_01_03}", 	TempString1)
	'-----------------------------------------------------------------------------------------------------------------------

	'-----------------------------------------------------------------------------------------------------------------------
	'ģ�飺�����б�
	Dim Info_Class, Info_Class_Title
	Info_Class	= Request("Info_Class")
	
	TempString1	= LoadTemplate("03_16_Module_InfoList")

	Sql	= "Select [Info_Class_Title] From [Monolith_Info_Class] Where [Info_Class_Level]='" & Info_Class & "'"
	Rs.Open Sql, Conn, 1, 1
	if Rs.Bof Or Rs.Eof then
		Response.write "����Ŀ�����ڣ�"	
		Response.end
	else
		Info_Class_Title	= Rs(0)
	end if
	Rs.Close

	TempString1	= Replace(TempString1, "{Info_Class_Title}", Info_Class_Title)
	TempString1	= Replace(TempString1, "{Info_Class_Level}", Info_Class)

	TempString1	= Split(TempString1, "<!--Split-->")
	TempString2	= ""
	
	Dim Info_Date_Arr, Info_Url

	Dim ThisPage , Page, Pages, j
  ThisPage							= Request.ServerVariables("PATH_INFO")
	Page									= clng(request("Page"))

	'------------------0-------------1--------------2------------3----------------4-------------5
	Sql = "Select [Info_Title], [Info_Author], [Info_Date], [Info_FileName], [Info_Class], [Info_Class_Name] From [Monolith_Info]"
	Sql = Sql + " Where [Info_Apply]=1 And Left([Info_Class], " & Len(Info_Class) & ")='" & Info_Class & "' "
	Sql	= Sql + " Order By [Info_ID] Desc"

	Rs.Open Sql, Conn ,1, 1
	if Not (Rs.bof Or Rs.Eof) then
		Rs.PageSize=20
		if page=0 then page=1
		pages=rs.pagecount
		if page > pages then page=pages
		rs.AbsolutePage=page
		For j=1 to rs.PageSize 
			TempString2	= TempString2 & TempString1(1)
		  TempString2	= Replace(TempString2, "{Info_Title}",	CutLongStrDot(Rs(0), 58, 0))
		  TempString2	= Replace(TempString2, "{Info_Title_All}",	Rs(0))
		  TempString2	= Replace(TempString2, "{Info_Date}",				Rs(2))

			Info_Date_Arr = Split(Rs(2), "-")
			Info_Url = InfoSettingArr(7) & "/" & Info_Date_Arr(0) & "/" & Info_Date_Arr(1) & Info_Date_Arr(2) & "/" & Rs(3)
	
		  TempString2	= Replace(TempString2, "{Info_Url}",						Info_Url)
		  TempString2	= Replace(TempString2, "{Info_Class}",					Rs(4))
		  TempString2	= Replace(TempString2, "{Info_Class_Title_Sub}",Rs(5))
			Rs.MoveNext
    	if Rs.Eof then Exit For
    Next

		TempString2	= TempString2	+ "<tr class=""row2"" height=""30"" valign=""bottom"">"
		TempString2	= TempString2	+ "<form method=""Post"" action=""" & ThisPage & """ style=""MARGIN-BOTTOM:0px"">"
		TempString2	= TempString2	+ "<input type=""hidden"" name=""Info_Class"" value=""" & Info_Class & """>"
		TempString2	= TempString2	+ "<td colspan=""3"" align=""center"">"

	  ThisPage = ThisPage & "?1=1"
	  if Info_Class	<> "" then ThisPage = ThisPage + "&Info_Class=" + Info_Class
	  if Page<2 then
	  	TempString2	= TempString2	+ "��&nbsp;<strike>��ҳ</strike>&nbsp;��"
	  	TempString2	= TempString2	+ "<strike>��һҳ</strike>&nbsp;��"
	  else
	  	TempString2	= TempString2	+ "��&nbsp;<a href=""" & ThisPage & "&page=1"">��ҳ</a>&nbsp;��"
	  	TempString2	= TempString2	+ "<a href=""" & ThisPage & "&page=" & Page-1 & """>��һҳ</a>&nbsp;��"
	  end if
		if rs.pagecount-page<1 then
			TempString2	= TempString2	+ "<strike>��һҳ</strike>"
			TempString2	= TempString2	+ "��<strike>βҳ</strike>��"
		else
			TempString2	= TempString2	+ "<a href=""" & ThisPage & "&page=" & page+1 & """>��һҳ</a>&nbsp;��"
			TempString2	= TempString2	+ "<a href=""" & ThisPage & "&page=" & rs.pagecount &""">βҳ</a>��"
		end if
		TempString2	= TempString2	+ "ҳ�Σ�<strong><font color=red>" & Page & "</font>/" & rs.pagecount & "</strong>&nbsp;ҳ&nbsp;��"
		TempString2	= TempString2	+ "	��<b><font color=red>" & rs.recordcount & "</font></b>����¼&nbsp;��"
		TempString2	= TempString2	+ "<b>" & rs.pagesize & "</b>����¼/ҳ&nbsp;��"
		TempString2	= TempString2	+ "ת����<input type=""text"" name=""page"" size=4 maxlength=10 class=""InputBox"" value=""" & page & """>&nbsp;"
		TempString2	= TempString2	+ "<input class=""InputBox"" type=""submit""  value="" Goto ""  name=""cndok"">"
		TempString2	= TempString2	+ "</td>"
		TempString2	= TempString2	+ "</form>"
		TempString2	= TempString2	+ "</tr>"
	end if
	Rs.Close

  'Sqlִ�д���+1
  SqlQueryNum	= SqlQueryNum + 1
	
	TheString		= Replace(TheString, "{Monolith_Module_02_01}", 	TempString1(0) & TempString2 & TempString1(2))
	'-----------------------------------------------------------------------------------------------------------------------


	'-----------------------------------------------------------------------------------------------------------------------
	'ģ�飺������Ŀ
	TempString1	= LoadTemplate("03_13_Module_NewsClass")

	TempString2	= ""

	Dim Info_Class_Level_Len, i
	Dim Tablecellspacing, Tablecellpadding
	Tablecellspacing	= 1
	Tablecellpadding	= 2
	Info_Class_Level_Len	= 0
	
	'--------------------0------------------1------------------2-----------------3------------------4
	Sql	= "Select [Info_Class_ID], [Info_Class_Parent_ID], [Info_Class_Level], [Info_Class_Title], [Info_Class_HasChild] From [Monolith_Info_Class] Order By [Info_Class_Level]"
	Rs.Open Sql, Conn ,1 ,1
	Do While Not (Rs.Bof Or Rs.Eof)

'A
		if Info_Class_Level_Len = 0 then
			TempString2 = TempString2 +  "<table border=""0"" width=""100%"" cellspacing=""" & Tablecellspacing & """ cellpadding=""" & Tablecellpadding & """>"	& Chr(13)
		end if

'B
		if Info_Class_Level_Len = Len(Rs(2)) then
			TempString2 = TempString2 +  "</td>"	& Chr(13)
			TempString2 = TempString2 +  "</tr>"	& Chr(13)
		end if

'C
		if Info_Class_Level_Len < Len(Rs(2)) And Info_Class_Level_Len <> 0 then
			TempString2 = TempString2 +  "<table border=""0"" width=""100%"" cellspacing=""" & Tablecellspacing & """ cellpadding=""" & Tablecellpadding & """ style=""display:none"" id=""Node_" & Rs(1) & """>"	& Chr(13)
		end if

'D
		if Info_Class_Level_Len > Len(Rs(2)) then
			j = (Info_Class_Level_Len-Len(Rs(2))) \ 2
		  for i = 1 to j
	  		TempString2 = TempString2 +  "</td>"															& Chr(13)
	  		TempString2 = TempString2 +  "</tr>"															& Chr(13)
		  	TempString2 = TempString2 +  "</table>"														& Chr(13)
			next
	  		TempString2 = TempString2 +  "</td>"															& Chr(13)
	  		TempString2 = TempString2 +  "</tr>"															& Chr(13)
		end if

		TempString2 = TempString2 +  "<tr>"																	& Chr(13)
		TempString2 = TempString2 +  "<td>"																	& Chr(13)
		TempString2 = TempString2 +  "<img src=""../images/blank.gif"" height=""1"" width=""" & (Len(Rs(2))-2) / 2 * 10 & """ >"	& Chr(13)

		if Rs(4) then
			TempString2 = TempString2 +  "<span class=""Menu"" onmouseover=""HL_Menu(this,0)"" onmouseout=""HL_Menu(this,1)"" onclick=""ExtendNode('Node_" & Rs(0) & "');""><img  id=""Img_Node_" & Rs(0) & """ src=""../images/shrink.gif"">&nbsp;" & Rs(3) & "</span>"	& Chr(13)
		else
			TempString2 = TempString2 +  "<span class=""Menu"" onmouseover=""HL_Menu(this,0)"" onmouseout=""HL_Menu(this,1)"" onclick=""GoToUrl('?Info_Class=" & Rs(2) & "', 0);""><img src=""/images/leaf.gif"">&nbsp;" & Rs(3) & "</span>"	& Chr(13)
		end if

		Dim Info_Class_Title_Len, Info_Class_Title_Normal
		Info_Class_Title_Normal	=	16
		Info_Class_Title_Len		=	CountStr(Rs(3))

		Info_Class_Level_Len	=  Len(Rs(2))
		Rs.MoveNext
	Loop
	Rs.Close

	TempString2 = TempString2 +  "</td>"																		& Chr(13)
	TempString2 = TempString2 +  "</tr>"																		& Chr(13)
 	TempString2 = TempString2 +  "</table>"																	& Chr(13)
	TempString2 = TempString2 +  "</td>"																		& Chr(13)
	TempString2 = TempString2 +  "</tr>"																		& Chr(13)
 	TempString2 = TempString2 +  "</table>"																	& Chr(13)

	TempString1	= Replace(TempString1, "{News_Class}", TempString2)

	
  'Sqlִ�д���+1
  SqlQueryNum	= SqlQueryNum + 1

	TheString		= Replace(TheString, "{Monolith_Module_01_02}", 	TempString1)
	'-----------------------------------------------------------------------------------------------------------------------


	'-----------------------------------------------------------------------------------------------------------------------
	'δ�õ���ģ��
	TheString		= Replace(TheString, "{Monolith_Module_01_04}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_01_05}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_01_06}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_01_07}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_01_08}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_01_09}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_01_10}" , "")

	TheString		= Replace(TheString, "{Monolith_Module_02_02}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_03}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_04}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_05}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_06}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_07}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_08}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_09}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_10}" , "")
	'-----------------------------------------------------------------------------------------------------------------------

	'Foot
	TempString1	= LoadTemplate("01_02_Foot")
  TempString1	= Replace(TempString1, "{Now}",	Now())
  TempString1	= Replace(TempString1, "{ExecuteTime}",	FormatNumber((Timer()-Startime)*1000,5) & "���롣")
  TempString1	= Replace(TempString1, "{SqlQueryNum}",	SqlQueryNum)
	
	TheString		= Replace(TheString, "{Monolith_Foot}", 	TempString1)

	'�ָ���
	TheString		= Replace(TheString, "{Monolith_Sep}", 	LoadTemplate("01_03_Sep"))
	Response.write TheString

	Call CloseConn
%>