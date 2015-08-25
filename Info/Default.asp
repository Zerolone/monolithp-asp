<!-- #include virtual =	"/Inc/Monolith_Conn.asp" -->
<!-- #include virtual = "/Inc/Monolith_Function.asp" -->
<!-- #include virtual = "/Inc/LoginUserCheck.asp" -->
<!-- #include virtual = "/Manage/Info/InfoSetting.asp" -->
<%
	'---网站字符串----临时字符串1---临时字符串2---临时字符串3---临时字符串4
	Dim TheString,    TempString1,  TempString2,  TempString3,  TempString4

	TheString		= ""
	'调首页
	TheString		= TheString & LoadTemplate("02_01_Default")

	'调头
	TheString		= Replace(TheString, "{Monolith_Head}", LoadTemplate("02_02_Default_Head"))

	'Css
	TheString		= Replace(TheString, "{Monolith_Css}", 	LoadTemplate("01_01_Css"))

	'-----------------------------------------------------------------------------------------------------------------------
	'模块：导航
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
	'模块：新闻栏目
	TempString1	= LoadTemplate("03_13_Module_NewsClass")

	TempString2	= ""

	Dim Info_Class_Level_Len, i, j
	Dim Tablecellspacing, Tablecellpadding
	Tablecellspacing	= 1
	Tablecellpadding	= 2
	Info_Class_Level_Len	= 0
	
	'------------------0----------------1-----------------------2-------------------3-------------------4
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
			TempString2 = TempString2 +  "<span class=""Menu"" onmouseover=""HL_Menu(this,0)"" onmouseout=""HL_Menu(this,1)"" onclick=""GoToUrl('List.asp?Info_Class=" & Rs(2) & "', 0);""><img src=""/images/leaf.gif"">&nbsp;" & Rs(3) & "</span>"	& Chr(13)
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

	
  'Sql执行次数+1
  SqlQueryNum	= SqlQueryNum + 1

	TheString		= Replace(TheString, "{Monolith_Module_01_02}", 	TempString1)
	'-----------------------------------------------------------------------------------------------------------------------

	'-----------------------------------------------------------------------------------------------------------------------
	'模块：用户信息
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
	'模块：广告
	TempString1	= LoadTemplate("03_06_Module_Ad")
	'---------------------------0
	Sql = "Select Top 1 [Advertisement_Content] From [Monolith_Advertisement] Where [Advertisement_Class_ID]=2 And [Advertisement_Active]=1 Order By [Advertisement_Order] ASC"
	Rs.Open Sql, Conn, 1, 1
	if Not (Rs.Bof Or Rs.Eof) then
		TempString1 = Replace(TempString1, "{Advertisement_String}",Rs(0))
	end if
	Rs.Close
  'Sql执行次数+1
  SqlQueryNum	= SqlQueryNum + 1

	TheString		= Replace(TheString, "{Monolith_Module_01_04}", 	TempString1)
	'-----------------------------------------------------------------------------------------------------------------------

	'-----------------------------------------------------------------------------------------------------------------------
	'模块：文字链接
	TempString1	= ""
	TempString1	= LoadTemplate("03_14_Module_Links")
	TempString1	= Split(TempString1, "<!--Split-->")
	TempString2	= ""

	'-------------------0--------------1
	Sql = "Select [Links_Title], [Links_Url] From [Monolith_Links] Where [Links_Kind]=1 And [Links_Active]=true Order By [Links_Order]"
	Rs.Open Sql, Conn, 1, 1
	Do While Not(Rs.Bof Or Rs.Eof)
		TempString2	= TempString2 & TempString1(1)
 		TempString2 = Replace(TempString2, "{Links_Title_Left}", 	Rs(0))
		TempString2 = Replace(TempString2, "{Links_URL_Left}", 		Rs(1))
		Rs.MoveNext
		if Not(Rs.Bof Or Rs.Eof) then
	 		TempString2 = Replace(TempString2, "{Links_Title_Right}", 	Rs(0))
			TempString2 = Replace(TempString2, "{Links_URL_Right}", 		Rs(1))
			Rs.MoveNext
		else
	 		TempString2 = Replace(TempString2, "{Links_Title_Right}", 	"")
			TempString2 = Replace(TempString2, "{Links_URL_Right}", 		"")
		end if
	Loop
	Rs.Close
		
	'Sql执行次数+1
	SqlQueryNum	= SqlQueryNum + 1
		
	TheString		= Replace(TheString, "{Monolith_Module_03_01}", 	TempString1(0) & TempString2 & TempString1(2))
	'-----------------------------------------------------------------------------------------------------------------------

	'-----------------------------------------------------------------------------------------------------------------------
	'模块：图片链接
	TempString1	= ""
	TempString1	= LoadTemplate("03_15_Module_Links")
	TempString1	= Replace(TempString1, "{Links_Code}", "文字链接")
	TempString1	= Split(TempString1, "<!--Split-->")
	TempString2	= ""

	'-------------------0--------------1------------2
	Sql = "Select [Links_Title], [Links_Url], [Links_Pic] From [Monolith_Links] Where [Links_Kind]=2 And [Links_Active]=true Order By [Links_Order]"
	Rs.Open Sql, Conn, 1, 1
	Do While Not(Rs.Bof Or Rs.Eof)
		TempString2	= TempString2 & TempString1(1)
 		TempString2 = Replace(TempString2, "{Links_Title_Left}", 	Rs(0))
		TempString2 = Replace(TempString2, "{Links_URL_Left}", 		Rs(1))
		TempString2 = Replace(TempString2, "{Links_Pic_Left}", 		Rs(2))
		Rs.MoveNext
		if Not(Rs.Bof Or Rs.Eof) then
	 		TempString2 = Replace(TempString2, "{Links_Title_Right}", 	Rs(0))
			TempString2 = Replace(TempString2, "{Links_URL_Right}", 		Rs(1))
			TempString2 = Replace(TempString2, "{Links_Pic_Right}", 		Rs(2))
			Rs.MoveNext
		else
	 		TempString2 = Replace(TempString2, "{Links_Title_Right}", 	"磐石")
			TempString2 = Replace(TempString2, "{Links_URL_Right}", 		"http://zerolone/")
			TempString2 = Replace(TempString2, "{Links_Pic_Right}", 		"http://zerolone/images/logo2.jpg")
		end if
	Loop
	Rs.Close
		
	'Sql执行次数+1
	SqlQueryNum	= SqlQueryNum + 1
		
	TheString		= Replace(TheString, "{Monolith_Module_03_02}", 	TempString1(0) & TempString2 & TempString1(2))
	'-----------------------------------------------------------------------------------------------------------------------

	'-----------------------------------------------------------------------------------------------------------------------
	'模块：信息列表
	TempString1	= ""
	TempString2	= LoadTemplate("03_16_Module_InfoList")
	
	Dim Info_Date_Arr, Info_Url
	
	'------------------------0-------------------1
	Sql = "Select [Info_Class_Title], [Info_Class_Level] From [Monolith_Info_Class] Where [Info_Class_Parent_ID]=0 "
	Sql	= Sql + " Order By [Info_Class_Level] Asc"
	Rs.Open Sql, Conn ,1, 1
	Do While Not(Rs.Bof Or Rs.Eof)
		TempString1	= TempString1 & TempString2
	  TempString1	= Replace(TempString1, "{Info_Class_Title}",	Rs(0))
	  TempString1	= Replace(TempString1, "{Info_Class_Level}",	Rs(1))
		
		TempString3	= Split(TempString1, "<!--Split-->")
		TempString4 = ""

		Dim Rs_News
		Set Rs_News	= Server.CreateObject("Adodb.RecordSet")
		'-------------------------0-------------1--------------2------------3----------------4-------------5
		Sql = "Select Top 10 [Info_Title], [Info_Author], [Info_Date], [Info_FileName], [Info_Class], [Info_Class_Name] "
		Sql = Sql + " From [Monolith_Info]"
		Sql = Sql + " Where [Info_Apply]=1"
		Sql = Sql + " And Left([Info_Class], 2)='" & Rs(1) & "' "
		Sql	= Sql + " Order By [Info_Order] Asc, [Info_ID] Desc"
		Rs_News.Open Sql, Conn, 1, 1
		Do While Not(Rs_News.Bof Or Rs_News.Eof)
			TempString4	= TempString4 & TempString3(1)
		  TempString4	= Replace(TempString4, "{Info_Title}",	CutLongStrDot(Rs_News(0), 58, 0))
		  TempString4	= Replace(TempString4, "{Info_Title_All}",	Rs_News(0))
		  TempString4	= Replace(TempString4, "{Info_Date}",				Rs_News(2))

			Info_Date_Arr = Split(Rs_News(2), "-")
			Info_Url = InfoSettingArr(7) & "/" & Info_Date_Arr(0) & "/" & Info_Date_Arr(1) & Info_Date_Arr(2) & "/" & Rs_News(3)
	
		  TempString4	= Replace(TempString4, "{Info_Url}",						Info_Url)
		  TempString4	= Replace(TempString4, "{Info_Class}",					Rs_News(4))
		  TempString4	= Replace(TempString4, "{Info_Class_Title_Sub}",	Rs_News(5))
			Rs_News.MoveNext
		Loop
		Rs_News.Close
	
	  'Sql执行次数+1
	  SqlQueryNum	= SqlQueryNum + 1
		
		TempString1		= TempString3(0) & TempString4 & TempString3(2)

		Rs.MoveNext
	Loop
	Rs.Close

  'Sql执行次数+1
  SqlQueryNum	= SqlQueryNum + 1
	
	TheString		= Replace(TheString, "{Monolith_Module_02_01}", 	TempString1)
	'-----------------------------------------------------------------------------------------------------------------------
	

	'-----------------------------------------------------------------------------------------------------------------------
	'未用到的模块
	TheString		= Replace(TheString, "{Monolith_Module_01_01}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_01_02}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_01_03}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_01_04}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_01_05}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_01_06}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_01_07}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_01_08}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_01_09}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_01_10}" , "")

	TheString		= Replace(TheString, "{Monolith_Module_02_01}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_02}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_03}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_04}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_05}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_06}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_07}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_08}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_09}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_10}" , "")

	TheString		= Replace(TheString, "{Monolith_Module_03_01}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_03_02}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_03_03}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_03_04}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_03_05}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_03_06}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_03_07}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_03_08}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_03_09}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_03_10}" , "")
	'-----------------------------------------------------------------------------------------------------------------------

	'Foot
	TempString1	= LoadTemplate("01_02_Foot")
  TempString1	= Replace(TempString1, "{Now}",	Now())
  TempString1	= Replace(TempString1, "{ExecuteTime}",	FormatNumber((Timer()-Startime)*1000,5) & "毫秒。")
  TempString1	= Replace(TempString1, "{SqlQueryNum}",	SqlQueryNum)
	
	TheString		= Replace(TheString, "{Monolith_Foot}", 	TempString1)

	'分隔线
	TheString		= Replace(TheString, "{Monolith_Sep}", 	LoadTemplate("01_03_Sep"))
	Response.write TheString

	Call CloseConn
%>