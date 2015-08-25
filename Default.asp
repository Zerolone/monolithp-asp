<!-- #include virtual = "/Inc/Monolith_Conn.asp" -->
<!-- #include virtual = "/Inc/Monolith_Function.asp" -->
<!-- #include virtual = "/Inc/LoginUserCheck.asp" -->
<!-- #include virtual = "/Manage/News/NewsSetting.asp" -->
<%
	'---网站字符串----临时字符串1---临时字符串1
	Dim TheString,    TempString1,  TempString2 

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
	'模块：论坛新帖子
	TempString1	= LoadTemplate("03_02_Module_NewSubject")
	TempString1	= Split(TempString1, "<!--Split-->")
	TempString2	= ""
	
 	'------------------------------0---------------------1--------------------------2---------------------------3---------------------------4-----------------------------5---------------------------6-------------------------7
	Sql = "Select Top 5 [MyBBS_BBS].[BBS_ID], [MyBBS_BBS].[BBS_Subject], [MyBBS_BBS].[BBS_Board_ID], [MyBBS_BBS].[BBS_PostUser], [MyBBS_BBS].[BBS_PostUserID], [MyBBS_BBS].[BBS_PostTime], [MyBBS_BBS].[BBS_Icon], [MyBBS_Board].[Board_Name] From [MyBBS_BBS], [MyBBS_Board]  Where [MyBBS_BBS].[BBS_Board_ID]=[MyBBS_Board].[Board_ID] And [MyBBS_BBS].[BBS_Root]=0 order by [MyBBS_BBS].[BBS_ID] Desc"
	Rs.Open Sql, Conn, 1, 1
	Do While Not(Rs.Bof Or Rs.Eof)
		TempString2	= TempString2 & TempString1(1)
 		TempString2 = Replace(TempString2, "{BBS_ID}", 						Rs(0))
	  TempString2 = Replace(TempString2, "{BBS_Subject}"	,			Rs(1))
	  TempString2 = Replace(TempString2, "{Board_ID}"	,					Rs(2))
	  TempString2 = Replace(TempString2, "{BBS_PostUser}"	,			Rs(3))
	  TempString2 = Replace(TempString2, "{BBS_PostUserID}",		Rs(4))
	  TempString2 = Replace(TempString2, "{BBS_PostTime}"	,			Rs(5))
	  TempString2 = Replace(TempString2, "{BBS_Icon}"	,					Rs(6))
	  TempString2 = Replace(TempString2, "{BBS_Board_Name}",		Rs(7))
		Rs.MoveNext
	Loop
	Rs.Close

	'Sql执行次数+1
	SqlQueryNum	= SqlQueryNum + 1
		
	TheString		= Replace(TheString, "{Monolith_Module_01_02}", 	TempString1(0) & TempString2 & TempString1(2))
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
		
	TheString		= Replace(TheString, "{Monolith_Module_03_05}", 	TempString1(0) & TempString2 & TempString1(2))
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
		
	TheString		= Replace(TheString, "{Monolith_Module_01_05}", 	TempString1(0) & TempString2 & TempString1(2))
	'-----------------------------------------------------------------------------------------------------------------------


	'-----------------------------------------------------------------------------------------------------------------------
	'模块：公告
	TempString1	= LoadTemplate("03_04_Module_Bulletin")

	Dim View_01, View_02, View_03, View_04, View_05, View_06, View_07, Bulletin_String
	View_01	= 1
	View_02	= 5
	View_03	= 20
	View_04	= "100%"
	View_05	= 5
	View_06 = 1
	View_07	= ""

	'---------------------------------------0---------------1-----------------2
	Sql = "Select Top " & View_02 & " [Bulletin_ID], [Bulletin_Title], [Bulletin_Date] From [Monolith_Bulletin] Order By [Bulletin_ID] Desc"
	Rs.Open Sql, Conn, 1, 1
	Do While Not(Rs.Bof Or Rs.Eof)
		Bulletin_String = Bulletin_String + "<A href=/Bulletin/Bulletin.asp?Bulletin_ID=" & Rs(0) & "  target='_blank'>" & rs(1)
    if View_06 = 1 then Bulletin_String = Bulletin_String & "&nbsp;[" & Rs(2) & "]"
    Bulletin_String = Bulletin_String & "</A>"
    if View_01 > 2 then
      Bulletin_String = Bulletin_String & "<BR><BR>"
    else
    	Bulletin_String = Bulletin_String & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    end if
	  Rs.MoveNext
	Loop
	Rs.Close

	Bulletin_String	=	"<marquee onmouseover='this.scrollAmount=0' onmouseout='this.scrollAmount=" & View_05 & "' scrollamount='" & View_05 & "' direction='" & View_01 & "' height='" & View_03 & "' width='" & View_04 & "' bgcolor='" & View_07 & "'>" & Bulletin_String & "</marquee></td>"

  'Sql执行次数+1
  SqlQueryNum	= SqlQueryNum + 1

	TempString1	= LoadTemplate("03_04_Module_Bulletin")
	TempString1	= Replace(TempString1, "{Bulletin_String}", Bulletin_String)
	TheString		= Replace(TheString, "{Monolith_Module_02_01}", 	TempString1)
	'-----------------------------------------------------------------------------------------------------------------------

	'-----------------------------------------------------------------------------------------------------------------------
	'模块：新闻
	TempString1	= ""
	TempString2	= LoadTemplate("03_05_Module_News")
	Dim News_Date_Arr, News_Url
	
	'------------------------0-------------1--------------2------------3------------4--------------5------------6
	Sql = "Select Top 5 [News_Title], [News_Author], [News_Date], [News_Memo], [News_Remark], [News_Hits], [News_FileName] From [Monolith_News] Where [News_Apply]=1 And Right([News_Area], 1)=1 "
	Sql	= Sql + " Order By [News_Order] Asc"
	Rs.Open Sql, Conn ,1, 1
	Do While Not(Rs.Bof Or Rs.Eof)
		TempString1	= TempString1 & TempString2
	  TempString1	= Replace(TempString1, "{News_Title}",	Rs(0))
	  TempString1	= Replace(TempString1, "{News_Author}",	Rs(1))
	  TempString1	= Replace(TempString1, "{News_Date}",		Rs(2))
	  TempString1	= Replace(TempString1, "{News_Memo}",		Rs(3))
	  TempString1	= Replace(TempString1, "{News_Post}",		Rs(4))
	  TempString1	= Replace(TempString1, "{News_Hits}",		Rs(5))

		News_Date_Arr = Split(Rs(2), "-")
		News_Url = NewsSettingArr(7) & "/" & News_Date_Arr(0) & "/" & News_Date_Arr(1) & News_Date_Arr(2) & "/" & Rs(6)

	  TempString1	= Replace(TempString1, "{News_Url}",	News_Url)

	  TempString1	= Replace(TempString1, "{}",	"")
	  TempString1	= Replace(TempString1, "{}",	"")
	  TempString1	= Replace(TempString1, "{}",	"")
	  TempString1	= Replace(TempString1, "{}",	"")
		Rs.MoveNext
	Loop
	Rs.Close

  'Sql执行次数+1
  SqlQueryNum	= SqlQueryNum + 1
	
	TheString		= Replace(TheString, "{Monolith_Module_02_02}", 	TempString1)
	'-----------------------------------------------------------------------------------------------------------------------

	'-----------------------------------------------------------------------------------------------------------------------
	'模块：广告
	TempString1	= LoadTemplate("03_06_Module_Ad")
	'---------------------------0
	Sql = "Select Top 1 [Advertisement_Content] From [Monolith_Advertisement] Where [Advertisement_Class_ID]=1 And [Advertisement_Active]=1 Order By [Advertisement_Order] ASC"
	Rs.Open Sql, Conn, 1, 1
	if Not (Rs.Bof Or Rs.Eof) then
		TempString1 = Replace(TempString1, "{Advertisement_String}",Rs(0))
	end if
	Rs.Close
  'Sql执行次数+1
  SqlQueryNum	= SqlQueryNum + 1

	TheString		= Replace(TheString, "{Monolith_Module_03_01}", 	TempString1)
	'-----------------------------------------------------------------------------------------------------------------------

	'-----------------------------------------------------------------------------------------------------------------------
	'模块：投票
	TempString1	= LoadTemplate("03_07_Module_Vote")
	TempString2	= ""
	'------------------0-------------1---------------2---------------3---------------4---------------5---------------6---------------7---------------8---------------9
	Sql = "Select [Vote_Title], [Vote_Answer1], [Vote_Answer2], [Vote_Answer3], [Vote_Answer4], [Vote_Answer5], [Vote_Answer6], [Vote_Answer7], [Vote_Answer8], [Vote_ID] From [Monolith_Vote] Where [Vote_Active]=True"
	Rs.Open Sql, Conn ,1, 1
	if Not(Rs.Bof Or Rs.Eof) then
		for i=1 to 8
			if Trim(Rs(i)) <> "" then
				TempString2 = TempString2 & "<tr>" & chr(13)
				TempString2 = TempString2 & "<td class='row1' align='left'>&nbsp;&nbsp;<input id='Vote_" & i & "' type='radio' value='" & i & "' name='Monolith_Vote'><label for='Vote_" & i & "'>&nbsp;&nbsp;" & Rs(i) & "</label></td>" & chr(13)
				TempString2 = TempString2 & "</tr>" & chr(13)
			end if
		next
		TempString1	= Replace(TempString1, "{Vote_Title}", 	Rs(0))
		TempString1	= Replace(TempString1, "{Vote_ID}", 		Rs(9))
	end if
	Rs.Close
			
  'Sql执行次数+1
  SqlQueryNum	= SqlQueryNum + 1	

	TempString1	= Replace(TempString1, "{Vote_String}", TempString2)
	TheString		= Replace(TheString, "{Monolith_Module_03_02}", 	TempString1)
	'-----------------------------------------------------------------------------------------------------------------------

	'-----------------------------------------------------------------------------------------------------------------------
	'模块：论坛发帖排行
	TempString1	= LoadTemplate("03_08_Module_Sort")
	TempString1	= Split(TempString1, "<!--Split-->")
	TempString2	= ""

 	'-------------------------0-----------1-------------2
	Sql = "Select Top 5 [Users_ID], [Users_Name], [Users_Subject] From [Monolith_Users] [MyBBS_Board] order by [Users_Subject] Desc"
	Rs.Open Sql, Conn, 1, 1
	Do While Not(Rs.Bof Or Rs.Eof)
		TempString2	= TempString2 & TempString1(1)
		TempString2 = Replace(TempString2, "{Users_ID}", 			Rs(0))
		TempString2 = Replace(TempString2, "{Users_Name}", 		Rs(1))
		TempString2 = Replace(TempString2, "{Users_Subject}", Rs(2))
		Rs.MoveNext
	Loop
	Rs.Close
	
  'Sql执行次数+1
  SqlQueryNum	= SqlQueryNum + 1

	TheString		= Replace(TheString, "{Monolith_Module_03_03}", 	TempString1(0) & TempString2 & TempString1(2))
	'-----------------------------------------------------------------------------------------------------------------------



	'-----------------------------------------------------------------------------------------------------------------------
	'模块：日历
	Dim Calendar_String
	
	Dim Date_Time, R_Year, R_Month, R_Day, i, j, Today
	Date_Time=trim(request("Date_Time"))
	if Date_Time = "" then Date_Time = Date()
	R_Year		= Year(Date_Time)
	R_Month		= Month(Date_Time)
	R_Day			= Day(Date_Time)
	
	Calendar_String		= "<tr>" & chr(13)
	Calendar_String		= Calendar_String	& "<td class='Board_Sub' align='center' width='100%' colspan='7'>" & Date() & "</td>" & chr(13)
	Calendar_String		= Calendar_String	& "</tr>" & chr(13)
	Calendar_String		= Calendar_String	& "  <td class='Board_Sub' align='center' width='14%'>S</td>" & chr(13)
	Calendar_String		= Calendar_String	& "  <td class='Board_Sub' align='center' width='14%'>M</td>" & chr(13)
	Calendar_String		= Calendar_String	& "  <td class='Board_Sub' align='center' width='14%'>T</td>" & chr(13)
	Calendar_String		= Calendar_String	& "  <td class='Board_Sub' align='center' width='14%'>W</td>" & chr(13)
	Calendar_String		= Calendar_String	& "  <td class='Board_Sub' align='center' width='14%'>T</td>" & chr(13)
	Calendar_String		= Calendar_String	& "  <td class='Board_Sub' align='center' width='14%'>F</td>" & chr(13)
	Calendar_String		= Calendar_String	& "  <td class='Board_Sub' align='center' width='14%'>S</td>" & chr(13)
	Calendar_String		= Calendar_String	& "</tr>" & chr(13)
	Calendar_String		= Calendar_String	& "<tr>" & chr(13)
	
	j = WeekDay(CDate(R_Year & "-" & R_Month & "-1"))
	j	= j - 1
	
	if j<> 0 then
		For i=1 to j
			Calendar_String		= Calendar_String	& "<td class='row2'>&nbsp;</td>" & chr(13)
		Next
	end if

	i = j
	Today=1

	Do While Today < 32
		if IsDate(R_Year & "-" & R_Month & "-" & Today) then
			Calendar_String		= Calendar_String	& "<td class='row2' align='center' onmouseover=this.className='row1' onmouseout =this.className='row2'>" & Today & "</td>" & chr(13)
			i=i+1
		end if
		if i mod 7=0 then Calendar_String		= Calendar_String	& "</tr>" & chr(13) & "<tr align=center>" & chr(13)
		Today=Today+1
	Loop

	j = i mod 7
	if j <> 0 then j = 7 - j

'	Response.write "<script>alert('"&j&"')</script>"

	if j <> 0 then
		For i=1 to j
			Calendar_String		= Calendar_String	& "<td class='row2'>&nbsp;</td>" & chr(13)
		Next
	end if

	Calendar_String = Calendar_String & "</tr>"


	TempString1	= LoadTemplate("03_09_Module_Calendar")
  TempString1	= Replace(TempString1, "{Calendar}",	Calendar_String)
	TheString		= Replace(TheString, "{Monolith_Module_03_04}", 	TempString1)
	'-----------------------------------------------------------------------------------------------------------------------

	'-----------------------------------------------------------------------------------------------------------------------
	'未用到的模块
	TheString		= Replace(TheString, "{Monolith_Module_01_05}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_01_06}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_01_07}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_01_08}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_01_09}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_01_10}" , "")

	TheString		= Replace(TheString, "{Monolith_Module_02_03}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_04}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_05}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_06}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_07}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_08}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_09}" , "")
	TheString		= Replace(TheString, "{Monolith_Module_02_10}" , "")

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