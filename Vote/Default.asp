<!-- #include virtual =	"/Inc/Monolith_Conn.asp" -->
<!-- #include virtual = "/Inc/Monolith_Function.asp" -->
<%
	Dim Bulletin_ID, i, Vote_Count, Vote_Count_Percent
	Bulletin_ID = Request("Bulletin_ID")
	CheckNum Bulletin_ID

	'---网站字符串----临时字符串1---临时字符串1
	Dim TheString,    TempString1,  TempString2 

	TheString		= ""
	'调首页
	TheString		= TheString & LoadTemplate("04_05_Vote")

	'调头
	TheString		= Replace(TheString, "{Monolith_Head}", LoadTemplate("02_02_Default_Head"))

	'Css
	TheString		= Replace(TheString, "{Monolith_Css}", 	LoadTemplate("01_01_Css"))

	'Foot
	TempString1	= LoadTemplate("01_02_Foot")
  TempString1	= Replace(TempString1, "{Now}",	Now())
  TempString1	= Replace(TempString1, "{ExecuteTime}",	FormatNumber((Timer()-Startime)*1000,5) & "毫秒。")
  TempString1	= Replace(TempString1, "{SqlQueryNum}",	SqlQueryNum)
	
	TheString		= Replace(TheString, "{Monolith_Foot}", 	TempString1)

	'分隔线
	TheString		= Replace(TheString, "{Monolith_Sep}", 	LoadTemplate("01_03_Sep"))
		
	Sql	= "Select Sum(Vote_Count1 + Vote_Count2 + Vote_Count3 + Vote_Count4 + Vote_Count5 + Vote_Count6 + Vote_Count7 + Vote_Count8) As Vote_Count From [Monolith_Vote] Where [Vote_Active] = True"
	Rs.Open Sql, Conn, 1, 1
	Vote_Count	= Rs(0)
	Rs.Close
	
	'------------------0-------------1---------------2---------------3---------------4---------------5---------------6---------------7---------------8
	Sql = "Select [Vote_Title], [Vote_Answer1], [Vote_Answer2], [Vote_Answer3], [Vote_Answer4], [Vote_Answer5], [Vote_Answer6], [Vote_Answer7], [Vote_Answer8],"
	'------------------9-------------10------------11------------12------------13------------14------------15------------16
	Sql	= Sql & " [Vote_Count1],[Vote_Count2],[Vote_Count3],[Vote_Count4],[Vote_Count5],[Vote_Count6],[Vote_Count7],[Vote_Count8] From [Monolith_Vote] Where [Vote_Active]=True"

	Rs.Open Sql, Conn ,1, 1
	if Not(Rs.Bof Or Rs.Eof) then
		for i=1 to 8
			if Trim(Rs(i)) <> "" then
				Vote_Count_Percent	= Rs(i+8) / Vote_Count * 100
				
				TempString2 = TempString2 & "<tr>" & chr(13)
				TempString2 = TempString2 & "<td class='row1' align='left'>&nbsp;&nbsp;" & Rs(i) & "</td>"
				TempString2	= TempString2 & "<td class='row2'>&nbsp;" & Rs(i+8) & "票， 占总数的：" & Vote_Count_Percent & "%"
				TempString2	= TempString2 & "<td class='row1'>&nbsp<img src=/images/bar.gif width='" & Vote_Count_Percent * 5 &"' height='10'></td>" & chr(13)
				TempString2 = TempString2 & "</tr>" & chr(13)
			end if
		next
		TheString	= Replace(TheString, "{Vote_Title}", 	Rs(0))
	end if
	
	TheString		= Replace(TheString, "{Vote_Content}", TempString2)
	Response.write TheString

	Call CloseRs
	Call CloseConn
%>