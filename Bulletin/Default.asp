<!-- #include virtual =	"/Inc/Monolith_Conn.asp" -->
<!-- #include virtual = "/Inc/Monolith_Function.asp" -->
<%
	'---ÍøÕ¾×Ö·û´®----ÁÙÊ±×Ö·û´®1---ÁÙÊ±×Ö·û´®1
	Dim TheString,    TempString1,  TempString2 

	TheString		= ""
	'µ÷Ê×Ò³
	TheString		= TheString & LoadTemplate("04_02_Bulletin")

	'µ÷Í·
	TheString		= Replace(TheString, "{Monolith_Head}", LoadTemplate("02_02_Default_Head"))

	'Css
	TheString		= Replace(TheString, "{Monolith_Css}", 	LoadTemplate("01_01_Css"))

	'Foot
	TempString1	= LoadTemplate("01_02_Foot")
  TempString1	= Replace(TempString1, "{Now}",	Now())
  TempString1	= Replace(TempString1, "{ExecuteTime}",	FormatNumber((Timer()-Startime)*1000,5) & "ºÁÃë¡£")
  TempString1	= Replace(TempString1, "{SqlQueryNum}",	SqlQueryNum)
	
	TheString		= Replace(TheString, "{Monolith_Foot}", 	TempString1)

	'·Ö¸ôÏß
	TheString		= Replace(TheString, "{Monolith_Sep}", 	LoadTemplate("01_03_Sep"))

	'------------------
	TempString1	= Split(TheString, "<!--Split-->")
	TempString2	= ""

	'----------------------0-----------------1-------------------2
	Sql = "Select [Bulletin_Title], [Bulletin_Content], [Bulletin_Date] From [Monolith_Bulletin] Order By [Bulletin_ID] Desc"
	Rs.Open Sql, Conn, 1, 1
	Do While Not (Rs.Bof Or Rs.Eof)
		TempString2	= TempString2 & TempString1(1)
		TempString2	= Replace(TempString2, "{Bulletin_Title}" , 	Rs(0))
		TempString2	= Replace(TempString2, "{Bulletin_Content}" , Replace(Rs(1), chr(10) , "<br>"))
		TempString2	= Replace(TempString2, "{Bulletin_Date}" , 		Rs(2))
		Rs.MoveNext
	Loop
	Rs.Close
	'------------------

	TheString		= TempString1(0) & TempString2 & TempString1(2)

	Response.write TheString

	Call CloseConn
%>