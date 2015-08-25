<!-- #include virtual =	"/Inc/Monolith_Conn.asp" -->
<!-- #include virtual = "/Inc/Monolith_Function.asp" -->
<%
	'---ÍøÕ¾×Ö·û´®----ÁÙÊ±×Ö·û´®1---ÁÙÊ±×Ö·û´®1
	Dim TheString,    TempString1,  TempString2 

	TheString		= ""
	'µ÷Ê×Ò³
	TheString		= TheString & LoadTemplate("04_01_Calendar")

	'µ÷Í·
	TheString		= Replace(TheString, "{Monolith_Head}", LoadTemplate("02_02_Default_Head"))

	'Css
	TheString		= Replace(TheString, "{Monolith_Css}", 	LoadTemplate("01_01_Css"))

	'--------
	Dim Calendar_String

	Dim Date_Time, R_Year, R_Month, R_Day, i, j, Today
	Date_Time=trim(request("Date_Time"))
	if Date_Time = "" then Date_Time = Date()
	R_Year		= Year(Date_Time)
	R_Month		= Month(Date_Time)
	R_Day			= Day(Date_Time)

	TheString = Replace(TheString, "{Calendar_LastMonth}", 	DateAdd("m",-1,Date_Time))
	TheString = Replace(TheString, "{Calendar_Today}", 			formatdatetime(Date_Time,1))
	TheString = Replace(TheString, "{Calendar_NextMonth}", 	DateAdd("m",1,Date_Time))

	Calendar_String	= "<tr align=center>" & Chr(13)
	
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

	TheString = Replace(TheString, "{Calendar_String}", 	Calendar_String)
	'----------------

	'Foot
	TempString1	= LoadTemplate("01_02_Foot")
  TempString1	= Replace(TempString1, "{Now}",	Now())
  TempString1	= Replace(TempString1, "{ExecuteTime}",	FormatNumber((Timer()-Startime)*1000,5) & "ºÁÃë¡£")
  TempString1	= Replace(TempString1, "{SqlQueryNum}",	SqlQueryNum)
	
	TheString		= Replace(TheString, "{Monolith_Foot}", 	TempString1)

	'·Ö¸ôÏß
	TheString		= Replace(TheString, "{Monolith_Sep}", 	LoadTemplate("01_03_Sep"))
	Response.write TheString

	Call CloseConn
%>