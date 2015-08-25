<!-- #include virtual =	"/Inc/Monolith_Conn.asp" -->
<!-- #include virtual = "/Inc/Monolith_Function.asp" -->
<%
	'---ÍøÕ¾×Ö·û´®----ÁÙÊ±×Ö·û´®1---ÁÙÊ±×Ö·û´®1
	Dim TheString,    TempString1,  TempString2 

	TheString		= ""
	'µ÷Ê×Ò³
	TheString		= TheString & LoadTemplate("05_01_Login")

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

	Response.write TheString

	Call CloseConn
%>