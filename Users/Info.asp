<!-- #include virtual="/Head.asp"-->
<%
  Dim TheString

 	TheString =	Monolith_Template_09
	Response.write TheString	

  TheString	= Monolith_Template_Foot 
  TheString	= Replace(TheString, "{Now}",	Now())
  TheString	= Replace(TheString, "{ExecuteTime}",	FormatNumber((Timer()-Startime)*1000,5) & "ºÁÃë¡£")
  TheString	= Replace(TheString, "{SqlQueryNum}",	SqlQueryNum)
  Response.write TheString
%>