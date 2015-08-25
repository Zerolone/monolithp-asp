<!-- #include file="Head.asp"-->
<%
  Dim TheString

  Dim ThisPage, FormPage
  ThisPage = "ListSave.asp"
  FormPage	= Request.ServerVariables("HTTP_REFERER")
%>
<!-- #include file="BoardString.asp"-->
<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线-->
<%
	TheString	= MyBBS_Template_04_Body
	TheString	= Replace(TheString, "{ThisPage}",			ThisPage)
	TheString	= Replace(TheString, "{FormPage}",			FormPage)
	TheString	= Replace(TheString, "{Board_ID}",			Board_ID)
	TheString	= Replace(TheString, "{Board_Name}",		Board_Name)
	Response.write TheString

  TheString	= MyBBS_Template_Foot 
  TheString	= Replace(TheString, "{Now}",	Now())
  TheString	= Replace(TheString, "{ExecuteTime}",	FormatNumber((Timer()-Startime)*1000,5) & "毫秒。")
  TheString	= Replace(TheString, "{SqlQueryNum}",	SqlQueryNum)
  Response.write TheString
%>