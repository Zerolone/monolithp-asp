<%
	Dim User_Agent
	User_Agent	= Request.ServerVariables("Http_User_Agent")
	Response.Write User_Agent
	
%>