<%
	Dim User_Agent
	User_Agent	= Request.ServerVariables("Http_User_Agent")
	
'	User_Agent	= "Monolith Client"
  
  if User_Agent <> "Monolith Client" then
%>
<script language="javascript">
	window.open("/Manage/Login_Fail.asp","_parent");
</script>
<%end if%>	
