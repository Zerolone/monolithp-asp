<%
	Dim IP
  IP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
  if IP = ""  then ip=Request.ServerVariables("REMOTE_ADDR")
  
  if Session("Admin_Login") <> 1 then
%>
<script language="javascript">
	window.open("/Manage/Login.asp","_parent");
</script>
<%end if%>