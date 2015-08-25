<!-- #include virtual = "/Inc/Monolith_Conn.asp" -->
<!-- #include virtual = "/Inc/Monolith_Function.asp" -->
<%
	Dim Vote_ID
	Vote_ID				= Request("Vote_ID")
	CheckNum(Vote_ID)
	
	Dim Monolith_Vote
	Monolith_Vote	=	Request("Monolith_Vote")
	
	Sql	= "Update [Monolith_Vote] Set [Vote_Count" & Monolith_Vote & "]=[Vote_Count"& Monolith_Vote &"]+1 Where [Vote_ID]=" & Vote_ID
	Conn.Execute(Sql)
	Call CloseConn
%>
<script language="javascript">
	alert("提交投票成功！");
</script>