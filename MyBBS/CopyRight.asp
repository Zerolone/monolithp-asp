<%
	Dim EndTime

  Endtime = Timer()
  TheString = FormatNumber((Endtime-Startime)*1000,5) & "���롣"
%>
<!--�ָ���--><img src="images/Blank.gif" border="0" width="1" height="5"><!--�ָ���--><table bgcolor="#8394B2" cellspacing="0" width="100%" height="26" cellpadding="0">
	<tr>
		<td>
		<table cellspacing="0" cellpadding="0" width="100%">
			<tr>
				<td align="right" width="50%"><a href="#"><b>
				<font color="#FFFFFF">�򻯰汾</font></b></a></td>
				<td align="right" width="50%">������ʱ��: <%=Now%>&nbsp;&nbsp;</td>
			</tr>
		</table>
		</td>
	</tr>
</table>
<!--�ָ���--><img src="images/Blank.gif" border="0" width="1" height="5"><!--�ָ���--><table bgcolor="#EEEEEE" cellspacing="0" width="100%" height="26" cellpadding="0">
	<tr>
		<td align="center">�� Powered by Monolith - MyBBs<br>[ Script Execution time: <%=TheString%> ]   [ <%=SqlQueryNum%> queries used ] </td>
	</tr>
</table>