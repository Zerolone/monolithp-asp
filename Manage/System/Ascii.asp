<!-- #include virtual ="/Manage/Check.asp"--><link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<br>
<br>
<br>
<br>
<table align="center">
	<tr>
	<%
		For i=1 to 128
			Response.write "  <td height=22 width=90 onMouseOver=this.className='tdin' onMouseOut =this.className=''>chr(" & i & ")<font color=""#FF0000"">" & chr(i) & "</font></td>"
			if i mod 8 =0 then Response.write "</tr><tr>"
		Next
	%>
	</tr>
</table>