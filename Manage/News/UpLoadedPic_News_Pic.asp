<!-- #include virtual ="/Inc/Monolith_Conn.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<html>

<head>
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<script language=javascript>
function AddTo(PicSrc)
{
	parent.document.form2.UpPicSelectList.value = PicSrc;
	parent.document.form2.InsertUpLoadPic0.click();
}
</script>
</head>

<body>

<%
		Dim Page, Pages, j, ThisPage
		ThisPage       = Request.ServerVariables("PATH_INFO")		
		%>
<div align="center">
<table border="1" width="460" height="220" bordercolor="#C0C0C0">
	<tr>
		<%
			Page         = clng(request("Page"))
			Sql = "Select [UpPic_FileName] From [Monolith_UpPic] Order By [UpPic_ID] DESC"
			Rs.Open Sql, Conn ,1, 1
if Not (Rs.bof Or Rs.Eof) then
  Rs.PageSize=6
  if page=0 then page=1
  pages=rs.pagecount
  if page > pages then page=pages
  rs.AbsolutePage=page
  for j=1 to rs.PageSize 
			%> <td align="center" valign="top" >
		<img src="<%=Rs(0)%>" width="111" height="95" onclick="AddTo(this.src);"></td>
		<%
		Rs.MoveNext
    if rs.eof then exit for
    if j mod 3 = 0 then Response.write "</tr><tr>"
  next
			%> </tr>
</table>
</div>
<table width="460" border="0" cellpadding="0" cellspacing="0">
	<tr bgcolor="#FFFFFF">
		<form method="Post" action="<%=ThisPage%>">
			<td height="22" align="center"><%
   ThisPage = ThisPage & "?1=1"
  if Page<2 then
     response.write "��ҳ ��һҳ&nbsp;"
   else
     response.write "<a href=" & ThisPage & "&page=1>��ҳ</a>&nbsp;"
  response.write "<a href=" & ThisPage & "&page=" & Page-1 & ">��һҳ</a>&nbsp;"
   end if
   if rs.pagecount-page<1 then
     response.write "��һҳ βҳ"
   else
     response.write "<a href=" & ThisPage & "&page=" & (page+1) & ">��һҳ</a> "
  response.write "<a href=" & ThisPage & "&page="&rs.pagecount&">βҳ</a>"
   end if
   response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&Page&"</font>/"&rs.pagecount&"</strong>ҳ "
   response.write "&nbsp;��<b><font color='#FF0000'>"&rs.recordcount&"</font></b>�� <b>"&rs.pagesize&"</b>��/ҳ"
   response.write " ת����<input type='text' name='page' size=4 maxlength=10 class=input value="&page&">"
   response.write " <input class=input type='submit'  value=' Goto '  name='cndok'></span></p>"
 end if
 Rs.Close
 Set Rs= Nothing
%> </td>
		</form>
	</tr>
</table>
</body>

</html>
