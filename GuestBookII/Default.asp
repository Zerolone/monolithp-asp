<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->

<html>

<head>
<title>���԰�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Css/Manage.Css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<%
		Dim GuestBook_Content, GuestBook_Replay, GuestBook_UserName, GuestBook_UserMail

		Dim Page, Pages, j, ThisPage

  ThisPage				= Request.ServerVariables("PATH_INFO")
	page						= clng(request("page"))

	'---------------------0---------------------1---------------------2--------------------3-------------------4-----------------------5
	Sql="Select [GuestBook_UserName], [GuestBook_Content], [GuestBook_Postdate], [GuestBook_Replay], [GuestBook_ReplayDate], [GuestBook_UserMail] From [Monolith_GuestBookII] Where [GuestBook_Active] order by [GuestBook_Postdate] desc"
	Rs.open sql,conn,1,3
	if Request("send")="ok" then
		if Request("GuestBook_UserName")="" or Request("GuestBook_Content")="" then
			response.write "<script language='javascript'>"
			response.write "alert('��Ч���룡��������д������');"
			response.write "location.href='" & ThisPage & "';"			
			response.write "</script>"
		else
			Rs.Addnew
			Rs(0)	= MonolithEncode(Request("GuestBook_UserName"))
			Rs(1)	= MonolithEncode(Request("GuestBook_Content"))
			Rs(5)	= MonolithEncode(Request("GuestBook_EMail"))
			Rs.Update
			Response.Write "<script>alert('��ӳɹ��� �ȴ�����Ա��ˣ�');window.open('"& ThisPage &"', '_self');</script>"
		end if
	end if
%>
<div align="center">
<table border="1" cellpadding="2" cellspacing="0" width="80%">
	<%
	if not (rs.bof or rs.eof) then
    rs.PageSize=5
		if page=0 then page=1
		pages=rs.pagecount
		if page > pages then page=pages
		rs.AbsolutePage=page  
		for j=1 to rs.PageSize 
		%>
	<tr>
		<td>&nbsp;<b><%=Rs(0)%> ������</b>   ���䣺<%=Rs(5)%></td>
		<td>����ʱ�䣺<%=Rs(2)%>&nbsp; </td>
	</tr>
	<tr>
		<td>��</td>
		<td>&nbsp; </td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="#ffffff" style="padding:5pt">
		<span class="newstext"><%=Rs(1)%> <%if Rs(3) <>"" then%>
		</span>
		<table width="100%" border="0" cellpadding="5" cellspacing="2">
			<tr>
				<td><b><font color="#ff0000">�ͷ��ظ���</font></b>&nbsp;
				<font color="#003399"><%=Rs(3)%> [<%=Rs(4)%>]</font></td>
			</tr>
		</table>
		<%end if%> </td>
	</tr>
	

	<%
      	rs.movenext
	  		if rs.eof then exit for
				next
%>
</table>
</div>
<%
end if
%>
<table border="0" width="80%" align="center">
	<form method="Post" action="<%=ThisPage%>" style="MARGIN-BOTTOM:0px">
		<tr>
			<td colspan="8" align="center"><%
			  ThisPage = ThisPage & "?1=1"
			  if Page<2 then
			%> �� <strike>��ҳ</strike> �� <strike>��һҳ</strike> �� <%else%> �� 
			<a href="<%=ThisPage%>&page=1">��ҳ</a> �� 
			<a href="<%=ThisPage%>&page=<%=Page-1%>">��һҳ</a> �� <%
				end if
				if rs.pagecount-page<1 then
			%> <strike>��һҳ</strike> �� <strike>βҳ</strike> �� <%else%>
			<a href="<%=ThisPage%>&page=<%=page+1%>">��һҳ</a> �� 
			<a href="<%=ThisPage%>&page=<%=rs.pagecount%>">βҳ</a> �� <%end if%> ҳ�Σ�<strong><font color="red"><%=Page%></font>/<%=rs.pagecount%></strong> 
			ҳ �� ��<b><font color="#FF0000"><%=rs.recordcount%></font></b>����¼ �� <b>
			<%=rs.pagesize%></b>����¼/ҳ �� ת����<input type="text" name="page" size="4" maxlength="10" class="InputBox" value="<%=page%>">
			<input class="InputBox" type="submit" value=" Goto (Alt + G) " name="cndok" accesskey="G">
			</td>
		</tr>
	</form>
</table>
<table border="0" cellpadding="3" cellspacing="1" width="80%" align="center" bgcolor="#ffffff">
	<tr>
		<td height="10" colspan="2">��</td>
	</tr>
	<script language="JavaScript">
  function checkForm(){
	if (javaTrim(document.feedback.GuestBook_UserName.value) ==""){
	alert("��û����ȷ��д�û�����");
	document.feedback.GuestBook_GuestBook_UserName.focus();
	document.feedback.GuestBook_UserName.value="";
	return false;
	}
	if (javaTrim(document.feedback.GuestBook_Content.value) ==""){
	alert("����д���ύ�����ݡ�");
	document.feedback.GuestBook_Content.focus();
	document.feedback.GuestBook_Content.value="";
	return false;
	}
	return true;
}
function javaTrim(str) {
   for (var i=0; (str.charAt(i)==' ') && i<str.length; i++);
   if (i == str.length) return ''; //whole string is space
   var newstr = str.substr(i);
   for (var i=newstr.length-1; newstr.charAt(i)==' ' && i>=0; i--);
   newstr = newstr.substr(0,i+1);
   return newstr;
}
//�ж��������ֶ���
function textCounter(maxlimit) {
if (document.feedback.GuestBook_Content.value.length > maxlimit) 
{alert("  ����ֽ̫С�ˣ�д���� "+maxlimit+" ���֣�����һҳ�ɣ�^!^ ");
 document.feedback.GuestBook_Content.value = document.feedback.GuestBook_Content.value.substring(0, maxlimit);
return false; 
 }
return true;
}
</script>
	<form action="<%=ThisPage%>" method="post" onsubmit="return checkForm();" name="feedback">
		<tr>
			<td width="143" align="right">������ </td>
			<td>
			<input name="GuestBook_UserName" class="field" id="GuestBook_UserName" size="18" maxlength="10"> 
			* �����䣺 
			<input name="GuestBook_email" class="field" id="GuestBook_email" size="18" maxlength="30">
			</td>
		</tr>
		<tr>
			<td align="right" valign="top" width="143">���ݣ� </td>
			<td>
			<textarea name="GuestBook_Content" cols="50" rows="4" class="field" id="GuestBook_Content" onkeydown="textCounter(200);"></textarea>
			</td>
		</tr>
		<tr align="center">
			<td colspan="2"><input type="submit" name="Submit" value="�ύ����">
			<input type="hidden" name="send" value="ok"></td>
		</tr>
	</form>
</table>

</body>

</html>
