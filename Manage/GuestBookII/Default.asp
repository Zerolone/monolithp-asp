<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>�ޱ����ĵ�</title>
</head>

<script language="javascript" src="/Js/All.js"></script>
<body>

<script language="JavaScript">
<!--
function CheckAll(form)
{
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
}
//-->
</script>
<form name="Prodlist" action="Del.asp" method="post">
</form>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="8"><b><font color="#FFFFFF">&nbsp;���԰����</font></b></td>
	</tr>
	<tr height="20">
		<td width="35" bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;����</nobr></font></td>
		<td bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;��������</nobr></font></td>
		<td bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;��������</nobr></font></td>
		<td bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;�ظ�����</nobr></font></td>
		<td bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;�ظ�����</nobr></font></td>
		<td bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;�Ƿ񼤻�</nobr></font></td>
		<td bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;����</nobr></font></td>
	</tr>
	<%
		Dim GuestBook_Content, GuestBook_Replay, GuestBook_UserName
	
		Dim Page, Pages, j, ThisPage
	  ThisPage							= Request.ServerVariables("PATH_INFO")
		Page									= clng(request("Page"))
		Dim i, TrBgColor

		'-----------------------0---------------1---------------------2---------------------3-------------------4-----------------------5
		Sql = "Select [GuestBook_ID], [GuestBook_UserName], [GuestBook_Content], [GuestBook_Postdate], [GuestBook_Replay], [GuestBook_ReplayDate], [GuestBook_Active] From [Monolith_GuestBookII] order by [GuestBook_PostDate] desc"
		Rs.open Sql, Conn, 1, 1
		if Not (Rs.bof Or Rs.Eof) then
			Rs.PageSize=19
			if page=0 then page=1
			pages=rs.pagecount
			if page > pages then page=pages
			rs.AbsolutePage=page
			for j=1 to rs.PageSize 

				if j mod 2 = 0 then 
					TrBgColor	= ""
				else 
					TrBgColor	= "bgcolor=""#FFFFFF"""
				end if
%>
	<tr <%=trbgcolor%> height="21">
		<td> &nbsp;<%=Rs(1)%></td>
		<td> &nbsp;<%=Left(Rs(2),20)%></td>
		<td>&nbsp;<%=Rs(3)%>&nbsp; </td>
		<td>&nbsp;<%=Left(Rs(4),20)%></td>
		<td><%=Rs(5)%></td>
		<td><%=Rs(6)%></td>
		<td>��<a href="Reply.asp?GuestBook_ID=<%=Rs(0)%>" target="tabWin" onclick="fnClick();">��˻ظ�</a>��<a href="Del.asp?GuestBook_ID=<%=Rs(0)%>&Page=<%=Page%>">ɾ��</a>���� </td>
	</tr>
	<%
    Rs.MoveNext
    if rs.eof then exit for
  next

		Dim ColorMod
		
		if TrBgColor="" then
			ColorMod=0
		else
			ColorMod=1
		end if	

		if j < Rs.PageSize then
		  for i=1 to Rs.PageSize-j
				if i mod 2 = ColorMod then 
					TrBgColor	= ""
				else 
					TrBgColor	= "bgcolor=""#FFFFFF"""
				end if
	%>
	<tr <%=trbgcolor%> height="22">
		<td>��</td>
		<td>��</td>
		<td>��</td>
		<td>��</td>
		<td>��</td>
		<td>��</td>
		<td>��</td>
	</tr>
	<%	
			next
		end if
		if TrBgColor="" then
			TrBgColor	= "bgcolor=""#FFFFFF"""
		else
			TrBgColor = ""
		end if
	%>
	<form>
	</form>
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
	<%
	end if
	Call CloseRs
	Call CloseConn
%>
	</tr>
</table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->
