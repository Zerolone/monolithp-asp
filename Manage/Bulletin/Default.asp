<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>�ޱ����ĵ�</title>
</head>
<script language="javascript" src="/Js/All.js"></script>
<body>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="4"><b><font color="#FFFFFF">&nbsp;�������</font></b></td>
	</tr>
	  <tr height="20" bgcolor="#FFFFFF">
    <td width="45%"  bgcolor="#999999"><font color="#FFFFFF">&nbsp;�������1</font></td>
    <td width="37%"  bgcolor="#999999"><font color="#FFFFFF">&nbsp;�������2</font></td>
    <td width="8%" bgcolor="#999999"><font color="#FFFFFF">&nbsp;��������</font></td>
    <td width="8%"  bgcolor="#999999"><font color="#FFFFFF">&nbsp;����</font></td>
  </tr>
	<%
		Dim ThisPage , Page, Pages, j
	  ThisPage							= Request.ServerVariables("PATH_INFO")
		Page									= Request("Page")
		Dim i, TrBgColor
		'--------------------0---------------1-----------------2------------------3
		Sql = "Select [Bulletin_ID], [Bulletin_Title], [Bulletin_Title2], [Bulletin_Date] From [Monolith_Bulletin] Order By [Bulletin_ID] Desc"
		Rs.Open Sql, Conn, 1, 1
		if Not (Rs.bof Or Rs.Eof) then
			if Page = "" Or Page = 0 then Page = 1
			Page				= CLng(Page)
			Rs.Pagesize	= 19
			Pages				= Rs.Pagecount
			if Page > Pages then Page=Pages
			rs.AbsolutePage=Page
			for j=1 to rs.Pagesize 
				if j mod 2 = 0 then 
					TrBgColor	= ""
				else 
					TrBgColor	= "bgcolor=""#FFFFFF"""
				end if
		%>
  <tr  <%=TrBgColor%> height="22">
    <td width="45%">&nbsp;<%=Rs(1)%></td>
    <td>&nbsp;<%=Rs(2)%></td>
    <td>&nbsp;<%=Rs(3)%></td>
    <td>��<a href="Edit.asp?Bulletin_ID=<%=Rs(0)%>&Page=<%=Page%>">�޸�</a>��<a href="Del.asp?Bulletin_ID=<%=Rs(0)%>&Page=<%=Page%>">ɾ��</a>��</td>
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

		if j < Rs.Pagesize then
		  for i=1 to Rs.Pagesize-j
				if i mod 2 = ColorMod then 
					TrBgColor	= ""
				else 
					TrBgColor	= "bgcolor=""#FFFFFF"""
				end if
	%>
	<tr <%=trbgcolor%> height="22">
		<td width="45%">��</td>
		<td width="37%">��</td>
		<td width="8%">��</td>
		<td width="8%">��</td>
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
	<tr <%=trbgcolor%> height="30" valign="bottom">
		<form method="Post" action="<%=ThisPage%>" style="MARGIN-BOTTOM:0px">
			<td colspan="4" align="center">
			<%
			  ThisPage = ThisPage & "?1=1"
			  if Page<2 then
			%>
				��&nbsp;<strike>��ҳ</strike>&nbsp;��
				<strike>��һҳ</strike>&nbsp;��
			<%else%>
				��&nbsp;<a href="<%=ThisPage%>&Page=1">��ҳ</a>&nbsp;��
				<a href="<%=ThisPage%>&Page=<%=Page-1%>">��һҳ</a>&nbsp;��
			<%
				end if
				if rs.Pagecount-Page<1 then
			%>
				<strike>��һҳ</strike>
				��
				<strike>βҳ</strike>
				��
			<%else%>
				<a href="<%=ThisPage%>&Page=<%=Page+1%>">��һҳ</a>&nbsp;��
				<a href="<%=ThisPage%>&Page=<%=rs.Pagecount%>">βҳ</a>
				��
			<%end if%>
			ҳ�Σ�<strong><font color=red><%=Page%></font>/<%=rs.Pagecount%></strong>&nbsp;ҳ&nbsp;��
			��<b><font color="#FF0000"><%=rs.recordcount%></font></b>����¼&nbsp;��
			<b><%=rs.Pagesize%></b>����¼/ҳ&nbsp;��
			ת����<input type="text" name="Page" size=4 maxlength=10 class="InputBox" value="<%=Page%>">
			<input class="InputBox" type="submit"  value=" Goto (Alt + G) "  name="cndok" accesskey="G">
		<%
	end if
	Call CloseRs
	Call CloseConn
%> </td>
		</form>
	</tr>
</table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->