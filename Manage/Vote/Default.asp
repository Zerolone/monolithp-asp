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
		<td colspan="3"><b><font color="#FFFFFF">&nbsp;�������</font></b></td>
	</tr>  
	<tr height="20">
    <td bgcolor="#999999"><font color="#FFFFFF">&nbsp;�������</font></td>
    <td bgcolor="#999999"><font color="#FFFFFF">&nbsp;�Ƿ�����</font></td>
    <td width="100" bgcolor="#999999"><font color="#FFFFFF">&nbsp;����</font></td>
  </tr>
	<%
		Dim Page, Pages, j, ThisPage
	  ThisPage							= Request.ServerVariables("PATH_INFO")
		Page									= clng(request("Page"))
		Dim i, TrBgColor
		'------------------0----------1-------------2
		Sql = "SELECT [Vote_ID], [Vote_Title], [Vote_Active] FROM [Monolith_Vote] Order By [Vote_Active] ASC "
		Rs.Open Sql, Conn ,1, 1
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
  <tr  <%=TrBgColor%> height="22">
    <td>&nbsp;<%=Rs(1)%></td>
    <td>&nbsp;<%=Rs(2)%></td>
    <td>��<a href="Edit.asp?Vote_ID=<%=Rs(0)%>&Page=<%=Page%>">�޸�</a>��<a href="Del.asp?Vote_ID=<%=Rs(0)%>&Page=<%=Page%>">ɾ��</a>��</td>
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
		<td width="18%" >��</td>
		<td >��</td>
		<td >��</td>
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
			<td colspan="3" align="center">
			<%
			  ThisPage = ThisPage & "?1=1"
			  if Page<2 then
			%>
				��&nbsp;<strike>��ҳ</strike>&nbsp;��
				<strike>��һҳ</strike>&nbsp;��
			<%else%>
				��&nbsp;<a href="<%=ThisPage%>&page=1">��ҳ</a>&nbsp;��
				<a href="<%=ThisPage%>&page=<%=Page-1%>">��һҳ</a>&nbsp;��
			<%
				end if
				if rs.pagecount-page<1 then
			%>
				<strike>��һҳ</strike>
				��
				<strike>βҳ</strike>
				��
			<%else%>
				<a href="<%=ThisPage%>&page=<%=page+1%>">��һҳ</a>&nbsp;��
				<a href="<%=ThisPage%>&page=<%=rs.pagecount%>">βҳ</a>
				��
			<%end if%>
			ҳ�Σ�<strong><font color=red><%=Page%></font>/<%=rs.pagecount%></strong>&nbsp;ҳ&nbsp;��
			��<b><font color="#FF0000"><%=rs.recordcount%></font></b>����¼&nbsp;��
			<b><%=rs.pagesize%></b>����¼/ҳ&nbsp;��
			ת����<input type="text" name="page" size=4 maxlength=10 class="InputBox" value="<%=page%>">
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