<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" --><!-- #include virtual ="/Manage/Check.asp"--><%
	Dim Info_Class
	Info_Class						= Request("Info_Class")
	Dim ThisPage , Page, Pages, j
  ThisPage							= Request.ServerVariables("PATH_INFO")
	Page									= clng(request("Page"))
	Dim i, TrBgColor
	
	Dim Info_ID
	Info_ID	= Request("Info_ID")
'	Response.write "<hr>" & Info_ID & "<hr>"
	if Info_ID <> "" and Request("submit") = " ͨ �� �� �� (Alt + Y) " then
	  Sql="update [Monolith_Info] set [Info_Apply]=1 where [Info_ID] in (" & Info_ID & ")"
	  Conn.Execute(Sql)
'	  Response.write "<script language='javascript'>alert('����ͨ����ˣ�');location.href=""" & ThisPage & """;</script>"
	end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>�ޱ����ĵ�</title>
<script language="javascript" src="/Js/All.js"></script>
<script language="JavaScript">
function CheckBoxSelect(CheckBoxName, Code)
{
	switch (Code){
	case 1:
	  for(i=0; i<= CheckBoxName.length-1;i++)
	  {
	  	CheckBoxName(i).checked=true;
	  }
	  break;

	case 2:
	  for(i=0; i<= CheckBoxName.length-1;i++)
	  {
	  	CheckBoxName(i).checked=CheckBoxName(i).checked==true?false:true;
	  }
	  break;

	case 3:
	  for(i=0; i<= CheckBoxName.length-1;i++)
	  {
	  	CheckBoxName(i).checked=false;
	  }
	  break;

	 }
}
</script>
</head>

<body>

<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="4"><b><font color="#FFFFFF">&nbsp;��Ϣ���� &gt;&gt; ��Ϣ���</font></b></td>
	</tr>
	<tr height="20" bgcolor="#FFFFFF">
		<td width="*" bgcolor="#999999"><font color="#FFFFFF">&nbsp;��Ϣ����</font></td>
		<td width="1" bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;��Ŀ���</nobr></font></td>
		<td width="1" bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;��Ŀ����</nobr></font></td>
		<td width="1" bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;��������</nobr></font></td>
	</tr>
	<%
		'-----------------0------------1-------------2------------3------------------4
		Sql = "Select [Info_ID], [Info_Title], [Info_Class], [Info_Class_Name], [Info_Author] From [Monolith_Info] Where [Info_Apply]=0 "
		if Info_Class	<> "" then Sql = Sql + " And [Info_Class] Like '" & Info_Class & "%' "
		Sql	= Sql + " Order By [Info_ID] Desc"
		
		Rs.Open Sql, Conn ,1, 1
		if Not (Rs.bof Or Rs.Eof) then
			Rs.PageSize=17
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
	<form action="<%=ThisPage%>" method="POST" name="form2" style="MARGIN-BOTTOM:0px">
		<tr <%=trbgcolor%>>
			<td height="22">
			<input type="checkbox" name="Info_ID" value="<%=Rs(0)%>" id="Info_ID_<%=Rs(0)%>"><label for="Info_ID_<%=Rs(0)%>">
			<%=Rs(1)%></label></td>
			<td><%=Rs(2)%></td>
			<td><%=Rs(3)%></td>
			<td><%=Rs(4)%></td>
		</tr>
		<%
			Rs.MoveNext
			i = i +1
    	if rs.eof then exit for
		next
		if j < Rs.PageSize then
		  for i=1 to Rs.PageSize-j
				if i mod 2 = 0 then 
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
		<tr <%=trbgcolor%>>
			<td colspan="4" height="44">
			<p align="center">
			<input type="button" value=" ȫ ѡ (Alt + A) " onclick="CheckBoxSelect(Info_ID, 1)" class="InputBox" accesskey="A">
			<input type="button" value=" �� ѡ (Alt + I) " onclick="CheckBoxSelect(Info_ID, 2)" class="InputBox" accesskey="I">
			<input type="button" value=" ȫ �� ѡ (Alt + N) " onclick="CheckBoxSelect(Info_ID, 3)" class="InputBox" accesskey="N">
			<input name="Submit" type="submit" value=" ͨ �� �� �� (Alt + Y) " class="InputBox" accesskey="Y"></p>
			</td>
		</tr>
	</form>
	<%
		if TrBgColor="" then
			TrBgColor	= "bgcolor=""#FFFFFF"""
		else
			TrBgColor = ""
		end if
	%>
	<tr <%=trbgcolor%> height="30" valign="bottom">
		<form method="Post" action="<%=ThisPage%>" style="MARGIN-BOTTOM:0px">
			<input type="hidden" name="Info_Class" value="<%=Info_Class%>">
			<td colspan="4" align="center"><%
			  ThisPage = ThisPage & "?1=1"
			  if Info_Class	<> "" then ThisPage = ThisPage + "&Info_Class=" + Info_Class
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
			<input class="InputBox" type="submit" value=" Goto " name="cndok"> <%
	end if
	Call CloseRs
	Call CloseConn
%> </td>
		</form>
	</tr>
</table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->

</body>

</html>
