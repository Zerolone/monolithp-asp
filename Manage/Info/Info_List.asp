<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual = "/Manage/News/NewsSetting.asp" --><!-- #include virtual ="/Manage/Check.asp"-->

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>�ޱ����ĵ�</title>
<script language="javascript" src="/Js/All.js"></script>
<script language="JavaScript">
function cform(){
 if(!confirm("�Ƿ�Ҫɾ�����ţ�"))
 return false;
}
</script>
</head>

<body>

<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="5"><b><font color="#FFFFFF">&nbsp;��Ϣ���� &gt;&gt; ��Ϣ�쿴</font></b></td>
	</tr>
	<tr height="20" bgcolor="#FFFFFF">
		<td width="*" bgcolor="#999999"><font color="#FFFFFF">&nbsp;��Ϣ����</font></td>
		<td width="1" bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;��Ŀ���</nobr></font></td>
		<td width="1" bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;��Ŀ����</nobr></font></td>
		<td width="1" bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;����</nobr></font></td>
		<td width="1" bgcolor="#999999"><font color="#FFFFFF"><nobr>&nbsp;����</nobr></font></td>
	</tr>
	<%
		Dim Info_Class, Info_Class_Len, Info_Date_Arr, Info_Url, Info_Area, Info_Area_Len, Info_Area_Code, i
		Info_Class						= Request("Info_Class")
		Info_Area							= Request("Info_Area")
		Info_Class_Len				= Len(Info_Class)		
		
		if Info_Area <> "" then
			Info_Area_Len					= Len(Info_Area)
			Info_Area_Code					= 1
			For i=2 to Info_Area_Len
				Info_Area_Code	= 10 * Info_Area_Code
			Next
		end if
		
		Dim ThisPage , Page, Pages, j
	  ThisPage							= Request.ServerVariables("PATH_INFO")
		Page									= clng(request("Page"))
		Dim TrBgColor
		'-----------------0------------1-------------2------------3------------------4--------------5----------------6
		Sql = "Select [Info_ID], [Info_Title], [Info_Class], [Info_Class_Name], [Info_Author], [Info_FileName], [Info_Date] From [Monolith_Info] Where [Info_Apply]=1 "
		if Info_Class	<> "" then Sql = Sql + " And Left([Info_Class]," & Info_Class_Len & ") ='" & Info_Class & "' "
		if Info_Area	<> "" then Sql = Sql + " And Right([Info_Area], " & Info_Area_Len & ") >= " & Info_Area_Code

		if Info_Area	<> "" then
			Sql = Sql + " Order By [Info_Order] Asc"
		else
			Sql	= Sql + " Order By [Info_ID] Desc"
		end if
		
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

				Info_Date_Arr = Split(Rs(6), "-")
				Info_Url = NewsSettingArr(7) & "/" & Info_Date_Arr(0) & "/" & Info_Date_Arr(1) & Info_Date_Arr(2) & "/" & Rs(5)
				
		%>
	<tr <%=trbgcolor%>>
		<td height="22">&nbsp;<a href="<%=Info_Url%>" title="<%=Rs(1)%>" target="_blank"><%=CutLongStrDot(Rs(1), 60, 1)%></a></td>
		<td height="22">&nbsp;<%=Rs(2)%></td>
		<td height="22">&nbsp;<%=Rs(3)%></td>
		<td height="22">&nbsp;<%=Rs(4)%></td>
		<td>��<a href="Add.asp?Info_ID=<%=Rs(0)%>" target="tabWin" onclick="fnClick();">�޸�</a>��<a onclick="return cform();" href="Info_Del.asp?Info_ID=<%=Rs(0)%>&Page=<%=Page%>&Info_Class=<%=Info_Class%>">ɾ��</a>��<a href="Info_ToHtml.asp?Info_ID=<%=Rs(0)%>&Page=<%=Page%>" target="tabWin" onclick="fnClick();">HTML</a>��</td>
	</tr>
	<%
			Rs.MoveNext
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
	<tr <%=trbgcolor%> height="30" valign="bottom">
		<form method="Post" action="<%=ThisPage%>" style="MARGIN-BOTTOM:0px">
			<input type="hidden" name="Info_Class" value="<%=Info_Class%>">
			<input type="hidden" name="Info_Area" value="<%=Info_Area%>">
			<td colspan="5" align="center">
			<%
			  ThisPage = ThisPage & "?1=1"
			  if Info_Class	<> "" then ThisPage = ThisPage + "&Info_Class=" + Info_Class
			  if Info_Area	<> "" then ThisPage = ThisPage + "&Info_Area="  + Info_Area
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