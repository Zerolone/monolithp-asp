<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Monolith-MyBBS</title>

<style>
<!--
.forminput  { font-size: 11px; font-family: verdana, helvetica, sans-serif; vertical-align: middle }


td { font-family: Tahoma,Verdana, Tahoma, Arial,Helvetica, sans-serif; font-size: 12px; word-break: break-all }
body{ font-family: Tahoma,Verdana, Tahoma, Arial,Helvetica, sans-serif; font-size: 12px; word-break: break-all }

a:link, a:visited, a:active{	background: transparent;	color: #3A4F6C;	text-decoration: none;}
a:hover{ color: #5176B5;}

.InputBox2						{	border: 1px solid #999999	}

.TdTitle {color: #3A4F6C;  font-weight:bold}
.darkrow1 { background-color: #C2CFDF; color:#4C77B6; }
.row4 { background-color: #E4EAF2 }
.row2 { background-color: #DFE6EF }
.desc { font-size:10px; color:#434951 }
.darkrow2 { background-color: #BCD0ED; color:#3A4F6C; }

.pformstrip { background-color: #D1DCEB; color:#3A4F6C;font-weight:bold;padding:7px;margin-top:1px }
.forminput  { font-size: 11px; font-family: verdana, helvetica, sans-serif; vertical-align: middle }
.pformleft  { background-color: #F5F9FD; padding:6px; margin-top:1px;width:25%; border-top:1px solid #C2CFDF; border-right:1px solid #C2CFDF; }

.pformright { background-color: #F5F9FD; padding:6px; margin-top:1px;border-top:1px solid #C2CFDF; }


-->
</style>
</head>

<body>

<table border="1" width="100%">
	<tr>
		<td><%
		Dim IsLogin
    Dim TheString
		dim UserName , Password
		if Request.Cookies.count<>0 then

	  Username=request.Cookies(Monolith_CookiesName)("Username")
	  Password=request.Cookies(Monolith_CookiesName)("Password")
		''���˷Ƿ��ַ�

  sql= "SELECT * FROM [Monolith_Users] where [Username] = '" & Username & "' and [password]='" & Password & "'"
  set rs=conn.Execute(sql)
  if not (rs.BOF or rs.eof) then
    TheString	= UserName & "����ӭ������<a href='Logout.asp'>�������˳�</a>"
	IsLogin		= 1
  end if
  rs.Close
end if
if IsLogin <> 1 then  TheString = "<a href='Reg.asp'>�û�ע��</a>&nbsp;&nbsp;<a href='Login.asp'>�û���¼</a>"
Response.write TheString
%> </td>
		<td>
		<p><a href="Default.asp">��ҳ</a></p>
		</td>
		<td>
		<p>
		<a href="http://www.showde.com/zero/index.asp?vt=bycat&cat_id=49" target="_blank">
		��������</a></p>
		</td>
		<td>
		<p><a href="AboutUs.asp">��������</a></p>
		</td>
	</tr>
</table>
<% if Username="Zerolone" then %>
<table width="100%" border="1" bgcolor="#CCCCCC">
	<tr>
		<td><a href="AdminBoardAdd.asp">��Ӱ��</a></td>
		<td><a href="AdminBoard.asp">���а��/����</a></td>
		<td>��</td>
		<td>��</td>
		<td><a href="AdminAboutUs.asp">�޸Ĺ�������</a></td>
	</tr>
</table>
<%end if%> <hr>
<table id="table1" cellspacing="1" width="100%" border="0" style="border-style: solid; border-width: 1px;" bordercolor="#000000">
	<tr>
		<td background="images/tile_back.gif">
		<img border="0" src="images/logo4.gif" width="221" height="67"></td>
	</tr>
	<tr height="26">
		<td background="images/tile_sub.gif">
		<table cellspacing="0" cellpadding="0" width="100%" class="TdTitle">
			<tr>
				<td><a href="#">&nbsp;Welcome to MyBBs.&nbsp;&nbsp; Current Version 
				1.0(������)</a></td>
				<td width="233"><a href="#">
				<img src="images/atb_help.gif" border="0"> ����</a>&nbsp;&nbsp;
				<a href="#"><img src="images/atb_search.gif" border="0"> ����</a>&nbsp;&nbsp;
				<a href="#"><img src="images/atb_members.gif" border="0"> �û�</a>&nbsp;&nbsp;
				<a href="#"><img src="images/atb_calendar.gif" border="0"> ����</a>&nbsp;
				</td>
			</tr>
		</table>
		</td>
	</tr>
</table>