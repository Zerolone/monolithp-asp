<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!-- #include file="connection.asp" -->
<html >
<link rel="stylesheet" type="text/css" href="main.css">
<body class=clblue>
<%		dim strSQL,rsSysInfo,showSysop
 	    showSysop=mid(Application("Sysop"),2,len(Application("Sysop"))-2)	

		%>
<div align=center><center>
<Br>
<form action="set.asp" name=frmTitle>
      <table width=85% class=InfoMan border=1 cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
        <tr bgcolor="#3399CC"> 
          <td colspan=2 align=center>վ���ռ�֮ϵͳ����</td>
        </tr>
        <tr bgcolor="#3399CC"> 
          <td width=6% align=right>1&nbsp;</td>
          <td> &nbsp;&nbsp;������������ 
            <INPUT TYPE="text" size=30 NAME="title" value="<%=Application("Title")%>" >&nbsp;&nbsp;
		  <INPUT TYPE="submit" name="SetIt" value="����">
  </td></tr>
        <tr bgcolor="#3399CC"> 
          <td width=6% align=right>2&nbsp;</td>
          <td> &nbsp;&nbsp;����Ա������ 
            <INPUT TYPE="text" size=30 NAME="sysop" value="<%=showSysop%>" >&nbsp;&nbsp;
		  <INPUT TYPE="submit" name="SetIt" value="����Ա">
  </td></tr>
 </table>
</form>
<%		dim rsBook,iUser,i,byOrder,deleteid,misspassid
		dim rsBoard,rsBook1,strFtitle,strBname
		set rsBook=server.CreateObject("ADODB.RecordSet")
		set rsBook1=server.CreateObject("ADODB.RecordSet")
		set rsBoard=server.CreateObject("ADODB.RecordSet")
%>
</center></div>
</body>
</html>
