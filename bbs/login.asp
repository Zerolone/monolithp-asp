<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!-- #include file="connection.asp" -->
<!-- #include file="string.asp" -->
<!-- #include file="std.asp" -->
<!-- #include file="encrypt.asp" -->

<%if request("userid")="" then  %>
<html>
<title><%=Application("Title")%></title>
<link rel="stylesheet" type="text/css" href="main.css">
<SCRIPT LANGUAGE="JavaScript">
<!--//
function check()
{
	if (document.frmaLogin.userid.value.length<1) {
		alert("用户名不能为空");
		document.frmaLogin.userid.focus();
	}
	else if (document.frmaLogin.PASSWORD.value.length<1) {
		alert("用户密码不能为空");
		document.frmaLogin.PASSWORD.focus();
	}
	else
        document.frmaLogin.submit();
}

function GoGuest()
{
 document.frmaLogin.userid.value="guest";
 document.frmaLogin.PASSWORD.value="guest";
 document.frmaLogin.submit();
}

function GoReg()
{
	top.frmMain.location.href="adduser.asp"	
}

//-->
</SCRIPT>
<style type="text/css">
<!--
.white2 {  font-size: 10pt; color: #FFFFFF; text-decoration: none}
-->
</style>

<body class=clblack>
<script language="javascript">
	<%if( session("Msg")=1 ) then %>
		alert("<% =session("StrMsg")%>");
	<%	session("StrMsg")=" "  
		session("msg")=0
	  end if %>
</script>
<FORM name=frmaLogin METHOD="POST" ACTION="login.asp">
  <table border=0 color=FFFFFF  align=center width=168>
    <tr>
      <td height="91"> 
        <TABLE  BORDER=1 class=login align=left width=83% height="104" cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#99FF99" >
          <tr bgcolor="#3399CC"> 
            <TD align=center COLSPAN=2 color=#ffffff height="2"><font color=#ffffff><b><font size="2">登录论坛</font></b></font></TD>
          </tr>
          <TR bgcolor="#99FF99" > 
            <TD  ALIGN=CENTER colspan="2"> <a href="javascript:GoReg()" class=Alogin><font color="#FFFFFF"><span class="white2"><font color="#000000">免费注册</font></span></font></a> 
            </TD>
          </tr>
          <TR bgcolor="#99FF99"> 
            <TD  ALIGN=RIGHT width="34%"><b><font size="2">代号:</font></b></TD>
            <TD width="66%" > <b><font size="2"> 
              <INPUT TYPE="text" NAME="userid" value='<%=request("uid")%>' SIZE="8" MAXLENGTH="20" >
              </font></b></TD>
          </TR>
          <TR bgcolor="#99FF99"> 
            <TD  ALIGN=RIGHT width="34%" height="7"><b><font size="2">密码:</font></b></TD>
            <TD width="66%" height="7" > <b><font size="2"> 
              <INPUT TYPE="password" NAME="PASSWORD" SIZE="8" MAXLENGTH="20">
              </font></b></TD>
          <TR align=center bgcolor="#99FF99"> 
            <TD  COLSPAN=2 ALIGN=CENTER height="4"> <b><font size="2"> 
              <INPUT TYPE=BUTTON  VALUE='进   入' onClick="check()" class=login>
              </font></b></TD>
          
        </TABLE>
</td></tr></table>
</FORM>
<div align="center">
  <SCRIPT LANGUAGE="JavaScript">
<!--//
document.frmaLogin.userid.focus();
//-->
</SCRIPT>
  <br>
  <a href="javascript:GoReg()" class=Alogin><font color="#FFFFFF"><span class="white2"><b>免费注册</b></span></font></a><b><a href="javascript:GoGuest()" class=Alogin><br>
  <span class="white2">游客进入*无须密码</span></a> </b><br>
  <hr>
  <br>
</div>
</BODY>
</HTML>

<%	else
		dim StrSQL,rsUserInfo,StrMsg,blOK,strUserid,strPassword
		set rsUserInfo=Server.CreateObject("ADODB.RECORDSET")

		strUserid=request("userid")
		strPassword=request("password")
		strSQL="select * from UserInfo where userid='"+strUserid+"'"
		rsUserInfo.Open strSQL,myconn
		blOK=false
		if not rsUserInfo.BOF and not rsUserInfo.EOF then
			if (rsUserInfo("password"))=mistake((strPassword)) then
				blOK=true
				strPassword="****"
			else
				strMsg="密码错误，请重新输入！"
			end if
			dim rsSysLogin 
			set rsSysLogin=Server.CreateObject("ADODB.RECORDSET")
			strSQL="insert into syslogin(userid,password,ip,logindate,loginOK) values('"
			strSQL=strSQL+strUserid+"','"+strPassword+"','"+Request.ServerVariables("REMOTE_ADDR")+"',"
			strSQL=strSQL+"now(),"+CSTR(blOK)+")"
			rsSysLogin.Open StrSQL,myconn
		else
			StrMsg="用户名错误，请重新输入！"
		end if
	    if not blOk then 
			session("strMsg")=strMsg
			session("msg")=1
			response.redirect "login.asp?uid="+strUserid
		else
			session("userid")=rsUserInfo("userid")
			session("nickname")=rsUserInfo("nickname")
			initLogin
			session("email")=rsUserInfo("email")
			session("uid")=rsUserInfo("id")
			session("sign")=rsUserInfo("sign")&" "
			
			dim LastLogin,iLogin
			LastLogin=rsUserInfo("LastLogin")
			iLogin=rsUserInfo("iLogin")
			rsUserInfo.close
			strSQL="update UserInfo set ilogin=ilogin+1, iPerience=iPerience+1,LastLogin=now() where userid='"+session("userid")+"'"
			rsUserInfo.Open strSQL,myconn
			session("msg")=1
			strMsg="亲爱的"+session("userid")+":\n\n"
			strMsg=strMsg+"这是你第"+cstr(iLogin+1)+"次光临寒舍\n"
			strMsg=strMsg+"上次光临时间："+FormatDateTime(LastLogin,1)
			session("strMsg")=strMsg
			response.redirect "FullScreen.htm"
		end if
		myconn.close
	end if 
%>
<%	set rsUserInfo=nothing
	set myconn=nothing
%>