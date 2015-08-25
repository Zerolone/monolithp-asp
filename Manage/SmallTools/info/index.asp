
<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Dim Url,Html,objRegExp
domain= Request.QueryString("domain") 
rankcheck= Request("rank")
whoischeck = Request("whois")
alexacheck= Request("alexa")
baiducheck= Request("baidu")
yisoucheck= Request("yisou")
googlecheck= Request("google")
speedcheck= Request("speed")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
 <link rel=stylesheet href="css.css">
	  	  <style><!--
body,td,div,p,a,font,span{
    font-family:arial,sans-serif;
       }
//-->
</style>
</head>
 <body bgcolor="#FFFFFF" text="#000000">
<form action="index.asp" method="get" name="form" id="form">
  <div align="center">
    <p><font face="Arial" size="3">www. 
      <input type="text" name="domain" size="20" value="<%=domain%>">
      &nbsp; &nbsp;</font> 
      <input type="submit" name="Submit" value="查询">
      <br>
      <br>
      <input type="checkbox" name="rank" value="checked" <%=rankcheck%>>
      <font size=""-1""> Google Rank查询 </font> 
      <input type="checkbox" name="google" value="checked" <%=googlecheck%>>
      <font size=""-1"">Google收录查询 </font> 
      <input type="checkbox" name="baidu" value="checked" <%=baiducheck%> >
      <font size=""-1"">Baidu收录查询 </font> 
      <input type="checkbox" name="yisou" value="checked" <%=baiducheck%> >
      <font size=""-1"">Yisou收录查询 </font> 
      <input type="checkbox" name="alexa" value="checked"  <%=alexacheck%> >
      <font size=""-1""> Alexa排名查询 </font> </p>
    </div>
</form>
 <% if domain<>"" then 
str1=domain
str2="."
If instr(str1,str2)>0 Then
 if  rankcheck="" and whoischeck="" and  alexacheck= "" and baiducheck="" and  googlecheck="" and speedcheck="" then
Response.Write "<font size=""-1""> 至少选一样把，朋友！</font>"
		Response.End
		end if
%>
<table width="700" align="center" cellpadding="2" cellspacing="1">
  <tr> 
    <td style="background-color: #d8e4f1">
      <table width="100%" align="center" cellpadding="6" cellspacing="0">
        <% if rankcheck<>"" then %>
        <tr> 
          <td colspan="2" style="background-color: #efefef"> <font size=""-1""> 
            Google Rank查询</font></td>
        </tr>
        <tr> 
          <td colspan="2">
            <iframe frameborder=0 height=20 name=rank scrolling=no 
            src="rank.asp?domain=www.<%=Request("domain")%>" 
            width="100%"></iframe>
          </td>
        </tr>
        <%end if %>
        <% if googlecheck<>"" then %>
        <tr> 
          <td colspan="2" style="background-color: #efefef"><font size=""-1"">Google收录查询</font></td>
        </tr>
        <tr> 
          <td colspan="2"> 
             <iframe frameborder=0 height=20 name=google scrolling=no 
            src="google.asp?domain=<%=Request("domain")%>" 
            width="100%"></iframe>
          </td>
        </tr>
        <%end if %>
        <% if baiducheck<>"" then %>
        <tr> 
          <td colspan="2" style="background-color: #efefef"> <font size=""-1"">Baidu收录查询</font></td>
        </tr>
        <tr> 
          <td colspan="2">
            <iframe frameborder=0 height=20 name=baidu scrolling=no 
            src="baidu.asp?domain=<%=Request("domain")%>" 
            width="100%"></iframe>
          </td>
        </tr>
        <%end if %>
        <% if yisoucheck<>"" then %>
        <tr> 
          <td colspan="2" style="background-color: #efefef"> <font size=""-1"">Yisou收录查询</font></td>
        </tr>
        <tr> 
          <td colspan="2">
            <iframe frameborder=0 height=20 name=baidu scrolling=no 
            src="yisou.asp?domain=<%=Request("domain")%>" 
            width="100%"></iframe>
          </td>
        </tr>
        <%end if %>
        <% if alexacheck<>"" then %>
        <tr> 
          <td colspan="2" style="background-color: #efefef"> <font size="-1"> 
            Alexa排名查询</font></td>
        </tr>
        <tr> 
          <td colspan="2">
            <iframe frameborder=0 height=110 name=alexa scrolling=no 
            src="alexa.asp?domain=<%=Request("domain")%>" 
            width="100%"></iframe>
      
         </td>
        </tr>
        <%end if %>
        <% if whoischeck<>"" then %>
        <tr> 
          <td colspan="2" style="background-color: #efefef"> <font size=""-1""> 
            Whois查询</font></td>
        </tr>
        <tr> 
          <td colspan="2">
               <iframe frameborder=0 height=280 name=whois scrolling=no 
            src="http://www.vipso.com/getwhois.asp?domain=<%=domain%>&se=<%=se%>" 
            width="100%"></iframe>
          </td>
        </tr>
        <%end if %>
        </tr>
      </table>
    </td>
  </tr>
</table>
<form action="index.asp" method="get" name="form" id="form">
  <div align="center"> 
    <p><font face="Arial" size="3">www. 
      <input type="text" name="domain" size="20" value="<%=domain%>">
      &nbsp; &nbsp;</font> 
      <input type="submit" name="Submit" value="查询">
      <br>
      <br>
      <input type="checkbox" name="rank" value="checked" <%=rankcheck%>>
      <font size=""-1""> Google Rank查询 </font> 
      <input type="checkbox" name="google" value="checked" <%=googlecheck%>>
      <font size=""-1"">Google收录查询 </font> 
      <input type="checkbox" name="baidu" value="checked" <%=baiducheck%> >
      <font size=""-1"">Baidu收录查询 </font> 
      <input type="checkbox" name="yisou" value="checked" <%=baiducheck%> >
      <font size=""-1"">Yisou收录查询 </font> 
      <input type="checkbox" name="alexa" value="checked"  <%=alexacheck%> >
      <font size=""-1""> Alexa排名查询 </font> 
  </div>
</form>
<%
else 
response.write"<font size=""-1""> 请输入正确的域名</font>"
end if
end if 
%>