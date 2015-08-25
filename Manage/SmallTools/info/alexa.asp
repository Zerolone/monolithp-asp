
<%
On Error Resume Next
Server.ScriptTimeOut=9999999
Function getHTTPPage(Path)
        t = GetBody(Path)
        getHTTPPage=BytesToBstr(t,"GB2312")
End function

Function GetBody(url) 
        on error resume next
        Set Retrieval = CreateObject("Microsoft.XMLHTTP") 
        With Retrieval 
        .Open "Get", url, False, "", "" 
        .Send 
        GetBody = .ResponseBody
        End With 
        Set Retrieval = Nothing 
End Function

Function BytesToBstr(body,Cset)
        dim objstream
        set objstream = Server.CreateObject("adodb.stream")
        objstream.Type = 1
        objstream.Mode =3
        objstream.Open
        objstream.Write body
        objstream.Position = 0
        objstream.Type = 2
        objstream.Charset = Cset
        BytesToBstr = objstream.ReadText 
        objstream.Close
        set objstream = nothing
End Function
Function Newstring(wstr,strng)
        Newstring=Instr(lcase(wstr),lcase(strng))
        if Newstring<=0 then Newstring=Len(wstr)
End Function
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel=stylesheet href="css.css">
<style>
<!--
 body,td,div,p,a,font,span{
 font-family:arial,sans-serif;
}
//-->
</style>
<body  leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#d8e4f1">
<table width="100%" border="0" cellpadding="0" cellspacing="0"><tr><td width="58%">
<%
if request("submit")<>"" or request("domain")<>"" then 
url="http://www.alexa.com/data/details/traffic_details?q=&url="&replace(request("domain"),"http://","")
getsms=gethttppage(url)
if instr(getsms,"Traffic rank:</span>") then
star=instr(getsms,"Traffic rank:</span>")+120
endd=instr(star,getsms,"</table>")+8
getstr=mid(getsms,star,endd-star)
getstr=replace(getstr,">Today<",">今日排名<")
getstr=replace(getstr,"1 wk. Avg.","一周平均")
getstr=replace(getstr,"3 mos. Avg.","三月平均")
getstr=replace(getstr,"3 mos. Change","三月浮动")
getstr=replace(getstr,"#CCCC99","#efefef")
getstr=replace(getstr,"<span class=""body"">","<font size=""-1"">")
response.write getstr
picurl="http://traffic.alexa.com/graph?w=200&h=100&r=6m&u="&request("domain")&"/&u="
response.write "<br>  <a href =""alexa.asp?url="&request("domain")&""" target=""_blank"">点击察看alexa中文化说明</a>&nbsp;&nbsp;  <a href =""http://www.alexa.com/data/details/?url="&request("domain")&""" target=""_blank"">点击察看alexa该网址说明</a><td width=""50%""><img src="&picurl&">"
end if
else response.write "您可能输入错误，请重新搜索<br>"
end if
%>

</td></tr></table></body></html>