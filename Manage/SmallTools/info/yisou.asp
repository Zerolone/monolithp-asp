
<%
On Error Resume Next
Server.ScriptTimeOut=9999999
Function getHTTPPage(Path)
        t = GetBody(Path)
        getHTTPPage=BytesToBstr(t,"gb2312")
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
<font size="-1">
<%
if request("submit")<>"" or request("domain")<>"" then 
Dim wstr,str,url,start,over,id
ID = Request.QueryString("domain")
url="http://www.yisou.com/search?p=site%3A"&id&""
        wstr=getHTTPPage(url) 

if instr(wstr,"�Ҳ��������Ĳ�ѯ") then
none="google��ʱû����¼ "&id&" ����</font><a href=""http://submit.search.yahoo.com/free/request"" target=""_blank"">����</a><font size=""-1"">����ύ��վ��</font><br>"
response.write ""&none&""
else
                                                                                           
start=Newstring(wstr,"&nbsp;&nbsp;Լ&nbsp;<b>")
        over=Newstring(wstr,"��&nbsp;��&nbsp;")
    body=mid(wstr,start+0,over-start)

body = replace(body,"&nbsp;&nbsp;Լ&nbsp;","Yisou��¼"&request("domain")&"Լ </font><a font='#FB5E3C'>")
body = replace(body,"</b>&nbsp;","</b></font><a font='#FB5E3C'> ƪ")


response.write body
response.write "</font><a href=""http://www.yisou.com/search?p=site%3A"&request("domain")&""" target=""_blank""> ����쿴��ϸ����</a><br>"

end if
else response.write "�����������������������<br>"


end if
%>
</body>      
</html>

