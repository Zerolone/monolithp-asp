
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
url="http://www.baidu.com/baidu?word=site%3A"&id&""
        wstr=getHTTPPage(url) 
                                                                                           
start=Newstring(wstr,"找到相关网页")
        over=Newstring(wstr,"秒</nobr>")
    body=mid(wstr,start+0,over-start)

body = replace(body,"找到相关网页",""&id&" 在baidu上的搜索结果：找到相关网页")

response.write body
response.write "秒。</font><a href=""http://www.baidu.com/baidu?word=site%3A"&request("domain")&""" target=""_blank""> 点击察看详细内容</a><br>"
else response.write "您可能输入错误，请重新搜索<br>"
end if
%>
</body>      
</html>

