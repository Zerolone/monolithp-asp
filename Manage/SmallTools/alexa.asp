<!--
********************************ALEXA ��վ��������С͵***********************
*  ���ߣ�����   QQ��103201165                                               *
*  ��ַ��http://www.07550755.com ������������                               *
*             ��ӭ����   ��ͬ�ٽ�                                           *
*****************************************************************************
-->

<html>
<head>
<title>��վ����������ѯ</title>
</head>
<table align=center border=1 width='778' cellspacing=1 height=300 cellpadding=1 bgcolor=#5f79a1 bordercolor=#FFFFFF>
<tr><td bgcolor=ffffff width=100% align=center>
<table align=center ><tr><td>
<form name="se" method=post action="">
��������Ҫ��ѯ������</td><td><input type=text name="url" value="www."></td>
<td><input type=submit name=submit value="��ѯ"></td></tr>
</table>
</form>
<%
if request("submit")<>"" or request("url")<>"" then 
url="http://www.alexa.com/data/details/traffic_details?q=&url="&replace(request("url"),"http://","")
getsms=gethttppage(url)
if instr(getsms,"Traffic rank:</span>") then
picurl="http://traffic.alexa.com/graph?w=379&h=216&r=6m&u="&request("url")&"/&u="
response.write "<img src="&picurl&">"
star=instr(getsms,"Traffic rank:</span>")+120
endd=instr(star,getsms,"</table>")+8
getstr=mid(getsms,star,endd-star)
getstr=replace(getstr,">Today<",">����<")
getstr=replace(getstr,"1 wk. Avg.","һ����ƽ��")
getstr=replace(getstr,"3 mos. Avg.","������ƽ��")
getstr=replace(getstr,"3 mos. Change","����������")
response.write getstr
response.write "<br>�����������ԣ�http://www.alexa.com,��ϸ��Ϣ���Ե�½������վ��ѯ"
end if
end if
%>
<%
	function getHTTPPage(url) 
		on error resume next 
		dim http 
		set http=Server.createobject("Microsoft.XMLHTTP") 
		Http.open "GET",url,false 
		Http.send() 
		if Http.readystate<>4 then
			exit function 
		end if 
		getHTTPPage=bytes2BSTR(Http.responseBody) 
		set http=nothing
		if err.number<>0 then err.Clear  
	end function 
	Function bytes2BSTR(vIn) 
		dim strReturn 
		dim i1,ThisCharCode,NextCharCode 
		strReturn = "" 
		For i1 = 1 To LenB(vIn) 
			ThisCharCode = AscB(MidB(vIn,i1,1)) 
			If ThisCharCode < &H80 Then 
				strReturn = strReturn & Chr(ThisCharCode) 
			Else 
				NextCharCode = AscB(MidB(vIn,i1+1,1)) 
				strReturn = strReturn & Chr(CLng(ThisCharCode) * &H100 + CInt(NextCharCode)) 
				i1 = i1 + 1 
			End If 
		Next 
		bytes2BSTR = strReturn 
	End Function 
	function getHTTPpic(url) 
		on error resume next 
		dim http 
		set http=Server.createobject("Microsoft.XMLHTTP") 
		Http.open "GET",url,false 
		Http.send() 
		if Http.readystate<>4 then
			exit function 
		end if 
		getHTTPPage=Http.responseBody 
		set http=nothing
		if err.number<>0 then err.Clear  
	end function 
%>
</td></tr></table>
