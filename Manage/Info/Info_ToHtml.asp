<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include file ="InfoSetting.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">

<%
	Server.ScriptTimeout=999999

	Dim Info_ID
	Info_ID					= Request("Info_ID")
	CheckNum Info_ID

	Dim Info_Title, Info_Title2, Info_Class, Info_Author, Info_From, Info_KeyWord, Info_Memo, Info_Pic1, Info_Pic2, Info_FileName, Info_Content, Info_Date, Info_Template

	'-------------------0--------------1-------------2-------------3-------------4--------------5-------------6------------7-----------8---------------9--------------10-----------11
	Sql = "Select [Info_Title], [Info_Title2], [Info_Class], [Info_Author], [Info_From], [Info_KeyWord], [Info_Memo], [Info_Pic1], [Info_Pic2], [Info_FileName], [Info_Date], [Info_Template] From [Monolith_Info] Where [Info_ID]=" & Info_ID
	Rs.open Sql, conn ,1 ,1
	if Not(Rs.Bof Or Rs.Eof) then
		Info_Title		= Rs(0)
		Info_Title2		= Rs(1)
		Info_Class		= Rs(2)
		Info_Author		= Rs(3)
		Info_From			= Rs(4)
		Info_KeyWord	= Rs(5)
		Info_Memo			= Rs(6)
		Info_Pic1			= Rs(7)
		Info_Pic2			= Rs(8)
		Info_FileName	= Rs(9)
		Info_Date			= Rs(10)
		Info_Template	= Rs(11)		
	else
		Response.write "该信息不存在！"
		Response.end
	end if
	Rs.Close

  if IsNull(Info_Title)		then	Info_Title		= "　"
  if IsNull(Info_Class) 	then	Info_Class		= "　"
  if IsNull(Info_Author)	then	Info_Author		= "　"
  if IsNull(Info_From)		then	Info_From			= "　"
  if IsNull(Info_Date)		then	Info_Date			= "　"
	
	if Info_Date = "" then
		Response.write "日期出错！"
		Response.end
	end if

	if Info_Template = "" then
		Response.write "模版出错！"
		Response.end
	end if

  Sql = "Select [Template_Content] From [Monolith_Template] Where [Template_ID]=" & Info_Template
  Rs.Open Sql, Conn , 1, 1
  If Not(Rs.Bof Or Rs.Eof)then
		Info_Template = Rs(0)
  else
    Response.write "该模版不存在！"
		Response.End
  end if
  rs.close
	
  '标题	：{Info_Title}
  Info_Template = Replace(Info_Template,"{Info_Title}",			Info_Title)
  
  '发布日期：{Info_Date}
  Info_Template = Replace(Info_Template,"{Info_Date}",			Info_Date)
  
  '作者	： {Info_Author}
  Info_Template = Replace(Info_Template,"{Info_Author}",		Info_Author)
  
  '来源	：{Info_From}
  Info_Template = Replace(Info_Template,"{Info_From}",			Info_From)
  
  '文章ID：{Info_ID}
  Info_Template = Replace(Info_Template,"{Info_ID}",				Info_ID)

  '文章Class：{Info_Class}
  Info_Template = Replace(Info_Template,"{Info_Class}",			Info_Class)

 
  '----------------

	Dim MyFile, FileSavePath, UrlPath, TheYear, TheMonth, TheDay, FSO

  Set MyFile=Server.CreateObject("Scripting.FileSystemObject")
  set FSO = server.CreateObject("Scripting.FileSystemObject")
  
	TheYear		= Year(Info_Date)
  TheMonth	= Month(Info_Date)
  TheDay		= Day(Info_Date)

	FileSavePath	= InfoSettingArr(6)
	UrlPath				= InfoSettingArr(7)
	if right(UrlPath,1) <> "/"			then UrlPath=UrlPath & "/" 
	if right(FileSavePath,1) <> "/" then FileSavePath=FileSavePath & "/" 

	UrlPath				= UrlPath & TheYear & "/" & TheMonth & TheDay & "/"

	FileSavePath	= FileSavePath & TheYear
  FileSavePath	= FileSavePath	& "/" & TheMonth & TheDay & "/"

	if CreateFolderII(FileSavePath) = false then
		Response.write "文件路径错误！请察看信息发布系统设置"
		Response.end
	end if

	if right(FileSavePath,1) <> "/" then FileSavePath=FileSavePath & "/" 
	if right(UrlPath,1) <> "/"			then UrlPath=UrlPath & "/" 	
  
	Dim tf, Fname, AllFileName, TheFileName, i , j, ThePage

	Fname							= Replace(Info_FileName, ".htm", "")

	TheFileName	= Fname & ".htm"
  '信息全称：{Info_FileName}
  Info_Template = Replace(Info_Template,"{Info_FileName}", UrlPath & TheFileName)
 
	Set tf 			= MyFile.CreateTextFile(FileSavePath & TheFileName ,True)
	tf.Write (Info_Template)
  tf.Close
  AllFileName = AllFileName & "<li><a target='_blank' href='" & UrlPath & TheFilename & "'>" & TheFilename & "</a>"
%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="4"><b><font color="#FFFFFF">&nbsp;信息管理 &gt;&gt; 静态页面生成</font></b></td>
	</tr>
	<tr height="20" bgcolor="#FFFFFF">
		<td width="100%" bgcolor="#999999"><font color="#FFFFFF">&nbsp;文件名</font></td>
	</tr>
	<tr bgcolor="#FFFFFF">
			<td height="22">　<strong><%=AllFileName%></strong></td>
	</tr>
</table>
<table><tr><td height="432"></td></tr></table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->