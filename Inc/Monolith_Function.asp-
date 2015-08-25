<%
'-----------------------------------------------
'作者:Zerolone ---------------------------------
'日期:20040923----------------------------------
'功能:检测ID是否为整形--------------------------
'参数:integer-----------------------------------
'-----------------------------------------------
Function CheckNum(ID)
	if not IsNumeric(ID) then
		response.write "<script>alert('非法参数');history.back()</script>"
		response.end
	end if
End Function

'-----------------------------------------------
'作者:Zerolone ---------------------------------
'日期:20041121----------------------------------
'功能:检测ID是否为整形， 返回T/F----------------
'参数:integer-----------------------------------
'-----------------------------------------------
Function CheckNumII(ID)
	if not IsNumeric(ID) then
		CheckNumII = false
	else
		CheckNumII = true
	end if
End Function


'-----------------------------------------------
'作者:------------------------------------------
'日期:------------------------------------------
'功能:生成随机的十位的数字----------------------
'参数:------------------------------------------
'-----------------------------------------------
Function MakeRndNum()
	Dim fname
	fname = now()
	fname = replace(fname,"-","")
	fname = replace(fname," ","") 
	fname = replace(fname,":","")
	fname = replace(fname,"PM","")
	fname = replace(fname,"AM","")
	fname = replace(fname,"上午","")
	fname = replace(fname,"下午","")
	fname = int(fname) + int((10-1+1)*Rnd + 1)
	MakeRndNum=fname
End Function 

'-----------------------------------------------
'作者:Zerolone ---------------------------------
'日期:20040923----------------------------------
'功能:去除多余的字符, 根据参数是否需要显示省略号
'参数:String, integer, 1/0----------------------
'-----------------------------------------------
Function CutLongStrDot(str, maxLength, hasDot)
	Dim i, j
	If Not ISNull(Str) then
		j = 0
		str = Trim(str)
		For i = 1 to Len(str)
			IF Abs(Asc(Mid(str, i, 1))) > 255 Then j = j + 2 Else j = j + 1
			IF j >= maxLength Then
				str = Left(str, i-1)
				if hasDot = 1 then	str = str & "..."
				Exit For
			End IF
		Next
		CutLongStrDot = str
	End if
End Function

'-----------------------------------------------
'作者:Zerolone ---------------------------------
'日期:20040924----------------------------------
'功能:计算字符长度------------------------------
'参数:String------------------------------------
'-----------------------------------------------
Function CountStr(str)
	Dim i, j
	If Not ISNull(Str) then
		j = 0
		str = Trim(str)
		For i = 1 to Len(str)
			IF Abs(Asc(Mid(str, i, 1))) > 255 Then j = j + 2 Else j = j + 1
		Next
		CountStr = j
	End if
End Function

'-----------------------------------------------
'作者:Zerolone ---------------------------------
'日期:20040923----------------------------------
'功能:过滤HTML语法------------------------------
'参数:String------------------------------------
'-----------------------------------------------
Function MonolithEncode(TheString)
	TheString				= Replace(TheString, chr(34), "&quot;")
	TheString				= Replace(TheString, "'", "＇")
	TheString				= Replace(TheString, "<", "&lt;")
	TheString				= Replace(TheString, ">", "&gt;")
'	TheString				= Replace(TheString, ";", "；")
	TheString				= Replace(TheString, "-", "—")
	TheString				= Replace(TheString, Chr(10), "<BR>")
	TheString				= Replace(TheString, " ", "&nbsp;")
	MonolithEncode	= TheString
End Function

'-----------------------------------------------
'作者:Zerolone ---------------------------------
'日期:20040923----------------------------------
'功能:把HTML代码写成一行------------------------
'参数:String------------------------------------
'-----------------------------------------------
Function MonolithEncodeII(TheString)
//	TheString				= Replace(TheString, chr(34), "\&quot;")
	TheString				= Replace(TheString, chr(34), "'")
	TheString				= Replace(TheString, Chr(10), "")
	TheString				= Replace(TheString, Chr(13), "")
	MonolithEncodeII	= TheString
End Function


'-----------------------------------------------
'作者:Zerolone ---------------------------------
'日期:20050220----------------------------------
'功能:还原HTML语法------------------------------
'参数:String------------------------------------
'-----------------------------------------------
Function MonolithDecode(TheString)
	TheString				= Replace(TheString, "&lt;", "<")
	TheString				= Replace(TheString, "&gt;", ">")
	TheString				= Replace(TheString, "<BR>", Chr(10))
	TheString				= Replace(TheString, "&nbsp;", " ")
	MonolithDecode	= TheString
End Function

'-----------------------------------------------
'作者:Zerolone ---------------------------------
'日期:20050110----------------------------------
'功能:过滤HTML语法------------------------------
'参数:String------------------------------------
'-----------------------------------------------
Function Monolith_Encode(TheString)
  if IsNull(TheString) then
    TheString = ""
   Exit Function 
  end if

	TheString				= Replace(TheString, chr(34), "&quot;")
	TheString				= Replace(TheString, "'", "＇")
	TheString				= Replace(TheString, "<", "&lt;")
	TheString				= Replace(TheString, ">", "&gt;")
	TheString				= Replace(TheString, ";", "；")
	TheString				= Replace(TheString, "-", "—")
	TheString				= Replace(TheString, Chr(13)&Chr(10), "<br>")
	Monolith_Encode	= TheString
End Function

'-----------------------------------------------
'作者:Zerolone ---------------------------------
'日期:20050328----------------------------------
'功能:还原HTML语法------------------------------
'参数:String------------------------------------
'-----------------------------------------------
Function Monolith_Decode(TheString)
	TheString				= Replace(TheString, "&quot;",	chr(34))
	TheString				= Replace(TheString, "＇", 			"'")
	TheString				= Replace(TheString, "&lt;", 		"<")
	TheString				= Replace(TheString, "&gt;", 		">")
	TheString				= Replace(TheString, "；", 			";")
	TheString				= Replace(TheString, "—", 			"-")
	TheString				= Replace(TheString, "<br>", 		Chr(13)&Chr(10))
	Monolith_Decode	= TheString
End Function



'-----------------------------------------------
'作者:Zerolone ---------------------------------
'日期:20040923----------------------------------
'功能:根据参数进行空格--------------------------
'参数:integer-----------------------------------
'-----------------------------------------------
Sub LoopNBSP(num)
	Dim i
	for i =0 to num
		Response.write "&nbsp;"
	next
End Sub

'-----------------------------------------------
'作者:------------------------------------------
'日期:20041008----------------------------------
'功能:读取文件, 不检测文件是否存在--------------
'参数:String, 文件名----------------------------
'-----------------------------------------------
Function FSOFileRead(FileName)
	Dim objFSO,objCountFile,FiletempData
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objCountFile = objFSO.OpenTextFile(Server.MapPath(filename),1,True)
	FSOFileRead = objCountFile.ReadAll
	objCountFile.Close
	Set objCountFile=Nothing
	Set objFSO = Nothing
End Function

'-----------------------------------------------
'作者:------------------------------------------
'日期:20041008----------------------------------
'功能:修改文件指定内容, 不检测文件是否存在------
'参数:String,String,String 文件名,目标,替换内容-
'版本:1.1---------------------------------------
'-----------------------------------------------
Function FSOFileReplaceString(Filename, ToFileName, TheString,TheReplaceString)
	FiletempData	= FSOFileRead(Filename)
	FiletempData	= Replace(FiletempData, TheString, TheReplaceString)
	
	Dim objFSO,objCountFile,FiletempData
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objCountFile=objFSO.CreateTextFile(Server.MapPath(ToFileName),True)
	objCountFile.Write FiletempData 
	objCountFile.Close
	Set objCountFile=Nothing	
	Set objFSO = Nothing
End Function

'-----------------------------------------------
'版本:1.1---------------------------------------
'Function FSOFileReplaceString(Filename, ToFileName, TheString,TheReplaceString)
'	Dim objFSO,objCountFile,FiletempData
'
'	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
'	Set objCountFile = objFSO.OpenTextFile(Server.MapPath(Filename),1,True)
'	
'	FiletempData = objCountFile.ReadAll
'	objCountFile.Close
'
'	
'	FiletempData=Replace(FiletempData, TheString, TheReplaceString)
'	Set objCountFile=objFSO.CreateTextFile(Server.MapPath(ToFileName),True)
'	objCountFile.Write FiletempData 
'	objCountFile.Close
'	Set objCountFile=Nothing	
'	Set objFSO = Nothing
'End Function
'-----------------------------------------------

'-----------------------------------------------
'作者:------------------------------------------
'日期:20041013----------------------------------
'功能:远程存储文件------------------------------
'参数:本地文件名， 远程地址---------------------
'版本:1.0---------------------------------------
'-----------------------------------------------
	Sub SaveRemoteFile(LocalFileName,RemoteFileUrl)
		Dim Ads, Retrieval, GetRemoteData
		On Error Resume Next
		Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
		With Retrieval
			.Open "Get", RemoteFileUrl, False, "", ""
			.SetRequestHeader "Referer","http://www.digichina.cn/" 
			.Send
			GetRemoteData = .ResponseBody
		End With
		Set Retrieval = Nothing
		Set Ads = Server.CreateObject("Adodb.Stream")
		With Ads
			.Type = 1
			.Open
			.Write GetRemoteData
'			.SaveToFile Server.MapPath(LocalFileName), 2
			.SaveToFile LocalFileName , 2
			.Cancel()
			.Close()
		End With
		Set Ads=nothing
	End Sub

'-----------------------------------------------
'作者:------------------------------------------
'日期:20041121----------------------------------
'功能:出错提示----------------------------------
'参数:挺示内容， 图标， 跳转地址----------------
'版本:1.0---------------------------------------
'-----------------------------------------------
	Sub ErrMsg(message,sflag,url)
%>
	<script language=vbscript>
		msgbox "<%=message%>",<%=sflag%>,"提示！！"
		location.href="<%=url%>"
	</script>
<%
	Response.end
	End sub

'-----------------------------------------------
'作者:------------------------------------------
'日期:20041123----------------------------------
'功能:一级一级的创建文件夹， 存在则不建立-------
'参数:文件夹路经--------------------------------
'版本:1.0---------------------------------------
'-----------------------------------------------
Function CreateFolderII(FolderPath)
	Dim FSO, FolderPath_Arr, i, FolderPath_Tmp, TheFolderPath
	
	FolderPath = Replace(FolderPath,"\","/")
	Set FSO = Server.CreateObject("Scripting.FileSystemObject")

	FolderPath_Arr = split(FolderPath,"/")

	For i = 0 to UBound(FolderPath_Arr)
		if i = 0 then
			FolderPath_Tmp = FolderPath_Arr(0) & "/"
		Else
			FolderPath_Tmp = FolderPath_Tmp & FolderPath_Arr(i) & "/"
		end if
		if i > 0 then
			TheFolderPath = left(FolderPath_Tmp, len(FolderPath_Tmp)-1)
			if not FSO.FolderExists(TheFolderPath) then FSO.CreateFolder TheFolderPath
		end if
	Next
	set FSO = nothing
	if err.number<>0 then
		CreateFolderII = false
		err.Clear
	else
		CreateFolderII = true
	end if
End Function

'-----------------------------------------------
'作者:------------------------------------------
'日期:20041123----------------------------------
'功能:组件对象是否存在---------------------------
'参数:组件对象---------------------------------
'版本:1.0---------------------------------------
'-----------------------------------------------
Function ObjInstall(ObjStr)
  on error resume next
  ObjInstall=false
  dim xtestobj
  err=0
  set xtestobj=server.createobject(ObjStr)
  if err=0 then ObjInstall=true
  set xtestobj=nothing
  err=0
End function

'-----------------------------------------------
'作者:Zerolone----------------------------------
'日期:20041224----------------------------------
'功能:去除非法字符---------------------------
'参数:字符串---------------------------------
'版本:1.0---------------------------------------
'-----------------------------------------------
Function CheckStr(str)
  if isnull(str) then
    CheckStr = ""
   exit function 
  end if
  CheckStr=replace(str,"'","''")
  CheckStr=replace(str," ","")
  CheckStr=replace(str,"<","&lt;")
  CheckStr=replace(str,">","&gt;")
  CheckStr=replace(str,";","；")
  CheckStr=replace(str, chr(34), "&quot;")
End Function

'-----------------------------------------------
'作者:Zerolone----------------------------------
'日期:20050219----------------------------------
'功能:调用相应的模版----------------------------
'参数:字符串------------------------------------
'版本:1.0---------------------------------------
'-----------------------------------------------
Function LoadTemplate(Template_Title)
 	Sql						= "Select [Template_Content] From [Monolith_Template] Where [Template_Title]='" & Template_Title & "'"
 	Rs.Open Sql, Conn, 1, 1
 	if Rs.Bof Or Rs.Eof then
		Response.write "后台配制错误， 请点击<a href=#>这里进入</a>本站简化版本！"
		Response.end
 	else
 		LoadTemplate	= Rs(0)
 	end if
 	Rs.Close
				
	'Sql执行次数+1
	SqlQueryNum	= SqlQueryNum + 1
End Function

'-----------------------------------------------
'作者:Zerolone----------------------------------
'日期:20050329----------------------------------
'功能:错误提示----------------------------------
'参数:字符串， 返回URL--------------------------
'版本:1.0---------------------------------------
'-----------------------------------------------
Sub errmsg(message,url)
%>
<script language=javascript>
	alert("<%=message%>");
	location.href="<%=url%>";
</script>
<%
Response.end
End Sub

'-----------------------------------------------
'作者:Zerolone----------------------------------
'日期:20050329----------------------------------
'功能:检查是否合法字符串------------------------
'参数:字符串， 返回URL--------------------------
'版本:1.0---------------------------------------
'-----------------------------------------------
Function Pass_Name(str)
	Pass_Name=true
	Dim Str_Arr
	Dim i
	Str_Arr		= "=%|?|&|;|,|'|chr(34)|.|chr(9)| |$|chr(255)|:|#|`|\|(|[|-|~"
	Str_Arr		= Split(Str_Arr, "|")
	For i=LBound(Str_Arr) to UBound(Str_Arr)
		if Instr(str, Str_Arr(i)) > 0 then Pass_Name=false
	Next
End Function











%>