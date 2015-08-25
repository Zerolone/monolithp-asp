<%
'-----------------------------------------------
'����:Zerolone ---------------------------------
'����:20040923----------------------------------
'����:���ID�Ƿ�Ϊ����--------------------------
'����:integer-----------------------------------
'-----------------------------------------------
Function CheckNum(ID)
	if not IsNumeric(ID) then
		response.write "<script>alert('�Ƿ�����');history.back()</script>"
		response.end
	end if
End Function

'-----------------------------------------------
'����:Zerolone ---------------------------------
'����:20041121----------------------------------
'����:���ID�Ƿ�Ϊ���Σ� ����T/F----------------
'����:integer-----------------------------------
'-----------------------------------------------
Function CheckNumII(ID)
	if not IsNumeric(ID) then
		CheckNumII = false
	else
		CheckNumII = true
	end if
End Function


'-----------------------------------------------
'����:------------------------------------------
'����:------------------------------------------
'����:���������ʮλ������----------------------
'����:------------------------------------------
'-----------------------------------------------
Function MakeRndNum()
	Dim fname
	fname = now()
	fname = replace(fname,"-","")
	fname = replace(fname," ","") 
	fname = replace(fname,":","")
	fname = replace(fname,"PM","")
	fname = replace(fname,"AM","")
	fname = replace(fname,"����","")
	fname = replace(fname,"����","")
	fname = int(fname) + int((10-1+1)*Rnd + 1)
	MakeRndNum=fname
End Function 

'-----------------------------------------------
'����:Zerolone ---------------------------------
'����:20040923----------------------------------
'����:ȥ��������ַ�, ���ݲ����Ƿ���Ҫ��ʾʡ�Ժ�
'����:String, integer, 1/0----------------------
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
'����:Zerolone ---------------------------------
'����:20040924----------------------------------
'����:�����ַ�����------------------------------
'����:String------------------------------------
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
'����:Zerolone ---------------------------------
'����:20040923----------------------------------
'����:����HTML�﷨------------------------------
'����:String------------------------------------
'-----------------------------------------------
Function MonolithEncode(TheString)
	TheString				= Replace(TheString, chr(34), "&quot;")
	TheString				= Replace(TheString, "'", "��")
	TheString				= Replace(TheString, "<", "&lt;")
	TheString				= Replace(TheString, ">", "&gt;")
'	TheString				= Replace(TheString, ";", "��")
	TheString				= Replace(TheString, "-", "��")
	TheString				= Replace(TheString, Chr(10), "<BR>")
	TheString				= Replace(TheString, " ", "&nbsp;")
	MonolithEncode	= TheString
End Function

'-----------------------------------------------
'����:Zerolone ---------------------------------
'����:20040923----------------------------------
'����:��HTML����д��һ��------------------------
'����:String------------------------------------
'-----------------------------------------------
Function MonolithEncodeII(TheString)
//	TheString				= Replace(TheString, chr(34), "\&quot;")
	TheString				= Replace(TheString, chr(34), "'")
	TheString				= Replace(TheString, Chr(10), "")
	TheString				= Replace(TheString, Chr(13), "")
	MonolithEncodeII	= TheString
End Function


'-----------------------------------------------
'����:Zerolone ---------------------------------
'����:20050220----------------------------------
'����:��ԭHTML�﷨------------------------------
'����:String------------------------------------
'-----------------------------------------------
Function MonolithDecode(TheString)
	TheString				= Replace(TheString, "&lt;", "<")
	TheString				= Replace(TheString, "&gt;", ">")
	TheString				= Replace(TheString, "<BR>", Chr(10))
	TheString				= Replace(TheString, "&nbsp;", " ")
	MonolithDecode	= TheString
End Function

'-----------------------------------------------
'����:Zerolone ---------------------------------
'����:20050110----------------------------------
'����:����HTML�﷨------------------------------
'����:String------------------------------------
'-----------------------------------------------
Function Monolith_Encode(TheString)
  if IsNull(TheString) then
    TheString = ""
   Exit Function 
  end if

	TheString				= Replace(TheString, chr(34), "&quot;")
	TheString				= Replace(TheString, "'", "��")
	TheString				= Replace(TheString, "<", "&lt;")
	TheString				= Replace(TheString, ">", "&gt;")
	TheString				= Replace(TheString, ";", "��")
	TheString				= Replace(TheString, "-", "��")
	TheString				= Replace(TheString, Chr(13)&Chr(10), "<br>")
	Monolith_Encode	= TheString
End Function

'-----------------------------------------------
'����:Zerolone ---------------------------------
'����:20050328----------------------------------
'����:��ԭHTML�﷨------------------------------
'����:String------------------------------------
'-----------------------------------------------
Function Monolith_Decode(TheString)
	TheString				= Replace(TheString, "&quot;",	chr(34))
	TheString				= Replace(TheString, "��", 			"'")
	TheString				= Replace(TheString, "&lt;", 		"<")
	TheString				= Replace(TheString, "&gt;", 		">")
	TheString				= Replace(TheString, "��", 			";")
	TheString				= Replace(TheString, "��", 			"-")
	TheString				= Replace(TheString, "<br>", 		Chr(13)&Chr(10))
	Monolith_Decode	= TheString
End Function



'-----------------------------------------------
'����:Zerolone ---------------------------------
'����:20040923----------------------------------
'����:���ݲ������пո�--------------------------
'����:integer-----------------------------------
'-----------------------------------------------
Sub LoopNBSP(num)
	Dim i
	for i =0 to num
		Response.write "&nbsp;"
	next
End Sub

'-----------------------------------------------
'����:------------------------------------------
'����:20041008----------------------------------
'����:��ȡ�ļ�, ������ļ��Ƿ����--------------
'����:String, �ļ���----------------------------
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
'����:------------------------------------------
'����:20041008----------------------------------
'����:�޸��ļ�ָ������, ������ļ��Ƿ����------
'����:String,String,String �ļ���,Ŀ��,�滻����-
'�汾:1.1---------------------------------------
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
'�汾:1.1---------------------------------------
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
'����:------------------------------------------
'����:20041013----------------------------------
'����:Զ�̴洢�ļ�------------------------------
'����:�����ļ����� Զ�̵�ַ---------------------
'�汾:1.0---------------------------------------
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
'����:------------------------------------------
'����:20041121----------------------------------
'����:������ʾ----------------------------------
'����:ͦʾ���ݣ� ͼ�꣬ ��ת��ַ----------------
'�汾:1.0---------------------------------------
'-----------------------------------------------
	Sub ErrMsg(message,sflag,url)
%>
	<script language=vbscript>
		msgbox "<%=message%>",<%=sflag%>,"��ʾ����"
		location.href="<%=url%>"
	</script>
<%
	Response.end
	End sub

'-----------------------------------------------
'����:------------------------------------------
'����:20041123----------------------------------
'����:һ��һ���Ĵ����ļ��У� �����򲻽���-------
'����:�ļ���·��--------------------------------
'�汾:1.0---------------------------------------
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
'����:------------------------------------------
'����:20041123----------------------------------
'����:��������Ƿ����---------------------------
'����:�������---------------------------------
'�汾:1.0---------------------------------------
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
'����:Zerolone----------------------------------
'����:20041224----------------------------------
'����:ȥ���Ƿ��ַ�---------------------------
'����:�ַ���---------------------------------
'�汾:1.0---------------------------------------
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
  CheckStr=replace(str,";","��")
  CheckStr=replace(str, chr(34), "&quot;")
End Function

'-----------------------------------------------
'����:Zerolone----------------------------------
'����:20050219----------------------------------
'����:������Ӧ��ģ��----------------------------
'����:�ַ���------------------------------------
'�汾:1.0---------------------------------------
'-----------------------------------------------
Function LoadTemplate(Template_Title)
 	Sql						= "Select [Template_Content] From [Monolith_Template] Where [Template_Title]='" & Template_Title & "'"
 	Rs.Open Sql, Conn, 1, 1
 	if Rs.Bof Or Rs.Eof then
		Response.write "��̨���ƴ��� ����<a href=#>�������</a>��վ�򻯰汾��"
		Response.end
 	else
 		LoadTemplate	= Rs(0)
 	end if
 	Rs.Close
				
	'Sqlִ�д���+1
	SqlQueryNum	= SqlQueryNum + 1
End Function

'-----------------------------------------------
'����:Zerolone----------------------------------
'����:20050329----------------------------------
'����:������ʾ----------------------------------
'����:�ַ����� ����URL--------------------------
'�汾:1.0---------------------------------------
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
'����:Zerolone----------------------------------
'����:20050329----------------------------------
'����:����Ƿ�Ϸ��ַ���------------------------
'����:�ַ����� ����URL--------------------------
'�汾:1.0---------------------------------------
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