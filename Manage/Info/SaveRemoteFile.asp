<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include file ="NewsSetting.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Server.ScriptTimeOut = 999999

	Dim News_Content, News_Date, TheString
	Dim NewFileName, UrlPath, TheYear, TheMonth, TheDay, FSO, FileSavePath

	News_Date			= Request("News_Date")

	News_Content	= Request("News_Content")
	TheString			= News_Content
	News_Content	= Replace(News_Content, "<img" , "<IMG")
	News_Content	= Replace(News_Content, "src=" , "src=")

	News_Content	= Monolith_Catch(News_Content, 1, "<IMG", ">", "")
	News_Content	= Monolith_Catch(News_Content, 1, "http", """", "")
	News_Content	= Replace(News_Content, chr(34), "")

	Sql = ""

	TheYear		= Year(News_Date)
  TheMonth	= Month(News_Date)
  TheDay		= Day(News_Date)

  set FSO = server.CreateObject("Scripting.FileSystemObject")

  UrlPath		= NewsSettingArr(2) & "/" & TheYear & "/" & TheMonth & TheDay & "/"

	FileSavePath	= NewsSettingArr(0)
	
	Response.write UrlPath
	
	FileSavePath	= FileSavePath & TheYear
  FileSavePath	= FileSavePath	& "/" & TheMonth & TheDay & "/"
	if CreateFolderII(FileSavePath) = false then
		Response.write "<script language='javascript'>alert('�ļ�·��������쿴����ϵͳ����');</script>"
		Response.end
	end if

	if right(FileSavePath,1) <> "/" then FileSavePath=FileSavePath & "/" 
	if right(UrlPath,1) <> "/"			then UrlPath=UrlPath & "/" 

	TheString = ""

	Dim News_Content_Arr, i
	News_Content_Arr				= Split(News_Content, "|")
	For i = LBound(News_Content_Arr) to UBound(News_Content_Arr)
		SaveRemoteFile FileSavePath + GetFileName(News_Content_Arr(i)) , News_Content_Arr(i)
		Sql = "insert into [Monolith_UpPic] (UpPic_FileName) values ('" & UrlPath & GetFileName(News_Content_Arr(i)) & "')"
		Conn.execute(Sql)
'------------------------------------
'		TheString		= Replace(TheString, News_Content_Arr(i), UrlPath & GetFileName(News_Content_Arr(i)))
'		Response.write FileSavePath + GetFileName(News_Content_Arr(i)) &"<br>"& UrlPath & GetFileName(News_Content_Arr(i)) & "<hr>"
'------------------------------------
		TheString = "parent.frames.Monolith_Info_Body.document.body.innerHTML=parent.frames.Monolith_Info_Body.document.body.innerHTML.replace('" & News_Content_Arr(i) & "','" & UrlPath & GetFileName(News_Content_Arr(i)) & "');" & TheString
	Next
	
	'=================================
	'===========����ͼƬ==============
	'=================================
	Function Monolith_Catch(Source, iBegin, S1, S2, TheString)
		'Source								Դ�ļ�
		'iBegin								��ʼλ��
		'TheString						���ó����ַ���
		'S1										ģ��ǰ�ַ���
		'S2										ģ����ַ���

		Dim LenS1							'ģ��ǰ�ַ�������
		Dim iEnd							'ģ��ǰ�ַ�������λ��
		Dim jBegin						'ģ����ַ�����ʼλ��
		Dim jEnd							'ģ����ַ�������λ��
		Dim TempString				'��ʱ�ַ���

		LenS1			= Len(S1)
		iEnd			= InStr(iBegin, Source, S1)

		if iEnd		= 0 then
			Monolith_Catch	= TheString
		else
			jBegin		= iEnd + LenS1
			jEnd			= InStr(jBegin, Source, S2)

			TempString = S1 + Mid(Source, iEnd + LenS1, jEnd-jBegin+Len(S2))

			if TheString = "" then
				TheString	= TempString
			else
				TheString = TheString & "|" & TempString
			end if
'			Monolith_Catch	= TheString
			Monolith_Catch	= Monolith_Catch(Source, iEnd+ LenS1, S1, S2, TheString)
		end if
	End Function


 Function GetFileName(FullPath)
  If FullPath <> "" Then
   GetFileName = mid(FullPath,InStrRev(FullPath, "/")+1)
  Else
   GetFileName = ""
  End If
 End  Function
%>

<script>
	<%=TheString%>
	parent.LayerPicToRemote.style.visibility	=	"hidden";
	alert("ͼƬ���ػ��ɹ���");
</script>
