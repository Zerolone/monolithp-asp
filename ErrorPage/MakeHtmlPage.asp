<!-- #include file ="../Inc/Monolith_Conn.asp" -->
<!-- #include file ="../Inc/Monolith_Function.asp" -->
<!-- #include file ="../Manage/News/NewsSetting.asp" -->
<%
	Server.ScriptTimeout=999999

	Dim News_ID, News_Title, News_Title2, News_Class, News_Author, News_From, News_KeyWord, News_Memo, News_Pic1, News_Pic2, News_FileName, News_Content, News_Date, News_Template
	
	'��ȡ�ļ���
	News_FileName	= Request("FileName")

	if InStr(News_FileName, "Info") > 0 then
		Response.Redirect "MakeHtmlPage_Info.asp?FileName=" & News_FileName
		Response.end
	end if

	if InStr(News_FileName, "News") = 0 then 
		Response.write "����ʵ��ļ������ڣ�"
		Response.write "<a href='#'>����</a>"
		Response.end
	end if

	News_FileName	= Right(News_FileName, Len(News_FileName)-InStrRev(News_FileName, "/"))

	Response.write "��ǰ�ļ�������, �ļ��ؽ���...."
	
'	Response.write News_FileName

	'-------------------0--------------1-------------2-------------3-------------4--------------5-------------6------------7-----------8---------------9--------------10-------------11-------------12--------------13
	Sql = "Select [News_Title], [News_Title2], [News_Class], [News_Author], [News_From], [News_KeyWord], [News_Memo], [News_Pic1], [News_Pic2], [News_FileName], [News_Content], [News_Date], [News_Template], [News_ID] From [Monolith_News] Where [News_FileName]='" & News_FileName & "'"
	Rs.open Sql, conn ,1 ,1
	if Not(Rs.Bof Or Rs.Eof) then
		News_Title		= Rs(0)
		News_Title2		= Rs(1)
		News_Class		= Rs(2)
		News_Author		= Rs(3)
		News_From			= Rs(4)
		News_KeyWord	= Rs(5)
		News_Memo			= Rs(6)
		News_Pic1			= Rs(7)
		News_Pic2			= Rs(8)
		News_FileName	= Rs(9)
		News_Content	= Rs(10)
		News_Date			= Rs(11)
		News_Template	= Rs(12)
		News_ID				= Rs(13)
	else
		Response.write "�����Ų����ڣ�"
		Response.end
	end if
	Rs.Close

  if IsNull(News_Title)		then	News_Title		= "��"
  if IsNull(News_Class) 	then	News_Class		= "��"
  if IsNull(News_Author)	then	News_Author		= "��"
  if IsNull(News_From)		then	News_From			= "��"
  if IsNull(News_Content) then	News_Content	= "��"
  if IsNull(News_Date)		then	News_Date			= "��"
	
	if News_Date = "" then
		Response.write "���ڳ���"
		Response.end
	end if

	if News_Template = "" then
		Response.write "ģ�����"
		Response.end
	end if

  Sql = "Select [Template_Content] From [Monolith_Template] Where [Template_ID]=" & News_Template
  Rs.Open Sql, Conn , 1, 1
  If Not(Rs.Bof Or Rs.Eof)then
		News_Template = Rs(0)
  else
    Response.write "��ģ�治���ڣ�"
		Response.End
  end if
  rs.close
	
  '���±���	��{News_Title}
  News_Template = Replace(News_Template,"{News_Title}",			News_Title)
  
  '�������ڣ�{News_Date}
  News_Template = Replace(News_Template,"{News_Date}",			News_Date)
  
  '����	�� {News_Author}
  News_Template = Replace(News_Template,"{News_Author}",		News_Author)
  
  '��Դ	��{News_From}
  News_Template = Replace(News_Template,"{News_From}",			News_From)
  
  '����ID��{News_ID}
  News_Template = Replace(News_Template,"{News_ID}",				News_ID)
  
  '----------------

	Dim MyFile, FileSavePath, UrlPath, TheYear, TheMonth, TheDay, FSO

  Set MyFile=Server.CreateObject("Scripting.FileSystemObject")
  set FSO = server.CreateObject("Scripting.FileSystemObject")
  
	TheYear		= Year(News_Date)
  TheMonth	= Month(News_Date)
  TheDay		= Day(News_Date)

	FileSavePath	= NewsSettingArr(6)
	UrlPath				= NewsSettingArr(7)
	if right(UrlPath,1) <> "/"			then UrlPath=UrlPath & "/" 
	if right(FileSavePath,1) <> "/" then FileSavePath=FileSavePath & "/" 

	UrlPath				= UrlPath & TheYear & "/" & TheMonth & TheDay & "/"

	FileSavePath	= FileSavePath & TheYear
  FileSavePath	= FileSavePath	& "/" & TheMonth & TheDay & "/"

	if CreateFolderII(FileSavePath) = false then
		Response.write "�ļ�·��������쿴����ϵͳ����"
		Response.end
	end if

	if right(FileSavePath,1) <> "/" then FileSavePath=FileSavePath & "/" 
	if right(UrlPath,1) <> "/"			then UrlPath=UrlPath & "/" 	
  
	Dim News_Content_Arr, MaxPages, tf, TheContent, TheText, Fname, AllFileName, TheFileName, i , j, ThePage

  News_Content_Arr	=	Split(News_Content,"[---��ҳ---]")
  MaxPages					=	UBound(News_Content_Arr)

	Fname							= Replace(News_FileName, ".htm", "")

  if MaxPages = 0 then
  '---------Singer Page
		TheFileName	= Fname & ".htm"
	  '����ȫ�ƣ�{News_FileName}
	  News_Template = Replace(News_Template,"{News_FileName}", UrlPath & TheFileName)

    Set tf 			= MyFile.CreateTextFile(FileSavePath & TheFileName ,True)
    TheContent 	= Replace(News_Template,"{News_Content}",News_Content)
		TheText 		= Replace(TheContent,"{��ҳ}","")
		tf.Write (TheText)
    tf.Close
	  AllFileName = AllFileName & "<li><a target='_blank' href='" & UrlPath & TheFilename & "'>" & TheFilename & "</a>"
	'----------Singer Page End
  else

	'----------Not Singer Page
    for i = 0 to MaxPages
      TheContent = Replace(News_Template,"{News_Content}" ,News_Content_Arr(i))
		  '-------------Page name
			if i = 0 then
		    TheFileName	= Fname & ".htm"
		  else
				TheFileName 	= Fname & "_" & i + 1 & ".htm"
		  end if
	  
		  '����ȫ�ƣ�{News_FileName}
		  TheContent = Replace(TheContent,"{News_FileName}", UrlPath & TheFileName)

	  '--------------��һҳ
	  if i = 1 then
		ThePage		= "<a class=black href='" & Fname & ".htm'>��һҳ</a>&nbsp;"
	  elseif i >1 then
		ThePage		= "<a class=black href='" & Fname & "_" & i & ".htm'>��һҳ</a>&nbsp;"
	  end if

	  '--------@#$&^*)
	  For j = 0 to MaxPages
			if j <> i Then
			  if j = 0 then
			    ThePage	= ThePage & "<a class=black href='" & Fname & ".htm'>[" & j + 1  & "]</a>&nbsp;"
				else
			    ThePage	= ThePage & "<a class=black href='" & Fname & "_" & j + 1 & ".htm'>[" & j + 1  & "]</a>&nbsp;"
			  end if
			else
				ThePage	= ThePage & "<font color='#FF0000'><B>["& j + 1 &"]</B></font>&nbsp;"
			end if
	  Next

	  '--------------��һҳ
	  if i < Maxpages then ThePage		= ThePage & "<a class=black href='" & Fname & "_" & i + 2 & ".htm'>��һҳ</a>"
		
	  Set tf = MyFile.CreateTextFile(FileSavePath & TheFileName ,True)
      TheText = Replace(TheContent,"{��ҳ}",ThePage)
	  
	  tf.Write (TheText)
	  tf.Close
	  AllFileName = AllFileName & "<li><a target='_blank' href='" & UrlPath & TheFilename & "'>" & TheFilename & "</a>"
    next
  end if 
%>
<input type="button" value="�ļ��ؽ���ɣ� ��������, 3�����Զ���ת" name="B3" onclick="window.open('<%=Request("FileName")%>', '_parent');">
<script>
	window.setTimeout("TimeStar()", 3000);
	function TimeStar()
	{
		window.open("<%=Request("FileName")%>", "_parent");
	}
</script>
