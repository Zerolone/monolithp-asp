<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include Virtual ="/Manage/Info/InfoSetting.asp" -->
<%
	Server.ScriptTimeout=999999

	Dim Info_ID, Info_Title, Info_Title2, Info_Class, Info_Author, Info_From, Info_KeyWord, Info_Memo, Info_Pic1, Info_Pic2, Info_FileName, Info_Content, Info_Date, Info_Template
	
	'��ȡ�ļ���
	Info_FileName	= Request("FileName")

	Info_FileName	= Right(Info_FileName, Len(Info_FileName)-InStrRev(Info_FileName, "/"))

	Response.write "��ǰ�ļ�������, �ļ��ؽ���...."
	

	'-------------------0--------------1-------------2-------------3-------------4--------------5-------------6------------7-----------8---------------9--------------10-------------11-------------12--------------13
	Sql = "Select [Info_Title], [Info_Title2], [Info_Class], [Info_Author], [Info_From], [Info_KeyWord], [Info_Memo], [Info_Pic1], [Info_Pic2], [Info_FileName], [Info_Date], [Info_Template], [Info_ID] From [Monolith_Info] Where [Info_FileName]='" & Info_FileName & "'"
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
		Info_ID				= Rs(12)
	else
		Response.write "�����Ų����ڣ�"
		Response.end
	end if
	Rs.Close

  if IsNull(Info_Title)		then	Info_Title		= "��"
  if IsNull(Info_Class) 	then	Info_Class		= "��"
  if IsNull(Info_Author)	then	Info_Author		= "��"
  if IsNull(Info_From)		then	Info_From			= "��"
  if IsNull(Info_Date)		then	Info_Date			= "��"
	
	if Info_Date = "" then
		Response.write "���ڳ���"
		Response.end
	end if

	if Info_Template = "" then
		Response.write "ģ�����"
		Response.end
	end if

  Sql = "Select [Template_Content] From [Monolith_Template] Where [Template_ID]=" & Info_Template
  Rs.Open Sql, Conn , 1, 1
  If Not(Rs.Bof Or Rs.Eof)then
		Info_Template = Rs(0)
  else
    Response.write "��ģ�治���ڣ�"
		Response.End
  end if
  rs.close
	
  '���±���	��{Info_Title}
  Info_Template = Replace(Info_Template,"{Info_Title}",			Info_Title)
  
  '�������ڣ�{Info_Date}
  Info_Template = Replace(Info_Template,"{Info_Date}",			Info_Date)
  
  '����	�� {Info_Author}
  Info_Template = Replace(Info_Template,"{Info_Author}",		Info_Author)
  
  '��Դ	��{Info_From}
  Info_Template = Replace(Info_Template,"{Info_From}",			Info_From)
  
  '����ID��{Info_ID}
  Info_Template = Replace(Info_Template,"{Info_ID}",				Info_ID)
  
  '����Class��{Info_Class}
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
		Response.write "�ļ�·��������쿴��Ϣ����ϵͳ����"
		Response.end
	end if

	if right(FileSavePath,1) <> "/" then FileSavePath=FileSavePath & "/" 
	if right(UrlPath,1) <> "/"			then UrlPath=UrlPath & "/" 	
  
	Dim tf, Fname, AllFileName, TheFileName, i , j, ThePage

	Fname							= Replace(Info_FileName, ".htm", "")

	TheFileName	= Fname & ".htm"
  '��Ϣȫ�ƣ�{Info_FileName}
  Info_Template = Replace(Info_Template,"{Info_FileName}", UrlPath & TheFileName)
 
	Set tf 			= MyFile.CreateTextFile(FileSavePath & TheFileName ,True)
	tf.Write (Info_Template)
  tf.Close
%>
<input type="button" value="�ļ��ؽ���ɣ� ��������, 3�����Զ���ת" name="B3" onclick="window.open('<%=Request("FileName")%>', '_parent');">
<script>
	window.setTimeout("TimeStar()", 3000);
	function TimeStar()
	{
		window.open("<%=Request("FileName")%>", "_parent");
	}
</script>
