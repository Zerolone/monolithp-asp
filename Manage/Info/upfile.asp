<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/UpLoadClass.asp"-->
<!-- #include file ="NewsSetting.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<%
	Server.ScriptTimeOut=5000
	dim request2,formPath,formName,intTemp
	Dim NewFileName, TheString, UrlPath, TheYear, TheMonth, TheDay, FSO, FileSavePath

	Sql = ""
  TheYear		= Year(Now)
  TheMonth	= Month(Now)
  TheDay		= Day(Now)

  set FSO = server.CreateObject("Scripting.FileSystemObject")

  UrlPath		= NewsSettingArr(2) & "/" & TheYear & "/" & TheMonth & TheDay & "/"

	FileSavePath	= NewsSettingArr(0)
	if not FSO.FolderExists(FileSavePath) then FSO.CreateFolder(FileSavePath)
	FileSavePath	= FileSavePath & TheYear
  if not FSO.FolderExists(FileSavePath) then FSO.CreateFolder(FileSavePath)
  FileSavePath	= FileSavePath	& "/" & TheMonth & TheDay & "/"
  if not FSO.FolderExists(FileSavePath) then FSO.CreateFolder(FileSavePath)

	set request2=new UpLoadClass
	request2.FileType	= NewsSettingArr(3)
	request2.SavePath = FileSavePath
	request2.AutoSave	= NewsSettingArr(4)
	request2.open() 


	'��ȡRequest
	'response.Write("<br>�ʼ����⣺"&request2.Form("strTitle"))
	
	for intTemp=1 to Ubound(request2.FileItem)
		'��ȡ���ļ��ؼ����ƣ�ע��FileItem�±��1��ʼ
		formName=request2.FileItem(intTemp)
		
		NewFileName	= request2.form(formName)

		'��ʾ�ļ�����״̬
		select case request2.form(formName&"_Err")
'			case -1:
'				response.write "û���ļ��ϴ�"
			case 0:
'				response.write "�ϴ��ļ��ɹ�"
				TheString = TheString & "<img src='" & UrlPath & NewFileName & "'>"
				Sql = "insert into [Monolith_UpPic] (UpPic_FileName) values ('" & UrlPath & NewFileName & "')"
				Conn.Execute(Sql)
'			case 1:
'				response.write "�ļ�̫�󣬾ܾ��ϴ�"
'			case 2:
'				response.write "�ļ���ʽ���ԣ��ܾ��ϴ�"
'			case 3:
'				response.write "�ļ�̫���Ҹ�ʽ���ԣ��ܾ��ϴ�"
		end select

	next
	Call CloseConn
	set request2=nothing
	
	Response.write "<script>"
	if TheString = "" then
		Response.write "alert('�ļ�ȫ��û���ϴ���');"
	else
		Response.write "alert('�ļ�ȫ���ϴ��ɹ���');"
	end if
	Response.write "</script>"
%>
<script>
var str;
str="";
str= "<%=TheString%>";
parent.document.form1.UpLoadPic.value=str;
parent.document.form1.InsertUpLoadPic.style.display = "";
</script>
