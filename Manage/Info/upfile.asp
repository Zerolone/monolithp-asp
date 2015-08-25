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


	'读取Request
	'response.Write("<br>邮件标题："&request2.Form("strTitle"))
	
	for intTemp=1 to Ubound(request2.FileItem)
		'获取表单文件控件名称，注意FileItem下标从1开始
		formName=request2.FileItem(intTemp)
		
		NewFileName	= request2.form(formName)

		'显示文件保存状态
		select case request2.form(formName&"_Err")
'			case -1:
'				response.write "没有文件上传"
			case 0:
'				response.write "上传文件成功"
				TheString = TheString & "<img src='" & UrlPath & NewFileName & "'>"
				Sql = "insert into [Monolith_UpPic] (UpPic_FileName) values ('" & UrlPath & NewFileName & "')"
				Conn.Execute(Sql)
'			case 1:
'				response.write "文件太大，拒绝上传"
'			case 2:
'				response.write "文件格式不对，拒绝上传"
'			case 3:
'				response.write "文件太大且格式不对，拒绝上传"
		end select

	next
	Call CloseConn
	set request2=nothing
	
	Response.write "<script>"
	if TheString = "" then
		Response.write "alert('文件全部没有上传！');"
	else
		Response.write "alert('文件全部上传成功！');"
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
