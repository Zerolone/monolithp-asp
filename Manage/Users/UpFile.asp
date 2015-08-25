<!-- #include virtual ="/Inc/Monolith_Function.asp"-->
<!-- #include virtual ="/Inc/UpLoadClass.asp"-->
<!-- #include virtual ="/Manage/Check.asp"-->
<html>
<head>
<title>文件上传</title>
</head>

<%
	dim request2,formPath,formName,intTemp
	Dim NewFileName, TheString, UrlPath, TheYear, TheMonth, TheDay, FSO, FileSavePath
	
	UrlPath	= "/Advertisement/Src/"
	UrlPath	= Server.MapPath(UrlPath)
	CreateFolderII(UrlPath)


	set request2=new UpLoadClass
	request2.FileType	= "gif/jpg/swf"
	request2.SavePath = UrlPath & "/"
	request2.AutoSave	= 0
	request2.open() 

' 	response.write "<FIELDSET align=center><LEGEND align=center>文件上传发生错误 </LEGEND><br>请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</fieldset>"
	
	for intTemp=1 to Ubound(request2.FileItem)
		'获取表单文件控件名称，注意FileItem下标从1开始
		formName=request2.FileItem(intTemp)
		NewFileName	= request2.form(formName)

		'显示文件保存状态
'		select case request2.form(formName&"_Err")
'			case -1:
'				response.write "没有文件上传"
'			case 0:
'				response.write "上传文件成功"
'			case 1:
'				response.write "文件太大，拒绝上传"
'			case 2:
'				response.write "文件格式不对，拒绝上传"
'			case 3:
'				response.write "文件太大且格式不对，拒绝上传"
'		end select
	next

' response.write "<center><FIELDSET align=center><LEGEND align=center><font color=red>文件上传成功 </font></LEGEND><br>[ <a href=# onclick=""Addpic('"&NewFileName&"')"">点击这里添加到编辑器中</a> ]</fieldset>"
' response.write "<center><FIELDSET align=center><LEGEND align=center><font color=red>文件上传成功 </font></LEGEND><br>[ <a href=# onclick=""Addpic('" & Replace(NewFileName,UrlPath,"")  & "')"">点击这里添加到编辑器中</a> ]</fieldset>"
%>
<script  language="JavaScript">
	parent.document.form1.Advertisement_FileName.value = "<%=NewFileName%>";
	parent.document.form1.Submit_AddFileName.click();
</script>

</body>
</html>