<!-- #include virtual ="/Inc/Monolith_Function.asp"-->
<!-- #include virtual ="/Inc/UpLoadClass.asp"-->
<!-- #include virtual ="/Manage/Check.asp"-->
<html>
<head>
<title>�ļ��ϴ�</title>
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

' 	response.write "<FIELDSET align=center><LEGEND align=center>�ļ��ϴ��������� </LEGEND><br>����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</fieldset>"
	
	for intTemp=1 to Ubound(request2.FileItem)
		'��ȡ���ļ��ؼ����ƣ�ע��FileItem�±��1��ʼ
		formName=request2.FileItem(intTemp)
		NewFileName	= request2.form(formName)

		'��ʾ�ļ�����״̬
'		select case request2.form(formName&"_Err")
'			case -1:
'				response.write "û���ļ��ϴ�"
'			case 0:
'				response.write "�ϴ��ļ��ɹ�"
'			case 1:
'				response.write "�ļ�̫�󣬾ܾ��ϴ�"
'			case 2:
'				response.write "�ļ���ʽ���ԣ��ܾ��ϴ�"
'			case 3:
'				response.write "�ļ�̫���Ҹ�ʽ���ԣ��ܾ��ϴ�"
'		end select
	next

' response.write "<center><FIELDSET align=center><LEGEND align=center><font color=red>�ļ��ϴ��ɹ� </font></LEGEND><br>[ <a href=# onclick=""Addpic('"&NewFileName&"')"">���������ӵ��༭����</a> ]</fieldset>"
' response.write "<center><FIELDSET align=center><LEGEND align=center><font color=red>�ļ��ϴ��ɹ� </font></LEGEND><br>[ <a href=# onclick=""Addpic('" & Replace(NewFileName,UrlPath,"")  & "')"">���������ӵ��༭����</a> ]</fieldset>"
%>
<script  language="JavaScript">
	parent.document.form1.Advertisement_FileName.value = "<%=NewFileName%>";
	parent.document.form1.Submit_AddFileName.click();
</script>

</body>
</html>