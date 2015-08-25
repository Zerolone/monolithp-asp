<!-- #include virtual ="/Manage/Check.asp"--><%
	TheString = ""
	Dim Advertisement_Class_Name
	i = 1
	Set Cmd = Server.CreateObject("ADODB.Command")
	Set Cmd.ActiveConnection = Conn
	Cmd.CommandText	= "DigiChina_Advertisement_AD"
	Cmd.CommandType	= 4
	Cmd.Parameters.Append Cmd.CreateParameter("@Advertisement_Class_ID",  3, 1,4)
	Cmd.Parameters.Append Cmd.CreateParameter("@Advertisement_Count",  3, 1,4)
	Cmd("@Advertisement_Class_ID")	= Advertisement_Class_ID
	Cmd("@Advertisement_Count")			= Advertisement_Count	
	Set Rs = Cmd.ExeCute
	TheString = "<script language='javascript'>"
	Do While Not (Rs.Bof or Rs.Eof)
	  TheString = TheString & "Bar" & i & ".innerHTML=""" & RTrim(Rs(0)) & """;"
		Rs.MoveNext
		i = i + 1
	Loop
	TheString = TheString & "</script>"
	Rs.Close
	Set Cmd=Nothing	
%>