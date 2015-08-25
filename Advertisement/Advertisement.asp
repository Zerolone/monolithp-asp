<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<%
	Dim TheString, i, Cmd, Advertisement_Class_Name, Advertisement_Class_ID, Advertisement_Count
	TheString = ""
	i = 1
	Advertisement_Class_ID		= Request("Advertisement_Class_ID")
	Advertisement_Count				= Request("Advertisement_Count")
	Set Cmd = Server.CreateObject("ADODB.Command")
	Set Cmd.ActiveConnection = Conn
	Cmd.CommandText	= "DigiChina_Advertisement_AD"
	Cmd.CommandType	= 4
	Cmd.Parameters.Append Cmd.CreateParameter("@Advertisement_Class_ID",  3, 1,4)
	Cmd.Parameters.Append Cmd.CreateParameter("@Advertisement_Count",  3, 1,4)
	Cmd("@Advertisement_Class_ID")	= Advertisement_Class_ID
	Cmd("@Advertisement_Count")		= Advertisement_Count	
	Set Rs = Cmd.ExeCute
	Do While Not (Rs.Bof or Rs.Eof)
	  TheString = TheString & "Bar" & i & ".innerHTML=""" & RTrim(Rs(0)) & """;"
		Rs.MoveNext
		i = i + 1
	Loop
	Rs.Close
	Set Cmd=Nothing	
%>
<%=TheString%>