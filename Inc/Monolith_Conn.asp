<%
  Option Explicit

  Response.Buffer = true
  Response.ExpiresAbsolute=now()-1
  Response.Expires=0
  Response.CacheControl="no-cache"

	'起始时间  
	Dim Startime
	Startime = Timer()
  
  'Sql执行次数
  Dim SqlQueryNum
  SqlQueryNum	= 0

	'数据库联接语句
  Dim DBPath, Rs, Conn, Sql, ConnStr
  DBPath	= Server.MapPath("\") & "\Data\Monolith.mdb"
'	Response.write DBPath

  Set Rs	= Server.CreateObject("ADODB.Recordset")
  Set Conn	= Server.CreateObject("ADODB.Connection")
  ConnStr	= "provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath
  Conn.Open ConnStr
	
	'全局CookiesName
  Dim Monolith_CookiesName
  Monolith_CookiesName	= "Monolith"

	'关闭Conn
	Sub CloseConn
		Conn.Close
		Set Conn = nothing
	End Sub

	'关闭Rs
	Sub CloseRs
		Rs.Close
		Set Rs = nothing
	End Sub
%>