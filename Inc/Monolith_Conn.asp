<%
  Option Explicit

  Response.Buffer = true
  Response.ExpiresAbsolute=now()-1
  Response.Expires=0
  Response.CacheControl="no-cache"

	'��ʼʱ��  
	Dim Startime
	Startime = Timer()
  
  'Sqlִ�д���
  Dim SqlQueryNum
  SqlQueryNum	= 0

	'���ݿ��������
  Dim DBPath, Rs, Conn, Sql, ConnStr
  DBPath	= Server.MapPath("\") & "\Data\Monolith.mdb"
'	Response.write DBPath

  Set Rs	= Server.CreateObject("ADODB.Recordset")
  Set Conn	= Server.CreateObject("ADODB.Connection")
  ConnStr	= "provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath
  Conn.Open ConnStr
	
	'ȫ��CookiesName
  Dim Monolith_CookiesName
  Monolith_CookiesName	= "Monolith"

	'�ر�Conn
	Sub CloseConn
		Conn.Close
		Set Conn = nothing
	End Sub

	'�ر�Rs
	Sub CloseRs
		Rs.Close
		Set Rs = nothing
	End Sub
%>