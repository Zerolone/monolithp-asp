<%
Option Explicit

Response.Buffer = true
Response.ExpiresAbsolute=now()-1
Response.Expires=0
Response.CacheControl="no-cache"

dim Startime

Startime = Timer()

dim Endtime, SqlQueryNum, TheString
dim conn,rs,Sql, Connstr

SqlQueryNum	= 0

set conn 	= Server.CreateObject("ADODB.Connection")
set rs		= Server.CreateObject("ADODB.Recordset")

ConnStr		= "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.mappath("Data/MyBBS.mdb")
Conn.Open ConnStr
%>
<html>
<head>