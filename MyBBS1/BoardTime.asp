<%
  Endtime = Timer()
  TheString = "<hr>执行时间：" & FormatNumber((Endtime-Startime)*1000,5) & "毫秒。查询数据库" & SqlQueryNum & "次"
  Response.write TheString
%>