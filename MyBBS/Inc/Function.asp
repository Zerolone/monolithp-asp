<%
function RandomChar()
  dim ranNum
  randomize
  ranNum=int(26*rnd)+97
  RandomChar = chr(ranNum)
  ranNum=int(26*rnd)+97
  RandomChar = RandomChar & chr(ranNum)
end function

function CheckStr(str)
  if isnull(str) then
    CheckStr = ""
   exit function 
  end if
  CheckStr=replace(str,"'","''")
  CheckStr=replace(str," ","")
  CheckStr=replace(str,"<","&lt;")
  CheckStr=replace(str,">","&gt;")
  CheckStr=replace(str,";","ฃป")
  CheckStr=replace(str, chr(34), "&quot;")
end function
	
function ReplaceNull(str)
  if IsNull(str) then	str = "กก"
end function
%>