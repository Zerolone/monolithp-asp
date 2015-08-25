<html> 
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript> 
<!-- 
function MakeColor(ThisColor) { 
//document.bgColor = ThisColor; 
AAA.BBB.value = ThisColor;
} 

function MakeColor1(ThisColor) { 
CCC.bgColor = ThisColor;
} 

//--> 
</SCRIPT> 
<center> 
<table cellspacing=2 Border="0"> 
	<tr> 
	<%
		Dim I1
		Dim I2
		Dim I3
		For I1 = 0 to 15 step 3
			For I2 = 0 to 15 step 3 
				For I3 = 0 to 15 step 3 
					Color = Hex(I1) & Hex(I1) & Hex(I2) & Hex(I2) & Hex(I3) & Hex(I3)
	%> 
		<td bgcolor="#<%=Color%>" OnClick="MakeColor('#<%=Color%>');" Onmousemove="MakeColor1('#<%=Color%>');">¡¡&nbsp;</td> 
	<% 
				Next 
			Next 
	%> 
	</tr> 
	<tr> 
	<% 
		Next 
	%> 
	</tr> 
</table>
<table id="CCC"><tr><td>&nbsp;asdf</td></tr></table>
<form name="AAA">
  <input name="BBB">
</form>
</center> 
</html>