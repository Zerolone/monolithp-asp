<%	const iPage=30

   function strFilter(str)
		str=Replace(str,"'","''")
		str=replace(str,"|","/")
		strFilter=str
   end function

   function AddCRLF(str,total)
		dim num,pos,i
		pos=1
		num=0
		while (pos>0 and num<total)
			pos=pos+1
			pos=instr(pos,str,chr(13))
			if pos>0 then
				num=num+1
			end if 
		wend
		for i=1 to total-num
			str=str+chr(13)+chr(10)
		next 
		AddCRLF=str
   end function
   
   function ShowBody(str)
		dim Spos,Epos,strURL,tmpStr,tmpPos
		Spos=1

		str=Replace(str,chr(34),"&quot;")
		str=Replace(str,"<","&lt")
		str=Replace(str,">","&gt")
		str=Replace(str,chr(13)+chr(10),"<br>")
		str=Replace(str,chr(9),"&nbsp;&nbsp;&nbsp;&nbsp;")
		Spos=instr(Spos,str," http://")
		if sPos>0 then 
			str=Replace(str," ","…")
			while (Spos>0)
				Spos=instr(Spos,str,"…http://")
				if Spos>0 then
				   Epos=instr(Spos+1,str,"…")
				   if Epos>0 then
					  strURL=mid(str,Spos+1,Epos-Spos-1)
					  strURL=lcase(strURL)
					  if instr(strURL,"gif")>0 or instr(strURL,"jpg")>0 then
						strURL="<img src='"+strURL+"' >"
					  else
						strURL="<a class=showdoc target=_blank href="+chr(34)+strURL+chr(34)+" >"+strURL+"</a>"
					  end if
	   				  tmpStr=left(str,Spos-1)+strURL+right(str,len(str)-Epos)
					  str=tmpstr	

				   end if
				   spos=spos+5
				end if
			wend
			str=Replace(str,"…","&nbsp;")
		else
     		str=Replace(str," ","&nbsp;")
		end if
		showbody=str
   end function
%>

