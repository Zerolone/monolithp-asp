<script language="javascript">
function tb_addnew(Table, Str1, Str2, Str3, Str4, Str5)
{
//var ls_t=document.all("mytable")

var ls_t=document.all(Table);

maxcell=ls_t.rows(0).cells.length;
mynewrow = ls_t.insertRow();
//    for(i=0;i<maxcell;i++)
//    {
    mynewcell=mynewrow.insertCell();
//    mynewcell.innerText= Str1;
    mynewcell.innerHTML= Str1;
    mynewcell=mynewrow.insertCell();
    mynewcell.innerText= Str2;
    mynewcell=mynewrow.insertCell();
    mynewcell.innerText= Str3;

    mynewcell=mynewrow.insertCell();
    mynewcell.innerText= Str4;
    mynewcell=mynewrow.insertCell();
    mynewcell.innerText= Str5;

//    }
}
</script>
<table width="100%" border="1" bordercolor="#FF0000">
  <tr> 
    <td>头</td>
  </tr>
  <tr> 
    <td><table width="100%" border="1" bordercolor="#0000FF" id=mytable>
        <tr> 
          <td>1</td>
          <td>2</td>
          <td>3</td>
          <td>4</td>
          <td>5</td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td>尾</td>
  </tr>
</table>
<p>
  <input type="button" name="Submit" value=" 添 加 一 列 " onClick="tb_addnew('mytable','<table border=1><tr><td>AAA</td></tr><tr><td>BBB</td></tr></table>','b','c','e','d')">
  <input type="button" name="Submit" value=" 添 加 一 列 " onClick="tb_addnew('mytable','<table border=1><tr><td>CCC</td></tr><tr><td>DDD</td></tr></table>','b','c','e','d')">
</p>
<table width="100%" border="1" bordercolor="#FF0000">
  <tr> 
    <td>头</td>
  </tr>
  <tr> 
    <td><table width="100%" border="1" bordercolor="#0000FF" id=mytable1>
        <tr> 
          <td><table width="100%" border="1">
              <tr> 
                <td>1</td>
              </tr>
              <tr> 
                <td>2</td>
              </tr>
              <tr> 
                <td>3</td>
              </tr>
              <tr> 
                <td>4</td>
              </tr>
              <tr> 
                <td>5</td>
              </tr>
            </table></td>
          <td><table width="100%" border="1">
              <tr> 
                <td>1</td>
              </tr>
              <tr> 
                <td>2</td>
              </tr>
              <tr> 
                <td>3</td>
              </tr>
              <tr> 
                <td>4</td>
              </tr>
              <tr> 
                <td>5</td>
              </tr>
            </table></td>			
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td>尾</td>
  </tr>
</table>
<p>
  <input type="button" name="Submit" value=" 添 加 一 列 " onClick="tb_addnew1('mytable1','a','b','c','e','d')">
  <input type="button" name="Submit" value=" 添 加 一 列 " onClick="tb_addnew1('mytable1','a','b','c','e','d')">
</p>
<script language="javascript">
function tb_addnew1(Table, Str1, Str2, Str3, Str4, Str5)
{
//var ls_t=document.all("mytable")

var ls_t=document.all(Table);

//maxcell=ls_t.rows(0).cells.length;
//mynewrow = ls_t.insertRow();
//mynewrow = ls_t;
//    for(i=0;i<maxcell;i++)
//    {
    mynewrow=ls_t.cells;
    mynewcell=mynewrow.insertCell();
	
//    mynewcell.innerText= Str1;

    mynewcell.innerHTML= Str1;
//    }
}
</script>