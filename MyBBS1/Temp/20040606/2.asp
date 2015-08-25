<SCRIPT LANGUAGE="JScript">
function numberCells() 
{
/*    var count=0;
    for (i=0; i < document.all.mytable.rows.length; i++)
    {
        for (j=0; j < document.all.mytable.rows(i).cells.length; j++) 
        {
            document.all.mytable.rows(i).cells(j).innerText = count;
            count++;
        }
    }
*/
}

function tb_addnew()
{
var ls_t=document.all("mytable")
maxcell=ls_t.rows(0).cells.length;
mynewrow = ls_t.insertRow();
    for(i=0;i<maxcell;i++)
    {
    mynewcell=mynewrow.insertCell();
    mynewcell.innerText="a"+i;

    }
}

function tb_delete()
{
var ls_t=document.all("mytable");

ls_t.deleteRow() ;
}

function tb_addcell()
{
//var ls_t=document.all("mytable")
//maxcell=ls_t.rows(0).cells.length;
//mynewrow = ls_t.insertRow();
//    for(i=0;i<maxcell;i++)
 //   {
    mynewrow=document.all("myrow");
    mynewcell=mynewrow.insertCell();
    mynewcell.innerText="aAAAA";
//    }
}


</SCRIPT>
<BODY onload="numberCells()">
<TABLE id=mytable border=1>
<TR id=myrow>
    <Td></Td>
  </TR>
</TABLE>
<input type=button value="新增" onclick="tb_addnew()">
<input type=button value="删除" onclick="tb_delete()" >
<input type=button value="新增Cell" onclick="tb_addcell()">