//编写日期		： 20040908
//编写人			： 
//修改人			： Zerolone
//功能				： 鼠标移入之后显示边框， 类似XP按钮效果
//参数				： obj 对象 state 状态 0为移入 1为移出
function HL_Menu(obj,state)
{
	switch (state)
	{
		case 0:
			obj.className="MenuOver"
			break;
		case 1:
			obj.className="Menu"
			break;
	}
} 

//编写日期		： 20050106
//编写人			： Zerolone
//功能				： 按钮点击事件, 跳转到指定Url
//参数				： 地址, 跳转方式
function GoToUrl(Url, Code)
{
	if(Code==0)
		{
			window.open(Url, "_self");
		}
	else
		{
			window.open(Url, "_blank");
		}
}


//编写日期		： 20040908
//编写人			： 
//修改人			： Zerolone
//功能				： 按钮点击事件, 扩展和收缩菜单
//参数				： 表格ID		
function ExtendNode(TableName)
{
	var TheImg;
	TheImg		= eval("Img_" + TableName);
	TableName	= eval(TableName);

	if(TableName.style.display == "none")
	{
		TableName.style.display = "";
		TheImg.src							= "/images/unwrap.gif";
	}
	else
	{
		TableName.style.display = "none";
		TheImg.src							= "/images/shrink.gif";
	}
}


//编写日期		： 20040908
//编写人			： 
//修改人			： Zerolone
//功能				： 按钮点击事件, 显示到右边窗口的一个TabSheet
//参数				： 
function fnClick()
{
	var oEl = event.srcElement;
	window.top.Frame_Right.AddWin(oEl.href,oEl.innerText);
}

function ffClick(Url, Text)
{
	window.top.Frame_Right.AddWin(Url,Text);
}