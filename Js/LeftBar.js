function RemoveAll()
{
	top.Frame_Right.win.removeall();
}

/*hideMenu() start
 *��ڲ������ޡ�
 *�� �� ֵ���ޡ�
 *��    �ã��������ܡ�
 */
function HideLeftBar()
  {top.Frame_Main.cols="0,*";
   top.frames[1].LeftBar.style.display="";
   show=false;
  }
//hideMenu() end


/*showOrHide() start
 *��ڲ������ޡ�
 *�� �� ֵ���ޡ�
 *��    �ã�����ܱ��ı��Сʱ������ܵĴ�С��
 *          �����ȵ���20������ָܷ���������ߣ���
 *          ���������ܣ�����hideMenu()��������
 *          �����ȴ���20������ָܷ��߱������϶�����
 *          ����ȫ��ʾ���ܣ�����showMenu()��������
 */
var show=false;//��ֻ֤����һ��showMenu()
function showOrHide()
  {if(document.body.clientWidth<=20)
     HideLeftBar();
   else
     if (!show)
       {top.frames[1].ShowLeftBar();
        show=true;
       }
  }
//showOrHide() end