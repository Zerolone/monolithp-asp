unit main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, OleCtrls, SHDocVw, ComCtrls, ToolWin, ImgList,
  ExtCtrls, Buttons;

type
  TForm1 = class(TForm)
    WebBrowser1: TWebBrowser;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    Panel1: TPanel;
    Label1: TLabel;
    Edit1: TEdit;
    StatusBar1: TStatusBar;
    BitBtn1: TBitBtn;
    ToolButton3: TToolButton;

    procedure ToolButton8Click(Sender: TObject);
    procedure ToolButton1Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure ToolButton4Click(Sender: TObject);
    procedure ToolButton5Click(Sender: TObject);
    procedure ToolButton6Click(Sender: TObject);
    procedure WebBrowser1DownloadBegin(Sender: TObject);
    procedure WebBrowser1DownloadComplete(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure WebBrowser1NewWindow2(Sender: TObject; var ppDisp: IDispatch;
      var Cancel: WordBool);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  Form2: TForm1;

implementation

{$R *.dfm}



procedure TForm1.ToolButton8Click(Sender: TObject);
begin
  Close;
end;

procedure TForm1.ToolButton1Click(Sender: TObject);
begin
  Try
    WebBrowser1.GoBack;
  except
    showmessage('无效网页');
    exit;
  end;

end;

procedure TForm1.ToolButton2Click(Sender: TObject);
begin
   Try
     WebBrowser1.GoForward;
   except
     showmessage('无效网页');
     exit;
   end;

end;

procedure TForm1.ToolButton4Click(Sender: TObject);
begin
  WebBrowser1.Stop;
end;

procedure TForm1.ToolButton5Click(Sender: TObject);
begin
  WebBrowser1.Refresh;
end;

procedure TForm1.ToolButton6Click(Sender: TObject);
begin
  WebBrowser1.GoHome;
  edit1.Text:='about:blank';
end;

procedure TForm1.WebBrowser1DownloadBegin(Sender: TObject);
begin
     form1.Caption:=Edit1.text+'...';
     StatusBar1.Panels[0].Text:='正在连接地址：'+Edit1.text;
end;

procedure TForm1.WebBrowser1DownloadComplete(Sender: TObject);
begin
   form1.caption:=WebBrowser1.LocationURL ;
   StatusBar1.Panels[0].Text:='完成'+WebBrowser1.LocationURL ;
end;

procedure TForm1.BitBtn1Click(Sender: TObject);
begin
  WebBrowser1.Navigate(Edit1.Text);
end;

procedure TForm1.WebBrowser1NewWindow2(Sender: TObject;
  var ppDisp: IDispatch; var Cancel: WordBool);
begin

   Showmessage('AAA');

 //  Form2.Create();
   Form2.WebBrowser1.Navigate(Edit1.Text);
end;

end.
