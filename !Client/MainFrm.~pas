unit MainFrm;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, OleCtrls, SHDocVw, ComCtrls, ShellAPI, ImgList;

type
  TMainForm = class(TForm)
    Monolith_WB: TWebBrowser;
    Monolith_MainMenu: TMainMenu;
    S1: TMenuItem;
    MenuItem_Login: TMenuItem;
    N1: TMenuItem;
    MenuItem_Exit: TMenuItem;
    Help1: TMenuItem;
    MenuItem_MonolithClient: TMenuItem;
    Setting1: TMenuItem;
    MenuItem_Ini: TMenuItem;
    Monolith_SB: TStatusBar;
    MenuItem_Open: TMenuItem;
    N3: TMenuItem;
    MenuItem_TakeEasy: TMenuItem;
    MenuItem_ShowAllSe: TMenuItem;
    MenuItem_Free: TMenuItem;
    MenuItem_DayTips: TMenuItem;
    MainHotImg: TImageList;
    procedure MenuItem_Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Monolith_WBDownloadBegin(Sender: TObject);
    procedure Monolith_WBDownloadComplete(Sender: TObject);
    procedure Monolith_WBStatusTextChange(Sender: TObject;
      const Text: WideString);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;
  Monolith_Headers:OleVariant;


implementation

uses SplashFrm;

{$R *.dfm}

procedure TMainForm.MenuItem_Click(Sender: TObject);
var
  TheTag:integer;
  //Ini
  Page_Login, Page_Portal:string;
begin
  TheTag:=TMenuItem(Sender).Tag;
  Case TheTag Of
  1://系统｜　&System
    begin
      Page_Login        :=SplashForm.Edt_Page_Login.Text;
      Monolith_WB.Navigate(Page_Login, EmptyParam, EmptyParam, EmptyParam, Monolith_Headers);
    end;

  2://系统参数设置｜　&Ini
    begin
      SplashForm.Img_Splash.Visible     := false;
      SplashForm.Setting_Page.Visible   := True;
      SplashForm.Setting_Page.Left      := 0;
      SplashForm.Setting_Page.Top       := 0;
      SplashForm.BorderStyle            := bsSingle;
      SplashForm.Visible                := True;

      SplashForm.AutoSize               := False;
      MainForm.Visible                  := False;
      SplashForm.AutoSize               := True;
    end;

  3://前台打开Monolith Portal｜　&Open Portal
    begin
      Page_Portal       :=SplashForm.Edt_Page_Portal.Text;
      ShellExecute(handle, 'Open', 'Explorer.EXE', PChar(Page_Portal), nil, SW_SHOW);
    end;

  4://文本收集器｜　&Take Easy
    begin
      ShellExecute(handle, 'Open', 'TakeEasy.exe', nil, nil, SW_SHOW);
    end;

  5://进程察看　｜　&ShowAllSe
    begin
      ShellExecute(handle, 'Open', 'ShowAllSe.exe', nil, nil, SW_SHOW);
    end;

  6://光驱控制　｜　&Free
    begin
      ShellExecute(handle, 'Open', 'Free.exe', nil, nil, SW_SHOW);
    end;

  7://Tips管理　｜　&DayTips
    begin
      ShellExecute(handle, 'Open', 'DayTips.exe', nil, nil, SW_SHOW);
    end;

  8://关于｜　Monolith Client
    begin
      SplashForm.Img_Splash.Visible     := true;
      SplashForm.Setting_Page.Visible   := False;
      SplashForm.BorderStyle            := bsNone;
      SplashForm.Visible                := True;

      SplashForm.AutoSize               := False;
      MainForm.Visible                  := False;
      SplashForm.AutoSize               := True;
    end;

  else
    //No Function
  end;
end;

procedure TMainForm.FormCreate(Sender: TObject);
begin
  Monolith_Headers           :='User-Agent:Monolith Client' + #10#13;
end;

procedure TMainForm.Monolith_WBDownloadBegin(Sender: TObject);
begin
  Monolith_SB.Panels[0].Text := '开始加载';
end;

procedure TMainForm.Monolith_WBDownloadComplete(Sender: TObject);
begin
  Monolith_SB.Panels[0].Text := '加载完毕';
end;

procedure TMainForm.Monolith_WBStatusTextChange(Sender: TObject;
  const Text: WideString);
begin
  Monolith_SB.Panels[0].Text := Text;
end;

end.
