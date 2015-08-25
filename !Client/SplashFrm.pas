unit SplashFrm;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, IniFiles;

type
  TSplashForm = class(TForm)
    Img_Splash: TImage;
    Setting_Page: TGroupBox;
    Label1: TLabel;
    Edt_Page_Login: TEdit;
    Label2: TLabel;
    Edt_Page_Portal: TEdit;
    Label_About: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure Img_SplashClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  SplashForm: TSplashForm;
  Monolith_ini:TiniFile;
  
implementation

uses MainFrm;

{$R *.dfm}

procedure TSplashForm.FormCreate(Sender: TObject);
begin
  //初始化设置
  Setting_Page.Hide;

  Monolith_ini:=TIniFile.Create(ExtractFilePath(Application.ExeName)+'Monolith_ini.ini');
  Try
    Edt_Page_Login.Text         := Monolith_ini.ReadString('浏览器参数设置', 'Page_Login', 'http://zerolone/manage/login.asp');
    Edt_Page_Portal.Text        := Monolith_ini.ReadString('浏览器参数设置', 'Page_Portal', 'http://zerolone/');

  Except
    ShowMessage('未找到'+ExtractFilePath(Application.ExeName)+'Monolith_ini.ini文件，系统使用默认设置。');
  End;
end;

procedure TSplashForm.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  CanClose:=False;
  Monolith_ini.WriteString('浏览器参数设置', 'Page_Login', Edt_Page_Login.Text);
  Monolith_ini.WriteString('浏览器参数设置', 'Page_Portal', Edt_Page_Portal.Text);

  MainForm.Visible      := True;
  CanClose:=True;
end;

procedure TSplashForm.Img_SplashClick(Sender: TObject);
begin
  SplashForm.Close;
end;

end.
