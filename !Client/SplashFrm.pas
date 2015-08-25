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
  //��ʼ������
  Setting_Page.Hide;

  Monolith_ini:=TIniFile.Create(ExtractFilePath(Application.ExeName)+'Monolith_ini.ini');
  Try
    Edt_Page_Login.Text         := Monolith_ini.ReadString('�������������', 'Page_Login', 'http://zerolone/manage/login.asp');
    Edt_Page_Portal.Text        := Monolith_ini.ReadString('�������������', 'Page_Portal', 'http://zerolone/');

  Except
    ShowMessage('δ�ҵ�'+ExtractFilePath(Application.ExeName)+'Monolith_ini.ini�ļ���ϵͳʹ��Ĭ�����á�');
  End;
end;

procedure TSplashForm.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  CanClose:=False;
  Monolith_ini.WriteString('�������������', 'Page_Login', Edt_Page_Login.Text);
  Monolith_ini.WriteString('�������������', 'Page_Portal', Edt_Page_Portal.Text);

  MainForm.Visible      := True;
  CanClose:=True;
end;

procedure TSplashForm.Img_SplashClick(Sender: TObject);
begin
  SplashForm.Close;
end;

end.
