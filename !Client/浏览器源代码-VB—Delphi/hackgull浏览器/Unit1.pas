unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, OleCtrls, SHDocVw_TLB, Menus, ToolWin, ComCtrls;

type
  TForm1 = class(TForm)
    StatusBar1: TStatusBar;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    ctrlp1: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    N9: TMenuItem;
    N10: TMenuItem;
    N11: TMenuItem;
    N12: TMenuItem;
    N13: TMenuItem;
    N14: TMenuItem;
    N15: TMenuItem;
    N16: TMenuItem;
    delphi1: TMenuItem;
    WebBrowser1: TWebBrowser;
    OpenDialog1: TOpenDialog;
    PopupMenu1: TPopupMenu;
    CoolBar1: TCoolBar;
    ComboBox1: TComboBox;
    N17: TMenuItem;
    procedure ComboBox1KeyPress(Sender: TObject; var Key: Char);
    procedure N2Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure N14Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.ComboBox1KeyPress(Sender: TObject; var Key: Char);
begin
if Key=#13 then
WebBrowser1.Navigate(ComboBox1.Text);
if key=#27 then
WebBrowser1.GoBack;
end;
procedure TForm1.N2Click(Sender: TObject);
begin
if  OpenDialog1.Execute then
 WebBrowser1.Navigate(OpenDialog1.FileName);
end;

procedure TForm1.N5Click(Sender: TObject);
begin
Form1.Close;
end;

procedure TForm1.N14Click(Sender: TObject);
begin
WebBrowser1.Refresh;
end;

end.
