program MonolithClient;

uses
  Forms,
  MainFrm in 'MainFrm.pas' {MainForm},
  SplashFrm in 'SplashFrm.pas' {SplashForm};

{$R *.res}

begin
  SplashForm:=TsplashForm.Create(application.Owner);
  SplashForm.Show;
  SplashForm.update;

  Application.Initialize;
  Application.Title := 'Monolith Portal Client V1';
  Application.CreateForm(TMainForm, MainForm);

  SplashForm.Close;
//  SplashForm.Free;

  Application.Run;
end.
