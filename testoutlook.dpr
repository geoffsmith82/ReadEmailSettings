program testoutlook;

uses
  Forms,
  frmtestoutlook in 'frmtestoutlook.pas' {frmMain},
  WindowsAccountSettings in 'WindowsAccountSettings.pas',
  EmailSettings in 'EmailSettings.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
