program testoutlook;

uses
  Forms,
  frmtestoutlook in 'frmtestoutlook.pas' {Form2},
  EmailSettings in 'EmailSettings.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm2, Form2);
  Application.Run;
end.
